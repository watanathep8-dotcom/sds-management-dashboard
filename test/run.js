"use strict";
/*
 * Dependency-free test suite:  node test/run.js
 *
 * Backend: end-to-end through doGet/doPost of apps-script/Code.gs running on
 * the in-memory Apps Script mock (test/gas-mock.js), seeded with the catalog
 * embedded in docs/index.html.
 * Frontend: syntax check of every inline <script> in docs/index.html.
 */
const assert = require("assert");
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const { createSeededBackend, CACHE_VALUE_MAX_BYTES } = require("./gas-mock");

const repo = path.resolve(__dirname, "..");
const PASSWORD = "test-password";
const tests = [];
function test(name, fn) { tests.push({ name, fn }); }

const clone = value => JSON.parse(JSON.stringify(value));
const okData = result => {
  assert.ok(result.json, "expected JSON, got " + result.mime + ": " + String(result.text).slice(0, 200));
  assert.strictEqual(result.json.ok, true, "expected ok:true, got " + JSON.stringify(result.json).slice(0, 300));
  return result.json.data;
};
const errorOf = result => {
  assert.strictEqual(result.json.ok, false, "expected ok:false, got " + JSON.stringify(result.json).slice(0, 300));
  return result.json.error;
};
const version = env => env.properties.get("SDS_DATA_VERSION") || "0";
const writesOf = stats => stats.sheetWrites + stats.propWrites + stats.driveWrites + stats.triggerWrites + stats.flushes;

function individualActions(env) {
  return {
    rows: okData(env.get({ action: "api" })),
    columnWidths: okData(env.get({ action: "columnWidths" })),
    columnVisibility: okData(env.get({ action: "columnVisibility" })),
    hazardDefinitions: okData(env.get({ action: "hazardDefinitions" }))
  };
}
function bootstrapOf(env) {
  const data = okData(env.get({ action: "bootstrap" }));
  return { version: data.version, rows: data.rows, columnWidths: data.columnWidths, columnVisibility: data.columnVisibility, hazardDefinitions: data.hazardDefinitions };
}
function withoutVersion(bootstrap) {
  const copy = clone(bootstrap);
  delete copy.version;
  return copy;
}
function assertBootstrapMatchesIndividual(env, message) {
  // Uncached ground truth straight from the sheet.
  const truth = clone(env.context.listSds());
  const boot = bootstrapOf(env);
  const single = individualActions(env);
  assert.deepStrictEqual(clone(boot.rows), truth, (message || "") + " bootstrap rows != sheet");
  assert.deepStrictEqual(withoutVersion(boot), clone(single), (message || "") + " bootstrap != individual actions");
  return boot;
}

/* ------------------------------------------------------------------ */
/* Reads                                                                */
/* ------------------------------------------------------------------ */

test("action=api returns every SDS row from the sheet", () => {
  const env = createSeededBackend();
  const rows = okData(env.get({ action: "api" }));
  assert.ok(rows.length > 50, "seeded rows: " + rows.length);
  assert.deepStrictEqual(clone(rows), clone(env.context.listSds()));
  assert.ok(rows.every(row => row.id && typeof row.chemical === "string" && Array.isArray(row.hazards)));
});

test("bootstrap equals api + columnWidths + columnVisibility + hazardDefinitions", () => {
  const env = createSeededBackend();
  assertBootstrapMatchesIndividual(env, "fresh install:");
  // With saved settings as well.
  okData(env.post({ action: "saveColumnWidths", columnWidths: { 1: 90, 3: 250 }, password: PASSWORD }));
  okData(env.post({ action: "saveColumnVisibility", columnVisibility: { 2: false, 7: false }, password: PASSWORD }));
  okData(env.post({ action: "saveHazardDefinitions", hazardDefinitions: [{ id: "x", name: "pictogram-4.png", image: "assets/pictograms/pictogram-4.png" }, { id: "y", name: "Custom", aliases: ["c"] }] }));
  const boot = assertBootstrapMatchesIndividual(env, "with settings:");
  assert.strictEqual(boot.columnWidths["1"], 90);
  assert.strictEqual(boot.columnVisibility["2"], false);
  assert.strictEqual(boot.hazardDefinitions[1].name, "Custom");
});

test("bootstrap reads the sheet once when cold, not at all when warm, never locks", () => {
  const env = createSeededBackend();
  env.resetStats();
  const first = okData(env.get({ action: "bootstrap" }));
  let stats = env.snapshotStats();
  assert.strictEqual(stats.sheetReads, 1, "cold bootstrap sheet reads");
  assert.strictEqual(stats.propReads, 1, "cold bootstrap property reads");
  assert.strictEqual(stats.locks, 0, "no lock on reads");
  assert.strictEqual(writesOf(stats), 0, "reads write nothing");
  assert.ok(stats.cacheMaxValueBytes <= CACHE_VALUE_MAX_BYTES);

  env.resetStats();
  const second = okData(env.get({ action: "bootstrap" }));
  stats = env.snapshotStats();
  assert.strictEqual(stats.sheetReads, 0, "warm bootstrap must come from cache");
  assert.strictEqual(stats.sheetOpens, 0);
  assert.strictEqual(stats.locks, 0);
  assert.deepStrictEqual(clone(second), clone(first));

  env.resetStats();
  okData(env.get({ action: "api" }));
  okData(env.get({ action: "columnWidths" }));
  okData(env.get({ action: "columnVisibility" }));
  okData(env.get({ action: "hazardDefinitions" }));
  stats = env.snapshotStats();
  assert.strictEqual(stats.sheetReads, 0, "api is served from the same cache");
  assert.strictEqual(stats.locks, 0);
  assert.strictEqual(writesOf(stats), 0);
});

test("cache works without any cached entry (expired/evicted) and with a corrupt chunk", () => {
  const env = createSeededBackend();
  const truth = clone(env.context.listSds());
  okData(env.get({ action: "bootstrap" }));
  // Corrupt one chunk: must be detected and re-read from the sheet.
  const chunkKey = Array.from(env.cache.keys()).find(key => /_0$/.test(key));
  env.cache.get(chunkKey).value = "{broken";
  env.resetStats();
  assert.deepStrictEqual(clone(okData(env.get({ action: "bootstrap" })).rows), truth);
  assert.strictEqual(env.snapshotStats().sheetReads, 1);
  env.cache.clear();
  assert.deepStrictEqual(clone(okData(env.get({ action: "api" }))), truth);
});

test("JSONP callback and old GET actions still work; unknown action serves the HTML page", () => {
  const env = createSeededBackend();
  const jsonp = env.get({ action: "bootstrap", callback: "cb_1" });
  assert.strictEqual(jsonp.mime, "application/javascript");
  assert.ok(/^cb_1\(\{"ok":true/.test(jsonp.text));
  const api = env.get({ action: "api", callback: "cb_2" });
  assert.ok(/^cb_2\(\{"ok":true,"data":\[/.test(api.text));
  assert.ok(env.get({}).html, "no action -> HtmlService page");
  assert.ok(env.get({ action: "nope" }).html);
});

test("file action returns an uploaded PDF", () => {
  const env = createSeededBackend();
  const saved = okData(env.post({
    action: "save",
    record: { chemicalEnglish: "PDF holder" },
    file: { name: "a.pdf", mimeType: "application/pdf", base64: Buffer.from("%PDF-1.4 test").toString("base64") }
  }));
  assert.ok(saved.pdfFileId);
  const file = env.get({ action: "file", fileId: saved.pdfFileId }).json;
  assert.strictEqual(file.ok, true);
  assert.strictEqual(Buffer.from(file.base64, "base64").toString(), "%PDF-1.4 test");
});

/* ------------------------------------------------------------------ */
/* Writes + cache invalidation                                          */
/* ------------------------------------------------------------------ */

test("save new record without password; response carries fresh bootstrap; cache invalidated", () => {
  const env = createSeededBackend();
  const before = bootstrapOf(env); // warm the cache
  const v0 = version(env);
  const response = env.post({ action: "save", record: { chemicalEnglish: "New Chemical", cas: "1-2-3", hazards: ["pictogram-7.png"] } });
  const saved = okData(response);
  assert.ok(saved.id && saved.chemical === "New Chemical");
  assert.notStrictEqual(version(env), v0, "version bumped");
  const boot = response.json.bootstrap;
  assert.ok(boot && Array.isArray(boot.rows), "save returns bootstrap");
  assert.strictEqual(boot.rows.length, before.rows.length + 1);
  assert.ok(boot.rows.some(row => row.id === saved.id));
  env.resetStats();
  const after = bootstrapOf(env);
  assert.strictEqual(env.snapshotStats().sheetReads, 0, "save re-warmed the cache for the new version");
  assert.deepStrictEqual(clone(after), clone(boot));
  assertBootstrapMatchesIndividual(env, "after save:");
});

test("updating an existing record needs the password; wrong/missing password changes nothing", () => {
  const env = createSeededBackend();
  const before = bootstrapOf(env);
  const target = before.rows[0];
  const v0 = version(env);
  for (const password of [undefined, "", "wrong"]) {
    const error = errorOf(env.post({ action: "save", record: { id: target.id, chemicalEnglish: "Hacked" }, password }));
    assert.ok(/Incorrect password/.test(error), error);
  }
  assert.strictEqual(version(env), v0, "failed writes do not bump the version");
  env.resetStats();
  assert.deepStrictEqual(clone(bootstrapOf(env)), clone(before));
  assert.strictEqual(env.snapshotStats().sheetReads, 0, "cache still valid after rejected writes");

  const response = env.post({ action: "save", record: { id: target.id, chemicalEnglish: "Renamed", cas: target.cas }, password: PASSWORD });
  const saved = okData(response);
  assert.strictEqual(saved.chemicalEnglish, "Renamed");
  // Fields the client omitted are kept (existing behaviour).
  assert.strictEqual(saved.supplier, target.supplier);
  assert.deepStrictEqual(clone(saved.firstAidMeasures), clone(target.firstAidMeasures));
  assert.strictEqual(saved.sequence, target.sequence);
  const after = assertBootstrapMatchesIndividual(env, "after update:");
  assert.strictEqual(after.rows.find(row => row.id === target.id).chemicalEnglish, "Renamed");
  assert.strictEqual(after.rows.length, before.rows.length, "update does not append");
  assert.deepStrictEqual(clone(response.json.bootstrap), clone(after));
});

test("save takes the script lock and reads the sheet data once before writing", () => {
  const env = createSeededBackend();
  const target = env.context.listSds()[5];
  env.resetStats();
  okData(env.post({ action: "save", record: { id: target.id, chemicalEnglish: "X" }, password: PASSWORD }));
  const stats = env.snapshotStats();
  assert.strictEqual(stats.locks, 1);
  // header check + data rows (save) + one fresh read for the returned bootstrap
  assert.ok(stats.sheetReads <= 3, "sheet reads during save: " + stats.sheetReads);
});

test("delete needs the password; deleted records disappear from api and bootstrap", () => {
  const env = createSeededBackend();
  const before = bootstrapOf(env);
  const target = before.rows[2];
  const v0 = version(env);
  assert.ok(/Incorrect password/.test(errorOf(env.post({ action: "delete", id: target.id, password: "nope" }))));
  assert.strictEqual(version(env), v0);
  const response = env.post({ action: "delete", id: target.id, password: PASSWORD });
  assert.deepStrictEqual(clone(okData(response)), { deleted: true });
  assert.notStrictEqual(version(env), v0);
  assert.ok(!response.json.bootstrap.rows.some(row => row.id === target.id), "delete response bootstrap");
  const after = assertBootstrapMatchesIndividual(env, "after delete:");
  assert.strictEqual(after.rows.length, before.rows.length - 1);
  assert.ok(!after.rows.some(row => row.id === target.id));
  assert.ok(!okData(env.get({ action: "api" })).some(row => row.id === target.id));
  // Deleting again: nothing to delete, no version bump.
  const v1 = version(env);
  assert.deepStrictEqual(clone(okData(env.post({ action: "delete", id: target.id, password: PASSWORD }))), { deleted: false });
  assert.strictEqual(version(env), v1);
});

test("delete trashes the record's uploaded PDF", () => {
  const env = createSeededBackend();
  const saved = okData(env.post({ action: "save", record: { chemicalEnglish: "With file" }, file: { name: "x.pdf", base64: Buffer.from("pdf").toString("base64") } }));
  okData(env.post({ action: "delete", id: saved.id, password: PASSWORD }));
  assert.strictEqual(env.driveFiles.get(saved.pdfFileId).trashed, true);
});

test("table settings saves need the password and invalidate the cache", () => {
  const env = createSeededBackend();
  bootstrapOf(env);
  const v0 = version(env);
  assert.ok(/Incorrect password/.test(errorOf(env.post({ action: "saveColumnWidths", columnWidths: { 1: 100 }, password: "bad" }))));
  assert.ok(/Incorrect password/.test(errorOf(env.post({ action: "saveColumnVisibility", columnVisibility: { 1: false } }))));
  assert.strictEqual(version(env), v0);

  const widths = okData(env.post({ action: "saveColumnWidths", columnWidths: { 1: 999, 2: 10, 3: "abc" }, password: PASSWORD }));
  assert.strictEqual(widths["1"], 180, "clamped to max");
  assert.strictEqual(widths["2"], 75, "clamped to min");
  assert.strictEqual(widths["3"], 180, "invalid -> fallback");
  const v1 = version(env);
  assert.notStrictEqual(v1, v0);
  assert.deepStrictEqual(clone(bootstrapOf(env).columnWidths), clone(widths));

  const allHidden = {};
  for (let i = 1; i <= 15; i++) allHidden[i] = false;
  assert.ok(/At least one heading/.test(errorOf(env.post({ action: "saveColumnVisibility", columnVisibility: allHidden, password: PASSWORD }))));
  assert.strictEqual(version(env), v1);
  const visibility = okData(env.post({ action: "saveColumnVisibility", columnVisibility: { 4: false, 5: "false" }, password: PASSWORD }));
  assert.strictEqual(visibility["4"], false);
  assert.strictEqual(visibility["5"], false);
  assert.strictEqual(visibility["1"], true);
  assert.notStrictEqual(version(env), v1);
  assertBootstrapMatchesIndividual(env, "after visibility:");
});

test("hazard definition saves invalidate the cache", () => {
  const env = createSeededBackend();
  bootstrapOf(env);
  const v0 = version(env);
  assert.ok(errorOf(env.post({ action: "saveHazardDefinitions", hazardDefinitions: [] })));
  assert.strictEqual(version(env), v0);
  const saved = okData(env.post({ action: "saveHazardDefinitions", hazardDefinitions: [{ id: "a", name: "Alpha", image: "data:image/png;base64,xx" }] }));
  assert.strictEqual(saved[0].image, "", "data: URLs are dropped");
  assert.notStrictEqual(version(env), v0);
  assert.deepStrictEqual(clone(bootstrapOf(env).hazardDefinitions), clone(saved));
  assertBootstrapMatchesIndividual(env, "after hazards:");
});

test("stale cache entries of an older version are never served after a write", () => {
  const env = createSeededBackend();
  const old = bootstrapOf(env);
  const oldKeys = Array.from(env.cache.keys());
  okData(env.post({ action: "save", record: { chemicalEnglish: "Fresh" } }));
  // Old entries are still physically in the cache...
  assert.ok(oldKeys.every(key => env.cache.has(key)));
  // ...but every read now sees the new data, even if the post-write re-warm was lost.
  for (const key of Array.from(env.cache.keys())) if (!oldKeys.includes(key)) env.cache.delete(key);
  const fresh = bootstrapOf(env);
  assert.strictEqual(fresh.rows.length, old.rows.length + 1);
  assert.notStrictEqual(fresh.version, old.version);
  assert.ok(okData(env.get({ action: "api" })).some(row => row.chemical === "Fresh"));
});

test("verifyPassword and lockout after repeated wrong passwords", () => {
  const env = createSeededBackend();
  assert.deepStrictEqual(clone(okData(env.post({ action: "verifyPassword", password: PASSWORD }))), { valid: true });
  for (let i = 0; i < 10; i++) {
    assert.ok(/Incorrect password/.test(errorOf(env.post({ action: "verifyPassword", password: "wrong" + i }))));
  }
  assert.strictEqual(env.snapshotStats().sleeps, 10, "each wrong attempt is slowed down");
  // Now even the right password is refused for every protected action.
  assert.ok(/Too many wrong attempts/.test(errorOf(env.post({ action: "verifyPassword", password: PASSWORD }))));
  const target = env.context.listSds()[0];
  assert.ok(/Too many wrong attempts/.test(errorOf(env.post({ action: "delete", id: target.id, password: PASSWORD }))));
  assert.ok(/Too many wrong attempts/.test(errorOf(env.post({ action: "save", record: { id: target.id, chemicalEnglish: "x" }, password: PASSWORD }))));
  assert.ok(/Too many wrong attempts/.test(errorOf(env.post({ action: "saveColumnWidths", columnWidths: { 1: 90 }, password: PASSWORD }))));
  // Reads are unaffected and the cache survives the lockout entries.
  assert.ok(okData(env.get({ action: "bootstrap" })).rows.length > 0);
  // Lock expires (the failure counter lives in CacheService).
  env.cache.delete("SDS_PASSWORD_FAILURES");
  assert.deepStrictEqual(clone(okData(env.post({ action: "verifyPassword", password: PASSWORD }))), { valid: true });
});

test("missing action password configuration is reported, not bypassed", () => {
  const env = createSeededBackend();
  env.properties.delete("SDS_ACTION_PASSWORD");
  assert.ok(/not configured/.test(errorOf(env.post({ action: "delete", id: "1", password: "" }))));
});

test("unsupported POST action", () => {
  const env = createSeededBackend();
  assert.ok(/Unsupported action/.test(errorOf(env.post({ action: "dropTable" }))));
});

/* ------------------------------------------------------------------ */
/* Cache size limits                                                    */
/* ------------------------------------------------------------------ */

test("large payloads are chunked under the 100 KB per-key limit (Thai text)", () => {
  const env = createSeededBackend();
  const thai = "ข้อมูลความปลอดภัยสารเคมี ".repeat(400);
  for (let i = 0; i < 40; i++) {
    okData(env.post({ action: "save", record: { chemicalEnglish: "Big " + i, emergencyResponse: thai, composition: "😀" + thai } }));
  }
  env.cache.clear();
  env.resetStats();
  const cold = okData(env.get({ action: "bootstrap" }));
  const bytes = Buffer.byteLength(JSON.stringify(cold));
  assert.ok(bytes > 1024 * 1024, "payload should be > 1 MB, is " + bytes);
  assert.ok(env.snapshotStats().cacheMaxValueBytes <= CACHE_VALUE_MAX_BYTES);
  env.resetStats();
  const warm = okData(env.get({ action: "bootstrap" }));
  assert.strictEqual(env.snapshotStats().sheetReads, 0, "served from chunks");
  assert.deepStrictEqual(clone(warm), clone(cold));
});

test("payloads beyond the chunk budget skip caching but are still served correctly", () => {
  const env = createSeededBackend();
  env.context.BOOTSTRAP_CACHE_MAX_CHUNKS = 2;
  env.cache.clear();
  const first = okData(env.get({ action: "bootstrap" }));
  env.resetStats();
  const second = okData(env.get({ action: "bootstrap" }));
  assert.strictEqual(env.snapshotStats().sheetReads, 1, "not cached -> read again");
  assert.deepStrictEqual(clone(second), clone(first));
  assert.deepStrictEqual(clone(second.rows), clone(env.context.listSds()));
});

/* ------------------------------------------------------------------ */
/* keepWarm + triggers                                                  */
/* ------------------------------------------------------------------ */

test("keepWarm refreshes the cache and writes nothing", () => {
  const env = createSeededBackend();
  const propsBefore = clone(Object.fromEntries(env.properties));
  const sheetBefore = clone(env.sheet()._state.values);
  env.cache.clear();
  env.resetStats();
  const result = env.context.keepWarm();
  const stats = env.snapshotStats();
  assert.strictEqual(writesOf(stats), 0, "keepWarm wrote: " + JSON.stringify(stats));
  assert.strictEqual(stats.locks, 0);
  assert.strictEqual(result.cached, true);
  assert.deepStrictEqual(clone(Object.fromEntries(env.properties)), propsBefore);
  assert.deepStrictEqual(clone(env.sheet()._state.values), sheetBefore);
  env.resetStats();
  okData(env.get({ action: "bootstrap" }));
  assert.strictEqual(env.snapshotStats().sheetReads, 0, "bootstrap after keepWarm is a cache hit");
  // keepWarm also picks up manual edits made directly in the sheet.
  env.sheet()._state.values[1][1] = "Edited in Sheets";
  env.sheet()._state.values[1][19] = "Edited in Sheets";
  env.context.keepWarm();
  assert.strictEqual(okData(env.get({ action: "bootstrap" })).rows[0].chemical, "Edited in Sheets");
});

test("keepWarm skips (instead of repairing) a sheet whose header needs fixing", () => {
  const env = createSeededBackend();
  env.sheet()._state.values[0][3] = "tampered";
  env.cache.clear();
  env.resetStats();
  const result = env.context.keepWarm();
  assert.strictEqual(writesOf(env.snapshotStats()), 0);
  assert.strictEqual(result.cached, false);
  // A normal request still repairs the header exactly like before.
  okData(env.get({ action: "bootstrap" }));
  assert.strictEqual(env.sheet()._state.values[0][3], "supplier");
});

test("installKeepWarmTrigger is idempotent; removeKeepWarmTrigger removes it", () => {
  const env = createSeededBackend();
  env.triggers.push({ handler: "somethingElse", id: "other" });
  const first = env.context.installKeepWarmTrigger();
  assert.strictEqual(first.installed, true);
  const second = env.context.installKeepWarmTrigger();
  assert.strictEqual(second.installed, false);
  const keepWarmTriggers = env.triggers.filter(t => t.handler === "keepWarm");
  assert.strictEqual(keepWarmTriggers.length, 1);
  assert.strictEqual(keepWarmTriggers[0].everyMinutes, 10);
  // Duplicates created by hand are collapsed to one.
  env.triggers.push({ handler: "keepWarm", id: "dup" });
  env.context.installKeepWarmTrigger();
  assert.strictEqual(env.triggers.filter(t => t.handler === "keepWarm").length, 1);
  assert.deepStrictEqual(clone(env.context.removeKeepWarmTrigger()), { removed: 1 });
  assert.strictEqual(env.triggers.filter(t => t.handler === "keepWarm").length, 0);
  assert.strictEqual(env.triggers.length, 1, "other triggers untouched");
});

/* ------------------------------------------------------------------ */
/* Legacy compatibility: the previous Code.gs answers the same          */
/* ------------------------------------------------------------------ */

test("new backend returns exactly what the previous backend returned (api + settings)", () => {
  const { execSync } = require("child_process");
  let oldCode;
  try {
    oldCode = execSync("git show HEAD~0:apps-script/Code.gs", { cwd: repo, encoding: "utf8", stdio: ["ignore", "pipe", "ignore"] });
  } catch (error) {
    console.log("    (skipped: git not available)");
    return;
  }
  if (!/listSds/.test(oldCode) || /getBootstrap_/.test(oldCode)) {
    // HEAD already contains this change; compare against the commit before it.
    try {
      const log = execSync("git log --format=%H -n 20 -- apps-script/Code.gs", { cwd: repo, encoding: "utf8" }).trim().split("\n");
      const base = log.find(hash => !/getBootstrap_/.test(execSync(`git show ${hash}:apps-script/Code.gs`, { cwd: repo, encoding: "utf8" })));
      if (!base) { console.log("    (skipped: no pre-bootstrap Code.gs in history)"); return; }
      oldCode = execSync(`git show ${base}:apps-script/Code.gs`, { cwd: repo, encoding: "utf8" });
    } catch (error) {
      console.log("    (skipped: " + error.message + ")");
      return;
    }
  }
  const tmp = path.join(require("os").tmpdir(), "sds-old-code-" + process.pid + ".gs");
  fs.writeFileSync(tmp, oldCode);
  try {
    const oldEnv = createSeededBackend({ codePath: tmp });
    const newEnv = createSeededBackend();
    const script = [
      env => env.get({ action: "api" }),
      env => env.post({ action: "save", record: { chemicalEnglish: "Cmp", cas: "50-00-0" } }),
      env => env.post({ action: "saveColumnWidths", columnWidths: { 2: 120 }, password: PASSWORD }),
      env => env.post({ action: "saveColumnVisibility", columnVisibility: { 3: false }, password: PASSWORD }),
      env => env.post({ action: "saveHazardDefinitions", hazardDefinitions: [{ name: "Z" }] }),
      env => env.post({ action: "delete", id: env.context.listSds()[1].id, password: PASSWORD }),
      env => env.post({ action: "verifyPassword", password: "bad" }),
      env => env.get({ action: "api" }),
      env => env.get({ action: "columnWidths" }),
      env => env.get({ action: "columnVisibility" }),
      env => env.get({ action: "hazardDefinitions" })
    ];
    script.forEach((step, index) => {
      const a = step(oldEnv).json;
      const b = step(newEnv).json;
      const strip = value => {
        const copy = clone(value);
        delete copy.bootstrap; // additive field of the new backend
        if (copy.data && copy.data.updatedAt) delete copy.data.updatedAt;
        if (copy.data && /^\d{13}$/.test(copy.data.id || "")) delete copy.data.id;
        if (Array.isArray(copy.data)) copy.data.forEach(row => { delete row.updatedAt; if (/^\d{13}$/.test(row.id || "") && Number(row.id) > 1700000000000) delete row.id; });
        return copy;
      };
      assert.deepStrictEqual(strip(b), strip(a), "step " + index + " differs");
    });
  } finally {
    fs.unlinkSync(tmp);
  }
});

/* ------------------------------------------------------------------ */
/* Frontend                                                             */
/* ------------------------------------------------------------------ */

function inlineScripts(html) {
  const scripts = [];
  const pattern = /<script\b([^>]*)>([\s\S]*?)<\/script>/gi;
  let match;
  while ((match = pattern.exec(html))) {
    if (/\bsrc\s*=/.test(match[1])) continue;
    const line = html.slice(0, match.index).split("\n").length;
    scripts.push({ attributes: match[1], code: match[2], line });
  }
  return scripts;
}

test("inline scripts in docs/index.html parse", () => {
  const html = fs.readFileSync(path.join(repo, "docs", "index.html"), "utf8");
  const scripts = inlineScripts(html);
  assert.ok(scripts.length >= 3, "found " + scripts.length + " inline scripts");
  scripts.forEach(script => {
    try {
      new vm.Script(script.code, { filename: "index.html:" + script.line });
    } catch (error) {
      throw new Error("Syntax error in inline script starting at line " + script.line + ": " + error.message);
    }
  });
  // config.js must load before the early bootstrap request in <head>.
  const head = html.slice(0, html.indexOf("</head>"));
  assert.ok(head.indexOf('src="config.js') !== -1 && head.indexOf('src="config.js') < head.indexOf("sdsEarlyBootstrap"));
  assert.strictEqual((html.match(/src="config\.js/g) || []).length, 1, "config.js is loaded once");
});

test("early bootstrap script requests action=bootstrap once and is skipped without apiUrl", () => {
  const html = fs.readFileSync(path.join(repo, "docs", "index.html"), "utf8");
  const early = inlineScripts(html).find(script => /sdsEarlyBootstrap/.test(script.code) && script.code.length < 5000);
  assert.ok(early, "early bootstrap script present");
  const run = config => {
    const calls = [];
    const sandbox = {
      window: { SDS_CONFIG: config, location: { href: "https://example.github.io/sds/" } },
      URL, String, Date,
      fetch: (url, options) => { calls.push({ url, options }); return Promise.resolve({ ok: true, json: () => Promise.resolve({ ok: true, data: { rows: [] } }) }); }
    };
    sandbox.window.fetch = sandbox.fetch;
    vm.runInNewContext(early.code, sandbox);
    return { calls, window: sandbox.window };
  };
  const remote = run({ apiUrl: "https://script.google.com/macros/s/X/exec" });
  assert.strictEqual(remote.calls.length, 1);
  assert.ok(/action=bootstrap/.test(remote.calls[0].url));
  assert.strictEqual(remote.calls[0].options.credentials, "omit");
  assert.ok(remote.window.sdsEarlyBootstrap && remote.window.sdsEarlyBootstrap.request);
  assert.strictEqual(run({ apiUrl: "" }).calls.length, 0);
});

/* ------------------------------------------------------------------ */

(async () => {
  let failed = 0;
  for (const { name, fn } of tests) {
    try {
      await fn();
      console.log("  ok   " + name);
    } catch (error) {
      failed += 1;
      console.log("  FAIL " + name);
      console.log("       " + String(error && error.stack || error).split("\n").slice(0, 6).join("\n       "));
    }
  }
  console.log(`\n${tests.length - failed}/${tests.length} passed`);
  process.exitCode = failed ? 1 : 0;
})();
