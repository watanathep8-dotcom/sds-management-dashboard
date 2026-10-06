"use strict";
/*
 * Local preview harness: serves docs/ and a mock Apps Script web app.
 *
 *   node test/server.js [--port 8787] [--docs docs] [--code apps-script/Code.gs]
 *                       [--latency 1500] [--op-latency 200]
 *
 * - GET  /exec?action=...  -> doGet  of Code.gs running on the in-memory mock
 * - POST /exec             -> doPost of Code.gs
 * - /config.js             -> docs/config.js with apiUrl pointed at /exec
 * - /__log                 -> JSON list of API calls since the page was last loaded
 *
 * Every API response is delayed to imitate Apps Script: a fixed base cost
 * per request (--latency) plus --op-latency per Sheet read/write and smaller
 * costs for Properties/Cache/Drive calls, measured from the mock's counters.
 * Nothing here talks to Google.
 */
const http = require("http");
const fs = require("fs");
const path = require("path");
const { createSeededBackend } = require("./gas-mock");

const repo = path.resolve(__dirname, "..");
const args = process.argv.slice(2);
function arg(name, fallback) {
  const index = args.indexOf("--" + name);
  return index !== -1 && args[index + 1] !== undefined ? args[index + 1] : fallback;
}
const port = Number(arg("port", 8787));
const docsDir = path.resolve(arg("docs", path.join(repo, "docs")));
const codePath = path.resolve(arg("code", path.join(repo, "apps-script", "Code.gs")));
const baseLatency = Number(arg("latency", 1500));
const opLatency = Number(arg("op-latency", 200));

const backend = createSeededBackend({ codePath, indexHtml: path.join(docsDir, "index.html") });
let log = [];
let pageLoadedAt = Date.now();

const MIME = {
  ".html": "text/html; charset=utf-8", ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8", ".json": "application/json", ".png": "image/png",
  ".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".svg": "image/svg+xml", ".pdf": "application/pdf",
  ".ico": "image/x-icon", ".webp": "image/webp", ".gif": "image/gif"
};

function simulatedCost(before, after) {
  const d = key => after[key] - before[key];
  return baseLatency
    + opLatency * (d("sheetReads") + d("sheetWrites") + d("flushes"))
    + 100 * d("sheetOpens")
    + 20 * (d("propReads") + d("propWrites") + d("cacheGets") + d("cachePuts"))
    + 300 * d("driveWrites")
    + 1000 * d("sleeps");
}

function handleApi(req, res, url, body) {
  const started = Date.now();
  const before = backend.snapshotStats();
  let output;
  let action;
  try {
    if (req.method === "POST") {
      const payload = JSON.parse(body || "{}");
      action = payload.action;
      output = backend.context.doPost({ postData: { contents: body, type: "text/plain" } });
    } else {
      const parameter = Object.fromEntries(url.searchParams);
      action = parameter.action;
      output = backend.context.doGet({ parameter });
    }
  } catch (error) {
    res.writeHead(500, { "Content-Type": "text/plain" });
    res.end(String(error && error.stack || error));
    return;
  }
  const after = backend.snapshotStats();
  const delay = simulatedCost(before, after);
  const text = output.getContent();
  setTimeout(() => {
    const ended = Date.now();
    const entry = {
      method: req.method, action: action || "(none)",
      start: started - pageLoadedAt, end: ended - pageLoadedAt, ms: ended - started,
      bytes: Buffer.byteLength(text), sheetReads: after.sheetReads - before.sheetReads,
      cacheHit: after.cacheHits > before.cacheHits
    };
    log.push(entry);
    console.log(`[api] ${entry.method} ${entry.action} +${entry.start}ms -> +${entry.end}ms (${entry.ms}ms, ${entry.bytes} B, sheetReads=${entry.sheetReads}${entry.cacheHit ? ", cache hit" : ""})`);
    res.writeHead(200, { "Content-Type": output.getMimeType() + "; charset=utf-8", "Cache-Control": "no-store" });
    res.end(text);
  }, delay);
}

function serveStatic(req, res, url) {
  let pathname = decodeURIComponent(url.pathname);
  if (pathname === "/" || pathname === "") pathname = "/index.html";
  if (pathname === "/index.html") { pageLoadedAt = Date.now(); log = []; }
  if (pathname === "/config.js") {
    const source = fs.readFileSync(path.join(docsDir, "config.js"), "utf8");
    const rewritten = source.replace(/apiUrl:\s*"[^"]*"/, `apiUrl: "http://localhost:${port}/exec"`);
    res.writeHead(200, { "Content-Type": MIME[".js"], "Cache-Control": "no-store" });
    res.end(rewritten);
    return;
  }
  const candidates = [path.join(docsDir, pathname), path.join(repo, "docs", pathname)];
  const file = candidates.find(candidate => {
    const resolved = path.resolve(candidate);
    return (resolved.startsWith(docsDir) || resolved.startsWith(path.join(repo, "docs"))) &&
      fs.existsSync(resolved) && fs.statSync(resolved).isFile();
  });
  if (!file) { res.writeHead(404); res.end("Not found"); return; }
  res.writeHead(200, { "Content-Type": MIME[path.extname(file).toLowerCase()] || "application/octet-stream", "Cache-Control": "no-store" });
  fs.createReadStream(file).pipe(res);
}

const server = http.createServer((req, res) => {
  const url = new URL(req.url, `http://localhost:${port}`);
  if (url.pathname === "/__log") {
    res.writeHead(200, { "Content-Type": "application/json", "Cache-Control": "no-store" });
    res.end(JSON.stringify(log, null, 2));
    return;
  }
  if (url.pathname === "/exec") {
    let body = "";
    req.on("data", chunk => { body += chunk; });
    req.on("end", () => handleApi(req, res, url, body));
    return;
  }
  serveStatic(req, res, url);
});

server.listen(port, () => {
  console.log(`SDS preview on http://localhost:${port}/  (docs=${docsDir}, code=${codePath}, latency=${baseLatency}+${opLatency}/op)`);
  console.log(`Action password for this mock: test-password`);
});
