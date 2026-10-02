/**
 * SDS Management backend for Google Apps Script.
 *
 * Run setupSystem() once from the Apps Script editor. It creates the
 * spreadsheet and Drive folder used by this app and stores their IDs in
 * Script Properties.
 */

var SDS_HEADERS = [
  "id", "chemical", "cas", "supplier", "revision", "revisionDate",
  "status", "signalWord", "hazards", "pdfFileId", "pdfName", "updatedAt",
  "reviewDate", "thaiSds", "language", "flashPoint", "emergencyResponse", "sdsLanguages",
  "chemicalThai", "chemicalEnglish", "thaiPdfFileId", "thaiPdfName",
  "englishPdfFileId", "englishPdfName", "pdfUrl", "thaiPdfUrl", "englishPdfUrl",
  "composition", "ratio", "personalProtectiveEquipment", "firstAidMeasures", "firefightingMeasures",
  "sequence", "group", "responsibleParty", "recorder", "dateOfUse", "registerSds", "components"
];

// Editing and deleting SDS records require a shared password. It is stored
// only in Script Properties (Project Settings > Script properties >
// SDS_ACTION_PASSWORD) so it never appears in this public source code.
// Repeated wrong attempts temporarily lock the check to slow down guessing.
var ACTION_PASSWORD_PROPERTY = "SDS_ACTION_PASSWORD";
var PASSWORD_FAIL_CACHE_KEY = "SDS_PASSWORD_FAILURES";
var PASSWORD_MAX_FAILURES = 10;
var PASSWORD_LOCK_SECONDS = 900;

// The public GitHub dashboard is intentionally allowed to create and update
// SDS records. Deletion is also allowed when the shared delete password is
// supplied; the explicit switch keeps this public-write policy easy to audit.
var PUBLIC_SDS_WRITE_ENABLED = true;

// Shared UI settings are stored in Script Properties so every viewer of the
// public dashboard receives the same table layout. This setting contains only
// validated column widths; it does not contain SDS records or credentials.
var COLUMN_WIDTHS_PROPERTY = "SDS_TABLE_COLUMN_WIDTHS_V1";
var COLUMN_VISIBILITY_PROPERTY = "SDS_TABLE_COLUMN_VISIBILITY_V1";
// Hazard/pictogram definitions edited from the dashboard Settings dialog.
// Only names, aliases, icon ids and asset paths are stored (no uploads), so
// the JSON fits comfortably inside a single Script Property value.
var HAZARD_DEFINITIONS_PROPERTY = "SDS_HAZARD_DEFINITIONS_V1";
var HAZARD_DEFINITIONS_MAX_ITEMS = 60;
var HAZARD_DEFINITIONS_MAX_BYTES = 8500;
var TABLE_COLUMN_INDICES = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15"];
var COLUMN_WIDTH_LIMITS = {
  "1": { min: 55, max: 180, fallback: 70 },
  "2": { min: 75, max: 220, fallback: 100 },
  "3": { min: 130, max: 380, fallback: 180 },
  "4": { min: 150, max: 340, fallback: 200 },
  "5": { min: 130, max: 380, fallback: 180 },
  "6": { min: 80, max: 220, fallback: 100 },
  "7": { min: 85, max: 220, fallback: 100 },
  "8": { min: 90, max: 220, fallback: 100 },
  "9": { min: 140, max: 420, fallback: 200 },
  "10": { min: 140, max: 420, fallback: 200 },
  "11": { min: 85, max: 220, fallback: 100 },
  "12": { min: 85, max: 220, fallback: 100 },
  "13": { min: 85, max: 240, fallback: 100 },
  "14": { min: 55, max: 160, fallback: 70 },
  "15": { min: 55, max: 160, fallback: 70 }
};

// New installations start empty. SDS records may be submitted from the public
// dashboard; edits and deletes still require an administrator session.
var SAMPLE_ROWS = [];

function doGet(e) {
  var params = (e && e.parameter) || {};

  if (params.action === "api") {
    var response = { ok: true, data: listSds() };
    return jsonResponse_(response, params.callback);
  }

  if (params.action === "file") {
    try {
      return jsonResponse_(getSdsFile(params.fileId || params.id), params.callback);
    } catch (error) {
      return jsonResponse_({ ok: false, error: error.message }, params.callback);
    }
  }

  if (params.action === "columnWidths") {
    return jsonResponse_({ ok: true, data: getColumnWidths_() }, params.callback);
  }

  if (params.action === "columnVisibility") {
    return jsonResponse_({ ok: true, data: getColumnVisibility_() }, params.callback);
  }

  if (params.action === "hazardDefinitions") {
    return jsonResponse_({ ok: true, data: getHazardDefinitions_() }, params.callback);
  }

  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("ระบบจัดการ SDS | SDS Management")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var body = (e && e.postData && e.postData.contents) || "{}";
    var payload = JSON.parse(body);
    var action = payload.action;
    var result;

    if (action === "save") {
      result = saveSds(payload.record || {}, payload.file || null, payload.password);
    } else if (action === "delete") {
      result = deleteSds(payload.id, payload.password);
    } else if (action === "verifyPassword") {
      result = verifyActionPassword(payload.password);
    } else if (action === "saveColumnWidths") {
      result = saveColumnWidths_(payload.columnWidths);
    } else if (action === "saveColumnVisibility") {
      result = saveColumnVisibility_(payload.columnVisibility);
    } else if (action === "saveHazardDefinitions") {
      result = saveHazardDefinitions_(payload.hazardDefinitions);
    } else {
      throw new Error("ไม่รองรับคำสั่งนี้ / Unsupported action.");
    }

    return jsonResponse_({ ok: true, data: result });
  } catch (error) {
    return jsonResponse_({ ok: false, error: error.message });
  }
}

function getColumnWidths_() {
  var properties = PropertiesService.getScriptProperties();
  var raw = properties.getProperty(COLUMN_WIDTHS_PROPERTY);
  var saved = {};
  if (raw) {
    try { saved = JSON.parse(raw) || {}; } catch (ignored) { saved = {}; }
  }

  var normalized = {};
  Object.keys(COLUMN_WIDTH_LIMITS).forEach(function(index) {
    if (saved[index] === undefined || saved[index] === null || saved[index] === "") return;
    normalized[index] = normalizeColumnWidth_(saved[index], COLUMN_WIDTH_LIMITS[index]);
  });
  return normalized;
}

// Public wrappers are used by the Apps Script-hosted dashboard through
// google.script.run. The underscore-suffixed helpers remain internal.
function getColumnWidths() {
  return getColumnWidths_();
}

function saveColumnWidths_(widths) {
  if (!widths || typeof widths !== "object" || Array.isArray(widths)) {
    throw new Error("รูปแบบค่าความกว้างตารางไม่ถูกต้อง / Invalid column width settings.");
  }

  var normalized = {};
  Object.keys(COLUMN_WIDTH_LIMITS).forEach(function(index) {
    var limit = COLUMN_WIDTH_LIMITS[index];
    var value = widths[index];
    if (value === undefined || value === null || value === "") value = limit.fallback;
    normalized[index] = normalizeColumnWidth_(value, limit);
  });

  // Deliberately store only the validated layout object. This endpoint does
  // not accept SDS records, files, or arbitrary Script Properties.
  PropertiesService.getScriptProperties().setProperty(
    COLUMN_WIDTHS_PROPERTY,
    JSON.stringify(normalized)
  );
  return normalized;
}

function saveColumnWidths(widths) {
  return saveColumnWidths_(widths);
}

function getColumnVisibility_() {
  var properties = PropertiesService.getScriptProperties();
  var raw = properties.getProperty(COLUMN_VISIBILITY_PROPERTY);
  var saved = {};
  if (raw) {
    try { saved = JSON.parse(raw) || {}; } catch (ignored) { saved = {}; }
  }

  var normalized = {};
  TABLE_COLUMN_INDICES.forEach(function(index) {
    var value = saved[index];
    normalized[index] = value !== false && String(value).toLowerCase() !== "false";
  });
  return normalized;
}

function getColumnVisibility() {
  return getColumnVisibility_();
}

function saveColumnVisibility_(visibility) {
  if (!visibility || typeof visibility !== "object" || Array.isArray(visibility)) {
    throw new Error("รูปแบบค่าการแสดงหัวข้อตารางไม่ถูกต้อง / Invalid table heading settings.");
  }

  var normalized = {};
  var visibleCount = 0;
  TABLE_COLUMN_INDICES.forEach(function(index) {
    var value = visibility[index];
    var visible = value === undefined || value === null || value === "" || value === true || String(value).toLowerCase() === "true";
    normalized[index] = visible;
    if (visible) visibleCount += 1;
  });

  if (!visibleCount) {
    throw new Error("ต้องแสดงอย่างน้อย 1 หัวข้อ / At least one heading must remain visible.");
  }

  PropertiesService.getScriptProperties().setProperty(
    COLUMN_VISIBILITY_PROPERTY,
    JSON.stringify(normalized)
  );
  return normalized;
}

function saveColumnVisibility(visibility) {
  return saveColumnVisibility_(visibility);
}

function getHazardDefinitions_() {
  var raw = PropertiesService.getScriptProperties().getProperty(HAZARD_DEFINITIONS_PROPERTY);
  if (!raw) return [];
  try {
    var saved = JSON.parse(raw);
    return Array.isArray(saved) ? saved : [];
  } catch (ignored) {
    return [];
  }
}

function getHazardDefinitions() {
  return getHazardDefinitions_();
}

function saveHazardDefinitions_(definitions) {
  if (!Array.isArray(definitions) || !definitions.length) {
    throw new Error("รูปแบบค่า Hazard ไม่ถูกต้อง / Invalid hazard settings.");
  }
  if (definitions.length > HAZARD_DEFINITIONS_MAX_ITEMS) {
    throw new Error("จำนวน Hazard มากเกินไป / Too many hazard definitions.");
  }

  var text = function(value, max) {
    return String(value === undefined || value === null ? "" : value).trim().slice(0, max);
  };
  var used = {};
  var normalized = [];
  definitions.forEach(function(item) {
    if (!item || typeof item !== "object") return;
    var name = text(item.name, 120);
    var key = name.toLowerCase();
    if (!name || used[key]) return;
    used[key] = true;
    var image = text(item.image, 300);
    normalized.push({
      id: text(item.id, 80),
      name: name,
      aliases: (Array.isArray(item.aliases) ? item.aliases : [])
        .map(function(alias) { return text(alias, 120); })
        .filter(String)
        .slice(0, 20),
      iconId: text(item.iconId, 40),
      icon: text(item.icon, 16),
      // Only asset paths or https URLs; data: URLs would overflow the property.
      image: /^(assets\/|https:\/\/)/.test(image) ? image : "",
      className: text(item.className, 80)
    });
  });

  if (!normalized.length) {
    throw new Error("ต้องมี Hazard อย่างน้อย 1 รายการ / At least one hazard is required.");
  }
  var json = JSON.stringify(normalized);
  if (Utilities.newBlob(json).getBytes().length > HAZARD_DEFINITIONS_MAX_BYTES) {
    throw new Error("ข้อมูล Hazard ยาวเกินกว่าที่บันทึกได้ / Hazard settings are too large to store.");
  }

  PropertiesService.getScriptProperties().setProperty(HAZARD_DEFINITIONS_PROPERTY, json);
  return normalized;
}

function saveHazardDefinitions(definitions) {
  return saveHazardDefinitions_(definitions);
}

function normalizeColumnWidth_(value, limit) {
  var number = Number(value);
  if (!isFinite(number)) number = limit.fallback;
  return Math.round(Math.min(limit.max, Math.max(limit.min, number)));
}

/**
 * One-time setup. Run this function from the Apps Script editor and approve
 * the requested Sheets/Drive permissions.
 */
function setupSystem() {
  var properties = PropertiesService.getScriptProperties();
  var spreadsheet = SpreadsheetApp.create("SDS Management Database");
  var sheet = spreadsheet.getSheets()[0];
  sheet.setName("SDS");
  sheet.getRange(1, 1, 1, SDS_HEADERS.length).setValues([SDS_HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, SDS_HEADERS.length).setFontWeight("bold");
  sheet.autoResizeColumns(1, SDS_HEADERS.length);

  if (SAMPLE_ROWS.length) {
    sheet.getRange(2, 1, SAMPLE_ROWS.length, SDS_HEADERS.length).setValues(SAMPLE_ROWS);
  }

  var folder = DriveApp.createFolder("SDS PDF Documents");
  var activeEmail = Session.getActiveUser().getEmail();
  var values = {
    SDS_SPREADSHEET_ID: spreadsheet.getId(),
    SDS_FOLDER_ID: folder.getId(),
    ADMIN_EMAILS: activeEmail || Session.getEffectiveUser().getEmail() || ""
  };
  properties.setProperties(values, true);

  return {
    spreadsheetUrl: spreadsheet.getUrl(),
    folderUrl: folder.getUrl(),
    adminEmails: values.ADMIN_EMAILS
  };
}

/**
 * Imports the 98-record catalog published with the GitHub dashboard into the
 * new Google Sheet. PDF links remain public GitHub asset links so viewers do
 * not need to sign in to Google Drive. Thai and English links are kept
 * separately for the QR-language chooser.
 *
 * Run once after setupSystem(). It is idempotent by record id.
 */
function migrateCatalogFromGithub() {
  assertAdmin_();
  var sourceUrl = "https://raw.githubusercontent.com/watanathep8-dotcom/sds-management-dashboard/main/docs/index.html";
  var response = UrlFetchApp.fetch(sourceUrl, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    throw new Error("โหลด catalog จาก GitHub ไม่สำเร็จ: HTTP " + response.getResponseCode());
  }

  var html = response.getContentText();
  var dataMatch = html.match(/const defaultData\s*=\s*(\[[\s\S]*?\])\s*;/);
  if (!dataMatch) throw new Error("ไม่พบ defaultData ใน GitHub catalog");
  var records = JSON.parse(dataMatch[1]);

  var thaiNames = {};
  var namesMatch = html.match(/const thaiChemicalNames = (\{[\s\S]*?\n  \});/);
  if (namesMatch) {
    try { thaiNames = JSON.parse(namesMatch[1]); } catch (ignored) {}
  }

  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  var existing = {};
  if (lastRow >= 2) {
    var current = sheet.getRange(2, 1, lastRow - 1, SDS_HEADERS.length).getValues();
    current.forEach(function(row, index) {
      if (row[0] !== "" && row[0] !== null) existing[String(row[0])] = index + 2;
    });
  }

  var rowsToAppend = [];
  var imported = 0;
  var englishPdfLinks = 0;
  var thaiPdfLinks = 0;
  var updatedAt = new Date().toISOString();

  records.forEach(function(seed) {
    var id = String(seed.id || "");
    if (!id) return;
    var chemicalEnglish = clean_(seed.chemicalEnglish || seed.chemical || seed.chemicalThai);
    var chemicalThai = clean_(seed.chemicalThai || thaiNames[seed.chemical] || chemicalEnglish);
    var englishPdfUrl = clean_(seed.englishPdfUrl || seed.pdfUrl);
    var thaiPdfUrl = clean_(seed.thaiPdfUrl);
    var englishPdfName = clean_(seed.englishPdfName || seed.pdfName);
    var thaiPdfName = clean_(seed.thaiPdfName);
    var language = clean_(seed.language) === "Thai" ? "Thai" : "English";
    var sdsLanguages = normalizeLanguages_(seed.sdsLanguages, language, seed.thaiSds === true);
    var row = [
      id,
      chemicalEnglish,
      clean_(seed.cas),
      clean_(seed.supplier),
      clean_(seed.revision),
      clean_(seed.revisionDate),
      normalizeStatus_(seed.status),
      clean_(seed.signalWord),
      (Array.isArray(seed.hazards) ? seed.hazards : []).map(clean_).filter(String).join("|"),
      "",
      englishPdfName,
      updatedAt,
      clean_(seed.reviewDate),
      seed.thaiSds === true,
      language,
      clean_(seed.flashPoint),
      clean_(seed.emergencyResponse),
      sdsLanguages.join("|"),
      chemicalThai,
      chemicalEnglish,
      "",
      thaiPdfName,
      "",
      englishPdfName,
      englishPdfUrl,
      thaiPdfUrl,
      englishPdfUrl,
      clean_(seed.composition),
      clean_(seed.ratio),
      normalizePpe_(seed.personalProtectiveEquipment).join("|"),
      serializeMeasures_(seed.firstAidMeasures),
      serializeMeasures_(seed.firefightingMeasures),
      Number(seed.sequence) || "",
      clean_(seed.group),
      clean_(seed.responsibleParty),
      clean_(seed.recorder),
      clean_(seed.dateOfUse),
      clean_(seed.registerSds),
      serializeComponents_(seed.components)
    ];

    if (existing[id]) sheet.getRange(existing[id], 1, 1, SDS_HEADERS.length).setValues([row]);
    else rowsToAppend.push(row);
    imported += 1;
    if (englishPdfUrl) englishPdfLinks += 1;
    if (thaiPdfUrl) thaiPdfLinks += 1;
  });

  if (rowsToAppend.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, SDS_HEADERS.length).setValues(rowsToAppend);
  }
  sheet.autoResizeColumns(1, SDS_HEADERS.length);

  return {
    imported: imported,
    appended: rowsToAppend.length,
    englishPdfLinks: englishPdfLinks,
    thaiPdfLinks: thaiPdfLinks,
    totalInSheet: listSds().length,
    source: sourceUrl
  };
}

function listSds() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var values = sheet.getRange(2, 1, lastRow - 1, SDS_HEADERS.length).getValues();
  return values.filter(function(row) {
    return row[0] !== "" && row[0] !== null;
  }).map(rowToObject_);
}

/**
 * Returns an uploaded PDF through the Apps Script web app instead of sending
 * viewers to Google Drive. The web app must be deployed to execute as the
 * owner and allow anonymous access for public viewers.
 */
function getSdsFile(fileId) {
  var id = clean_(fileId);
  if (!id) throw new Error("ไม่พบรหัสไฟล์ SDS");

  var file = DriveApp.getFileById(id);
  var blob = file.getBlob();
  return {
    ok: true,
    name: file.getName(),
    mimeType: blob.getContentType() || "application/pdf",
    base64: Utilities.base64Encode(blob.getBytes())
  };
}

function saveSds(record, fileData, password) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var sheet = getSheet_();
    var requestedId = clean_(record && record.id);
    var existing = findRowById_(sheet, requestedId);
    if (existing) {
      // A client that omits a field (e.g. one the edit form has no input for)
      // must not wipe the stored value.
      var stored = rowToObject_(sheet.getRange(existing, 1, 1, SDS_HEADERS.length).getValues()[0]);
      record = Object.assign({}, record);
      Object.keys(stored).forEach(function(key) {
        if (record[key] === undefined) record[key] = stored[key];
      });
    }
    if (!PUBLIC_SDS_WRITE_ENABLED && !isAdmin_()) {
      throw new Error("ยังไม่เปิดให้ผู้ใช้ทั่วไปบันทึกข้อมูล SDS / Public SDS writing is disabled");
    }
    var normalized = normalizeRecord_(record);
    existing = existing || findRowById_(sheet, normalized.id);
    // Updating an existing record requires the shared password (same as
    // deletion); creating a new record does not. Administrators are exempt.
    if (existing && !isAdmin_()) {
      assertActionPassword_(password);
    }
    var oldRow = existing ? sheet.getRange(existing, 1, 1, SDS_HEADERS.length).getValues()[0] : null;
    var pdfFileId = oldRow ? String(oldRow[9] || "") : "";
    var pdfName = oldRow ? String(oldRow[10] || "") : "";
    var thaiPdfFileId = oldRow ? String(oldRow[20] || "") : "";
    var thaiPdfName = oldRow ? String(oldRow[21] || "") : "";
    var englishPdfFileId = oldRow ? String(oldRow[22] || "") : "";
    var englishPdfName = oldRow ? String(oldRow[23] || "") : "";
    var pdfUrl = oldRow ? String(oldRow[24] || "") : "";
    var thaiPdfUrl = oldRow ? String(oldRow[25] || "") : "";
    var englishPdfUrl = oldRow ? String(oldRow[26] || "") : "";

    if (fileData && fileData.base64) {
      var folder = getFolder_();
      var bytes = Utilities.base64Decode(fileData.base64);
      var blob = Utilities.newBlob(bytes, fileData.mimeType || "application/pdf", fileData.name || "SDS.pdf");
      var file = folder.createFile(blob);
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (sharingError) {
        // Workspace administrators may disable public link sharing.
        console.warn(sharingError.message);
      }

      if (pdfFileId) {
        try { DriveApp.getFileById(pdfFileId).setTrashed(true); } catch (ignored) {}
      }
      pdfFileId = file.getId();
      pdfName = file.getName();
      englishPdfFileId = pdfFileId;
      englishPdfName = pdfName;
      pdfUrl = "";
      englishPdfUrl = "";
    }

    var row = [
      normalized.id,
      normalized.chemical,
      normalized.cas,
      normalized.supplier,
      normalized.revision,
      normalized.revisionDate,
      normalized.status,
      normalized.signalWord,
      normalized.hazards.join("|"),
      pdfFileId,
      pdfName,
      new Date().toISOString(),
      normalized.reviewDate,
      normalized.thaiSds,
      normalized.language,
      normalized.flashPoint,
      normalized.emergencyResponse,
      normalized.sdsLanguages.join("|"),
      normalized.chemicalThai,
      normalized.chemicalEnglish,
      thaiPdfFileId,
      thaiPdfName,
      englishPdfFileId,
      englishPdfName,
      pdfUrl,
      thaiPdfUrl,
      englishPdfUrl,
      normalized.composition,
      normalized.ratio,
      normalized.personalProtectiveEquipment.join("|"),
      serializeMeasures_(normalized.firstAidMeasures),
      serializeMeasures_(normalized.firefightingMeasures),
      normalized.sequence,
      normalized.group,
      normalized.responsibleParty,
      normalized.recorder,
      normalized.dateOfUse,
      normalized.registerSds,
      serializeComponents_(normalized.components)
    ];

    if (existing) {
      sheet.getRange(existing, 1, 1, SDS_HEADERS.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    return rowToObject_(row);
  } finally {
    lock.releaseLock();
  }
}

function deleteSds(id, password) {
  assertActionPassword_(password);
  if (!PUBLIC_SDS_WRITE_ENABLED && !isAdmin_()) {
    throw new Error("ไม่มีสิทธิ์แก้ไขข้อมูล กรุณาเปิด Apps Script ด้วยบัญชีผู้ดูแลระบบ");
  }
  var sheet = getSheet_();
  var rowNumber = findRowById_(sheet, id);
  if (!rowNumber) return { deleted: false };

  var row = sheet.getRange(rowNumber, 1, 1, SDS_HEADERS.length).getValues()[0];
  var fileIds = [row[9], row[20], row[22]].map(function(value) { return String(value || "").trim(); });
  fileIds.filter(function(fileId, index) { return fileId && fileIds.indexOf(fileId) === index; }).forEach(function(fileId) {
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch (ignored) {}
  });
  sheet.deleteRow(rowNumber);
  return { deleted: true };
}

function getSheet_() {
  var id = PropertiesService.getScriptProperties().getProperty("SDS_SPREADSHEET_ID");
  if (!id) throw new Error("ยังไม่ได้ตั้งค่าระบบ กรุณารัน setupSystem() ก่อน");
  var sheet = SpreadsheetApp.openById(id).getSheetByName("SDS");
  if (sheet.getMaxColumns() < SDS_HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), SDS_HEADERS.length - sheet.getMaxColumns());
  }
  var header = sheet.getRange(1, 1, 1, SDS_HEADERS.length).getValues()[0];
  if (header.join("|") !== SDS_HEADERS.join("|")) {
    sheet.getRange(1, 1, 1, SDS_HEADERS.length).setValues([SDS_HEADERS]);
  }
  return sheet;
}

function getFolder_() {
  var id = PropertiesService.getScriptProperties().getProperty("SDS_FOLDER_ID");
  if (!id) throw new Error("ยังไม่ได้ตั้งค่าโฟลเดอร์ PDF กรุณารัน setupSystem() ก่อน");
  return DriveApp.getFolderById(id);
}

function findRowById_(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var target = String(id || "");
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === target) return i + 2;
  }
  return null;
}

function normalizeRecord_(record) {
  var hazards = Array.isArray(record.hazards) ? record.hazards : [];
  hazards = hazards.map(function(item) { return clean_(item); }).filter(String).filter(function(item, index, list) {
    return list.indexOf(item) === index;
  }).slice(0, HAZARD_DEFINITIONS_MAX_ITEMS);
  var thaiSds = record.thaiSds === true || String(record.thaiSds || "").toLowerCase() === "true";
  var language = clean_(record.language);
  if (["English", "Thai"].indexOf(language) === -1) language = thaiSds ? "Thai" : "English";
  var sdsLanguages = normalizeLanguages_(record.sdsLanguages, language, thaiSds);

  return {
    // The original UI renders IDs inside inline onclick handlers, so keep
    // generated IDs numeric for compatibility with that view.
    id: String(record.id || Date.now()),
    chemical: clean_(record.chemicalEnglish || record.chemicalThai || record.chemical),
    chemicalThai: clean_(record.chemicalThai),
    chemicalEnglish: clean_(record.chemicalEnglish),
    cas: clean_(record.cas),
    supplier: clean_(record.supplier),
    revision: clean_(record.revision),
    revisionDate: clean_(record.revisionDate),
    reviewDate: clean_(record.reviewDate),
    status: normalizeStatus_(record.status),
    signalWord: ["Danger", "Warning", ""].indexOf(record.signalWord || "") !== -1 ? (record.signalWord || "") : "",
    hazards: hazards,
    thaiSds: thaiSds,
    language: language,
    flashPoint: clean_(record.flashPoint),
    emergencyResponse: clean_(record.emergencyResponse),
    sdsLanguages: sdsLanguages,
    composition: clean_(record.composition),
    ratio: clean_(record.ratio),
    personalProtectiveEquipment: normalizePpe_(record.personalProtectiveEquipment),
    firstAidMeasures: normalizeMeasures_(record.firstAidMeasures),
    firefightingMeasures: normalizeMeasures_(record.firefightingMeasures),
    sequence: Number(record.sequence) || "",
    group: clean_(record.group),
    responsibleParty: clean_(record.responsibleParty),
    recorder: clean_(record.recorder),
    dateOfUse: clean_(record.dateOfUse),
    registerSds: clean_(record.registerSds),
    components: normalizeComponents_(record.components)
  };
}

function normalizeStatus_(status) {
  var value = clean_(status);
  if (value === "Active") return "Current";
  if (value === "Expiring") return "Review Due";
  if (value === "Expired") return "Update Required";
  return ["Current", "Review Due", "Update Required"].indexOf(value) !== -1 ? value : "Current";
}

function clean_(value) {
  return String(value === undefined || value === null ? "" : value).trim();
}

function normalizeLanguages_(value, language, thaiSds) {
  var values = Array.isArray(value) ? value : String(value || "").split(/[|,]/);
  var result = values.map(function(item) { return clean_(item); }).filter(function(item) {
    return item === "Thai" || item === "English";
  });
  if (thaiSds && result.indexOf("Thai") === -1) result.push("Thai");
  if (!result.length && ["Thai", "English"].indexOf(language) !== -1) result.push(language);
  return result.filter(function(item, index) { return result.indexOf(item) === index; });
}

function normalizePpe_(value) {
  var allowed = ["V", "W", "X", "Y", "Z", "AA"];
  var values = Array.isArray(value) ? value : String(value || "").split(/[|,]/);
  return values.map(function(item) { return clean_(item).toUpperCase(); }).filter(function(item, index, list) {
    return allowed.indexOf(item) !== -1 && list.indexOf(item) === index;
  });
}

function normalizeMeasures_(value) {
  if (!Array.isArray(value)) return [];
  return value.map(function(item) {
    return {
      thai: clean_(item && item.thai),
      english: clean_(item && item.english),
      value: clean_(item && (item.value !== undefined ? item.value : item.detail))
    };
  }).filter(function(item) { return item.thai || item.english || item.value; });
}

function serializeMeasures_(value) {
  return JSON.stringify(normalizeMeasures_(value));
}

function parseMeasures_(value) {
  if (!value) return [];
  try { return normalizeMeasures_(JSON.parse(String(value))); } catch (ignored) { return []; }
}

function normalizeComponents_(value) {
  if (!Array.isArray(value)) return [];
  return value.map(function(item) {
    return {
      composition: clean_(item && item.composition),
      cas: clean_(item && item.cas),
      ratio: clean_(item && item.ratio)
    };
  }).filter(function(item) {
    return item.composition || item.cas || item.ratio;
  });
}

function serializeComponents_(value) {
  return JSON.stringify(normalizeComponents_(value));
}

function parseComponents_(value) {
  if (!value) return [];
  try { return normalizeComponents_(JSON.parse(String(value))); } catch (ignored) { return []; }
}

function rowToObject_(row) {
  var fileId = String(row[9] || "");
  var storedPdfUrl = String(row[24] || "");
  var legacyChemical = String(row[1] || "");
  var chemicalThai = String(row[18] || "");
  var chemicalEnglish = String(row[19] || "");
  if (!chemicalThai && !chemicalEnglish) {
    if (/[\u0E00-\u0E7F]/.test(legacyChemical)) chemicalThai = legacyChemical;
    else chemicalEnglish = legacyChemical;
  }
  var thaiSds = String(row[13] || "").toLowerCase() === "true";
  var language = ["English", "Thai"].indexOf(String(row[14] || "")) !== -1
    ? String(row[14])
    : (thaiSds ? "Thai" : "English");
  return {
    id: String(row[0]),
    chemical: chemicalEnglish || chemicalThai || legacyChemical,
    chemicalThai: chemicalThai,
    chemicalEnglish: chemicalEnglish,
    cas: String(row[2] || ""),
    supplier: String(row[3] || ""),
    revision: String(row[4] || ""),
    revisionDate: formatDateValue_(row[5]),
    status: normalizeStatus_(row[6]),
    signalWord: String(row[7] || ""),
    hazards: String(row[8] || "").split("|").filter(String).slice(0, HAZARD_DEFINITIONS_MAX_ITEMS),
    pdfFileId: fileId,
    pdfName: String(row[10] || ""),
    pdfUrl: storedPdfUrl || (fileId ? "https://drive.google.com/file/d/" + encodeURIComponent(fileId) + "/preview" : ""),
    thaiPdfFileId: String(row[20] || ""),
    thaiPdfName: String(row[21] || ""),
    englishPdfFileId: String(row[22] || ""),
    englishPdfName: String(row[23] || ""),
    thaiPdfUrl: String(row[25] || ""),
    englishPdfUrl: String(row[26] || ""),
    updatedAt: String(row[11] || ""),
    reviewDate: formatDateValue_(row[12]),
    thaiSds: thaiSds,
    language: language,
    flashPoint: String(row[15] || ""),
    emergencyResponse: String(row[16] || ""),
    sdsLanguages: normalizeLanguages_(String(row[17] || "").split("|"), language, thaiSds),
    composition: String(row[27] || ""),
    ratio: String(row[28] || ""),
    personalProtectiveEquipment: normalizePpe_(String(row[29] || "").split("|")),
    firstAidMeasures: parseMeasures_(row[30]),
    firefightingMeasures: parseMeasures_(row[31]),
    sequence: Number(row[32]) || "",
    group: String(row[33] || ""),
    responsibleParty: String(row[34] || ""),
    recorder: String(row[35] || ""),
    dateOfUse: String(row[36] || ""),
    registerSds: String(row[37] || ""),
    components: parseComponents_(row[38])
  };
}

function formatDateValue_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(value);
}

function verifyActionPassword(password) {
  assertActionPassword_(password);
  return { valid: true };
}

function assertActionPassword_(password) {
  var expected = PropertiesService.getScriptProperties().getProperty(ACTION_PASSWORD_PROPERTY) || "";
  if (!expected) {
    throw new Error("ยังไม่ได้ตั้งรหัสใน Script Properties (" + ACTION_PASSWORD_PROPERTY + ") / Action password is not configured.");
  }

  var cache = CacheService.getScriptCache();
  var failures = Number(cache.get(PASSWORD_FAIL_CACHE_KEY) || 0);
  if (failures >= PASSWORD_MAX_FAILURES) {
    throw new Error("ใส่รหัสผิดหลายครั้งเกินไป กรุณารอ 15 นาที / Too many wrong attempts. Try again in 15 minutes.");
  }

  if (String(password || "") !== expected) {
    cache.put(PASSWORD_FAIL_CACHE_KEY, String(failures + 1), PASSWORD_LOCK_SECONDS);
    Utilities.sleep(1000);
    throw new Error("รหัสไม่ถูกต้อง / Incorrect password.");
  }
}

function assertAdmin_() {
  if (!isAdmin_()) {
    throw new Error("ไม่มีสิทธิ์แก้ไขข้อมูล กรุณาเปิด Apps Script ด้วยบัญชีผู้ดูแลระบบ");
  }
}

function isAdmin_() {
  var email = String(Session.getActiveUser().getEmail() || "").toLowerCase();
  var configured = PropertiesService.getScriptProperties().getProperty("ADMIN_EMAILS") || "";
  var admins = configured.split(",").map(function(item) { return item.trim().toLowerCase(); }).filter(String);
  return Boolean(email && admins.indexOf(email) !== -1);
}

function jsonResponse_(value, callback) {
  var json = JSON.stringify(value);
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
