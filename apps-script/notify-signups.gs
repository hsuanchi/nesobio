const SPREADSHEET_ID = "";
const SHEET_NAME = "notify_signups";
const HEADERS = ["timestamp", "email", "source", "pageUrl", "userAgent"];

function doGet() {
  return jsonResponse({ ok: true, message: "NESO BIO notify endpoint is ready." });
}

function doPost(e) {
  const params = e && e.parameter ? e.parameter : {};
  const email = String(params.email || "").trim().toLowerCase();
  const honeypot = String(params.website || "").trim();

  if (honeypot) {
    return jsonResponse({ ok: true });
  }

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return jsonResponse({ ok: false, error: "Invalid email." });
  }

  const sheet = getSheet();
  sheet.appendRow([
    new Date(),
    email,
    String(params.source || ""),
    String(params.pageUrl || ""),
    String(params.userAgent || ""),
  ]);

  return jsonResponse({ ok: true });
}

function getSheet() {
  const spreadsheet = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  const existingHeaders = headerRange.getValues()[0];
  const needsHeaders = HEADERS.some((header, index) => existingHeaders[index] !== header);

  if (needsHeaders) {
    headerRange.setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
