// ═══════════════════════════════════════════════════════════════
// Workshop Registration — Google Apps Script Backend
// Admins: ykashwin@gmail.com & kkanth09@gmail.com
// ═══════════════════════════════════════════════════════════════

const SPREADSHEET_ID = "17JPzyzjIza0PzR7pnN3GTwMtSug2KVIlS9UdRs0Rop0";
const SHEET_NAME     = "Registrations";
const ADMIN_EMAILS   = ["ykashwin@gmail.com", "kkanth09@gmail.com"];
const ADMIN_PASSWORD = "Password2026";

// ── Access check ──────────────────────────────────────────────
function isAdmin(params) {
  return (params.pwd || "") === ADMIN_PASSWORD;
}

// ── Sheet setup ───────────────────────────────────────────────
function getSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = [
      "Reg Number", "Timestamp", "Name", "Email", "Mobile",
      "Company", "Source", "Feedback", "Photo (Base64)", "Attended", "Attended At"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#2E1F0F")
      .setFontColor("#F5ECD8");
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, headers.length, 150);
    sheet.setColumnWidth(9, 60);
  }
  return sheet;
}

// ── POST: save new registration (open to all) ─────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet();
    sheet.appendRow([
      data.regNumber || "",
      data.timestamp || new Date().toLocaleString(),
      data.name      || "",
      data.email     || "",
      data.mobile    || "",
      data.company   || "",
      data.source    || "",
      data.feedback  || "",
      data.photo     || "",
      "No",
      ""
    ]);
    return jsonResponse({ success: true });
  } catch(err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ── GET: admin-protected actions ──────────────────────────────
function doGet(e) {
  try {
    const action = e.parameter.action || "";

    // All lookup/admin actions require password
    if (action === "getAll" || action === "search" || action === "markAttended") {
      if (!isAdmin(e.parameter)) {
        return jsonResponse({ success: false, error: "UNAUTHORIZED" });
      }
    }

    if (action === "getAll")       return handleGetAll();
    if (action === "search")       return handleSearch(e.parameter);
    if (action === "markAttended") return handleMarkAttended(e.parameter);

    return jsonResponse({ success: false, error: "Unknown action" });
  } catch(err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ── Get all registrations ─────────────────────────────────────
function handleGetAll() {
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();
  const records = [];
  for (let i = 1; i < rows.length; i++) records.push(rowToRecord(rows[i]));
  return jsonResponse({ success: true, records });
}

// ── Search by reg number or name ──────────────────────────────
function handleSearch(params) {
  const regQuery  = (params.reg  || "").trim().toUpperCase();
  const nameQuery = (params.name || "").trim().toLowerCase();
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();
  const matches = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const regMatch  = regQuery  && String(row[0]).toUpperCase().includes(regQuery);
    const nameMatch = nameQuery && String(row[2]).toLowerCase().includes(nameQuery);
    if (regMatch || nameMatch) matches.push(rowToRecord(row));
  }
  return jsonResponse({ success: true, records: matches });
}

// ── Mark / unmark attended ────────────────────────────────────
function handleMarkAttended(params) {
  const regNumber = (params.reg || "").trim().toUpperCase();
  const undo      = params.undo === "true";
  const time      = params.time || new Date().toLocaleString();
  const sheet     = getSheet();
  const rows      = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toUpperCase() === regNumber) {
      sheet.getRange(i + 1, 10).setValue(undo ? "No" : "Yes");
      sheet.getRange(i + 1, 11).setValue(undo ? "" : time);
      const updated = sheet.getRange(i + 1, 1, 1, 11).getValues()[0];
      return jsonResponse({ success: true, record: rowToRecord(updated) });
    }
  }
  return jsonResponse({ success: false, error: "Not found" });
}

// ── Row → record object ───────────────────────────────────────
function rowToRecord(row) {
  return {
    regNumber:  row[0],  timestamp: row[1],  name:       row[2],
    email:      row[3],  mobile:    row[4],  company:    row[5],
    source:     row[6],  feedback:  row[7],  photo:      row[8],
    attended:   row[9],  attendedAt:row[10]
  };
}

// ── JSON response ─────────────────────────────────────────────
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
