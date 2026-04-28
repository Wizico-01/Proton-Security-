// ============================================================
// PROTON SECURITY SERVICES LIMITED — SITREP RECEIVER
// Paste into Google Apps Script:
// Extensions > Apps Script > Delete all > Paste > Save
// Then run initialSetup() ONCE, then Deploy as Web App
// ============================================================

var HEADERS = [
  "DATE", "TIME", "SHIFT", "LOCATION", "ZONE", "AREA", "OC",
  "SUBMITTED BY", "RANK", "EXPECTED GUARDS", "ACTUAL GUARDS",
  "SHORT-MANNING", "ABSENT GUARD(S)", "REASON(S)", "INCIDENT / REMARKS"
];

function doPost(e) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("DAILY SITREP");

    if (!sheet) {
      sheet = ss.insertSheet("DAILY SITREP");
      setupHeaders(sheet);
    }

    var data     = JSON.parse(e.postData.contents);
    var absentees = [];
    var reasons   = [];

    try {
      var ab = JSON.parse(data.absentees || "[]");
      ab.forEach(function(a) {
        absentees.push(a.name);
        reasons.push(a.reason);
      });
    } catch(err) {}

    sheet.appendRow([
      data.date,
      data.time,
      data.shift       || "Day",
      data.location,
      data.zone        || "",
      data.area        || "",
      data.oc          || "",
      data.guardName,
      data.rank        || "",
      data.expected,
      data.actual,
      data.shortmanning,
      absentees.join(", ") || "None",
      reasons.join(", ")   || "N/A",
      data.incident    || ""
    ]);

    // Color short-manning cell
    var lastRow   = sheet.getLastRow();
    var shortCell = sheet.getRange(lastRow, 12);
    if (data.shortmanning > 0) {
      shortCell.setBackground("#fdecea").setFontColor("#a93226").setFontWeight("bold");
    } else {
      shortCell.setBackground("#e8f9ef").setFontColor("#1a7a40").setFontWeight("bold");
    }

    // Alternate row shading
    if (lastRow % 2 === 0) {
      sheet.getRange(lastRow, 1, 1, HEADERS.length).setBackground("#f8f9ff");
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function setupHeaders(sheet) {
  var hr = sheet.getRange(1, 1, 1, HEADERS.length);
  hr.setValues([HEADERS])
    .setBackground("#1a2a6c")
    .setFontColor("#D4A017")
    .setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center");
  sheet.setFrozenRows(1);

  // Column widths
  var widths = [90, 70, 80, 200, 120, 120, 150, 180, 100, 130, 110, 120, 200, 150, 200];
  widths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
}

// Run this ONCE manually before deploying
function initialSetup() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DAILY SITREP") || ss.insertSheet("DAILY SITREP");
  sheet.clearContents();
  setupHeaders(sheet);
  Logger.log("✓ Proton Sitrep sheet is ready!");
}
