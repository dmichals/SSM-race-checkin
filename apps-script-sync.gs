/**
 * SSM Race Check-In Sync Endpoint
 * -------------------------------
 * Paste this into the Script Editor of your Sailing Series Manager
 * spreadsheet (Extensions > Apps Script), then deploy as a Web App.
 *
 * Deploy steps:
 *  1. Deploy > New deployment > type: Web app
 *  2. Execute as: Me
 *  3. Who has access: Anyone (needed so the page can POST without login)
 *  4. Click Deploy, copy the /exec URL
 *  5. Paste that URL into SYNC_URL in race-checkin.html
 *
 * Adjust SHEET_NAME and the column order below to match your
 * actual SSM check-in / results sheet.
 */

const SHEET_NAME = "RaceLog";   // <-- change to your sheet's tab name
const HEADERS = ["Sail #", "Boat Name", "Check-in Time", "Finish Time", "Status", "Start Area", "Weather Mark", "JAM Course", "Spinnaker Course", "Sec4 Start", "Sec3 Start", "Sec2 Start", "Committee Boat", "Recorded By", "Synced At"];

// --- Roster settings ---
// Sheet that holds your boat list, used to preload Sail # + Boat Name
// so the check-in page can autocomplete offline. Adjust column numbers
// (1 = A, 2 = B, ...) to match your actual roster layout.
const ROSTER_SHEET_NAME = "Boats";    // <-- change to your roster tab name
const ROSTER_SAIL_COL = 1;            // column A: Sail #
const ROSTER_NAME_COL = 2;            // column B: Boat Name
const ROSTER_HEADER_ROWS = 1;         // number of header rows to skip

/**
 * Serves the boat roster as JSON so the check-in page can cache it
 * locally (call this once while online, e.g. before leaving the dock).
 * Test in a browser: <your /exec url>?action=roster
 */
function doGet(e) {
  if (e.parameter.action === "roster") {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ROSTER_SHEET_NAME);
    if (!sheet) {
      return ContentService.createTextOutput(
        JSON.stringify({ ok: false, error: "Roster sheet '" + ROSTER_SHEET_NAME + "' not found" })
      ).setMimeType(ContentService.MimeType.JSON);
    }
    const lastRow = sheet.getLastRow();
    const numRows = lastRow - ROSTER_HEADER_ROWS;
    let roster = [];
    if (numRows > 0) {
      const sailVals = sheet.getRange(ROSTER_HEADER_ROWS + 1, ROSTER_SAIL_COL, numRows, 1).getValues();
      const nameVals = sheet.getRange(ROSTER_HEADER_ROWS + 1, ROSTER_NAME_COL, numRows, 1).getValues();
      roster = sailVals
        .map((row, i) => ({ sail: String(row[0]).trim(), name: String(nameVals[i][0]).trim() }))
        .filter(b => b.sail !== "");
    }
    return ContentService.createTextOutput(
      JSON.stringify({ ok: true, roster: roster })
    ).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(
    JSON.stringify({ ok: false, error: "Unknown GET action" })
  ).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const records = payload.records || [];

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADERS);
    } else if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
    }

    const now = new Date();
    const rows = records.map(r => [
      r.boat || "",
      r.boatName || "",
      r.checkinTime ? new Date(r.checkinTime) : "",
      r.finishTime ? new Date(r.finishTime) : "",
      r.status || "",
      r.startArea || "",
      r.weatherMark || "",
      r.jamCourse || "",
      r.spinCourse || "",
      r.section4Start || "",
      r.section3Start || "",
      r.section2Start || "",
      r.committeeBoat || "",
      r.recordedBy || "",
      now
    ]);

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
    }

    return ContentService.createTextOutput(
      JSON.stringify({ ok: true, written: rows.length })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ ok: false, error: err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// Optional: quick manual test from the Script Editor.
function testDoPost() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        records: [
          { boat: "42", checkinTime: new Date().toISOString(), finishTime: null }
        ]
      })
    }
  };
  Logger.log(doPost(fakeEvent).getContent());
}
