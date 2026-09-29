// IJGB Earth Science Questionnaire: receives answers and adds one row per submission.
// Setup: paste this into Extensions > Apps Script in your Google Sheet, then Deploy > New deployment > Web app.

const SUBMIT_CODE = "ijgb-fall26"; // Must match SUBMIT_CODE in index.html
const SHEET_NAME = "Responses";

const HEADERS = [
  "submitted_at", "name",
  "learn_first", "learn_stuck", "learn_instr", "learn_focus",
  "rate_air", "rate_water", "rate_land", "rate_life", "rate_ice",
  "top1", "top2", "connect", "lens1", "lens2",
  "watched1", "read1", "summary1", "surprise1", "question1", "answer1",
  "watched2", "read2", "summary2", "surprise2", "question2", "answer2",
  "final_choice", "final_why"
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.code !== SUBMIT_CODE) return out("rejected");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow(HEADERS.map(h => clean(data[h])));
    return out("ok");
  } catch (err) {
    return out("error");
  } finally {
    lock.releaseLock();
  }
}

// Stops text that starts with = + - @ from running as a formula in the sheet.
function clean(v) {
  const s = v === undefined || v === null ? "" : String(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function out(msg) {
  return ContentService.createTextOutput(msg);
}
