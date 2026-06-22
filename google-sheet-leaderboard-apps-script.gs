const SPREADSHEET_ID = "1TOhQPZhbimDMlyyOBSveTPyTOMQq1yqq7upu2WZoYaA";
const SHEET_NAME = "Pokemon 3D Word Quest";

function doGet() {
  const sheet = getLeaderboardSheet_();
  return json_({
    ok: true,
    spreadsheetId: SPREADSHEET_ID,
    sheet: SHEET_NAME,
    rows: Math.max(0, sheet.getLastRow() - 1),
    message: "Pokemon 3D Word Quest Web App is available."
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const payload = JSON.parse((e.postData && e.postData.contents) || "{}");
    const sheet = getLeaderboardSheet_();
    sheet.appendRow([
      new Date(),
      payload.id || "",
      payload.name || "",
      payload.character || "",
      payload.week || "",
      payload.mode || "",
      Number(payload.score || 0),
      Number(payload.correct || 0),
      Number(payload.total || 0),
      Number(payload.lives || 0),
      payload.grade || "",
      Number(payload.durationSeconds || 0),
      payload.sheetName || "",
      payload.submittedAt || "",
      payload.userAgent || ""
    ]);
    return json_({
      ok: true,
      spreadsheetId: SPREADSHEET_ID,
      sheet: SHEET_NAME
    });
  } catch (error) {
    return json_({
      ok: false,
      spreadsheetId: SPREADSHEET_ID,
      sheet: SHEET_NAME,
      error: String(error)
    });
  } finally {
    lock.releaseLock();
  }
}

function getLeaderboardSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "Submitted At",
      "Submission ID",
      "Name",
      "Character",
      "Week",
      "Mode",
      "Score",
      "Correct",
      "Total",
      "Lives",
      "Grade",
      "Duration Seconds",
      "Client Sheet Name",
      "Client Submitted At",
      "User Agent"
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function json_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
