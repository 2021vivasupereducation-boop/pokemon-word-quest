const SPREADSHEET_ID = '1TOhQPZhbimDMlyyOBSveTPyTOMQq1yqq7upu2WZoYaA';
const DEFAULT_SHEET_NAME = 'Pokemon 3D Word Quest';
const NEW_SHEET_NAME = 'Pokemon 7月 Y3 4th&8月 1st';

const STANDARD_HEADERS = [
  'Server Timestamp',
  'Client Time',
  'Student Name',
  'Character',
  'Week',
  'Mode',
  'Score',
  'Max Score',
  'Percent',
  'Correct',
  'Wrong',
  'Lives Left',
  'Total Questions',
  'Game ID',
  'Game Title',
  'User Agent'
];

const LEGACY_HEADERS = [
  'Submitted At',
  'Submission ID',
  'Name',
  'Character',
  'Week',
  'Mode',
  'Score',
  'Correct',
  'Total',
  'Lives',
  'Grade',
  'Duration Seconds',
  'Client Sheet Name',
  'Client Submitted At',
  'User Agent'
];

function doGet() {
  return json_({
    ok: true,
    spreadsheetId: SPREADSHEET_ID,
    sheets: [DEFAULT_SHEET_NAME, NEW_SHEET_NAME],
    message: 'Pokemon Word Quest score endpoint is ready.',
    serverTime: new Date().toISOString()
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const payload = parsePayload_(e);
    const sheetName = resolveSheetName_(payload);
    const useLegacyFormat = sheetName === DEFAULT_SHEET_NAME && payload.gameId !== 'pokemon-july-bip-word-quest';
    const sheet = ensureSheet_(sheetName, useLegacyFormat ? LEGACY_HEADERS : STANDARD_HEADERS);

    if (useLegacyFormat) {
      sheet.appendRow([
        new Date(),
        payload.id || '',
        payload.name || payload.studentName || '',
        payload.character || '',
        payload.week || '',
        payload.mode || '',
        Number(payload.score || 0),
        Number(payload.correct || 0),
        Number(payload.total || payload.totalQuestions || 0),
        Number(payload.lives || payload.livesLeft || 0),
        payload.grade || '',
        Number(payload.durationSeconds || 0),
        payload.sheetName || '',
        payload.submittedAt || payload.clientTime || '',
        payload.userAgent || ''
      ]);
    } else {
      const maxScore = Number(payload.maxScore || (Number(payload.total || payload.totalQuestions || 0) * 10) || 0);
      const score = Number(payload.score || 0);
      const percent = Number(payload.percent || (maxScore ? Math.round((score / maxScore) * 100) : 0));
      const totalQuestions = Number(payload.totalQuestions || payload.total || 0);
      sheet.appendRow([
        new Date(),
        payload.clientTime || payload.submittedAt || '',
        payload.studentName || payload.name || '',
        payload.character || '',
        payload.week || '',
        payload.mode || '',
        score,
        maxScore,
        percent,
        Number(payload.correct || 0),
        Number(payload.wrong || Math.max(0, totalQuestions - Number(payload.correct || 0))),
        Number(payload.livesLeft || payload.lives || 0),
        totalQuestions,
        payload.gameId || '',
        payload.gameTitle || '',
        payload.userAgent || ''
      ]);
    }

    return json_({ ok: true, sheetName: sheetName });
  } catch (error) {
    return json_({ ok: false, error: String(error && error.message ? error.message : error) });
  } finally {
    lock.releaseLock();
  }
}

function resolveSheetName_(payload) {
  if (payload.gameId === 'pokemon-july-bip-word-quest' || payload.sheetName === NEW_SHEET_NAME) {
    return NEW_SHEET_NAME;
  }
  return DEFAULT_SHEET_NAME;
}

function ensureSheet_(sheetName, headers) {
  if (!SPREADSHEET_ID || SPREADSHEET_ID === 'PASTE_SPREADSHEET_ID_HERE') {
    throw new Error('Set SPREADSHEET_ID before publishing this Apps Script web app.');
  }

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const hasHeaders = firstRow.some((cell) => cell);
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function parsePayload_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return {};
  }

  const contents = e.postData.contents;
  try {
    return JSON.parse(contents);
  } catch (error) {
    const params = {};
    contents.split('&').forEach((pair) => {
      const parts = pair.split('=');
      if (parts[0]) {
        params[decodeURIComponent(parts[0])] = decodeURIComponent(parts.slice(1).join('=') || '');
      }
    });
    return params;
  }
}

function json_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
