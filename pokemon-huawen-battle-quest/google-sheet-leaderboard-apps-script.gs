const SPREADSHEET_ID = '1TOhQPZhbimDMlyyOBSveTPyTOMQq1yqq7upu2WZoYaA';
const DEFAULT_SHEET_NAME = 'Mario 华文拼音部首闯关';
const GAME_SHEETS = {
  'mario-huawen-pinyin-radical-quest': 'Mario 华文拼音部首闯关',
  'pokemon-huawen-battle-quest': 'Pokemon 华文拼音部首对战',
  'mario-july-bip-word-quest': 'Mario 7月 Y3 2nd&3rd',
  'pokemon-july-bip-word-quest': 'Pokemon 7月 Y3 4th&8月 1st',
  'pokemon-y1-july-2nd-3rd-word-quest': 'Pokemon 7月 Y1 2nd&3rd Week',
  'mario-y1-july-4th-august-1st-word-quest': 'Mario 7月 Y1 4th&8月 1st'
};

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
  'Sheet Name',
  'Answers JSON',
  'User Agent'
];

function doGet() {
  return json_({
    ok: true,
    spreadsheetId: SPREADSHEET_ID,
    sheets: Object.keys(GAME_SHEETS).map(function(gameId) {
      return GAME_SHEETS[gameId];
    }),
    message: 'Mario Huawen score endpoint is ready.',
    serverTime: new Date().toISOString()
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const payload = parsePayload_(e);
    const sheetName = resolveSheetName_(payload);
    const sheet = ensureSheet_(sheetName, STANDARD_HEADERS);
    const totalQuestions = Number(payload.totalQuestions || payload.total || 0);
    const maxScore = Number(payload.maxScore || totalQuestions * 10 || 0);
    const score = Number(payload.score || 0);
    const percent = Number(payload.percent || (maxScore ? Math.round((score / maxScore) * 100) : 0));

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
      payload.sheetName || sheetName,
      JSON.stringify(payload.answers || []),
      payload.userAgent || ''
    ]);

    return json_({ ok: true, sheetName: sheetName });
  } catch (error) {
    return json_({ ok: false, error: String(error && error.message ? error.message : error) });
  } finally {
    lock.releaseLock();
  }
}

function resolveSheetName_(payload) {
  if (payload && GAME_SHEETS[payload.gameId]) {
    return GAME_SHEETS[payload.gameId];
  }
  if (payload && payload.sheetName) {
    const knownSheets = Object.keys(GAME_SHEETS).map(function(gameId) {
      return GAME_SHEETS[gameId];
    });
    if (knownSheets.indexOf(payload.sheetName) !== -1) {
      return payload.sheetName;
    }
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
  const hasHeaders = firstRow.some(function(cell) {
    return cell;
  });
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
    contents.split('&').forEach(function(pair) {
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
