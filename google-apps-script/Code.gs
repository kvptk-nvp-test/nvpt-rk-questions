const SPREADSHEET_ID = "1l9sQU2Hz5F9rsfK_N-r4K3LlY7CTIiRMyEnm06WPM8w";
const DEFAULT_THEME_PREFIX = "Тема ";

function doGet(e) {
  return handleRequest_(e);
}

function doPost(e) {
  return handleRequest_(e);
}

function handleRequest_(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    const params = extractParams_(e);
    const themeId = cleanString_(params.theme_id);
    const themeCode = buildThemeCode_(params.theme_code, themeId);
    const targetSheetName = sanitizeSheetName_(themeCode);
    const correctCount = toNumber_(params.correct_count);
    const totalQuestions = toNumber_(params.total_questions);
    const percent = toNumber_(params.score_percent);

    if (!themeId || !themeCode) {
      return textResponse_("missing theme_id or theme_code");
    }

    const row = [
      new Date(),
      buildFullName_(params.name, params.surname),
      buildScoreText_(correctCount, totalQuestions),
      percent
    ];

    const spreadsheet = getSpreadsheet_();
    appendRow_(spreadsheet, targetSheetName, row);
    normalizeThemeSheetOrder_(spreadsheet);

    return textResponse_("ok");
  } catch (error) {
    return textResponse_(String(error && error.message ? error.message : error));
  } finally {
    lock.releaseLock();
  }
}

function extractParams_(e) {
  const queryParams = e && e.parameter ? e.parameter : {};
  const postData = e && e.postData && e.postData.contents ? e.postData.contents : "";

  if (!postData) {
    return queryParams;
  }

  try {
    const json = JSON.parse(postData);
    return Object.assign({}, queryParams, json);
  } catch (error) {
    return queryParams;
  }
}

function getSpreadsheet_() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    return active;
  }

  throw new Error("Spreadsheet not found. Bind the script to a sheet or set SPREADSHEET_ID.");
}

function appendRow_(spreadsheet, sheetName, row) {
  const sheet = getOrCreateSheet_(spreadsheet, sheetName);
  ensureHeader_(sheet);
  sheet.appendRow(row);
}

function getOrCreateSheet_(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  return sheet;
}

function ensureHeader_(sheet) {
  const header = [[
    "Дата и время",
    "Имя и фамилия",
    "Верных",
    "Баллы (%)"
  ]];

  sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), header[0].length)).clearContent();
  sheet.getRange(1, 1, 1, header[0].length).setValues(header);
  sheet.getRange(1, 1, 1, header[0].length).setFontWeight("bold");
  sheet.setFrozenRows(1);
  sheet.getRange("A:A").setNumberFormat("dd.MM.yyyy HH:mm:ss");
}

function buildThemeCode_(rawThemeCode, themeId) {
  if (!themeId) {
    return "";
  }

  const normalizedThemeId = normalizeThemeId_(themeId) || cleanString_(rawThemeCode);
  return DEFAULT_THEME_PREFIX + normalizedThemeId;
}

function normalizeThemeId_(themeId) {
  const match = String(themeId).match(/\d+/);
  return match ? String(Number(match[0])) : cleanString_(themeId);
}

function sanitizeSheetName_(value) {
  return cleanString_(value)
    .replace(/[\[\]\*\/\\\?\:]/g, "_")
    .slice(0, 100);
}

function normalizeThemeSheetOrder_(spreadsheet) {
  const themeSheets = spreadsheet
    .getSheets()
    .map(function(sheet) {
      return {
        sheet: sheet,
        themeNumber: extractThemeNumber_(sheet.getName())
      };
    })
    .filter(function(item) {
      return item.themeNumber !== null;
    })
    .sort(function(left, right) {
      return left.themeNumber - right.themeNumber;
    });

  themeSheets.forEach(function(item, index) {
    spreadsheet.setActiveSheet(item.sheet);
    spreadsheet.moveActiveSheet(index + 1);
  });
}

function extractThemeNumber_(sheetName) {
  const match = cleanString_(sheetName).match(/^Тема\s+(\d+)$/i);
  return match ? Number(match[1]) : null;
}

function buildFullName_(name, surname) {
  return [cleanString_(name), cleanString_(surname)].filter(Boolean).join(" ");
}

function buildScoreText_(correctCount, totalQuestions) {
  return correctCount + "/" + totalQuestions;
}

function cleanString_(value) {
  return String(value || "").trim();
}

function toNumber_(value) {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : 0;
}

function textResponse_(text) {
  return ContentService
    .createTextOutput(text)
    .setMimeType(ContentService.MimeType.TEXT);
}
