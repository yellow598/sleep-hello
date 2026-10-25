/**
 * 好睡計劃 — Google 試算表同步後端（Google Apps Script 網頁應用程式）
 *
 * 部署：執行身分「我」，存取權「任何人」。
 * 以「同步密碼」(TOKEN) 擋掉沒有密碼的請求；密碼只放在你自己的 Apps Script 裡，
 * 不要把填好真實密碼的版本放上 GitHub。
 *
 * 試算表欄位順序固定如 COLUMNS；日期與時間欄位強制存成純文字，
 * 避免 Google 試算表把 "2026-10-25"、"23:30" 自動轉成日期物件。
 */

const TOKEN = 'REPLACE_WITH_YOUR_SECRET';
const SHEET_NAME = '睡眠日誌';

// [欄位 key, 試算表標題, 是否純文字]
const COLUMNS = [
  ['date',       '日期',            true],
  ['bedTime',    '上床時間',        true],
  ['lightsTime', '熄燈時間',        true],
  ['latency',    '入睡花費(分)',    false],
  ['awakenings', '夜醒次數',        false],
  ['wasoMinutes','夜醒總時間(分)',  false],
  ['wakeTime',   '最後醒來',        true],
  ['outBedTime', '離床時間',        true],
  ['tib',        '在床時數(小時)',  false],
  ['tst',        '睡眠時數(小時)',  false],
  ['se',         '睡眠效率(%)',     false],
  ['quality',    '睡眠品質(1-5)',   false],
  ['notes',      '備註',            true],
  ['caffeine',   '咖啡因',          true],
  ['nap',        '午睡',            true],
  ['stress',     '壓力(1-5)',       false],
  ['energy',     '精力(1-5)',       false],
  ['updatedAt',  '更新時間',        true]
];

function doGet(e) {
  try {
    checkToken_(e.parameter && e.parameter.token);
    return json_({ status: 'ok', logs: readAll_() });
  } catch (err) {
    return json_({ status: 'error', message: String(err.message || err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const body = JSON.parse(e.postData.contents);
    checkToken_(body.token);
    lock.waitLock(15000);

    if (body.action === 'save') {
      (body.logs || []).forEach(upsert_);
    } else if (body.action === 'delete') {
      deleteByDate_(body.date);
    } else if (body.action === 'clear') {
      clearAll_();
    } else {
      throw new Error('unknown action');
    }
    return json_({ status: 'ok' });
  } catch (err) {
    return json_({ status: 'error', message: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

// ---------- 內部函式 ----------

function checkToken_(token) {
  if (!TOKEN || TOKEN === 'REPLACE_WITH_YOUR_SECRET' || token !== TOKEN) {
    throw new Error('unauthorized');
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS.map(c => c[1])]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    // 文字欄位先設成純文字格式，再寫入資料
    COLUMNS.forEach((c, i) => {
      if (c[2]) sheet.getRange(2, i + 1, sheet.getMaxRows() - 1, 1).setNumberFormat('@');
    });
  } else {
    // 標題列和程式裡定義的不同時（例如改了欄位名稱），自動更新成最新的標題
    const headers = COLUMNS.map(c => c[1]);
    const current = sheet.getRange(1, 1, 1, COLUMNS.length).getValues()[0];
    if (current.join('|') !== headers.join('|')) {
      sheet.getRange(1, 1, 1, COLUMNS.length).setValues([headers]).setFontWeight('bold');
    }
  }
  return sheet;
}

// 把儲存格內容一律轉成字串或數字（防止被手動編輯成日期物件）
function cell_(value, isText, key) {
  if (value instanceof Date) {
    const tz = Session.getScriptTimeZone();
    return Utilities.formatDate(value, tz, key === 'date' ? 'yyyy-MM-dd' : 'HH:mm');
  }
  if (value === '' || value === null) return '';
  return isText ? String(value) : (isNaN(Number(value)) ? String(value) : Number(value));
}

function readAll_() {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const rows = sheet.getRange(2, 1, last - 1, COLUMNS.length).getValues();
  return rows
    .filter(r => r[0] !== '')
    .map(r => {
      const obj = {};
      COLUMNS.forEach((c, i) => { obj[c[0]] = cell_(r[i], c[2], c[0]); });
      return obj;
    })
    .sort((a, b) => (a.date < b.date ? 1 : -1)); // 新的日期在前，與網頁一致
}

function findRow_(sheet, date) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const dates = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < dates.length; i++) {
    if (cell_(dates[i][0], true, 'date') === date) return i + 2;
  }
  return -1;
}

// 同一天只會有一列：已存在就覆蓋，不存在就新增
function upsert_(log) {
  if (!log || !/^\d{4}-\d{2}-\d{2}$/.test(log.date)) throw new Error('invalid date');
  const sheet = getSheet_();
  log.updatedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const row = COLUMNS.map(c => (log[c[0]] === undefined || log[c[0]] === null ? '' : cell_(log[c[0]], c[2], c[0])));

  let target = findRow_(sheet, log.date);
  if (target === -1) target = sheet.getLastRow() + 1;
  const range = sheet.getRange(target, 1, 1, COLUMNS.length);
  COLUMNS.forEach((c, i) => { if (c[2]) sheet.getRange(target, i + 1).setNumberFormat('@'); });
  range.setValues([row]);
}

function deleteByDate_(date) {
  const sheet = getSheet_();
  const row = findRow_(sheet, date);
  if (row !== -1) sheet.deleteRow(row);
}

function clearAll_() {
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last > 1) sheet.deleteRows(2, last - 1);
}
