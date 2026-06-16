const SPREADSHEET_ID = '1iKBXYvK-wJ43oVpANLpT5R8wqXvc2VbGjrGqCp6zk7o';

const SHEETS = {
  records: 'Records',
  actionList: 'ActionList',
  settings: 'Settings',
};

function doGet(e) {
  const action = ((e.parameter && e.parameter.action) || '').trim();
  if (!action) {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('공정부적합 대시보드')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (action !== 'read') {
    return jsonOutput({ok: false, error: 'Unsupported action: ' + action}, e.parameter.callback);
  }
  return jsonOutput(readState(), e.parameter.callback);
}

function doPost(e) {
  try {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    return jsonOutput(handleWriteAction(body));
  } catch (err) {
    return jsonOutput({ok: false, error: err.message});
  }
}

function apiRead() {
  return readState();
}

function apiPost(body) {
  return handleWriteAction(body || {});
}

function handleWriteAction(body) {
  switch (body.action) {
    case 'upsertRecords':
      return upsertJsonRows(SHEETS.records, body.records || []);
    case 'replaceActionList':
      return replaceJsonRows(SHEETS.actionList, body.items || []);
    case 'saveSetting':
      return saveSetting(body.key, body.value);
    default:
      return {ok: false, error: 'Unsupported action: ' + body.action};
  }
}

function readState() {
  const settings = readSettings();
  return {
    ok: true,
    records: readJsonRows(SHEETS.records),
    actionItems: readJsonRows(SHEETS.actionList),
    note: settings.al_note_v1 || '',
    updatedAt: new Date().toISOString(),
  };
}

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Missing sheet: ' + name);
  return sheet;
}

function recordId(item) {
  return [item['문서번호'] || '', item['순번'] || '1'].join('_');
}

function shouldScaleLossAmount(item, amount) {
  const docNo = String(item['문서번호'] || '');
  return amount > 0 && amount < 1000 && /^부적합보고서\d{2}-/.test(docNo);
}

function normalizeRecordAmount(item) {
  const copy = Object.assign({}, item);
  if ('금액' in copy) {
    const amount = Math.round(Number(copy['금액']) || 0);
    copy['금액'] = shouldScaleLossAmount(copy, amount) ? amount * 10000 : amount;
  }
  return copy;
}

function itemId(item) {
  return item.id || Utilities.getUuid();
}

function readJsonRows(sheetName) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  return values
    .filter(row => row[0] && row[1])
    .map(row => JSON.parse(row[1]))
    .map(item => sheetName === SHEETS.records ? normalizeRecordAmount(item) : item);
}

function upsertJsonRows(sheetName, items) {
  const sheet = getSheet(sheetName);
  const ids = existingIdMap(sheet);
  const now = new Date().toISOString();
  const appends = [];
  let saved = 0;

  items.forEach(item => {
    const normalized = sheetName === SHEETS.records ? normalizeRecordAmount(item) : item;
    const id = recordId(normalized);
    if (!id || id === '_1') return;
    const row = [id, JSON.stringify(normalized), now, 'dashboard'];
    if (ids[id]) {
      sheet.getRange(ids[id], 1, 1, 4).setValues([row]);
    } else {
      appends.push(row);
    }
    saved++;
  });

  if (appends.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appends.length, 4).setValues(appends);
  }
  return {ok: true, saved};
}

function replaceJsonRows(sheetName, items) {
  const sheet = getSheet(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  if (!items.length) return {ok: true, saved: 0};

  const now = new Date().toISOString();
  const rows = items.map(item => {
    const id = itemId(item);
    item.id = id;
    return [id, JSON.stringify(item), now, 'dashboard'];
  });
  sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  return {ok: true, saved: rows.length};
}

function existingIdMap(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return ids.reduce((acc, row, index) => {
    if (row[0]) acc[row[0]] = index + 2;
    return acc;
  }, {});
}

function readSettings() {
  const sheet = getSheet(SHEETS.settings);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  return rows.reduce((acc, row) => {
    if (row[0]) acc[row[0]] = row[1] || '';
    return acc;
  }, {});
}

function saveSetting(key, value) {
  if (!key) throw new Error('Missing setting key');
  const sheet = getSheet(SHEETS.settings);
  const ids = existingSettingMap(sheet);
  const row = [key, value || '', new Date().toISOString()];
  if (ids[key]) {
    sheet.getRange(ids[key], 1, 1, 3).setValues([row]);
  } else {
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, 3).setValues([row]);
  }
  return {ok: true, saved: 1};
}

function existingSettingMap(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return keys.reduce((acc, row, index) => {
    if (row[0]) acc[row[0]] = index + 2;
    return acc;
  }, {});
}

function jsonOutput(payload, callback) {
  const json = JSON.stringify(payload);
  if (callback) {
    return ContentService
      .createTextOutput(String(callback).replace(/[^\w.$]/g, '') + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
