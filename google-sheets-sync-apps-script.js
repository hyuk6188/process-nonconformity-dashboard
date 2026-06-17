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
  if (action.toLowerCase().startsWith('repairlosstotal')) {
    return jsonOutput(
      repairRecordsTotalForDate(e.parameter.date || '2026-06-15', Number(e.parameter.amount || 1679415)),
      e.parameter.callback
    );
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
    .map(row => JSON.parse(row[1]));
}

function upsertJsonRows(sheetName, items) {
  const sheet = getSheet(sheetName);
  const ids = existingIdMap(sheet);
  const now = new Date().toISOString();
  const appends = [];
  let saved = 0;

  items.forEach(item => {
    const id = recordId(item);
    if (!id || id === '_1') return;
    const row = [id, JSON.stringify(item), now, 'dashboard'];
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

function repairRecordsTotalForDate(dateText, targetAmount) {
  const sheet = getSheet(SHEETS.records);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {ok: false, error: 'No records'};

  const range = sheet.getRange(2, 1, lastRow - 1, 4);
  const rows = range.getValues();
  const matches = [];
  let currentTotal = 0;

  rows.forEach((row, index) => {
    if (!row[0] || !row[1]) return;
    const item = JSON.parse(row[1]);
    if (item['발생일'] !== dateText) return;
    const amount = Math.round(Number(item['금액']) || 0);
    matches.push({index, item, amount});
    currentTotal += amount;
  });

  if (!matches.length) return {ok: false, error: 'No records for ' + dateText};
  if (!currentTotal) return {ok: false, error: 'Current total is zero'};

  const target = Math.round(Number(targetAmount) || 0);
  const scaled = matches.map(match => {
    const exact = match.amount * target / currentTotal;
    return {...match, next: Math.floor(exact), remainder: exact - Math.floor(exact)};
  });
  let remainder = target - scaled.reduce((sum, match) => sum + match.next, 0);
  scaled.sort((a, b) => b.remainder - a.remainder);
  scaled.forEach(match => {
    if (remainder > 0) {
      match.next++;
      remainder--;
    }
  });

  const now = new Date().toISOString();
  scaled.forEach(match => {
    match.item['금액'] = match.next;
    rows[match.index][1] = JSON.stringify(match.item);
    rows[match.index][2] = now;
    rows[match.index][3] = 'dashboard-repair';
  });
  range.setValues(rows);

  return {
    ok: true,
    date: dateText,
    records: matches.length,
    previousTotal: currentTotal,
    targetTotal: target,
    savedTotal: scaled.reduce((sum, match) => sum + match.next, 0),
  };
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
