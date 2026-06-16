const SPREADSHEET_ID = '1iKBXYvK-wJ43oVpANLpT5R8wqXvc2VbGjrGqCp6zk7o';

const SHEETS = {
  records: 'Records',
  actionList: 'ActionList',
  settings: 'Settings',
};

function doGet(e) {
  const action = (e.parameter.action || 'read').trim();
  if (action !== 'read') {
    return jsonOutput({ok: false, error: 'Unsupported action: ' + action});
  }
  return jsonOutput(readState());
}

function doPost(e) {
  try {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    switch (body.action) {
      case 'upsertRecords':
        return jsonOutput(upsertJsonRows(SHEETS.records, body.records || []));
      case 'replaceActionList':
        return jsonOutput(replaceJsonRows(SHEETS.actionList, body.items || []));
      case 'saveSetting':
        return jsonOutput(saveSetting(body.key, body.value));
      default:
        return jsonOutput({ok: false, error: 'Unsupported action: ' + body.action});
    }
  } catch (err) {
    return jsonOutput({ok: false, error: err.message});
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
  return [item['문서번호'] || item['臾몄꽌踰덊샇'] || '', item['순번'] || item['?쒕쾲'] || '1'].join('_');
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

function jsonOutput(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
