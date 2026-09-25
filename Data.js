// ============================================================
// Data.gs - データアクセス + 外部スプレッドシート
// ============================================================

// エラーをログ + 監査ログ + 直近エラーを ScriptProperties に記録
function recordDataError_(source, err) {
  var msg = err && err.message ? err.message : String(err);
  Logger.log(source + ' エラー: ' + msg);
  try { writeAuditLog_('error.data', source, msg); } catch (e) {}
  try { PropertiesService.getScriptProperties().setProperty('LAST_ERROR', new Date().toISOString() + ' ' + source + ': ' + msg); } catch (e) {}
}

function getActiveWorkflows_() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_SETTINGS);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
    var wfs = [];
    // status === '有効' のみ(archived や '無効' は除外)
    data.forEach(function(r) { if (r[7] === '有効') wfs.push({ id: r[0], name: r[1], type: r[2], execSpace: r[3], targetSpace: r[4], creator: r[5], created: r[6], status: r[7], approvers: r[9] || '', externalSheet: r[10] || '' }); });
    return wfs;
  } catch (err) { recordDataError_('getActiveWorkflows_', err); return []; }
}

function getWorkflowById_(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_SETTINGS);
    if (!sheet || sheet.getLastRow() < 2) return null;
    var lastCol = Math.max(sheet.getLastColumn(), 15);
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
    // archived は返さない(復元 API 経由でのみ archived を扱う)
    for (var i = 0; i < data.length; i++) { if (data[i][0] === id && data[i][7] !== 'archived') return { id: data[i][0], name: data[i][1], type: data[i][2], execSpace: data[i][3], targetSpace: data[i][4], creator: data[i][5], created: data[i][6], status: data[i][7], approvers: data[i][9] || '', externalSheet: data[i][10] || '', approverUserIds: data[i][12] || '', mentionText: data[i][13] || '', collaborators: data[i][14] || '' }; }
    return null;
  } catch (err) { recordDataError_('getWorkflowById_', err); return null; }
}

// includeDeleted=true を渡すと deleted=true の行も含める(管理画面の差分表示用)。
// デフォルトは deleted を除外し、有効フィールドのみ返す。
function getWorkflowFields_(wid, includeDeleted) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureFieldSchema_(ss);
    var sheet = ss.getSheetByName(SHEET_FIELDS);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
    var fields = [];
    data.forEach(function(r) {
      if (r[0] !== wid) return;
      var isDel = r[7] === true || r[7] === 'TRUE' || r[7] === 'true';
      if (isDel && !includeDeleted) return;
      fields.push({
        name: r[1], type: r[2], options: r[3],
        required: r[4] === true || r[4] === 'TRUE' || r[4] === 'true',
        order: r[5], fieldId: r[6], deleted: isDel,
        dataCol: Number(r[8]) || 0
      });
    });
    fields.sort(function(a, b) { return a.order - b.order; });
    return fields;
  } catch (err) { recordDataError_('getWorkflowFields_', err); return []; }
}

function getWorkflowButtons_(wid) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_BUTTONS);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
    var btns = [];
    data.forEach(function(r) { if (r[0] === wid) btns.push({ name: r[1], status: r[2], color: r[3], order: r[4], threadReply: r[5] === true || r[5] === 'TRUE' || r[5] === 'true', replyMsg: r[6] || '' }); });
    btns.sort(function(a, b) { return a.order - b.order; });
    return btns;
  } catch (err) { recordDataError_('getWorkflowButtons_', err); return []; }
}

function writeToExternalSheet_(wf, rowData) {
  if (!wf.externalSheet) return;
  try {
    var extId = wf.externalSheet.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (!extId) return;
    var extSs = SpreadsheetApp.openById(extId[1]);
    var extSheet = extSs.getSheetByName(wf.name);
    if (!extSheet) { extSheet = extSs.insertSheet(wf.name, extSs.getSheets().length); extSheet.appendRow(rowData.headers); }
    extSheet.appendRow(rowData.values);
  } catch (err) { recordDataError_('writeToExternalSheet_', err); }
}

function updateExternalSheet_(wf, sheetName, rowIndex, col, value) {
  if (!wf.externalSheet) return;
  try {
    var extId = wf.externalSheet.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (!extId) return;
    var extSs = SpreadsheetApp.openById(extId[1]);
    var extSheet = extSs.getSheetByName(sheetName);
    if (extSheet && rowIndex <= extSheet.getLastRow()) extSheet.getRange(rowIndex, col).setValue(value);
  } catch (err) { recordDataError_('updateExternalSheet_', err); }
}
