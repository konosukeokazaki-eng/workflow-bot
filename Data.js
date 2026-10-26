// ============================================================
// Data.gs - データアクセス + 外部スプレッドシート
// ============================================================

// カンマ区切り送信先スペースIDをパースして配列化する。空要素/空白は除外。
// 送信先スペース(targetSpace)/実行スペース(execSpace)共通で利用する。
// 単一ID("spaces/AAA")でも配列(["spaces/AAA"])で返す。空文字なら [] を返す。
function parseSpaceIds_(raw) {
  return String(raw || '').split(',').map(function(s){ return s.trim(); }).filter(function(s){ return s; });
}

// プルダウン/チェックボックスの単一選択肢を解析する。
// 入力: "営業部=sales@example.com;mgr@example.com" または "営業部"
// 出力: { label: '営業部', emails: ['sales@example.com', 'mgr@example.com'] }
// '=' を含まなければ emails=[]（従来通りメンションなし）。
// '=' 以降が空文字でも emails=[] 扱い。
function parseDropdownOption_(raw) {
  var s = String(raw || '').trim();
  if (!s) return { label: '', emails: [] };
  var eqIdx = s.indexOf('=');
  if (eqIdx < 0) return { label: s, emails: [] };
  var label = s.substring(0, eqIdx).trim();
  var tail = s.substring(eqIdx + 1).trim();
  if (!tail) return { label: label, emails: [] };
  // メール複数はセミコロン区切り。カンマは選択肢区切りと衝突するため使えない。
  var emails = tail.split(';').map(function(e){ return e.trim(); }).filter(function(e){ return e && e.indexOf('@') > 0; });
  return { label: label, emails: emails };
}

// options 文字列を全選択肢に分解。dropdown/checkbox 共通。
// 返り値: [{ label, emails }]
function parseFieldOptionsForMentions_(options) {
  return String(options || '').split(',').map(function(o){ return parseDropdownOption_(o); }).filter(function(p){ return p.label; });
}

// 従来互換用: options から label のみの配列を返す（UI/プレビュー表示用）
function extractLabelsFromOptions_(options) {
  return parseFieldOptionsForMentions_(options).map(function(p){ return p.label; });
}

// email のリスト → Google Chat の userId("users/XXX" ではなく生ID "XXX") 配列に解決する。
// 承認者ピッカー (api_getSpaceMembers) と同じ経路: Admin SDK Directory API の users.get(email) を使う。
// 解決できなかった email は捨てる(ログ記録のみ)。結果は ScriptCache に 6h キャッシュする。
// 戻り値: ['UUU...', 'VVV...'] (重複排除済み)
function resolveEmailsToUserIds_(emails) {
  var uniq = [];
  var seen = {};
  (emails || []).forEach(function(e){ var k = String(e||'').trim().toLowerCase(); if (k && !seen[k]) { seen[k] = true; uniq.push(k); } });
  if (!uniq.length) return [];
  var cache = CacheService.getScriptCache();
  var result = [];
  var missing = [];
  uniq.forEach(function(em){
    var cv = cache.get('e2uid_' + em);
    if (cv) { if (cv !== '-') result.push(cv); }
    else missing.push(em);
  });
  if (missing.length) {
    var token;
    try { token = ScriptApp.getOAuthToken(); } catch (e) { Logger.log('resolveEmailsToUserIds_ token err: ' + e.message); return result; }
    missing.forEach(function(em){
      try {
        var resp = UrlFetchApp.fetch('https://admin.googleapis.com/admin/directory/v1/users/' + encodeURIComponent(em) + '?projection=basic&viewType=domain_public', { method: 'get', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
        if (resp.getResponseCode() < 400) {
          var du = JSON.parse(resp.getContentText());
          if (du && du.id) {
            result.push(du.id);
            try { cache.put('e2uid_' + em, du.id, 21600); } catch (e2) {}
            return;
          }
        } else {
          Logger.log('resolveEmailsToUserIds_ miss email=' + em + ' http=' + resp.getResponseCode());
        }
        try { cache.put('e2uid_' + em, '-', 1800); } catch (e3) {}
      } catch (e4) { Logger.log('resolveEmailsToUserIds_ email=' + em + ' err=' + e4.message); }
    });
  }
  // 最終的な重複排除
  var dedup = [], seen2 = {};
  result.forEach(function(id){ if (!seen2[id]) { seen2[id] = true; dedup.push(id); } });
  return dedup;
}

// 送信値(values: fieldName→value) と fields 配列から、選択された選択肢に紐づく
// メンション先 email を全フィールド横断で集約する。
// dropdown: 単一選択 / checkbox: カンマ区切り複数選択 の両方に対応。
// 戻り値: email 配列(重複排除、小文字化)
function collectMentionEmailsFromValues_(fields, values) {
  var out = [], seen = {};
  (fields || []).forEach(function(f){
    if (f.type !== 'dropdown' && f.type !== 'checkbox') return;
    var parsed = parseFieldOptionsForMentions_(f.options);
    if (!parsed.length) return;
    var map = {};
    parsed.forEach(function(p){ map[p.label] = p.emails; });
    var raw = String(values[f.name] || '');
    if (!raw) return;
    var selected = (f.type === 'checkbox') ? raw.split(',').map(function(s){ return s.trim(); }) : [raw.trim()];
    selected.forEach(function(label){
      var ems = map[label];
      if (!ems) return;
      ems.forEach(function(e){ var k = String(e||'').toLowerCase(); if (k && !seen[k]) { seen[k] = true; out.push(k); } });
    });
  });
  return out;
}

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
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).getValues();
    var wfs = [];
    // status === '有効' のみ(archived や '無効' は除外)
    data.forEach(function(r) { if (r[7] === '有効') wfs.push({ id: r[0], name: r[1], type: r[2], execSpace: r[3], targetSpace: r[4], creator: r[5], created: r[6], status: r[7], approvers: r[9] || '', externalSheet: r[10] || '', description: r[15] || '' }); });
    return wfs;
  } catch (err) { recordDataError_('getActiveWorkflows_', err); return []; }
}

function getWorkflowById_(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_SETTINGS);
    if (!sheet || sheet.getLastRow() < 2) return null;
    var lastCol = Math.max(sheet.getLastColumn(), 16);
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
    // archived は返さない(復元 API 経由でのみ archived を扱う)
    for (var i = 0; i < data.length; i++) { if (data[i][0] === id && data[i][7] !== 'archived') return { id: data[i][0], name: data[i][1], type: data[i][2], execSpace: data[i][3], targetSpace: data[i][4], creator: data[i][5], created: data[i][6], status: data[i][7], approvers: data[i][9] || '', externalSheet: data[i][10] || '', approverUserIds: data[i][12] || '', mentionText: data[i][13] || '', collaborators: data[i][14] || '', description: data[i][15] || '' }; }
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
    var lastCol = Math.max(9, sheet.getLastColumn());
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
    var btns = [];
    var truthy = function(v) { return v === true || v === 'TRUE' || v === 'true'; };
    data.forEach(function(r) {
      if (r[0] !== wid) return;
      var type = String(r[9] || '').trim() || 'status'; // 列J: 空=旧データは status
      btns.push({
        name: r[1], status: r[2], color: r[3], order: r[4],
        threadReply: truthy(r[5]), replyMsg: r[6] || '',
        type: type,
        linkedWorkflowId: r[10] ? Number(r[10]) : 0, // 列K
        inheritThread: truthy(r[11]) // 列L
      });
    });
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
