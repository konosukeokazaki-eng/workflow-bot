// ============================================================
// AdminAPI.gs - 管理画面のサーバー側API
// ============================================================

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getAdminPage_() {
  var tpl = HtmlService.createTemplateFromFile('admin');
  // doGet で既に isAdmin_ を通過済み。初期データは内部関数から直接取得し、
  // requireAdmin_ を経由しない(Session コンテキスト差異による誤検知を回避)。
  tpl.initialData = JSON.stringify(getAllWorkflowsForRender_());
  tpl.initialSpaces = '[]';
  tpl.initialAdmins = JSON.stringify(getAdminsForRender_());
  tpl.currentUser = Session.getActiveUser().getEmail();
  tpl.appVersion = VERSION;
  tpl.contactEmail = CONTACT_EMAIL;
  tpl.envName = ENV_NAME;
  return tpl.evaluate().setTitle('ワークフローBot 管理画面');
}

// 書き込み系 API(api_saveWorkflow / api_deleteWorkflow / api_restoreWorkflow /
// api_toggleStatus / api_saveAdmins)後に必ず呼ぶ。
function invalidateRenderCache_() {
  try { CacheService.getScriptCache().removeAll(['workflows_all', 'admins_all']); } catch (e) {}
}

function getAllWorkflowsForRender_() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('workflows_all');
    if (cached) return JSON.parse(cached);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_SETTINGS);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var lastCol = Math.max(sheet.getLastColumn(), 14);
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
    // archived を除外する。復元一覧は api_getArchivedWorkflows で別途取得。
    var result = data.filter(function(r) { return r[7] !== 'archived'; }).map(function(r) { return { id: r[0], name: r[1], type: r[2], execSpace: r[3], targetSpace: r[4], creator: r[5], created: r[6], status: r[7], approvers: r[9] || '', externalSheet: r[10] || '', approverUserIds: r[12] || '', mentionText: r[13] || '' }; });
    try { cache.put('workflows_all', JSON.stringify(result), 300); } catch (e) {}
    return result;
  } catch (err) { Logger.log('getAllWorkflowsForRender_ エラー: ' + err.message); return []; }
}

function getAdminsForRender_() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('admins_all');
    if (cached) return JSON.parse(cached);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_ADMINS);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var result = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(function(r) { return String(r[0]).trim(); }).filter(function(e) { return e; });
    try { cache.put('admins_all', JSON.stringify(result), 300); } catch (e) {}
    return result;
  } catch (err) { Logger.log('getAdminsForRender_ エラー: ' + err.message); return []; }
}

function api_getWorkflows() {
  try {
    requireAdmin_();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_SETTINGS);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var lastCol = Math.max(sheet.getLastColumn(), 14);
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
    return data.filter(function(r) { return r[7] !== 'archived'; }).map(function(r) { return { id: r[0], name: r[1], type: r[2], execSpace: r[3], targetSpace: r[4], creator: r[5], created: r[6], status: r[7], approvers: r[9] || '', externalSheet: r[10] || '', approverUserIds: r[12] || '', mentionText: r[13] || '' }; });
  } catch (err) { recordDataError_('api_getWorkflows', err); return []; }
}

function api_getArchivedWorkflows() {
  try {
    requireAdmin_();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_SETTINGS);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
    return data.filter(function(r) { return r[7] === 'archived'; }).map(function(r) { return { id: r[0], name: r[1], type: r[2], archivedAt: r[11] || '', creator: r[5], approvers: r[9] || '' }; });
  } catch (err) { recordDataError_('api_getArchivedWorkflows', err); return []; }
}

function api_restoreWorkflow(wid) {
  var lock = LockService.getScriptLock();
  try {
    requireAdmin_();
    if (!lock.tryLock(15000)) return { success: false, error: '他の管理操作が進行中です。10秒ほど待ってから再度お試しください。' };
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var st = ss.getSheetByName(SHEET_SETTINGS);
    if (st.getLastRow() < 2) return { success: false, error: 'ワークフローが見つかりません' };
    var rows = st.getRange(2, 1, st.getLastRow() - 1, 12).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (rows[i][0] === wid && rows[i][7] === 'archived') {
        st.getRange(i + 2, 8).setValue('無効');
        st.getRange(i + 2, 12).setValue('');
        writeAuditLog_('workflow.restore', 'wid=' + wid, rows[i][1]);
        invalidateRenderCache_();
        return { success: true, name: rows[i][1] };
      }
    }
    return { success: false, error: 'アーカイブ済みのワークフローが見つかりません' };
  } catch (err) { Logger.log('api_restoreWorkflow エラー: ' + err.message); try { notifyAdminOnError_(err, 'api_restoreWorkflow'); } catch (e2) {} return { success: false, error: err.message }; }
  finally { try { lock.releaseLock(); } catch (e) {} }
}

// アーカイブ済みワークフローの物理削除。
// _設定 の行、_項目設定 / _ボタン設定 の該当行、およびデータシート(あれば)を削除する。
// archived 状態のワークフローのみ対象。有効/無効のワークフローは削除できない(安全策)。
function api_purgeWorkflow(wid) {
  var lock = LockService.getScriptLock();
  try {
    requireAdmin_();
    if (!lock.tryLock(30000)) return { success: false, error: '他の管理操作が進行中です。10秒ほど待ってから再度お試しください。' };
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var st = ss.getSheetByName(SHEET_SETTINGS);
    if (!st || st.getLastRow() < 2) return { success: false, error: 'ワークフローが見つかりません' };
    var rows = st.getRange(2, 1, st.getLastRow() - 1, 8).getValues();
    var wn = '', foundRow = -1;
    for (var i = 0; i < rows.length; i++) {
      if (rows[i][0] === wid) {
        if (rows[i][7] !== 'archived') return { success: false, error: 'アーカイブ済みのワークフローのみ完全削除できます。先にアーカイブしてください。' };
        wn = rows[i][1];
        foundRow = i + 2;
        break;
      }
    }
    if (foundRow < 0) return { success: false, error: 'ワークフローが見つかりません' };
    st.deleteRow(foundRow);
    var fl = ss.getSheetByName(SHEET_FIELDS);
    if (fl && fl.getLastRow() >= 2) {
      var fd = fl.getRange(2, 1, fl.getLastRow() - 1, 1).getValues();
      for (var j = fd.length - 1; j >= 0; j--) { if (fd[j][0] === wid) fl.deleteRow(j + 2); }
    }
    var bt = ss.getSheetByName(SHEET_BUTTONS);
    if (bt && bt.getLastRow() >= 2) {
      var bd = bt.getRange(2, 1, bt.getLastRow() - 1, 1).getValues();
      for (var k = bd.length - 1; k >= 0; k--) { if (bd[k][0] === wid) bt.deleteRow(k + 2); }
    }
    if (wn) { var ds = ss.getSheetByName(wn); if (ds) ss.deleteSheet(ds); }
    writeAuditLog_('workflow.purge', 'wid=' + wid, wn);
    invalidateRenderCache_();
    return { success: true, name: wn };
  } catch (err) { Logger.log('api_purgeWorkflow エラー: ' + err.message); try { notifyAdminOnError_(err, 'api_purgeWorkflow'); } catch (e2) {} return { success: false, error: err.message }; }
  finally { try { lock.releaseLock(); } catch (e) {} }
}

function api_getFields(wid) {
  try { requireAdmin_(); return getWorkflowFields_(wid); } catch (err) { recordDataError_('api_getFields', err); return []; }
}

function api_getButtons(wid) {
  try { requireAdmin_(); return getWorkflowButtons_(wid); } catch (err) { recordDataError_('api_getButtons', err); return []; }
}

function api_getWorkflowDetail(wid) {
  try { requireAdmin_(); return { fields: getWorkflowFields_(wid), buttons: getWorkflowButtons_(wid) }; }
  catch (err) { recordDataError_('api_getWorkflowDetail', err); return { fields: [], buttons: [] }; }
}

function api_getSpaces() {
  try {
    requireAdmin_();
    var resp = UrlFetchApp.fetch('https://chat.googleapis.com/v1/spaces', { method: 'get', headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    var r = JSON.parse(resp.getContentText());
    if (!r.spaces) return [];
    return r.spaces.filter(function(s) { return s.type === 'ROOM' && s.displayName; }).map(function(s) { return { id: s.name, name: s.displayName }; });
  } catch (err) { recordDataError_('api_getSpaces', err); return []; }
}

// 指定スペースの HUMAN メンバーを {userId, email, displayName} 配列で返す。
// Bot がスペース未参加ならサービスアカウントで参加させてから再試行する。
// メール解決は People API を使う(directory.readonly スコープが必要)。
// 結果は CacheService に 1 時間キャッシュする。
function api_getSpaceMembers(spaceId) {
  try {
    requireAdmin_();
    if (!spaceId) return [];
    var cache = CacheService.getScriptCache();
    var cacheKey = 'space_members_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, spaceId)).substring(0, 24);
    var cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);

    var saToken = getServiceAccountToken_();
    var listUrl = 'https://chat.googleapis.com/v1/' + spaceId + '/members?pageSize=200';
    var resp = UrlFetchApp.fetch(listUrl, { method: 'get', headers: { 'Authorization': 'Bearer ' + saToken }, muteHttpExceptions: true });
    var code = resp.getResponseCode();
    // 403/404 のときは Bot 未参加の可能性が高いので追加してリトライ
    if (code === 403 || code === 404) {
      try { addBotToSpace_(spaceId); } catch (e) {}
      Utilities.sleep(1500);
      resp = UrlFetchApp.fetch(listUrl, { method: 'get', headers: { 'Authorization': 'Bearer ' + saToken }, muteHttpExceptions: true });
      code = resp.getResponseCode();
    }
    if (code >= 400) throw new Error('spaces.members.list 失敗: HTTP ' + code + ' ' + resp.getContentText().substring(0, 200));

    var data = JSON.parse(resp.getContentText());
    var memberships = data.memberships || [];
    // HUMAN のみ、ID 抽出
    var members = memberships.filter(function(m) { return m.member && m.member.type === 'HUMAN' && m.state === 'JOINED'; }).map(function(m) {
      var id = String(m.member.name).replace('users/', '');
      return { userId: id, email: '', displayName: m.member.displayName || '' };
    });

    // People API で email を解決(バッチで最大 50 件)
    var userToken = ScriptApp.getOAuthToken();
    for (var i = 0; i < members.length; i += 50) {
      var chunk = members.slice(i, i + 50);
      var params = chunk.map(function(m) { return 'resourceNames=people/' + encodeURIComponent(m.userId); }).join('&') + '&personFields=emailAddresses';
      try {
        var pr = UrlFetchApp.fetch('https://people.googleapis.com/v1/people:batchGet?' + params, { method: 'get', headers: { 'Authorization': 'Bearer ' + userToken }, muteHttpExceptions: true });
        if (pr.getResponseCode() < 400) {
          var pd = JSON.parse(pr.getContentText());
          (pd.responses || []).forEach(function(rsp) {
            if (!rsp.person) return;
            var rn = String(rsp.requestedResourceName || rsp.person.resourceName || '').replace('people/', '');
            var emails = rsp.person.emailAddresses || [];
            var primary = emails.filter(function(e) { return e.metadata && e.metadata.primary; })[0] || emails[0];
            if (primary && primary.value) {
              for (var j = 0; j < members.length; j++) { if (members[j].userId === rn) { members[j].email = primary.value; break; } }
            }
          });
        }
      } catch (e) { Logger.log('people.batchGet 失敗: ' + e.message); }
    }

    // 名前で並び替え
    members.sort(function(a, b) { return (a.displayName || '').localeCompare(b.displayName || '', 'ja'); });
    try { cache.put(cacheKey, JSON.stringify(members), 3600); } catch (e) {}
    return members;
  } catch (err) { recordDataError_('api_getSpaceMembers', err); try { notifyAdminOnError_(err, 'api_getSpaceMembers'); } catch (e) {} return []; }
}

function api_getAdmins() {
  try {
    requireAdmin_();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_ADMINS);
    if (!sheet || sheet.getLastRow() < 2) return [];
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(function(r) { return String(r[0]).trim(); }).filter(function(e) { return e; });
  } catch (err) { recordDataError_('api_getAdmins', err); return []; }
}

function api_saveWorkflow(data) {
  var lock = LockService.getScriptLock();
  try {
    requireAdmin_();
    var nameErr = validateWorkflowName_(data && data.name);
    if (nameErr) return { success: false, error: nameErr };
    var approverErr = validateEmailList_(data.approvers);
    if (approverErr) return { success: false, error: approverErr };
    var urlErr = validateSheetUrl_(data.externalSheet);
    if (urlErr) return { success: false, error: urlErr };
    if (!lock.tryLock(30000)) return { success: false, error: '他の管理操作が進行中です。10秒ほど待ってから再度お試しください。' };
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var stSheet = ss.getSheetByName(SHEET_SETTINGS);
    var flSheet = ss.getSheetByName(SHEET_FIELDS);
    var btSheet = ss.getSheetByName(SHEET_BUTTONS);
    // 13 列目(承認者Chat User ID) / 14 列目(メンション文言)ヘッダを初回だけ用意する
    if (stSheet.getLastColumn() < 13 || stSheet.getRange(1, 13).getValue() !== '承認者UserID') stSheet.getRange(1, 13).setValue('承認者UserID');
    if (stSheet.getLastColumn() < 14 || stSheet.getRange(1, 14).getValue() !== 'メンション文言') stSheet.getRange(1, 14).setValue('メンション文言');
    var approverUserIds = String(data.approverUserIds || '').split(',').map(function(s) { return s.trim(); }).filter(function(s) { return s; }).join(',');
    var mentionText = String(data.mentionText || '').trim();
    var isNew = !data.id, wid, oldName = '';
    // 重複チェック(採番前に実施)。アーカイブ済み(status='archived')は無視する。
    var allRows = stSheet.getLastRow() >= 2 ? stSheet.getRange(2, 1, stSheet.getLastRow() - 1, 8).getValues() : [];
    for (var d = 0; d < allRows.length; d++) {
      if (allRows[d][1] === data.name && allRows[d][0] !== data.id && allRows[d][7] !== 'archived') {
        return { success: false, error: '同じ名前のワークフローが既に存在します: ' + data.name };
      }
    }
    if (isNew) {
      wid = 1;
      if (stSheet.getLastRow() >= 2) { stSheet.getRange(2, 1, stSheet.getLastRow() - 1, 1).getValues().forEach(function(r) { if (r[0] >= wid) wid = r[0] + 1; }); }
      stSheet.appendRow([wid, data.name, data.type, data.execSpace, data.targetSpace, Session.getActiveUser().getEmail(), new Date(), '有効', '', data.approvers || '', data.externalSheet || '', '', approverUserIds, mentionText]);
    } else {
      wid = data.id;
      var rows = stSheet.getRange(2, 1, stSheet.getLastRow() - 1, 11).getValues();
      for (var i = 0; i < rows.length; i++) { if (rows[i][0] === wid) { var rn = i + 2; oldName = rows[i][1]; stSheet.getRange(rn, 2).setValue(data.name); stSheet.getRange(rn, 3).setValue(data.type); stSheet.getRange(rn, 4).setValue(data.execSpace); stSheet.getRange(rn, 5).setValue(data.targetSpace); stSheet.getRange(rn, 10).setValue(data.approvers || ''); stSheet.getRange(rn, 11).setValue(data.externalSheet || ''); stSheet.getRange(rn, 13).setValue(approverUserIds); stSheet.getRange(rn, 14).setValue(mentionText); break; } }
    }
    if (flSheet.getLastRow() >= 2) { var fd = flSheet.getRange(2, 1, flSheet.getLastRow() - 1, 6).getValues(); for (var j = fd.length - 1; j >= 0; j--) { if (fd[j][0] === wid) flSheet.deleteRow(j + 2); } }
    if (data.fields) data.fields.forEach(function(f, idx) { flSheet.appendRow([wid, f.name, f.type, f.options || '', f.required || false, idx + 1]); });
    if (btSheet.getLastRow() >= 2) { var bd = btSheet.getRange(2, 1, btSheet.getLastRow() - 1, 9).getValues(); for (var k = bd.length - 1; k >= 0; k--) { if (bd[k][0] === wid) btSheet.deleteRow(k + 2); } }
    if (data.buttons) data.buttons.forEach(function(b, idx) { btSheet.appendRow([wid, b.name, b.status, b.color, idx + 1, b.threadReply || false, b.replyMsg || '', false, '']); });
    var sn = oldName || data.name;
    var ds = ss.getSheetByName(sn);
    if (ds) {
      if (oldName && oldName !== data.name) ds.setName(data.name);
      var nh = ['タイムスタンプ', '送信者', '送信者メール'];
      if (data.fields) data.fields.forEach(function(f) { nh.push(f.name); });
      if (data.type === '申請・承認') nh.push('ステータス');
      nh.push('メッセージ名');
      ds.getRange(1, 1, 1, nh.length).setValues([nh]);
    } else {
      ds = ss.insertSheet(data.name);
      var h = ['タイムスタンプ', '送信者', '送信者メール'];
      if (data.fields) data.fields.forEach(function(f) { h.push(f.name); });
      if (data.type === '申請・承認') h.push('ステータス');
      h.push('メッセージ名');
      ds.appendRow(h);
    }
    if (data.externalSheet) {
      try {
        var eid = data.externalSheet.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (eid) { var ess = SpreadsheetApp.openById(eid[1]); var es = ess.getSheetByName(data.name); if (!es) { es = ess.insertSheet(data.name); var eh = ['タイムスタンプ', '送信者', '送信者メール']; if (data.fields) data.fields.forEach(function(f) { eh.push(f.name); }); if (data.type === '申請・承認') eh.push('ステータス'); es.appendRow(eh); } }
      } catch (err) { Logger.log('外部シート作成エラー: ' + err.message); }
    }
    writeAuditLog_(isNew ? 'workflow.create' : 'workflow.update', 'wid=' + wid, data.name);
    invalidateRenderCache_();
    return { success: true, id: wid };
  } catch (err) { Logger.log('api_saveWorkflow エラー: ' + err.message); try { notifyAdminOnError_(err, 'api_saveWorkflow'); } catch (e2) {} return { success: false, error: err.message }; }
  finally { try { lock.releaseLock(); } catch (e) {} }
}

// 論理削除: status を 'archived' にして _設定.archivedAt(12列目) に日時を記録する。
// 30 日経過後に monthlyArchive_ 内の purgeArchivedWorkflows_ で物理削除される。
function api_deleteWorkflow(wid) {
  var lock = LockService.getScriptLock();
  try {
    requireAdmin_();
    if (!lock.tryLock(30000)) return { success: false, error: '他の管理操作が進行中です。10秒ほど待ってから再度お試しください。' };
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var st = ss.getSheetByName(SHEET_SETTINGS);
    if (st.getLastRow() < 2) return { success: false, error: 'ワークフローが見つかりません' };
    // archivedAt 列(L列)を初回だけ用意する
    if (st.getRange(1, 12).getValue() !== 'アーカイブ日時') st.getRange(1, 12).setValue('アーカイブ日時');
    var rows = st.getRange(2, 1, st.getLastRow() - 1, 8).getValues();
    var wn = '';
    for (var i = 0; i < rows.length; i++) {
      if (rows[i][0] === wid) {
        wn = rows[i][1];
        st.getRange(i + 2, 8).setValue('archived');
        st.getRange(i + 2, 12).setValue(new Date());
        break;
      }
    }
    writeAuditLog_('workflow.archive', 'wid=' + wid, wn);
    invalidateRenderCache_();
    return { success: true };
  } catch (err) { Logger.log('api_deleteWorkflow エラー: ' + err.message); try { notifyAdminOnError_(err, 'api_deleteWorkflow'); } catch (e2) {} return { success: false, error: err.message }; }
  finally { try { lock.releaseLock(); } catch (e) {} }
}

function api_toggleStatus(wid) {
  try {
    requireAdmin_();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_SETTINGS);
    if (sheet.getLastRow() < 2) return { success: false };
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
    for (var i = 0; i < rows.length; i++) { if (rows[i][0] === wid) { var ns = (rows[i][7] === '有効') ? '無効' : '有効'; sheet.getRange(i + 2, 8).setValue(ns); writeAuditLog_('workflow.toggle', 'wid=' + wid, ns); invalidateRenderCache_(); return { success: true, status: ns }; } }
    return { success: false };
  } catch (err) { Logger.log('api_toggleStatus エラー: ' + err.message); try { notifyAdminOnError_(err, 'api_toggleStatus'); } catch (e2) {} return { success: false, error: err.message }; }
}

function api_saveAdmins(admins) {
  try {
    requireAdmin_();
    if (!admins || !Array.isArray(admins) || admins.length === 0) return { success: false, error: '管理者を1人以上指定してください' };
    for (var i = 0; i < admins.length; i++) {
      var e = String(admins[i] || '').trim();
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e)) return { success: false, error: '不正なメール形式: ' + admins[i] };
    }
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_ADMINS);
    if (!sheet) { sheet = ss.insertSheet(SHEET_ADMINS); sheet.appendRow(['メールアドレス']); }
    if (sheet.getLastRow() >= 2) sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).clearContent();
    admins.forEach(function(e, i) { sheet.getRange(i + 2, 1).setValue(String(e).trim()); });
    writeAuditLog_('admins.save', '', admins.join(', '));
    try { invalidateAdminCache_(); } catch (e3) {}
    invalidateRenderCache_();
    return { success: true };
  } catch (err) { Logger.log('api_saveAdmins エラー: ' + err.message); try { notifyAdminOnError_(err, 'api_saveAdmins'); } catch (e2) {} return { success: false, error: err.message }; }
}

// 現在のユーザのプライバシー同意状態を返す。列 B(privacyAcceptedAt)を確認する。
function api_getPrivacyStatus() {
  try {
    var email = requireAdmin_();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_ADMINS);
    if (!sheet || sheet.getLastRow() < 2) return { accepted: true };
    // ヘッダーに privacyAcceptedAt 列がなければ追加
    if (sheet.getRange(1, 2).getValue() !== '同意日時') sheet.getRange(1, 2).setValue('同意日時');
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === email.toLowerCase()) {
        return { accepted: !!data[i][1], acceptedAt: data[i][1] || null };
      }
    }
    return { accepted: false };
  } catch (err) { recordDataError_('api_getPrivacyStatus', err); return { accepted: true, error: err.message }; }
}

function api_recordPrivacyAcceptance() {
  try {
    var email = requireAdmin_();
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_ADMINS);
    if (!sheet || sheet.getLastRow() < 2) return { success: false, error: '管理者シートが見つかりません' };
    if (sheet.getRange(1, 2).getValue() !== '同意日時') sheet.getRange(1, 2).setValue('同意日時');
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === email.toLowerCase()) {
        sheet.getRange(i + 2, 2).setValue(new Date());
        writeAuditLog_('privacy.accept', email, VERSION);
        return { success: true };
      }
    }
    return { success: false, error: '管理者エントリが見つかりません' };
  } catch (err) { Logger.log('api_recordPrivacyAcceptance エラー: ' + err.message); return { success: false, error: err.message }; }
}

function api_getQuotaStatus() {
  try { requireAdmin_(); return getQuotaStatus_(); }
  catch (err) { recordDataError_('api_getQuotaStatus', err); return { error: err.message, items: [] }; }
}

function api_deployButton(wid) {
  try {
    requireAdmin_();
    var wf = getWorkflowById_(wid);
    if (!wf) return { success: false, error: 'ワークフローが見つかりません' };
    addBotToSpace_(wf.execSpace);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var st = ss.getSheetByName(SHEET_SETTINGS);
    var rows = st.getRange(2, 1, st.getLastRow() - 1, 11).getValues();
    for (var i = 0; i < rows.length; i++) { if (rows[i][0] === wid && rows[i][8]) { try { UrlFetchApp.fetch('https://chat.googleapis.com/v1/' + rows[i][8], { method: 'delete', headers: { 'Authorization': 'Bearer ' + getServiceAccountToken_() }, muteHttpExceptions: true }); } catch (e) {} break; } }
    var label = (wf.type === '申請・承認') ? '申請する' : '記録する';
    var icon = (wf.type === '申請・承認') ? '📋' : '📝';
    var t = getServiceAccountToken_();
    var resp = UrlFetchApp.fetch('https://chat.googleapis.com/v1/' + wf.execSpace + '/messages', { method: 'post', contentType: 'application/json', headers: { 'Authorization': 'Bearer ' + t }, payload: JSON.stringify({ cardsV2: [{ cardId: 'launcher_' + wid, card: { header: { title: icon + ' ' + wf.name }, sections: [{ widgets: [{ buttonList: { buttons: [{ text: label, color: COLOR_MAP.blue, onClick: { action: { function: FUNCTION_URL, interaction: 'OPEN_DIALOG', parameters: [{ key: 'actionName', value: 'openWorkflow' }, { key: 'workflowId', value: String(wid) }] } } }] } }] }] } }] }), muteHttpExceptions: true });
    var r = JSON.parse(resp.getContentText());
    if (r.name) { for (var j = 0; j < rows.length; j++) { if (rows[j][0] === wid) { st.getRange(j + 2, 9).setValue(r.name); break; } } writeAuditLog_('workflow.deploy', 'wid=' + wid, wf.name); return { success: true }; }
    return { success: false, error: resp.getContentText() };
  } catch (err) { Logger.log('api_deployButton エラー: ' + err.message); try { notifyAdminOnError_(err, 'api_deployButton'); } catch (e2) {} return { success: false, error: err.message }; }
}

function initialSetup() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var s1 = ss.getSheetByName(SHEET_SETTINGS); if (!s1) { s1 = ss.insertSheet(SHEET_SETTINGS); s1.appendRow(['ID','ワークフロー名','種類','実行スペースID','送信先スペースID','作成者','作成日','ステータス','ボタンカードメッセージ名','承認者','外部スプレッドシートURL']); }
  var s2 = ss.getSheetByName(SHEET_FIELDS); if (!s2) { s2 = ss.insertSheet(SHEET_FIELDS); s2.appendRow(['ワークフローID','項目名','入力タイプ','選択肢','必須','順番']); }
  var s3 = ss.getSheetByName(SHEET_BUTTONS); if (!s3) { s3 = ss.insertSheet(SHEET_BUTTONS); s3.appendRow(['ワークフローID','ボタン名','ステータス値','色','順番','スレッド返信','返信メッセージ','DM通知','DMメッセージ']); }
  var s4 = ss.getSheetByName(SHEET_ADMINS); if (!s4) { s4 = ss.insertSheet(SHEET_ADMINS); s4.appendRow(['メールアドレス']); s4.appendRow([Session.getActiveUser().getEmail()]); }
  var s5 = ss.getSheetByName(SHEET_AUDIT); if (!s5) { s5 = ss.insertSheet(SHEET_AUDIT); s5.appendRow(['タイムスタンプ','実行者','アクション','対象','詳細']); }
  if (ss.getSheetByName(SHEET_SETTINGS).getRange(1, 11).getValue() !== '外部スプレッドシートURL') ss.getSheetByName(SHEET_SETTINGS).getRange(1, 11).setValue('外部スプレッドシートURL');
  Logger.log('初期セットアップ完了');
}
