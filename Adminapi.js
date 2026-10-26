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
  tpl.initialLinks = JSON.stringify(getLinkRelations_());
  var _cu = Session.getActiveUser().getEmail();
  tpl.currentUser = _cu;
  tpl.currentRole = getEffectiveUserRole_(_cu) || ROLE_OWNER;
  tpl.appVersion = VERSION;
  tpl.contactEmail = CONTACT_EMAIL;
  tpl.envName = ENV_NAME;
  // 別タブで開くヘルプ用: 現在の実行URL(@HEADなら/dev, 本番なら/exec)を返す。
  // FUNCTION_URL はChat用の固定バージョンなので、新タブが古いバージョンを開いてしまう。
  try { tpl.selfUrl = ScriptApp.getService().getUrl(); } catch (e) { tpl.selfUrl = FUNCTION_URL; }
  return tpl.evaluate().setTitle('ワークフローBot 管理画面');
}

// 運用マニュアル HTML を返す。管理者/申請者を問わず閲覧可(認証なし)。
function renderManualPage_() {
  var tpl = HtmlService.createTemplateFromFile('Manual');
  tpl.appVersion = VERSION;
  return tpl.evaluate().setTitle('Workflow Bot 運用マニュアル').addMetaTag('viewport', 'width=device-width,initial-scale=1');
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
    var full;
    if (cached) {
      full = JSON.parse(cached);
    } else {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = ss.getSheetByName(SHEET_SETTINGS);
      if (!sheet || sheet.getLastRow() < 2) return [];
      var lastCol = Math.max(sheet.getLastColumn(), 15);
      var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
      full = data.filter(function(r) { return r[7] !== 'archived'; }).map(function(r) { return { id: r[0], name: r[1], type: r[2], execSpace: r[3], targetSpace: r[4], creator: r[5], created: r[6], status: r[7], approvers: r[9] || '', externalSheet: r[10] || '', approverUserIds: r[12] || '', mentionText: r[13] || '', collaborators: r[14] || '' }; });
      try { cache.put('workflows_all', JSON.stringify(full), 300); } catch (e) {}
    }
    return _annotateWorkflowsForUser_(full);
  } catch (err) { Logger.log('getAllWorkflowsForRender_ エラー: ' + err.message); return []; }
}

// 現在のユーザ視点で閲覧可能なもののみに絞り、canEdit/canDelete を付与する。
function _annotateWorkflowsForUser_(list) {
  var email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch (e) {}
  var role = getEffectiveUserRole_(email) || ROLE_OWNER;
  var out = [];
  for (var i = 0; i < list.length; i++) {
    var w = list[i];
    if (!canViewWorkflow_(email, w, role)) continue;
    w.canEdit = canEditWorkflow_(email, w, role);
    w.canDelete = canDeleteWorkflow_(email, w, role);
    out.push(w);
  }
  return out;
}

function getAdminsForRender_() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('admins_all');
    if (cached) return JSON.parse(cached);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureAdminSchema_(ss);
    var sheet = ss.getSheetByName(SHEET_ADMINS);
    if (!sheet || sheet.getLastRow() < 2) return [];
    var lastCol = Math.max(sheet.getLastColumn(), 3);
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
    var result = data.map(function(r) {
      var email = String(r[0] || '').trim();
      if (!email) return null;
      var role = String(r[2] || '').trim().toLowerCase();
      if (role !== ROLE_EDITOR && role !== ROLE_MEMBER) role = ROLE_OWNER;
      return { email: email, role: role };
    }).filter(function(e) { return e; });
    try { cache.put('admins_all', JSON.stringify(result), 300); } catch (e) {}
    return result;
  } catch (err) { Logger.log('getAdminsForRender_ エラー: ' + err.message); return []; }
}

function api_getWorkflows() {
  try { requireAdmin_(); return getAllWorkflowsForRender_(); }
  catch (err) { recordDataError_('api_getWorkflows', err); return []; }
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
    var caller = requireAdmin_();
    if (!canManageAdmins_(getEffectiveUserRole_(caller))) return { success: false, error: 'この操作はオーナー権限が必要です' };
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
    var caller = requireAdmin_();
    if (!canManageAdmins_(getEffectiveUserRole_(caller))) return { success: false, error: 'この操作はオーナー権限が必要です' };
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
  try { requireAdmin_(); return { fields: getWorkflowFields_(wid), buttons: getWorkflowButtons_(wid), parents: findLaunchParents_(wid) }; }
  catch (err) { recordDataError_('api_getWorkflowDetail', err); return { fields: [], buttons: [], parents: [] }; }
}

// 全ワークフローのlaunch連動関係を [{parent, child, buttonName, color, inheritThread, status}] で返す。
// 画面サイドバーの階層表示用。
function getLinkRelations_() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var bt = ss.getSheetByName(SHEET_BUTTONS);
    if (!bt || bt.getLastRow() < 2) return [];
    var lastCol = Math.max(12, bt.getLastColumn());
    var data = bt.getRange(2, 1, bt.getLastRow() - 1, lastCol).getValues();
    var truthy = function(v) { return v === true || v === 'TRUE' || v === 'true'; };
    var out = [];
    data.forEach(function(r) {
      if (String(r[9] || '').trim() !== 'launch') return;
      var child = Number(r[10]);
      if (!child) return;
      out.push({ parent: Number(r[0]), child: child, buttonName: r[1], color: r[3] || 'gray', inheritThread: truthy(r[11]), status: r[2] || '' });
    });
    return out;
  } catch (err) { recordDataError_('getLinkRelations_', err); return []; }
}

function api_getLinkRelations() {
  try { requireAdmin_(); return getLinkRelations_(); }
  catch (err) { recordDataError_('api_getLinkRelations', err); return []; }
}

// 指定WF ID を launch ボタンから参照している親 WF を全て返す。
// 戻り値: [{ wfId, wfName, buttonName, inheritThread }]
function findLaunchParents_(childWid) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var bt = ss.getSheetByName(SHEET_BUTTONS);
    if (!bt || bt.getLastRow() < 2) return [];
    var lastCol = Math.max(12, bt.getLastColumn());
    var data = bt.getRange(2, 1, bt.getLastRow() - 1, lastCol).getValues();
    var parents = [];
    var truthy = function(v) { return v === true || v === 'TRUE' || v === 'true'; };
    data.forEach(function(r) {
      if (String(r[9] || '').trim() === 'launch' && Number(r[10]) === Number(childWid)) {
        var parentWf = getWorkflowById_(r[0]);
        if (parentWf) parents.push({ wfId: r[0], wfName: parentWf.name, buttonName: r[1], inheritThread: truthy(r[11]) });
      }
    });
    return parents;
  } catch (err) { recordDataError_('findLaunchParents_', err); return []; }
}

function api_getSpaces() {
  try {
    var caller = requireAdmin_();
    // Service Account の Domain-Wide Delegation でアクセスユーザーに impersonate し、
    // そのユーザーが参加しているスペース一覧を取得する。
    // 失敗時(DWD未設定・スコープ未許可など)は従来のデプロイ者トークンにフォールバック。
    var token;
    try { token = getImpersonatedToken_(caller, 'https://www.googleapis.com/auth/chat.spaces.readonly'); }
    catch (e) { recordDataError_('api_getSpaces.impersonate', e); token = ScriptApp.getOAuthToken(); }
    var out = [], pageToken = '';
    do {
      var url = 'https://chat.googleapis.com/v1/spaces?pageSize=200' + (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');
      var resp = UrlFetchApp.fetch(url, { method: 'get', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
      var code = resp.getResponseCode();
      if (code >= 400) { recordDataError_('api_getSpaces', new Error('HTTP ' + code + ' ' + resp.getContentText().substring(0, 300))); return []; }
      var r = JSON.parse(resp.getContentText());
      (r.spaces || []).forEach(function(s) {
        var isRoom = (s.spaceType === 'SPACE' || s.type === 'ROOM');
        if (isRoom && s.displayName) out.push({ id: s.name, name: s.displayName });
      });
      pageToken = r.nextPageToken || '';
    } while (pageToken);
    return out;
  } catch (err) { recordDataError_('api_getSpaces', err); return []; }
}

// 指定スペースの HUMAN メンバーを {userId, email, displayName} 配列で返す。
// Bot がスペース未参加ならサービスアカウントで参加させてから再試行する。
// メール解決は People API を使う(directory.readonly スコープが必要)。
// 結果は CacheService に 1 時間キャッシュする。
function api_getSpaceMembers(spaceId) {
  try {
    var caller = requireAdmin_();
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

    // People API で email を解決(バッチで最大 50 件)。sources にドメインプロファイル/連絡先を追加して Workspace 社員のメールを引けるようにする。
    var userToken = ScriptApp.getOAuthToken();
    var sourcesParam = '&sources=READ_SOURCE_TYPE_PROFILE&sources=READ_SOURCE_TYPE_DOMAIN_CONTACT&sources=READ_SOURCE_TYPE_CONTACT';
    for (var i = 0; i < members.length; i += 50) {
      var chunk = members.slice(i, i + 50);
      var params = chunk.map(function(m) { return 'resourceNames=people/' + encodeURIComponent(m.userId); }).join('&') + '&personFields=emailAddresses' + sourcesParam;
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
    // People API で取れなかったユーザーは Admin Directory API でフォールバック。
    // スペース一覧取得と同じ DWD impersonation 経由にすることで、デプロイ者が
    // Workspace 管理者でなくても admin.directory.user.readonly スコープで解決できるようにする。
    var missing = members.filter(function(m) { return !m.email; });
    if (missing.length > 0) {
      var adminToken = null;
      try { adminToken = getImpersonatedToken_(caller, 'https://www.googleapis.com/auth/admin.directory.user.readonly'); }
      catch (impErr) { Logger.log('directory impersonate 失敗(デプロイ者トークンにフォールバック): ' + impErr.message); adminToken = userToken; }
      for (var k = 0; k < missing.length; k++) {
        try {
          var dr = UrlFetchApp.fetch('https://admin.googleapis.com/admin/directory/v1/users/' + encodeURIComponent(missing[k].userId) + '?projection=basic&viewType=domain_public', { method: 'get', headers: { 'Authorization': 'Bearer ' + adminToken }, muteHttpExceptions: true });
          var drc = dr.getResponseCode();
          if (drc < 400) {
            var du = JSON.parse(dr.getContentText());
            if (du.primaryEmail) {
              for (var jj = 0; jj < members.length; jj++) { if (members[jj].userId === missing[k].userId) { members[jj].email = du.primaryEmail; break; } }
            }
          } else {
            Logger.log('directory.users.get HTTP ' + drc + ' uid=' + missing[k].userId + ' body=' + dr.getContentText().substring(0, 200));
          }
        } catch (e2) { Logger.log('directory.users.get 例外: ' + missing[k].userId + ' ' + e2.message); }
      }
    }
    var stillMissing = members.filter(function(m) { return !m.email; }).length;
    if (stillMissing > 0) Logger.log('api_getSpaceMembers: email未解決 ' + stillMissing + '/' + members.length + '名 space=' + spaceId);

    // 名前で並び替え
    members.sort(function(a, b) { return (a.displayName || '').localeCompare(b.displayName || '', 'ja'); });
    try { cache.put(cacheKey, JSON.stringify(members), 3600); } catch (e) {}
    return members;
  } catch (err) { recordDataError_('api_getSpaceMembers', err); try { notifyAdminOnError_(err, 'api_getSpaceMembers'); } catch (e) {} return []; }
}

function api_getAdmins() {
  try { requireAdmin_(); return getAdminsForRender_(); }
  catch (err) { recordDataError_('api_getAdmins', err); return []; }
}

// フィールドの追加/削除/改名/並び替えでデータシートの列がズレないよう、
// fieldId ベースの差分検出でデータシートを最小限だけ変更する。
// - 追加: メッセージ名の直前に列挿入 → dataCol を割り当て
// - 削除: 論理削除(deleted=true)。データシート列は保持
// - 改名: ヘッダー行の該当セルのみ更新
// - 並び替え: データシート無変更。order 値のみ更新
function api_saveWorkflow(data) {
  var lock = LockService.getScriptLock();
  try {
    var caller = requireAdmin_();
    var callerRole = getEffectiveUserRole_(caller);
    var nameErr = validateWorkflowName_(data && data.name);
    if (nameErr) return { success: false, error: nameErr };
    var approverErr = validateEmailList_(data.approvers);
    if (approverErr) return { success: false, error: approverErr };
    if (data.type === '申請・承認' && !String(data.approvers || '').trim()) {
      return { success: false, error: '「申請・承認」型は承認者を1名以上設定してください（空の場合、意図しない人物が承認できてしまいます）' };
    }
    var collabErr = validateEmailList_(data.collaborators);
    if (collabErr) return { success: false, error: '共同編集者: ' + collabErr };
    var urlErr = validateSheetUrl_(data.externalSheet);
    if (urlErr) return { success: false, error: urlErr };
    // 既存編集は canEdit チェック
    if (data.id) {
      var existingWf = getWorkflowById_(data.id);
      if (existingWf && !canEditWorkflow_(caller, existingWf, callerRole)) {
        return { success: false, error: 'このワークフローを編集する権限がありません' };
      }
    }
    if (!lock.tryLock(30000)) return { success: false, error: '他の管理操作が進行中です。10秒ほど待ってから再度お試しください。' };
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureFieldSchema_(ss);
    var stSheet = ss.getSheetByName(SHEET_SETTINGS);
    var flSheet = ss.getSheetByName(SHEET_FIELDS);
    var btSheet = ss.getSheetByName(SHEET_BUTTONS);
    if (stSheet.getLastColumn() < 13 || stSheet.getRange(1, 13).getValue() !== '承認者UserID') stSheet.getRange(1, 13).setValue('承認者UserID');
    if (stSheet.getLastColumn() < 14 || stSheet.getRange(1, 14).getValue() !== 'メンション文言') stSheet.getRange(1, 14).setValue('メンション文言');
    if (stSheet.getLastColumn() < 15 || stSheet.getRange(1, 15).getValue() !== '共同編集者') stSheet.getRange(1, 15).setValue('共同編集者');
    if (stSheet.getLastColumn() < 16 || stSheet.getRange(1, 16).getValue() !== '説明文') stSheet.getRange(1, 16).setValue('説明文');
    var approverUserIds = String(data.approverUserIds || '').split(',').map(function(s) { return s.trim(); }).filter(function(s) { return s; }).join(',');
    var mentionText = String(data.mentionText || '').trim();
    var description = String(data.description || '').trim();
    // 承認者・共同編集者メールは不可視文字を除去して保存する
    data.approvers = sanitizeEmailList_(data.approvers);
    var collaboratorsClean = sanitizeEmailList_(data.collaborators);
    var isNew = !data.id, wid, oldName = '';
    var allRows = stSheet.getLastRow() >= 2 ? stSheet.getRange(2, 1, stSheet.getLastRow() - 1, 8).getValues() : [];
    for (var d = 0; d < allRows.length; d++) {
      if (allRows[d][1] === data.name && allRows[d][0] !== data.id && allRows[d][7] !== 'archived') {
        return { success: false, error: '同じ名前のワークフローが既に存在します: ' + data.name };
      }
    }
    if (isNew) {
      wid = 1;
      if (stSheet.getLastRow() >= 2) { stSheet.getRange(2, 1, stSheet.getLastRow() - 1, 1).getValues().forEach(function(r) { if (r[0] >= wid) wid = r[0] + 1; }); }
      stSheet.appendRow([wid, data.name, data.type, data.execSpace, data.targetSpace, caller, new Date(), '有効', '', data.approvers || '', data.externalSheet || '', '', approverUserIds, mentionText, collaboratorsClean, description]);
    } else {
      wid = data.id;
      var rows = stSheet.getRange(2, 1, stSheet.getLastRow() - 1, 11).getValues();
      for (var i = 0; i < rows.length; i++) { if (rows[i][0] === wid) { var rn = i + 2; oldName = rows[i][1]; stSheet.getRange(rn, 2).setValue(data.name); stSheet.getRange(rn, 3).setValue(data.type); stSheet.getRange(rn, 4).setValue(data.execSpace); stSheet.getRange(rn, 5).setValue(data.targetSpace); stSheet.getRange(rn, 10).setValue(data.approvers || ''); stSheet.getRange(rn, 11).setValue(data.externalSheet || ''); stSheet.getRange(rn, 13).setValue(approverUserIds); stSheet.getRange(rn, 14).setValue(mentionText); stSheet.getRange(rn, 15).setValue(collaboratorsClean); stSheet.getRange(rn, 16).setValue(description); break; } }
    }

    // データシートを準備。新規シートは最小ヘッダー(ts/sender/email + status? + messageName)のみで作り、
    // フィールド列は後続の diff.added で一貫して追加する(重複防止)。
    var sn = oldName || data.name;
    var ds = ss.getSheetByName(sn);
    if (!ds) {
      // 新規データシートはタブ一覧の右端に追加(既存のシステムシートより後ろに並べて視認性を上げる)
      ds = ss.insertSheet(data.name, ss.getSheets().length);
      var h = ['タイムスタンプ', '送信者', '送信者メール'];
      if (data.type === '申請・承認') h.push('ステータス');
      h.push('メッセージ名');
      ds.appendRow(h);
    } else if (oldName && oldName !== data.name) {
      ds.setName(data.name);
    }
    // ステータス列は type 変更に追従: 申請・承認で無ければ挿入、テンプレなら不要
    var mCol = findMessageCol_(ds);
    var sCol = findStatusCol_(ds);
    if (data.type === '申請・承認' && sCol === -1) {
      // メッセージ名の直前にステータス列を挿入
      if (mCol > 0) { ds.insertColumnBefore(mCol); ds.getRange(1, mCol).setValue('ステータス'); }
      else { ds.getRange(1, ds.getLastColumn() + 1).setValue('ステータス'); ds.getRange(1, ds.getLastColumn() + 1).setValue('メッセージ名'); }
    }

    // 差分算出
    var oldFieldRows = readFieldRows_(ss, wid);
    var oldActive = oldFieldRows.filter(function(r) { return !r.deleted; });
    var newFields = data.fields || [];
    var diff = computeFieldDiff_(oldActive, newFields);

    // 追加フィールドの列挿入: メッセージ名の直前に順次挿入し、dataCol を確定
    var byId = {};
    oldFieldRows.forEach(function(r) { byId[r.fieldId] = r; });
    // 追加フィールドは「ステータス」列(なければ「メッセージ名」列)の直前に挿入する。
    // 挿入後は毎回位置を取り直す(挿入で全ての後続列がシフトするため)。
    diff.added.forEach(function(a) {
      var beforeCol = findStatusCol_(ds);
      if (beforeCol <= 0) beforeCol = findMessageCol_(ds);
      var insertAt;
      if (beforeCol > 0) { ds.insertColumnBefore(beforeCol); insertAt = beforeCol; }
      else { insertAt = ds.getLastColumn() + 1; }
      ds.getRange(1, insertAt).setValue(a.newField.name);
      a.assignedDataCol = insertAt;
      if (!a.newField.fieldId) a.newField.fieldId = generateFieldId_();
      // 既存フィールド行(byId)の dataCol も、挿入位置以降なら +1 シフトする
      Object.keys(byId).forEach(function(fid) {
        var r = byId[fid];
        if (r.dataCol >= insertAt) r.dataCol += 1;
      });
    });

    // 改名: ヘッダー行を更新
    diff.renamed.forEach(function(r) {
      var col = byId[r.old.fieldId] ? byId[r.old.fieldId].dataCol : 0;
      if (col > 0) ds.getRange(1, col).setValue(r.newField.name);
    });

    // _項目設定 の書き換え: 既存行を削除し、旧行(削除フィールド保持) + 新行(追加/継続) で再構築
    if (flSheet.getLastRow() >= 2) {
      var fd = flSheet.getRange(2, 1, flSheet.getLastRow() - 1, 9).getValues();
      for (var j = fd.length - 1; j >= 0; j--) if (fd[j][0] === wid) flSheet.deleteRow(j + 2);
    }
    // 削除された旧フィールドは deleted=true で残す(dataCol 維持)
    diff.deleted.forEach(function(o) {
      flSheet.appendRow([wid, o.name, o.type, o.options || '', o.required || false, o.order || 0, o.fieldId, true, o.dataCol || 0]);
    });
    // もともと deleted=true だった行もそのまま残す
    oldFieldRows.filter(function(r) { return r.deleted; }).forEach(function(o) {
      flSheet.appendRow([wid, o.name, o.type, o.options || '', o.required || false, o.order || 0, o.fieldId, true, o.dataCol || 0]);
    });
    // 現在有効なフィールド行(継続 + 新規)を書き込む
    newFields.forEach(function(n, idx) {
      var fid = n.fieldId || null;
      var dataCol = 0;
      if (fid && byId[fid]) dataCol = byId[fid].dataCol;
      // 追加フィールドは diff.added 側で dataCol 割当済み
      var addedRec = null;
      diff.added.forEach(function(a) { if (a.newField === n) addedRec = a; });
      if (addedRec) { fid = addedRec.newField.fieldId; dataCol = addedRec.assignedDataCol; }
      if (!fid) fid = generateFieldId_();
      if (!dataCol) dataCol = 3 + (idx + 1); // フォールバック
      flSheet.appendRow([wid, n.name, n.type, n.options || '', n.required || false, idx + 1, fid, false, dataCol]);
    });

    // ボタン設定（列J=type, K=連動ワークフローID, L=スレッド継承）
    ensureButtonsSchema_(btSheet);
    if (btSheet.getLastRow() >= 2) { var bd = btSheet.getRange(2, 1, btSheet.getLastRow() - 1, 12).getValues(); for (var k = bd.length - 1; k >= 0; k--) { if (bd[k][0] === wid) btSheet.deleteRow(k + 2); } }
    if (data.buttons) data.buttons.forEach(function(b, idx2) {
      var btype = (b.type === 'launch') ? 'launch' : 'status';
      var linkedId = 0;
      if (btype === 'launch') {
        linkedId = Number(b.linkedWorkflowId || 0);
        if (!linkedId || linkedId === wid) throw new Error('ボタン「' + (b.name || '') + '」の連動ワークフローが不正です');
      }
      btSheet.appendRow([
        wid, b.name,
        b.status || '',
        b.color || 'gray', idx2 + 1,
        btype === 'launch' ? false : !!b.threadReply,
        btype === 'launch' ? '' : (b.replyMsg || ''),
        false, '',
        btype,
        linkedId || '',
        btype === 'launch' ? !!b.inheritThread : false
      ]);
    });

    // 外部シート
    if (data.externalSheet) {
      try {
        var eid = data.externalSheet.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (eid) {
          var ess = SpreadsheetApp.openById(eid[1]);
          var es = ess.getSheetByName(data.name);
          if (!es) {
            es = ess.insertSheet(data.name, ess.getSheets().length);
            var eh = ['タイムスタンプ', '送信者', '送信者メール'];
            newFields.forEach(function(f) { eh.push(f.name); });
            if (data.type === '申請・承認') eh.push('ステータス');
            es.appendRow(eh);
          }
        }
      } catch (err) { Logger.log('外部シート作成エラー: ' + err.message); }
    }

    // UI の並び順に物理列を揃える。追加/削除/並び替えの後に一度だけ呼ぶ。
    // 内部で dataCol も更新される。
    try { canonicalizeDataSheetLayout_(ss, ds, wid, data.type === '申請・承認'); }
    catch (e) { Logger.log('canonicalize skip: ' + e.message); }

    writeAuditLog_(isNew ? 'workflow.create' : 'workflow.update', 'wid=' + wid, data.name + ' [' + describeFieldDiff_(diff) + ']');
    invalidateRenderCache_();

    // 以前はここで Bot 参加状態 + 外部シート書込権限を同期チェックして warnings を返していたが、
    // 実行スペースが多いワークフローでは数秒~十数秒要してレスポンスが遅かったため非同期化した。
    // クライアントは保存完了後に api_verifyWorkflowSetup(wid) を別途呼び、必要なら警告モーダルを出す。
    return { success: true, id: wid, diff: describeFieldDiff_(diff), warnings: [] };
  } catch (err) { Logger.log('api_saveWorkflow エラー: ' + err.message); try { notifyAdminOnError_(err, 'api_saveWorkflow'); } catch (e2) {} return { success: false, error: err.message }; }
  finally { try { lock.releaseLock(); } catch (e) {} }
}

// 保存後にクライアントが非同期で呼び、Bot参加状態や外部シート接続性を検証して警告を返す。
// 保存クリティカルパスから外すために分離した。
function api_verifyWorkflowSetup(wid) {
  try {
    requireAdmin_();
    var wf = getWorkflowById_(wid);
    if (!wf) return { warnings: [] };
    var warnings = [];
    try {
      // 送信先スペースはカンマ区切り複数対応。全スペースについて Bot 参加チェックする。
      var tgtList = parseSpaceIds_(wf.targetSpace);
      var tgtSeen = {};
      tgtList.forEach(function(sid){
        tgtSeen[sid] = true;
        var vt = verifyBotAccess_(sid);
        if (!vt.ok) warnings.push('送信先スペース(' + sid + ')に Bot が参加していません: ' + vt.error);
      });
      if (wf.execSpace) {
        var execList = parseSpaceIds_(wf.execSpace);
        execList.forEach(function(sid){
          if (tgtSeen[sid]) return; // 送信先と重複する実行スペースは重複チェックしない
          var ve = verifyBotAccess_(sid);
          if (!ve.ok) warnings.push('実行スペース(' + sid + ')に Bot が参加していません: ' + ve.error);
        });
      }
    } catch (e) { warnings.push('スペース参加状態の確認に失敗: ' + e.message); }
    if (wf.externalSheet) {
      try {
        var eidChk = String(wf.externalSheet).match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (eidChk) SpreadsheetApp.openById(eidChk[1]).getSheets().length;
      } catch (e) { warnings.push('外部スプレッドシートを開けません(共有権限/URL 確認): ' + e.message); }
    }
    return { warnings: warnings };
  } catch (err) { return { warnings: [], error: err.message }; }
}

// UI が保存前に差分を確認するために呼ぶ。実際の変更は行わない。
// 戻り値: { added:[name...], deleted:[name...], renamed:[{from,to}...], reordered:bool, hasData:bool, dataRows:number }
function api_previewFieldDiff(payload) {
  try {
    requireAdmin_();
    var wid = payload && payload.id;
    if (!wid) return { added: (payload.fields || []).map(function(f) { return f.name; }), deleted: [], renamed: [], reordered: false, hasData: false, dataRows: 0 };
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureFieldSchema_(ss);
    var oldFields = readFieldRows_(ss, wid).filter(function(r) { return !r.deleted; });
    var diff = computeFieldDiff_(oldFields, payload.fields || []);
    var ds = ss.getSheetByName(payload.name || (getWorkflowById_(wid) || {}).name);
    var dataRows = 0;
    if (ds) dataRows = Math.max(0, ds.getLastRow() - 1);
    return {
      added: diff.added.map(function(a) { return a.newField.name; }),
      deleted: diff.deleted.map(function(d) { return d.name; }),
      renamed: diff.renamed.map(function(r) { return { from: r.old.name, to: r.newField.name }; }),
      reordered: diff.reordered,
      hasData: dataRows > 0,
      dataRows: dataRows
    };
  } catch (err) { return { error: err.message }; }
}

// 論理削除: status を 'archived' にして _設定.archivedAt(12列目) に日時を記録する。
// 30 日経過後に monthlyArchive_ 内の purgeArchivedWorkflows_ で物理削除される。
function api_deleteWorkflow(wid) {
  var lock = LockService.getScriptLock();
  try {
    var caller = requireAdmin_();
    var callerRole = getEffectiveUserRole_(caller);
    var wfChk = getWorkflowById_(wid);
    if (wfChk && !canDeleteWorkflow_(caller, wfChk, callerRole)) return { success: false, error: 'このワークフローを削除する権限がありません' };
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
    var caller = requireAdmin_();
    var callerRole = getEffectiveUserRole_(caller);
    var wfChk = getWorkflowById_(wid);
    if (wfChk && !canEditWorkflow_(caller, wfChk, callerRole)) return { success: false, error: 'このワークフローの有効/無効を切り替える権限がありません' };
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_SETTINGS);
    if (sheet.getLastRow() < 2) return { success: false };
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
    for (var i = 0; i < rows.length; i++) { if (rows[i][0] === wid) { var ns = (rows[i][7] === '有効') ? '無効' : '有効'; sheet.getRange(i + 2, 8).setValue(ns); writeAuditLog_('workflow.toggle', 'wid=' + wid, ns); invalidateRenderCache_(); return { success: true, status: ns }; } }
    return { success: false };
  } catch (err) { Logger.log('api_toggleStatus エラー: ' + err.message); try { notifyAdminOnError_(err, 'api_toggleStatus'); } catch (e2) {} return { success: false, error: err.message }; }
}

// admins は文字列配列(後方互換) または {email, role} 配列を受け付ける。
// owner ロールのみ実行可。少なくとも1名 owner を残し、呼び出し元自身を外せない。
function api_saveAdmins(admins) {
  try {
    var caller = requireAdmin_();
    var callerRole = getEffectiveUserRole_(caller);
    if (!canManageAdmins_(callerRole)) return { success: false, error: 'この操作はオーナー権限が必要です' };
    if (!admins || !Array.isArray(admins) || admins.length === 0) return { success: false, error: '管理者を1人以上指定してください' };
    var cleaned = admins.map(function(a) {
      var isObj = a && typeof a === 'object';
      var email = normalizeEmail_(isObj ? a.email : a);
      var role = String((isObj && a.role) || ROLE_OWNER).toLowerCase();
      if (role !== ROLE_EDITOR && role !== ROLE_MEMBER && role !== ROLE_OWNER) role = ROLE_OWNER;
      return email ? { email: email, role: role } : null;
    }).filter(function(a) { return a; });
    for (var i = 0; i < cleaned.length; i++) {
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(cleaned[i].email)) return { success: false, error: '不正なメール形式: ' + cleaned[i].email };
    }
    if (!cleaned.some(function(a) { return a.role === ROLE_OWNER; })) return { success: false, error: 'オーナーを最低1名残してください' };
    var callerNorm = normalizeEmail_(caller);
    if (!cleaned.some(function(a) { return a.email === callerNorm; })) return { success: false, error: '自分自身を管理者から外すことはできません' };
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ensureAdminSchema_(ss);
    var sheet = ss.getSheetByName(SHEET_ADMINS);
    if (!sheet) { sheet = ss.insertSheet(SHEET_ADMINS); sheet.appendRow(['メールアドレス', '同意日時', '権限']); }
    // 既存の 同意日時 を保持しつつ書き換える
    var existing = {};
    if (sheet.getLastRow() >= 2) {
      var old = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
      old.forEach(function(r) { var e = normalizeEmail_(r[0]); if (e) existing[e] = r[1] || ''; });
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clearContent();
    }
    cleaned.forEach(function(a, idx) {
      sheet.getRange(idx + 2, 1).setValue(a.email);
      sheet.getRange(idx + 2, 2).setValue(existing[a.email] || '');
      sheet.getRange(idx + 2, 3).setValue(a.role);
    });
    writeAuditLog_('admins.save', '', cleaned.map(function(a) { return a.email + '(' + a.role + ')'; }).join(', '));
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
    var execList = String(wf.execSpace || '').split(',').map(function(s){return s.trim();}).filter(function(s){return s;});
    if (!execList.length) return { success: false, error: '実行スペースが設定されていません' };
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var st = ss.getSheetByName(SHEET_SETTINGS);
    var rows = st.getRange(2, 1, st.getLastRow() - 1, 11).getValues();
    // 既存のランチャーカード(カンマ区切り複数の場合あり)を全て削除
    var t = getServiceAccountToken_();
    for (var i = 0; i < rows.length; i++) {
      if (rows[i][0] === wid && rows[i][8]) {
        String(rows[i][8]).split(',').map(function(s){return s.trim();}).filter(function(s){return s;}).forEach(function(mn){
          try { UrlFetchApp.fetch('https://chat.googleapis.com/v1/' + mn, { method: 'delete', headers: { 'Authorization': 'Bearer ' + t }, muteHttpExceptions: true }); } catch (e) {}
        });
        break;
      }
    }
    var label = (wf.type === '申請・承認') ? '申請する' : '記録する';
    var icon = (wf.type === '申請・承認') ? '📋' : '📝';
    var postedNames = [];
    var errors = [];
    execList.forEach(function(sid){
      try { addBotToSpace_(sid); } catch (e) {}
      try {
        var _hdr = { title: icon + ' ' + wf.name };
        if (wf.description) _hdr.subtitle = wf.description;
        var resp = UrlFetchApp.fetch('https://chat.googleapis.com/v1/' + sid + '/messages', { method: 'post', contentType: 'application/json', headers: { 'Authorization': 'Bearer ' + t }, payload: JSON.stringify({ cardsV2: [{ cardId: 'launcher_' + wid, card: { header: _hdr, sections: [{ widgets: [{ buttonList: { buttons: [{ text: label, color: COLOR_MAP.blue, onClick: { action: { function: cardActionFunc_(), interaction: 'OPEN_DIALOG', parameters: [{ key: 'actionName', value: 'openWorkflow' }, { key: 'workflowId', value: String(wid) }] } } }] } }] }] } }] }), muteHttpExceptions: true });
        var r = JSON.parse(resp.getContentText());
        if (r.name) postedNames.push(r.name);
        else errors.push(sid + ': ' + resp.getContentText());
      } catch (e) { errors.push(sid + ': ' + e.message); }
    });
    if (postedNames.length) {
      for (var j = 0; j < rows.length; j++) { if (rows[j][0] === wid) { st.getRange(j + 2, 9).setValue(postedNames.join(',')); break; } }
      writeAuditLog_('workflow.deploy', 'wid=' + wid, wf.name + ' (' + postedNames.length + '/' + execList.length + ')');
      if (errors.length) return { success: true, warning: '一部スペースへの投稿に失敗: ' + errors.join(' / ') };
      return { success: true };
    }
    return { success: false, error: errors.join(' / ') || '投稿に失敗しました' };
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
