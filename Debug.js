// ============================================================
// Debug.gs - 権限ロール別の挙動シミュレーション
// Apps Script エディタから runRoleDebugSuite() を手動実行して使う。
// 読み取り系のみ(mutating API は呼ばない)。実行者本人の UserCache を
// 一時的に書き換えるだけなので他ユーザーには影響しない。
// ============================================================

function runRoleDebugSuite() {
  var email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!email) throw new Error('実行者メールが取得できません');
  var realRole = getUserRole_(email);
  if (realRole !== ROLE_OWNER) throw new Error('このデバッグ関数は実オーナーのみ実行可能です(現在: ' + realRole + ')');

  var cache = CacheService.getUserCache();
  var saved = getDebugRoleOverride_();
  var roles = [ROLE_OWNER, ROLE_EDITOR, ROLE_MEMBER];
  var summary = [];
  var perRole = {};

  try {
    roles.forEach(function(r) {
      _setOverride_(cache, r);
      var effective = getEffectiveUserRole_(email);
      var list = getAllWorkflowsForRender_();
      perRole[r] = list;
      summary.push({
        asRole: r,
        effective: effective,
        visible: list.length,
        editable: list.filter(function(w){return w.canEdit;}).length,
        deletable: list.filter(function(w){return w.canDelete;}).length,
        canSeeMasterSheet: effective === ROLE_OWNER,
        canManageAdmins: canManageAdmins_(effective),
        canArchive: _simulateCanArchive_(effective),
        canToggleStatus: _simulateCanToggleStatus_(effective)
      });
    });
  } finally {
    // 必ず元に戻す
    if (saved) cache.put(DEBUG_ROLE_CACHE_KEY, saved, DEBUG_ROLE_TTL);
    else cache.remove(DEBUG_ROLE_CACHE_KEY);
  }

  _writeDebugReport_(email, realRole, summary, perRole);
  Logger.log('✅ デバッグスイート完了。スプレッドシートの「_デバッグレポート」シートを確認してください。');
  return { ok: true, roles: summary };
}

function _setOverride_(cache, role) {
  if (role === ROLE_OWNER) cache.remove(DEBUG_ROLE_CACHE_KEY);
  else cache.put(DEBUG_ROLE_CACHE_KEY, role, DEBUG_ROLE_TTL);
}

// Adminapi.js の権限判定ロジックと同じ条件を評価する(実際に api_* を呼ばない)。
function _simulateCanArchive_(role) { return role === ROLE_OWNER || role === ROLE_EDITOR; }
function _simulateCanToggleStatus_(role) { return role === ROLE_OWNER || role === ROLE_EDITOR; }

function _writeDebugReport_(email, realRole, summary, perRole) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('_デバッグレポート');
  if (!sh) sh = ss.insertSheet('_デバッグレポート');
  sh.clear();

  sh.appendRow(['=== ロール別権限シミュレーション ===']);
  sh.appendRow(['生成日時', new Date()]);
  sh.appendRow(['実行者', email]);
  sh.appendRow(['実権限', realRole]);
  sh.appendRow(['バージョン', VERSION]);
  sh.appendRow(['']);

  sh.appendRow(['--- サマリ ---']);
  sh.appendRow(['ロール', '可視WF数', '編集可', '削除可', '管理データ閲覧', '管理者編集', 'アーカイブ可', 'ON/OFF切替可']);
  summary.forEach(function(r) {
    sh.appendRow([
      r.asRole,
      r.visible, r.editable, r.deletable,
      r.canSeeMasterSheet ? '○' : '×',
      r.canManageAdmins   ? '○' : '×',
      r.canArchive        ? '○' : '×',
      r.canToggleStatus   ? '○' : '×'
    ]);
  });

  sh.appendRow(['']);
  sh.appendRow(['--- WF別 可視/編集/削除マトリクス ---']);
  sh.appendRow(['WF ID', 'WF名', '作成者', '共同編集者',
                'owner可視', 'editor可視', 'member可視',
                'owner編集', 'editor編集', 'member編集',
                'owner削除', 'editor削除', 'member削除']);
  var ownerList = perRole[ROLE_OWNER] || [];
  var byId = {};
  [ROLE_OWNER, ROLE_EDITOR, ROLE_MEMBER].forEach(function(r) {
    byId[r] = {};
    (perRole[r] || []).forEach(function(w) { byId[r][w.id] = w; });
  });
  var mark = function(w, prop) { return (w && w[prop]) ? '○' : '×'; };
  var vis = function(role, id) { return byId[role][id] ? '○' : '×'; };
  ownerList.forEach(function(w) {
    sh.appendRow([
      w.id, w.name, w.creator, w.collaborators || '',
      vis(ROLE_OWNER, w.id),  vis(ROLE_EDITOR, w.id), vis(ROLE_MEMBER, w.id),
      mark(byId[ROLE_OWNER][w.id], 'canEdit'), mark(byId[ROLE_EDITOR][w.id], 'canEdit'), mark(byId[ROLE_MEMBER][w.id], 'canEdit'),
      mark(byId[ROLE_OWNER][w.id], 'canDelete'), mark(byId[ROLE_EDITOR][w.id], 'canDelete'), mark(byId[ROLE_MEMBER][w.id], 'canDelete')
    ]);
  });

  sh.appendRow(['']);
  sh.appendRow(['--- 期待される挙動との差分チェック ---']);
  sh.appendRow(['項目', '結果', '備考']);
  var ownerSum = summary[0], editorSum = summary[1], memberSum = summary[2];
  var checks = [
    ['owner は全WF可視',               ownerSum.visible >= (editorSum.visible || 0),           '通常 owner >= editor'],
    ['editor も原則全WF可視(role定義)', editorSum.visible === ownerSum.visible,                 'editor は view 全可'],
    ['member は作成者/共同編集者のみ可視', memberSum.visible <= ownerSum.visible,                 '限定表示'],
    ['owner のみ管理者編集可',          ownerSum.canManageAdmins && !editorSum.canManageAdmins && !memberSum.canManageAdmins, ''],
    ['owner のみ管理データ閲覧',        ownerSum.canSeeMasterSheet && !editorSum.canSeeMasterSheet && !memberSum.canSeeMasterSheet, ''],
    ['member は削除不可',               memberSum.deletable === 0,                               '削除は owner のみ']
  ];
  checks.forEach(function(c) { sh.appendRow([c[0], c[1] ? '✅ OK' : '❌ NG', c[2]]); });

  sh.setFrozenRows(1);
  try { sh.autoResizeColumns(1, 13); } catch (e) {}
}
