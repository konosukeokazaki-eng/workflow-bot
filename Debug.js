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

// ============================================================
// runEnvironmentDiagnostics() - 認証/データ/Chat/整合性の環境バグ検出
// 読み取り系のみ。実書き込みは行わない(セル書換えテストもロールバックする)。
// Apps Script エディタから実オーナーが手動で実行する。
// ============================================================
function runEnvironmentDiagnostics() {
  var email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!email) throw new Error('実行者メールが取得できません');
  var realRole = getUserRole_(email);
  if (realRole !== ROLE_OWNER) throw new Error('このデバッグ関数は実オーナーのみ実行可能です(現在: ' + realRole + ')');

  var results = [];
  var add = function(category, name, ok, detail, fix) {
    results.push({ category: category, name: name, ok: !!ok, detail: detail || '', fix: fix || '' });
  };

  // ---------- 認証系 ----------
  add('認証', 'Session.getActiveUser().getEmail() が取れる', !!email, email,
    '取れない場合: Chat 経由では空になることがある。commonEventObject.user.email の併用を検討。');

  var oauthOk = false, oauthDetail = '';
  try { var t = ScriptApp.getOAuthToken(); oauthOk = !!t; oauthDetail = 'len=' + (t ? t.length : 0); }
  catch (e) { oauthDetail = e.message; }
  add('認証', 'ScriptApp.getOAuthToken() 取得', oauthOk, oauthDetail, '取れない場合: スコープ再承認が必要。');

  var saOk = false, saDetail = '';
  try { getServiceAccountToken_(); saOk = true; saDetail = 'ok'; }
  catch (e) { saDetail = e.message; }
  add('認証', 'サービスアカウントトークン', saOk, saDetail,
    'NG の場合: ScriptProperty SERVICE_ACCOUNT_KEY を確認。鍵が失効/ローテされた可能性。');

  var dwdOk = false, dwdDetail = '';
  try {
    var it = getImpersonatedToken_(email, 'https://www.googleapis.com/auth/chat.spaces.readonly');
    dwdOk = !!it; dwdDetail = 'impersonate=' + email + ' ok';
  } catch (e) { dwdDetail = e.message; }
  add('認証', 'DWD impersonation (chat.spaces.readonly)', dwdOk, dwdDetail,
    'NG の場合: Workspace 管理コンソールで SA クライアント ID に該当スコープを委任。これが NG だと「デプロイ者のスペースだけ表示」バグが出る。');

  // 本当にユーザー自身の見えるスペース一覧が取れるか(= デプロイ者スペースだけバグの検出)
  var spacesOk = false, spacesDetail = '';
  try {
    var it2 = getImpersonatedToken_(email, 'https://www.googleapis.com/auth/chat.spaces.readonly');
    var resp = UrlFetchApp.fetch('https://chat.googleapis.com/v1/spaces?pageSize=1', {
      method: 'get', headers: { 'Authorization': 'Bearer ' + it2 }, muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    spacesOk = (code === 200);
    spacesDetail = 'HTTP ' + code + ' ' + resp.getContentText().substring(0, 150);
  } catch (e) { spacesDetail = e.message; }
  add('認証', 'spaces.list (impersonated)', spacesOk, spacesDetail,
    'NG だと UI の「スペース選択」が実行者本人の見えるスペースを引けず、代替で deployer のスペースを拾う挙動になる。');

  // ---------- データ系 ----------
  var ssOk = false, ssDetail = '';
  var ss = null;
  try { ss = SpreadsheetApp.openById(SPREADSHEET_ID); ssOk = true; ssDetail = ss.getName(); }
  catch (e) { ssDetail = e.message; }
  add('データ', 'スプレッドシートを開ける', ssOk, ssDetail,
    'NG: ScriptProperty SPREADSHEET_ID / 権限設定を確認。');

  if (ss) {
    [SHEET_SETTINGS, SHEET_FIELDS, SHEET_BUTTONS, SHEET_ADMINS, SHEET_AUDIT, SHEET_TOKENS].forEach(function(n) {
      var sh = ss.getSheetByName(n);
      add('データ', 'シート存在: ' + n, !!sh, sh ? ('rows=' + sh.getLastRow()) : 'なし',
        !sh ? '初回アクセス時に自動生成される。手動で /admin を開くか ensure*Schema_ を実行。' : '');
    });

    // setValue readback: 「ステータスが変わらない」バグ検出
    var tmp = ss.getSheetByName('_診断一時') || ss.insertSheet('_診断一時');
    var writeOk = false, writeDetail = '';
    try {
      var marker = 'diag_' + Date.now();
      tmp.getRange('A1').setValue(marker);
      SpreadsheetApp.flush();
      var readBack = tmp.getRange('A1').getValue();
      writeOk = (readBack === marker);
      writeDetail = 'wrote=' + marker + ' read=' + readBack;
      tmp.getRange('A1').clearContent();
    } catch (e) { writeDetail = e.message; }
    add('データ', 'setValue → flush → 読み戻し一致', writeOk, writeDetail,
      'NG: 書き込みが反映されない根深いバグ。SpreadsheetApp.flush() 位置、権限、同時編集ロックを確認。');
    try { ss.deleteSheet(tmp); } catch (e) {}
  }

  // LockService 取得可否
  var lockOk = false, lockDetail = '';
  try {
    var lk = LockService.getScriptLock();
    lk.tryLock(100);
    if (lk.hasLock()) { lockOk = true; lockDetail = 'acquired'; lk.releaseLock(); }
    else lockDetail = 'tryLock timeout (他プロセスが保持中かも)';
  } catch (e) { lockDetail = e.message; }
  add('データ', 'LockService 取得', lockOk, lockDetail,
    'NG: 保存処理が待たされ「処理中」が長引く原因。');

  // ---------- Chat 連携 ----------
  add('Chat', 'FUNCTION_URL 設定', !!FUNCTION_URL, FUNCTION_URL || '(未設定)',
    '未設定だと HTTP モードでカードボタンが機能しない。');
  add('Chat', 'CONTACT_EMAIL 設定', !!CONTACT_EMAIL, CONTACT_EMAIL || '(未設定)', '');
  add('Chat', 'CHAT_CONNECTION_MODE', CHAT_CONNECTION_MODE === 'http' || CHAT_CONNECTION_MODE === 'apps_script',
    CHAT_CONNECTION_MODE, 'http か apps_script のみ有効。');

  // ---------- 整合性 ----------
  if (ss) {
    var allWfs = (function() { try { return getAllWorkflowsForRender_() || []; } catch (e) { return []; } })();
    // 承認型なのに承認者が空
    var noApprover = allWfs.filter(function(w) { return w.type === '申請・承認' && !(w.approvers && String(w.approvers).trim()); });
    add('整合性', '承認型WFの承認者未設定', noApprover.length === 0,
      noApprover.length === 0 ? '0件' : noApprover.map(function(w){return '#'+w.id+' '+w.name;}).join(', '),
      '承認者が空だと承認カードを誰にも送れない。');

    // ボタンの linkedWorkflowId が壊れていないか
    var brokenLinks = [];
    allWfs.forEach(function(w) {
      try {
        var btns = getWorkflowButtons_(w.id);
        btns.forEach(function(b) {
          if (b.type === 'launch' && b.linkedWorkflowId) {
            var child = getWorkflowById_(b.linkedWorkflowId);
            if (!child) brokenLinks.push('WF#' + w.id + ' btn=' + b.name + ' → #' + b.linkedWorkflowId + ' が存在しない');
          }
        });
      } catch (e) {}
    });
    add('整合性', '起動ボタンのlinkedWorkflowIdが有効', brokenLinks.length === 0,
      brokenLinks.length === 0 ? '0件' : brokenLinks.join(' / '),
      '子WFが削除/アーカイブされると親の起動ボタンが機能しなくなる。該当ボタンを削除するか子WFを復元。');

    // 管理者シートに owner が1人も居ない
    var adSh = ss.getSheetByName(SHEET_ADMINS);
    var ownerExists = false;
    if (adSh && adSh.getLastRow() >= 2) {
      var rows = adSh.getRange(2, 1, adSh.getLastRow() - 1, Math.max(adSh.getLastColumn(), 3)).getValues();
      ownerExists = rows.some(function(r) {
        var role = String(r[2] || '').trim().toLowerCase();
        return r[0] && (role === ROLE_OWNER || role === '');
      });
    }
    add('整合性', 'owner が最低1人存在', ownerExists, ownerExists ? 'ok' : 'owner不在',
      'NG: /admin へ誰もアクセスできなくなる。SHEET_ADMINS に手動で owner を追加。');

    // 外部シート参照の到達性(到達不能なら警告のみ)
    var extNG = [];
    allWfs.forEach(function(w) {
      if (!w.externalSheet) return;
      var m = String(w.externalSheet).match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (!m) { extNG.push('WF#' + w.id + ' URL形式不正'); return; }
      try { SpreadsheetApp.openById(m[1]); }
      catch (e) { extNG.push('WF#' + w.id + ' ' + e.message.substring(0, 60)); }
    });
    add('整合性', '外部スプレッドシートに全てアクセス可', extNG.length === 0,
      extNG.length === 0 ? '0件' : extNG.join(' / '),
      '外部シートに書き込めない WF がある。共有設定を見直すか externalSheet を外す。');
  }

  _writeEnvDiagReport_(email, results);
  var ng = results.filter(function(r){return !r.ok;}).length;
  Logger.log('✅ 環境診断完了。NG=' + ng + ' / 全' + results.length + '件。「_環境診断レポート」シートを確認してください。');
  return { ok: ng === 0, total: results.length, ng: ng, results: results };
}

// ============================================================
// runButtonBehaviorDiagnostics() - 各WFのボタン挙動を実データ非破壊で検証
// - /apply コマンド投入時のカード生成
// - openWorkflow ダイアログの生成(全フィールド含有)
// - submitWorkflow に空入力 → 必須バリデーションが働くか(DB書込に到達する前の経路)
// - 承認トークン発行→検証 ラウンドトリップ
// - 無効トークンでの handleApproval_ がエラーページを返すか
// - 起動ボタンの linkedWorkflowId からダイアログを生成できるか
// 書き込み・Chat API 送信は一切行わない。
// ============================================================
function runButtonBehaviorDiagnostics() {
  var email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!email) throw new Error('実行者メールが取得できません');
  var realRole = getUserRole_(email);
  if (realRole !== ROLE_OWNER) throw new Error('このデバッグ関数は実オーナーのみ実行可能です(現在: ' + realRole + ')');

  var results = [];
  var add = function(wf, scenario, ok, detail, fix) {
    results.push({ wfId: wf ? wf.id : '', wfName: wf ? wf.name : '-', scenario: scenario, ok: !!ok, detail: detail || '', fix: fix || '' });
  };

  var wfs = getActiveWorkflows_();
  if (wfs.length === 0) {
    _writeButtonDiagReport_(email, results, '有効なWFがありません');
    return { ok: true, total: 0 };
  }

  // --- 共通: 擬似イベント ---
  var fakeSpace = 'spaces/__DIAG__';
  var fakeUser = { displayName: '診断ユーザー', email: email };

  wfs.forEach(function(wf) {
    // 1) /apply ダイアログ生成
    try {
      var dlg = buildWorkflowDialog_(wf.id);
      var hasCard = !!(dlg && dlg.action && dlg.action.navigations && dlg.action.navigations[0] &&
                       (dlg.action.navigations[0].pushCard || dlg.action.navigations[0].updateCard));
      var card = dlg && dlg.action && dlg.action.navigations && dlg.action.navigations[0] &&
                 (dlg.action.navigations[0].pushCard || dlg.action.navigations[0].updateCard);
      var widgets = card && card.sections && card.sections[0] && card.sections[0].widgets || [];
      var hasSubmit = widgets.some(function(w) {
        return w.buttonList && w.buttonList.buttons && w.buttonList.buttons.some(function(b) {
          return b.onClick && b.onClick.action && b.onClick.action.parameters &&
                 b.onClick.action.parameters.some(function(p){ return p.key==='actionName' && p.value==='submitWorkflow'; });
        });
      });
      add(wf, 'ダイアログ生成', hasCard && hasSubmit,
        'widgets=' + widgets.length + ' submit=' + hasSubmit,
        !hasCard ? 'buildWorkflowDialog_ がカードを返していない' : (!hasSubmit ? '送信ボタンが欠落' : ''));
    } catch (e) {
      add(wf, 'ダイアログ生成', false, 'throw: ' + e.message, 'buildWorkflowDialog_ が例外');
    }

    // 2) フィールドがダイアログに全て含まれているか
    try {
      var fields = getWorkflowFields_(wf.id);
      var dlg2 = buildWorkflowDialog_(wf.id);
      var card2 = dlg2.action.navigations[0].pushCard || dlg2.action.navigations[0].updateCard;
      var widgets2 = card2.sections[0].widgets;
      var missing = [];
      fields.forEach(function(f) {
        if (f.type === 'checkbox_qty') {
          var found = widgets2.some(function(w) { return w.textParagraph && String(w.textParagraph.text || '').indexOf(f.name) !== -1; });
          if (!found) missing.push(f.name);
        } else {
          var key = 'field_' + f.order;
          var found2 = widgets2.some(function(w) {
            return (w.textInput && w.textInput.name === key) ||
                   (w.selectionInput && w.selectionInput.name === key) ||
                   (w.dateTimePicker && w.dateTimePicker.name === key);
          });
          if (!found2) missing.push(f.name);
        }
      });
      add(wf, '全フィールド配置', missing.length === 0,
        missing.length === 0 ? ('fields=' + fields.length) : ('欠落: ' + missing.join(',')),
        'fieldId/order のズレか type 未対応の可能性');
    } catch (e) {
      add(wf, '全フィールド配置', false, 'throw: ' + e.message, '');
    }

    // 3) 空入力送信 → 必須バリデーションが働くか(DB書込に到達しない)
    try {
      var requiredFields = getWorkflowFields_(wf.id).filter(function(f){return f.required;});
      if (requiredFields.length === 0) {
        add(wf, '必須バリデーション', true, '必須項目なし(スキップ)', '');
      } else {
        var fakeEvent = {
          chat: { user: fakeUser, buttonClickedPayload: { space: { name: fakeSpace } } },
          commonEventObject: {
            parameters: { actionName: 'submitWorkflow', workflowId: String(wf.id) },
            formInputs: {} // 空
          }
        };
        var resp = handleSubmitWorkflow_(fakeEvent);
        // バリデーション失敗なら updateCard が返るはず(エラー文言付き)
        var nav = resp && resp.action && resp.action.navigations && resp.action.navigations[0];
        var errCard = nav && nav.updateCard;
        var errText = '';
        if (errCard) {
          var ews = errCard.sections[0].widgets;
          var head = ews[0] && ews[0].textParagraph && ews[0].textParagraph.text || '';
          errText = head;
        }
        var ok = !!errCard && errText.indexOf('必須項目') !== -1;
        add(wf, '必須バリデーション', ok,
          ok ? ('必須=' + requiredFields.length + ' エラー表示OK') : ('応答=' + JSON.stringify(resp).substring(0,120)),
          '空入力でDB書込に進んでしまうと重大。handleSubmitWorkflow_ の required チェックを確認。');
      }
    } catch (e) {
      add(wf, '必須バリデーション', false, 'throw: ' + e.message, '');
    }

    // 4) 承認型のみ: トークン発行→検証
    if (wf.type === '申請・承認') {
      try {
        var tok = generateToken_(wf.id, 2);
        var v1 = validateToken_(tok, wf.id, 2);
        var v2 = validateToken_('invalid_token_xxx', wf.id, 2);
        var v3 = validateToken_(tok, wf.id, 99); // 他行で使えないこと
        var ok = v1.ok && !v2.ok && !v3.ok;
        add(wf, 'トークン発行→検証', ok,
          'self=' + v1.ok + ' invalid=' + v2.ok + ' crossRow=' + v3.ok,
          !ok ? 'トークン検証ロジックに問題。validateToken_ を確認。' : '');
      } catch (e) {
        add(wf, 'トークン発行→検証', false, 'throw: ' + e.message, '');
      }

      // 5) 無効トークンで handleApproval_ がエラーページ返すか
      try {
        var fakeE = { parameter: { token: 'bogus', wf: String(wf.id), row: '2', btn: '0', name: encodeURIComponent('test') } };
        var html = handleApproval_(fakeE);
        var htmlStr = html.getContent ? html.getContent() : String(html);
        var ok2 = htmlStr.indexOf('無効なリクエスト') !== -1 || htmlStr.indexOf('🚫') !== -1;
        add(wf, '無効トークン拒否', ok2, '応答長=' + htmlStr.length,
          !ok2 ? 'bogus トークンで承認されてしまうかも。' : '');
      } catch (e) {
        add(wf, '無効トークン拒否', false, 'throw: ' + e.message, '');
      }

      // 6) 承認者未設定の警告(= 整合性チェックと重複だが、WF単位で可視化)
      var approverStr = String(wf.approvers || '').trim();
      add(wf, '承認者設定', !!approverStr, approverStr || '(空)',
        !approverStr ? '承認カードを誰にも送れない。管理画面で承認者を追加。' : '');
    }

    // 7) 起動ボタン: linkedWorkflowId → ダイアログ生成可能か
    try {
      var btns = getWorkflowButtons_(wf.id);
      var launches = btns.filter(function(b){return b.type === 'launch';});
      if (launches.length === 0) {
        add(wf, '起動ボタン', true, '起動ボタンなし(スキップ)', '');
      } else {
        var ngs = [];
        launches.forEach(function(b) {
          if (!b.linkedWorkflowId) { ngs.push('"' + b.name + '"→未設定'); return; }
          try {
            var child = buildWorkflowDialog_(b.linkedWorkflowId);
            var hasCard = !!(child && child.action && child.action.navigations && child.action.navigations[0] &&
                             (child.action.navigations[0].pushCard || child.action.navigations[0].updateCard));
            if (!hasCard) ngs.push('"' + b.name + '"→#' + b.linkedWorkflowId + ' カード生成失敗');
          } catch (e) { ngs.push('"' + b.name + '"→throw ' + e.message); }
        });
        add(wf, '起動ボタン', ngs.length === 0,
          ngs.length === 0 ? ('OK ' + launches.length + '件') : ngs.join(' / '),
          '子WFが削除/壊れている可能性');
      }
    } catch (e) {
      add(wf, '起動ボタン', false, 'throw: ' + e.message, '');
    }

    // 8) /apply コマンド経由で可視か (execSpace フィルタ)
    try {
      var hasExec = String(wf.execSpace || '').trim().length > 0;
      add(wf, '実行スペース設定', hasExec, wf.execSpace || '(未設定)',
        !hasExec ? '/apply コマンドでスペース絞り込みが機能しない。空=全スペース扱い。' : '');
    } catch (e) {}
  });

  _writeButtonDiagReport_(email, results, null);
  var ng = results.filter(function(r){return !r.ok;}).length;
  Logger.log('✅ ボタン挙動診断完了。NG=' + ng + ' / 全' + results.length + '件。「_ボタン挙動レポート」シート参照。');
  return { ok: ng === 0, total: results.length, ng: ng };
}

function _writeButtonDiagReport_(email, results, note) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('_ボタン挙動レポート');
  if (!sh) sh = ss.insertSheet('_ボタン挙動レポート');
  sh.clear();
  sh.appendRow(['=== ボタン挙動シミュレーションレポート ===']);
  sh.appendRow(['生成日時', new Date()]);
  sh.appendRow(['実行者', email]);
  sh.appendRow(['バージョン', VERSION]);
  if (note) sh.appendRow(['備考', note]);
  sh.appendRow(['']);
  sh.appendRow(['WF ID', 'WF名', 'シナリオ', '結果', '詳細', '対処ヒント']);
  results.forEach(function(r) {
    sh.appendRow([r.wfId, r.wfName, r.scenario, r.ok ? '✅ OK' : '❌ NG', r.detail, r.fix]);
  });
  sh.setFrozenRows(note ? 7 : 6);
  try { sh.autoResizeColumns(1, 6); } catch (e) {}
}

// 全診断を一括実行するランチャー
function runAllDiagnostics() {
  var out = { roleSuite: null, env: null, button: null };
  try { out.roleSuite = runRoleDebugSuite(); } catch (e) { out.roleSuite = { error: e.message }; }
  try { out.env       = runEnvironmentDiagnostics(); } catch (e) { out.env = { error: e.message }; }
  try { out.button    = runButtonBehaviorDiagnostics(); } catch (e) { out.button = { error: e.message }; }
  Logger.log('=== runAllDiagnostics ===\n' + JSON.stringify(out, null, 2));
  return out;
}

function _writeEnvDiagReport_(email, results) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('_環境診断レポート');
  if (!sh) sh = ss.insertSheet('_環境診断レポート');
  sh.clear();
  sh.appendRow(['=== 環境診断レポート ===']);
  sh.appendRow(['生成日時', new Date()]);
  sh.appendRow(['実行者', email]);
  sh.appendRow(['バージョン', VERSION]);
  sh.appendRow(['']);
  sh.appendRow(['カテゴリ', '項目', '結果', '詳細', '対処ヒント']);
  results.forEach(function(r) {
    sh.appendRow([r.category, r.name, r.ok ? '✅ OK' : '❌ NG', r.detail, r.fix]);
  });
  sh.setFrozenRows(6);
  try { sh.autoResizeColumns(1, 5); } catch (e) {}
}
