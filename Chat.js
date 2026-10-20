// ============================================================
// Chat.gs - Google Chat 連携（コマンド、カード送信、承認処理）
// ============================================================

// 数量付きチェックボックスの options を "<選択肢>###<qtype>###<数量候補>" から構造体に分解する。
function parseQtyOptions_(raw) {
  var s = String(raw || '');
  var parts = s.split('###');
  return { opts: parts[0] || '', qtype: parts[1] || 'number', qopts: parts[2] || '' };
}

function handleApplyCommand_(event) {
  var workflows = getActiveWorkflows_();
  var currentSpace = '';
  if (event.chat && event.chat.appCommandPayload && event.chat.appCommandPayload.space) currentSpace = event.chat.appCommandPayload.space.name || '';
  if (currentSpace) workflows = workflows.filter(function(wf) { return String(wf.execSpace || '').split(',').map(function(s){return s.trim();}).indexOf(currentSpace) !== -1; });
  if (workflows.length === 0) return { action: { navigations: [{ pushCard: { header: { title: 'ワークフロー' }, sections: [{ widgets: [{ textParagraph: { text: 'このスペースに紐づいたワークフローはありません。' } }] }] } }] } };
  if (workflows.length === 1) return buildWorkflowDialog_(workflows[0].id);
  var items = workflows.map(function(wf) { return { text: wf.name, value: String(wf.id), selected: false }; });
  items[0].selected = true;
  return { action: { navigations: [{ pushCard: { header: { title: '📋 ワークフロー選択', subtitle: '実行するワークフローを選んでください' }, sections: [{ widgets: [{ selectionInput: { label: 'ワークフロー', name: 'workflowId', type: 'DROPDOWN', items: items } }, { buttonList: { buttons: [{ text: '次へ', onClick: { action: { function: cardActionFunc_(), parameters: [{ key: 'actionName', value: 'openWorkflow' }] } } }] } }] }] } }] } };
}

function openWorkflowDialog_(event) {
  var inputs = (event.commonEventObject && event.commonEventObject.formInputs) || {};
  var params = (event.commonEventObject && event.commonEventObject.parameters) || {};
  var wid = '';
  if (inputs.workflowId && inputs.workflowId.stringInputs) wid = inputs.workflowId.stringInputs.value[0];
  if (!wid && params.workflowId) wid = params.workflowId;
  if (!wid) return createMessage_('ワークフローが選択されていません。');
  var chain = parseInt(params.chain || '1', 10);
  var inheritFrom = params.inheritThreadFrom || '';
  // 連動ボタンから直接呼ばれる場合は親WF/行指定で親メッセージ名を解決する
  if (!inheritFrom && params.parentWorkflowId && params.parentRow) {
    try {
      var pwf = getWorkflowById_(parseInt(params.parentWorkflowId, 10));
      if (pwf) {
        var pss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var pds = pss.getSheetByName(pwf.name);
        if (pds) {
          var phdrs = pds.getRange(1, 1, 1, pds.getLastColumn()).getValues()[0];
          for (var pi = 0; pi < phdrs.length; pi++) {
            if (phdrs[pi] === 'メッセージ名') { inheritFrom = pds.getRange(parseInt(params.parentRow, 10), pi + 1).getValue() || ''; break; }
          }
        }
      }
    } catch (e) { Logger.log('parent resolve skip: ' + e.message); }
  }
  if (chain > 3) return createMessage_('連鎖上限(3階層)を超えています。');
  return buildWorkflowDialog_(parseInt(wid, 10), chain, inheritFrom);
}

function buildWorkflowDialog_(workflowId, chain, inheritFrom) {
  return buildWorkflowDialogCard_(workflowId, null, null, 'push', chain, inheritFrom);
}

// 共通のカード生成。prefill(name -> string値) と errors(必須未入力の項目名配列) を渡すと
// 入力値を保持したまま再描画し、エラー行を先頭に挿入する。
// navMode: 'push'(初回) or 'update'(再描画)
function buildWorkflowDialogCard_(workflowId, prefill, errors, navMode, chain, inheritFrom) {
  var wf = getWorkflowById_(workflowId);
  if (!wf) return createMessage_('ワークフローが見つかりません');
  var fields = getWorkflowFields_(workflowId);
  var widgets = [];
  if (errors && errors.length) {
    widgets.push({ textParagraph: { text: '<font color="#c62828"><b>⚠️ 必須項目が未入力です:</b> ' + errors.join(', ') + '</font>' } });
  }
  fields.forEach(function(f) {
    var pv = (prefill && prefill[f.name] != null) ? String(prefill[f.name]) : '';
    var label = f.name + (f.required ? ' *' : '');
    switch (f.type) {
      case 'text': widgets.push({ textInput: { label: label, name: 'field_' + f.order, type: 'SINGLE_LINE', value: pv } }); break;
      case 'textarea': widgets.push({ textInput: { label: label, name: 'field_' + f.order, type: 'MULTIPLE_LINE', value: pv } }); break;
      case 'dropdown': widgets.push({ selectionInput: { label: label, name: 'field_' + f.order, type: 'DROPDOWN', items: f.options.split(',').map(function(o, i) { var t = o.trim(); return { text: t, value: t, selected: pv ? (pv === t) : (i === 0) }; }) } }); break;
      case 'checkbox': var sels = pv ? pv.split(',').map(function(s){return s.trim();}) : []; widgets.push({ selectionInput: { label: label, name: 'field_' + f.order, type: 'CHECK_BOX', items: f.options.split(',').map(function(o) { var t = o.trim(); return { text: t, value: t, selected: sels.indexOf(t) !== -1 }; }) } }); break;
      case 'date':
        var dp = { label: label, name: 'field_' + f.order, type: 'DATE_ONLY' };
        if (pv) { var pd = new Date(pv); if (!isNaN(pd.getTime())) dp.valueMsEpoch = pd.getTime(); }
        widgets.push({ dateTimePicker: dp });
        break;
      case 'checkbox_qty':
        var qp = parseQtyOptions_(f.options);
        var qItems = qp.opts.split(',').map(function(s){return s.trim();}).filter(function(s){return s;});
        var pvMap = {};
        if (pv) { String(pv).split(',').forEach(function(seg){ var m = seg.match(/^\s*(.+?)\s*×\s*(\S+)\s*$/); if (m) pvMap[m[1]] = m[2]; }); }
        widgets.push({ textParagraph: { text: '<b>' + label + '</b>' } });
        qItems.forEach(function(itemName, oi) {
          var key = 'field_' + f.order + '_q_' + oi;
          var curVal = pvMap[itemName] || '';
          var itemLabel = itemName + '（数量）';
          if (qp.qtype === 'dropdown') {
            var qopts = qp.qopts.split(',').map(function(s){return s.trim();}).filter(function(s){return s;});
            if (qopts.indexOf('0') === -1) qopts.unshift('0');
            var items = [];
            qopts.forEach(function(q, qi){ items.push({ text: q, value: q, selected: curVal ? (curVal === q) : (qi === 0) }); });
            widgets.push({ selectionInput: { label: itemLabel, name: key, type: 'DROPDOWN', items: items } });
          } else {
            widgets.push({ textInput: { label: itemLabel, name: key, type: 'SINGLE_LINE', value: curVal, hintText: '半角数字（空欄=未選択）' } });
          }
        });
        break;
    }
  });
  var submitParams = [{ key: 'actionName', value: 'submitWorkflow' }, { key: 'workflowId', value: String(workflowId) }];
  if (chain) submitParams.push({ key: 'chain', value: String(chain) });
  if (inheritFrom) submitParams.push({ key: 'inheritThreadFrom', value: String(inheritFrom) });
  widgets.push({ buttonList: { buttons: [{ text: '送信', onClick: { action: { function: cardActionFunc_(), parameters: submitParams } } }] } });
  var icon = (wf.type === '申請・承認') ? '📋' : '📝';
  var _hdr = { title: icon + ' ' + wf.name };
  if (wf.description) _hdr.subtitle = wf.description;
  var card = { header: _hdr, sections: [{ widgets: widgets }] };
  var nav = (navMode === 'update') ? { updateCard: card } : { pushCard: card };
  return { action: { navigations: [nav] } };
}

function handleSubmitWorkflow_(event) {
  var params = (event.commonEventObject && event.commonEventObject.parameters) || {};
  var inputs = (event.commonEventObject && event.commonEventObject.formInputs) || {};
  var workflowId = parseInt(params.workflowId || '0', 10);
  var chain = parseInt(params.chain || '1', 10);
  var inheritFrom = String(params.inheritThreadFrom || '');
  var wf = getWorkflowById_(workflowId);
  if (!wf) return createMessage_('ワークフローが見つかりません。');
  var fields = getWorkflowFields_(workflowId);
  var senderName = '', senderEmail = '';
  if (event.chat && event.chat.user) { senderName = event.chat.user.displayName || ''; senderEmail = event.chat.user.email || ''; }
  // 送信先スペース未設定時のために、ボタンが押された実行スペースを特定する。
  var currentSpace = '';
  if (event.chat) {
    if (event.chat.buttonClickedPayload && event.chat.buttonClickedPayload.space) currentSpace = event.chat.buttonClickedPayload.space.name || '';
    else if (event.chat.appCommandPayload && event.chat.appCommandPayload.space) currentSpace = event.chat.appCommandPayload.space.name || '';
  }
  var values = {}, errors = [];
  fields.forEach(function(f) {
    var v;
    if (f.type === 'checkbox_qty') {
      var qp = parseQtyOptions_(f.options);
      var qItems = qp.opts.split(',').map(function(s){return s.trim();}).filter(function(s){return s;});
      var pairs = [];
      qItems.forEach(function(itemName, oi) {
        var qv = String(getInputValue_(inputs, 'field_' + f.order + '_q_' + oi) || '').trim();
        if (!qv) return;
        if (qp.qtype === 'number') {
          var n = parseInt(qv.replace(/[^0-9]/g, ''), 10);
          if (!isNaN(n) && n > 0) pairs.push(itemName + '×' + n);
        } else if (qv !== '0') {
          pairs.push(itemName + '×' + qv);
        }
      });
      v = pairs.join(', ');
    } else {
      v = getInputValue_(inputs, 'field_' + f.order);
    }
    values[f.name] = v;
    if (f.required && !String(v).trim()) errors.push(f.name);
  });
  if (errors.length > 0) return buildWorkflowDialogCard_(workflowId, values, errors, 'update', chain, inheritFrom);
  // 並行申請時の rowIndex 取り違え防止のため、appendRow→getLastRow→承認カード送信までを直列化する。
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return createMessage_('現在混み合っています。もう一度送信してください。');
  try {
    // データシートは fieldId → dataCol の対応で書き込む。列の追加/削除/並び替えに強い。
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dataSheet = ss.getSheetByName(wf.name);
    if (!dataSheet) {
      // 保存時に作られていなかった場合の遅延作成もタブ右端に配置
      dataSheet = ss.insertSheet(wf.name, ss.getSheets().length);
      var initHeaders = ['タイムスタンプ', '送信者', '送信者メール'];
      fields.forEach(function(f) { initHeaders.push(f.name); });
      if (wf.type === '申請・承認') initHeaders.push('ステータス');
      initHeaders.push('メッセージ名');
      dataSheet.appendRow(initHeaders);
    }
    // 手動で列が消された場合の防衛: 必須システム列(ステータス/メッセージ名など)を自動復旧
    try { ensureDataSheetSchema_(dataSheet, wf.type === '申請・承認'); } catch (e) { Logger.log('ensureDataSheetSchema_ skip: ' + e.message); }
    var lastCol = dataSheet.getLastColumn();
    var headerRow = dataSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var row = new Array(lastCol);
    for (var ci = 0; ci < lastCol; ci++) row[ci] = '';
    row[0] = new Date(); row[1] = senderName; row[2] = senderEmail;
    fields.forEach(function(f) {
      var col = f.dataCol && f.dataCol > 0 ? f.dataCol : -1;
      if (col > 0 && col <= lastCol) row[col - 1] = values[f.name];
    });
    var statusIdx = -1, msgIdx = -1;
    for (var hi = 0; hi < headerRow.length; hi++) { if (headerRow[hi] === 'ステータス') statusIdx = hi; if (headerRow[hi] === 'メッセージ名') msgIdx = hi; }
    if (wf.type === '申請・承認' && statusIdx >= 0) row[statusIdx] = '承認待ち';
    if (msgIdx >= 0) row[msgIdx] = '';
    dataSheet.appendRow(row);
    SpreadsheetApp.flush();
    var rowIndex = dataSheet.getLastRow();
    var headers = headerRow.slice();
    writeToExternalSheet_(wf, { headers: headers, values: row.slice(0, msgIdx >= 0 ? msgIdx : row.length) });
    var destSpace = wf.targetSpace || currentSpace;
    if (wf.type === '申請・承認') sendApprovalCard_(wf, fields, values, senderName, rowIndex, destSpace, chain, inheritFrom);
    else sendInfoCard_(wf, fields, values, senderName, destSpace, inheritFrom);
    return { action: { navigations: [{ endNavigation: { action: 'CLOSE_DIALOG' } }], notification: { text: '✅ 送信しました（' + wf.name + '）' } } };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function sendApprovalCard_(wf, fields, values, senderName, rowIndex, destSpace, chain, inheritFrom) {
  var widgets = [];
  widgets.push({ decoratedText: { topLabel: '送信者', text: senderName } });
  fields.forEach(function(f) { widgets.push({ decoratedText: { topLabel: f.name, text: String(values[f.name]) || '(なし)' } }); });
  widgets.push({ decoratedText: { topLabel: 'ステータス', text: '⏳ 承認待ち' } });
  var buttons = getWorkflowButtons_(wf.id);
  if (buttons.length === 0) buttons = [{ name: '承認', status: '承認', color: 'green' }, { name: '却下', status: '却下', color: 'red' }];
  var secToken = generateToken_(wf.id, rowIndex);
  var chainVal = chain || 1;
  var chatBtns = buttons.map(function(b, idx) {
    var color = COLOR_MAP[b.color] || COLOR_MAP.gray;
    if (b.type === 'launch' && b.linkedWorkflowId) {
      var lp = [
        { key: 'actionName', value: 'openWorkflow' },
        { key: 'workflowId', value: String(b.linkedWorkflowId) },
        { key: 'chain', value: String(chainVal + 1) }
      ];
      if (b.inheritThread) {
        lp.push({ key: 'parentWorkflowId', value: String(wf.id) });
        lp.push({ key: 'parentRow', value: String(rowIndex) });
      }
      return { text: b.name, color: color, onClick: { action: { function: cardActionFunc_(), parameters: lp, interaction: 'OPEN_DIALOG' } } };
    }
    return { text: b.name, color: color, onClick: { openLink: { url: FUNCTION_URL + '?action=btn&row=' + rowIndex + '&wf=' + wf.id + '&btn=' + idx + '&name=' + encodeURIComponent(senderName) + '&token=' + secToken + '&chain=' + chainVal } } };
  });
  widgets.push({ buttonList: { buttons: chatBtns } });
  // 承認者 User ID が保存されていれば message.text にメンションを付与して通知を飛ばす。
  // カード内 widget のメンションは通知が飛ばないため、message.text に載せる必要がある。
  var mentionText = '';
  if (wf.approverUserIds) {
    var ids = String(wf.approverUserIds).split(',').map(function(s) { return s.trim(); }).filter(function(s) { return s; });
    if (ids.length > 0) {
      var msg = String(wf.mentionText || '承認をお願いします');
      msg = msg.replace(/\{申請者\}/g, senderName).replace(/\{ワークフロー\}/g, wf.name);
      mentionText = ids.map(function(id) { return '<users/' + id + '>'; }).join(' ') + ' ' + msg;
    }
  }
  try {
    var token = getServiceAccountToken_();
    var payload = { cardsV2: [{ cardId: 'approval_' + wf.id + '_' + rowIndex, card: { header: { title: '🔔 新しい申請', subtitle: wf.name + ' - ' + senderName }, sections: [{ header: '申請内容', widgets: widgets }] } }] };
    if (mentionText) payload.text = mentionText;
    var sendTo = destSpace || wf.targetSpace;
    var urlSuffix = '';
    if (inheritFrom) {
      var parentSpace = String(inheritFrom).split('/messages/')[0];
      if (parentSpace) sendTo = parentSpace;
      var parentMsgId = String(inheritFrom).split('/').pop().split('.')[0];
      payload.thread = { name: parentSpace + '/threads/' + parentMsgId };
      urlSuffix = '?messageReplyOption=REPLY_MESSAGE_FALLBACK_TO_NEW_THREAD';
    }
    if (!sendTo) throw new Error('送信先スペースが特定できません');
    var resp = UrlFetchApp.fetch('https://chat.googleapis.com/v1/' + sendTo + '/messages' + urlSuffix, { method: 'post', contentType: 'application/json', headers: { 'Authorization': 'Bearer ' + token }, payload: JSON.stringify(payload), muteHttpExceptions: true });
    var result = JSON.parse(resp.getContentText());
    if (result.name) { SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(wf.name).getRange(rowIndex, SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(wf.name).getLastColumn()).setValue(result.name); }
  } catch (err) { Logger.log('カード送信エラー: ' + err.message); }
}

function sendInfoCard_(wf, fields, values, senderName, destSpace, inheritFrom) {
  try {
    var widgets = [];
    widgets.push({ decoratedText: { topLabel: '送信者', text: senderName } });
    fields.forEach(function(f) { widgets.push({ decoratedText: { topLabel: f.name, text: String(values[f.name]) || '(なし)' } }); });
    var sendTo = destSpace || wf.targetSpace;
    var payload = { cardsV2: [{ cardId: 'info_' + wf.id + '_' + Date.now(), card: { header: { title: '📝 ' + wf.name, subtitle: senderName }, sections: [{ widgets: widgets }] } }] };
    var urlSuffix = '';
    if (inheritFrom) {
      var parentSpace = String(inheritFrom).split('/messages/')[0];
      if (parentSpace) sendTo = parentSpace;
      var parentMsgId = String(inheritFrom).split('/').pop().split('.')[0];
      payload.thread = { name: parentSpace + '/threads/' + parentMsgId };
      urlSuffix = '?messageReplyOption=REPLY_MESSAGE_FALLBACK_TO_NEW_THREAD';
    }
    if (!sendTo) throw new Error('送信先スペースが特定できません');
    var token = getServiceAccountToken_();
    UrlFetchApp.fetch('https://chat.googleapis.com/v1/' + sendTo + '/messages' + urlSuffix, { method: 'post', contentType: 'application/json', headers: { 'Authorization': 'Bearer ' + token }, payload: JSON.stringify(payload), muteHttpExceptions: true });
  } catch (err) { Logger.log('定型文送信エラー: ' + err.message); }
}

// 連動ボタン: 親メッセージのスレッドに次ワークフローの起動カードを投稿する。
// inheritThread=true のとき次のワークフロー送信先も親スレッドに継承させる (inheritThreadFrom を引き継ぐ)。
// chain は投稿時点の階層カウント(次ワークフロー送信で chain+1 が記録される)。
function postLauncherToThread_(linkedWf, parentMessageName, inheritThread, chain, applicantName) {
  try {
    if (!parentMessageName) throw new Error('親メッセージが特定できません');
    var parentSpace = String(parentMessageName).split('/messages/')[0];
    var parentMsgId = String(parentMessageName).split('/').pop().split('.')[0];
    var threadName = parentSpace + '/threads/' + parentMsgId;
    var params = [{ key: 'actionName', value: 'openWorkflow' }, { key: 'workflowId', value: String(linkedWf.id) }, { key: 'chain', value: String(chain) }];
    if (inheritThread) params.push({ key: 'inheritThreadFrom', value: parentMessageName });
    var card = {
      header: { title: '🚀 次のワークフロー', subtitle: linkedWf.name },
      sections: [{ widgets: [
        { textParagraph: { text: (applicantName || '前のワークフロー') + ' から連動起動されました' } },
        { buttonList: { buttons: [{ text: '起動', onClick: { action: { function: cardActionFunc_(), parameters: params, interaction: 'OPEN_DIALOG' } } }] } }
      ] }]
    };
    var token = getServiceAccountToken_();
    UrlFetchApp.fetch('https://chat.googleapis.com/v1/' + parentSpace + '/messages?messageReplyOption=REPLY_MESSAGE_FALLBACK_TO_NEW_THREAD', { method: 'post', contentType: 'application/json', headers: { 'Authorization': 'Bearer ' + token }, payload: JSON.stringify({ cardsV2: [{ cardId: 'launcher_' + linkedWf.id + '_' + Date.now(), card: card }], thread: { name: threadName } }), muteHttpExceptions: true });
  } catch (err) { Logger.log('postLauncherToThread_ エラー: ' + err.message); try { notifyAdminOnError_(err, 'postLauncherToThread_'); } catch (e) {} }
}

function handleApproval_(e) {
  var lock = LockService.getScriptLock();
  try {
    var rowIndex = parseInt(e.parameter.row || '0', 10);
    var workflowId = parseInt(e.parameter.wf || '0', 10);
    var btnIndex = parseInt(e.parameter.btn || '0', 10);
    var applicantName = decodeURIComponent(e.parameter.name || '不明');
    var wf = getWorkflowById_(workflowId);
    if (!wf || rowIndex < 2) return htmlPage_('🚫', '無効なリクエスト', '', '#666');
    var tokenCheck = validateToken_(e.parameter.token || '', workflowId, rowIndex);
    if (!tokenCheck.ok) return htmlPage_('🚫', '無効なリクエスト', 'このリンクは使用できません(' + tokenCheck.reason + ')', '#666');
    var buttons = getWorkflowButtons_(workflowId);
    if (buttons.length === 0) buttons = [{ name: '承認', status: '承認', color: 'green', threadReply: false, replyMsg: '' }, { name: '却下', status: '却下', color: 'red', threadReply: false, replyMsg: '' }];
    var btn = buttons[btnIndex];
    if (!btn) return htmlPage_('🚫', '無効なボタン', '', '#666');
    var incomingChain = parseInt(e.parameter.chain || '1', 10);
    // 連動起動タイプ: ステータス更新は行わず、親スレッドに次ワークフロー起動カードを投稿する
    if (btn.type === 'launch') {
      if (incomingChain >= 3) return htmlPage_('⚠️', '連鎖上限', '3階層までの連鎖に制限されています', '#666');
      if (!btn.linkedWorkflowId) return htmlPage_('⚠️', 'エラー', '連動先ワークフローが未設定です', '#666');
      var linkedWf = getWorkflowById_(btn.linkedWorkflowId);
      if (!linkedWf) return htmlPage_('⚠️', 'エラー', '起動対象ワークフローが見つかりません', '#666');
      var ssL = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheetL = ssL.getSheetByName(wf.name);
      if (!sheetL) return htmlPage_('⚠️', 'エラー', 'データシートが見つかりません', '#666');
      var headersL = sheetL.getRange(1, 1, 1, sheetL.getLastColumn()).getValues()[0];
      var msgColL = -1;
      for (var ci = 0; ci < headersL.length; ci++) if (headersL[ci] === 'メッセージ名') { msgColL = ci + 1; break; }
      var parentMsg = msgColL > 0 ? sheetL.getRange(rowIndex, msgColL).getValue() : '';
      postLauncherToThread_(linkedWf, parentMsg, btn.inheritThread, incomingChain + 1, applicantName);
      writeAuditLog_('approval.launch', 'wid=' + workflowId + ' row=' + rowIndex, 'linked=' + btn.linkedWorkflowId);
      return htmlPage_('🚀', '起動しました', '次のワークフロー「' + linkedWf.name + '」をスレッドに投稿しました', '#1a73e8');
    }
    var newStatus = btn.status;
    var emoji = { green: '✅', blue: '🔵', yellow: '🟡', orange: '🟠', red: '❌', gray: '⚫' }[btn.color] || '⚪';
    var htmlColor = { green: '#1a7a1a', blue: '#1a73e8', yellow: '#b06000', orange: '#e65100', red: '#c62828', gray: '#666' }[btn.color] || '#333';
    if (!lock.tryLock(15000)) return htmlPage_('⏳', '混雑中', '他の承認処理が進行中です。10秒ほど待ってからもう一度お試しください', '#666');
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(wf.name);
    if (!sheet) return htmlPage_('⚠️', 'エラー', 'データシートが見つかりません', '#666');
    // 手動で列が消された場合の防衛: ステータス/メッセージ名列を自動復旧
    try { ensureDataSheetSchema_(sheet, wf.type === '申請・承認'); } catch (e) { Logger.log('ensureDataSheetSchema_ skip: ' + e.message); }
    var fields = getWorkflowFields_(workflowId);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var statusCol = -1, msgCol = -1;
    for (var c = 0; c < headers.length; c++) { if (headers[c] === 'ステータス') statusCol = c + 1; if (headers[c] === 'メッセージ名') msgCol = c + 1; }
    if (statusCol === -1) return htmlPage_('⚠️', 'エラー', 'ステータス列が見つかりません', '#666');
    var currentStatus = sheet.getRange(rowIndex, statusCol).getValue();
    if (currentStatus !== '承認待ち') return htmlPage_('⚠️', '処理済み', 'この申請は既に「' + currentStatus + '」されています', '#666');
    var currentUser = normalizeEmail_(Session.getActiveUser().getEmail());
    if (wf.approvers) {
      var approverList = wf.approvers.split(',').map(function(a) { return normalizeEmail_(a); });
      if (!currentUser || approverList.indexOf(currentUser) === -1) return htmlPage_('🚫', '権限がありません', '', '#666');
    } else {
      // 承認者未設定の過去ワークフローは管理者のみに制限(新規保存は承認者必須のため、ここに来るのは旧データのみ)
      if (!currentUser || !isAdmin_(currentUser)) return htmlPage_('🚫', '権限がありません', '承認者が設定されていません。管理画面でワークフローに承認者を設定してください。', '#666');
    }
    sheet.getRange(rowIndex, statusCol).setValue(newStatus);
    SpreadsheetApp.flush();
    var writeBack = sheet.getRange(rowIndex, statusCol).getValue();
    Logger.log('approval setValue row=' + rowIndex + ' expected=' + newStatus + ' actual=' + writeBack);
    updateExternalSheet_(wf, wf.name, rowIndex, statusCol, newStatus);
    var messageName = msgCol > 0 ? sheet.getRange(rowIndex, msgCol).getValue() : '';
    if (messageName) updateApprovalCard_(wf, fields, sheet, rowIndex, messageName, applicantName, newStatus, emoji, headers);
    if (btn.threadReply && btn.replyMsg && messageName) {
      // 送信先が空(=実行スペース返信)でも、messageName からスペースを抽出して返信先を特定する。
      var replySpace = wf.targetSpace || (String(messageName).split('/messages/')[0] || '');
      sendThreadReply_(replySpace, messageName, btn.replyMsg, applicantName, newStatus);
    }
    writeAuditLog_('approval.' + (newStatus === '承認' ? 'approve' : 'decide'), 'wid=' + workflowId + ' row=' + rowIndex, newStatus + '/' + applicantName + (tokenCheck.legacy ? ' [legacy-token]' : ''));
    try { incrementDailyCounter_('approval', 1); } catch (e4) {}
    return htmlPage_(emoji, newStatus, applicantName + ' の申請を「' + newStatus + '」にしました', htmlColor);
  } catch (err) {
    Logger.log('handleApproval_ エラー: ' + err.message);
    try { notifyAdminOnError_(err, 'handleApproval_'); } catch (e3) {}
    return htmlPage_('⚠️', 'エラー', '処理中にエラーが発生しました。管理者に連絡してください: ' + CONTACT_EMAIL, '#666');
  }
  finally { try { lock.releaseLock(); } catch (e2) {} }
}

function updateApprovalCard_(wf, fields, sheet, rowIndex, messageName, applicantName, newStatus, emoji, headers) {
  var widgets = [];
  widgets.push({ decoratedText: { topLabel: '送信者', text: applicantName } });
  fields.forEach(function(f) {
    var col = -1;
    for (var c = 0; c < headers.length; c++) { if (headers[c] === f.name) { col = c + 1; break; } }
    widgets.push({ decoratedText: { topLabel: f.name, text: String(col > 0 ? sheet.getRange(rowIndex, col).getValue() : '') || '(なし)' } });
  });
  widgets.push({ decoratedText: { topLabel: 'ステータス', text: emoji + ' ' + newStatus } });
  try {
    var token = getServiceAccountToken_();
    var upd = UrlFetchApp.fetch('https://chat.googleapis.com/v1/' + messageName + '?updateMask=cardsV2', { method: 'put', contentType: 'application/json', headers: { 'Authorization': 'Bearer ' + token }, payload: JSON.stringify({ cardsV2: [{ cardId: 'approval_' + wf.id + '_' + rowIndex, card: { header: { title: emoji + ' ' + newStatus, subtitle: wf.name + ' - ' + applicantName }, sections: [{ widgets: widgets }] } }] }), muteHttpExceptions: true });
    if (upd.getResponseCode() >= 400) {
      var errMsg = 'approval card update HTTP ' + upd.getResponseCode() + ' ' + upd.getContentText().substring(0, 300);
      Logger.log(errMsg);
      try { notifyAdminOnError_(new Error(errMsg), 'updateApprovalCard_'); } catch (e2) {}
    }
  } catch (err) { Logger.log('カード更新エラー: ' + err.message); try { notifyAdminOnError_(err, 'updateApprovalCard_'); } catch (e3) {} }
}

function sendThreadReply_(spaceId, messageName, msgTemplate, applicantName, newStatus) {
  try {
    var text = msgTemplate.replace(/{申請者}/g, applicantName).replace(/{ステータス}/g, newStatus);
    var msgId = messageName.split('/').pop().split('.')[0];
    var threadName = messageName.split('/messages/')[0] + '/threads/' + msgId;
    var token = getServiceAccountToken_();
    UrlFetchApp.fetch('https://chat.googleapis.com/v1/' + spaceId + '/messages?messageReplyOption=REPLY_MESSAGE_FALLBACK_TO_NEW_THREAD', { method: 'post', contentType: 'application/json', headers: { 'Authorization': 'Bearer ' + token }, payload: JSON.stringify({ text: text, thread: { name: threadName } }), muteHttpExceptions: true });
  } catch (err) { Logger.log('スレッド返信エラー: ' + err.message); }
}

function htmlPage_(emoji, title, msg, color) {
  return HtmlService.createHtmlOutput('<html><body style="font-family:sans-serif;text-align:center;padding:60px 20px;background:#f5f5f5"><div style="background:#fff;border-radius:12px;padding:40px;max-width:360px;margin:0 auto"><p style="font-size:48px">' + emoji + '</p><p style="font-size:20px;font-weight:bold;margin:16px 0;color:' + color + '">' + title + '</p>' + (msg ? '<p style="color:#666">' + msg + '</p>' : '') + '<button onclick="window.close();open(location,\'_self\').close();" style="margin-top:20px;padding:12px 40px;border:none;border-radius:8px;background:' + color + ';color:#fff;cursor:pointer;font-size:15px">閉じる</button><p style="color:#aaa;font-size:11px;margin-top:8px">閉じない場合はタブを手動で閉じてください</p></div></body></html>');
}