// ============================================================
// Chat.gs - Google Chat 連携（コマンド、カード送信、承認処理）
// ============================================================

function handleApplyCommand_(event) {
  var workflows = getActiveWorkflows_();
  var currentSpace = '';
  if (event.chat && event.chat.appCommandPayload && event.chat.appCommandPayload.space) currentSpace = event.chat.appCommandPayload.space.name || '';
  if (currentSpace) workflows = workflows.filter(function(wf) { return wf.execSpace === currentSpace; });
  if (workflows.length === 0) return { action: { navigations: [{ pushCard: { header: { title: 'ワークフロー' }, sections: [{ widgets: [{ textParagraph: { text: 'このスペースに紐づいたワークフローはありません。' } }] }] } }] } };
  if (workflows.length === 1) return buildWorkflowDialog_(workflows[0].id);
  var items = workflows.map(function(wf) { return { text: wf.name, value: String(wf.id), selected: false }; });
  items[0].selected = true;
  return { action: { navigations: [{ pushCard: { header: { title: '📋 ワークフロー選択', subtitle: '実行するワークフローを選んでください' }, sections: [{ widgets: [{ selectionInput: { label: 'ワークフロー', name: 'workflowId', type: 'DROPDOWN', items: items } }, { buttonList: { buttons: [{ text: '次へ', onClick: { action: { function: FUNCTION_URL, parameters: [{ key: 'actionName', value: 'openWorkflow' }] } } }] } }] }] } }] } };
}

function openWorkflowDialog_(event) {
  var inputs = (event.commonEventObject && event.commonEventObject.formInputs) || {};
  var params = (event.commonEventObject && event.commonEventObject.parameters) || {};
  var wid = '';
  if (inputs.workflowId && inputs.workflowId.stringInputs) wid = inputs.workflowId.stringInputs.value[0];
  if (!wid && params.workflowId) wid = params.workflowId;
  if (!wid) return createMessage_('ワークフローが選択されていません。');
  return buildWorkflowDialog_(parseInt(wid, 10));
}

function buildWorkflowDialog_(workflowId) {
  return buildWorkflowDialogCard_(workflowId, null, null, 'push');
}

// 共通のカード生成。prefill(name -> string値) と errors(必須未入力の項目名配列) を渡すと
// 入力値を保持したまま再描画し、エラー行を先頭に挿入する。
// navMode: 'push'(初回) or 'update'(再描画)
function buildWorkflowDialogCard_(workflowId, prefill, errors, navMode) {
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
    }
  });
  widgets.push({ buttonList: { buttons: [{ text: '送信', onClick: { action: { function: FUNCTION_URL, parameters: [{ key: 'actionName', value: 'submitWorkflow' }, { key: 'workflowId', value: String(workflowId) }] } } }] } });
  var icon = (wf.type === '申請・承認') ? '📋' : '📝';
  var card = { header: { title: icon + ' ' + wf.name }, sections: [{ widgets: widgets }] };
  var nav = (navMode === 'update') ? { updateCard: card } : { pushCard: card };
  return { action: { navigations: [nav] } };
}

function handleSubmitWorkflow_(event) {
  var params = (event.commonEventObject && event.commonEventObject.parameters) || {};
  var inputs = (event.commonEventObject && event.commonEventObject.formInputs) || {};
  var workflowId = parseInt(params.workflowId || '0', 10);
  var wf = getWorkflowById_(workflowId);
  if (!wf) return createMessage_('ワークフローが見つかりません。');
  var fields = getWorkflowFields_(workflowId);
  var senderName = '', senderEmail = '';
  if (event.chat && event.chat.user) { senderName = event.chat.user.displayName || ''; senderEmail = event.chat.user.email || ''; }
  var values = {}, errors = [];
  fields.forEach(function(f) { var v = getInputValue_(inputs, 'field_' + f.order); values[f.name] = v; if (f.required && !String(v).trim()) errors.push(f.name); });
  if (errors.length > 0) return buildWorkflowDialogCard_(workflowId, values, errors, 'update');
  // 並行申請時の rowIndex 取り違え防止のため、appendRow→getLastRow→承認カード送信までを直列化する。
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return createMessage_('現在混み合っています。もう一度送信してください。');
  try {
    // データシートは fieldId → dataCol の対応で書き込む。列の追加/削除/並び替えに強い。
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var dataSheet = ss.getSheetByName(wf.name);
    if (!dataSheet) {
      dataSheet = ss.insertSheet(wf.name);
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
    if (wf.type === '申請・承認') sendApprovalCard_(wf, fields, values, senderName, rowIndex);
    else sendInfoCard_(wf, fields, values, senderName);
    return { action: { navigations: [{ endNavigation: { action: 'CLOSE_DIALOG' } }], notification: { text: '✅ 送信しました（' + wf.name + '）' } } };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function sendApprovalCard_(wf, fields, values, senderName, rowIndex) {
  var widgets = [];
  widgets.push({ decoratedText: { topLabel: '送信者', text: senderName } });
  fields.forEach(function(f) { widgets.push({ decoratedText: { topLabel: f.name, text: String(values[f.name]) || '(なし)' } }); });
  widgets.push({ decoratedText: { topLabel: 'ステータス', text: '⏳ 承認待ち' } });
  var buttons = getWorkflowButtons_(wf.id);
  if (buttons.length === 0) buttons = [{ name: '承認', status: '承認', color: 'green' }, { name: '却下', status: '却下', color: 'red' }];
  var secToken = generateToken_(wf.id, rowIndex);
  var chatBtns = buttons.map(function(b, idx) {
    return { text: b.name, color: COLOR_MAP[b.color] || COLOR_MAP.gray, onClick: { openLink: { url: FUNCTION_URL + '?action=btn&row=' + rowIndex + '&wf=' + wf.id + '&btn=' + idx + '&name=' + encodeURIComponent(senderName) + '&token=' + secToken } } };
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
    var resp = UrlFetchApp.fetch('https://chat.googleapis.com/v1/' + wf.targetSpace + '/messages', { method: 'post', contentType: 'application/json', headers: { 'Authorization': 'Bearer ' + token }, payload: JSON.stringify(payload), muteHttpExceptions: true });
    var result = JSON.parse(resp.getContentText());
    if (result.name) { SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(wf.name).getRange(rowIndex, SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(wf.name).getLastColumn()).setValue(result.name); }
  } catch (err) { Logger.log('カード送信エラー: ' + err.message); }
}

function sendInfoCard_(wf, fields, values, senderName) {
  try {
    var widgets = [];
    widgets.push({ decoratedText: { topLabel: '送信者', text: senderName } });
    fields.forEach(function(f) { widgets.push({ decoratedText: { topLabel: f.name, text: String(values[f.name]) || '(なし)' } }); });
    var token = getServiceAccountToken_();
    UrlFetchApp.fetch('https://chat.googleapis.com/v1/' + wf.targetSpace + '/messages', { method: 'post', contentType: 'application/json', headers: { 'Authorization': 'Bearer ' + token }, payload: JSON.stringify({ cardsV2: [{ cardId: 'info_' + wf.id + '_' + Date.now(), card: { header: { title: '📝 ' + wf.name, subtitle: senderName }, sections: [{ widgets: widgets }] } }] }), muteHttpExceptions: true });
  } catch (err) { Logger.log('定型文送信エラー: ' + err.message); }
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
    if (wf.approvers) {
      var currentUser = normalizeEmail_(Session.getActiveUser().getEmail());
      var approverList = wf.approvers.split(',').map(function(a) { return normalizeEmail_(a); });
      if (!currentUser || approverList.indexOf(currentUser) === -1) return htmlPage_('🚫', '権限がありません', '', '#666');
    }
    sheet.getRange(rowIndex, statusCol).setValue(newStatus);
    updateExternalSheet_(wf, wf.name, rowIndex, statusCol, newStatus);
    var messageName = msgCol > 0 ? sheet.getRange(rowIndex, msgCol).getValue() : '';
    if (messageName) updateApprovalCard_(wf, fields, sheet, rowIndex, messageName, applicantName, newStatus, emoji, headers);
    if (btn.threadReply && btn.replyMsg && messageName) sendThreadReply_(wf.targetSpace, messageName, btn.replyMsg, applicantName, newStatus);
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
    UrlFetchApp.fetch('https://chat.googleapis.com/v1/' + messageName + '?updateMask=cardsV2', { method: 'put', contentType: 'application/json', headers: { 'Authorization': 'Bearer ' + token }, payload: JSON.stringify({ cardsV2: [{ cardId: 'approval_' + wf.id + '_' + rowIndex, card: { header: { title: emoji + ' ' + newStatus, subtitle: wf.name + ' - ' + applicantName }, sections: [{ widgets: widgets }] } }] }), muteHttpExceptions: true });
  } catch (err) { Logger.log('カード更新エラー: ' + err.message); }
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