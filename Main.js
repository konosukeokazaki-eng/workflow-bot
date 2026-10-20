// ============================================================
// Main.gs - 定数 + doPost + doGet + ルーティング
// ============================================================

var VERSION = '1.5.19';

var SHEET_SETTINGS   = '_設定';
var SHEET_FIELDS     = '_項目設定';
var SHEET_BUTTONS    = '_ボタン設定';
var SHEET_ADMINS     = '_管理者';
var SHEET_AUDIT      = '_監査ログ';
var SHEET_TOKENS     = '_トークン履歴';

var TOKEN_TTL_DAYS = 30;

// スクリプトプロパティから取得。未設定時はフォールバック値を使用。
// dev/stg/prod 切替はスクリプトプロパティで行うこと。
function getConfig_(key, fallback) {
  try {
    var v = PropertiesService.getScriptProperties().getProperty(key);
    return (v && v.length > 0) ? v : fallback;
  } catch (e) { return fallback; }
}

// 機密値はスクリプトプロパティで設定する。リポジトリにフォールバックを残さない。
// 設定キー: SPREADSHEET_ID / FUNCTION_URL / CONTACT_EMAIL / ENV_NAME
var SPREADSHEET_ID = getConfig_('SPREADSHEET_ID', '');
var FUNCTION_URL   = getConfig_('FUNCTION_URL',   '');
var CONTACT_EMAIL  = getConfig_('CONTACT_EMAIL',  '');
var ENV_NAME       = getConfig_('ENV_NAME',       'prod');
// Chat API 接続モード: 'http' or 'apps_script'
// カード内ボタンの action.function にセットする値を切り替える。
// - 'http'       : FUNCTION_URL (本番の Web App URL) を URL として渡す
// - 'apps_script': 'onCardClick' という Apps Script 関数名を渡す
var CHAT_CONNECTION_MODE = getConfig_('CHAT_CONNECTION_MODE', 'http');
function cardActionFunc_() { return CHAT_CONNECTION_MODE === 'apps_script' ? 'onCardClick' : FUNCTION_URL; }

var COLOR_MAP = {
  green:  { red: 0.2, green: 0.7, blue: 0.2, alpha: 1 },
  blue:   { red: 0.1, green: 0.46, blue: 0.82, alpha: 1 },
  yellow: { red: 0.9, green: 0.7, blue: 0.1, alpha: 1 },
  orange: { red: 0.9, green: 0.5, blue: 0.1, alpha: 1 },
  red:    { red: 0.8, green: 0.2, blue: 0.2, alpha: 1 },
  gray:   { red: 0.5, green: 0.5, blue: 0.5, alpha: 1 }
};

// Chat API 接続方式(HTTP or Apps Script)の両方に対応できるよう、
// doPost で受けた Chat イベントを分岐する内部関数と、Apps Script ネイティブトリガー関数を併設する。
function dispatchChatEvent_(event) {
  try {
    if (event.chat && event.chat.appCommandPayload) {
      var cmdId = event.chat.appCommandPayload.appCommandMetadata && event.chat.appCommandPayload.appCommandMetadata.appCommandId;
      if (cmdId === 1 || cmdId === '1') return handleApplyCommand_(event);
      return createMessage_('不明なコマンドです。');
    } else if (event.chat && event.chat.buttonClickedPayload) {
      var actionName = '';
      if (event.commonEventObject && event.commonEventObject.parameters) actionName = event.commonEventObject.parameters.actionName || '';
      switch (actionName) {
        case 'openWorkflow': return openWorkflowDialog_(event);
        case 'submitWorkflow': return handleSubmitWorkflow_(event);
        default: return createMessage_('不明なアクションです: ' + actionName);
      }
    } else if (event.chat && event.chat.messagePayload) {
      return createMessage_('ワークフローを実行するにはボタンを押すか、/apply コマンドを使ってください。');
    } else if (event.chat && event.chat.addedToSpacePayload) {
      return createMessage_('Workflow Bot を追加いただきありがとうございます。/apply コマンドでワークフローを実行できます。');
    }
    return createMessage_('ワークフローを実行するにはボタンを押すか、/apply コマンドを使ってください。');
  } catch (err) {
    Logger.log('ERROR: ' + err.message + '\n' + err.stack);
    try { notifyAdminOnError_(err, 'dispatchChatEvent_'); } catch (e) {}
    return createMessage_('エラーが発生しました。管理者に連絡してください: ' + CONTACT_EMAIL);
  }
}

function doPost(e) {
  var event = JSON.parse(e.postData.contents);
  Logger.log('EVENT: ' + JSON.stringify(event).substring(0, 5000));
  var result = dispatchChatEvent_(event);
  Logger.log('RESP: ' + JSON.stringify(result).substring(0, 5000));
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// Apps Script ネイティブ Chat トリガー (Chat API 接続が「Apps Script」の場合に呼ばれる)
// HTTP モードの応答形式(action.navigations / hostAppDataAction) を Apps Script 向け
// (actionResponse.dialogAction / 素の message オブジェクト) に変換して返す。
function toAppsScriptResponse_(response) {
  if (!response) return {};
  if (response.action && response.action.navigations && response.action.navigations.length) {
    var nav = response.action.navigations[0];
    if (nav.pushCard || nav.updateCard) {
      var card = nav.pushCard || nav.updateCard;
      return { actionResponse: { type: 'DIALOG', dialogAction: { dialog: { body: card } } } };
    }
  }
  if (response.hostAppDataAction && response.hostAppDataAction.chatDataAction) {
    var cda = response.hostAppDataAction.chatDataAction;
    if (cda.createMessageAction && cda.createMessageAction.message) return cda.createMessageAction.message;
  }
  return response;
}
function logEvent_(name, event) { try { Logger.log(name + ' EVENT: ' + JSON.stringify(event).substring(0, 2000)); } catch (e) {} }
function onAppCommand(event) { logEvent_('onAppCommand', event); var r = toAppsScriptResponse_(dispatchChatEvent_(event)); Logger.log('onAppCommand RESP: ' + JSON.stringify(r).substring(0,1500)); return r; }
function onCardClick(event)  { logEvent_('onCardClick', event);  return toAppsScriptResponse_(dispatchChatEvent_(event)); }
function onMessage(event)    { logEvent_('onMessage', event);    return toAppsScriptResponse_(dispatchChatEvent_(event)); }
function onAddedToSpace(event) { logEvent_('onAddedToSpace', event); return toAppsScriptResponse_(dispatchChatEvent_(event)); }
function onRemovedFromSpace(event) { Logger.log('onRemovedFromSpace'); return {}; }

function doGet(e) {
  if (e.parameter.ping === '1') return healthCheck_();
  if (e.parameter.action) return handleApproval_(e);
  var user = Session.getActiveUser().getEmail();
  if (!isAdmin_(user)) {
    return HtmlService.createHtmlOutput('<html><body style="font-family:sans-serif;text-align:center;padding:60px 20px;background:#f5f5f5"><div style="background:#fff;border-radius:12px;padding:40px;max-width:400px;margin:0 auto"><p style="font-size:48px">🔒</p><p style="font-size:18px;font-weight:bold;margin:16px 0">アクセス権限がありません</p><p style="color:#666">' + user + '</p><p style="color:#999;font-size:12px;margin-top:24px">問い合わせ: ' + CONTACT_EMAIL + '</p></div></body></html>').setTitle('アクセス拒否');
  }
  return getAdminPage_();
}

function healthCheck_() {
  var sheetOk = false, adminCount = 0, workflowCount = 0;
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var st = ss.getSheetByName(SHEET_SETTINGS);
    var ad = ss.getSheetByName(SHEET_ADMINS);
    sheetOk = !!(st && ad);
    if (st && st.getLastRow() >= 2) workflowCount = st.getLastRow() - 1;
    if (ad && ad.getLastRow() >= 2) adminCount = ad.getLastRow() - 1;
  } catch (e) { sheetOk = false; }
  // サービスアカウントトークンが取得できるか(鍵ローテ/失効の早期検知)
  var saTokenOk = false, saTokenError = '';
  try { getServiceAccountToken_(); saTokenOk = true; }
  catch (e) { saTokenError = e.message; }
  var overallOk = sheetOk && saTokenOk;
  var body = { ok: overallOk, version: VERSION, env: ENV_NAME, sheetOk: sheetOk, serviceAccountOk: saTokenOk, serviceAccountError: saTokenError, workflows: workflowCount, admins: adminCount, checkedAt: new Date().toISOString() };
  return ContentService.createTextOutput(JSON.stringify(body)).setMimeType(ContentService.MimeType.JSON);
}

function createMessage_(t) { return { hostAppDataAction: { chatDataAction: { createMessageAction: { message: { text: t } } } } }; }

function getInputValue_(inputs, key) {
  if (inputs[key]) {
    if (inputs[key].stringInputs && inputs[key].stringInputs.value) return inputs[key].stringInputs.value.join(', ');
    if (inputs[key].dateInput) return inputs[key].dateInput.msSinceEpoch ? new Date(parseInt(inputs[key].dateInput.msSinceEpoch)).toLocaleDateString('ja-JP') : '';
  }
  return '';
}
