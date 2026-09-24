// ============================================================
// Main.gs - 定数 + doPost + doGet + ルーティング
// ============================================================

var VERSION = '1.2.7';

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

var SPREADSHEET_ID = getConfig_('SPREADSHEET_ID', '1iHLqkdJyCZUP6eivZiMovZvt6kbrWnzTDsTV2O5GUS8');
var FUNCTION_URL   = getConfig_('FUNCTION_URL',   'https://script.google.com/macros/s/AKfycbwwTU7juIOgPDJiqRNNfksYsKls0ypfHjkykSWvI571cwjmqeIgtyYCHUZlscoYSyVu/exec');
var CONTACT_EMAIL  = getConfig_('CONTACT_EMAIL',  'konosuke.okazaki@cmind-co.jp');
var ENV_NAME       = getConfig_('ENV_NAME',       'prod');

var COLOR_MAP = {
  green:  { red: 0.2, green: 0.7, blue: 0.2, alpha: 1 },
  blue:   { red: 0.1, green: 0.46, blue: 0.82, alpha: 1 },
  yellow: { red: 0.9, green: 0.7, blue: 0.1, alpha: 1 },
  orange: { red: 0.9, green: 0.5, blue: 0.1, alpha: 1 },
  red:    { red: 0.8, green: 0.2, blue: 0.2, alpha: 1 },
  gray:   { red: 0.5, green: 0.5, blue: 0.5, alpha: 1 }
};

function doPost(e) {
  var event = JSON.parse(e.postData.contents);
  Logger.log('EVENT: ' + JSON.stringify(event).substring(0, 1000));
  var result = {};
  try {
    if (event.chat && event.chat.appCommandPayload) {
      var cmdId = event.chat.appCommandPayload.appCommandMetadata.appCommandId;
      if (cmdId === 1 || cmdId === '1') result = handleApplyCommand_(event);
      else result = createMessage_('不明なコマンドです。');
    } else if (event.chat && event.chat.buttonClickedPayload) {
      var actionName = '';
      if (event.commonEventObject && event.commonEventObject.parameters) actionName = event.commonEventObject.parameters.actionName || '';
      switch (actionName) {
        case 'openWorkflow': result = openWorkflowDialog_(event); break;
        case 'submitWorkflow': result = handleSubmitWorkflow_(event); break;
        default: result = createMessage_('不明なアクションです: ' + actionName);
      }
    } else {
      result = createMessage_('ワークフローを実行するにはボタンを押すか、/apply コマンドを使ってください。');
    }
  } catch (err) {
    Logger.log('ERROR: ' + err.message + '\n' + err.stack);
    try { notifyAdminOnError_(err, 'doPost'); } catch (e) {}
    result = createMessage_('エラーが発生しました。管理者に連絡してください: ' + CONTACT_EMAIL);
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

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
  var body = { ok: sheetOk, version: VERSION, env: ENV_NAME, sheetOk: sheetOk, workflows: workflowCount, admins: adminCount, checkedAt: new Date().toISOString() };
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
