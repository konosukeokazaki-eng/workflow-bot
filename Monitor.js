// ============================================================
// Monitor.gs - ヘルスチェック定期実行と管理者アラート
// ============================================================
//
// 前提:
//   - スクリプトプロパティ ADMIN_SPACE_ID に管理者通知用 Chat スペース名を設定する
//     (未設定時はメール通知にフォールバック)
//   - トリガー "monitorHealth_" を 5 分毎に設定する
//   - トリガー "dailyBackup_" を毎日 00:30 に設定する
//   - トリガー "monthlyArchive_" を毎月 1 日 01:00 に設定する
//   これらの初期セットアップは installMonitorTriggers() で一括登録可能。
// ============================================================

function sendAdminChat_(title, text) {
  var spaceId = getConfig_('ADMIN_SPACE_ID', '');
  if (spaceId) {
    try {
      var token = getServiceAccountToken_();
      var payload = {
        cardsV2: [{
          cardId: 'monitor_' + Date.now(),
          card: {
            header: { title: title, subtitle: ENV_NAME + ' / v' + VERSION },
            sections: [{ widgets: [{ textParagraph: { text: text } }] }]
          }
        }]
      };
      UrlFetchApp.fetch('https://chat.googleapis.com/v1/' + spaceId + '/messages', {
        method: 'post',
        contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + token },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      return true;
    } catch (e) { Logger.log('sendAdminChat_ 失敗: ' + e.message); }
  }
  try {
    MailApp.sendEmail(CONTACT_EMAIL, '[workflow-bot ' + ENV_NAME + '] ' + title, text);
    return true;
  } catch (e) { Logger.log('sendAdminChat_ メールも失敗: ' + e.message); return false; }
}

// 例外を管理者に通知。同一エラーの連発を 10 分間抑制する。
function notifyAdminOnError_(err, context) {
  try {
    var msg = err && err.message ? err.message : String(err);
    var stack = err && err.stack ? err.stack : '';
    var cache = CacheService.getScriptCache();
    var key = 'err_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, context + '|' + msg)).substring(0, 24);
    if (cache.get(key)) return;
    cache.put(key, '1', 600);
    var body = '発生箇所: ' + context + '\n' +
               'エラー: ' + msg + '\n\n' +
               (stack ? 'スタック:\n' + stack.substring(0, 800) : '');
    sendAdminChat_('⚠️ 実行時エラー', body);
    try { writeAuditLog_('error.runtime', context, msg); } catch (e) {}
  } catch (e) { Logger.log('notifyAdminOnError_ 失敗: ' + e.message); }
}

// 状態遷移のみ通知する死活監視。5 分毎のトリガーで呼ぶ。
function monitorHealth_() {
  var props = PropertiesService.getScriptProperties();
  var prev = props.getProperty('LAST_HEALTH_STATE') || 'unknown';
  var status = 'ok', reasons = [];
  var workflowCount = 0, adminCount = 0;
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var st = ss.getSheetByName(SHEET_SETTINGS);
    var ad = ss.getSheetByName(SHEET_ADMINS);
    if (!st) { status = 'ng'; reasons.push('_設定シートが見つかりません'); }
    if (!ad) { status = 'ng'; reasons.push('_管理者シートが見つかりません'); }
    if (st && st.getLastRow() >= 2) workflowCount = st.getLastRow() - 1;
    if (ad && ad.getLastRow() >= 2) adminCount = ad.getLastRow() - 1;
    if (adminCount === 0) { status = 'ng'; reasons.push('管理者が0人です'); }
  } catch (e) { status = 'ng'; reasons.push('スプレッドシート接続失敗: ' + e.message); }

  props.setProperty('LAST_HEALTH_STATE', status);
  props.setProperty('LAST_HEALTH_CHECK', new Date().toISOString());

  if (prev !== status) {
    if (status === 'ng') {
      sendAdminChat_('🚨 ヘルスチェック異常', '直前まで正常だったヘルスチェックが異常を検出しました。\n\n理由:\n- ' + reasons.join('\n- ') + '\n\nワークフロー数: ' + workflowCount + '\n管理者数: ' + adminCount);
    } else if (prev === 'ng') {
      sendAdminChat_('✅ ヘルスチェック復旧', 'ヘルスチェックが正常に戻りました。\n\nワークフロー数: ' + workflowCount + '\n管理者数: ' + adminCount);
    }
  }
  return { status: status, reasons: reasons, workflowCount: workflowCount, adminCount: adminCount };
}

// ---------- Quota カウンタ ----------
// 想定日次上限に対する使用状況を追跡する。
// キー: METRIC_YYYYMMDD_<name> をスクリプトプロパティに保存する。

var QUOTA_LIMITS = {
  approval: 900,      // capacity-plan.md 想定: 300人 × 20申請 × 3操作 / 20営業日
  urlfetch: 20000,    // Google Apps Script Quota
  execTimeSec: 5400   // 90 分/日 = 5400 秒
};

function _metricKey_(name, date) {
  var d = date || new Date();
  var ymd = Utilities.formatDate(d, 'JST', 'yyyyMMdd');
  return 'METRIC_' + ymd + '_' + name;
}

function incrementDailyCounter_(name, incrementBy) {
  try {
    var props = PropertiesService.getScriptProperties();
    var key = _metricKey_(name);
    var cur = parseInt(props.getProperty(key) || '0', 10);
    var next = cur + (incrementBy || 1);
    props.setProperty(key, String(next));
    var limit = QUOTA_LIMITS[name];
    if (limit && cur < limit * 0.8 && next >= limit * 0.8) {
      sendAdminChat_('⚠️ Quota 使用率 80% 到達', name + ' の日次使用量が閾値を超えました。\n\n現在: ' + next + ' / 想定上限: ' + limit + '\n\n負荷急増か仕様変更の可能性があります。');
    }
  } catch (e) { Logger.log('incrementDailyCounter_ 失敗: ' + e.message); }
}

function getQuotaStatus_() {
  var props = PropertiesService.getScriptProperties();
  var out = { date: Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd'), items: [] };
  Object.keys(QUOTA_LIMITS).forEach(function(name) {
    var used = parseInt(props.getProperty(_metricKey_(name)) || '0', 10);
    var limit = QUOTA_LIMITS[name];
    out.items.push({ name: name, used: used, limit: limit, ratio: Math.round(used / limit * 1000) / 10 });
  });
  return out;
}

// トリガーを一括登録する初期化関数(GAS エディタから手動で 1 度実行する)
// 末尾 _ を付けない: private 関数(末尾 _)は GAS UI の実行ドロップダウンに出ないため。
function installMonitorTriggers() {
  var existing = ScriptApp.getProjectTriggers();
  var wanted = ['monitorHealth_', 'dailyBackup_', 'monthlyArchive_'];
  existing.forEach(function(t) { if (wanted.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('monitorHealth_').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('dailyBackup_').timeBased().atHour(0).nearMinute(30).everyDays(1).create();
  ScriptApp.newTrigger('monthlyArchive_').timeBased().onMonthDay(1).atHour(1).create();
  Logger.log('トリガー登録完了: monitorHealth_(5分毎), dailyBackup_(毎日00:30), monthlyArchive_(月初01:00)');
}
