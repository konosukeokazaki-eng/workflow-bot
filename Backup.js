// ============================================================
// Backup.gs - 日次バックアップと月次アーカイブ
// ============================================================
//
// 前提:
//   - スクリプトプロパティ BACKUP_FOLDER_ID にコピー先 Drive フォルダ ID を設定
//   - トリガー "dailyBackup_" を毎日 00:30 に設定(installMonitorTriggers で自動登録)
//   - トリガー "monthlyArchive_" を毎月 1 日 01:00 に設定
//
// 保管ポリシー:
//   - スプレッドシートコピー: 30 世代(30日超は自動削除)
//   - 監査ログ: 1 年オンライン、以降は _監査ログ_アーカイブ シートへ移動
// ============================================================

var BACKUP_RETENTION_DAYS = 30;
var AUDIT_ARCHIVE_DAYS = 365;
var SHEET_AUDIT_ARCHIVE = '_監査ログ_アーカイブ';

// GAS UI から手動テスト実行するためのラッパー(末尾 _ を付けない)。
function runDailyBackupNow() { dailyBackup_(); }
function runMonthlyArchiveNow() { monthlyArchive_(); }

function dailyBackup_() {
  var folderId = getConfig_('BACKUP_FOLDER_ID', '');
  if (!folderId) {
    Logger.log('BACKUP_FOLDER_ID 未設定のためバックアップスキップ');
    sendAdminChat_('⚠️ バックアップ設定未完了', 'スクリプトプロパティ BACKUP_FOLDER_ID が未設定です。Drive の保管先フォルダ ID を設定してください。');
    return;
  }
  try {
    var folder = DriveApp.getFolderById(folderId);
    var src = DriveApp.getFileById(SPREADSHEET_ID);
    var name = 'backup_workflow-bot_' + Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmm');
    src.makeCopy(name, folder);
    purgeOldBackups_(folder);
    writeAuditLog_('backup.daily', name, 'ok');
  } catch (err) {
    Logger.log('dailyBackup_ エラー: ' + err.message);
    notifyAdminOnError_(err, 'dailyBackup_');
  }
}

function purgeOldBackups_(folder) {
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - BACKUP_RETENTION_DAYS);
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (files.hasNext()) {
    var f = files.next();
    if (f.getName().indexOf('backup_workflow-bot_') === 0 && f.getDateCreated() < cutoff) {
      f.setTrashed(true);
    }
  }
}

// 監査ログの 1 年経過分を _監査ログ_アーカイブ シートへ移動
// 実行に時間がかかる可能性があるため、1 回あたり最大 5000 行に制限
function monthlyArchive_() {
  var ss;
  try { ss = SpreadsheetApp.openById(SPREADSHEET_ID); }
  catch (e) { notifyAdminOnError_(e, 'monthlyArchive_ open'); return; }

  archiveAuditLog_(ss);
  purgeArchivedWorkflows_(ss);
}

function archiveAuditLog_(ss) {
  try {
    var src = ss.getSheetByName(SHEET_AUDIT);
    if (!src || src.getLastRow() < 2) return;
    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - AUDIT_ARCHIVE_DAYS);
    var lastRow = src.getLastRow();
    var maxRowsPerRun = 5000;
    var rowsToRead = Math.min(lastRow - 1, maxRowsPerRun);
    var data = src.getRange(2, 1, rowsToRead, 5).getValues();
    var move = [];
    var keepStart = -1;
    for (var i = 0; i < data.length; i++) {
      var ts = data[i][0];
      if (ts instanceof Date && ts < cutoff) {
        move.push(data[i]);
      } else {
        keepStart = i;
        break;
      }
    }
    if (move.length === 0) return;
    var archive = ss.getSheetByName(SHEET_AUDIT_ARCHIVE);
    if (!archive) { archive = ss.insertSheet(SHEET_AUDIT_ARCHIVE); archive.appendRow(['タイムスタンプ','実行者','アクション','対象','詳細']); }
    archive.getRange(archive.getLastRow() + 1, 1, move.length, 5).setValues(move);
    src.deleteRows(2, move.length);
    writeAuditLog_('archive.audit', String(move.length), 'moved to ' + SHEET_AUDIT_ARCHIVE);
  } catch (err) {
    Logger.log('archiveAuditLog_ エラー: ' + err.message);
    notifyAdminOnError_(err, 'archiveAuditLog_');
  }
}

// 30 日以上前に archived になったワークフローを完全削除する。
function purgeArchivedWorkflows_(ss) {
  try {
    var st = ss.getSheetByName(SHEET_SETTINGS);
    if (!st || st.getLastRow() < 2) return;
    var rows = st.getRange(2, 1, st.getLastRow() - 1, 12).getValues();
    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - 30);
    var purged = 0;
    for (var i = rows.length - 1; i >= 0; i--) {
      var status = rows[i][7];
      var archivedAt = rows[i][11];
      if (status === 'archived' && archivedAt instanceof Date && archivedAt < cutoff) {
        var wid = rows[i][0], name = rows[i][1];
        try {
          var ds = ss.getSheetByName(name);
          if (ds && ss.getSheets().length > 1) ss.deleteSheet(ds);
        } catch (e) {}
        st.deleteRow(i + 2);
        var fl = ss.getSheetByName(SHEET_FIELDS);
        if (fl && fl.getLastRow() >= 2) {
          var fd = fl.getRange(2, 1, fl.getLastRow() - 1, 1).getValues();
          for (var j = fd.length - 1; j >= 0; j--) if (fd[j][0] === wid) fl.deleteRow(j + 2);
        }
        var bt = ss.getSheetByName(SHEET_BUTTONS);
        if (bt && bt.getLastRow() >= 2) {
          var bd = bt.getRange(2, 1, bt.getLastRow() - 1, 1).getValues();
          for (var k = bd.length - 1; k >= 0; k--) if (bd[k][0] === wid) bt.deleteRow(k + 2);
        }
        writeAuditLog_('workflow.purge', 'wid=' + wid, name);
        purged++;
      }
    }
    if (purged > 0) sendAdminChat_('🗑️ ワークフロー完全削除', purged + ' 件のアーカイブ済みワークフローを完全削除しました(30日経過)');
  } catch (err) {
    Logger.log('purgeArchivedWorkflows_ エラー: ' + err.message);
    notifyAdminOnError_(err, 'purgeArchivedWorkflows_');
  }
}
