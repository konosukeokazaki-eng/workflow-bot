// ============================================================
// Schema.gs - フィールドIDベースのスキーマ管理 (列ズレ防止)
// ============================================================
// 目的:
//   - _項目設定 に fieldId / deleted / dataCol 列を追加
//   - フィールドの追加・削除・並び替え・改名でデータシートの列がズレないようにする
//   - 削除はソフト削除(列はデータシートに残す)
// ============================================================

// 一意なフィールドIDを生成
function generateFieldId_() {
  return 'f_' + Utilities.getUuid().replace(/-/g, '').substring(0, 16);
}

// _項目設定 に fieldId (7), deleted (8), dataCol (9) 列を保証する。
// 既存行にIDが無ければ採番し、dataCol は order から復元する。
// 冪等: 何度呼んでも安全。
function ensureFieldSchema_(ss) {
  var fl = ss.getSheetByName(SHEET_FIELDS);
  if (!fl) return;
  var lastCol = fl.getLastColumn();
  var needFieldId = lastCol < 7 || fl.getRange(1, 7).getValue() !== 'fieldId';
  var needDeleted = lastCol < 8 || fl.getRange(1, 8).getValue() !== 'deleted';
  var needDataCol = lastCol < 9 || fl.getRange(1, 9).getValue() !== 'dataCol';
  if (needFieldId) fl.getRange(1, 7).setValue('fieldId');
  if (needDeleted) fl.getRange(1, 8).setValue('deleted');
  if (needDataCol) fl.getRange(1, 9).setValue('dataCol');

  if (fl.getLastRow() < 2) return;
  var rows = fl.getRange(2, 1, fl.getLastRow() - 1, 9).getValues();
  var dirty = false;
  for (var i = 0; i < rows.length; i++) {
    if (!rows[i][6]) { rows[i][6] = generateFieldId_(); dirty = true; }
    if (rows[i][7] === '' || rows[i][7] == null) { rows[i][7] = false; dirty = true; }
    // dataCol は order から復元: 先頭3列(ts,sender,email)後にorder順で並ぶ
    if (!rows[i][8]) { rows[i][8] = 3 + Number(rows[i][5] || 0); dirty = true; }
  }
  if (dirty) fl.getRange(2, 1, rows.length, 9).setValues(rows);
}

// 特定 wid のフィールド行を { rowNum, values[] } で全件取得(deleted 含む)
function readFieldRows_(ss, wid) {
  var fl = ss.getSheetByName(SHEET_FIELDS);
  if (!fl || fl.getLastRow() < 2) return [];
  var data = fl.getRange(2, 1, fl.getLastRow() - 1, 9).getValues();
  var out = [];
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === wid) {
      out.push({
        rowNum: i + 2,
        wid: data[i][0], name: data[i][1], type: data[i][2],
        options: data[i][3], required: data[i][4],
        order: data[i][5], fieldId: data[i][6],
        deleted: data[i][7] === true || data[i][7] === 'TRUE' || data[i][7] === 'true',
        dataCol: Number(data[i][8]) || 0
      });
    }
  }
  return out;
}

// 新旧フィールドを fieldId で突合して差分を出す。
// UIから来るnewFieldsに fieldId が無いものは「新規追加」扱い。
// 戻り値: { added:[{new,...}], deleted:[{old,...}], renamed:[{old,new,...}],
//          updated:[{old,new,...}], reordered:bool, unchanged:[] }
function computeFieldDiff_(oldFields, newFields) {
  var oldById = {};
  oldFields.forEach(function(o) { if (o.fieldId) oldById[o.fieldId] = o; });
  var seenIds = {};
  var added = [], updated = [], renamed = [];
  var reordered = false;
  newFields.forEach(function(n, idx) {
    var o = n.fieldId ? oldById[n.fieldId] : null;
    if (!o) {
      added.push({ newField: n, newOrder: idx + 1 });
    } else {
      seenIds[o.fieldId] = true;
      var isRenamed = String(o.name) !== String(n.name);
      var isTypeChanged = String(o.type) !== String(n.type);
      var isOptChanged = String(o.options || '') !== String(n.options || '');
      var isReqChanged = !!o.required !== !!n.required;
      var isOrderChanged = Number(o.order) !== (idx + 1);
      if (isOrderChanged) reordered = true;
      if (isRenamed) renamed.push({ old: o, newField: n, newOrder: idx + 1 });
      if (isTypeChanged || isOptChanged || isReqChanged || isRenamed || isOrderChanged) {
        updated.push({ old: o, newField: n, newOrder: idx + 1 });
      }
    }
  });
  var deleted = oldFields.filter(function(o) { return !o.deleted && !seenIds[o.fieldId]; });
  return { added: added, updated: updated, renamed: renamed, deleted: deleted, reordered: reordered };
}

// 差分をヒトが読める短文で説明(監査ログ・警告表示用)
function describeFieldDiff_(diff) {
  var parts = [];
  if (diff.added.length) parts.push('追加:' + diff.added.map(function(a){return a.newField.name;}).join('/'));
  if (diff.deleted.length) parts.push('削除:' + diff.deleted.map(function(d){return d.name;}).join('/'));
  if (diff.renamed.length) parts.push('改名:' + diff.renamed.map(function(r){return r.old.name+'→'+r.newField.name;}).join('/'));
  if (diff.reordered) parts.push('並替あり');
  return parts.length ? parts.join(' ') : '変更なし';
}

// データシートに必須のシステム列(タイムスタンプ/送信者/送信者メール/(承認型ならステータス)/メッセージ名)が
// 揃っているかを確認し、欠けていれば末尾に補完する。管理者が手動で列を削除した場合の防衛策。
// isApproval: 対象ワークフローが「申請・承認」型なら true(ステータス列を要求)
// 戻り値: 補った列名の配列(空配列なら変更なし)
function ensureDataSheetSchema_(ds, isApproval) {
  if (!ds) return [];
  var lastCol = ds.getLastColumn();
  var headers = lastCol > 0 ? ds.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var required = ['タイムスタンプ', '送信者', '送信者メール'];
  if (isApproval) required.push('ステータス');
  required.push('メッセージ名');
  var added = [];
  required.forEach(function(name) {
    if (headers.indexOf(name) === -1) {
      var newCol = ds.getLastColumn() + 1;
      ds.getRange(1, newCol).setValue(name);
      headers.push(name);
      added.push(name);
    }
  });
  return added;
}

// データシートで「メッセージ名」列の位置を返す(1-indexed)。なければ -1。
function findMessageCol_(ds) {
  if (!ds || ds.getLastColumn() === 0) return -1;
  var headers = ds.getRange(1, 1, 1, ds.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) if (headers[i] === 'メッセージ名') return i + 1;
  return -1;
}

// データシートで「ステータス」列の位置を返す(1-indexed)。なければ -1。
function findStatusCol_(ds) {
  if (!ds || ds.getLastColumn() === 0) return -1;
  var headers = ds.getRange(1, 1, 1, ds.getLastColumn()).getValues()[0];
  for (var i = 0; i < headers.length; i++) if (headers[i] === 'ステータス') return i + 1;
  return -1;
}

// データシートの列レイアウトを正規化する:
//   [タイムスタンプ, 送信者, 送信者メール, ...activeFields(order順), ステータス?, メッセージ名, ...deletedFields]
// activeFields/deletedFields は _項目設定 の現状値(fieldId, dataCol)を基に決定する。
// 実装は全データの読み出し → 新順で1回だけ setValues で書き戻し(原子的で安全)。
// 書き戻し後、_項目設定 の dataCol を新位置に更新する。
// 呼び出し前提: ensureFieldSchema_ 済み、_項目設定 は最新状態、データシートには不要な列がない。
function canonicalizeDataSheetLayout_(ss, ds, wid, isApproval) {
  var lastRow = ds.getLastRow();
  var lastCol = ds.getLastColumn();
  if (lastCol === 0) return;
  var headerRow = ds.getRange(1, 1, 1, lastCol).getValues()[0];

  var fieldRows = readFieldRows_(ss, wid);
  var active = fieldRows.filter(function(r) { return !r.deleted; }).slice().sort(function(a, b) { return a.order - b.order; });
  var deleted = fieldRows.filter(function(r) { return r.deleted; });

  // 特殊列の現在位置を検索
  var findByHeader = function(name) { for (var i = 0; i < headerRow.length; i++) if (headerRow[i] === name) return i + 1; return -1; };
  var tsCol = findByHeader('タイムスタンプ');
  var senderCol = findByHeader('送信者');
  var emailCol = findByHeader('送信者メール');
  var statusCol = findByHeader('ステータス');
  var msgCol = findByHeader('メッセージ名');

  // 目標順(1-indexed の旧列を並べる)
  var orderList = [];
  if (tsCol > 0) orderList.push(tsCol);
  if (senderCol > 0) orderList.push(senderCol);
  if (emailCol > 0) orderList.push(emailCol);
  active.forEach(function(f) { if (f.dataCol > 0 && f.dataCol <= lastCol) orderList.push(f.dataCol); });
  if (isApproval && statusCol > 0) orderList.push(statusCol);
  if (msgCol > 0) orderList.push(msgCol);
  deleted.forEach(function(f) { if (f.dataCol > 0 && f.dataCol <= lastCol) orderList.push(f.dataCol); });

  // すべての既存列を過不足なくカバーしているか確認
  if (orderList.length !== lastCol) {
    // 見落とし列を末尾に足して整合を取る(データを消さないための安全策)
    var seen = {};
    orderList.forEach(function(c) { seen[c] = true; });
    for (var c = 1; c <= lastCol; c++) if (!seen[c]) orderList.push(c);
  }
  var duplicates = {};
  var deduped = [];
  orderList.forEach(function(c) { if (!duplicates[c]) { duplicates[c] = true; deduped.push(c); } });
  orderList = deduped;

  // 既に正しい順序なら何もしない
  var alreadyOrdered = true;
  for (var i = 0; i < orderList.length; i++) if (orderList[i] !== i + 1) { alreadyOrdered = false; break; }

  if (!alreadyOrdered) {
    var allData = ds.getRange(1, 1, lastRow, lastCol).getValues();
    var reordered = allData.map(function(row) { return orderList.map(function(c) { return row[c - 1]; }); });
    ds.getRange(1, 1, reordered.length, reordered[0].length).setValues(reordered);
  }

  // 新しい dataCol を fieldId ごとに算出して _項目設定 に反映
  var newColByFid = {};
  var pos = 1;
  if (tsCol > 0) pos++;
  if (senderCol > 0) pos++;
  if (emailCol > 0) pos++;
  active.forEach(function(f) { newColByFid[f.fieldId] = pos++; });
  if (isApproval && statusCol > 0) pos++;
  if (msgCol > 0) pos++;
  deleted.forEach(function(f) { newColByFid[f.fieldId] = pos++; });

  var fl = ss.getSheetByName(SHEET_FIELDS);
  if (fl && fl.getLastRow() >= 2) {
    var flData = fl.getRange(2, 1, fl.getLastRow() - 1, 9).getValues();
    for (var j = 0; j < flData.length; j++) {
      if (flData[j][0] === wid && flData[j][6] && newColByFid[flData[j][6]]) {
        var target = newColByFid[flData[j][6]];
        if (Number(flData[j][8]) !== target) fl.getRange(j + 2, 9).setValue(target);
      }
    }
  }
}
