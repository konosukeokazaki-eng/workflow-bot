// ============================================================
// Auth.gs - 認証・セキュリティ・Bot管理・バリデーション・監査ログ
// ============================================================

// メールアドレスから全角空白/NBSP/ゼロ幅文字などの不可視文字を除去し、小文字化する。
// 目的: 承認者/管理者リストに手貼りされた際の見えないゴミによる「本人なのに権限なし」誤判定を防ぐ。
function normalizeEmail_(s) {
  if (s == null) return '';
  return String(s).replace(/[\u0000-\u0020\u00A0\u200B-\u200D\u3000\uFEFF]/g, '').toLowerCase();
}

// 5 分キャッシュ。書き込み系 API(api_saveAdmins など)後は invalidateAdminCache_ を呼ぶこと。
function isAdmin_(email) {
  var e = normalizeEmail_(email);
  if (!e) return false;
  var cache = null, key = null;
  try {
    cache = CacheService.getScriptCache();
    key = 'admin_' + e;
    var cached = cache.get(key);
    if (cached === 'y') return true;
    if (cached === 'n') return false;
  } catch (ex) {}
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_ADMINS);
    if (!sheet || sheet.getLastRow() < 2) { if (cache) try { cache.put(key, 'y', 300); } catch (ex) {} return true; }
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (normalizeEmail_(data[i][0]) === e) {
        if (cache) try { cache.put(key, 'y', 300); } catch (ex) {}
        return true;
      }
    }
    if (cache) try { cache.put(key, 'n', 300); } catch (ex) {}
    return false;
  } catch (err) { Logger.log('isAdmin_ エラー: ' + err.message); return false; }
}

// 管理者リスト変更時に呼び出す。ScriptCache は list 削除不可なので、admins シート全員のキーを個別に削除する。
function invalidateAdminCache_() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_ADMINS);
    if (!sheet || sheet.getLastRow() < 2) return;
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    var keys = data.map(function(r) { return 'admin_' + normalizeEmail_(r[0]); }).filter(function(k) { return k !== 'admin_'; });
    if (keys.length > 0) CacheService.getScriptCache().removeAll(keys);
  } catch (e) { Logger.log('invalidateAdminCache_ 失敗: ' + e.message); }
}

// api_* 関数の先頭で呼ぶ。非管理者ならエラーを throw する。
function requireAdmin_() {
  var email = '';
  try { email = Session.getActiveUser().getEmail(); } catch (e) {}
  if (!email || !isAdmin_(email)) {
    throw new Error('権限がありません');
  }
  return email;
}

// 新トークン形式: base64(sha256(wid_row_issuedAt_SPREADSHEET_ID)).slice(0,16) + '.' + issuedAt
// 旧トークン形式(タイムスタンプなし)も後方互換で許容する。
function generateToken_(wid, rowIndex, issuedAtOpt) {
  var issuedAt = issuedAtOpt || Math.floor(Date.now() / 1000);
  var raw = wid + '_' + rowIndex + '_' + issuedAt + '_' + SPREADSHEET_ID;
  var hash = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw)).substring(0, 16);
  return hash + '.' + issuedAt;
}

// 旧形式(タイムスタンプなし)のトークンを生成。後方互換チェック用。
function generateLegacyToken_(wid, rowIndex) {
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, wid + '_' + rowIndex + '_' + SPREADSHEET_ID)).substring(0, 16);
}

// 検証: 新形式(TTL チェックあり) / 旧形式(TTL チェックなし・遷移期間のみ許容)
function validateToken_(token, wid, rowIndex) {
  if (!token) return { ok: false, reason: 'トークンなし' };
  var parts = String(token).split('.');
  if (parts.length === 2) {
    var hash = parts[0], issuedAt = parseInt(parts[1], 10);
    if (!issuedAt || isNaN(issuedAt)) return { ok: false, reason: 'トークン破損' };
    var ageDays = (Math.floor(Date.now() / 1000) - issuedAt) / 86400;
    if (ageDays > TOKEN_TTL_DAYS) return { ok: false, reason: '期限切れ(' + Math.floor(ageDays) + '日経過)' };
    if (generateToken_(wid, rowIndex, issuedAt) !== token) return { ok: false, reason: '署名不一致' };
    return { ok: true, issuedAt: issuedAt };
  }
  // 旧形式(1.1.0 未満で発行されたトークン)
  if (token === generateLegacyToken_(wid, rowIndex)) {
    Logger.log('レガシートークン受理: wid=' + wid + ' row=' + rowIndex);
    return { ok: true, legacy: true };
  }
  return { ok: false, reason: '無効なトークン' };
}

function getServiceAccountToken_() {
  var kj = JSON.parse(PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT_KEY'));
  var h = Utilities.base64EncodeWebSafe(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  var now = Math.floor(Date.now() / 1000);
  var c = Utilities.base64EncodeWebSafe(JSON.stringify({ iss: kj.client_email, scope: 'https://www.googleapis.com/auth/chat.bot https://www.googleapis.com/auth/chat.memberships.app', aud: 'https://oauth2.googleapis.com/token', iat: now, exp: now + 3600 }));
  var si = h + '.' + c;
  var sig = Utilities.base64EncodeWebSafe(Utilities.computeRsaSha256Signature(si, kj.private_key));
  var tr = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', { method: 'post', contentType: 'application/x-www-form-urlencoded', payload: 'grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=' + si + '.' + sig, muteHttpExceptions: true });
  var td = JSON.parse(tr.getContentText());
  if (td.access_token) return td.access_token;
  throw new Error('トークン取得失敗');
}

function addBotToSpace_(sid) {
  try { var t = getServiceAccountToken_(); UrlFetchApp.fetch('https://chat.googleapis.com/v1/' + sid + '/members', { method: 'post', contentType: 'application/json', headers: { 'Authorization': 'Bearer ' + t }, payload: JSON.stringify({ member: { name: 'users/app', type: 'BOT' } }), muteHttpExceptions: true }); } catch (e) { Logger.log('Bot追加エラー: ' + e.message); }
}

// 指定スペースに Bot が参加しているか検証し、未参加なら追加を試みる。
// 戻り値: { ok: true } | { ok: false, error: string }
// 保存時にサイレント失敗(送信したのに Chat にカードが流れない)を早期検出するために使う。
function verifyBotAccess_(spaceId) {
  if (!spaceId) return { ok: false, error: 'スペースIDが未設定です' };
  try {
    var token = getServiceAccountToken_();
    var url = 'https://chat.googleapis.com/v1/' + spaceId;
    var resp = UrlFetchApp.fetch(url, { method: 'get', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
    var code = resp.getResponseCode();
    if (code === 200) return { ok: true };
    if (code === 403 || code === 404) {
      // Bot 未参加の可能性が高いので追加してリトライ
      try {
        var addResp = UrlFetchApp.fetch('https://chat.googleapis.com/v1/' + spaceId + '/members', { method: 'post', contentType: 'application/json', headers: { 'Authorization': 'Bearer ' + token }, payload: JSON.stringify({ member: { name: 'users/app', type: 'BOT' } }), muteHttpExceptions: true });
        var addCode = addResp.getResponseCode();
        if (addCode >= 400 && addCode !== 409) {
          return { ok: false, error: 'Bot をスペースに追加できません(HTTP ' + addCode + ')。手動でスペースに Bot を招待してください。' };
        }
      } catch (e) {
        return { ok: false, error: 'Bot 追加時に例外: ' + e.message };
      }
      Utilities.sleep(1000);
      var resp2 = UrlFetchApp.fetch(url, { method: 'get', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
      if (resp2.getResponseCode() === 200) return { ok: true };
      return { ok: false, error: 'Bot をスペースに参加させられませんでした。スペース設定を確認してください。' };
    }
    return { ok: false, error: 'スペース確認失敗 HTTP ' + code };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// ---------- 監査ログ ----------

function writeAuditLog_(action, target, detail) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_AUDIT);
    if (!sheet) { sheet = ss.insertSheet(SHEET_AUDIT); sheet.appendRow(['タイムスタンプ', '実行者', 'アクション', '対象', '詳細']); }
    var actor = '';
    try { actor = Session.getActiveUser().getEmail() || ''; } catch (e) {}
    sheet.appendRow([new Date(), actor, action, String(target || ''), String(detail || '').substring(0, 500)]);
  } catch (err) { Logger.log('writeAuditLog_ エラー: ' + err.message); }
}

// ---------- バリデーション ----------

function validateWorkflowName_(name) {
  if (!name || typeof name !== 'string') return 'ワークフロー名は必須です';
  var trimmed = name.trim();
  if (trimmed.length === 0) return 'ワークフロー名は必須です';
  if (trimmed.length > 80) return 'ワークフロー名は80文字以内で入力してください';
  // Google Sheets のシート名で使えない文字を弾く
  if (/[\/\\?\*\[\]:]/.test(trimmed)) return 'シート名として使えない文字が含まれています: / \\ ? * [ ] :';
  return null;
}

function validateEmailList_(str) {
  if (!str) return null;
  var list = String(str).split(',').map(function(s) { return normalizeEmail_(s); }).filter(function(s) { return s; });
  for (var i = 0; i < list.length; i++) {
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(list[i])) return '不正なメール形式: ' + list[i];
  }
  return null;
}

// カンマ区切りメール文字列を正規化(不可視文字除去+小文字化)して返す。空要素は除外。
function sanitizeEmailList_(str) {
  if (!str) return '';
  return String(str).split(',').map(function(s) { return normalizeEmail_(s); }).filter(function(s) { return s; }).join(',');
}

function validateSheetUrl_(url) {
  if (!url) return null;
  var m = String(url).match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (!m) return '外部スプレッドシートURLの形式が不正です';
  return null;
}
