/**
 * 汚濁負荷量 共有バックエンド v2 (Google Apps Script Web App)
 * - ユーザー個別 ID/パスワード認証（退職者は個別に停止/削除＝即ログイン不可）
 * - セッショントークン / LockService 直列化 / 日単位マージ / rev 管理
 * - データ用スプレッドシートは未指定なら自動生成（このGASの所有Googleアカウントのドライブ内）
 *
 * 必要なスクリプトプロパティ（最小）:
 *   ODAKU_ADMIN_USER  初期管理者ID（省略時 admin）
 *   ODAKU_ADMIN_PASS  初期管理者パスワード（省略時 change-me-now）
 *   ODAKU_SHEET_ID    省略可（無ければ自動生成してここに記録）
 */

var P_SHEET = 'ODAKU_SHEET_ID';
var P_ADMIN_U = 'ODAKU_ADMIN_USER';
var P_ADMIN_P = 'ODAKU_ADMIN_PASS';
var SESSION_HOURS = 12;
var HASH_ROUNDS = 1000;

function props_() { return PropertiesService.getScriptProperties(); }

function ensureSpreadsheet_() {
  var p = props_(); var id = p.getProperty(P_SHEET); var ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) { ss = SpreadsheetApp.create('汚濁負荷量_データ'); p.setProperty(P_SHEET, ss.getId()); }
  return ss;
}
function sheet_(name, headers) {
  var s = ensureSpreadsheet_(); var sh = s.getSheetByName(name);
  if (!sh) { sh = s.insertSheet(name); sh.appendRow(headers); }
  return sh;
}
function usersSheet_() { return sheet_('USERS', ['userId', 'name', 'salt', 'hash', 'active', 'role', 'createdAt', 'updatedAt']); }
function sessSheet_()  { return sheet_('SESSIONS', ['token', 'userId', 'createdAt', 'expiresAt']); }
function dataSheet_()  { return sheet_('DATA', ['key', 'fy', 'plant', 'month', 'json', 'rev', 'updatedAt', 'updatedBy']); }
function metaSheet_()  { return sheet_('META', ['plant', 'json', 'updatedAt', 'updatedBy']); }

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function safeParse_(s) { try { return s ? JSON.parse(s) : null; } catch (e) { return null; } }
function emptyGrid_() { var a = []; for (var i = 0; i < 31; i++) a.push({}); return a; }
function keyOf_(fy, plant, month) { return fy + '__' + plant + '__' + month; }
function uuid_() { return Utilities.getUuid().replace(/-/g, ''); }
function toBool_(v) { return v === true || v === 'TRUE' || v === 'true' || v === 1; }
function hash_(salt, pw) {
  var s = salt + '|' + pw;
  for (var i = 0; i < HASH_ROUNDS; i++) {
    s = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s + i));
  }
  return s;
}

function ensureSetup_() {
  usersSheet_(); sessSheet_(); dataSheet_(); metaSheet_();
  var sh = usersSheet_();
  if (sh.getLastRow() <= 1) {
    var p = props_();
    upsertUser_(p.getProperty(P_ADMIN_U) || 'admin', '管理者', p.getProperty(P_ADMIN_P) || 'change-me-now', 'admin', true);
  }
}

function findUserRow_(userId) {
  var sh = usersSheet_(); var last = sh.getLastRow(); if (last < 2) return -1;
  var v = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]) === String(userId)) return i + 2;
  return -1;
}
function getUser_(userId) {
  var r = findUserRow_(userId); if (r < 0) return null;
  var v = usersSheet_().getRange(r, 1, 1, 8).getValues()[0];
  return { row: r, userId: v[0], name: v[1], salt: v[2], hash: v[3], active: toBool_(v[4]), role: v[5] };
}
function upsertUser_(userId, name, password, role, active) {
  var sh = usersSheet_(); var r = findUserRow_(userId); var now = new Date(); var salt, hash;
  if (password) { salt = uuid_(); hash = hash_(salt, password); }
  if (r < 0) {
    if (!password) { salt = ''; hash = ''; }
    sh.appendRow([userId, name || userId, salt, hash, active !== false, role || 'user', now, now]);
  } else {
    var cur = sh.getRange(r, 1, 1, 8).getValues()[0];
    if (!password) { salt = cur[2]; hash = cur[3]; }
    sh.getRange(r, 1, 1, 8).setValues([[userId, name || cur[1], salt, hash, active !== false, role || cur[5] || 'user', cur[6] || now, now]]);
  }
}

function pruneSessions_() {
  var sh = sessSheet_(); var last = sh.getLastRow(); if (last < 2) return;
  var v = sh.getRange(2, 1, last - 1, 4).getValues(); var now = Date.now(); var keep = [];
  for (var i = 0; i < v.length; i++) if (new Date(v[i][3]).getTime() > now) keep.push(v[i]);
  if (keep.length !== v.length) {
    sh.getRange(2, 1, v.length, 4).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, 4).setValues(keep);
  }
}
function auth_(token) {
  if (!token) return null;
  var sh = sessSheet_(); var last = sh.getLastRow(); if (last < 2) return null;
  var v = sh.getRange(2, 1, last - 1, 4).getValues();
  for (var i = 0; i < v.length; i++) {
    if (v[i][0] === token) {
      if (Date.now() > new Date(v[i][3]).getTime()) return null;
      var u = getUser_(v[i][1]); if (!u || !u.active) return null; return u;
    }
  }
  return null;
}

function doGet() { return json_({ ok: true, service: 'odaku', time: Date.now() }); }

function doPost(e) {
  try {
    ensureSetup_();
    var req = JSON.parse(e.postData.contents);
    var act = req.action;
    if (act === 'login') return json_(login_(req));
    if (act === 'logout') return json_(logout_(req));
    var u = auth_(req.token);
    if (!u) return json_({ ok: false, error: 'AUTH' });
    switch (act) {
      case 'me':        return json_({ ok: true, user: pubUser_(u) });
      case 'bootstrap': return json_(bootstrap_(req));
      case 'saveDays':  return json_(saveDays_(req, u));
      case 'saveMeta':  return json_(saveMeta_(req, u));
      case 'pull':      return json_(pull_(req));
      case 'adminListUsers': return json_(adminList_(u));
      case 'adminUpsertUser': return json_(adminUpsert_(req, u));
      case 'adminDeleteUser': return json_(adminDelete_(req, u));
      default: return json_({ ok: false, error: 'BAD_ACTION' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function pubUser_(u) { return { userId: u.userId, name: u.name, role: u.role }; }

function login_(req) {
  pruneSessions_();
  var u = getUser_(req.userId);
  if (!u || !u.active || hash_(u.salt, req.password) !== u.hash) return { ok: false, error: 'LOGIN' };
  var token = uuid_() + uuid_(); var now = Date.now(); var exp = now + SESSION_HOURS * 3600 * 1000;
  sessSheet_().appendRow([token, u.userId, new Date(now), new Date(exp)]);
  return { ok: true, token: token, user: pubUser_(u), expiresAt: exp };
}
function logout_(req) {
  var sh = sessSheet_(); var last = sh.getLastRow(); if (last < 2) return { ok: true };
  var v = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) if (v[i][0] === req.token) { sh.deleteRow(i + 2); break; }
  return { ok: true };
}

// ---- 管理（admin のみ）----
function requireAdmin_(u) { if (!u || u.role !== 'admin') throw 'FORBIDDEN'; }
function adminList_(u) {
  requireAdmin_(u);
  var sh = usersSheet_(); var last = sh.getLastRow(); var out = [];
  if (last >= 2) {
    var v = sh.getRange(2, 1, last - 1, 8).getValues();
    for (var i = 0; i < v.length; i++) out.push({ userId: v[i][0], name: v[i][1], active: toBool_(v[i][4]), role: v[i][5], updatedAt: v[i][7] });
  }
  return { ok: true, users: out };
}
function adminUpsert_(req, u) {
  requireAdmin_(u);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { upsertUser_(req.userId, req.name, req.password || '', req.role || 'user', req.active !== false); return { ok: true }; }
  finally { lock.releaseLock(); }
}
function adminDelete_(req, u) {
  requireAdmin_(u);
  if (req.userId === u.userId) return { ok: false, error: 'SELF' };
  var sh = usersSheet_(); var r = findUserRow_(req.userId); if (r > 0) sh.deleteRow(r);
  // そのユーザーのセッションも全削除
  var ssh = sessSheet_(); var last = ssh.getLastRow();
  if (last >= 2) { var v = ssh.getRange(2, 1, last - 1, 2).getValues(); for (var i = v.length - 1; i >= 0; i--) if (v[i][1] === req.userId) ssh.deleteRow(i + 2); }
  return { ok: true };
}

// ---- データ ----
function readMeta_() {
  var ms = metaSheet_(); var mv = ms.getDataRange().getValues(); var meta = {};
  for (var i = 1; i < mv.length; i++) if (mv[i][0]) meta[mv[i][0]] = safeParse_(mv[i][1]) || {};
  return meta;
}
function readFy_(fy) {
  var sh = dataSheet_(); var vals = sh.getDataRange().getValues(); var out = {};
  for (var i = 1; i < vals.length; i++) {
    var r = vals[i]; if (String(r[1]) !== String(fy)) continue;
    out[r[0]] = { plant: r[2], month: Number(r[3]), grid: safeParse_(r[4]) || emptyGrid_(), rev: Number(r[5] || 0) };
  }
  return out;
}
function bootstrap_(req) { return { ok: true, fy: req.fy, rows: readFy_(req.fy), meta: readMeta_(), serverTime: Date.now() }; }
function findRow_(sh, key) {
  var last = sh.getLastRow(); if (last < 2) return -1;
  var v = sh.getRange(1, 1, last, 1).getValues();
  for (var i = 1; i < v.length; i++) if (v[i][0] === key) return i + 1;
  return -1;
}
function saveDays_(req, u) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var fy = req.fy, plant = req.plant, month = Number(req.month), changed = req.changedDays || {};
    var key = keyOf_(fy, plant, month); var sh = dataSheet_(); var rownum = findRow_(sh, key); var grid, rev;
    if (rownum < 0) {
      grid = emptyGrid_(); rev = 0; rownum = sh.getLastRow() + 1;
      sh.getRange(rownum, 1, 1, 8).setValues([[key, fy, plant, month, '', 0, '', '']]);
    } else {
      var row = sh.getRange(rownum, 1, 1, 8).getValues()[0];
      grid = safeParse_(row[4]); if (!grid || !grid.length) grid = emptyGrid_(); rev = Number(row[5] || 0);
    }
    Object.keys(changed).forEach(function (idx) { grid[Number(idx)] = changed[idx] || {}; });
    rev++; var now = new Date();
    sh.getRange(rownum, 5, 1, 4).setValues([[JSON.stringify(grid), rev, now, u.name || u.userId]]);
    return { ok: true, key: key, grid: grid, rev: rev, updatedAt: now.getTime() };
  } finally { lock.releaseLock(); }
}
function saveMeta_(req, u) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = metaSheet_(); var rownum = findRow_(sh, req.plant); var now = new Date(); var j = JSON.stringify(req.meta || {});
    if (rownum < 0) sh.appendRow([req.plant, j, now, u.name || u.userId]);
    else sh.getRange(rownum, 2, 1, 3).setValues([[j, now, u.name || u.userId]]);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
function pull_(req) {
  var clientRevs = req.revs || {}; var rows = readFy_(req.fy); var changed = {};
  Object.keys(rows).forEach(function (k) { if (Number(clientRevs[k]) !== rows[k].rev) changed[k] = rows[k]; });
  return { ok: true, changed: changed, meta: readMeta_(), serverTime: Date.now() };
}
