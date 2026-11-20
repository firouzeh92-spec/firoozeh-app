/**
 * فایلینگ شخصی - Google Apps Script (Container-bound به گوگل شیت)
 * تب‌ها: Files (فایل‌های مشترک) | Users | Saved (ذخیره‌های هر کاربر) | Settings
 */

const SH = { FILES: 'Files', USERS: 'Users', SAVED: 'Saved', SETTINGS: 'Settings', CUST: 'Customers' };

const FILE_COLS = ['Code', 'Date', 'DealType', 'PropertyType', 'Address', 'Area', 'Foundation',
  'TotalFloors', 'TotalUnits', 'Floor', 'Rooms', 'Direction', 'Facade', 'Flooring', 'Wall',
  'Cabinet', 'Cooler', 'Age', 'Elevator', 'Parking', 'Description', 'Price', 'Deposit', 'Rent',
  'Owner', 'Mobile', 'Source', 'AddedBy'];
const USER_COLS = ['Username', 'DisplayName', 'Role', 'Salt', 'Hash', 'Active', 'CreatedAt'];
const SAVED_COLS = ['Username', 'Code', 'Status', 'Note', 'SavedAt', 'UpdatedAt', 'Snapshot'];

const CUST_COLS = ['ID', 'Username', 'Name', 'Phone', 'Deal', 'Areas', 'Budget', 'Rent', 'Notes', 'Log', 'CreatedAt', 'UpdatedAt'];

const SESSION_DAYS = 30;
const DEFAULT_SIGNATURE = 'برای بازدید و اطلاعات بیشتر تماس بگیرید';

/* ===================== وب اپ ===================== */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('فایلینگ مشاورین املاک فیروزه')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ===================== منوی شیت و راه‌اندازی ===================== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('فایلینگ مشاورین املاک فیروزه')
    .addItem('۱. راه‌اندازی اولیه (ساخت تب‌ها)', 'setup')
    .addItem('۲. ساخت / ریست کاربر مدیر', 'createAdmin')
    .addItem('۳. حذف فایل‌های تکراری (بر اساس Code)', 'removeDuplicatesMenu')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());
  ensureSheet_(ss, SH.FILES, FILE_COLS);
  ensureSheet_(ss, SH.USERS, USER_COLS);
  ensureSheet_(ss, SH.SAVED, SAVED_COLS);
  ensureSheet_(ss, SH.CUST, CUST_COLS);
  const st = ensureSheet_(ss, SH.SETTINGS, ['Key', 'Value']);
  if (st.getLastRow() < 2) st.appendRow(['AdSignature', DEFAULT_SIGNATURE]);
  secret_();
  SpreadsheetApp.getUi().alert('تب‌ها ساخته شد. حالا از منوی فایلینگ مشاورین املاک فیروزه، «ساخت / ریست کاربر مدیر» را بزنید.');
}

function ensureSheet_(ss, name, cols) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function createAdmin() {
  const ui = SpreadsheetApp.getUi();
  setup_();
  const u = ui.prompt('نام کاربری مدیر (انگلیسی، بدون فاصله)');
  if (u.getSelectedButton() !== ui.Button.OK) return;
  const username = u.getResponseText().trim().toLowerCase();
  const p = ui.prompt('رمز عبور مدیر (حداقل ۶ کاراکتر)');
  if (p.getSelectedButton() !== ui.Button.OK) return;
  const pass = p.getResponseText();
  if (!/^[a-z0-9_.]{3,}$/.test(username) || pass.length < 6) {
    ui.alert('نام کاربری یا رمز معتبر نیست.');
    return;
  }
  upsertUser_({ username: username, name: 'مدیر', role: 'admin', password: pass, active: true });
  ui.alert('کاربر مدیر ساخته شد. حالا وب اپ را Deploy کنید و با همین مشخصات وارد شوید.');
}

function setup_() {
  if (!PropertiesService.getScriptProperties().getProperty('SS_ID')) setup();
}

function removeDuplicatesMenu() {
  const n = removeDuplicates_();
  SpreadsheetApp.getUi().alert(n + ' ردیف تکراری حذف شد.');
}

function removeDuplicates_() {
  const sh = sheet_(SH.FILES);
  const last = sh.getLastRow();
  if (last < 3) return 0;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(String);
  const ci = headers.indexOf('Code');
  const codes = sh.getRange(2, ci + 1, last - 1, 1).getDisplayValues();
  const seen = {};
  const del = [];
  for (let i = codes.length - 1; i >= 0; i--) { // آخرین نسخه حفظ می‌شود
    const c = norm_(codes[i][0]);
    if (!c) continue;
    if (seen[c]) del.push(i + 2); else seen[c] = 1;
  }
  del.forEach(function (r) { sh.deleteRow(r); }); // از پایین به بالا
  return del.length;
}

/* ===================== ابزارهای شیت ===================== */

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}
function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('تب ' + name + ' پیدا نشد. راه‌اندازی اولیه را اجرا کنید.');
  return sh;
}
function norm_(v) {
  return String(v == null ? '' : v)
    .replace(/[۰-۹]/g, function (d) { return '۰۱۲۳۴۵۶۷۸۹'.indexOf(d); })
    .replace(/[٠-٩]/g, function (d) { return '٠١٢٣٤٥٦٧٨٩'.indexOf(d); })
    .replace(/\u064A/g, '\u06CC').replace(/\u0643/g, '\u06A9').trim();
}
function readObjects_(sh) {
  const vals = sh.getDataRange().getDisplayValues();
  if (vals.length < 2) return [];
  const headers = vals[0].map(function (h) { return String(h).trim(); });
  const out = [];
  for (let i = 1; i < vals.length; i++) {
    const o = {};
    let empty = true;
    for (let j = 0; j < headers.length; j++) {
      if (!headers[j]) continue;
      o[headers[j]] = vals[i][j];
      if (vals[i][j] !== '') empty = false;
    }
    if (!empty) { o._row = i + 1; out.push(o); }
  }
  return out;
}
function appendText_(sh, cols, obj) {
  // همه به صورت متن ذخیره می‌شود تا صفر ابتدای موبایل حذف نشود
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(String);
  const row = headers.map(function (h) { return obj[h] == null ? '' : String(obj[h]); });
  const r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, row.length).setNumberFormat('@').setValues([row]);
  return r;
}
function getSettings_() {
  const o = {};
  readObjects_(sheet_(SH.SETTINGS)).forEach(function (r) { o[r.Key] = r.Value; });
  return o;
}

/* ===================== احراز هویت ===================== */

function secret_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); }
  return s;
}
function hash_(salt, pass) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pass, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function sign_(payload) {
  const sig = Utilities.computeHmacSha256Signature(payload, secret_());
  return Utilities.base64EncodeWebSafe(sig);
}
function makeToken_(username) {
  const exp = Date.now() + SESSION_DAYS * 86400000;
  const payload = Utilities.base64EncodeWebSafe(username + '|' + exp, Utilities.Charset.UTF_8);
  return payload + '.' + sign_(payload);
}
function auth_(token) {
  if (!token || token.indexOf('.') < 0) throw new Error('AUTH');
  const parts = token.split('.');
  if (sign_(parts[0]) !== parts[1]) throw new Error('AUTH');
  const dec = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
  const bits = dec.split('|');
  if (Number(bits[1]) < Date.now()) throw new Error('AUTH');
  const u = findUser_(bits[0]);
  if (!u || String(u.Active).toLowerCase() === 'false') throw new Error('AUTH');
  return { username: u.Username, name: u.DisplayName || u.Username, role: u.Role === 'admin' ? 'admin' : 'user' };
}
function findUser_(username) {
  const key = String(username).trim().toLowerCase();
  const list = readObjects_(sheet_(SH.USERS));
  for (let i = 0; i < list.length; i++) if (String(list[i].Username).toLowerCase() === key) return list[i];
  return null;
}
function requireAdmin_(user) { if (user.role !== 'admin') throw new Error('دسترسی مجاز نیست.'); }

function api_login(username, password) {
  const key = 'fail_' + String(username).trim().toLowerCase();
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get(key) || 0);
  if (fails >= 5) throw new Error('تعداد تلاش بیش از حد. ۱۰ دقیقه بعد دوباره امتحان کنید.');
  const u = findUser_(username);
  const ok = u && String(u.Active).toLowerCase() !== 'false' && hash_(u.Salt, password) === u.Hash;
  if (!ok) {
    cache.put(key, String(fails + 1), 600);
    throw new Error('نام کاربری یا رمز عبور اشتباه است.');
  }
  cache.remove(key);
  return {
    token: makeToken_(u.Username),
    user: { username: u.Username, name: u.DisplayName || u.Username, role: u.Role === 'admin' ? 'admin' : 'user' }
  };
}

/* ===================== داده‌ها ===================== */

function api_bootstrap(token) {
  const user = auth_(token);
  const files = readObjects_(sheet_(SH.FILES)).map(function (f) { delete f._row; return f; });
  const saved = readObjects_(sheet_(SH.SAVED))
    .filter(function (s) { return String(s.Username).toLowerCase() === user.username.toLowerCase(); })
    .map(function (s) { delete s._row; return s; });
  const st = getSettings_();
  return { user: user, files: files, customers: custList_(user), saved: saved, settings: { adSignature: st.AdSignature || DEFAULT_SIGNATURE } };
}

function api_saveItem(token, code, status, note, snapshot) {
  const user = auth_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(SH.SAVED);
    const now = new Date().toISOString();
    const rows = readObjects_(sh);
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (String(r.Username).toLowerCase() === user.username.toLowerCase() && norm_(r.Code) === norm_(code)) {
        sh.getRange(r._row, 3, 1, 4).setNumberFormat('@').setValues([[status || '', note || '', r.SavedAt, now]]);
        return { Username: user.username, Code: code, Status: status || '', Note: note || '', SavedAt: r.SavedAt, UpdatedAt: now };
      }
    }
    const obj = { Username: user.username, Code: code, Status: status || '', Note: note || '', SavedAt: now, UpdatedAt: now,
      Snapshot: JSON.stringify(snapshot || {}).slice(0, 45000) };
    appendText_(sh, SAVED_COLS, obj);
    delete obj.Snapshot;
    return obj;
  } finally { lock.releaseLock(); }
}

function api_unsaveItem(token, code) {
  const user = auth_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(SH.SAVED);
    const rows = readObjects_(sh);
    for (let i = rows.length - 1; i >= 0; i--) {
      const r = rows[i];
      if (String(r.Username).toLowerCase() === user.username.toLowerCase() && norm_(r.Code) === norm_(code)) sh.deleteRow(r._row);
    }
    return true;
  } finally { lock.releaseLock(); }
}

function api_getSnapshot(token, code) {
  const user = auth_(token);
  const rows = readObjects_(sheet_(SH.SAVED));
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i].Username).toLowerCase() === user.username.toLowerCase() && norm_(rows[i].Code) === norm_(code)) {
      try { return JSON.parse(rows[i].Snapshot); } catch (e) { return null; }
    }
  }
  return null;
}

function api_addFile(token, f) {
  const user = auth_(token);
  if (!f || !String(f.Address || '').trim()) throw new Error('آدرس الزامی است.');
  if (!String(f.DealType || '').trim()) throw new Error('نوع معامله الزامی است.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const code = 'M' + Utilities.formatDate(new Date(), 'Asia/Tehran', 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
    const obj = {};
    FILE_COLS.forEach(function (c) { obj[c] = f[c] == null ? '' : String(f[c]).trim(); });
    obj.Code = code;
    obj.Source = 'دستی';
    obj.AddedBy = user.username;
    appendText_(sheet_(SH.FILES), FILE_COLS, obj);
    return obj;
  } finally { lock.releaseLock(); }
}

function api_deleteFile(token, code) {
  const user = auth_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(SH.FILES);
    const rows = readObjects_(sh);
    for (let i = rows.length - 1; i >= 0; i--) {
      const r = rows[i];
      if (norm_(r.Code) !== norm_(code)) continue;
      const mine = r.Source === 'دستی' && String(r.AddedBy).toLowerCase() === user.username.toLowerCase();
      if (user.role !== 'admin' && !mine) throw new Error('فقط مدیر یا ثبت‌کننده می‌تواند حذف کند.');
      sh.deleteRow(r._row);
      return true;
    }
    return false;
  } finally { lock.releaseLock(); }
}

/* ===================== مدیریت ===================== */

function upsertUser_(u) {
  const sh = sheet_(SH.USERS);
  const existing = findUser_(u.username);
  const salt = Utilities.getUuid();
  if (existing) {
    const r = existing._row;
    sh.getRange(r, 2, 1, 3).setNumberFormat('@').setValues([[u.name || existing.DisplayName, u.role || existing.Role, existing.Salt]]);
    if (u.password) {
      sh.getRange(r, 4, 1, 2).setNumberFormat('@').setValues([[salt, hash_(salt, u.password)]]);
    }
    if (u.active !== undefined) sh.getRange(r, 6).setNumberFormat('@').setValue(u.active ? 'TRUE' : 'FALSE');
  } else {
    if (!u.password) throw new Error('رمز عبور الزامی است.');
    appendText_(sh, USER_COLS, {
      Username: u.username.toLowerCase(), DisplayName: u.name || u.username, Role: u.role || 'user',
      Salt: salt, Hash: hash_(salt, u.password), Active: u.active === false ? 'FALSE' : 'TRUE',
      CreatedAt: new Date().toISOString()
    });
  }
}

function api_listUsers(token) {
  requireAdmin_(auth_(token));
  return readObjects_(sheet_(SH.USERS)).map(function (u) {
    return { Username: u.Username, DisplayName: u.DisplayName, Role: u.Role, Active: String(u.Active).toLowerCase() !== 'false' };
  });
}

function api_saveUser(token, u) {
  const me = auth_(token);
  requireAdmin_(me);
  if (!/^[a-zA-Z0-9_.]{3,}$/.test(String(u.username || ''))) throw new Error('نام کاربری باید انگلیسی و حداقل ۳ حرف باشد.');
  if (u.password && u.password.length < 6) throw new Error('رمز عبور حداقل ۶ کاراکتر باشد.');
  if (u.username.toLowerCase() === me.username.toLowerCase() && (u.active === false || u.role === 'user')) {
    throw new Error('نمی‌توانید دسترسی خودتان را کم کنید.');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { upsertUser_(u); } finally { lock.releaseLock(); }
  return api_listUsers(token);
}

function api_setSignature(token, text) {
  requireAdmin_(auth_(token));
  const sh = sheet_(SH.SETTINGS);
  const rows = readObjects_(sh);
  for (let i = 0; i < rows.length; i++) {
    if (rows[i].Key === 'AdSignature') { sh.getRange(rows[i]._row, 2).setNumberFormat('@').setValue(text); return true; }
  }
  sh.appendRow(['AdSignature', text]);
  return true;
}

function api_removeDuplicates(token) {
  requireAdmin_(auth_(token));
  return removeDuplicates_();
}


/* ===================== مشتریان ===================== */

function custSheet_() {
  return ensureSheet_(ss_(), SH.CUST, CUST_COLS);
}
function custList_(user) {
  return readObjects_(custSheet_())
    .filter(function (c) { return user.role === 'admin' || String(c.Username).toLowerCase() === user.username.toLowerCase(); })
    .map(function (c) { delete c._row; return c; });
}
function custFind_(id) {
  const list = readObjects_(custSheet_());
  for (let i = 0; i < list.length; i++) if (String(list[i].ID) === String(id)) return list[i];
  return null;
}
function custCanEdit_(user, c) {
  if (!c) throw new Error('مشتری پیدا نشد.');
  if (user.role !== 'admin' && String(c.Username).toLowerCase() !== user.username.toLowerCase()) throw new Error('دسترسی مجاز نیست.');
}
function custWrite_(row, obj) {
  const sh = custSheet_();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(String);
  const vals = headers.map(function (h) { return obj[h] == null ? '' : String(obj[h]); });
  sh.getRange(row, 1, 1, vals.length).setNumberFormat('@').setValues([vals]);
}
function custDate_(s) {
  const d = String(s || '').replace(/[۰-۹]/g, function (x) { return x.charCodeAt(0) - 1776; }).replace(/[٠-٩]/g, function (x) { return x.charCodeAt(0) - 1632; });
  const m = d.match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  return m ? m[1] + '/' + ('0' + m[2]).slice(-2) + '/' + ('0' + m[3]).slice(-2) : '';
}
function custClean_(c) {
  const t = function (v) { return norm_(v).slice(0, 2000); };
  return {
    Name: t(c.Name), Phone: t(c.Phone), Deal: c.Deal === 'rent' ? 'rent' : 'sale',
    Areas: t(c.Areas), Budget: t(c.Budget).replace(/[^\d]/g, ''), Rent: t(c.Rent).replace(/[^\d]/g, ''),
    Notes: t(c.Notes), CreatedAt: custDate_(c.Date), UpdatedAt: custDate_(c.Updated)
  };
}
function api_saveCustomer(token, c) {
  const user = auth_(token);
  if (!c || !String(c.Name || '').trim()) throw new Error('نام مشتری را وارد کنید.');
  const clean = custClean_(c);
  if (c.ID) {
    const old = custFind_(c.ID);
    custCanEdit_(user, old);
    if (!clean.CreatedAt) clean.CreatedAt = old.CreatedAt;
    if (!clean.UpdatedAt) clean.UpdatedAt = old.UpdatedAt;
    const obj = Object.assign({}, old, clean);
    delete obj._row;
    custWrite_(old._row, obj);
    return obj;
  }
  const obj = Object.assign({ ID: Utilities.getUuid().slice(0, 8), Username: user.username, Log: '[]' }, clean);
  appendText_(custSheet_(), CUST_COLS, obj);
  return obj;
}
function api_deleteCustomer(token, id) {
  const user = auth_(token);
  const c = custFind_(id);
  custCanEdit_(user, c);
  custSheet_().deleteRow(c._row);
  return true;
}
function api_addCustomerLog(token, id, date, text) {
  const user = auth_(token);
  const c = custFind_(id);
  custCanEdit_(user, c);
  if (!String(text || '').trim()) throw new Error('خلاصه مکالمه را بنویسید.');
  let log = [];
  try { log = JSON.parse(c.Log || '[]'); } catch (e) {}
  log.unshift({ d: norm_(date), t: String(text).slice(0, 1000) });
  c.Log = JSON.stringify(log);
  c.UpdatedAt = custDate_(date) || c.UpdatedAt;
  const row = c._row; delete c._row;
  custWrite_(row, c);
  return c;
}
function api_deleteCustomerLog(token, id, index) {
  const user = auth_(token);
  const c = custFind_(id);
  custCanEdit_(user, c);
  let log = [];
  try { log = JSON.parse(c.Log || '[]'); } catch (e) {}
  log.splice(Number(index), 1);
  c.Log = JSON.stringify(log);
  const row = c._row; delete c._row;
  custWrite_(row, c);
  return c;
}


/* ===================== درگاه اپ اندروید (فقط افزوده شده؛ هیچ تابع قبلی تغییر نکرده) ===================== */

const ANDROID_API_ = ['api_login', 'api_bootstrap', 'api_saveItem', 'api_unsaveItem', 'api_getSnapshot', 'api_addFile',
  'api_deleteFile', 'api_listUsers', 'api_saveUser', 'api_setSignature', 'api_removeDuplicates', 'api_saveCustomer',
  'api_deleteCustomer', 'api_addCustomerLog', 'api_deleteCustomerLog'];

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    if (ANDROID_API_.indexOf(req.fn) < 0) throw new Error('درخواست نامعتبر');
    const fns = { api_login: api_login, api_bootstrap: api_bootstrap, api_saveItem: api_saveItem, api_unsaveItem: api_unsaveItem,
      api_getSnapshot: api_getSnapshot, api_addFile: api_addFile, api_deleteFile: api_deleteFile, api_listUsers: api_listUsers,
      api_saveUser: api_saveUser, api_setSignature: api_setSignature, api_removeDuplicates: api_removeDuplicates,
      api_saveCustomer: api_saveCustomer, api_deleteCustomer: api_deleteCustomer, api_addCustomerLog: api_addCustomerLog,
      api_deleteCustomerLog: api_deleteCustomerLog };
    out = { ok: true, result: fns[req.fn].apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
