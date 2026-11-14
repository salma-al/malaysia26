/**
 * Malaysia 26 · shared trip data
 * Stores eating-list ticks, itinerary "done" ticks, day notes, the Malacca and travel
 * choices, the shopping list and the photo album (photos go to Google Drive folders) in this Google Sheet (shopping-list photos go to a Google Drive folder), and reads the photo list from your shared Google Photos album.
 *
 * ---------------------------------------------------------------------------
 * SETUP (about 5 minutes, free)
 *  1. Create a new Google Sheet (sheets.new). Name it "Malaysia 26".
 *  2. In the sheet: Extensions > Apps Script. Delete what's there and paste this file.
 *  3. Fill in the settings just below, then click Save (the disk icon).
 *  4. Deploy > New deployment > gear icon > "Web app".
 *       Execute as: Me      Who has access: Anyone
 *     Click Deploy > Authorize access > pick your Google account.
 *     If Google says "Google hasn't verified this app": Advanced > Go to project (unsafe).
 *     It's your own script, so that's expected.
 *  5. Copy the Web app URL (ends in /exec) and send it to Claude / paste it into
 *     index.html as SYNC_URL.
 *
 * Changing a setting later? Edit it, Save, then Deploy > Manage deployments >
 * pencil icon > Version: "New version" > Deploy. (The URL stays the same.)
 * ---------------------------------------------------------------------------
 */

// 1. A secret code you both type once on each device. Not stored in the public site.
var TRIP_CODE = 'change-me-to-something-secret';

// 2. Your two names, exactly as you want them shown (used for "paid by" and notes).
var PEOPLE = ['Salma', 'Rahma'];

// 3. Your booking PDFs: upload them to Google Drive (keep them private, don't share them),
//    then paste each file's Drive link. The site can only open them with the trip code.
var BOOKINGS = {
  langkawi: '',  // Expedia booking (Coconut Beach Villa)
  penang: ''     // Trip.com check-in voucher (22 Macalisterz)
};

// 4. Your shared Google Photos album link (Share > Create link). Leave '' for now if you don't have it yet.
var ALBUM_URL = '';

/* ======================= nothing to change below ======================= */

var LIMITS = { ticks: 1500, spend: 2000, shop: 500, note: 4000, what: 80, imgBytes: 3000000 };

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!codeOk_(p.code)) return reply_({ error: 'code' });
  if (p.what === 'photos') return reply_(photos_());
  if (p.what === 'img') return reply_(img_(p.id));
  if (p.what === 'booking') return reply_(booking_(p.id));
  return reply_(state_());
}

function doPost(e) {
  var b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return reply_({ error: 'bad request' }); }
  if (!codeOk_(b.code)) return reply_({ error: 'code' });
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var by = PEOPLE.indexOf(b.by) >= 0 ? b.by : '';
    if (b.a === 'tick') return reply_(tick_(b.key, !!b.on, by));
    if (b.a === 'note') return reply_(note_(b.key, b.text, by));
    if (b.a === 'spend.add') return reply_(spendAdd_(b.item));
    if (b.a === 'spend.del') return reply_(spendDel_(b.id));
    if (b.a === 'set') return reply_(set_(b.key, b.val, by));
    if (b.a === 'album.add') return reply_(albumAdd_(b.img, by));
    if (b.a === 'album.del') return reply_(albumDel_(b.id));
    if (b.a === 'shop.add') return reply_(shopAdd_(b.item, by));
    if (b.a === 'shop.edit') return reply_(shopEdit_(b.id, b.text, b.link, !!b.rmimg, b.img));
    if (b.a === 'shop.toggle') return reply_(shopToggle_(b.id, !!b.on));
    if (b.a === 'shop.del') return reply_(shopDel_(b.id));
    return reply_({ error: 'unknown action' });
  } catch (err) {
    return reply_({ error: 'failed' });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- storage ---------- */
function sheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sh.getRange('A:Z').setNumberFormat('@'); // keep everything as plain text
  }
  return sh;
}
function ticksSheet_() { return sheet_('ticks', ['key', 'by', 'updated']); }
function notesSheet_() { return sheet_('notes', ['key', 'text', 'by', 'updated']); }
function spendSheet_() { return sheet_('spending', ['id', 'date', 'paid by', 'what', 'amount', 'currency', 'for']); }
function settingsSheet_() { return sheet_('settings', ['setting', 'value', 'by', 'updated']); }
function shopSheet_() { return sheet_('shopping', ['id', 'added', 'by', 'item', 'link', 'photo (Drive file id)', 'bought']); }

function rows_(sh) {
  var n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, sh.getLastColumn()).getValues();
}
function rowOf_(sh, key) {
  var r = rows_(sh);
  for (var i = 0; i < r.length; i++) if (String(r[i][0]) === key) return i + 2;
  return 0;
}
function clean_(t, max) {
  t = String(t == null ? '' : t).slice(0, max);
  return /^[=+\-@]/.test(t) ? "'" + t : t; // never let text become a formula
}
function now_() { return new Date().toISOString(); }

function state_() {
  var ticks = {}, notes = {}, spend = [];
  rows_(ticksSheet_()).forEach(function (r) { if (r[0]) ticks[String(r[0])] = 1; });
  rows_(notesSheet_()).forEach(function (r) { if (r[0]) notes[String(r[0])] = { t: String(r[1]), by: String(r[2]) }; });
  rows_(spendSheet_()).forEach(function (r) {
    if (!r[0]) return;
    spend.push({ id: String(r[0]), at: String(r[1]), who: String(r[2]), what: String(r[3]), amt: Number(r[4]) || 0, cur: String(r[5]), split: String(r[6]) });
  });
  var shop = [];
  rows_(shopSheet_()).forEach(function (r) {
    if (!r[0]) return;
    shop.push({ id: String(r[0]), at: String(r[1]), by: String(r[2]), text: String(r[3]), link: String(r[4]), img: String(r[5]), got: String(r[6]) === '1' ? 1 : 0 });
  });
  var settings = {};
  rows_(settingsSheet_()).forEach(function (r) { if (r[0]) settings[String(r[0])] = String(r[1]); });
  return { ticks: ticks, notes: notes, spend: spend, shop: shop, settings: settings, people: PEOPLE };
}

function tick_(key, on, by) {
  if (typeof key !== 'string' || !/^(food|stop):[^<>]{1,140}$/.test(key)) return { error: 'bad key' };
  var sh = ticksSheet_(), row = rowOf_(sh, key);
  if (on && !row) {
    if (sh.getLastRow() > LIMITS.ticks) return { error: 'full' };
    sh.appendRow([clean_(key, 150), by, now_()]);
  } else if (!on && row) {
    sh.deleteRow(row);
  }
  return { ok: true };
}

function note_(key, text, by) {
  if (typeof key !== 'string' || !/^day:[A-Za-z0-9 ]{1,20}$/.test(key)) return { error: 'bad key' };
  var sh = notesSheet_(), row = rowOf_(sh, key);
  text = String(text || '').trim();
  if (!text) { if (row) sh.deleteRow(row); return { ok: true }; }
  var vals = [[key, clean_(text, LIMITS.note), by, now_()]];
  if (row) sh.getRange(row, 1, 1, 4).setValues(vals); else sh.appendRow(vals[0]);
  return { ok: true };
}

function spendAdd_(x) {
  if (!x || typeof x.id !== 'string' || !/^[a-z0-9]{4,30}$/.test(x.id)) return { error: 'bad item' };
  var amt = Number(x.amt);
  if (!(amt > 0 && amt < 1000000)) return { error: 'bad amount' };
  if (['MYR', 'EGP', 'GBP', 'USD'].indexOf(x.cur) < 0) return { error: 'bad currency' };
  if (['both', 'payer', 'other'].indexOf(x.split) < 0) x.split = 'both';
  if (PEOPLE.indexOf(x.who) < 0) return { error: 'bad person' };
  var sh = spendSheet_();
  if (rowOf_(sh, x.id)) return { ok: true }; // already saved
  if (sh.getLastRow() > LIMITS.spend) return { error: 'full' };
  sh.appendRow([x.id, clean_(x.at || now_(), 40), x.who, clean_(x.what, LIMITS.what), String(Math.round(amt * 100) / 100), x.cur, x.split]);
  return { ok: true };
}

function spendDel_(id) {
  if (typeof id !== 'string') return { error: 'bad id' };
  var sh = spendSheet_(), row = rowOf_(sh, id);
  if (row) sh.deleteRow(row);
  return { ok: true };
}

/* ---------- shared settings (e.g. Malacca day trip on/off) ---------- */
var SETTINGS = {
  malday: ['0', '1'],
  tr_klia: ['plane', 'trainferry'],
  tr_lgkpen: ['plane', 'ferry'],
  tr_penkl: ['plane', 'train', 'bus']
};
function set_(key, val, by) {
  if (!SETTINGS.hasOwnProperty(key) || SETTINGS[key].indexOf(String(val)) < 0) return { error: 'bad setting' };
  var sh = settingsSheet_(), row = rowOf_(sh, key);
  var vals = [[key, String(val), by, now_()]];
  if (row) sh.getRange(row, 1, 1, 4).setValues(vals); else sh.appendRow(vals[0]);
  return { ok: true };
}

/* ---------- shopping list ---------- */
function shopAdd_(x, by) {
  if (!x || typeof x.id !== 'string' || !/^[a-z0-9]{4,30}$/.test(x.id)) return { error: 'bad item' };
  var sh = shopSheet_();
  if (rowOf_(sh, x.id)) return { ok: true }; // already saved
  if (sh.getLastRow() > LIMITS.shop) return { error: 'full' };
  var text = String(x.text || '').slice(0, 200);
  var link = String(x.link || '').slice(0, 500);
  if (link && !/^https?:\/\/[^\s<>"]+$/i.test(link)) link = '';
  var fileId = saveImg_(x.id, x.img);
  if (!text && !link && !fileId) return { error: 'empty' };
  sh.appendRow([x.id, now_(), by, clean_(text, 200), clean_(link, 500), fileId, '0']);
  return { ok: true, img: fileId };
}
function saveImg_(id, img) {
  img = String(img || '');
  var m = /^data:image\/(jpeg|png|webp);base64,([A-Za-z0-9+\/=]+)$/.exec(img);
  if (!m || m[2].length >= LIMITS.imgBytes * 1.4) return '';
  var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), 'image/' + m[1], id + '.' + (m[1] === 'jpeg' ? 'jpg' : m[1]));
  var file = shopFolder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getId();
}
function shopEdit_(id, text, link, rmimg, img) {
  var sh = shopSheet_(), row = rowOf_(sh, String(id));
  if (!row) return { ok: true };
  text = String(text || '').slice(0, 200);
  link = String(link || '').slice(0, 500);
  if (link && !/^https?:\/\/[^\s<>"]+$/i.test(link)) link = '';
  sh.getRange(row, 4, 1, 2).setValues([[clean_(text, 200), clean_(link, 500)]]);
  var newId = img ? saveImg_(String(id), img) : '';
  if (rmimg || newId) {
    var oldId = String(sh.getRange(row, 6).getValue());
    if (oldId) { try { DriveApp.getFileById(oldId).setTrashed(true); } catch (err) {} }
    sh.getRange(row, 6).setValue(newId);
  }
  return { ok: true };
}
function shopToggle_(id, on) {
  var sh = shopSheet_(), row = rowOf_(sh, String(id));
  if (row) sh.getRange(row, 7).setValue(on ? '1' : '0');
  return { ok: true };
}
function shopDel_(id) {
  var sh = shopSheet_(), row = rowOf_(sh, String(id));
  if (!row) return { ok: true };
  var fileId = String(sh.getRange(row, 6).getValue());
  if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (err) {} } // goes to Drive trash, not deleted forever
  sh.deleteRow(row);
  return { ok: true };
}
function shopFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHOP_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (err) {} }
  var f = DriveApp.createFolder('Malaysia 26 shopping list photos');
  props.setProperty('SHOP_FOLDER', f.getId());
  return f;
}

/* ---------- photos: read the shared album page (unofficial; may break if Google changes it) ---------- */
function photos_() {
  var drive = [];
  var it = albumFolder_().getFiles();
  while (it.hasNext()) {
    var f = it.next();
    if (f.isTrashed() || String(f.getMimeType()).indexOf('image/') !== 0) continue;
    drive.push({ id: f.getId(), by: f.getDescription() || '', at: f.getDateCreated().toISOString() });
  }
  drive.sort(function (a, b) { return a.at < b.at ? 1 : -1; });
  var g = googlePhotos_();
  return { drive: drive, photos: g.photos, url: g.url };
}
function googlePhotos_() {
  if (!ALBUM_URL) return { url: '', photos: [] };
  var cache = CacheService.getScriptCache();
  var hit = cache.get('photos');
  if (hit) return JSON.parse(hit);
  var out = [], seen = {};
  try {
    var html = UrlFetchApp.fetch(ALBUM_URL, { followRedirects: true, muteHttpExceptions: true }).getContentText();
    var re = /\["(https:\/\/lh3\.googleusercontent\.com\/pw\/[^"\\]+)",(\d+),(\d+)/g, m;
    while ((m = re.exec(html)) && out.length < 600) {
      if (seen[m[1]]) continue;
      seen[m[1]] = 1;
      out.push({ u: m[1], w: +m[2], h: +m[3] });
    }
  } catch (err) { /* fall through with what we have */ }
  var res = { url: ALBUM_URL, photos: out };
  try { cache.put('photos', JSON.stringify(res), 600); } catch (err) {}
  return res;
}

/* ---------- booking PDFs (private files in your Drive) ---------- */
function booking_(key) {
  if (!BOOKINGS.hasOwnProperty(key) || !BOOKINGS[key]) return { error: 'missing' };
  var m = /[-\w]{25,}/.exec(BOOKINGS[key]);
  if (!m) return { error: 'missing' };
  var file = DriveApp.getFileById(m[0]);
  var bytes = file.getBlob().getBytes();
  if (bytes.length > 10000000) return { error: 'too big' };
  return { name: file.getName(), data: Utilities.base64Encode(bytes) };
}

/* ---------- album uploads (Drive folder "Malaysia 26 album") ---------- */
function albumFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('ALBUM_FOLDER');
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (err) {} }
  var folder = DriveApp.createFolder('Malaysia 26 album');
  props.setProperty('ALBUM_FOLDER', folder.getId());
  return folder;
}
function albumAdd_(img, by) {
  var m = /^data:image\/(jpeg|png|webp);base64,([A-Za-z0-9+\/=]+)$/.exec(String(img || ''));
  if (!m) return { error: 'bad image' };
  if (m[2].length > 8000000) return { error: 'too big' };
  var name = new Date().toISOString().replace(/[:.]/g, '-') + '.' + (m[1] === 'jpeg' ? 'jpg' : m[1]);
  var file = albumFolder_().createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), 'image/' + m[1], name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  if (by) file.setDescription(by);
  return { ok: true, id: file.getId() };
}
function inFolder_(file, folder) {
  var ps = file.getParents();
  while (ps.hasNext()) if (ps.next().getId() === folder.getId()) return true;
  return false;
}
function albumDel_(id) {
  try {
    var f = DriveApp.getFileById(String(id));
    if (inFolder_(f, albumFolder_())) f.setTrashed(true); // goes to Drive trash
  } catch (err) {}
  return { ok: true };
}
// Serves a photo's bytes when Google's direct image links don't load (only album and shopping-list files).
function img_(id) {
  try {
    var f = DriveApp.getFileById(String(id));
    if (!inFolder_(f, albumFolder_()) && !inFolder_(f, shopFolder_())) return { error: 'not allowed' };
    var b = f.getBlob();
    if (b.getBytes().length > 6000000) return { error: 'too big' };
    return { mime: b.getContentType(), data: Utilities.base64Encode(b.getBytes()) };
  } catch (err) { return { error: 'missing' }; }
}

/* ---------- helpers ---------- */
function codeOk_(c) { return typeof c === 'string' && c.length > 0 && c === TRIP_CODE && TRIP_CODE !== 'change-me-to-something-secret'; }
function reply_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
