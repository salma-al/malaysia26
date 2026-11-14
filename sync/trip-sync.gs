/**
 * Malaysia 26 · shared trip data
 * Stores eating-list ticks, itinerary "done" ticks, day notes, the Malacca and travel
 * choices, the shopping list and your bookings list in this Google Sheet. Album and
 * shopping-list photos, and booking attachments (kept private), are saved in Google Drive folders.
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


/* ======================= nothing to change below ======================= */

var LIMITS = { ticks: 1500, spend: 2000, shop: 500, note: 4000, what: 80, imgBytes: 3000000 };

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!codeOk_(p.code)) return reply_({ error: 'code' });
  if (p.what === 'photos') return reply_(photos_());
  if (p.what === 'img') return reply_(img_(p.id));
  if (p.what === 'booking') return reply_(booking_(p.id));
  if (p.what === 'bkfile') return reply_(bkFile_(p.id));
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
    if (b.a === 'album.add') return reply_(albumAdd_(b.img, by, b.taken));
    if (b.a === 'album.del') return reply_(albumDel_(b.id));
    if (b.a === 'shop.add') return reply_(shopAdd_(b.item, by));
    if (b.a === 'shop.edit') return reply_(shopEdit_(b.id, b.text, b.link, !!b.rmimg, b.img));
    if (b.a === 'shop.toggle') return reply_(shopToggle_(b.id, !!b.on));
    if (b.a === 'shop.del') return reply_(shopDel_(b.id));
    if (b.a === 'bk.add') return reply_(bkAdd_(b.item, b.file, by));
    if (b.a === 'bk.edit') return reply_(bkEdit_(b.id, b.title, !!b.rmfile, b.file));
    if (b.a === 'bk.del') return reply_(bkDel_(b.id));
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
function bkSheet_() { return sheet_('bookings', ['id', 'added', 'by', 'title', 'file (Drive file id)', 'file name', 'type']); }
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
  var bookings = [];
  rows_(bkSheet_()).forEach(function (r) {
    if (!r[0]) return;
    bookings.push({ id: String(r[0]), at: String(r[1]), by: String(r[2]), title: String(r[3]), file: String(r[4]), fname: String(r[5]), mime: String(r[6]) });
  });
  return { ticks: ticks, notes: notes, spend: spend, shop: shop, bookings: bookings, settings: settings, people: PEOPLE };
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
/* ---------- our bookings (title + optional photo/PDF, kept private in Drive) ---------- */
var BK_TYPES = { 'application/pdf': 'pdf', 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp', 'image/heic': 'heic' };
function bkFolder_() { return namedFolder_('BK_FOLDER', 'Malaysia 26 bookings'); }
function bkSave_(id, f) {
  if (!f || !BK_TYPES[f.mime] || typeof f.data !== 'string' || !/^[A-Za-z0-9+\/=]+$/.test(f.data)) return null;
  if (f.data.length > 14000000) return null; // about 10 MB
  var name = clean_(String(f.name || ('booking.' + BK_TYPES[f.mime])), 120);
  var file = bkFolder_().createFile(Utilities.newBlob(Utilities.base64Decode(f.data), f.mime, name)); // private: only served through this script
  return { id: file.getId(), name: name, mime: f.mime };
}
function bkAdd_(x, f, by) {
  if (!x || typeof x.id !== 'string' || !/^[a-z0-9]{4,30}$/.test(x.id)) return { error: 'bad item' };
  var title = String(x.title || '').trim().slice(0, 120);
  if (!title) return { error: 'empty' };
  var sh = bkSheet_();
  if (rowOf_(sh, x.id)) return { ok: true };
  if (sh.getLastRow() > 300) return { error: 'full' };
  var saved = f ? bkSave_(x.id, f) : null;
  if (f && !saved) return { error: 'bad file' };
  sh.appendRow([x.id, now_(), by, clean_(title, 120), saved ? saved.id : '', saved ? saved.name : '', saved ? saved.mime : '']);
  return { ok: true };
}
function bkEdit_(id, title, rmfile, f) {
  var sh = bkSheet_(), row = rowOf_(sh, String(id));
  if (!row) return { error: 'missing' };
  title = String(title || '').trim().slice(0, 120);
  if (title) sh.getRange(row, 4).setValue(clean_(title, 120));
  var saved = f ? bkSave_(String(id), f) : null;
  if (f && !saved) return { error: 'bad file' };
  if (rmfile || saved) {
    var oldId = String(sh.getRange(row, 5).getValue());
    if (oldId) { try { DriveApp.getFileById(oldId).setTrashed(true); } catch (err) {} }
    sh.getRange(row, 5, 1, 3).setValues([[saved ? saved.id : '', saved ? saved.name : '', saved ? saved.mime : '']]);
  }
  return { ok: true };
}
function bkDel_(id) {
  var sh = bkSheet_(), row = rowOf_(sh, String(id));
  if (!row) return { ok: true };
  var fileId = String(sh.getRange(row, 5).getValue());
  if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (err) {} } // Drive trash, not deleted forever
  sh.deleteRow(row);
  return { ok: true };
}
function bkFile_(id) {
  var sh = bkSheet_(), ok = rows_(sh).some(function (r) { return String(r[4]) === String(id) && id; });
  if (!ok) return { error: 'not allowed' };
  var f = DriveApp.getFileById(String(id)), bytes = f.getBlob().getBytes();
  if (bytes.length > 10000000) return { error: 'too big' };
  return { name: f.getName(), mime: f.getMimeType(), data: Utilities.base64Encode(bytes) };
}

function shopFolder_() { return namedFolder_('SHOP_FOLDER', 'Malaysia 26 shopping list photos'); }

function photos_() {
  var drive = [], seen = {};
  albumFolder_();
  var folders = albumFolders_(), files = [];
  folders.forEach(function (fo) { var fi = fo.getFiles(); while (fi.hasNext()) files.push(fi.next()); });
  files.forEach(function (f) {
    if (seen[f.getId()]) return; seen[f.getId()] = 1;
    if (f.isTrashed() || String(f.getMimeType()).indexOf('image/') !== 0) return;
    var meta = {}, desc = f.getDescription() || '';
    try { meta = JSON.parse(desc); } catch (err) { meta = { by: desc }; }
    drive.push({ id: f.getId(), by: meta.by || '', taken: meta.taken || '', at: f.getDateCreated().toISOString() });
  });
  drive.sort(function (a, b) { return (a.taken || a.at) < (b.taken || b.at) ? 1 : -1; });
  return { drive: drive };
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
function namedFolder_(prop, name) {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(prop);
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (err) {} }
  var it = DriveApp.getFoldersByName(name), folder = null;
  while (it.hasNext()) { var c = it.next(); if (!c.isTrashed()) { folder = c; break; } }
  if (!folder) folder = DriveApp.createFolder(name);
  props.setProperty(prop, folder.getId());
  return folder;
}
function albumFolder_() { return namedFolder_('ALBUM_FOLDER', 'Malaysia 26 album'); }
function albumFolders_() {
  var out = [], it = DriveApp.getFoldersByName('Malaysia 26 album');
  while (it.hasNext()) { var f = it.next(); if (!f.isTrashed()) out.push(f); }
  return out;
}
function inAlbum_(file) { return albumFolders_().some(function (fo) { return inFolder_(file, fo); }); }
function albumAdd_(img, by, taken) {
  var m = /^data:image\/(jpeg|png|webp);base64,([A-Za-z0-9+\/=]+)$/.exec(String(img || ''));
  if (!m) return { error: 'bad image' };
  if (m[2].length > 8000000) return { error: 'too big' };
  var name = new Date().toISOString().replace(/[:.]/g, '-') + '.' + (m[1] === 'jpeg' ? 'jpg' : m[1]);
  var file = albumFolder_().createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), 'image/' + m[1], name));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  taken = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(String(taken || '')) ? String(taken) : '';
  file.setDescription(JSON.stringify({ by: by || '', taken: taken }));
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
    if (!inAlbum_(f)) return { ok: true, deleted: false };
    f.setTrashed(true); // goes to Drive trash, so it can be restored
    return { ok: true, deleted: true };
  } catch (err) { return { ok: true, deleted: false }; }
}
// Serves a photo's bytes when Google's direct image links don't load (only album and shopping-list files).
function img_(id) {
  try {
    var f = DriveApp.getFileById(String(id));
    if (!inAlbum_(f) && !inFolder_(f, shopFolder_())) return { error: 'not allowed' };
    var b = f.getBlob();
    if (b.getBytes().length > 6000000) return { error: 'too big' };
    return { mime: b.getContentType(), data: Utilities.base64Encode(b.getBytes()) };
  } catch (err) { return { error: 'missing' }; }
}

/* ---------- helpers ---------- */
function codeOk_(c) { return typeof c === 'string' && c.length > 0 && c === TRIP_CODE && TRIP_CODE !== 'change-me-to-something-secret'; }
function reply_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
