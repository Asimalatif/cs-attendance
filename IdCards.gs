/**
 * IdCards.gs — Student identity cards inside the attendance app (admin only).
 *
 * The app page idcards.html draws the cards (same design as the printed "Section E" cards) from
 * Transform-ALL-COLLEGE-STRENGTH. This file only sends the students, their photos, and keeps a
 * "ID Cards Printed" log in the Users workbook so the admin can see who already has a card.
 * Only 1st year students are sent. Each year the admin sets the session / "Valid up to" date and uploads the
 * principal's signature + stamp (saved in the Drive folder "JCW ID Cards", one picture per session year).
 *
 * INSTALL:
 *   a) + > Script > name it "IdCards" > paste this file.
 *   b) doPost, next to the pc_ line:
 *        if (/^idc_/.test(String(req.action || ''))) return json_(idcRoute_(user, req));
 *   c) Run idcAuthorize once (Run ▸ Allow), then Deploy > Manage deployments > Edit > New version > Deploy.
 */
var IDC_CFG = {
  USERS: '1ARU8zMO2daJUfU_SIIo4baeSOG7qBPuni5TYRIRLBNY',     // Users workbook
  STRENGTH: 'Transform-ALL-COLLEGE-STRENGTH',
  LOG: 'ID Cards Printed',                                   // made by itself in the Users workbook
  SIGN_FOLDER: 'JCW ID Cards',                               // Drive folder for the principal's sign + stamp of each year
  PROP: 'IDC_SETTINGS',                                      // card settings (Script Properties)
  FIRST_YEAR: /1st/i,                                        // only 1st year students get cards
  PHOTO_MAX: 1500000,                                        // bytes; bigger photos use Drive's thumbnail
  BATCH: 8                                                   // photos per call
};

function idcAuthorize() {
  SpreadsheetApp.openById(IDC_CFG.USERS).getName();
  idcFolder_().getName();                                     // makes the "JCW ID Cards" folder (needs full Drive permission)
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
  Logger.log('ID cards: permissions OK. Now Deploy → Manage deployments → Edit → New version.');
}

function idcRoute_(user, req) {
  try {
    if (!isAdmin_(user)) return { ok: false, error: 'Only the admin can make ID cards.' };
    switch (String(req.action || '')) {
      case 'idc_list':   return idcList_();
      case 'idc_photos': return idcPhotos_(req.items || []);
      case 'idc_mark':   return idcMark_(user, req.rolls || []);
      case 'idc_settings': return idcSaveSettings_(req.cfg || {});
      case 'idc_sign':   return idcSaveSign_(req.year, req.data);
    }
    return { ok: false, error: 'UNKNOWN_ACTION' };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}

function idcS_(v) { return String(v == null ? '' : v).trim(); }
function idcK_(v) { return idcS_(v).toUpperCase().replace(/[^A-Z0-9]/g, ''); }

/** every student of the strength sheet with the fields a card needs */
function idcList_() {
  var book = SpreadsheetApp.openById(IDC_CFG.USERS);
  var sh = book.getSheetByName(IDC_CFG.STRENGTH);
  if (!sh) throw new Error('Sheet ' + IDC_CFG.STRENGTH + ' not found.');
  var v = sh.getDataRange().getDisplayValues(), head = v[0].map(idcK_);
  var c = function () {
    for (var i = 0; i < arguments.length; i++) { var j = head.indexOf(arguments[i]); if (j >= 0) return j; }
    return -1;
  };
  var C = { roll: c('ENROLL', 'REG'), name: c('STUDENTNAME'), father: c('FATHERNAME'), program: c('PROGRAM'),
    cls: c('CLASS'), section: c('SECTION'), pcs: c('PCS'), group: c('GROUPONLY'), pic: c('PICURL', 'PHOTO'),
    pic2: c('PHOTO'), status: c('STATUS'), session: c('SESSION'), dob: c('DOB'), contact: c('STUDENTCONTACT') };
  if (C.roll < 0 || C.name < 0) throw new Error('Enroll# or Student Name column not found in ' + IDC_CFG.STRENGTH + '.');
  var g = function (r, k) { return C[k] >= 0 ? idcS_(r[C[k]]) : ''; };
  var printed = idcPrinted_(book), out = [];
  for (var i = 1; i < v.length; i++) {
    var r = v[i], roll = g(r, 'roll');
    if (!roll || !g(r, 'name') || !IDC_CFG.FIRST_YEAR.test(g(r, 'cls'))) continue;
    out.push({ roll: roll, name: g(r, 'name'), father: g(r, 'father'), program: g(r, 'program'), cls: g(r, 'cls'),
      section: g(r, 'section'), pcs: g(r, 'pcs'), group: /^null$/i.test(g(r, 'group')) ? '' : g(r, 'group'),
      pic: g(r, 'pic') || g(r, 'pic2'), status: g(r, 'status'), session: g(r, 'session'),
      printed: printed[idcK_(roll)] || '' });
  }
  var cfg = idcSettings_(), signs = idcSigns_();
  return { ok: true, list: out, at: new Date().toISOString(), cfg: cfg, signs: signs };
}

/* ================= yearly settings + principal sign / stamp ================= */
function idcSettings_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(IDC_CFG.PROP) || '{}'); } catch (e) { return {}; }
}
function idcSaveSettings_(cfg) {
  var s = JSON.stringify(cfg || {});
  if (s.length > 8500) throw new Error('Settings are too long.');
  PropertiesService.getScriptProperties().setProperty(IDC_CFG.PROP, s);
  return { ok: true };
}
function idcFolder_() {
  var it = DriveApp.getFoldersByName(IDC_CFG.SIGN_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(IDC_CFG.SIGN_FOLDER);
}
function idcSignName_(year) { return 'Principal sign and stamp ' + year + '.png'; }
/** { '2026': dataURL, ... } for every year that has a picture */
function idcSigns_() {
  var out = {}, it;
  try { it = DriveApp.getFoldersByName(IDC_CFG.SIGN_FOLDER); } catch (e) { return out; }
  if (!it.hasNext()) return out;                             // no picture uploaded yet: the folder is made on the first upload
  var files = it.next().getFiles();
  while (files.hasNext()) {
    var f = files.next(), m = f.getName().match(/^Principal sign and stamp (\d{4})\./);
    if (!m) continue;
    var b = f.getBlob();
    out[m[1]] = 'data:' + (b.getContentType() || 'image/png') + ';base64,' + Utilities.base64Encode(b.getBytes());
  }
  return out;
}
function idcSaveSign_(year, data) {
  year = idcS_(year);
  if (!/^\d{4}$/.test(year)) throw new Error('Session year missing.');
  var m = String(data || '').match(/^data:(image\/[\w+.-]+);base64,(.+)$/);
  if (!m) throw new Error('No picture received.');
  var folder = idcFolder_(), name = idcSignName_(year), old = folder.getFilesByName(name);
  while (old.hasNext()) old.next().setTrashed(true);
  folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), 'image/png', name));
  return { ok: true, year: year };
}

/** items: [{ roll, pic }] -> { photos: { roll: dataURL | '' } } */
function idcPhotos_(items) {
  var out = {};
  items.slice(0, IDC_CFG.BATCH).forEach(function (it) {
    var roll = idcS_(it.roll);
    try { out[roll] = idcPhoto_(idcS_(it.pic)); } catch (e) { out[roll] = ''; }
  });
  return { ok: true, photos: out };
}
function idcPhoto_(url) {
  if (!url || /^(null|n\/?a|-|0)$/i.test(url)) return '';
  var blob = null, id = (url.match(/[-\w]{25,}/) || [])[0];
  if (id && /drive|docs\.google|googleusercontent/i.test(url)) {
    var f = DriveApp.getFileById(id);
    blob = f.getBlob();
    if (blob.getBytes().length > IDC_CFG.PHOTO_MAX) blob = f.getThumbnail();
  } else if (/^https?:\/\//i.test(url)) {
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    if (res.getResponseCode() !== 200) return '';
    blob = res.getBlob();
    if (!/^image\//i.test(blob.getContentType() || '')) return '';
  }
  if (!blob) return '';
  var bytes = blob.getBytes();
  if (!bytes.length || bytes.length > IDC_CFG.PHOTO_MAX) return '';
  return 'data:' + (blob.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(bytes);
}

/* ================= printed log ================= */
function idcLogSheet_(book) {
  var sh = book.getSheetByName(IDC_CFG.LOG);
  if (!sh) {
    sh = book.insertSheet(IDC_CFG.LOG);
    sh.getRange(1, 1, 1, 4).setValues([['Enroll#', 'Last printed', 'Times', 'By']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function idcPrinted_(book) {
  var sh = book.getSheetByName(IDC_CFG.LOG), map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues().forEach(function (r) { if (r[0]) map[idcK_(r[0])] = r[1]; });
  return map;
}
function idcMark_(user, rolls) {
  if (!rolls.length) return { ok: true, n: 0 };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var book = SpreadsheetApp.openById(IDC_CFG.USERS), sh = idcLogSheet_(book);
    var n = sh.getLastRow(), rows = n > 1 ? sh.getRange(2, 1, n - 1, 4).getValues() : [], at = {};
    rows.forEach(function (r, i) { at[idcK_(r[0])] = i; });
    var now = Utilities.formatDate(new Date(), 'Asia/Karachi', 'dd-MM-yyyy HH:mm');
    var who = idcS_(user && (user.username || user.name || user.email) || user);
    rolls.forEach(function (roll) {
      var k = idcK_(roll); if (!k) return;
      if (at[k] != null) { var r = rows[at[k]]; r[1] = now; r[2] = (Number(r[2]) || 0) + 1; r[3] = who; }
      else { at[k] = rows.length; rows.push([idcS_(roll), now, 1, who]); }
    });
    if (rows.length) sh.getRange(2, 1, rows.length, 4).setNumberFormat('@').setValues(rows);
    return { ok: true, n: rolls.length, at: now };
  } finally { lock.releaseLock(); }
}
