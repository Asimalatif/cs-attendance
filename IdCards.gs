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
 *   Photos are read 12 at a time in parallel; the page keeps them on the device, so each photo is read from Drive once.
 *   Photos fixed or given in the app go to "JCW ID Cards/Student photos" and the sheet "ID Card Photos".
 *   c) Run idcAuthorize once (Run ▸ Allow), then Deploy > Manage deployments > Edit > New version > Deploy.
 */
var IDC_CFG = {
  USERS: '1ARU8zMO2daJUfU_SIIo4baeSOG7qBPuni5TYRIRLBNY',     // Users workbook
  STRENGTH: 'Transform-ALL-COLLEGE-STRENGTH',
  LOG: 'ID Cards Printed',                                   // made by itself in the Users workbook
  SIGN_FOLDER: 'JCW ID Cards',                               // Drive folder for the principal's sign + stamp of each year
  PROP: 'IDC_SETTINGS',                                      // card settings (Script Properties)
  FIRST_YEAR: /1st/i,                                        // only 1st year students get cards
  PHOTOS: 'ID Card Photos',                                  // sheet: photos the admin gave/fixed in the app (Enroll#, File ID, Updated, By)
  PHOTO_FOLDER: 'Student photos',                            // inside the "JCW ID Cards" folder
  PHOTO_MAX: 1500000,                                        // bytes; bigger photos use Drive's thumbnail
  BATCH: 12                                                  // photos per call (fetched in parallel)
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
      case 'idc_photo_save': return idcSavePhoto_(user, req.roll, req.data);
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
  var printed = idcPrinted_(book), own = idcOwnPhotos_(book), out = [];
  for (var i = 1; i < v.length; i++) {
    var r = v[i], roll = g(r, 'roll');
    if (!roll || !g(r, 'name') || !IDC_CFG.FIRST_YEAR.test(g(r, 'cls'))) continue;
    out.push({ roll: roll, name: g(r, 'name'), father: g(r, 'father'), program: g(r, 'program'), cls: g(r, 'cls'),
      section: g(r, 'section'), pcs: g(r, 'pcs'), group: /^null$/i.test(g(r, 'group')) ? '' : g(r, 'group'),
      pic: g(r, 'pic') || g(r, 'pic2'), status: g(r, 'status'), session: g(r, 'session'),
      printed: printed[idcK_(roll)] || '' });
    var o = own[idcK_(roll)];
    if (o) { var st = out[out.length - 1]; st.pic = 'https://drive.google.com/file/d/' + o.id + '/view'; st.own = o.at; }
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

/** items: [{ roll, pic }] -> { photos: { roll: dataURL | '' } }. Drive files and web links are fetched in parallel. */
function idcPhotos_(items) {
  items = items.slice(0, IDC_CFG.BATCH);
  var out = {}, reqs = [], who = [], tok = ScriptApp.getOAuthToken();
  items.forEach(function (it) {
    var roll = idcS_(it.roll), url = idcS_(it.pic);
    out[roll] = '';
    if (!url || /^(null|n\/?a|-|0)$/i.test(url)) return;
    var id = (url.match(/[-\w]{25,}/) || [])[0];
    if (id && /drive|docs\.google|googleusercontent/i.test(url)) {
      reqs.push({ url: 'https://www.googleapis.com/drive/v3/files/' + id + '?alt=media&supportsAllDrives=true', headers: { Authorization: 'Bearer ' + tok }, muteHttpExceptions: true });
      who.push({ roll: roll, id: id });
    } else if (/^https?:\/\//i.test(url)) {
      reqs.push({ url: url, muteHttpExceptions: true, followRedirects: true });
      who.push({ roll: roll });
    }
  });
  var res = [];
  try { res = reqs.length ? UrlFetchApp.fetchAll(reqs) : []; } catch (e) { res = []; }
  who.forEach(function (w, i) {
    try {
      var r = res[i], blob = r && r.getResponseCode() === 200 ? r.getBlob() : null;
      if (blob && !/^image\//i.test(blob.getContentType() || '') && !w.id) blob = null;
      if (w.id && (!blob || blob.getBytes().length > IDC_CFG.PHOTO_MAX)) {      // big or unreadable: one by one
        var f = DriveApp.getFileById(w.id); blob = f.getBlob();
        if (blob.getBytes().length > IDC_CFG.PHOTO_MAX) blob = f.getThumbnail();
      }
      if (!blob) return;
      var bytes = blob.getBytes();
      if (!bytes.length || bytes.length > IDC_CFG.PHOTO_MAX) return;
      var type = /^image\//i.test(blob.getContentType() || '') ? blob.getContentType() : 'image/jpeg';
      out[w.roll] = 'data:' + type + ';base64,' + Utilities.base64Encode(bytes);
    } catch (e) { out[w.roll] = ''; }
  });
  return { ok: true, photos: out };
}

/* ================= photos given or fixed in the app ================= */
function idcOwnPhotos_(book) {
  var sh = book.getSheetByName(IDC_CFG.PHOTOS), map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).getDisplayValues().forEach(function (r) { if (r[0] && r[1]) map[idcK_(r[0])] = { id: r[1], at: r[2] }; });
  return map;
}
function idcPhotoFolder_() {
  var root = idcFolder_(), it = root.getFoldersByName(IDC_CFG.PHOTO_FOLDER);
  return it.hasNext() ? it.next() : root.createFolder(IDC_CFG.PHOTO_FOLDER);
}
/** saves the (rotated / trimmed / new) photo of one student: Drive file + "ID Card Photos" row; fills Pic URL only if it is empty */
function idcSavePhoto_(user, roll, data) {
  roll = idcS_(roll);
  var m = String(data || '').match(/^data:image\/(jpeg|png);base64,(.+)$/);
  if (!roll || !m) throw new Error('No photo received.');
  var bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > 2000000) throw new Error('Photo too big.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var book = SpreadsheetApp.openById(IDC_CFG.USERS), folder = idcPhotoFolder_();
    var name = roll.replace(/[^\w-]+/g, '_') + '.jpg', old = folder.getFilesByName(name);
    while (old.hasNext()) old.next().setTrashed(true);
    var file = folder.createFile(Utilities.newBlob(bytes, 'image/' + m[1], name));
    var sh = book.getSheetByName(IDC_CFG.PHOTOS);
    if (!sh) { sh = book.insertSheet(IDC_CFG.PHOTOS); sh.getRange(1, 1, 1, 4).setValues([['Enroll#', 'File ID', 'Updated', 'By']]).setFontWeight('bold'); sh.setFrozenRows(1); }
    var now = Utilities.formatDate(new Date(), 'Asia/Karachi', 'dd-MM-yyyy HH:mm');
    var who = idcS_(user && (user.username || user.name || user.email) || user);
    var n = sh.getLastRow(), rows = n > 1 ? sh.getRange(2, 1, n - 1, 1).getDisplayValues() : [], at = -1;
    for (var i = 0; i < rows.length; i++) if (idcK_(rows[i][0]) === idcK_(roll)) { at = i + 2; break; }
    var row = [[roll, file.getId(), now, who]];
    if (at > 0) sh.getRange(at, 1, 1, 4).setNumberFormat('@').setValues(row);
    else sh.getRange(n + 1, 1, 1, 4).setNumberFormat('@').setValues(row);
    // a student who had no photo gets this one in the strength sheet too (existing photo links are never changed)
    var st = book.getSheetByName(IDC_CFG.STRENGTH), v = st.getDataRange().getDisplayValues(), head = v[0].map(idcK_);
    var cR = head.indexOf('ENROLL'), cP = head.indexOf('PICURL');
    if (cR >= 0 && cP >= 0) for (var r = 1; r < v.length; r++) if (idcK_(v[r][cR]) === idcK_(roll)) {
      if (!idcS_(v[r][cP]) || /^null$/i.test(v[r][cP])) st.getRange(r + 1, cP + 1).setValue('https://drive.google.com/file/d/' + file.getId() + '/view');
      break;
    }
    return { ok: true, id: file.getId(), at: now };
  } finally { lock.releaseLock(); }
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
