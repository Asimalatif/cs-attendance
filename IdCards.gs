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
 *   Students check their card on idcheck.html (no login). For that, put this line at the TOP of doPost,
 *   right after `req` is read and BEFORE the login check:
 *        if (/^idp_/.test(String(req.action || ''))) return json_(idcPublic_(req));
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
  BATCH: 12,                                                 // photos per call (fetched in parallel)
  REQS: 'ID Card Requests',                                  // sheet: what students sent from idcheck.html (made by itself)
  REQ_FOLDER: 'Student requests',                            // inside "JCW ID Cards": photos students sent, waiting for the admin
  CHECK_PROP: 'IDC_CHECK',                                   // { open: true/false, until: 'yyyy-mm-dd' } for the student page
  TRIES: 8                                                   // wrong checks allowed per Enroll# in 30 minutes
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
      case 'idc_reqs':   return idcReqList_(req.status || 'Pending', Number(req.from) || 0, Math.min(Number(req.n) || 20, 30));
      case 'idc_req_do': return idcReqDo_(user, req.items || []);
      case 'idc_check_set': PropertiesService.getScriptProperties().setProperty(IDC_CFG.CHECK_PROP, JSON.stringify(req.check || {})); return { ok: true };
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
  var chk = idcReqLatest_(book);
  out.forEach(function (st) { var q = chk[idcK_(st.roll)]; if (q) { st.chk = q.status; st.chkAt = q.when; st.chkAsk = q.asked; st.chkBy = q.by; st.chkOn = q.at; st.chkWhy = q.reason; } });
  // v7.5: left / migrated / cancelled students still in Transform get a warning on the card
  var gone = {};
  try {
    var ls = book.getSheetByName('Left Students');
    if (ls && ls.getLastRow() > 1) {
      var lv = ls.getRange(1, 1, ls.getLastRow(), ls.getLastColumn()).getDisplayValues(), lh = lv[0].map(idcK_), le = -1;
      for (var j = 0; j < lh.length; j++) if (/^ENROLL/.test(lh[j])) { le = j; break; }
      if (le >= 0) for (var k = 1; k < lv.length; k++) if (lv[k][le]) gone[idcK_(lv[k][le])] = idcS_(lv[k][2]) || 'Left';
    }
  } catch (e) {}
  out.forEach(function (st) {
    var s = String(st.status || '').trim();
    if (s && !/^(enrolled|active|regular|admitted)$/i.test(s)) st.gone = s;
    else if (gone[idcK_(st.roll)]) st.gone = gone[idcK_(st.roll)];
  });
  var cfg = idcSettings_(), signs = idcSigns_();
  return { ok: true, list: out, at: new Date().toISOString(), cfg: cfg, signs: signs, check: idcCheck_() };
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
  try { return idcStorePhoto_(SpreadsheetApp.openById(IDC_CFG.USERS), roll, bytes, 'image/' + m[1], idcWho_(user)); }
  finally { lock.releaseLock(); }
}
function idcWho_(user) { return idcS_(user && (user.username || user.name || user.email) || user); }
/** the student's card photo from now on: Drive file + "ID Card Photos" row; Pic URL filled only if empty. Call inside a lock. */
function idcStorePhoto_(book, roll, bytes, mime, who) {
  var folder = idcPhotoFolder_();
  var name = roll.replace(/[^\w-]+/g, '_') + '.jpg', old = folder.getFilesByName(name);
  while (old.hasNext()) old.next().setTrashed(true);
  var file = folder.createFile(Utilities.newBlob(bytes, mime || 'image/jpeg', name));
  var sh = book.getSheetByName(IDC_CFG.PHOTOS);
  if (!sh) { sh = book.insertSheet(IDC_CFG.PHOTOS); sh.getRange(1, 1, 1, 4).setValues([['Enroll#', 'File ID', 'Updated', 'By']]).setFontWeight('bold'); sh.setFrozenRows(1); }
  var now = Utilities.formatDate(new Date(), 'Asia/Karachi', 'dd-MM-yyyy HH:mm');
  var n = sh.getLastRow(), rows = n > 1 ? sh.getRange(2, 1, n - 1, 1).getDisplayValues() : [], at = -1;
  for (var i = 0; i < rows.length; i++) if (idcK_(rows[i][0]) === idcK_(roll)) { at = i + 2; break; }
  var row = [[roll, file.getId(), now, who]];
  if (at > 0) sh.getRange(at, 1, 1, 4).setNumberFormat('@').setValues(row);
  else sh.getRange(n + 1, 1, 1, 4).setNumberFormat('@').setValues(row);
  idcSetStrength_(book, roll, { PICURL: 'https://drive.google.com/file/d/' + file.getId() + '/view' }, true);
  idcIndexDrop_();
  return { ok: true, id: file.getId(), at: now };
}
/** writes columns (by header key, e.g. STUDENTNAME) of one student in the strength sheet; onlyEmpty = never overwrite */
function idcSetStrength_(book, roll, vals, onlyEmpty) {
  var st = book.getSheetByName(IDC_CFG.STRENGTH), v = st.getDataRange().getDisplayValues(), head = v[0].map(idcK_), cR = head.indexOf('ENROLL');
  if (cR < 0) return false;
  for (var r = 1; r < v.length; r++) if (idcK_(v[r][cR]) === idcK_(roll)) {
    Object.keys(vals).forEach(function (k) {
      var c = head.indexOf(k); if (c < 0 || !idcS_(vals[k])) return;
      if (onlyEmpty && idcS_(v[r][c]) && !/^null$/i.test(v[r][c])) return;
      st.getRange(r + 1, c + 1).setValue(vals[k]);
    });
    return true;
  }
  return false;
}

/* ================= students check their own card (idcheck.html, no login) ================= */
function idcCheck_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(IDC_CFG.CHECK_PROP) || '{}'); } catch (e) { return {}; }
}
/** small list of 1st year students kept in the script cache for 30 minutes, so 3000 students do not read the sheet 3000 times */
function idcIndex_() {
  var cache = CacheService.getScriptCache(), n = Number(cache.get('idcidx_n') || 0);
  if (n) {
    var parts = cache.getAll(Array.apply(null, Array(n)).map(function (_, i) { return 'idcidx_' + i; })), txt = '';
    for (var i = 0; i < n; i++) { if (parts['idcidx_' + i] == null) { txt = null; break; } txt += parts['idcidx_' + i]; }
    if (txt) try { return JSON.parse(txt); } catch (e) {}
  }
  var book = SpreadsheetApp.openById(IDC_CFG.USERS), sh = book.getSheetByName(IDC_CFG.STRENGTH);
  var v = sh.getDataRange().getDisplayValues(), head = v[0].map(idcK_);
  var c = function (k) { return head.indexOf(k); };
  var C = { roll: c('ENROLL'), name: c('STUDENTNAME'), father: c('FATHERNAME'), program: c('PROGRAM'), cls: c('CLASS'), section: c('SECTION'), pcs: c('PCS'),
    pic: c('PICURL') >= 0 ? c('PICURL') : c('PHOTO'), contact: c('STUDENTCONTACT') >= 0 ? c('STUDENTCONTACT') : c('CONTACT'), status: c('STATUS') };
  var own = idcOwnPhotos_(book), idx = {};
  for (var r = 1; r < v.length; r++) {
    var row = v[r], g = function (k) { return C[k] >= 0 ? idcS_(row[C[k]]) : ''; }, roll = g('roll');
    if (!roll || !IDC_CFG.FIRST_YEAR.test(g('cls'))) continue;
    var o = own[idcK_(roll)];
    idx[idcK_(roll)] = [roll, g('name'), g('father'), g('program'), g('pcs'), g('section'), o ? 'https://drive.google.com/file/d/' + o.id + '/view' : g('pic'), g('contact').replace(/\D/g, '').slice(-4), g('status')];
  }
  var txt2 = JSON.stringify(idx), size = 90000, put = {}, k = 0;
  for (var p = 0; p < txt2.length; p += size) put['idcidx_' + (k++)] = txt2.substr(p, size);
  put.idcidx_n = String(k);
  try { cache.putAll(put, 1800); } catch (e) {}
  return idx;
}
function idcIndexDrop_() { try { CacheService.getScriptCache().remove('idcidx_n'); } catch (e) {} }
function idcNameK_(s) { return idcS_(s).toUpperCase().replace(/[^A-Z ]+/g, ' ').replace(/\s+/g, ' ').trim(); }
/** Enroll# (or its number) + father's first name or the last 4 digits of the mobile number in the college record */
function idcVerify_(rollIn, key) {
  var idx = idcIndex_(), k = idcK_(rollIn), st = idx[k];
  if (!st) {
    var d = String(rollIn || '').replace(/\D+/g, '');
    if (d.length >= 3) { var hits = Object.keys(idx).filter(function (x) { return idx[x][0].split('-').pop() === d; }); if (hits.length === 1) st = idx[hits[0]]; }
  }
  if (!st) return { error: 'This Enroll# is not in the 1st year list. Write it like JCW-F26-1032.' };
  var cache = CacheService.getScriptCache(), ck = 'idcfail_' + idcK_(st[0]), fails = Number(cache.get(ck) || 0);
  if (fails >= IDC_CFG.TRIES) return { error: 'Too many wrong tries. Wait 30 minutes or ask the college office.' };
  var key2 = idcS_(key), ok = false;
  if (/^\d{4}$/.test(key2)) ok = !!st[7] && st[7] === key2;
  else {
    var a = idcNameK_(key2).split(' ')[0], f = idcNameK_(st[2]).split(' ');
    ok = a.length >= 3 && f.some(function (w) { return w === a || (w.length > 4 && a.length > 4 && w.slice(0, 5) === a.slice(0, 5)); });
  }
  if (!ok) { cache.put(ck, String(fails + 1), 1800); return { error: 'That does not match our record. Check the spelling of your father\'s first name, or give the last 4 digits of the mobile number you gave the college.' }; }
  return { st: st };
}
function idcReqSheet_(book) {
  var sh = book.getSheetByName(IDC_CFG.REQS);
  if (!sh) {
    sh = book.insertSheet(IDC_CFG.REQS);
    sh.getRange(1, 1, 1, 12).setValues([['When', 'Enroll#', 'Status', 'Name on record', 'Father on record', 'Correct name', 'Correct father name', 'Note', 'New photo (file id)', 'Checked by', 'Checked at', 'Reason']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
/** latest request of every student: { ENROLLKEY: { status, when, row } } */
function idcReqLatest_(book) {
  var sh = book.getSheetByName(IDC_CFG.REQS), map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  sh.getRange(2, 1, sh.getLastRow() - 1, 12).getDisplayValues().forEach(function (r, i) { if (r[1]) map[idcK_(r[1])] = { when: r[0], status: r[2], row: i + 2,
    asked: [r[5] ? 'name → ' + r[5] : '', r[6] ? 'father → ' + r[6] : '', r[8] ? 'new photo' : '', r[7] ? 'note: ' + r[7] : ''].filter(String).join('; '), by: r[9], at: r[10], reason: r[11] }; });
  return map;
}
function idcPublic_(req) {
  try {
    var chk = idcCheck_();
    if (!chk.open) return { ok: false, error: 'Checking ID cards is closed at the moment. Ask your class teacher when it opens.' };
    if (chk.until && Utilities.formatDate(new Date(), 'Asia/Karachi', 'yyyy-MM-dd') > chk.until) return { ok: false, error: 'The last date for checking ID cards (' + chk.until + ') has passed. Please visit the college office.' };
    var v = idcVerify_(req.roll, req.key);
    if (v.error) return { ok: false, error: v.error };
    var st = v.st, book;
    if (req.action === 'idp_find') {
      book = SpreadsheetApp.openById(IDC_CFG.USERS);
      var last = idcReqLatest_(book)[idcK_(st[0])] || null, photo = '';
      try { photo = idcPhotos_([{ roll: st[0], pic: st[6] }]).photos[st[0]] || ''; } catch (e) {}
      var cfg = idcSettings_(), sign = '';                       // (the principal's sign is not sent here, to keep this call light)
      return { ok: true, st: { roll: st[0], name: st[1], father: st[2], program: st[3], pcs: st[4], section: st[5] }, photo: photo,
        last: last ? { status: last.status, when: last.when } : null, cfg: cfg, sign: sign, until: chk.until || '' };
    }
    if (req.action === 'idp_send') {
      var name = idcS_(req.name).slice(0, 80), father = idcS_(req.father).slice(0, 80), note = idcS_(req.note).slice(0, 300), fileId = '';
      var allOk = req.allOk === true && !name && !father && !note && !req.photo;
      if (!allOk && !name && !father && !note && !req.photo) return { ok: false, error: 'Nothing to send.' };
      if (req.photo) {
        var pm = String(req.photo).match(/^data:image\/jpeg;base64,(.+)$/);
        if (!pm) return { ok: false, error: 'Photo not received. Try again.' };
        var bytes = Utilities.base64Decode(pm[1]);
        if (bytes.length > 400000) return { ok: false, error: 'Photo too big.' };
        var root = idcFolder_(), it = root.getFoldersByName(IDC_CFG.REQ_FOLDER), fol = it.hasNext() ? it.next() : root.createFolder(IDC_CFG.REQ_FOLDER);
        fileId = fol.createFile(Utilities.newBlob(bytes, 'image/jpeg', st[0].replace(/[^\w-]+/g, '_') + ' ' + Date.now() + '.jpg')).getId();
      }
      var lock = LockService.getScriptLock(); lock.waitLock(25000);
      try {
        book = SpreadsheetApp.openById(IDC_CFG.USERS);
        var sh = idcReqSheet_(book), prev = idcReqLatest_(book)[idcK_(st[0])];
        if (prev && prev.status === 'Pending') sh.getRange(prev.row, 3).setValue('Replaced');
        var now = Utilities.formatDate(new Date(), 'Asia/Karachi', 'dd-MM-yyyy HH:mm');
        sh.appendRow([now, st[0], allOk ? 'Confirmed' : 'Pending', st[1], st[2], name, father, note, fileId, '', '', '']);
        sh.getRange(sh.getLastRow(), 1, 1, 12).setNumberFormat('@');
        return { ok: true, status: allOk ? 'Confirmed' : 'Pending', when: now };
      } finally { lock.releaseLock(); }
    }
    return { ok: false, error: 'UNKNOWN_ACTION' };
  } catch (e) {
    console.error('idcPublic_ ' + req.action + ': ' + (e && e.stack || e));
    return { ok: false, error: 'The server is busy. Please try again in a minute.' };
  }
}

/* ================= admin: students' requests ================= */
function idcReqList_(status, from, n) {
  var book = SpreadsheetApp.openById(IDC_CFG.USERS), sh = book.getSheetByName(IDC_CFG.REQS);
  if (!sh || sh.getLastRow() < 2) return { ok: true, items: [], total: 0 };
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 12).getDisplayValues(), rows = [];
  v.forEach(function (r, i) { if (r[2] === status) rows.push({ row: i + 2, when: r[0], roll: r[1], status: r[2], oldName: r[3], oldFather: r[4], name: r[5], father: r[6], note: r[7], file: r[8], by: r[9], at: r[10], reason: r[11] }); });
  var page = rows.slice(from, from + n), pics = page.filter(function (x) { return x.file; }).map(function (x) { return { roll: x.row + '', pic: 'https://drive.google.com/file/d/' + x.file + '/view' }; });
  var got = {};
  for (var i = 0; i < pics.length; i += IDC_CFG.BATCH) { var p = idcPhotos_(pics.slice(i, i + IDC_CFG.BATCH)).photos; Object.keys(p).forEach(function (k) { got[k] = p[k]; }); }
  page.forEach(function (x) { x.photo = got[x.row + ''] || ''; });
  return { ok: true, items: page, total: rows.length };
}
/** items: [{ row, approve: true|false, reason, photo: true|false, name: true|false, father: true|false }] */
function idcReqDo_(user, items) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var book = SpreadsheetApp.openById(IDC_CFG.USERS), sh = idcReqSheet_(book), who = idcWho_(user), done = 0;
    var now = Utilities.formatDate(new Date(), 'Asia/Karachi', 'dd-MM-yyyy HH:mm');
    items.slice(0, 50).forEach(function (it) {
      var r = sh.getRange(Number(it.row), 1, 1, 12).getDisplayValues()[0];
      if (r[2] !== 'Pending') return;
      if (it.approve) {
        if (r[8] && it.photo !== false) { var f = DriveApp.getFileById(r[8]); idcStorePhoto_(book, r[1], f.getBlob().getBytes(), 'image/jpeg', who); }
        var set = {};
        if (r[5] && it.name !== false) set.STUDENTNAME = r[5].toUpperCase();
        if (r[6] && it.father !== false) set.FATHERNAME = r[6].toUpperCase();
        if (Object.keys(set).length) idcSetStrength_(book, r[1], set, false);
      }
      sh.getRange(Number(it.row), 3).setValue(it.approve ? 'Approved' : 'Rejected');
      sh.getRange(Number(it.row), 10, 1, 3).setValues([[who, now, idcS_(it.reason)]]);
      done++;
    });
    idcIndexDrop_();
    return { ok: true, n: done };
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
