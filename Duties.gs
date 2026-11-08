/**
 * Duties.gs — Exam duty lists (detention, grand test or any exam) on the admin side of the attendance app.
 *
 * The page duties.html reads the duty table your colleague makes in MS Word (open the .docx or paste the table)
 * and makes the teacher-wise, room-wise, date-wise, class-wise and subject-wise lists and PDFs ON THE DEVICE
 * (no server call, also offline). This file is only used to:
 *   - keep the lists: sheet "Duty Lists" in the Users workbook (one row per row of the Word table), so a list
 *     opens on any phone or laptop,
 *   - keep the name fixes: sheet "Duty Names" (Name as written | Standard name) — same idea as the old "names" sheet,
 *   - make a Google Sheet with all the lists in the Drive folder "JCW Duty Lists".
 * Teacher names and WhatsApp numbers come from the existing 'contacts' action.
 *
 * INSTALL:
 *   a) + > Script > name it "Duties" > paste this file.
 *   b) doPost, next to the idc_ line:
 *        if (/^dut_/.test(String(req.action || ''))) return json_(dutRoute_(user, req));
 *   c) Run dutAuthorize once (Run ▸ Allow), then Deploy > Manage deployments > Edit > New version > Deploy.
 */
var DUT_CFG = {
  USERS: '1ARU8zMO2daJUfU_SIIo4baeSOG7qBPuni5TYRIRLBNY',   // Users workbook
  LIST: 'Duty Lists',                                      // made by itself
  NAMES: 'Duty Names',                                     // made by itself
  FIXES: 'Duty Fixes',                                     // made by itself: rooms / classes / subjects / staff notes / room blocks
  STRENGTH: 'Transform-ALL-COLLEGE-STRENGTH',              // students (for the students' seating message)
  FOLDER: 'JCW Duty Lists',                                // Drive folder for the Google Sheets
  MAX_ROWS: 3000                                           // rows in one list (safety)
};
var DUT_HEAD = ['Exam ID', 'Exam', 'Date', 'Day', 'Time', 'Class', 'Subject', 'Room', 'Staff on duty (as written)', 'Teachers', 'Duty timings', 'Saved by', 'Saved at', 'Google Sheet'];

function dutAuthorize() {
  SpreadsheetApp.openById(DUT_CFG.USERS).getName();
  dutFolder_().getName();
  Logger.log('Duty lists: permissions OK. Now Deploy → Manage deployments → Edit → New version.');
}

function dutRoute_(user, req) {
  try {
    if (String(req.action || '') === 'dut_whoami') return { ok: true, admin: !!isAdmin_(user), ds: dutDsAdmin_(user) };
    if (!isAdmin_(user) && !dutDsAdmin_(user)) return { ok: false, error: 'Only the admin or a DS ADMIN can make duty lists.' };
    switch (String(req.action || '')) {
      case 'dut_list':       return dutList_();
      case 'dut_load':       return dutLoad_(dutS_(req.id));
      case 'dut_save':       return dutSave_(user, req.exam || {});
      case 'dut_delete':     return dutDelete_(dutS_(req.id));
      case 'dut_names_save': return dutNamesSave_(req.pairs || []);
      case 'dut_sheet':      return dutMakeSheet_(req);
      case 'dut_fix_save':   return dutFixSave_(req.items || []);
      case 'dut_students':   return dutStudents_();
      case 'dut_staff':      return dutStaff_();
    }
    return { ok: false, error: 'UNKNOWN_ACTION' };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}

/** role DS ADMIN in the Users sheet (column "Role", or a cell reading DS ADMIN in her row) = duty operator */
function dutDsAdmin_(user) {
  var u = String(user && user.username || '').trim().toLowerCase(); if (!u) return false;
  var cache = CacheService.getScriptCache(), ck = 'dutds_' + u, hit = cache.get(ck);
  if (hit) return hit === '1';
  var ok = false;
  try {
    var v = SpreadsheetApp.openById(DUT_CFG.USERS).getSheetByName('Users').getDataRange().getDisplayValues();
    var head = v[0].map(function (x) { return String(x).toUpperCase().replace(/[^A-Z]/g, ''); });
    var cu = head.indexOf('USERNAME'); if (cu < 0) cu = 4;
    var cr = head.indexOf('ROLE');
    for (var i = 1; i < v.length; i++) {
      if (String(v[i][cu]).trim().toLowerCase() !== u) continue;
      var cells = cr >= 0 ? [v[i][cr]] : v[i];
      ok = cells.some(function (x) { return /\bDS\s*-?\s*ADMIN\b/i.test(String(x)); });
      break;
    }
  } catch (e) { ok = false; }
  cache.put(ck, ok ? '1' : '0', 600);
  return ok;
}

function dutS_(v) { return String(v == null ? '' : v).trim(); }

function dutTab_(name, head) {
  var book = SpreadsheetApp.openById(DUT_CFG.USERS), sh = book.getSheetByName(name);
  if (!sh) {
    sh = book.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#6B3423').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange('A:Z').setNumberFormat('@');          // keep dates like 17/02/2026 as text
  }
  return sh;
}

function dutRows_() {
  var sh = dutTab_(DUT_CFG.LIST, DUT_HEAD), n = sh.getLastRow();
  return { sh: sh, rows: n < 2 ? [] : sh.getRange(2, 1, n - 1, DUT_HEAD.length).getDisplayValues() };
}

function dutWrite_(sh, rows) {
  var n = sh.getLastRow();
  if (n > 1) sh.getRange(2, 1, n - 1, DUT_HEAD.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, DUT_HEAD.length).setNumberFormat('@').setValues(rows);
}

function dutPairs_() {
  var sh = dutTab_(DUT_CFG.NAMES, ['Name as written', 'Standard name']), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 2).getDisplayValues().filter(function (r) { return r[0] && r[1]; });
}

function dutList_() {
  var d = dutRows_().rows, by = {}, order = [];
  d.forEach(function (r) {
    var id = r[0]; if (!id) return;
    if (!by[id]) { by[id] = { id: id, title: r[1], count: 0, dates: [], by: r[11], at: r[12] }; order.push(id); }
    by[id].count++;
    if (r[2]) by[id].dates.push(r[2]);
  });
  var key = function (s) { var m = String(s).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); return m ? m[3] + ('0' + m[2]).slice(-2) + ('0' + m[1]).slice(-2) : String(s); };
  var exams = order.map(function (id) {
    var x = by[id], ds = x.dates.slice().sort(function (a, b) { return key(a) < key(b) ? -1 : key(a) > key(b) ? 1 : 0; });
    return { id: id, title: x.title, count: x.count, first: ds[0] || '', last: ds[ds.length - 1] || '', by: x.by, at: x.at };
  }).reverse();                                       // newest first
  return { ok: true, exams: exams, pairs: dutPairs_(), fixes: dutFixes_() };
}

function dutLoad_(id) {
  var d = dutRows_().rows.filter(function (r) { return r[0] === id; });
  if (!d.length) return { ok: false, error: 'This list is not on the server any more.' };
  var sheet = d[0][13] || '';
  return { ok: true, exam: {
    id: id, title: d[0][1], timing: d[0][10], sheet: sheet, sheetUrl: sheet ? 'https://docs.google.com/spreadsheets/d/' + sheet + '/edit' : '',
    rows: d.map(function (r) { return { dt: r[2], day: r[3], time: r[4], cls: r[5], sub: r[6], room: r[7], staff: r[8] }; })
  } };
}

function dutSave_(user, ex) {
  var id = dutS_(ex.id), rows = ex.rows || [];
  if (!id) return { ok: false, error: 'No list id.' };
  if (rows.length > DUT_CFG.MAX_ROWS) return { ok: false, error: 'Too many rows (' + rows.length + ').' };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var R = dutRows_(), old = R.rows.filter(function (r) { return r[0] === id; });
    var sheet = old.length ? old[0][13] : '';
    var at = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'), by = user.teacher || user.username || '';
    var keep = R.rows.filter(function (r) { return r[0] !== id; });
    var add = rows.map(function (r) {
      return [id, dutS_(ex.title), dutS_(r.dt), dutS_(r.day), dutS_(r.time), dutS_(r.cls), dutS_(r.sub), dutS_(r.room), dutS_(r.staff), dutS_(r.teachers), dutS_(ex.timing), by, at, sheet];
    });
    dutWrite_(R.sh, keep.concat(add));
    return { ok: true, count: add.length };
  } finally { lock.releaseLock(); }
}

function dutDelete_(id) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var R = dutRows_();
    dutWrite_(R.sh, R.rows.filter(function (r) { return r[0] !== id; }));
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function dutNamesSave_(pairs) {
  var sh = dutTab_(DUT_CFG.NAMES, ['Name as written', 'Standard name']);
  var rows = (pairs || []).map(function (p) { return [dutS_(p[0]), dutS_(p[1])]; }).filter(function (r) { return r[0] && r[1]; });
  var n = sh.getLastRow();
  if (n > 1) sh.getRange(2, 1, n - 1, 2).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 2).setValues(rows);
  return { ok: true, count: rows.length };
}

function dutFolder_() {
  var it = DriveApp.getFoldersByName(DUT_CFG.FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(DUT_CFG.FOLDER);
}

/** req.tabs = [{ name, rows: [[...]], red: [row numbers to show in red] }] — the page sends the finished tables */
function dutMakeSheet_(req) {
  var ss = null, id = dutS_(req.fileId);
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create(dutS_(req.title) || 'Duty lists');
    DriveApp.getFileById(ss.getId()).moveTo(dutFolder_());
  }
  var first = ss.getSheets()[0], made = {};
  (req.tabs || []).forEach(function (t) {
    var name = dutS_(t.name).replace(/[\\\/?*\[\]:]/g, ' ').slice(0, 99) || 'List';
    var rows = t.rows || []; if (!rows.length) return;
    var w = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 1);
    rows = rows.map(function (r) { r = r.map(function (v) { return v == null ? '' : String(v); }); while (r.length < w) r.push(''); return r; });
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.clear(); made[name] = true;
    sh.getRange(1, 1, rows.length, w).setNumberFormat('@').setValues(rows).setVerticalAlignment('top').setWrap(true);
    sh.getRange(1, 1, 1, w).setFontWeight('bold').setBackground('#6B3423').setFontColor('#ffffff');
    if (rows.length > 1) sh.getRange(2, 1, rows.length - 1, w).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
    (t.red || []).forEach(function (i) { if (i > 0 && i < rows.length) sh.getRange(i + 1, 1, 1, w).setFontColor('#C81E1E').setFontWeight('bold'); });
    sh.setFrozenRows(1);
    for (var c = 1; c <= w; c++) sh.setColumnWidth(c, Math.min(420, Math.max(70, 8 * rows.reduce(function (m, r) { return Math.max(m, String(r[c - 1]).length); }, 4))));
  });
  if (first && !made[first.getName()] && ss.getSheets().length > 1) ss.deleteSheet(first);
  // remember the file on the saved list, so "Update Google Sheet" writes into the same file
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var R = dutRows_(), ch = false;
    R.rows.forEach(function (r) { if (r[0] === dutS_(req.id) && r[13] !== ss.getId()) { r[13] = ss.getId(); ch = true; } });
    if (ch) dutWrite_(R.sh, R.rows);
  } finally { lock.releaseLock(); }
  return { ok: true, fileId: ss.getId(), url: ss.getUrl() };
}

/* ---------- fixes: Type | As written | Standard | (room, class, subject, note, block) ---------- */
var DUT_FIX_HEAD = ['Type', 'As written', 'Standard / note'];
function dutFixes_() {
  var sh = dutTab_(DUT_CFG.FIXES, DUT_FIX_HEAD), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 3).getDisplayValues().filter(function (r) { return r[0] && r[1]; });
}
function dutFixSave_(items) {
  var sh = dutTab_(DUT_CFG.FIXES, DUT_FIX_HEAD);
  var rows = (items || []).map(function (r) { return [dutS_(r[0]), dutS_(r[1]), dutS_(r[2])]; }).filter(function (r) { return r[0] && r[1]; });
  var n = sh.getLastRow();
  if (n > 1) sh.getRange(2, 1, n - 1, 3).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, 3).setValues(rows);
  return { ok: true, count: rows.length };
}

/* ---------- students by class + section, with their Enroll# (for the students' seating message) ---------- */
function dutStudents_() {
  var sh = SpreadsheetApp.openById(DUT_CFG.USERS).getSheetByName(DUT_CFG.STRENGTH);
  if (!sh) throw new Error('Sheet ' + DUT_CFG.STRENGTH + ' not found.');
  var v = sh.getDataRange().getDisplayValues(), K = function (x) { return dutS_(x).toUpperCase().replace(/[^A-Z0-9]/g, ''); };
  var head = v[0].map(K), c = function () { for (var i = 0; i < arguments.length; i++) { var j = head.indexOf(arguments[i]); if (j >= 0) return j; } return -1; };
  var C = { roll: c('ENROLL', 'REG'), cls: c('CLASS'), section: c('SECTION'), program: c('PROGRAM'), group: c('GROUPONLY'), status: c('STATUS') };
  if (C.roll < 0) throw new Error('Enroll# column not found in ' + DUT_CFG.STRENGTH + '.');
  var g = function (r, k) { return C[k] >= 0 ? dutS_(r[C[k]]) : ''; }, by = {};
  for (var i = 1; i < v.length; i++) {
    var r = v[i], roll = g(r, 'roll');
    if (!roll || /migrat|cancel|disenrol|left|struck/i.test(g(r, 'status'))) continue;
    var grp = /^null$/i.test(g(r, 'group')) ? '' : g(r, 'group');
    var key = [g(r, 'cls'), g(r, 'section'), g(r, 'program'), grp].join('|');
    (by[key] = by[key] || { cls: g(r, 'cls'), section: g(r, 'section'), program: g(r, 'program'), group: grp, rolls: [] }).rolls.push(roll);
  }
  return { ok: true, list: Object.keys(by).map(function (k) { return by[k]; }), at: new Date().toISOString() };
}

/* ---------- staff: Users sheet = teachers now in the college, X-Users = teachers who left ----------
   Users columns: teacher, Status, designation, Department, User Name, Password, Contact (found by heading).
   Passwords and user names are never sent. */
function dutStaff_() {
  var book = SpreadsheetApp.openById(DUT_CFG.USERS), K = function (x) { return dutS_(x).toUpperCase().replace(/[^A-Z0-9]/g, ''); };
  var read = function (name) {
    var sh = book.getSheetByName(name); if (!sh) return [];
    var v = sh.getDataRange().getDisplayValues(); if (v.length < 2) return [];
    var head = v[0].map(K), c = function () { for (var i = 0; i < arguments.length; i++) { var j = head.indexOf(arguments[i]); if (j >= 0) return j; } return -1; };
    var C = { name: c('TEACHER', 'TEACHERNAME', 'NAME'), status: c('STATUS'), desig: c('DESIGNATION'), dept: c('DEPARTMENT', 'DEPT'), user: c('USERNAME', 'USER'), phone: c('CONTACT', 'PHONE', 'MOBILE', 'WHATSAPP') };
    if (C.name < 0) C.name = 0;
    var g = function (r, k) { return C[k] >= 0 ? dutS_(r[C[k]]) : ''; };
    return v.slice(1).map(function (r) { return { name: g(r, 'name'), status: g(r, 'status'), desig: g(r, 'desig'), dept: g(r, 'dept'), phone: g(r, 'phone'), user: g(r, 'user') }; })
      .filter(function (x) { return x.name && !/^admin$/i.test(x.name) && !/^admin$/i.test(x.user); })
      .map(function (x) { delete x.user; return x; });
  };
  return { ok: true, users: read('Users'), left: read('X-Users').map(function (x) { return { name: x.name, dept: x.dept, desig: x.desig }; }) };
}
