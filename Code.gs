var REPORTS_SHEET = 'Chief Proctor Reports';
var PROCTORS_SHEET = 'Chief Proctors';   // columns: Id | Name | Phone | Dept | Desig   (phone as digits with country code, e.g. 923001234567)
var REPORT_HEAD = ['Report ID','Received','Teacher','Username','Student','Roll No','Class','Subject','Month','Incident date','Lecture','Reason','Severity','Statement','Attendance','Sent to','Via','Guest','Status','Chief Proctor remark'];


// One attendance workbook per month
var BOOK_IDS = {
  SEPT: '1tKdOESL-cqP9IXafk8IbMNv8Y9vRG59a5gPJyCRSCXM',
  OCT:  '10X8dQGSMGvq0S369d6Go2jY2BVKtbU0XfQre8Xi7PBI',
  NOV:  '1B0KJVbMKiXjr8UkkUbsMwmaeZxbg74q59Hldgsskjcc',
  DEC:  '1O0BiNIys7EHzGwLFpfq0tQLW4jFcd2NxkJSkwxBPb3k',
  JAN:  '1H_1mzuNUuJ_sDfQTXmhU0p8MBYul6yb2eNFq0tqYxjE'
};
// Workbook that holds the sheets: Users, Timetable, Transform-ALL-COLLEGE-STRENGTH
var USERS_SS_ID = '1ARU8zMO2daJUfU_SIIo4baeSOG7qBPuni5TYRIRLBNY';

var ADMIN_USER      = 'jcw@uop.edu.pk';  // sees every section
var PROGRAM_FILTER  = '';                // e.g. 'FCS' to only use that Program from the Timetable. '' = all
var HEADER_ROW = 3;
var FIRST_ROW  = 4;

var DAY_COL1  = 7;                       // G = day 1
var LAST_DAY  = DAY_COL1 + 30;           // AK = day 31
var TA_COL    = LAST_DAY + 1;            // AL = T.A
var MT_COL    = LAST_DAY + 2;            // AM = M.T
var TITLE_PREFIX = 'JCW Attendance - ';

var VALID = { P: 1, A: 1, L: 1, B: 1 };  // marks that may be written. ONLY "P" is counted.

// Period-wise attendance
var PD_SHEET   = 'Period Data';          // hidden sheet in every month workbook (real marks of every period)
var MAX_PERIOD = 8;                      // periods 1..8
var NONP       = ['L', 'A', 'B'];        // student has NO P in any period: this letter wins first (most lenient first)

var SEASON       = ['SEPT', 'OCT', 'NOV', 'DEC', 'JAN'];
var SEASON_MONTH = { SEPT: 8, OCT: 9, NOV: 10, DEC: 11, JAN: 0 };
var MONTH_NAMES  = ['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE','JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER'];
var CACHE_TTL    = 600;
var CACHE_KEYS   = ['users_att2', 'tt_att3', 'master_att3', 'ptimes'];

// Big workbook settings
var SPARE_ROWS = 2;          // empty rows kept under the last student (so a teacher can still add a line by hand)
var BATCH_MS   = 240000;     // bulk jobs stop after 4 minutes and continue by themselves (Apps Script stops at 6)

// Disabled students (changed subject / section)
var OFF_TAG = 'LEFT SUBJECT';   // note on the Roll No cell of a disabled student

// Student remarks
var REMARKS_SHEET = 'Student Remarks';
var REMARK_HEAD   = ['Key', 'Teacher', 'Class', 'Roll No', 'Student', 'Tags', 'Latest comment', 'Comments', 'Data (do not edit)', 'Updated (ms)', 'Updated'];
var REMARK_MAX_LOG = 50, REMARK_MAX_TAGS = 10;

function doGet() {
  return json_({ ok: true, service: 'JCW Daily Attendance backend v7.5' });
}

function doPost(e) {
  var req = {};
  try { req = JSON.parse(e.postData.contents); } catch (err) {
    return json_({ ok: false, error: 'BAD_REQUEST' });
    
  }

  // v6.8: the UOP portal sync button sends student data with the sync key (no teacher login)
  if (req.action === 'portal_import') {
    try { return json_(portalImport_(req)); }
    catch (err) { return json_({ ok: false, error: String(err && err.message || err) }); }
  }
    if (req.action === 'portal_left') {                                  // v7.3: the sync panel's Left Students buttons (sync key)
    try { return json_(portalLeftDo_(req)); }
    catch (err) { return json_({ ok: false, error: String(err && err.message || err) }); }
  }
  // v6.7: the public student result page (no teacher login; roll no + private code)
  if (req.action === 'student_days') {                                // v7.2: her day-by-day attendance (roll no + code again)
    try { return json_(studentDays_(req.roll, req.code, req.mon)); }
    catch (err) { return json_({ ok: false, error: 'The result page is busy. Try again in a minute.' }); }
  }
  if (req.action === 'student_report') {
    try { return json_(studentReport_(req.roll, req.code)); }
    catch (err) { return json_({ ok: false, error: 'The result page is busy. Try again in a minute.' }); }
  }

    if (/^idp_/.test(String(req.action || ''))) return json_(idcPublic_(req));   // ID card check: students, no login

  var user = authenticate_(req.username, req.password);
  if (!user) return json_({ ok: false, error: 'BAD_LOGIN' });
  if (/^adm_/.test(String(req.action || ''))) return json_(admRoute_(user, req));
     if (/^gt_/.test(String(req.action || ''))) return json_(gtRoute_(user, req));
        if (/^bk_/.test(String(req.action || ''))) return json_(bkRoute_(user, req));
           if (/^dt_/.test(String(req.action || ''))) return json_(dtRoute_(user, req));
              if (/^wr_/.test(String(req.action || ''))) return json_(wrRoute_(user, req));
                 if (/^st_/.test(String(req.action || ''))) return json_(stRoute_(user, req));
                 if (/^br_/.test(String(req.action || ''))) return json_(brRoute_(user, req));
                 if (/^pc_/.test(String(req.action || ''))) return json_(pcRoute_(user, req));
                 if (/^idc_/.test(String(req.action || ''))) return json_(idcRoute_(user, req));
                 if (/^dut_/.test(String(req.action || ''))) return json_(dutRoute_(user, req));
                  if (/^rpt_/.test(String(req.action || ''))) return json_(rptRoute_(user, req));
                    if (/^bnk_/.test(String(req.action || ''))) return json_(bnkRoute_(user, req));
                      if (/^lv_/.test(String(req.action || ''))) return json_(lvRoute_(user, req));
                       if (/^shr_/.test(String(req.action || ''))) return json_(shrRoute_(user, req));
                        if (/^rsm_/.test(String(req.action || ''))) return json_(rsmRoute_(user, req));
                          if (/^ntt_/.test(String(req.action || ''))) return json_(nttRoute_(user, req));
                            if (/^at_/.test(String(req.action || ''))) return json_(atRoute_(user, req));
                 
             
                
  try {
    if (req.action === 'sheets') {
  return json_({ ok: true, teacher: user.teacher, admin: isAdmin_(user), sheets: listSheets_(user), proctors: getProctors_(), view: scopeName_(user) });
}
    if (req.action === 'whoami')       return json_(whoami_(user));
    if (req.action === 'overview')     return json_(overview_(user));
    if (req.action === 'overview_low') return json_(overviewLow_(user));
    if (req.action === 'overview_today') return json_(overviewToday_(user));
    if (req.action === 'report_data')  return json_(reportData_(user, req));
    if (req.action === 'reg_opts')     return json_(regOpts_(user));
    if (req.action === 'reg_find')     return json_(regFind_(user, req.q));
    if (req.action === 'reg_plan')     return json_(regPlan_(user, req));
    if (req.action === 'reg_days')     return json_(regDays_(user, req));
    if (req.action === 'tt_all')       return json_(ttAll_(user));
    if (req.action === 'contacts')     return json_(contacts_(user));
    if (req.action === 'report_save')  return json_(saveReports_(user, req.records || []));
if (req.action === 'reports_mine') return json_(getMyReports_(user));
if (req.action === 'report_delete') return json_(deleteReport_(user, req.id));
    if (req.action === 'guest_search') return json_(guestSearch_(user, req.sheet, req.q));
    if (req.action === 'guest_add')    return json_(guestAdd_(user, req.sheet, req.roll));
    if (req.action === 'guest_remove') return json_(guestRemove_(user, req.sheet, req.roll));
    if (req.action === 'sheet')         return json_(getSheetData_(user, req.sheet, req.create === true));
    if (req.action === 'save')          return json_(saveRecords_(user, req.records || []));
    if (req.action === 'delete_day')    return json_(deleteDay_(user, req.sheet, req.day, req.confirmDay, req.period));
    if (req.action === 'pdf')           return json_(exportPdf_(user, req.sheet));
    if (req.action === 'password')      return json_(changePassword_(user, req.oldPass, req.newPass));
    if (req.action === 'remarks_get')   return json_(getRemarks_(user));
    if (req.action === 'remarks_save')  return json_(saveRemarks_(user, req.records || []));
    // v6.7
    if (req.action === 'history')       return json_(history_(user, req));
    if (req.action === 'sub_list')      return json_(subList_(user, req));
    if (req.action === 'sub_save')      return json_(subSave_(user, req.rec || {}));
    if (req.action === 'sub_delete')    return json_(subDelete_(user, req.id));
    if (req.action === 'health')        return json_(health_(user, req.part));
    if (req.action === 'sheets_status') return json_(sheetsStatus_(user));
    if (req.action === 'sub_result')    return json_(subResult_(user, req.id, req.result));
    if (req.action === 'duty_save')     return json_(dutySave_(user, req.settings || {}));
    if (req.action === 'left_preview')  return json_(leftPreview_(user, req.statuses));
    if (req.action === 'left_move')     return json_(leftMove_(user, req.rolls || []));
    if (req.action === 'left_list')     return json_(leftList_(user));
    if (req.action === 'left_restore')  return json_(leftRestore_(user, req.roll));
    if (req.action === 'roster_check')  return json_(rosterCheck_(user, req.mon, req.cursor));
    if (req.action === 'roster_fix')    return json_(rosterFix_(user, req.sheet, req.roll, req.op));
    if (req.action === 'health_fix')    return json_(healthFix_(user, req.what));
    if (req.action === 'codes')         return json_(studentCodes_(user, req.cls, req.renew === true));
    if (req.action === 'results_info')  return json_(resultsInfo_(user));
    if (req.action === 'portal_key')    return json_(portalKey_(user, req.renew === true));
    if (req.action === 'portal_info')   return json_(portalInfo_(user));
    return json_({ ok: false, error: 'UNKNOWN_ACTION' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}
function deleteReport_(user, id) {
  id = String(id || '').trim();
  if (!id) return { ok: false, error: 'Missing report id.' };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = reportsSheet_(), n = sh.getLastRow() - 1;
    if (n < 1) return { ok: false, error: 'REPORT_NOT_FOUND' };
    var v = sh.getRange(2, 1, n, REPORT_HEAD.length).getValues();
    var me = String(user.username).trim().toLowerCase(), admin = isAdmin_(user);
    for (var i = 0; i < n; i++) {
      if (String(v[i][0]).trim() !== id) continue;
      if (!admin && String(v[i][3]).trim().toLowerCase() !== me) return { ok: false, error: 'This is not your report.' };
      var st = String(v[i][18] || 'Received').trim();
      if (!admin && st !== 'Received') return { ok: false, error: 'The Chief Proctor already acted on this report (' + st + '). It cannot be deleted.' };
      sh.deleteRow(i + 2);
      SpreadsheetApp.flush();
      var chk = sh.getLastRow() - 1, still = false;
      if (chk > 0) sh.getRange(2, 1, chk, 1).getValues().forEach(function (r) { if (String(r[0]).trim() === id) still = true; });
      return still ? { ok: false, error: 'Row could not be removed. Try again.' } : { ok: true, id: id, deleted: true };
    }
    return { ok: false, error: 'REPORT_NOT_FOUND' };
  } finally { lock.releaseLock(); }
}
function getProctors_() {
  var v = usersTable_();
  if (!v.length) return [];
  var n = function (s) { return String(s).toLowerCase().replace(/[^a-z0-9]/g, ''); };
  var hdr = v[0].map(n);
  var col = function (name, fallback) { var i = hdr.indexOf(name); return i === -1 ? fallback : i; };
  var iName = col('teacher', 0), iDes = col('designation', 2), iDept = col('department', 3),
      iPh = col('contact', 6), iRole = col('role', 8), iCP = col('chiefproctor', 9);
  var title = function (s) { return String(s).toLowerCase().replace(/\b\w/g, function (c) { return c.toUpperCase(); }); };
  var out = [];
  for (var i = 1; i < v.length; i++) {
    var tag = String(v[i][iCP]).trim().toUpperCase();
    if (!/^CP\d*$/.test(tag)) continue;                       // accepts CP, CP1, CP2 ...
    var role = String(v[i][iRole]).trim().toUpperCase();
    if (role === 'PRINCIPAL' || role === 'ADMIN') continue;   // not real proctors
    var phone = String(v[i][iPh]).replace(/\D/g, '');
    if (!phone) continue;
    out.push({
      id: 'CP' + (out.length + 1),                            // CP1, CP2 ... by row order
      name: title(v[i][iName]),
      dept: title(v[i][iDept]),
      desig: title(v[i][iDes]),
      phone: phone                                            // +923158813443 -> 923158813443
    });
  }
  return out;
}

function reportsSheet_() {
  var book = SpreadsheetApp.openById(USERS_SS_ID), sh = book.getSheetByName(REPORTS_SHEET);
  if (sh) return sh;
  sh = book.insertSheet(REPORTS_SHEET);
  sh.getRange(1, 1, 1, REPORT_HEAD.length).setValues([REPORT_HEAD]).setFontWeight('bold').setBackground('#D9D9D9');
  sh.setFrozenRows(1);
  sh.getRange(2, 6, 1000, 1).setNumberFormat('@');
  sh.getRange(2, 10, 1000, 2).setNumberFormat('@');   // incident date + lecture stay text
  sh.setColumnWidth(14, 400);
  return sh;
}

function saveReports_(user, records) {
  var okNames = allowed_(user), saved = [], failed = [], results = {};
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = reportsSheet_(), n = sh.getLastRow() - 1, have = {};
    if (n > 0) sh.getRange(2, 1, n, 1).getValues().forEach(function (r) { have[String(r[0])] = true; });
    var cps = getProctors_(), rows = [];
    records.forEach(function (rec) {
      try {
        if (!rec.id) throw new Error('Missing report id');
        if (have[rec.id]) { saved.push(rec.id); results[rec.id] = {}; return; }   // already stored: safe to re-send
        var meta = okNames[rec.sheet];
        if (!meta) throw new Error('Not allowed: ' + rec.sheet);
        var txt = String(rec.txt || '').trim();
        if (txt.length < 10) throw new Error('Statement is too short.');
        var cp = cps.filter(function (c) { return c.id === rec.cp; })[0];
        rows.push([rec.id, new Date(), user.teacher, user.username, String(rec.who || ''), String(rec.roll || ''), meta.cls, meta.sub, meta.mon,
                   String(rec.date || ''), String(rec.period || ''), String(rec.cat || ''), rec.sev === 'Serious' ? 'Serious' : 'Normal',
                   txt.slice(0, 1000), String(rec.att || ''), cp ? cp.name : '', String(rec.via || ''), rec.guest ? (rec.home || 'Guest') : '', 'Received', '']);
        have[rec.id] = true; saved.push(rec.id);
        if (rec.removeGuest) {
          try { removeGuestCore_(meta, String(rec.roll || '').trim()); results[rec.id] = { guestRemoved: true }; }
          catch (ge) { results[rec.id] = { guestRemoved: false, guestError: String(ge.message || ge) }; }
        } else results[rec.id] = {};
      } catch (err) { failed.push({ id: rec.id, error: String(err.message || err) }); }
    });
    if (rows.length) {
      var start = sh.getLastRow() + 1;
      sh.getRange(start, 1, rows.length, 1).setNumberFormat('@');sh.getRange(start, 10, rows.length, 2).setNumberFormat('@');
      sh.getRange(start, 6, rows.length, 1).setNumberFormat('@');sh.getRange(start, 10, rows.length, 2).setNumberFormat('@');
      sh.getRange(start, 1, rows.length, REPORT_HEAD.length).setValues(rows);
    }
    SpreadsheetApp.flush();
    return { ok: true, saved: saved, failed: failed, results: results };
  } finally { lock.releaseLock(); }
}

function getMyReports_(user) {
  var sh = reportsSheet_(), n = sh.getLastRow() - 1, out = [];
  if (n < 1) return { ok: true, reports: out };
  var me = String(user.username).trim().toLowerCase();
  sh.getRange(2, 1, n, REPORT_HEAD.length).getValues().forEach(function (r) {
    if (String(r[3]).trim().toLowerCase() !== me) return;
    out.push({ id: r[0], at: Utilities.formatDate(new Date(r[1]), Session.getScriptTimeZone(), 'd MMM yyyy HH:mm'), teacher: r[2], who: r[4], roll: r[5],
      cls: r[6], sub: r[7], mon: r[8], date: (r[9] instanceof Date) ? Utilities.formatDate(r[9], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[9]),
lecture: r[10], cat: r[11], sev: r[12], txt: r[13], att: r[14],
      sent: r[15], guest: !!r[17], home: r[17], status: r[18] || 'Received' });
  });
  out.reverse();
  return { ok: true, reports: out };
}
/* ---------- small helpers ---------- */

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function norm_(v) {
  return String(v == null ? '' : v).replace(/\s+/g, '').toUpperCase();
}

function normalizeName(name) {
  if (!name) return '';
  return name.toString().toLowerCase()
    .replace(/\b(dr|ms|mr|mrs)\b\.?/gi, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function colLetter_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function isAdmin_(user) {
  return String(user.username).trim().toLowerCase() === ADMIN_USER;
}

/** True when a Roll No cell note marks the student as disabled (left the subject). */
function isOffNote_(n) { n = String(n || ''); return n.indexOf(OFF_TAG) === 0 || n.indexOf(LEFT_TAG) === 0; }
var LEFT_TAG = 'LEFT COLLEGE';   // v7.0: note on the Roll No cell of a student who left the college (migrated, cancelled …)

function academicStartYear_() {
  var d = new Date(), y = d.getFullYear(), m = d.getMonth();
  return (m > 7 || (m === 7 && d.getDate() >= 15)) ? y : y - 1;
}

function monthYear_(tok) {
  var m = SEASON_MONTH[tok];
  var start = academicStartYear_();
  return { m: m, y: m >= 7 ? start : start + 1 };
}

/** Months shown to teachers: from SEPT up to the current month (all of them outside the season). */
function activeMonths_() {
  var m = new Date().getMonth();
  var idx = m >= 8 ? m - 8 : (m === 0 ? 4 : -1);
  return idx < 0 ? SEASON.slice() : SEASON.slice(0, idx + 1);
}

/** Season token of the current calendar month, or null outside Sept-Jan. */
function currentSeasonToken_() {
  var map = { 8: 'SEPT', 9: 'OCT', 10: 'NOV', 11: 'DEC', 0: 'JAN' };
  return map[new Date().getMonth()] || null;
}

/* ---------- one workbook per month ---------- */

var BOOK_CACHE_ = {};

/** Spreadsheet id of the workbook that holds the sheets of this month token (SEPT, OCT, NOV, DEC, JAN). */
function bookIdFor_(tok) {
  var id = BOOK_IDS[String(tok || '').toUpperCase()];
  if (!id) throw new Error('No workbook is set for month "' + tok + '". Check BOOK_IDS.');
  return id;
}

/** Opens the workbook of that month (opened once per request). */
function bookFor_(tok) {
  var id = bookIdFor_(tok);
  if (!BOOK_CACHE_[id]) BOOK_CACHE_[id] = SpreadsheetApp.openById(id);
  return BOOK_CACHE_[id];
}

/** All five month workbooks, in season order: [{ tok, book }] */
function allBooks_() {
  return SEASON.map(function (tok) { return { tok: tok, book: bookFor_(tok) }; });
}

/* ---------- cache helpers (values over 100KB are chunked) ---------- */

function cachePutBig_(key, obj) {
  var c = CacheService.getScriptCache();
  var s = JSON.stringify(obj), size = 30000, n = Math.max(1, Math.ceil(s.length / size)), o = {};
  for (var i = 0; i < n; i++) o[key + '_' + i] = s.substr(i * size, size);
  o[key + '_n'] = String(n);
  try { c.putAll(o, CACHE_TTL); } catch (e) {}
}

function cacheGetBig_(key) {
  var c = CacheService.getScriptCache();
  var n = Number(c.get(key + '_n'));
  if (!n) return null;
  var keys = [];
  for (var i = 0; i < n; i++) keys.push(key + '_' + i);
  var m = c.getAll(keys), s = '';
  for (var j = 0; j < keys.length; j++) {
    if (m[keys[j]] == null) return null;
    s += m[keys[j]];
  }
  try { return JSON.parse(s); } catch (e) { return null; }
}

function cacheClear_(k) {
  var c = CacheService.getScriptCache();
  var n = Number(c.get(k + '_n')) || 0, keys = [k + '_n'];
  for (var i = 0; i < n; i++) keys.push(k + '_' + i);
  c.removeAll(keys);
}

/** Run this by hand after you change the Timetable, Users or the strength sheet. */
function clearAppCaches() {
  CACHE_KEYS.concat(['master_att', 'users_att', 'tt_att', 'tt_att2', 'users_full']).forEach(cacheClear_);
    SEASON.forEach(function (t) { cacheClear_('ll2_' + t); });
}

/* ---------- login ---------- */

/** Column number of a Users header (lower-case, letters/digits only), or the fallback when the header is missing. */
function usersCol_(hdr, name, fallback) {
  var h = hdr.map(function (s) { return String(s).toLowerCase().replace(/[^a-z0-9]/g, ''); });
  var i = h.indexOf(name);
  return i === -1 ? fallback : i;
}

/** The whole Users sheet (as text), kept in the cache for 10 minutes. clearAppCaches / a password change empties it. */
var USERS_MEMO_ = null;
function usersTable_() {
  if (USERS_MEMO_) return USERS_MEMO_;
  var v = cacheGetBig_('users_full');
  if (!v) {
    var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName('Users');
    v = sh ? sh.getDataRange().getValues().map(function (r) { return r.map(function (x) { return x == null ? '' : String(x); }); }) : [];
    cachePutBig_('users_full', v);
  }
  return (USERS_MEMO_ = v);
}

function getUsers_() {
  var rows = cacheGetBig_('users_att2');
  if (rows) return rows;
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName('Users');
  if (!sh) throw new Error('Users sheet not found');
  var d = sh.getDataRange().getValues();
  var iRole = usersCol_(d[0], 'role', 8), iDept = usersCol_(d[0], 'department', 3);
  rows = [];
  for (var i = 1; i < d.length; i++) {
    if (!d[i][4]) continue;
    // teacher, username, password, role, department
    rows.push([String(d[i][0]).trim(), String(d[i][4]).trim(), String(d[i][5]),
               String(d[i][iRole] || '').trim(), String(d[i][iDept] || '').trim()]);
  }
  cachePutBig_('users_att2', rows);
  return rows;
}

function authenticate_(username, password) {
  var u = String(username || '').trim().toLowerCase();
  if (!u) return null;
  var rows = getUsers_();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][1].toLowerCase() === u && rows[i][2] === String(password)) {
      return { teacher: rows[i][0], username: rows[i][1], password: rows[i][2], role: rows[i][3] || '', dept: rows[i][4] || '' };
    }
  }
  return null;
}

function changePassword_(user, oldPass, newPass) {
  if (String(oldPass) !== user.password) return { ok: false, error: 'WRONG_OLD' };
  if (!newPass || String(newPass).length < 4) return { ok: false, error: 'TOO_SHORT' };
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName('Users');
  var d = sh.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][4]).trim().toLowerCase() === user.username.toLowerCase()) {
      sh.getRange(i + 1, 6).setValue(String(newPass));
      cacheClear_('users_att2'); cacheClear_('users_full');
      return { ok: true };
    }
  }
  return { ok: false, error: 'User not found' };
}

/* ---------- timetable -> this teacher's sections ---------- */

function getTimetable_() {
  var t = cacheGetBig_('tt_att3');
  if (t) return t;
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName('Timetable');
  if (!sh) throw new Error('Timetable sheet not found');
  var data = sh.getDataRange().getValues();
  t = { has: {}, rows: [] };
  if (data.length >= 2) {
    var h = data[0];
    // row = [teacher, program, class, pcs, subject, period, theory days, practical days, days, room, course name, section]
    var names = ['Teacher', 'Program', 'Class', 'PCS', 'Subject', 'Period', 'Theory', 'Practical', 'Days', 'Room', 'Course Name', 'Section'];
    var hl = h.map(function (x) { return String(x).trim().toLowerCase(); });
    var ix = names.map(function (n) { var i = h.indexOf(n); if (i === -1) i = hl.indexOf(n.toLowerCase()); t.has[n] = i !== -1; return i; });
    for (var i = 1; i < data.length; i++) {
      t.rows.push(ix.map(function (c) { return c === -1 ? '' : String(data[i][c] || '').trim(); }));
    }
  }
  cachePutBig_('tt_att3', t);
  return t;
}

/** [{cls:'FCS-1 D2', subject:'COMPUTER SCIENCE', teacher:'Asima Latif'}] - only this teacher's rows (admin: all). */
function getAssignments_(user) {
  var t = getTimetable_();
  ['PCS', 'Subject', 'Teacher'].forEach(function (n) {
    if (t.rows.length && !t.has[n]) throw new Error('Timetable must contain a "' + n + '" column.');
  });
  var admin = isAdmin_(user), me = normalizeName(user.teacher), seen = {}, out = [];
  t.rows.forEach(function (r) {
    var teacher = r[0], program = r[1], pcs = r[3], subject = r[4];
    if (!pcs || !subject) return;
    if (PROGRAM_FILTER && program.toUpperCase() !== PROGRAM_FILTER.toUpperCase()) return;
    if (!admin && normalizeName(teacher) !== me) return;
    var key = pcs.toUpperCase() + '|' + subject.toUpperCase();
    if (seen[key]) return;
    seen[key] = true;
    out.push({ cls: pcs, subject: subject.toUpperCase(), teacher: teacher });
  });
  return out;
}

/** "F.A-1 Optional Subjects" -> "F.A-1" (other classes unchanged) */
function baseClass_(cls) {
  return String(cls || '').replace(/\s*Optional\s*Subjects?/i, '').replace(/\s+/g, ' ').trim();
}

/** Section letter from a student's PCS: "FCS-1 D2" -> "D2", "F.A-1" -> "" */
function sectionOf_(classSec) {
  var p = String(classSec || '').trim().split(/\s+/);
  return p.length > 1 ? p[p.length - 1] : '';
}

function makeName_(cls, mon, subject) {
  return (baseClass_(cls) + ' ' + mon + ' ' + subject)
    .replace(/[\[\]\*\?:\/\\]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 100);
}

function listSheets_(user) {
  var months = activeMonths_(), out = [];
  getAssignments_(user).forEach(function (a) {
    months.forEach(function (mon) {
      out.push({ name: makeName_(a.cls, mon, a.subject), cls: a.cls, sub: a.subject, mon: mon, teacher: a.teacher });
    });
  });
  out.sort(function (x, y) {
    return x.cls.localeCompare(y.cls, undefined, { numeric: true }) ||
           x.sub.localeCompare(y.sub) ||
           SEASON.indexOf(x.mon) - SEASON.indexOf(y.mon);
  });
  return out;
}

/** name -> meta, only for sections this user may touch. */
function allowed_(user) {
  var map = {};
  listSheets_(user).forEach(function (s) { map[s.name] = s; });
  return map;
}

/* ---------- student roster for a new sheet ---------- */

function getMaster_() {
  var M = cacheGetBig_('master_att3');
  if (M) return M;
  M = { hasClass: false, rows: [] };      // rows: [enroll, name, classSec, [SUBJECTS...], fatherName]
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName('Transform-ALL-COLLEGE-STRENGTH');
  if (sh && sh.getLastRow() > 1) {
    var v = sh.getDataRange().getValues();
    var h = v[0].map(norm_);
    var find = function (list) { for (var i = 0; i < list.length; i++) { var x = h.indexOf(list[i]); if (x !== -1) return x; } return -1; };
    var ec = h.indexOf('ENROLL#');
    var nc = find(['NAME', 'STUDENTNAME', 'STUDENT', 'FULLNAME']);
    var cc = find(['CLASS-SEC', 'CLASSSEC', 'CLASS-SECTION', 'PCS', 'SECTION']);
    var fc = find(['FATHERNAME', 'FATHER', "FATHER'SNAME", 'FATHER/GUARDIAN', 'GUARDIANNAME']);
    var sc = ['SUB1', 'SUB2', 'SUB3', 'SUB4', 'SUB5', 'SUB6', 'SUB7'].map(function (s) { return h.indexOf(s); })
      .filter(function (i) { return i !== -1; });
    M.hasClass = cc !== -1;
    if (ec !== -1) {
      for (var i = 1; i < v.length; i++) {
        var e = v[i][ec];
        if (!e) continue;
        var subs = [];
        sc.forEach(function (c) { var s = String(v[i][c] || '').trim().toUpperCase(); if (s) subs.push(s); });
        M.rows.push([String(e).trim(), nc === -1 ? String(e).trim() : String(v[i][nc] || '').trim(),
                     cc === -1 ? '' : String(v[i][cc] || '').trim(), subs,
                     fc === -1 ? '' : String(v[i][fc] || '').trim()]);
      }
    }
  }
  cachePutBig_('master_att3', M);
  return M;
}

function enrollNum_(e) {
  var p = String(e || '').split('-');
  return parseInt(p[p.length - 1], 10) || 0;
}

/** Returns [[roll, name, father, classSec], ...] sorted by roll number. */
function getRoster_(book, meta) {
  var M = getMaster_();
  var subj = meta.sub.toUpperCase();
  var squash = function (s) { return String(s || '').replace(/\s+/g, '').toUpperCase(); };
  var base = baseClass_(meta.cls);
  var baseSq = squash(base);
  var isOpt = baseSq !== squash(meta.cls);
  var classOf = {};
  M.rows.forEach(function (r) { classOf[r[0]] = r[2] || ''; });
  var list = [];

  // 1) Master sheet: students of this class who take this subject
  if (M.hasClass) {
    M.rows.forEach(function (r) {
      var c = squash(r[2]);
      var classOk = isOpt ? (c === baseSq || c.indexOf(baseSq) === 0) : (c === baseSq);
      if (classOk && r[3].indexOf(subj) !== -1) list.push([r[0], r[1], r[4] || '', r[2] || base]);
    });
  }

  // 2) Fallback: roster of another attendance sheet of the same class, filtered by subject
  var fromSibs = function (sibs) {
    for (var k = 0; k < sibs.length && !list.length; k++) {
      try {
        var L = readLayout_(sibs[k]);
        var n = sibs[k].getLastRow() - FIRST_ROW + 1;
        if (n < 1) continue;
        var rolls = sibs[k].getRange(FIRST_ROW, L.rollCol, n, 1).getValues();
        var names = L.nameCol > 0 ? sibs[k].getRange(FIRST_ROW, L.nameCol, n, 1).getValues() : null;
        var taking = {}, any = false;
        M.rows.forEach(function (r) { if (r[3].indexOf(subj) !== -1) { taking[r[0]] = true; any = true; } });
        var fathers = {};
        M.rows.forEach(function (r) { fathers[r[0]] = r[4] || ''; });
        for (var i = 0; i < n; i++) {
          var roll = String(rolls[i][0]).trim();
          if (!roll) continue;
          if (any && !taking[roll]) continue;
          list.push([roll, names ? String(names[i][0]).trim() : roll, fathers[roll] || '', classOf[roll] || base]);
        }
      } catch (e) { /* try next sheet */ }
    }
  };

  // 2a) fast path: the same class + subject in the other months. Each month lives in its own workbook,
  //     so every month's workbook is checked with a direct lookup (no scan of all tabs).
  if (!list.length) {
    var fast = [];
    SEASON.forEach(function (tok) {
      var nm = makeName_(meta.cls, tok, meta.sub);
      if (nm === meta.name) return;
      try {
        var s = bookFor_(tok).getSheetByName(nm);
        if (s) fast.push(s);
      } catch (e) { /* month workbook not reachable, skip it */ }
    });
    fromSibs(fast);
  }

  // 2b) slow path: any sheet of the same class, in this month's workbook first, then the other months
  if (!list.length) {
    var sibs = [], seenIds = {};
    var scan = function (bk) {
      var id = bk.getId();
      if (seenIds[id]) return;
      seenIds[id] = true;
      bk.getSheets().forEach(function (s) {
        if (s.getName().toUpperCase().indexOf(base.toUpperCase() + ' ') === 0 && s.getName() !== meta.name) sibs.push(s);
      });
    };
    scan(book);
    SEASON.forEach(function (tok) { try { scan(bookFor_(tok)); } catch (e) {} });
    fromSibs(sibs);
  }

  list.sort(function (a, b) { return enrollNum_(a[0]) - enrollNum_(b[0]); });
  return list;
}

function titleCase_(s) {
  return String(s).toLowerCase().replace(/\b\w/g, function (c) { return c.toUpperCase(); });
}

/** Writes "Teacher: <name>" in A2 (only if A2 is empty or already a teacher stamp). */
function stampTeacher_(sheet, meta) {
  if (!meta || !meta.teacher) return;
  var cell = sheet.getRange(2, 1);
  var v = String(cell.getValue() || '');
  var want = 'Teacher: ' + meta.teacher;
  if (v === want) return;
  if (!v || /^Teacher:/i.test(v)) cell.setValue(want).setFontWeight('bold').setHorizontalAlignment('left');
}

/** Row 2, above each date: Mon, Tue, Wed... (blank for days that don't exist in the month). */
function setWeekdayRow_(sh, month, year, firstDayCol) {
  var col = firstDayCol || DAY_COL1;
  var names = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  var dim = new Date(year, month + 1, 0).getDate(), vals = [];
  for (var d = 1; d <= 31; d++) vals.push(d > dim ? '' : names[new Date(year, month, d).getDay()]);
  sh.getRange(2, col, 1, 31).setValues([vals])
    .setFontSize(8).setFontWeight('bold').setHorizontalAlignment('center');
}

/**
 * Removes the unused empty rows and columns of a sheet. A new Google sheet is 1,000 x 26 cells and every
 * cell (even empty) counts toward the 10,000,000 cell limit of the workbook.
 * usedRows = last row that holds data, usedCols = last column that holds data (at least the M.T column).
 */
function trimSheet_(sh, usedRows, usedCols) {
  var keepRows = Math.max(usedRows, FIRST_ROW) + SPARE_ROWS, maxR = sh.getMaxRows();
  if (maxR > keepRows) sh.deleteRows(keepRows + 1, maxR - keepRows);
  var keepCols = Math.max(usedCols || MT_COL, MT_COL), maxC = sh.getMaxColumns();
  if (maxC > keepCols) sh.deleteColumns(keepCols + 1, maxC - keepCols);
}

/* ---------- period-wise attendance: formulas, colours, hidden "Period Data" sheet ---------- */

/**
 * T.A of one row = number of lectures the student attended.
 * Cells look like P (1), 2P, 3P ... -> the number in front of the P (1 when there is no number).
 * A, L, B, 2A, 3L ... count 0.
 */
function taFormula_(row) {
  var r = colLetter_(DAY_COL1) + row + ':' + colLetter_(LAST_DAY) + row;
  return '=SUMPRODUCT(ARRAYFORMULA(IFERROR(VALUE(REGEXEXTRACT(' + r + '&"","^\\d+")),1)*(RIGHT(' + r + '&"",1)="P")))';
}

/** Colour rules of an attendance sheet: A red, L amber, B blue (cell CONTAINS the letter, so 3A, 2P+... work);
 *  low attendance (<73%) and low monthly test (<50%) highlighted. n = number of student rows. */
function applyRules_(sh, n) {
  var days = sh.getRange(FIRST_ROW, DAY_COL1, n, 31);
  var taL = colLetter_(TA_COL), mtL = colLetter_(MT_COL);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('A')
      .setBackground('#FF0000').setFontColor('#FFFFFF').setBold(true).setRanges([days]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('L')
      .setBackground('#FFE599').setFontColor('#7F6000').setBold(true).setRanges([days]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextContains('B')
      .setBackground('#9FC5E8').setFontColor('#0B3D6B').setBold(true).setRanges([days]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($' + taL + '$2>0,$' + taL + FIRST_ROW + '/$' + taL + '$2<0.73)')
      .setBackground('#F4C7C3').setFontColor('#990000')
      .setRanges([sh.getRange(FIRST_ROW, 4, n, 1), sh.getRange(FIRST_ROW, TA_COL, n, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($' + mtL + FIRST_ROW + '),$' + mtL + '$2>0,$' + mtL + FIRST_ROW + '/$' + mtL + '$2<0.5)')
      .setBackground('#F4C7C3').setFontColor('#990000')
      .setRanges([sh.getRange(FIRST_ROW, MT_COL, n, 1)]).build()
  ]);
}

/** The hidden sheet that keeps the real marks of every period (one row per attendance sheet). */
function pdSheet_(book) {
  var sh = book.getSheetByName(PD_SHEET);
  if (sh) return sh;
  sh = book.insertSheet(PD_SHEET);
  sh.getRange(1, 1, 1, 3).setValues([['Sheet name', 'Period data (do not edit)', 'Updated']]).setFontWeight('bold');
  sh.setColumnWidth(1, 320);
  sh.hideSheet();
  return sh;
}

/** Loads the period marks of one attendance sheet: { book, sh, row, d: { day: { period: { roll: mark } } } } */
function pdLoad_(book, name) {
  var pd = { book: book, sh: book.getSheetByName(PD_SHEET), row: 0, d: {} };
  if (!pd.sh || pd.sh.getLastRow() < 2) return pd;
  var f = pd.sh.getRange(2, 1, pd.sh.getLastRow() - 1, 1).createTextFinder(name)
    .matchCase(true).matchEntireCell(true).findNext();
  if (!f) return pd;
  pd.row = f.getRow();
  var raw;
  try { raw = JSON.parse(pd.sh.getRange(pd.row, 2).getValue() || '{}'); } catch (e) { raw = {}; }
  var rolls = raw.r || [];
  Object.keys(raw.d || {}).forEach(function (day) {
    pd.d[day] = {};
    Object.keys(raw.d[day]).forEach(function (p) {
      var s = String(raw.d[day][p]), m = {};
      for (var i = 0; i < rolls.length; i++) { var c = s.charAt(i); if (VALID[c]) m[rolls[i]] = c; }
      pd.d[day][p] = m;
    });
  });
  return pd;
}

/** Saves the period marks (compact: one string per period, one character per student of the stored roll list). */
function pdSave_(pd, name) {
  var rolls = [], seen = {};
  Object.keys(pd.d).forEach(function (day) {
    Object.keys(pd.d[day]).forEach(function (p) {
      Object.keys(pd.d[day][p]).forEach(function (r) { if (!seen[r]) { seen[r] = 1; rolls.push(r); } });
    });
  });
  var d = {};
  Object.keys(pd.d).forEach(function (day) {
    d[day] = {};
    Object.keys(pd.d[day]).forEach(function (p) {
      d[day][p] = rolls.map(function (r) { return pd.d[day][p][r] || '-'; }).join('');
    });
  });
  var json = JSON.stringify({ r: rolls, d: d });
  if (json.length > 45000) throw new Error('Period data of ' + name + ' is too large.');
  if (!pd.sh) pd.sh = pdSheet_(pd.book);
  if (!pd.row) pd.row = pd.sh.getLastRow() + 1;
  pd.sh.getRange(pd.row, 1, 1, 3).setValues([[name, json, new Date()]]);
}

/** Old single-letter cells of a day -> { 1: { roll: mark } } (old days become "period 1"). */
function legacyDay_(I, col) {
  var vals = I.sheet.getRange(FIRST_ROW, col, I.n, 1).getValues(), m = {}, any = false;
  for (var i = 0; i < I.n; i++) {
    var v = String(vals[i][0]).trim().toUpperCase();
    if (I.rolls[i] && VALID[v]) { m[I.rolls[i]] = v; any = true; }
  }
  return any ? { 1: m } : {};
}

/** Set of rolls that have a mark in any period of this day. */
function rollsOf_(dayPm) {
  var s = {};
  Object.keys(dayPm || {}).forEach(function (p) {
    Object.keys(dayPm[p]).forEach(function (r) { s[r] = true; });
  });
  return s;
}

/** The text of the day cell: P, 2P, 3P ... or (no P) A / 3A, L / 3L, B / 3B. */
function cellFor_(dayPm, roll) {
  var ps = Object.keys(dayPm || {}), held = ps.length, p = 0, seen = {};
  ps.forEach(function (k) {
    var m = dayPm[k][roll];
    if (!m) return;
    if (m === 'P') p++; else seen[m] = true;
  });
  if (p) return p === 1 ? 'P' : p + 'P';
  for (var i = 0; i < NONP.length; i++) {
    if (seen[NONP[i]]) return held === 1 ? NONP[i] : held + NONP[i];
  }
  return '';
}

/**
 * Rewrites the day column (only students that have, or had, period data; disabled students are never written)
 * and the "periods held" number in row 1, then reads the cells back to verify.
 * rollSet = rolls that had data BEFORE this change (so removed marks are cleared too).
 */
function writeDay_(I, pd, day, col, rollSet) {
  var pm = pd.d[String(day)] || {}, held = Object.keys(pm).length;
  var now = rollsOf_(pm);
  var rng = I.sheet.getRange(FIRST_ROW, col, I.n, 1), vals = rng.getValues(), want = {};
  for (var i = 0; i < I.n; i++) {
    var roll = I.rolls[i];
    if (!roll || I.off[i] || !(rollSet[roll] || now[roll])) continue;
    var v = cellFor_(pm, roll);
    vals[i][0] = v; want[i] = v;
  }
  rng.setValues(vals);
  I.sheet.getRange(1, col).setValue(held || '').setFontSize(8).setFontWeight('bold').setHorizontalAlignment('center');
  SpreadsheetApp.flush();
  var back = rng.getValues(), bad = 0;
  Object.keys(want).forEach(function (k) {
    if (String(back[k][0]).trim().toUpperCase() !== want[k]) bad++;
  });
  if (bad) throw new Error('Verification failed: ' + bad + ' cell(s) in ' + I.sheet.getName() + ' day ' + day + ' do not match. Tap Sync now to retry.');
}

/** Creates the monthly sheet (same layout every time) if it does not exist yet.
 *  `book` must be the workbook of meta.mon (use bookFor_(meta.mon)). */
function ensureSheet_(book, meta) {
  var sh = book.getSheetByName(meta.name);
  if (sh) return sh;

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    sh = book.getSheetByName(meta.name);
    if (sh) return sh;

    var roster = getRoster_(book, meta);
    if (!roster.length) {
      throw new Error('No students found for ' + meta.cls + ' / ' + meta.sub +
        '. Check the section and subject columns in Transform-ALL-COLLEGE-STRENGTH.');
    }

    var my = monthYear_(meta.mon);
        var gIdx = {};
    activeGuests_(meta.cls, meta.sub, meta.mon).forEach(function (g) {
      if (!roster.some(function (r) { return r[0] === g.roll; })) { gIdx[roster.length] = g; roster.push([g.roll, g.name, g.father, g.home]); }
    });
    var n = roster.length;
    var totalCols = MT_COL;

    sh = book.insertSheet(meta.name);

    // rows 1-2: grey banner
    sh.getRange(1, 1, 2, totalCols).setBackground('#CFCCCC');
    sh.getRange(1, 1).setValue(TITLE_PREFIX + titleCase_(MONTH_NAMES[my.m]) + ' ' + my.y)
      .setFontWeight('bold').setFontSize(11);
    stampTeacher_(sh, meta);
    sh.getRange(1, TA_COL).setValue('T.L').setFontSize(8).setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange(2, TA_COL).setValue(0).setFontWeight('bold').setHorizontalAlignment('center');
    setWeekdayRow_(sh, my.m, my.y);

    // row 3: headers
    var headers = ['Sr. No', 'Roll No', 'Section', 'StudentName', 'FatherName', 'Class-Sec'];
    for (var d = 1; d <= 31; d++) headers.push(d);
    headers.push('T.A', 'M.T');
    sh.getRange(HEADER_ROW, 1, 1, totalCols).setValues([headers])
      .setFontWeight('bold').setBackground('#D9D9D9').setHorizontalAlignment('center').setVerticalAlignment('middle');

    // students
    sh.getRange(FIRST_ROW, 2, n, 1).setNumberFormat('@');
    var rows = roster.map(function (r, i) { return [i + 1, r[0], sectionOf_(r[3]), r[1], r[2] || '', r[3] || baseClass_(meta.cls)]; });
        sh.getRange(FIRST_ROW, 1, n, 6).setValues(rows);
        Object.keys(gIdx).forEach(function (i) { var g = gIdx[i]; sh.getRange(FIRST_ROW + Number(i), 2).setNote(guestNote_(g.home, g.by, g.on)); });

    // T.A = number of lectures attended: P = 1, 2P = 2, 3P = 3 ... (A, L, B are NOT counted)
    var ta = [];
    for (var r = 0; r < n; r++) ta.push([taFormula_(FIRST_ROW + r)]);
    sh.getRange(FIRST_ROW, TA_COL, n, 1).setFormulas(ta).setHorizontalAlignment('center');

    // alignment
    sh.getRange(FIRST_ROW, 1, n, 3).setHorizontalAlignment('center');
    sh.getRange(FIRST_ROW, 4, n, 1).setHorizontalAlignment('left');
    sh.getRange(FIRST_ROW, 6, n, 1).setHorizontalAlignment('center');
    var days = sh.getRange(FIRST_ROW, DAY_COL1, n, 31);
    days.setHorizontalAlignment('center');
    sh.getRange(FIRST_ROW, MT_COL, n, 1).setHorizontalAlignment('center');

    // day column shading: Sunday black, Saturday pink, days not in this month grey
    var dim = new Date(my.y, my.m + 1, 0).getDate();
    for (var dd = 1; dd <= 31; dd++) {
      var bg = null;
      if (dd > dim) bg = '#B7B7B7';
      else {
        var wd = new Date(my.y, my.m, dd).getDay();
        if (wd === 0) bg = '#000000'; else if (wd === 6) bg = '#EAD1DC';
      }
      if (bg) sh.getRange(FIRST_ROW, DAY_COL1 + dd - 1, n, 1).setBackground(bg);
    }

    // A = red, L = amber, B = blue; low attendance (<73%) and low monthly test (<50%) highlighted
    applyRules_(sh, n);

    // borders, filter, freeze, widths
    sh.getRange(HEADER_ROW, 1, n + 1, totalCols)
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(HEADER_ROW, 1, n + 1, totalCols).createFilter();
    sh.setFrozenRows(HEADER_ROW);
    sh.setFrozenColumns(4);
    sh.setColumnWidth(1, 50);
    sh.setColumnWidth(2, 115);
    sh.setColumnWidth(3, 65);
    sh.setColumnWidth(4, 190);
    sh.setColumnWidth(5, 150);
    sh.setColumnWidth(6, 100);
    sh.setColumnWidths(DAY_COL1, 31, 34);
    sh.setColumnWidth(TA_COL, 50);
    sh.setColumnWidth(MT_COL, 50);

    // keep the workbook small: drop the empty rows/columns a new sheet comes with
    try { trimSheet_(sh, FIRST_ROW + n - 1, totalCols); } catch (e) { Logger.log('Trim skipped for ' + meta.name + ': ' + e.message); }

    SpreadsheetApp.flush();
    return sh;
  } finally {
    lock.releaseLock();
  }
}

/* ---------- subject / section changes: disable old row, add new admission ---------- */

/** Disables students who no longer take this subject/section, re-enables those who came back.
 *  Returns { off: n, on: n }. Only disables when allowOff is true (current month). */
function syncSubjectChanges_(book, sh, meta, allowOff) {
  var res = { off: 0, on: 0 };
  var M = getMaster_();
  if (!M.hasClass || !M.rows.length) return res;
  var roster = getRoster_(book, meta);
  if (!roster.length) return res;

  var inRoster = {}, inMaster = {};
  roster.forEach(function (r) { inRoster[String(r[0]).trim()] = true; });
  M.rows.forEach(function (r) { inMaster[String(r[0]).trim()] = true; });

  var L = readLayout_(sh);
  var total = sh.getLastRow() - FIRST_ROW + 1;
  if (total < 1) return res;
  var rg = sh.getRange(FIRST_ROW, L.rollCol, total, 1);
  var rolls = rg.getValues(), notes = rg.getNotes();

  var toOff = [], toOn = [], active = 0;
  for (var i = 0; i < total; i++) {
    var roll = String(rolls[i][0]).trim();
    if (!roll) continue;
    var off = isOffNote_(notes[i][0]);
    if (inRoster[roll]) { if (off) toOn.push(i); else active++; }
    else if (!off) { active++; if (inMaster[roll]) toOff.push(i); }   // still a student, but not in this subject
  }

  if (allowOff && toOff.length > 3 && toOff.length > active * 0.4) {
    Logger.log('SKIPPED disabling in ' + meta.name + ': ' + toOff.length + ' of ' + active +
      ' students would be disabled. Check the subject name / Transform sheet.');
    toOff = [];
  }
  if (!allowOff) toOff = [];

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMM yyyy');
  toOff.forEach(function (i) {
    var row = FIRST_ROW + i;
    sh.getRange(row, L.rollCol).setNote(OFF_TAG + ' - changed subject' );
    sh.getRange(row, 1, 1, 6).setBackground('#EFEFEF').setFontColor('#999999');
    sh.getRange(row, 4).setFontLine('line-through');
    res.off++;
  });
  toOn.forEach(function (i) {
    var row = FIRST_ROW + i;
    sh.getRange(row, L.rollCol).clearNote();
    sh.getRange(row, 1, 1, 6).setBackground(null).setFontColor('#000000');
    sh.getRange(row, 4).setFontLine('none');
    res.on++;
  });
  if (res.off || res.on) SpreadsheetApp.flush();
  return res;
}



/* ---------- read / save ---------- */

/** Finds Roll No / Name / M.T columns and the column of each day (1..31) from the header row. */
function readLayout_(sheet) {
  var lastCol = sheet.getLastColumn();
  return layoutOf_(sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0], sheet.getName());
}
/** same as readLayout_, from a header row that was already read */
function layoutOf_(hdr, sheetName) {

  var L = { rollCol: -1, nameCol: -1, mtCol: -1, dayCols: {} };
  for (var c = 0; c < hdr.length; c++) {
    var t = norm_(hdr[c]);
    if (t === 'ROLLNO' && L.rollCol < 0) L.rollCol = c + 1;
    else if (t === 'STUDENTNAME' && L.nameCol < 0) L.nameCol = c + 1;
    else if ((t === 'M.T' || t === 'MT') && L.mtCol < 0) L.mtCol = c + 1;
    else if (t !== '' && /^\d+$/.test(t)) {
      var d = Number(t);
      if (d >= 1 && d <= 31 && !L.dayCols[d]) L.dayCols[d] = c + 1;
    }
  }
  if (L.rollCol < 0) throw new Error('Roll No header not found in ' + sheetName);
  return L;
}

function dayColRange_(L) {
  var cols = Object.keys(L.dayCols).map(function (d) { return L.dayCols[d]; });
  if (!cols.length) return null;
  return { c1: Math.min.apply(null, cols), c2: Math.max.apply(null, cols) };
}

/**
 * Returns the sheet for the app. marks has the shape marks[day][period][roll] = 'P'|'A'|'L'|'B'.
 * Days saved before period-wise attendance existed (single letters in the cells) are returned as period 1.
 */
function getSheetData_(user, name, create) {
  var meta = allowed_(user)[name];
  if (!meta) return { ok: false, error: 'NOT_ALLOWED: ' + name };

  var book = bookFor_(meta.mon);          // the workbook of the SELECTED month
  var sheet = book.getSheetByName(name);
  if (!sheet) {
    if (!create) return { ok: true, missing: true, sheet: name };
    sheet = ensureSheet_(book, meta);
  }

  // v7.4: the whole sheet in ONE read, all notes in ONE read (was about ten reads)
  var lastR = Math.max(sheet.getLastRow(), HEADER_ROW), lastC = sheet.getLastColumn();
  var V = sheet.getRange(1, 1, lastR, lastC).getValues();
  var NT = sheet.getRange(HEADER_ROW, 1, lastR - HEADER_ROW + 1, lastC).getNotes();     // row 0 = header notes, then the students
  var L = layoutOf_(V[HEADER_ROW - 1], name);
  var n = lastR - FIRST_ROW + 1;
  var students = [], marks = {}, notes = {}, rolls = [];
  var days = Object.keys(L.dayCols).map(Number).sort(function (a, b) { return a - b; });
  var rg = dayColRange_(L);
      var pd = pdLoad_(book, name);
  var guestAt = guestAtMap_(meta);
  var colOf = function (c) { var out = []; for (var r = 0; r < n; r++) out.push([V[FIRST_ROW - 1 + r][c - 1]]); return out; };

  if (n > 0) {
    rolls = colOf(L.rollCol);
    var rollNotes = []; for (var q0 = 0; q0 < n; q0++) rollNotes.push([NT[q0 + 1][L.rollCol - 1]]);   // a LEFT SUBJECT note = disabled student
    var names = L.nameCol > 0 ? colOf(L.nameCol) : null;
    var rollList = [];
    for (var i = 0; i < n; i++) {
      var roll = String(rolls[i][0]).trim();
      if (!roll) continue;
      rollList.push(i);
          var st = { roll: roll, name: names ? String(names[i][0]).trim() : roll };
      var gnote = String(rollNotes[i][0] || '');
      if (isOffNote_(gnote)) { st.off = true; if (gnote.indexOf('guest removed') !== -1) st.offGuest = true; if (gnote.indexOf(LEFT_TAG) === 0) st.leftAs = (gnote.match(/-\s*(.*?) on the portal/) || [])[1] || 'Left college'; }
      if (gnote.indexOf('GUEST|') === 0) { var gp = gnote.split('|'); st.guest = { home: gp[1] || '', by: gp[2] || '', on: gp[3] || '' }; }
      if (guestAt[roll]) st.guestAt = guestAt[roll];
      students.push(st);
    }
    if (rg) {
      // the day columns of the whole month, from the one read above
      var block = []; for (var b0 = 0; b0 < n; b0++) block.push(V[FIRST_ROW - 1 + b0].slice(rg.c1 - 1, rg.c2));
      days.forEach(function (d) {
        if (pd.d[d] && Object.keys(pd.d[d]).length) { marks[d] = pd.d[d]; return; }   // period-wise day
        var j = L.dayCols[d] - rg.c1, m = {}, any = false;                              // old day: single letters = period 1
        rollList.forEach(function (i) {
          var v = String(block[i][j]).trim().toUpperCase();
          if (VALID[v]) { m[String(rolls[i][0]).trim()] = v; any = true; }
        });
        if (any) marks[d] = { 1: m };
      });
    }
  }
  if (rg) {
    var nr = NT[0].slice(rg.c1 - 1, rg.c2);
    days.forEach(function (d) { var t = nr[L.dayCols[d] - rg.c1]; if (t) notes[d] = t; });
  }

  var mt = { total: 0, marks: {} };
  if (L.mtCol > 0 && n > 0) {
    mt.total = Number(V[1][L.mtCol - 1]) || 0;
    var mv = colOf(L.mtCol);
    for (var q = 0; q < n; q++) {
      var rr = String(rolls[q][0]).trim(), mval = String(mv[q][0]).trim();
      if (rr && mval !== '') mt.marks[rr] = mval;
    }
  }

   try { mt = mtFromBook_(meta, mt); } catch (e) { Logger.log('Monthly test workbook skipped for ' + name + ': ' + e.message); }
  var my = monthYear_(meta.mon);
  var title = String(V[0][0] || '').trim();
  var lv = {};
  try { lv = leavesForSheet_(meta, students); } catch (e) { Logger.log('Approved leaves skipped for ' + name + ': ' + e.message); }
  return { ok: true, sheet: name, title: title, m: my.m, y: my.y, days: days, students: students, marks: marks, notes: notes, mt: mt, lv: lv };
}
 

/**
 * Total lectures held in the month, written in row 2 after the last day column (T.L).
 * A day counts as many lectures as the number in row 1 above its date (periods held).
 * Old days without that number count as 1 lecture when somebody has P or A. L and B alone do not count.
 */
function updateLectures_(sheet, L, n) {
  var rg = dayColRange_(L);
  if (!rg || n < 1) return;
  var width = rg.c2 - rg.c1 + 1;
  var held = sheet.getRange(1, rg.c1, 1, width).getValues()[0];
  var block = sheet.getRange(FIRST_ROW, rg.c1, n, width).getValues();
  var rolls = sheet.getRange(FIRST_ROW, L.rollCol, n, 1).getValues();
  var count = 0;
  Object.keys(L.dayCols).forEach(function (d) {
    var j = L.dayCols[d] - rg.c1, h = Number(held[j]) || 0;
    if (!h) {
      for (var i = 0; i < n; i++) {
        if (!String(rolls[i][0]).trim()) continue;          // ignore rows with no student
        var v = String(block[i][j]).trim().toUpperCase();
        if (v === 'P' || v === 'A') { h = 1; break; }
      }
    }
    count += h;
  });
  sheet.getRange(2, rg.c2 + 1).setValue(count);
}

function recalcLecturesAllSheets() {
  batchSheets_('recalcLecturesAllSheets', function (sh) {
    var L = readLayout_(sh);
    var n = sh.getLastRow() - FIRST_ROW + 1;
    if (n < 1) return;
    updateLectures_(sh, L, n);
  });
}

/** Puts the note on the date header of that day (yellow header). Empty text clears it. */
function setDayNote_(sheet, L, day, text) {
  var col = L.dayCols[Number(day)];
  if (!col) throw new Error('Day ' + day + ' column not found');
  var cell = sheet.getRange(HEADER_ROW, col);
  text = String(text == null ? '' : text).trim();
  if (text) cell.setNote(text).setBackground('#FFD966');
  else cell.clearNote().setBackground('#D9D9D9');
}

/** Monthly test: total in row 2 of the M.T column, each student's mark (number or A) in that column.
 *  Disabled students (I.off) are never written. */
function saveMt_(I, rec) {
  if (!I.L.mtCol || I.L.mtCol < 0) throw new Error('M.T column not found in ' + rec.sheet);
  var total = Number(rec.total);
  if (!(total > 0)) throw new Error('Total marks must be more than 0.');
  var marks = rec.marks || {};
  var rng = I.sheet.getRange(FIRST_ROW, I.L.mtCol, I.n, 1), vals = rng.getValues(), want = {}, wrote = 0;
  for (var i = 0; i < I.n; i++) {
    var roll = I.rolls[i];
    if (!roll || I.off[i] || !marks.hasOwnProperty(roll)) continue;
    var s = String(marks[roll]).trim().toUpperCase(), v;
    if (s === '') v = '';
    else if (s === 'A' || s === 'AB' || s === 'ABSENT') v = 'A';
    else {
      v = Number(s);
      if (isNaN(v) || v < 0 || v > total) throw new Error('Marks for ' + roll + ' must be between 0 and ' + total + ' (or A).');
    }
    var was = String(vals[i][0] == null ? '' : vals[i][0]).trim();       // v6.7: change history
    if (I.user && was !== '' && was.toUpperCase() !== String(v).toUpperCase()) auditAdd_(I.user, I.meta, '', '', roll, I.names ? I.names[i] : '', was, String(v), 'Monthly test');
    vals[i][0] = v; want[i] = v; wrote++;
  }
  rng.setValues(vals);
  I.sheet.getRange(1, I.L.mtCol).setValue('MT TOTAL').setFontSize(7).setFontWeight('bold').setHorizontalAlignment('center');
  I.sheet.getRange(2, I.L.mtCol).setValue(total).setFontWeight('bold').setHorizontalAlignment('center');
  SpreadsheetApp.flush();
  var back = I.sheet.getRange(FIRST_ROW, I.L.mtCol, I.n, 1).getValues(), bad = 0;
  Object.keys(want).forEach(function (k) {
    if (String(back[k][0]).trim().toUpperCase() !== String(want[k]).toUpperCase()) bad++;
  });
  if (bad) throw new Error('Verification failed: ' + bad + ' mark(s) in ' + rec.sheet + ' do not match. Tap Sync now to retry.');
     mtToBook_(I, rec, total);
  return { at: Date.now(), mt: true, wrote: wrote, total: total };
}

/**
 * records: [{ id, sheet, day, period, marks: { 'JCW-F26-2301': 'P'|'A'|'L'|'B'|'' }, note?: 'text' }]
 *   or monthly test: { id, kind:'mt', sheet, total, marks:{roll:'12'|'A'|''} }
 * period = 1..MAX_PERIOD (missing = 1). Only the marks of THAT period are replaced; the other periods of the
 * day stay. The day cell (P / 2P / 3A ...) is then rebuilt from all periods of the day.
 * Saving the same sheet/day/period again simply overwrites, so a re-sent record is harmless.
 * Every write is read back; a record is only reported as saved if the sheet really holds the values.
 * Each record is saved in the workbook of ITS month. Disabled students are never written.
 */
function saveRecords_(user, records) {
  var okNames = allowed_(user);
  var saved = [], failed = [], confirmed = {}, info = {}, touched = {}, createFail = {};

  // create missing sheets first, each in the workbook of its own month (ensureSheet_ takes the lock itself)
  records.forEach(function (rec) {
    if (!okNames[rec.sheet] || createFail[rec.sheet]) return;
    try {
      var bk = bookFor_(okNames[rec.sheet].mon);
      if (bk.getSheetByName(rec.sheet)) return;
      ensureSheet_(bk, okNames[rec.sheet]);
    }
    catch (err) { createFail[rec.sheet] = String(err.message || err); }
  });

  var lock = LockService.getScriptLock();
  lock.waitLock(60000);                                                // v7.4: wait up to 60 s for the teacher before (was 30 s)
  var locked = true;
  try {
    records.forEach(function (rec) {
      try {
        if (!okNames[rec.sheet]) throw new Error('Not allowed: ' + rec.sheet);
        if (createFail[rec.sheet]) throw new Error(createFail[rec.sheet]);
        var I = info[rec.sheet];
        if (!I) {
          var sheet = bookFor_(okNames[rec.sheet].mon).getSheetByName(rec.sheet);
          if (!sheet) throw new Error('Sheet not found: ' + rec.sheet);
          var L = readLayout_(sheet);
          var n = sheet.getLastRow() - FIRST_ROW + 1;
          if (n < 1) throw new Error('No students in ' + rec.sheet);
          var rollRange = sheet.getRange(FIRST_ROW, L.rollCol, n, 1);
          var rl = rollRange.getValues().map(function (r) { return String(r[0]).trim(); });
          var offs = rollRange.getNotes().map(function (r) { return isOffNote_(r[0]); });
          var idx = {};
          rl.forEach(function (r, i) { if (r) idx[r] = i; });
          var nmv = L.nameCol > 0 ? sheet.getRange(FIRST_ROW, L.nameCol, n, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : rl;
          I = info[rec.sheet] = { sheet: sheet, L: L, n: n, rolls: rl, off: offs, idx: idx, pd: null, names: nmv };
        }

        if (rec.kind === 'mt') {
          I.user = user; I.meta = okNames[rec.sheet];                     // v6.7: for the change history
          var cm = saveMt_(I, rec);
          touched[rec.sheet] = true; confirmed[rec.id] = cm; saved.push(rec.id);
          return;
        }

        var col = I.L.dayCols[Number(rec.day)];
        if (!col) throw new Error('Day ' + rec.day + ' column not found in ' + rec.sheet);

        var c = { at: Date.now() };
        var marks = rec.marks || {};
        var hasMarks = Object.keys(marks).length > 0;

        if (hasMarks) {
          var per = Number(rec.period) || 1;
          if (per < 1 || per > MAX_PERIOD) throw new Error('Period must be 1 to ' + MAX_PERIOD + '.');
          if (!I.pd) I.pd = pdLoad_(bookFor_(okNames[rec.sheet].mon), rec.sheet);
          var pd = I.pd, dk = String(Number(rec.day));
          if (!pd.d[dk]) pd.d[dk] = legacyDay_(I, col);          // old day with single letters = period 1
          var before = rollsOf_(pd.d[dk]);
          var pm = pd.d[dk][per] = pd.d[dk][per] || {};
          var old = {}; Object.keys(pm).forEach(function (r) { old[r] = pm[r]; });   // v6.7: for the change history
          var wrote = 0;
          Object.keys(marks).forEach(function (roll) {
            if (I.idx[roll] == null || I.off[I.idx[roll]]) return;     // unknown or disabled student: never written
            var m = String(marks[roll]).trim().toUpperCase();
            if (VALID[m]) { pm[roll] = m; wrote++; }
            else if (m === '') delete pm[roll];
          });
          if (!Object.keys(pm).length) delete pd.d[dk][per];
          if (!Object.keys(pd.d[dk]).length) delete pd.d[dk];
          pdSave_(pd, rec.sheet);
          writeDay_(I, pd, rec.day, col, before);
          Object.keys(marks).forEach(function (roll) {                  // v6.7: only CHANGES of a saved mark are logged
            var was = old[roll] || '', now = pm[roll] || '';
            if (was && was !== now && I.idx[roll] != null) auditAdd_(user, okNames[rec.sheet], rec.day, per, roll, I.names[I.idx[roll]], was, now, 'Attendance');
          });

          // counts of THIS period (active students only)
          c.p = 0; c.a = 0; c.l = 0; c.b = 0;
          var nowPm = (pd.d[dk] && pd.d[dk][per]) || {};
          Object.keys(nowPm).forEach(function (r) {
            if (I.idx[r] == null || I.off[I.idx[r]]) return;
            var v = nowPm[r];
            if (v === 'P') c.p++; else if (v === 'A') c.a++; else if (v === 'L') c.l++; else if (v === 'B') c.b++;
          });
          c.wrote = wrote;
          c.period = per;
        }

        if (rec.note !== undefined && rec.note !== null) {
          setDayNote_(I.sheet, I.L, rec.day, rec.note);
          var got = String(I.sheet.getRange(HEADER_ROW, col).getNote() || '').trim();
          if (got !== String(rec.note).trim()) throw new Error('Note could not be verified in ' + rec.sheet);
          c.note = true;
        }

        touched[rec.sheet] = true;
        confirmed[rec.id] = c;
        saved.push(rec.id);
      } catch (err) {
        failed.push({ id: rec.id, error: String(err.message || err) });
      }
    });

    Object.keys(touched).forEach(function (name) {
      try { stampTeacher_(info[name].sheet, okNames[name]); } catch (e) {}
    });
    try { auditFlush_(); } catch (e) { Logger.log('Change history skipped: ' + e.message); }   // v6.7
    SpreadsheetApp.flush();
    lock.releaseLock(); locked = false;                                // v7.4: the next teacher can save now

    // after the lock: lecture count (T.L) and the Attendance Index row of each sheet saved.
    // Both are worked out again from the sheet itself, so a save that runs at the same time cannot spoil them.
    Object.keys(touched).forEach(function (name) {
      try { var I = info[name]; updateLectures_(I.sheet, I.L, I.n); } catch (e) {}
      try { indexUpdate_(info[name].sheet, okNames[name]); } catch (e) { Logger.log('Index skipped for ' + name + ': ' + e.message); }
    });
    SpreadsheetApp.flush();
    return { ok: true, saved: saved, failed: failed, confirmed: confirmed };
  } finally {
    if (locked) lock.releaseLock();
  }
}

/* ---------- student remarks (tags + comments) ---------- */
/* Remarks are kept in ONE place: the SEPT workbook (BOOK_ID). */

function remarksSheet_() {
  var book = bookFor_('SEPT');
  var sh = book.getSheetByName(REMARKS_SHEET);
  if (sh) return sh;
  sh = book.insertSheet(REMARKS_SHEET);
  sh.getRange(1, 1, 1, REMARK_HEAD.length).setValues([REMARK_HEAD])
    .setFontWeight('bold').setBackground('#D9D9D9');
  sh.getRange(2, 1, 1000, 4).setNumberFormat('@');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 140); sh.setColumnWidth(3, 90);
  sh.setColumnWidth(4, 120); sh.setColumnWidth(5, 180); sh.setColumnWidth(6, 200);
  sh.setColumnWidth(7, 320); sh.setColumnWidth(8, 80);
  sh.hideColumns(1); sh.hideColumns(9); sh.hideColumns(10);
  return sh;
}

function remarkKey_(user, cls, roll) {
  return String(user.username).trim().toLowerCase() + '|' + cls + '|' + roll;
}

/** Everything this teacher has written: { 'FCS-1 D2|JCW-F26-2301': {tags:[], log:[{t,txt}], u:ms} } */
function getRemarks_(user) {
  var sh = remarksSheet_(), out = {};
  var n = sh.getLastRow() - 1;
  if (n < 1) return { ok: true, remarks: out };
  var prefix = String(user.username).trim().toLowerCase() + '|';
  var v = sh.getRange(2, 1, n, REMARK_HEAD.length).getValues();
  v.forEach(function (r) {
    var key = String(r[0]);
    if (key.indexOf(prefix) !== 0) return;
    var log = [];
    try { log = JSON.parse(r[8] || '[]'); } catch (e) {}
    out[String(r[2]) + '|' + String(r[3])] = {
      tags: String(r[5] || '').split(',').map(function (t) { return t.trim(); }).filter(Boolean),
      log: log,
      u: Number(r[9]) || 0
    };
  });
  return { ok: true, remarks: out };
}

/**
 * records: [{ cls, roll, name, tags:[...], log:[{t,txt}], u }]
 * Newest edit wins (compared by u). An empty tags+log record clears that student's remarks.
 * Returns saved:[ 'cls|roll' ] and skipped:[ 'cls|roll' ] (server already had a newer edit).
 */
function saveRemarks_(user, records) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = remarksSheet_(), saved = [], skipped = [], appendRows = [];
    var n = sh.getLastRow() - 1, rowOf = {}, uOf = {};
    if (n > 0) {
      var keys = sh.getRange(2, 1, n, 1).getValues(), us = sh.getRange(2, 10, n, 1).getValues();
      for (var i = 0; i < n; i++) { rowOf[String(keys[i][0])] = i + 2; uOf[String(keys[i][0])] = Number(us[i][0]) || 0; }
    }
    records.forEach(function (rec) {
      var cls = String(rec.cls || '').trim(), roll = String(rec.roll || '').trim();
      if (!cls || !roll) return;
      var id = cls + '|' + roll, key = remarkKey_(user, cls, roll), u = Number(rec.u) || Date.now();
      if (rowOf[key] && uOf[key] > u) { skipped.push(id); return; }

      var tags = (Array.isArray(rec.tags) ? rec.tags : []).map(function (t) { return String(t).trim().slice(0, 40); })
        .filter(Boolean).slice(0, REMARK_MAX_TAGS);
      var log = (Array.isArray(rec.log) ? rec.log : []).map(function (e) {
        return { t: Number(e.t) || Date.now(), txt: String(e.txt || '').trim().slice(0, 400) };
      }).filter(function (e) { return e.txt; }).slice(-REMARK_MAX_LOG);
      var last = log.length ? log[log.length - 1] : null;
      var row = [key, user.teacher, cls, roll, String(rec.name || ''), tags.join(', '),
                 last ? Utilities.formatDate(new Date(last.t), Session.getScriptTimeZone(), 'd MMM yyyy') + ': ' + last.txt : '',
                 log.length, JSON.stringify(log), u,
                 Utilities.formatDate(new Date(u), Session.getScriptTimeZone(), 'd MMM yyyy HH:mm')];
      if (rowOf[key]) sh.getRange(rowOf[key], 1, 1, REMARK_HEAD.length).setValues([row]);
      else { appendRows.push(row); rowOf[key] = -1; }
      uOf[key] = u;
      saved.push(id);
    });
    if (appendRows.length) {
      var start = sh.getLastRow() + 1;
      sh.getRange(start, 1, appendRows.length, 4).setNumberFormat('@');
      sh.getRange(start, 1, appendRows.length, REMARK_HEAD.length).setValues(appendRows);
    }
    SpreadsheetApp.flush();
    return { ok: true, saved: saved, skipped: skipped };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- delete attendance of one period (or of the whole day) ---------- */

/**
 * period given  -> only that period of the day is removed (the other periods stay and the day cell is rebuilt).
 * period empty  -> the whole day is removed (all periods).
 * When no period is left on the day, the day column, the "periods held" number and the day note are cleared.
 */
function deleteDay_(user, name, day, confirmDay, period) {
  day = Number(day);
  if (!(day >= 1 && day <= 31) || Number(confirmDay) !== day) return { ok: false, error: 'Day confirmation does not match.' };
  var meta = allowed_(user)[name];
  if (!meta) return { ok: false, error: 'NOT_ALLOWED: ' + name };
  var book = bookFor_(meta.mon), sheet = book.getSheetByName(name);
  if (!sheet) return { ok: false, error: 'This sheet does not exist yet.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var L = readLayout_(sheet), col = L.dayCols[day];
    if (!col) throw new Error('Day ' + day + ' column not found in ' + name);
    var n = sheet.getLastRow() - FIRST_ROW + 1;
    if (n < 1) throw new Error('No students in ' + name);

    var rr = sheet.getRange(FIRST_ROW, L.rollCol, n, 1);
    var I = {
      sheet: sheet, L: L, n: n,
      rolls: rr.getValues().map(function (r) { return String(r[0]).trim(); }),
      off: rr.getNotes().map(function (r) { return isOffNote_(r[0]); })
    };

    var pd = pdLoad_(book, name), dk = String(day), cnt = 0;
    if (!pd.d[dk]) pd.d[dk] = legacyDay_(I, col);
    var before = rollsOf_(pd.d[dk]);
    period = Number(period) || 0;
    try {                                                                // v6.7: change history of the deleted marks
      var nmv = L.nameCol > 0 ? sheet.getRange(FIRST_ROW, L.nameCol, n, 1).getValues() : null;
      Object.keys(pd.d[dk]).forEach(function (p) {
        if (period && Number(p) !== period) return;
        Object.keys(pd.d[dk][p]).forEach(function (r) {
          var i = I.rolls.indexOf(r);
          auditAdd_(user, meta, day, p, r, nmv && i >= 0 ? String(nmv[i][0]).trim() : '', pd.d[dk][p][r], '', 'Deleted');
        });
      });
    } catch (e) { Logger.log('Change history (delete) skipped: ' + e.message); }
    if (period) {
      cnt = Object.keys(pd.d[dk][period] || {}).length;
      delete pd.d[dk][period];
    } else {
      Object.keys(pd.d[dk]).forEach(function (p) { cnt += Object.keys(pd.d[dk][p]).length; });
      pd.d[dk] = {};
    }
    var empty = !Object.keys(pd.d[dk]).length;
    if (empty) delete pd.d[dk];
    pdSave_(pd, name);
    writeDay_(I, pd, day, col, before);
    if (empty) {
      sheet.getRange(FIRST_ROW, col, n, 1).clearContent();
      setDayNote_(sheet, L, day, '');
    }
    SpreadsheetApp.flush();
    updateLectures_(sheet, L, n);
    SpreadsheetApp.flush();
    try { indexUpdate_(sheet, meta); } catch (e) { Logger.log('Index skipped for ' + name + ': ' + e.message); }
    try { auditFlush_(); } catch (e) { Logger.log('Change history skipped: ' + e.message); }   // v6.7
    return { ok: true, count: cnt };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- delete one day's attendance (with backup) ---------- */
var DELLOG_SHEET = 'Deleted Attendance Log';

function logDeletedDay_(user, meta, day, deleted, note) {
  var book = bookFor_('SEPT');
  var sh = book.getSheetByName(DELLOG_SHEET);
  if (!sh) {
    sh = book.insertSheet(DELLOG_SHEET);
    sh.getRange(1, 1, 1, 8).setValues([['Deleted at', 'Teacher', 'Username', 'Sheet', 'Date', 'Marks deleted', 'Note', 'Data (roll: mark)']])
      .setFontWeight('bold').setBackground('#D9D9D9');
    sh.setFrozenRows(1);
  }
  var my = monthYear_(meta.mon);
  sh.appendRow([new Date(), user.teacher, user.username, meta.name,
                day + ' ' + MONTH_NAMES[my.m] + ' ' + my.y, Object.keys(deleted).length, note, JSON.stringify(deleted)]);
}

/* ---------- print: sheet as PDF ---------- */

/* ---------- ink saver: every PDF the server makes is black & white (the college prints on a black & white printer) ---------- */
var INK_SAVER = true;            // false = colour PDFs again
var INK_CSS_ = '<style>*{color:#000!important;background:#fff!important;background-color:#fff!important;box-shadow:none!important;text-shadow:none!important;' +
  '-webkit-print-color-adjust:economy!important}th,thead td{background:#ececec!important;background-color:#ececec!important}' +
  'table,th,td{border-color:#555!important}img{filter:grayscale(1)}</style>';
/** HTML that is turned into a PDF: add the black & white style */
function inkHtml_(html) {
  if (!INK_SAVER) return html;
  html = String(html);
  return /<\/head>/i.test(html) ? html.replace(/<\/head>/i, INK_CSS_ + '</head>') : INK_CSS_ + html;
}

function exportPdf_(user, name) {
  var meta = allowed_(user)[name];
  if (!meta) return { ok: false, error: 'NOT_ALLOWED: ' + name };
  var bookId = bookIdFor_(meta.mon);       // workbook of the selected month
  var book = bookFor_(meta.mon);
  var sh = book.getSheetByName(name);
  if (!sh) return { ok: false, error: 'This sheet does not exist yet.' };

  if (INK_SAVER) return exportPdfInk_(bookId, book, sh, name);          // v7.4.2: black & white copy of the sheet
  // hide notes from the printout: take them off, export, then put every one back
  var rng = sh.getDataRange();
  var notes = rng.getNotes();
  var hasNotes = notes.some(function (r) { return r.some(function (n) { return n; }); });
  if (hasNotes) { rng.clearNote(); }
  SpreadsheetApp.flush();

  var url = 'https://docs.google.com/spreadsheets/d/' + bookId + '/export?format=pdf' +
    '&gid=' + sh.getSheetId() +
    '&size=A4&portrait=false&fitw=true&gridlines=true&printtitle=false&sheetnames=false' +
    '&pagenum=UNDEFINED&shownotes=false&fzr=false&top_margin=0.4&bottom_margin=0.4&left_margin=0.3&right_margin=0.3';

  var resp;
  try {
    resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
  } finally {
    if (hasNotes) { rng.setNotes(notes); SpreadsheetApp.flush(); }   // notes are always restored
  }

  if (resp.getResponseCode() !== 200) {
    return { ok: false, error: 'PDF export failed (' + resp.getResponseCode() + ').' };
  }
  return { ok: true, name: name, pdf: Utilities.base64Encode(resp.getBlob().getBytes()) };
}

/** Black & white PDF of an attendance sheet: a temporary copy without colours, colour rules and notes is exported, then deleted.
 *  The teacher's sheet is never changed. */
function exportPdfInk_(bookId, book, sh, name) {
  var tmp = sh.copyTo(book), resp;
  try {
    tmp.setName('PDF ' + Date.now());
    var rng = tmp.getDataRange();
    tmp.setConditionalFormatRules([]);
    rng.setBackground(null).setFontColor('#000000').clearNote();
    SpreadsheetApp.flush();
    var url = 'https://docs.google.com/spreadsheets/d/' + bookId + '/export?format=pdf' + '&gid=' + tmp.getSheetId() +
      '&size=A4&portrait=false&fitw=true&gridlines=true&printtitle=false&sheetnames=false' +
      '&pagenum=UNDEFINED&shownotes=false&fzr=false&top_margin=0.4&bottom_margin=0.4&left_margin=0.3&right_margin=0.3';
    resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  } finally {
    try { book.deleteSheet(tmp); } catch (e) { Logger.log('Temporary PDF sheet not deleted: ' + e.message); }
  }
  if (resp.getResponseCode() !== 200) return { ok: false, error: 'PDF export failed (' + resp.getResponseCode() + ').' };
  return { ok: true, name: name, pdf: Utilities.base64Encode(resp.getBlob().getBytes()) };
}

/* ---------- bulk jobs that must not hit the 6-minute limit ---------- */

/** Deletes pending one-time continuation triggers of a function (never touches the monthly trigger). */
function clearContinuation_(handler) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === handler) ScriptApp.deleteTrigger(t);
  });
}

/**
 * Runs fn(sheet) on every attendance sheet of ALL five month workbooks. Stops after BATCH_MS, remembers
 * where it stopped and schedules itself again one minute later, until every sheet is done.
 * Returns true when finished.
 * `handler` is the name of the function that calls this (so the continuation runs the same job).
 * Tabs whose A1 does not start with TITLE_PREFIX (Period Data, Student Remarks ...) are skipped.
 */
function batchSheets_(handler, fn, toks) {
  var start = Date.now(), props = PropertiesService.getScriptProperties(), key = 'cursor_' + handler;
  clearContinuation_(handler);
  var sheets = [];
  var books = toks ? toks.map(function (t) { return { tok: t, book: bookFor_(t) }; }) : allBooks_();
  books.forEach(function (b) { b.book.getSheets().forEach(function (s) { sheets.push(s); }); });
  var i = Number(props.getProperty(key)) || 0, done = 0;
  for (; i < sheets.length; i++) {
    if (Date.now() - start > BATCH_MS) {
      props.setProperty(key, String(i));
      ScriptApp.newTrigger(handler).timeBased().after(60 * 1000).create();
      Logger.log(handler + ': ' + done + ' sheet(s) done this round, continuing automatically in 1 minute (at ' + i + ' of ' + sheets.length + ').');
      return false;
    }
    var sh = sheets[i];
    try {
      if (String(sh.getRange(1, 1).getValue()).indexOf(TITLE_PREFIX) !== 0) continue;
      fn(sh);
      done++;
    } catch (e) {
      Logger.log('Skipped ' + sh.getName() + ': ' + e.message);
    }
  }
  props.deleteProperty(key);
  Logger.log(handler + ': finished (' + sheets.length + ' tabs checked).');
  return true;
}

/**
 * Run ONCE by hand for sheets that already exist: removes their empty rows/columns so the workbooks stay far
 * below the 10,000,000 cell limit. Continues by itself until every sheet is done. Safe to run again.
 */
function trimAllSheets() {
  batchSheets_('trimAllSheets', function (sh) {
    var last = sh.getLastRow();
    if (last < HEADER_ROW) return;
    trimSheet_(sh, last, sh.getLastColumn());
  });
}

/** Shows how many of the 10,000,000 cells each month workbook uses (see View > Logs). */
function checkWorkbookCells() {
  var grand = 0;
  allBooks_().forEach(function (b) {
    var sheets = b.book.getSheets(), total = 0;
    sheets.forEach(function (s) { total += s.getMaxRows() * s.getMaxColumns(); });
    Logger.log(b.tok + ': ' + sheets.length + ' tabs, ' + total + ' of 10,000,000 cells used (' + Math.round(total / 100000) + '%).');
    grand += total;
  });
  return grand;
}

/* ---------- create this month's sheets for every teacher ---------- */

/**
 * Creates the current month's sheet for every section/subject in the Timetable (skips existing ones),
 * inside the workbook of the current month.
 * With hundreds of sections it works in rounds of about 4 minutes and continues by itself until done.
 */
function createCurrentMonthSheets() {
  var tok = currentSeasonToken_();
  if (!tok) return;
  clearContinuation_('createMonthSheetsContinue');
  var start = Date.now(), more = false;
  var book = bookFor_(tok);
  var have = {};
  book.getSheets().forEach(function (s) { have[s.getName()] = true; });
  var admin = { teacher: '', username: ADMIN_USER };
  var list = getAssignments_(admin);
  var job = sheetJob_(tok, list, have);                                // v6.9: progress shown in the app
  for (var i = 0; i < list.length; i++) {
    var a = list[i];
    var meta = { name: makeName_(a.cls, tok, a.subject), cls: a.cls, sub: a.subject, mon: tok, teacher: a.teacher };
    if (have[meta.name]) continue;
    if (Date.now() - start > BATCH_MS) { more = true; break; }
    try { ensureSheet_(book, meta); have[meta.name] = true; job.made++; }
    catch (err) { Logger.log('Skipped ' + meta.name + ': ' + err.message); job.failed.push(meta.cls + ' · ' + titleCase_(meta.sub) + ': ' + err.message); }
    sheetJobSave_(job, false);
  }
  sheetJobSave_(job, !more);
  if (more) {
    ScriptApp.newTrigger('createMonthSheetsContinue').timeBased().after(60 * 1000).create();
    Logger.log('createCurrentMonthSheets: continuing automatically in 1 minute.');
  }
}

/** Continuation of createCurrentMonthSheets (scheduled by itself, do not run by hand). */
function createMonthSheetsContinue() {
  createCurrentMonthSheets();
}

/** Run ONCE by hand: creates the sheets automatically on the 1st of every month. */
function setupMonthlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'createCurrentMonthSheets') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('createCurrentMonthSheets').timeBased().onMonthDay(1).atHour(1).create();
  createCurrentMonthSheets();   // also create this month's sheets right now
}

/** Run once by hand: asks Google for permission to open ALL five month workbooks and export PDFs. */
function authorizePdf() {
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
  allBooks_().forEach(function (b) { Logger.log(b.tok + ': ' + b.book.getName()); });
  Logger.log(ScriptApp.getOAuthToken() ? 'OK' : 'no token');
}

/* ---------- one-time maintenance helpers (run by hand if needed; they continue by themselves) ---------- */

/**
 * Run ONCE by hand after installing v5.4: gives every existing attendance sheet the new T.A formula
 * (counts 2P, 3P ...) and the new colour rules. Continues by itself until every sheet is done. Safe to run again.
 * Then run recalcLecturesAllSheets().
 */
function upgradePeriodFormatAllSheets() {
  batchSheets_('upgradePeriodFormatAllSheets', function (sh) {
    var n = sh.getLastRow() - FIRST_ROW + 1; if (n < 1) return;
    var f = [];
    for (var r = 0; r < n; r++) f.push([taFormula_(FIRST_ROW + r)]);
    sh.getRange(FIRST_ROW, TA_COL, n, 1).setFormulas(f);
    applyRules_(sh, n);
  });
}

function addLowAttendanceRule() {
  var L = colLetter_(TA_COL);
  batchSheets_('addLowAttendanceRule', function (sh) {
    var n = sh.getLastRow() - FIRST_ROW + 1; if (n < 1) return;
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($' + L + '$2>0,$' + L + FIRST_ROW + '/$' + L + '$2<0.73)')
      .setBackground('#F4C7C3').setFontColor('#990000')
      .setRanges([sh.getRange(FIRST_ROW, 4, n, 1), sh.getRange(FIRST_ROW, TA_COL, n, 1)]).build();
    var rules = sh.getConditionalFormatRules(); rules.push(rule); sh.setConditionalFormatRules(rules);
  });
}

function addMtLowRule() {
  var C = colLetter_(MT_COL);
  batchSheets_('addMtLowRule', function (sh) {
    var n = sh.getLastRow() - FIRST_ROW + 1; if (n < 1) return;
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($' + C + FIRST_ROW + '),$' + C + '$2>0,$' + C + FIRST_ROW + '/$' + C + '$2<0.5)')
      .setBackground('#F4C7C3').setFontColor('#990000')
      .setRanges([sh.getRange(FIRST_ROW, MT_COL, n, 1)]).build();
    var rules = sh.getConditionalFormatRules(); rules.push(rule); sh.setConditionalFormatRules(rules);
  });
}

/** Updates every existing attendance sheet (removes TOTAL LECTURES label, adds T.L and weekday row). */
function fixHeaderRowsAllSheets() {
  batchSheets_('fixHeaderRowsAllSheets', function (sh) {
    var title = String(sh.getRange(1, 1).getValue() || '');
    if (title.indexOf(TITLE_PREFIX) !== 0) return;
    var mt = title.match(/-\s*([A-Za-z]+)\s+(\d{4})/);
    if (!mt) return;
    var month = MONTH_NAMES.indexOf(mt[1].toUpperCase());
    if (month < 0) return;
    var L = readLayout_(sh), rg = dayColRange_(L);
    if (!rg) return;

    var row2 = sh.getRange(2, 1, 1, Math.max(sh.getLastColumn(), rg.c2 + 1)).getValues()[0];
    row2.forEach(function (v, i) {
      if (/TOTAL\s*LECTURES/i.test(String(v))) sh.getRange(2, i + 1).clearContent();
    });
    sh.getRange(1, rg.c2 + 1).setValue('T.L').setFontSize(8).setFontWeight('bold').setHorizontalAlignment('center');
    setWeekdayRow_(sh, month, Number(mt[2]), L.dayCols[1] || rg.c1);
  });
}
function debugPeriods() {
  var name = 'FCS-1 D2 OCT COMPUTER SCIENCE';
  var m = allowed_({username: ADMIN_USER, teacher: ''})[name];
  var pd = pdLoad_(bookFor_(m.mon), name);
  Logger.log('row=' + pd.row + ' data=' + JSON.stringify(pd.d));
}

/* ---------- guest students ---------- */
var GUESTS_SHEET = 'Guests';
var GUEST_HEAD = ['Roll','Name','Father','Home class','Host class','Subject','Added by','Added on','From month','Status'];

function guestsSheet_() {
  var book = SpreadsheetApp.openById(USERS_SS_ID), sh = book.getSheetByName(GUESTS_SHEET);
  if (sh) return sh;
  sh = book.insertSheet(GUESTS_SHEET);
  sh.getRange(1, 1, 1, GUEST_HEAD.length).setValues([GUEST_HEAD]).setFontWeight('bold').setBackground('#D9D9D9');
  sh.setFrozenRows(1);
  sh.getRange(2, 1, 2000, GUEST_HEAD.length).setNumberFormat('@');
  return sh;
}
function guestRows_() {
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(GUESTS_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, GUEST_HEAD.length).getValues();
}
function guestNote_(home, by, on) { return 'GUEST|' + home + '|' + by + '|' + on; }

/** Active guests of a class+subject that apply to this month (used when a new month sheet is created). */
function activeGuests_(cls, sub, mon) {
  var out = [], mi = SEASON.indexOf(mon);
  guestRows_().forEach(function (r) {
    if (String(r[9]) !== 'Active' || String(r[4]) !== cls || String(r[5]).toUpperCase() !== String(sub).toUpperCase()) return;
    if (SEASON.indexOf(String(r[8])) > mi) return;
    out.push({ roll: String(r[0]).trim(), name: r[1], father: r[2], home: r[3], by: r[6], on: r[7] });
  });
  return out;
}
/** roll -> [{cls, sub, by}] : where else this student is a guest (shown to her own teacher). */
/** "2", "P2", "II", "Period 2", "2nd" -> "2" */
function normPeriod_(p) {
  p = String(p || '').trim().toUpperCase().replace(/^P(ERIOD)?\s*[-.:]?\s*/, '');
  var R = { I: 1, II: 2, III: 3, IV: 4, V: 5, VI: 6, VII: 7, VIII: 8 };
  if (R[p]) return String(R[p]);
  var m = p.match(/\d+/);
  return m ? String(Number(m[0])) : '';
}
var DAYCODE_ = { M: 1, MO: 1, MON: 1, T: 2, TU: 2, TUE: 2, W: 3, WE: 3, WED: 3, TH: 4, THU: 4, THUR: 4, F: 5, FR: 5, FRI: 5, S: 6, SA: 6, SAT: 6 };
var DAYNAME_ = ['', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
/** "M-TH" -> {1,2,3,4}; "F,S" -> {5,6}; "W-S" -> {3..6}; empty -> every day (Mon-Sat). */
function daySet_(txt) {
  var out = {}, any = false;
  String(txt || '').toUpperCase().replace(/\s*-\s*/g, '-').split(/[,;\/&+\s]+/).forEach(function (part) {
    if (!part) return;
    var rg = part.split('-');
    if (rg.length === 2 && DAYCODE_[rg[0]] && DAYCODE_[rg[1]]) {
      for (var d = DAYCODE_[rg[0]]; d <= DAYCODE_[rg[1]]; d++) { out[d] = 1; any = true; }
    } else if (DAYCODE_[part]) { out[DAYCODE_[part]] = 1; any = true; }
  });
  if (!any) for (var k = 1; k <= 6; k++) out[k] = 1;
  return out;
}
/** {1,2,3,4} -> "Mon-Thu"; {5,6} -> "Fri-Sat"; {1,3} -> "Mon, Wed"; all six -> "" */
function dayLabel_(set) {
  var ds = Object.keys(set).map(Number).sort(function (a, b) { return a - b; });
  if (ds.length >= 6) return '';
  var parts = [], i = 0;
  while (i < ds.length) {
    var j = i; while (j + 1 < ds.length && ds[j + 1] === ds[j] + 1) j++;
    parts.push(j - i >= 1 ? DAYNAME_[ds[i]] + '-' + DAYNAME_[ds[j]] : DAYNAME_[ds[i]]);
    i = j + 1;
  }
  return parts.join(', ');
}
/** Timetable slots of one class + subject: [{ p:'2', days:{1:1,..} }] (one per period of each timetable row). */
var SLOTS_MEMO_ = {};
function slotsOf_(cls, sub) {
  var mk = String(cls) + '|' + String(sub).toUpperCase();
  if (SLOTS_MEMO_[mk]) return SLOTS_MEMO_[mk];
  var t = getTimetable_(), out = [];
  var c = String(cls || '').replace(/\s+/g, '').toUpperCase(), s = String(sub || '').trim().toUpperCase();
  t.rows.forEach(function (r) {
    if (String(r[3] || '').replace(/\s+/g, '').toUpperCase() !== c || String(r[4] || '').toUpperCase() !== s) return;
    var days = daySet_([r[6], r[7], r[8]].filter(function (x) { return String(x || '').trim(); }).join(','));
    String(r[5] || '').split(/[,\/&+;]|\band\b/i).forEach(function (x) {
      var rg = String(x).split('-').map(normPeriod_).filter(Boolean);
      if (rg.length === 2 && Number(rg[1]) > Number(rg[0]) && Number(rg[1]) - Number(rg[0]) < 6) {
        for (var n = Number(rg[0]); n <= Number(rg[1]); n++) out.push({ p: String(n), days: days });
      } else rg.forEach(function (n) { out.push({ p: n, days: days }); });
    });
  });
  SLOTS_MEMO_[mk] = out;
  return out;
}
/** Where the two classes meet at the same period on the same day: [{ p, days:{..} }] (empty = no clash, or timetable incomplete). */
function clashSlots_(hostCls, hostSub, myCls, mySub) {
  var a = slotsOf_(hostCls, hostSub), b = slotsOf_(myCls, mySub), out = [];
  a.forEach(function (x) {
    b.forEach(function (y) {
      if (x.p !== y.p) return;
      var both = {}, any = false;
      Object.keys(x.days).forEach(function (d) { if (y.days[d]) { both[d] = 1; any = true; } });
      if (any) out.push({ p: x.p, days: both });
    });
  });
  return out;
}
/** roll -> [{ cls, sub, by, p:'P2 (Mon-Thu)' }].
 *  Shown ONLY to the teacher of her own section whose lecture is at the same period on the same days as her guest class
 *  (that teacher will not see her in that lecture). Every other teacher of her section sees nothing.
 *  The teacher who added the guest sees "Guest from ..." on her own sheet, as before. */
function guestAtMap_(meta) {
  var map = {};
  guestRows_().forEach(function (r) {
    if (String(r[9]) !== 'Active') return;
    if (String(r[4]) === meta.cls && String(r[5]).toUpperCase() === meta.sub.toUpperCase()) return;
    var cl = [];
    try { cl = clashSlots_(r[4], r[5], meta.cls, meta.sub); } catch (e) {}
    if (!cl.length) return;
    var byP = {};
    cl.forEach(function (x) { byP[x.p] = byP[x.p] || {}; Object.keys(x.days).forEach(function (d) { byP[x.p][d] = 1; }); });
    var label = Object.keys(byP).sort(function (a, b) { return a - b; }).map(function (p) { var dl = dayLabel_(byP[p]); return 'P' + p + (dl ? ' (' + dl + ')' : ''); }).join(', ');
    var roll = String(r[0]).trim();
    (map[roll] = map[roll] || []).push({ cls: r[4], sub: r[5], by: r[6], p: label, clash: true });
  });
  return map;
}

function guestSearch_(user, name, q) {
  var meta = allowed_(user)[name];
  if (!meta) return { ok: false, error: 'NOT_ALLOWED: ' + name };
  q = String(q || '').trim().toLowerCase();
  if (q.length < 2) return { ok: true, results: [] };
  var sheet = bookFor_(meta.mon).getSheetByName(name), have = {};
  if (sheet && sheet.getLastRow() >= FIRST_ROW) {
    var L = readLayout_(sheet), cnt = sheet.getLastRow() - FIRST_ROW + 1;
    var rg = sheet.getRange(FIRST_ROW, L.rollCol, cnt, 1), v = rg.getValues(), nt = rg.getNotes();
    for (var i = 0; i < cnt; i++) if (!isOffNote_(nt[i][0])) have[String(v[i][0]).trim()] = true;
  }
  var out = [];
  getMaster_().rows.forEach(function (r) {
    if (out.length >= 15 || have[r[0]]) return;
    if (String(r[1]).toLowerCase().indexOf(q) === -1 && String(r[0]).toLowerCase().indexOf(q) === -1) return;
    out.push({ roll: r[0], name: r[1], cls: r[2] });
  });
  return { ok: true, results: out };
}

/** Adds (or re-enables) the guest row in one month sheet. m = master row [roll, name, classSec, subs, father]. */
function putGuestRow_(sh, roll, m, by, on) {
  var L = readLayout_(sh), last = sh.getLastRow(), n = last - FIRST_ROW + 1, row = 0;
  if (n < 1) return;
  var rolls = sh.getRange(FIRST_ROW, L.rollCol, n, 1).getValues();
  for (var i = 0; i < n; i++) if (String(rolls[i][0]).trim() === roll) { row = FIRST_ROW + i; break; }
  if (row) {                                   // she was a guest before and was removed: switch her on again
    sh.getRange(row, L.rollCol).setNote(guestNote_(m[2], by, on));
    sh.getRange(row, 1, 1, 6).setBackground(null).setFontColor('#000000');
    sh.getRange(row, 4).setFontLine('none');
    return;
  }
  row = last + 1;
  if (sh.getMaxRows() < row) sh.insertRowsAfter(sh.getMaxRows(), row - sh.getMaxRows());
  sh.getRange(last, 1, 1, MT_COL).copyTo(sh.getRange(row, 1, 1, MT_COL));   // same look and T.A formula
  sh.getRange(row, DAY_COL1, 1, 31).clearContent();
  sh.getRange(row, MT_COL).clearContent();
  sh.getRange(row, 1, 1, 6).setBackground(null).setFontColor('#000000');
  sh.getRange(row, 4).setFontLine('none');
  sh.getRange(row, L.rollCol).clearNote().setNumberFormat('@');
  sh.getRange(row, 1, 1, 6).setValues([[n + 1, roll, sectionOf_(m[2]), m[1], m[4] || '', m[2]]]);
  sh.getRange(row, L.rollCol).setNote(guestNote_(m[2], by, on));
  applyRules_(sh, row - FIRST_ROW + 1);
}

function guestAdd_(user, name, roll) {
  var meta = allowed_(user)[name];
  if (!meta) return { ok: false, error: 'NOT_ALLOWED: ' + name };
  roll = String(roll || '').trim();
  var m = getMaster_().rows.filter(function (r) { return r[0] === roll; })[0];
  if (!m) return { ok: false, error: 'Student not found.' };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var gs = guestsSheet_(), already = false;
    guestRows_().forEach(function (r) {
      if (String(r[0]).trim() === roll && String(r[4]) === meta.cls && String(r[5]).toUpperCase() === meta.sub.toUpperCase() && String(r[9]) === 'Active') already = true;
    });
    var on = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMM yyyy');
    if (!already) gs.appendRow([roll, m[1], m[4], m[2], meta.cls, meta.sub, user.teacher, on, meta.mon, 'Active']);
    var mi = SEASON.indexOf(meta.mon);
    SEASON.forEach(function (tok, k) {
      if (k < mi) return;
      try {
        var sh = bookFor_(tok).getSheetByName(makeName_(meta.cls, tok, meta.sub));
        if (sh) putGuestRow_(sh, roll, m, user.teacher, on);
      } catch (e) { Logger.log('guest_add ' + tok + ': ' + e.message); }
    });
    SpreadsheetApp.flush();
    return { ok: true, already: already };
  } finally { lock.releaseLock(); }
}

/** Switches the guest off from this month onwards (old attendance stays). Caller must hold the lock. */
function removeGuestCore_(meta, roll) {
  var gs = guestsSheet_(), n = gs.getLastRow() - 1;
  if (n > 0) {
    var v = gs.getRange(2, 1, n, GUEST_HEAD.length).getValues();
    for (var i = 0; i < n; i++) {
      if (String(v[i][0]).trim() === roll && String(v[i][4]) === meta.cls && String(v[i][5]).toUpperCase() === meta.sub.toUpperCase() && String(v[i][9]) === 'Active') gs.getRange(i + 2, 10).setValue('Removed');
    }
  }
  var mi = SEASON.indexOf(meta.mon);
  SEASON.forEach(function (tok, k) {
    if (k < mi) return;
    try {
      var sh = bookFor_(tok).getSheetByName(makeName_(meta.cls, tok, meta.sub));
      if (!sh) return;
      var L = readLayout_(sh), cnt = sh.getLastRow() - FIRST_ROW + 1;
      if (cnt < 1) return;
      var rolls = sh.getRange(FIRST_ROW, L.rollCol, cnt, 1).getValues();
      for (var j = 0; j < cnt; j++) {
        if (String(rolls[j][0]).trim() !== roll) continue;
        var row = FIRST_ROW + j;
        sh.getRange(row, L.rollCol).setNote(OFF_TAG + ' - guest removed');
        sh.getRange(row, 1, 1, 6).setBackground('#EFEFEF').setFontColor('#999999');
        sh.getRange(row, 4).setFontLine('line-through');
      }
    } catch (e) { Logger.log('guest_remove ' + tok + ': ' + e.message); }
  });
  SpreadsheetApp.flush();
  return { ok: true, guestRemoved: true };
}
function guestRemove_(user, name, roll) {
  var meta = allowed_(user)[name];
  if (!meta) return { ok: false, error: 'NOT_ALLOWED: ' + name };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return removeGuestCore_(meta, String(roll || '').trim()); } finally { lock.releaseLock(); }
}
/* =====================================================================================================
   v6.0  DASHBOARDS: attendance index + college / department overview
   One row per attendance sheet in the sheet "Attendance Index" (Users workbook). It is refreshed every
   time a teacher saves or deletes attendance, so the overview reads ONE sheet instead of 1,600.
   Run rebuildAttendanceIndex() once by hand after installing v6.0 (it continues by itself).
   ===================================================================================================== */
var INDEX_SHEET = 'Attendance Index';
var INDEX_HEAD  = ['Sheet', 'Month', 'Class', 'Subject', 'Teacher', 'Students', 'Lectures held', 'Attended (sum)',
                   'Possible (sum)', 'Avg %', 'Below 73%', 'Marked days', 'Low students (do not edit)', 'Updated',
                   'Test total', 'Students data (do not edit)'];
var LOW_PCT = 73;

var INDEX_OK_ = false;
function indexSheet_() {
  var book = SpreadsheetApp.openById(USERS_SS_ID), sh = book.getSheetByName(INDEX_SHEET);
  if (sh) {
    if (!INDEX_OK_ && String(sh.getRange(1, INDEX_HEAD.length).getValue()) !== INDEX_HEAD[INDEX_HEAD.length - 1]) {
      if (sh.getMaxColumns() < INDEX_HEAD.length) sh.insertColumnsAfter(sh.getMaxColumns(), INDEX_HEAD.length - sh.getMaxColumns());
      sh.getRange(1, 1, 1, INDEX_HEAD.length).setValues([INDEX_HEAD]).setFontWeight('bold').setBackground('#D9D9D9');
    }
    INDEX_OK_ = true;
    return sh;
  }
  sh = book.insertSheet(INDEX_SHEET);
  sh.getRange(1, 1, 1, INDEX_HEAD.length).setValues([INDEX_HEAD]).setFontWeight('bold').setBackground('#D9D9D9');
  sh.setFrozenRows(1);
  sh.getRange(2, 12, 1000, 1).setNumberFormat('@');      // "1,2,3" must stay text
  sh.setColumnWidth(1, 300); sh.setColumnWidth(13, 300);
  return sh;
}

/** Lectures a guest was marked in (P or A) from her row of day cells: P=1, 2P=2, 3A=3 ... (L and B count 0). */
function guestHeld_(cells) {
  var n = 0;
  cells.forEach(function (c) {
    var m = String(c || '').trim().toUpperCase().match(/^(\d*)([PA])$/);
    if (m) n += m[1] ? Number(m[1]) : 1;
  });
  return n;
}

/** Re-reads one attendance sheet and writes its summary row to the index. meta = { name, mon, cls, sub, teacher }. */
function indexUpdate_(sh, meta) {
  if (!sh || !meta) return;
  var L = readLayout_(sh), n = sh.getLastRow() - FIRST_ROW + 1, rg = dayColRange_(L);
  var row = [meta.name || sh.getName(), meta.mon, meta.cls, meta.sub, meta.teacher || '', 0, 0, 0, 0, '', 0, '', '[]', new Date(), 0, '[]'];
  if (n > 0 && rg) {
    var dayW = rg.c2 - rg.c1 + 1;
    var block = sh.getRange(FIRST_ROW, rg.c1, n, dayW + 1).getValues();      // all days + T.A
    var rr = sh.getRange(FIRST_ROW, L.rollCol, n, 1), rolls = rr.getValues(), rnotes = rr.getNotes();
    var names = L.nameCol > 0 ? sh.getRange(FIRST_ROW, L.nameCol, n, 1).getValues() : null;
    var held = Number(sh.getRange(2, rg.c2 + 1).getValue()) || 0;
    var hdrNotes = sh.getRange(HEADER_ROW, rg.c1, 1, dayW).getNotes()[0];
    var mts = L.mtCol > 0 ? sh.getRange(FIRST_ROW, L.mtCol, n, 1).getValues() : null;
    var mtTotal = L.mtCol > 0 ? Number(sh.getRange(2, L.mtCol).getValue()) || 0 : 0;
    var all = [];                                    // every student: [roll, name, attended, held, test, off(1/0), guest home]
    var marked = [];
    Object.keys(L.dayCols).forEach(function (d) {
      var j = L.dayCols[d] - rg.c1, any = !!hdrNotes[j];                     // a note (college closed ...) counts as done
      for (var i = 0; i < n && !any; i++) if (String(rolls[i][0]).trim() && String(block[i][j]).trim()) any = true;
      if (any) marked.push(Number(d));
    });
    marked.sort(function (a, b) { return a - b; });
    var students = 0, sumP = 0, sumPos = 0, low = [];
    for (var i = 0; i < n; i++) {
      var roll = String(rolls[i][0]).trim();
      if (!roll) continue;
      var note = String(rnotes[i][0] || ''), off = isOffNote_(note), guest = note.indexOf('GUEST|') === 0;
      var ta = Number(block[i][dayW]) || 0;
      var den = guest ? guestHeld_(block[i].slice(0, dayW)) : held;
      var mv = mts ? String(mts[i][0]).trim() : '';
      all.push([roll, names ? String(names[i][0]).trim() : roll, ta, den, mv === '' ? '' : (/^A/i.test(mv) ? 'A' : Number(mv)), off ? 1 : 0, guest ? (note.split('|')[1] || '') : '']);
      if (off) continue;
      students++;
      if (den > 0) {
        sumP += ta; sumPos += den;
        if (ta / den * 100 < LOW_PCT) low.push([roll, names ? String(names[i][0]).trim() : roll, ta, den]);
      }
    }
    var lowJson = JSON.stringify(low);
    while (lowJson.length > 45000 && low.length) { low.pop(); lowJson = JSON.stringify(low); }
    row[5] = students; row[6] = held; row[7] = sumP; row[8] = sumPos;
    row[9] = sumPos ? Math.round(sumP / sumPos * 1000) / 10 : '';
    row[10] = low.length; row[11] = marked.join(','); row[12] = lowJson;
    row[14] = mtTotal; row[15] = JSON.stringify(all).slice(0, 49000);
  }
  var ix = indexSheet_(), last = ix.getLastRow(), r = 0;
  var find = function () {
    var lr = ix.getLastRow(); if (lr < 2) return 0;
    var f = ix.getRange(2, 1, lr - 1, 1).createTextFinder(row[0]).matchCase(true).matchEntireCell(true).findNext();
    return f ? f.getRow() : 0;
  };
  r = last > 1 ? find() : 0;
  if (!r) {                                                            // v7.4: new sheet in the index: appendRow is one step,
    ix.appendRow([row[0]]);                                            //       so two saves at the same time never take the same row
    SpreadsheetApp.flush();
    r = find();
    if (!r) return;
  }
  ix.getRange(r, 12).setNumberFormat('@');
  ix.getRange(r, 1, 1, INDEX_HEAD.length).setValues([row]);
}

/** Run ONCE by hand after installing v6.1: fills the index for every month from September up to now
 *  (continues by itself in rounds of 4 minutes). Safe to run again. */
function rebuildAttendanceIndex() {
  var metas = allowed_({ username: ADMIN_USER, teacher: '' });
  batchSheets_('rebuildAttendanceIndex', function (sh) {
    var meta = metas[sh.getName()];
    if (meta) indexUpdate_(sh, meta);
  }, activeMonths_());
}

/** 'all' (admin, principal), 'dept' (HOD) or '' (teacher). */
function scopeOf_(user) {
  if (isAdmin_(user)) return { all: true, dept: '' };
  var role = String(user.role || '').toUpperCase();
  if (/ADMIN|PRINCIPAL/.test(role)) return { all: true, dept: '' };
  if (/\bHOD\b|HEAD OF DEP/.test(role)) return { all: false, dept: String(user.dept || '').trim().toUpperCase() };
  return null;
}
function scopeName_(user) { var s = scopeOf_(user); return s ? (s.all ? 'all' : 'dept') : ''; }
function whoami_(user) {
  var s = scopeOf_(user), out = { ok: true, view: scopeName_(user), dept: s && !s.all ? user.dept : '' };
  try { out.cp = isCP_(user); } catch (e) { out.cp = false; }
  try { out.tt = myTimetable_(user); } catch (e) { out.tt = []; }
  try { out.times = periodTimes_(); } catch (e) { out.times = []; }
  return out;
}

/** normalized teacher name -> department (from the Users sheet). */
function teacherDepts_() {
  var map = {}, v = usersTable_();
  if (!v.length) return map;
  var iD = usersCol_(v[0], 'department', 3);
  for (var i = 1; i < v.length; i++) if (v[i][0]) map[normalizeName(v[i][0])] = String(v[i][iD] || '').trim();
  return map;
}

/** name -> index row, for one month. */
function indexRows_(tok) {
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(INDEX_SHEET), out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, INDEX_HEAD.length).getValues().forEach(function (r) {
    if (String(r[1]) === tok) out[String(r[0])] = r;
  });
  return out;
}

/** The sections this admin / HOD may see in the current month: [{ a, dept, name, x }] */
function scopedSections_(user) {
  var scope = scopeOf_(user);
  if (!scope) return null;
  var tok = currentSeasonToken_();
  if (!tok) return { tok: null, list: [], scope: scope };
  var depts = teacherDepts_(), idx = indexRows_(tok), list = [];
  getAssignments_({ teacher: '', username: ADMIN_USER }).forEach(function (a) {
    var dept = depts[normalizeName(a.teacher)] || '';
    if (!scope.all && dept.toUpperCase() !== scope.dept) return;
    var name = makeName_(a.cls, tok, a.subject);
    list.push({ a: a, dept: dept, name: name, x: idx[name] || null });
  });
  return { tok: tok, list: list, scope: scope };
}

/** Admin / HOD overview of the current month: who marked today, section and subject percentages. */
function overview_(user) {
  var S = scopedSections_(user);
  if (!S) return { ok: false, error: 'NOT_ALLOWED' };
  var now = new Date(), day = now.getDate();
  var res = { ok: true, tok: S.tok, day: day, sunday: now.getDay() === 0, scope: S.scope.all ? 'College' : (S.scope.dept || 'Department'), rows: [] };
  if (!S.tok) return res;
  var my = monthYear_(S.tok);
  res.m = my.m; res.y = my.y;
  var dow = now.getDay();
  S.list.forEach(function (s) {
    var x = s.x, days = x ? String(x[11]).split(',').map(Number).filter(Boolean) : [];
    var sl = [], sched = null;
    try { sl = slotsOf_(s.a.cls, s.a.subject); } catch (e) {}
    if (sl.length) sched = sl.some(function (q) { return !!q.days[dow]; });   // null = not in the timetable
    res.rows.push({ sched: sched,
      name: s.name, cls: s.a.cls, sub: s.a.subject, teacher: s.a.teacher, dept: s.dept,
      students: x ? Number(x[5]) || 0 : 0, held: x ? Number(x[6]) || 0 : 0,
      att: x ? Number(x[7]) || 0 : 0, pos: x ? Number(x[8]) || 0 : 0,
      avg: x && x[9] !== '' ? Number(x[9]) : null, low: x ? Number(x[10]) || 0 : 0,
      today: days.indexOf(day) !== -1, last: days.length ? days[days.length - 1] : 0, marked: days.length
    });
  });
  return res;
}

/** v7.5: "Marked today" list for admin / HOD: which teacher marked which section and subject today,
 *  with periods and present / absent / leave counts (from the Period Data sheet, read once). */
function overviewToday_(user) {
  var S = scopedSections_(user);
  if (!S) return { ok: false, error: 'NOT_ALLOWED' };
  var now = new Date(), day = now.getDate(), out = [];
  if (!S.tok) return { ok: true, day: day, rows: out };
  var pd = {};
  try {
    var sh = bookFor_(S.tok).getSheetByName(PD_SHEET);
    if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { if (r[0]) pd[String(r[0])] = r[1]; });
  } catch (e) {}
  S.list.forEach(function (s) {
    var x = s.x; if (!x) return;
    var days = String(x[11]).split(',').map(Number);
    if (days.indexOf(day) === -1) return;
    var per = [], P = 0, A = 0, L = 0, raw = null;
    try { raw = pd[s.name] ? JSON.parse(pd[s.name]) : null; } catch (e) {}
    var dd = raw && raw.d ? raw.d[day] || raw.d[String(day)] : null;
    if (dd) Object.keys(dd).sort(function (a, b) { return Number(a) - Number(b); }).forEach(function (p) {
      var t = String(dd[p]), c = { p: p, P: 0, A: 0, L: 0 };
      for (var i = 0; i < t.length; i++) { var ch = t.charAt(i); if (ch === 'P' || ch === 'A' || ch === 'L') c[ch]++; }
      P += c.P; A += c.A; L += c.L; per.push(c);
    });
    var at = x[13] instanceof Date ? Utilities.formatDate(x[13], Session.getScriptTimeZone(), 'h:mm a') : String(x[13] || '');
    out.push({ teacher: s.a.teacher, dept: s.dept, cls: s.a.cls, sub: s.a.subject, name: s.name,
               students: Number(x[5]) || 0, per: per, P: P, A: A, L: L, at: at });
  });
  out.sort(function (a, b) { return String(a.teacher).localeCompare(String(b.teacher)) || String(a.cls).localeCompare(String(b.cls)); });
  return { ok: true, day: day, rows: out };
}

/** Students below 73% this month in every section this admin / HOD may see. */
function overviewLow_(user) {
  var S = scopedSections_(user);
  if (!S) return { ok: false, error: 'NOT_ALLOWED' };
  var out = [];
  S.list.forEach(function (s) {
    if (!s.x) return;
    var low = [];
    try { low = JSON.parse(s.x[12] || '[]'); } catch (e) {}
    low.forEach(function (l) {
      out.push({ roll: l[0], name: l[1], p: l[2], held: l[3], pct: Math.round(l[2] / l[3] * 1000) / 10,
                 cls: s.a.cls, sub: s.a.subject, teacher: s.a.teacher, dept: s.dept, sheet: s.name });
    });
  });
  out.sort(function (a, b) { return a.pct - b.pct; });
  return { ok: true, tok: S.tok, rows: out.slice(0, 4000), total: out.length };
}

/* =====================================================================================================
   v6.1  REPORTS: per-student attendance and monthly test data for any month range, from the index.
   Teacher: only her own sections/subjects.  HOD: her department.  Admin / Principal: whole college.
   req: { from:'SEPT', to:'OCT', cls:'FCS-1 D2' (optional), sub:'COMPUTER SCIENCE' (optional), roll (optional) }
   ===================================================================================================== */
function reportData_(user, req) {
  var scope = scopeOf_(user), own = null, depts = null;
  if (!scope) { own = {}; listSheets_(user).forEach(function (s) { own[s.name] = true; }); }
  else if (!scope.all) depts = teacherDepts_();
  var act = activeMonths_(), fi = Math.max(0, SEASON.indexOf(String(req.from || 'SEPT'))), ti = SEASON.indexOf(String(req.to || ''));
  if (ti < 0 || ti > act.length - 1) ti = act.length - 1;
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(INDEX_SHEET), out = [];
  var wantCls = String(req.cls || ''), wantSub = String(req.sub || '').toUpperCase(), wantRoll = String(req.roll || '');
  if (sh && sh.getLastRow() > 1) {
    var w = Math.min(sh.getLastColumn(), INDEX_HEAD.length);
    sh.getRange(2, 1, sh.getLastRow() - 1, w).getValues().forEach(function (r) {
      var name = String(r[0]), mon = String(r[1]), mi = SEASON.indexOf(mon);
      if (mi < fi || mi > ti) return;
      if (wantCls && String(r[2]) !== wantCls) return;
      if (wantSub && String(r[3]).toUpperCase() !== wantSub) return;
      if (own && !own[name]) return;
      if (depts && String(depts[normalizeName(r[4])] || '').toUpperCase() !== scope.dept) return;
      var st = [];
      try { st = JSON.parse(r[15] || '[]'); } catch (e) {}
      if (wantRoll) { st = st.filter(function (x) { return x[0] === wantRoll; }); if (!st.length) return; }
      out.push({ cls: String(r[2]), sub: String(r[3]), teacher: String(r[4]), mon: mon, held: Number(r[6]) || 0, mt: Number(r[14]) || 0, st: st });
    });
  }
  return { ok: true, from: SEASON[fi], to: SEASON[ti], view: scope ? (scope.all ? 'all' : 'dept') : '', rows: out };
}

/* =====================================================================================================
   v6.2  TIMETABLE IN THE APP: period timings, the teacher's own timetable, workload
   ===================================================================================================== */
/* Period timings and evening sections come from the timetable editor's own settings sheet ("NTT Settings",
   edited in the timetable web interface: Settings -> Timings / Evening sections). Nothing to set up here. */
var NTT_SETTINGS_SHEET = 'NTT Settings';
var DEFAULT_TIMES_ = ['08:00-09:00', '09:00-10:00', '10:00-11:00', '11:00-12:00', '12:00-13:00'];

/** "08:00", "8:00 AM", "01:00" (afternoon), a time cell -> "HH:MM" (24 h). College hours: 1:00-6:59 are taken as afternoon. */
function hhmm_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  var m = String(v || '').trim().match(/(\d{1,2})\s*[:.]\s*(\d{2})\s*(am|pm)?/i);
  if (!m) return '';
  var h = Number(m[1]), ap = (m[3] || '').toLowerCase();
  if (ap === 'pm' && h < 12) h += 12; else if (ap === 'am' && h === 12) h = 0; else if (!ap && h >= 1 && h <= 6) h += 12;
  return (h < 10 ? '0' : '') + h + ':' + m[2];
}
/** "08:00-09:00" / "08:00 – 09:00" -> { start:'08:00', end:'09:00' } or null */
function timeRange_(txt) {
  var m = String(txt || '').match(/(\d{1,2}\s*[:.]\s*\d{2}\s*(?:am|pm)?)\s*[–—-]+\s*(\d{1,2}\s*[:.]\s*\d{2}\s*(?:am|pm)?)/i);
  return m ? { start: hhmm_(m[1]), end: hhmm_(m[2]) } : null;
}
var NTT_MEMO_ = null;
/** { times:{ M:[{p,start,end}], E:[...] }, evening:{ 'PROGRAM|CLASS|SECTION':1, 'PROGRAM|CLASS|*':1 } } */
function nttSettings_() {
  if (NTT_MEMO_) return NTT_MEMO_;
  var c = cacheGetBig_('ptimes');
  if (c) return (NTT_MEMO_ = c);
  var out = { times: { M: [], E: [] }, evening: {} }, raw = { M: [], E: [] };
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(NTT_SETTINGS_SHEET);
  if (sh) sh.getDataRange().getValues().forEach(function (r, i) {
    if (!i) return;
    var type = String(r[0] || '').trim().toUpperCase(), key = String(r[1] || '').trim();
    if (type === 'EVENING' && key) out.evening[key.split('|').map(function (x) { return String(x).trim().toUpperCase().replace(/\s+/g, ' '); }).join('|')] = 1;
    if (type === 'TIME') { var m = key.match(/^([ME])\|(\d+)\|(.*)$/); if (m) raw[m[1]][Number(m[2]) - 1] = m[3]; }
  });
  if (!raw.M.some(Boolean)) raw.M = DEFAULT_TIMES_.slice();
  ['M', 'E'].forEach(function (k) {
    var n = Math.max(raw.M.length, raw[k].length);
    for (var i = 0; i < n; i++) {
      var t = timeRange_(raw[k][i]) || (k === 'E' ? null : null);      // an evening period left blank uses the morning timing (handled in the app)
      if (t && t.start && t.end) out.times[k].push({ p: String(i + 1), start: t.start, end: t.end });
    }
  });
  cachePutBig_('ptimes', out);
  return (NTT_MEMO_ = out);
}
/** kept for the app: { M:[...], E:[...] } */
function periodTimes_() { return nttSettings_().times; }
/** true when this Program / Class / Section is an evening section in the timetable settings */
function isEvening_(program, cls, section) {
  var ev = nttSettings_().evening, n = function (x) { return String(x || '').trim().toUpperCase().replace(/\s+/g, ' '); };
  return !!(ev[n(program) + '|' + n(cls) + '|' + n(section)] || ev[n(program) + '|' + n(cls) + '|*']);
}
/** "M-TH" -> [1,2,3,4]; empty -> [] */
function dayList_(txt) {
  if (!String(txt || '').trim()) return [];
  return Object.keys(daySet_(txt)).map(Number).sort(function (a, b) { return a - b; });
}
/** "2" -> ['2'], "5-6" -> ['5','6'], "II, IV" -> ['2','4'] */
function periodList_(txt) {
  var out = [];
  String(txt || '').split(/[,\/&+;]|\band\b/i).forEach(function (x) {
    var rg = String(x).split('-').map(normPeriod_).filter(Boolean);
    if (rg.length === 2 && Number(rg[1]) > Number(rg[0]) && Number(rg[1]) - Number(rg[0]) < 6) {
      for (var n = Number(rg[0]); n <= Number(rg[1]); n++) if (out.indexOf(String(n)) === -1) out.push(String(n));
    } else rg.forEach(function (n) { if (out.indexOf(n) === -1) out.push(n); });
  });
  return out;
}
/** One timetable row for the app: { teacher, cls, sub, ps:['2'], th:[days], pr:[days], room, course, program } */
function ttRow_(r) {
  var th = dayList_(r[6]), pr = dayList_(r[7]);
  if (!th.length && String(r[8] || '').trim()) th = dayList_(r[8]);
  if (!th.length && !pr.length) th = [1, 2, 3, 4, 5, 6];
  var eve = false;
  try { eve = isEvening_(r[1], r[2], r[11]) || /\(\s*EVENING\s*\)/i.test(String(r[9] || '')); } catch (e) {}
  return { teacher: r[0], cls: r[3], sub: String(r[4] || '').toUpperCase(), ps: periodList_(r[5]), th: th, pr: pr,
           room: r[9] || '', course: r[10] || '', program: r[1] || '', sh: eve ? 'E' : 'M' };
}
/** This teacher's own timetable rows. */
function myTimetable_(user) {
  var me = normalizeName(user.teacher), out = [];
  if (!me) return out;
  getTimetable_().rows.forEach(function (r) {
    if (r[3] && r[4] && normalizeName(r[0]) === me) out.push(ttRow_(r));
  });
  return out;
}
/** Admin / Principal / Chief Proctors: the whole timetable. HOD: her department. For Workload and Free teachers. */
function ttAll_(user) {
  var scope = scopeOf_(user), own = scope && !scope.all ? scope.dept : '';
  if (isCP_(user)) scope = { all: true, dept: '' };                    // Chief Proctors (also an HOD who is CP) see the whole college: duties go to any department
  if (!scope) return { ok: false, error: 'NOT_ALLOWED' };
  var depts = teacherDepts_(), rows = [];
  getTimetable_().rows.forEach(function (r) {
    if (!r[0] || !r[3] || !r[4]) return;
    var dept = depts[normalizeName(r[0])] || '';
    if (!scope.all && dept.toUpperCase() !== scope.dept) return;
    var x = ttRow_(r); x.dept = dept; rows.push(x);
  });
  var info = {}, all = usersInfo_();
  rows.forEach(function (x) {                                         // no phone numbers here: they go only through 'contacts'
    var k = normalizeName(x.teacher), a = all[k];
    if (a && !info[k]) info[k] = { status: a.status, desig: a.desig, dept: a.dept };
  });
  return { ok: true, v: 2, rows: rows, times: periodTimes_(), info: info, wdept: own };   // wdept: an HOD-CP's own department (her Workload tab)
}

/* =====================================================================================================
   v6.4  TEACHER CONTACTS: kept apart from every report.
   Admin, principal and Chief Proctors (also when HOD): the whole college. Other HODs: their own department. Everybody else: refused.
   ===================================================================================================== */
function contacts_(user) {
  var scope = scopeOf_(user);
  if (isCP_(user)) scope = { all: true, dept: '' };                    // a Chief Proctor (even if also HOD) reaches every teacher
  if (!scope) return { ok: false, error: 'NOT_ALLOWED' };
  var out = [], v = usersTable_();
  if (!v.length) return { ok: true, list: out };
  var h = v[0];
  var iSt = usersCol_(h, 'status', 1), iDs = usersCol_(h, 'designation', 2), iDp = usersCol_(h, 'department', 3),
      iPh = usersCol_(h, 'contact', 6), iWa = usersCol_(h, 'whatsapplink', 7);
  for (var i = 1; i < v.length; i++) {
    var name = String(v[i][0] || '').trim(); if (!name) continue;
    var dept = String(v[i][iDp] || '').trim();
    if (!scope.all && dept.toUpperCase() !== scope.dept) continue;
    var phone = String(v[i][iPh] || '').trim(), wa = String(v[i][iWa] || '').trim();
    if (!phone && !wa) continue;
    out.push({ name: name, dept: dept, desig: String(v[i][iDs] || '').trim(), status: String(v[i][iSt] || '').trim(), phone: phone, wa: wa });
  }
  out.sort(function (a, b) { return (a.dept || '~').localeCompare(b.dept || '~') || a.name.localeCompare(b.name); });
  return { ok: true, list: out, scope: scope.all ? 'all' : scope.dept };
}

/* =====================================================================================================
   v6.3  FREE TEACHERS (admin, principal, HOD, Chief Proctor)
   ===================================================================================================== */
/** normalized teacher name -> { phone, status, desig, dept } from the Users sheet */
function usersInfo_() {
  var map = {}, v = usersTable_();
  if (!v.length) return map;
  var h = v[0];
  var iSt = usersCol_(h, 'status', 1), iDs = usersCol_(h, 'designation', 2), iDp = usersCol_(h, 'department', 3), iPh = usersCol_(h, 'contact', 6);
  for (var i = 1; i < v.length; i++) {
    if (!v[i][0]) continue;
    map[normalizeName(v[i][0])] = { phone: String(v[i][iPh] || '').trim(), status: String(v[i][iSt] || '').trim(),
                                    desig: String(v[i][iDs] || '').trim(), dept: String(v[i][iDp] || '').trim() };
  }
  return map;
}
/** true when this user is marked CP / CP1 / CP2 ... in the Chief Proctor column of the Users sheet */
function isCP_(user) {
  var v = usersTable_();
  if (!v.length) return false;
  var iCP = usersCol_(v[0], 'chiefproctor', 9), me = String(user.username || '').trim().toLowerCase();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][4] || '').trim().toLowerCase() === me) return /^CP\d*$/.test(String(v[i][iCP] || '').trim().toUpperCase());
  }
  return false;
}
/* =====================================================================================================
   v6.6  REGISTERS & RESULTS (Reports screen -> "Registers & results" card)
   - reg_opts : the sections / subjects this user may report on (from the Attendance Index)
   - reg_find : student search (name, father name, roll no) inside those sections
   - reg_plan : the attendance sheets for a choice (months, section or all, subject, student)
   - reg_days : day-by-day cells of up to 30 sheets per call (the app asks in parts and builds the PDF itself)
   Same rights as Reports: teacher = own subjects, HOD = department, admin / principal = whole college.
   ===================================================================================================== */
var REG_MAX_SHEETS = 30;

/** All index rows: { name, mon, cls, sub, teacher, json } */
function regIndex_() {
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(INDEX_SHEET), out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  var w = Math.min(sh.getLastColumn(), INDEX_HEAD.length);
  sh.getRange(2, 1, sh.getLastRow() - 1, w).getValues().forEach(function (r) {
    if (!r[0]) return;
    out.push({ name: String(r[0]), mon: String(r[1]), cls: String(r[2]), sub: String(r[3]).toUpperCase(),
               teacher: String(r[4] || ''), json: String(r[15] || '') });
  });
  return out;
}

/** function(indexRow) -> true when this user may see that sheet */
function regFilter_(user) {
  var scope = scopeOf_(user), own = null, depts = null;
  if (!scope) { own = {}; listSheets_(user).forEach(function (s) { own[s.name] = true; }); }
  else if (!scope.all) depts = teacherDepts_();
  return function (r) {
    if (own) return !!own[r.name];
    if (depts) return String(depts[normalizeName(r.teacher)] || '').toUpperCase() === scope.dept;
    return true;
  };
}

function regOpts_(user) {
  var ok = regFilter_(user), map = {}, act = activeMonths_();
  regIndex_().forEach(function (r) {
    if (act.indexOf(r.mon) === -1 || !ok(r)) return;
    (map[r.cls] = map[r.cls] || {})[r.sub] = true;
  });
  var list = Object.keys(map).sort(function (a, b) { return a.localeCompare(b, undefined, { numeric: true }); })
    .map(function (c) { return { cls: c, subs: Object.keys(map[c]).sort() }; });
  return { ok: true, view: scopeName_(user), sections: list, months: act };
}

/** Students of the sections this user may see. q = part of a name, father name or roll no. */
function regFind_(user, q) {
  q = String(q || '').trim().toLowerCase();
  if (q.length < 2) return { ok: true, list: [] };
  var ok = regFilter_(user), act = activeMonths_(), clsSet = {};
  regIndex_().forEach(function (r) { if (act.indexOf(r.mon) !== -1 && ok(r)) clsSet[norm_(baseClass_(r.cls))] = true; });
  var bases = Object.keys(clsSet), out = [];
  getMaster_().rows.forEach(function (r) {
    if (out.length >= 40) return;
    var c = norm_(r[2]);
    if (!bases.some(function (b) { return c === b || c.indexOf(b) === 0; })) return;
    if (String(r[0]).toLowerCase().indexOf(q) === -1 && String(r[1]).toLowerCase().indexOf(q) === -1 &&
        String(r[4] || '').toLowerCase().indexOf(q) === -1) return;
    out.push({ roll: r[0], name: r[1], father: r[4] || '', cls: r[2] });
  });
  return { ok: true, list: out };
}

/** req: { from, to, cls ('' or '*' = all), sub ('' = all), roll ('' = everybody) } -> sheets, in order */
function regPlan_(user, req) {
  var ok = regFilter_(user), act = activeMonths_();
  var fi = Math.max(0, SEASON.indexOf(String(req.from || 'SEPT'))), ti = SEASON.indexOf(String(req.to || ''));
  if (ti < 0 || ti > act.length - 1) ti = act.length - 1;
  var cls = String(req.cls || ''), sub = String(req.sub || '').toUpperCase(), roll = String(req.roll || '').trim();
  if (cls === '*') cls = '';
  var list = [];
  regIndex_().forEach(function (r) {
    var mi = SEASON.indexOf(r.mon);
    if (mi < fi || mi > ti) return;
    if (cls && r.cls !== cls) return;
    if (sub && r.sub !== sub) return;
    if (roll && r.json.indexOf('["' + roll + '"') === -1) return;
    if (!ok(r)) return;
    list.push({ name: r.name, mon: r.mon, cls: r.cls, sub: r.sub, teacher: r.teacher });
  });
  list.sort(function (a, b) {
    return a.cls.localeCompare(b.cls, undefined, { numeric: true }) || a.sub.localeCompare(b.sub) ||
           SEASON.indexOf(a.mon) - SEASON.indexOf(b.mon);
  });
  return { ok: true, from: SEASON[fi], to: SEASON[ti], sheets: list };
}

/** req: { names:[sheet names], roll:'' } -> day-by-day data of each sheet */
function regDays_(user, req) {
  var names = (req.names || []).slice(0, REG_MAX_SHEETS), roll = String(req.roll || '').trim();
  var ok = regFilter_(user), meta = {};
  regIndex_().forEach(function (r) { meta[r.name] = r; });
  var out = [];
  names.forEach(function (nm) {
    var m = meta[nm];
    if (!m || !ok(m)) { out.push({ name: nm, missing: true }); return; }
    try { out.push(regSheetDays_(m, roll)); }
    catch (e) { out.push({ name: nm, missing: true, error: String(e && e.message || e) }); }
  });
  return { ok: true, sheets: out };
}

/** One sheet: { name, mon, cls, sub, teacher, m, y, days:[d], held:[n], dn:{d:note}, st:[[roll, name, [cells], off, guestHome]] } */
function regSheetDays_(meta, roll) {
  var sh = bookFor_(meta.mon).getSheetByName(meta.name);
  if (!sh) return { name: meta.name, missing: true };
  var lastR = sh.getLastRow(), lastC = sh.getLastColumn();
  if (lastR < HEADER_ROW) return { name: meta.name, missing: true };
  var v = sh.getRange(1, 1, lastR, lastC).getValues(), hdr = v[HEADER_ROW - 1];
  var rollCol = -1, nameCol = -1, dayCols = {};
  for (var c = 0; c < hdr.length; c++) {
    var t = norm_(hdr[c]);
    if (t === 'ROLLNO' && rollCol < 0) rollCol = c;
    else if (t === 'STUDENTNAME' && nameCol < 0) nameCol = c;
    else if (t !== '' && /^\d+$/.test(t)) { var d = Number(t); if (d >= 1 && d <= 31 && dayCols[d] == null) dayCols[d] = c; }
  }
  if (rollCol < 0) return { name: meta.name, missing: true };
  var days = Object.keys(dayCols).map(Number).sort(function (a, b) { return a - b; });
  var n = lastR - FIRST_ROW + 1;
  var rnotes = n > 0 ? sh.getRange(FIRST_ROW, rollCol + 1, n, 1).getNotes() : [];
  var hnotes = sh.getRange(HEADER_ROW, 1, 1, lastC).getNotes()[0];
  var held = days.map(function (d) {
    var h = Number(v[0][dayCols[d]]) || 0;
    if (!h) for (var i = 0; i < n; i++) {
      var r = v[FIRST_ROW - 1 + i];
      if (!String(r[rollCol]).trim()) continue;
      var x = String(r[dayCols[d]]).trim().toUpperCase();
      if (x === 'P' || x === 'A') { h = 1; break; }
    }
    return h;
  });
  var dn = {};
  days.forEach(function (d) { var t = hnotes[dayCols[d]]; if (t) dn[d] = String(t).slice(0, 60); });
  var st = [];
  for (var i = 0; i < n; i++) {
    var r = v[FIRST_ROW - 1 + i], rl = String(r[rollCol]).trim();
    if (!rl || (roll && enr_(rl) !== enr_(roll))) continue;
    var note = String((rnotes[i] && rnotes[i][0]) || '');
    var guest = note.indexOf('GUEST|') === 0 ? (note.split('|')[1] || '?') : '';
    st.push([rl, nameCol >= 0 ? String(r[nameCol]).trim() : rl,
             days.map(function (d) { return String(r[dayCols[d]]).trim().toUpperCase(); }),
             isOffNote_(note) ? 1 : 0, guest]);
  }
  var my = monthYear_(meta.mon);
  return { name: meta.name, mon: meta.mon, cls: meta.cls, sub: meta.sub, teacher: meta.teacher,
           m: my.m, y: my.y, days: days, held: held, dn: dn, st: st };
}

/* =====================================================================================================
   v6.7  CHANGE HISTORY, SUBSTITUTIONS, HEALTH CHECK, STUDENT RESULT PAGE
   All new sheets live in the Users workbook (USERS_SS_ID) and are created by themselves.
   ===================================================================================================== */
var TZ_ = function () { return Session.getScriptTimeZone(); };
var ymd_ = function (d) { return Utilities.formatDate(d, TZ_(), 'yyyy-MM-dd'); };

/* ---------- 1) change history: every CHANGE or DELETE of a saved mark (first marking is not logged) ---------- */
var AUDIT_SHEET = 'Change History';
var AUDIT_HEAD  = ['When', 'Teacher', 'Username', 'Sheet', 'Month', 'Class', 'Subject', 'Date', 'Period', 'Roll No', 'Student', 'Before', 'After', 'What'];
var AUDIT_ = [];

function auditAdd_(user, meta, day, per, roll, name, was, now, what) {
  var date = '';
  if (day && meta) { var my = monthYear_(meta.mon); date = Utilities.formatDate(new Date(my.y, my.m, Number(day)), TZ_(), 'd MMM yyyy'); }
  AUDIT_.push([new Date(), user.teacher, user.username, meta ? meta.name : '', meta ? meta.mon : '', meta ? meta.cls : '', meta ? meta.sub : '',
               date, per ? 'P' + per : '', String(roll), String(name || ''), String(was), String(now), what]);
}
function auditSheet_() {
  var book = SpreadsheetApp.openById(USERS_SS_ID), sh = book.getSheetByName(AUDIT_SHEET);
  if (sh) return sh;
  sh = book.insertSheet(AUDIT_SHEET);
  sh.getRange(1, 1, 1, AUDIT_HEAD.length).setValues([AUDIT_HEAD]).setFontWeight('bold').setBackground('#D9D9D9');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 140); sh.setColumnWidth(4, 260); sh.setColumnWidth(11, 180);
  return sh;
}
/** Writes the collected changes in one go (called at the end of a save / delete). */
function auditFlush_() {
  if (!AUDIT_.length) return;
  var rows = AUDIT_; AUDIT_ = [];
  var sh = auditSheet_(), start = sh.getLastRow() + 1;
  sh.getRange(start, 8, rows.length, 6).setNumberFormat('@');
  sh.getRange(start, 1, rows.length, AUDIT_HEAD.length).setValues(rows);
}
/** req: { roll?, sheet?, cls? } -> the newest 300 changes this user may see.
 *  Teacher: her own sheets. HOD: her department. Admin / principal: everything. */
function history_(user, req) {
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(AUDIT_SHEET);
  if (!sh || sh.getLastRow() < 2) return { ok: true, rows: [] };
  var scope = scopeOf_(user), own = null, depts = null, tOf = {};
  if (!scope) own = allowed_(user);
  else if (!scope.all) { depts = teacherDepts_(); regIndex_().forEach(function (r) { tOf[r.name] = r.teacher; }); }
  var roll = String(req.roll || '').trim().toUpperCase(), sheet = String(req.sheet || ''), cls = String(req.cls || '');
  var last = sh.getLastRow(), from = Math.max(2, last - 20000 + 1);
  var v = sh.getRange(from, 1, last - from + 1, AUDIT_HEAD.length).getValues(), out = [];
  for (var i = v.length - 1; i >= 0 && out.length < 300; i--) {
    var r = v[i];
    if (roll && String(r[9]).trim().toUpperCase() !== roll) continue;
    if (sheet && String(r[3]) !== sheet) continue;
    if (cls && String(r[5]) !== cls) continue;
    if (own && !own[String(r[3])]) continue;
    if (depts && String(depts[normalizeName(tOf[String(r[3])] || r[1])] || '').toUpperCase() !== scope.dept) continue;
    out.push({ at: r[0] instanceof Date ? Utilities.formatDate(r[0], TZ_(), 'd MMM yyyy HH:mm') : String(r[0]), by: r[1], sheet: r[3], mon: r[4],
               cls: r[5], sub: r[6], date: String(r[7]), p: String(r[8]), roll: String(r[9]), name: String(r[10]), was: String(r[11]), now: String(r[12]), what: String(r[13]) });
  }
  return { ok: true, rows: out };
}

/* ---------- 2) substitutions (Chief Proctor, admin, principal, HOD) ---------- */
var SUBS_SHEET = 'Substitutions';
var SUBS_HEAD  = ['ID', 'Date', 'Day', 'Shift', 'Period', 'Class', 'Subject', 'Absent teacher', 'Substitute', 'Assigned by', 'Username', 'Note', 'Saved at', 'Result', 'Result by', 'Result at'];

function canAssign_(user) { try { return !!scopeOf_(user) || isCP_(user); } catch (e) { return !!scopeOf_(user); } }
function subsSheet_() {
  var book = SpreadsheetApp.openById(USERS_SS_ID), sh = book.getSheetByName(SUBS_SHEET);
  if (sh) {                                                            // v6.9: older sheets get the Result columns
    if (String(sh.getRange(1, SUBS_HEAD.length).getValue()) !== SUBS_HEAD[SUBS_HEAD.length - 1]) sh.getRange(1, 1, 1, SUBS_HEAD.length).setValues([SUBS_HEAD]).setFontWeight('bold').setBackground('#D9D9D9');
    return sh;
  }
  sh = book.insertSheet(SUBS_SHEET);
  sh.getRange(1, 1, 1, SUBS_HEAD.length).setValues([SUBS_HEAD]).setFontWeight('bold').setBackground('#D9D9D9');
  sh.setFrozenRows(1);
  sh.getRange(2, 1, 2000, 5).setNumberFormat('@');
  return sh;
}
function subRow_(r) {
  return { id: String(r[0]), date: r[1] instanceof Date ? ymd_(r[1]) : String(r[1]), day: String(r[2]), sh: String(r[3] || 'M'), p: String(r[4]),
           cls: String(r[5]), sub: String(r[6]), absent: String(r[7]), teacher: String(r[8]), by: String(r[9]), user: String(r[10]), note: String(r[11]),
           result: String(r[13] || 'Assigned'), rby: String(r[14] || '') };
}
/** req: { from:'yyyy-MM-dd', to:'yyyy-MM-dd' } (default: this month). Teachers see only their own duties. */
function subList_(user, req) {
  var now = new Date(), from = String(req.from || ymd_(new Date(now.getFullYear(), now.getMonth(), 1)));
  var to = String(req.to || ymd_(new Date(now.getFullYear(), now.getMonth() + 1, 0)));
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(SUBS_SHEET), out = [];
  var all = canAssign_(user), me = normalizeName(user.teacher);
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(13, Math.min(sh.getLastColumn(), SUBS_HEAD.length))).getValues().forEach(function (r) {
    if (!r[0]) return;
    var x = subRow_(r);
    if (x.date < from || x.date > to) return;
    if (!all && normalizeName(x.teacher) !== me) return;
    out.push(x);
  });
  out.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : Number(a.p) - Number(b.p); });
  return { ok: true, from: from, to: to, rows: out, can: all, settings: dutySettings_() };
}
/** rec: { id, date:'yyyy-MM-dd', sh:'M'|'E', p, cls, sub, absent, teacher (the substitute), note } */
function subSave_(user, rec) {
  if (!canAssign_(user)) return { ok: false, error: 'NOT_ALLOWED' };
  var date = String(rec.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { ok: false, error: 'Choose the date.' };
  if (!String(rec.teacher || '').trim()) return { ok: false, error: 'Choose the substitute teacher.' };
  if (!normPeriod_(rec.p)) return { ok: false, error: 'Choose the period.' };
  var id = String(rec.id || Utilities.getUuid());
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = subsSheet_(), n = sh.getLastRow() - 1;
    if (n > 0 && sh.getRange(2, 1, n, 1).getValues().some(function (r) { return String(r[0]) === id; })) return { ok: true, id: id, already: true };
    var d = date.split('-'), wd = new Date(Number(d[0]), Number(d[1]) - 1, Number(d[2])).getDay();
    var row = [id, date, DAYNAME_[wd] || 'Sun', rec.sh === 'E' ? 'E' : 'M', normPeriod_(rec.p), String(rec.cls || '').slice(0, 60), String(rec.sub || '').slice(0, 60),
               String(rec.absent || '').slice(0, 80), String(rec.teacher).trim().slice(0, 80), user.teacher, user.username, String(rec.note || '').slice(0, 200), new Date(), 'Assigned', '', ''];
    var at = sh.getLastRow() + 1;
    sh.getRange(at, 1, 1, 5).setNumberFormat('@');
    sh.getRange(at, 1, 1, SUBS_HEAD.length).setValues([row]);
    SpreadsheetApp.flush();
    return { ok: true, id: id, row: subRow_(row) };
  } finally { lock.releaseLock(); }
}
/** The one who assigned it, or the admin, can remove a substitution. */
function subDelete_(user, id) {
  id = String(id || '');
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(SUBS_SHEET);
    if (!sh || sh.getLastRow() < 2) return { ok: false, error: 'Not found.' };
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, 13).getValues();
    for (var i = 0; i < v.length; i++) {
      if (String(v[i][0]) !== id) continue;
      if (!isAdmin_(user) && String(v[i][10]).trim().toLowerCase() !== String(user.username).trim().toLowerCase()) return { ok: false, error: 'Only the person who assigned it (or the admin) can remove it.' };
      sh.deleteRow(i + 2);
      return { ok: true, id: id };
    }
    return { ok: false, error: 'Not found.' };
  } finally { lock.releaseLock(); }
}

/* ---------- 3) health check (admin, principal) ---------- */
function nextSeasonToken_() {
  var t = currentSeasonToken_(); if (!t) return null;
  var i = SEASON.indexOf(t); return i >= 0 && i < SEASON.length - 1 ? SEASON[i + 1] : null;
}
function hasTrigger_(fn) { return ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === fn && t.getEventType() === ScriptApp.EventType.CLOCK; }); }
function health_(user, part) {
  var scope = scopeOf_(user);
  if (!scope || !scope.all) return { ok: false, error: 'Only the admin and the principal can run the health check.' };
  var sq = function (s) { return String(s || '').replace(/\s+/g, '').toUpperCase(); };
  var out = { ok: true, at: Utilities.formatDate(new Date(), TZ_(), 'd MMM yyyy HH:mm'), checks: [] };
  var add = function (key, title, level, items, help, cols) {       // v7.0: items are text or { t: text, r:[...] }; v7.2: cols = Excel columns of r
    var it = items.slice(0, 1500);
    out.checks.push({ key: key, title: title, level: level, n: items.length, help: help || '', cols: cols || ['Roll No', 'Student Name', 'Father Name', 'PCS', 'Detail / Status'],
      items: it.map(function (x) { return typeof x === 'string' ? x : x.t; }),
      rows: it.map(function (x) { return typeof x === 'string' ? ['', '', '', '', x] : x.r; }) });
  };
  part = String(part || '');                                           // v6.9: '' = everything; 'data' | 'month' | 'students' = one step (for the progress bar)
  out.part = part || 'all';
  var A = getAssignments_({ teacher: '', username: ADMIN_USER });
  if (part === 'students') { try { studentChecks_(add, A); } catch (e) { add('stuerr', 'Student data checks could not run', 'warn', [String(e.message || e)], ''); } return out; }
  if (part === 'month') { healthMonth_(add, A, out); return out; }

  // students (Transform-ALL-COLLEGE-STRENGTH)
  var M = getMaster_(), seen = {}, dup = {}, noCls = [], noSub = [];
  var MR = function (r, d) { return { t: r[0] + ' · ' + r[1] + (r[2] ? ' · ' + r[2] : '') + (d ? ' · ' + d : ''), r: [r[0], r[1], r[4] || '', r[2] || '', d || ''] }; };
  var firstRow = {};
  M.rows.forEach(function (r) {
    if (seen[r[0]]) dup[r[0]] = (dup[r[0]] || [seen[r[0]]]).concat([r[2] || '?']); else { seen[r[0]] = r[2] || '?'; firstRow[r[0]] = r; }
    if (!String(r[2] || '').trim()) noCls.push(MR(r, 'no Class-Sec'));
    if (!r[3].length) noSub.push(MR(r, 'no subjects'));
  });
  var dupList = Object.keys(dup).map(function (k) { return MR(firstRow[k], 'in ' + dup[k].join(' and ')); });
  add('dup', 'Same roll number more than once', dupList.length ? 'bad' : 'ok', dupList, 'Fix in Transform-ALL-COLLEGE-STRENGTH: one row per student. A student in two rows can appear in two sections.');
  add('nocls', 'Students without a Class-Sec', noCls.length ? 'warn' : 'ok', noCls, 'They are in no attendance sheet. Fill the Class-Sec column.');
  add('nosub', 'Students without subjects', noSub.length ? 'warn' : 'ok', noSub, 'They are in no attendance sheet. Fill SUB1 … SUB7.');

  // timetable vs Users and vs students
  var users = {}; getUsers_().forEach(function (u) { users[normalizeName(u[0])] = true; });
  var ttMiss = {};
  A.forEach(function (a) { var k = normalizeName(a.teacher); if (k && !users[k]) ttMiss[a.teacher] = true; });
  add('ttusers', 'Timetable teachers with no login', Object.keys(ttMiss).length ? 'warn' : 'ok', Object.keys(ttMiss).sort().map(function (t) { return { t: t, r: [t, 'in the Timetable, not in Users'] }; }), 'Add them to the Users sheet (same spelling as the Timetable), or correct the name in the Timetable.', ['Teacher', 'Detail']);
  var cnt = {};
  M.rows.forEach(function (r) { var c = sq(r[2]); r[3].forEach(function (s) { cnt[c + '|' + s] = (cnt[c + '|' + s] || 0) + 1; }); });
  var classes = {}; M.rows.forEach(function (r) { classes[sq(r[2])] = true; });
  var empty = [];
  A.forEach(function (a) {
    var b = sq(baseClass_(a.cls)), opt = b !== sq(a.cls), n = 0;
    Object.keys(classes).forEach(function (c) { if (c === b || (opt && c.indexOf(b) === 0)) n += cnt[c + '|' + a.subject] || 0; });
    if (!n) empty.push(secRow_(a, 'no student takes this subject'));
  });
  add('empty', 'Timetable classes with no students', empty.length ? 'bad' : 'ok', empty, 'No student of this section takes this subject in the strength sheet, so its attendance sheet cannot be made. Check the section and subject spelling.', SEC_COLS);

  if (part === 'data') return out;
  healthMonth_(add, A, out);
  try { studentChecks_(add, A); } catch (e) { add('stuerr', 'Student data checks could not run', 'warn', [String(e.message || e)], ''); }
  return out;
}
/** Health check step 'month': this month's and next month's sheets, automatic jobs, student page. */
var SEC_COLS = ['PCS', 'Subject', 'Teacher', 'Detail'];
var secRow_ = function (a, d) { return { t: a.cls + ' · ' + titleCase_(a.subject) + (a.teacher ? ' · ' + a.teacher : '') + (d ? ' · ' + d : ''), r: [a.cls, titleCase_(a.subject), a.teacher || '', d || ''] }; };
function healthMonth_(add, A, out) {
  // this month: sections with no attendance yet
  var tok = currentSeasonToken_();
  if (tok) {
    var idx = indexRows_(tok), none = [], day = new Date().getDate();
    A.forEach(function (a) {
      var x = idx[makeName_(a.cls, tok, a.subject)];
      if (!x || !String(x[11] || '').trim()) none.push(secRow_(a, x ? 'sheet made, nothing marked' : 'no sheet yet'));
    });
    add('nomark', 'Sections with no attendance in ' + titleCase_(MONTH_NAMES[monthYear_(tok).m]), none.length && day > 5 ? 'warn' : 'ok', none, 'Not one day marked this month. The teacher may not be using the app.', SEC_COLS);
    var have = {}; bookFor_(tok).getSheets().forEach(function (s) { have[s.getName()] = true; });
    var miss = A.filter(function (a) { return !have[makeName_(a.cls, tok, a.subject)]; }).map(function (a) { return secRow_(a, 'sheet not created'); });
    add('sheets', 'This month\'s sheets not created', miss.length ? 'warn' : 'ok', miss, 'Use "Create missing sheets". It runs in the background for a few minutes.', SEC_COLS);
    var nt = nextSeasonToken_();
    if (nt && day >= 20) {
      var have2 = {}; bookFor_(nt).getSheets().forEach(function (s) { have2[s.getName()] = true; });
      var miss2 = A.filter(function (a) { return !have2[makeName_(a.cls, nt, a.subject)]; }).map(function (a) { return secRow_(a, 'next month sheet not made'); });
      add('next', 'Next month\'s sheets not ready (' + titleCase_(MONTH_NAMES[monthYear_(nt).m]) + ')', miss2.length ? 'warn' : 'ok', miss2, 'Use "Prepare next month" so teachers can start on the 1st without waiting.', SEC_COLS);
    }
  }
  var trig = [];
  if (!hasTrigger_('createCurrentMonthSheets')) trig.push('Monthly sheet creation (1st of every month) is not switched on');
  if (!hasTrigger_('buildStudentResults')) trig.push('Nightly update of the student result page is not switched on');
  add('triggers', 'Automatic jobs', trig.length ? 'warn' : 'ok', trig, 'Use "Switch on automatic jobs".', ['Job']);
  var p = PropertiesService.getScriptProperties();
  out.results = { at: p.getProperty('resultsAt') || '', count: Number(p.getProperty('resultsCount')) || 0 };
}
/** what: 'sheets' | 'next' | 'triggers' | 'results' - the job starts in the background within a minute. */
function healthFix_(user, what) {
  var scope = scopeOf_(user);
  if (!scope || !scope.all) return { ok: false, error: 'NOT_ALLOWED' };
  var soon = function (fn) { clearContinuation_(fn); ScriptApp.newTrigger(fn).timeBased().after(30 * 1000).create(); };
  var mark = function (tok) { if (tok) PropertiesService.getScriptProperties().setProperty('sheetJob', JSON.stringify({ tok: tok, running: true, starting: true, made: 0, total: 0, failed: [], at: Date.now() })); };
  if (what === 'sheets') { mark(currentSeasonToken_()); soon('createMonthSheetsContinue'); return { ok: true, msg: 'Creating this month\'s missing sheets in the background (a few minutes).' }; }
  if (what === 'next') { mark(nextSeasonToken_()); soon('createNextMonthSheets'); return { ok: true, msg: 'Preparing next month\'s sheets in the background (a few minutes).' }; }
  if (what === 'results') { soon('buildStudentResults'); return { ok: true, msg: 'Updating the student result page in the background (one or two minutes).' }; }
  if (what === 'triggers') {
    ScriptApp.getProjectTriggers().forEach(function (t) {
      var f = t.getHandlerFunction();
      if ((f === 'createCurrentMonthSheets' || f === 'buildStudentResults') && t.getEventType() === ScriptApp.EventType.CLOCK) ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger('createCurrentMonthSheets').timeBased().onMonthDay(1).atHour(1).create();
    ScriptApp.newTrigger('buildStudentResults').timeBased().everyDays(1).atHour(2).create();
    return { ok: true, msg: 'Switched on: sheets on the 1st of every month (1 am) and the student result page every night (2 am).' };
  }
  return { ok: false, error: 'Unknown fix.' };
}
/** Creates next month's sheets ahead of time (same as createCurrentMonthSheets, continues by itself). */
function createNextMonthSheets() {
  var tok = nextSeasonToken_();
  clearContinuation_('createNextMonthSheets');
  if (!tok) return;
  var start = Date.now(), book = bookFor_(tok), have = {};
  book.getSheets().forEach(function (s) { have[s.getName()] = true; });
  var list = getAssignments_({ teacher: '', username: ADMIN_USER }), job = sheetJob_(tok, list, have);
  for (var i = 0; i < list.length; i++) {
    var a = list[i], meta = { name: makeName_(a.cls, tok, a.subject), cls: a.cls, sub: a.subject, mon: tok, teacher: a.teacher };
    if (have[meta.name]) continue;
    if (Date.now() - start > BATCH_MS) { sheetJobSave_(job, false); ScriptApp.newTrigger('createNextMonthSheets').timeBased().after(60 * 1000).create(); return; }
    try { ensureSheet_(book, meta); have[meta.name] = true; job.made++; }
    catch (err) { Logger.log('Skipped ' + meta.name + ': ' + err.message); job.failed.push(meta.cls + ' · ' + titleCase_(meta.sub) + ': ' + err.message); }
    sheetJobSave_(job, false);
  }
  sheetJobSave_(job, true);
}

/* ---------- 4) student result page: private code per student + results rebuilt every night ---------- */
var CODES_SHEET   = 'Student Codes';
var CODES_HEAD    = ['Roll No', 'Name', 'Class', 'Code', 'Created'];
var RESULTS_SHEET = 'Student Results';
var RESULTS_HEAD  = ['Roll No', 'Name', 'Class', 'Code', 'Data (do not edit)', 'Updated'];
var RESULT_TRIES  = 6;          // wrong codes allowed per roll number every 15 minutes

function codesSheet_() {
  var book = SpreadsheetApp.openById(USERS_SS_ID), sh = book.getSheetByName(CODES_SHEET);
  if (sh) return sh;
  sh = book.insertSheet(CODES_SHEET);
  sh.getRange(1, 1, 1, CODES_HEAD.length).setValues([CODES_HEAD]).setFontWeight('bold').setBackground('#D9D9D9');
  sh.setFrozenRows(1);
  sh.getRange(2, 1, 6000, 4).setNumberFormat('@');
  return sh;
}
function newCode_() { var c = ''; for (var i = 0; i < 6; i++) c += String(Math.floor(Math.random() * 10)); return c.charAt(0) === '0' ? '7' + c.slice(1) : c; }
/** Codes of every student of a Class-Sec (made when missing). renew = new codes for the whole section (old slips stop working). */
function studentCodes_(user, cls, renew) {
  var scope = scopeOf_(user);
  if (!scope || !scope.all) return { ok: false, error: 'Only the admin and the principal can print student codes.' };
  cls = String(cls || '').trim();
  if (!cls) {                                                          // no section: the Class-Sec list of the strength sheet
    var set = {}; getMaster_().rows.forEach(function (r) { if (String(r[2] || '').trim()) set[String(r[2]).trim()] = true; });
    return { ok: true, classes: Object.keys(set).sort(function (a, b) { return a.localeCompare(b, undefined, { numeric: true }); }) };
  }
  var want = norm_(cls), list = getMaster_().rows.filter(function (r) { return norm_(r[2]) === want; });
  if (!list.length) return { ok: false, error: 'No students of ' + cls + ' in the strength sheet.' };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = codesSheet_(), n = sh.getLastRow() - 1, rowOf = {}, codeOf = {};
    if (n > 0) sh.getRange(2, 1, n, 4).getValues().forEach(function (r, i) { var k = String(r[0]).trim().toUpperCase(); rowOf[k] = i + 2; codeOf[k] = String(r[3]); });
    var add = [], today = Utilities.formatDate(new Date(), TZ_(), 'd MMM yyyy'), out = [];
    list.sort(function (a, b) { return enrollNum_(a[0]) - enrollNum_(b[0]); }).forEach(function (r) {
      var k = String(r[0]).trim().toUpperCase(), code = codeOf[k];
      if (!code || renew) {
        code = newCode_();
        if (rowOf[k]) sh.getRange(rowOf[k], 1, 1, 5).setValues([[r[0], r[1], r[2], code, today]]);
        else add.push([r[0], r[1], r[2], code, today]);
      }
      out.push({ roll: r[0], name: r[1], father: r[4] || '', cls: r[2], code: code });
    });
    if (add.length) { var at = sh.getLastRow() + 1; sh.getRange(at, 1, add.length, 4).setNumberFormat('@'); sh.getRange(at, 1, add.length, 5).setValues(add); }
    SpreadsheetApp.flush();
    var p = PropertiesService.getScriptProperties();
    return { ok: true, cls: cls, list: out, made: add.length, resultsAt: p.getProperty('resultsAt') || '' };
  } finally { lock.releaseLock(); }
}
function resultsInfo_(user) {
  var scope = scopeOf_(user);
  if (!scope || !scope.all) return { ok: false, error: 'NOT_ALLOWED' };
  var p = PropertiesService.getScriptProperties(), sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(CODES_SHEET);
  return { ok: true, at: p.getProperty('resultsAt') || '', count: Number(p.getProperty('resultsCount')) || 0, codes: sh ? Math.max(0, sh.getLastRow() - 1) : 0,
           nightly: hasTrigger_('buildStudentResults') };
}
/** Builds the page data of every student who has a code, from the Attendance Index (all months so far).
 *  Runs every night by itself (switch on in the app: College > Health check > Switch on automatic jobs). */
function buildStudentResults() {
  clearContinuation_('buildStudentResults');
  var cs = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(CODES_SHEET);
  if (!cs || cs.getLastRow() < 2) return;
  var codes = {};
  cs.getRange(2, 1, cs.getLastRow() - 1, 4).getValues().forEach(function (r) {
    var k = String(r[0]).trim().toUpperCase(); if (k && r[3]) codes[k] = { roll: String(r[0]).trim(), name: String(r[1]), cls: String(r[2]), code: String(r[3]), rows: [] };
  });
  var act = activeMonths_(), ix = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(INDEX_SHEET);
  if (ix && ix.getLastRow() > 1) ix.getRange(2, 1, ix.getLastRow() - 1, Math.min(ix.getLastColumn(), INDEX_HEAD.length)).getValues().forEach(function (x) {
    var mon = String(x[1]);
    if (!x[0] || act.indexOf(mon) === -1 || !x[15]) return;
    var st = []; try { st = JSON.parse(x[15]); } catch (e) { return; }
    var sub = String(x[3]).toUpperCase(), teacher = String(x[4] || ''), mtT = Number(x[14]) || 0;
    st.forEach(function (s) {
      var c = codes[String(s[0]).trim().toUpperCase()]; if (!c) return;
      // [subject, teacher, month, attended, held, test ('' / number / 'A'), test total, left subject (1/0)]
      c.rows.push([sub, teacher, mon, Number(s[2]) || 0, Number(s[3]) || 0, s[4] === '' || s[4] == null ? '' : s[4], mtT, s[5] ? 1 : 0]);
    });
  });
  var now = new Date(), stamp = Utilities.formatDate(now, TZ_(), 'd MMM yyyy HH:mm'), rows = [];
  Object.keys(codes).forEach(function (k) {
    var c = codes[k];
    rows.push([c.roll, c.name, c.cls, c.code, JSON.stringify(c.rows).slice(0, 49000), stamp]);
  });
  var book = SpreadsheetApp.openById(USERS_SS_ID), sh = book.getSheetByName(RESULTS_SHEET);
  if (!sh) { sh = book.insertSheet(RESULTS_SHEET); sh.hideSheet(); }
  sh.clearContents();
  sh.getRange(1, 1, 1, RESULTS_HEAD.length).setValues([RESULTS_HEAD]).setFontWeight('bold');
  if (rows.length) { sh.getRange(2, 1, rows.length, 4).setNumberFormat('@'); sh.getRange(2, 1, rows.length, RESULTS_HEAD.length).setValues(rows); }
  var p = PropertiesService.getScriptProperties();
  p.setProperty('resultsAt', stamp); p.setProperty('resultsCount', String(rows.length));
}
/** Public (no login): roll no + code -> that student's attendance and tests. Wrong codes are limited per roll number. */
function studentReport_(roll, code) {
  roll = String(roll || '').trim().toUpperCase(); code = String(code || '').replace(/\D/g, '');
  if (!roll || code.length !== 6) return { ok: false, error: 'Enter your roll number and your 6-digit code.' };
  var cache = CacheService.getScriptCache(), key = 'srf_' + roll.replace(/[^A-Z0-9]/g, ''), tries = Number(cache.get(key)) || 0;
  if (tries >= RESULT_TRIES) return { ok: false, error: 'Too many wrong tries. Try again after 15 minutes.' };
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(RESULTS_SHEET), f = null;
  if (sh && sh.getLastRow() > 1) f = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(roll).matchCase(false).matchEntireCell(true).findNext();
  var r = f ? sh.getRange(f.getRow(), 1, 1, RESULTS_HEAD.length).getValues()[0] : null;
  if (!r || String(r[3]) !== code) {
    cache.put(key, String(tries + 1), 900);
    return { ok: false, error: 'The roll number or the code is not correct.' + (r ? '' : ' (If you got your code today, your page is ready tomorrow.)') };
  }
  cache.remove(key);
  var rows = []; try { rows = JSON.parse(r[4] || '[]'); } catch (e) {}
  return { ok: true, roll: String(r[0]), name: String(r[1]), cls: String(r[2]), updated: String(r[5]), low: LOW_PCT, season: SEASON, rows: rows };
}

/* =====================================================================================================
   v6.8  UOP PORTAL SYNC
   The "JCW Portal Sync" button (Tampermonkey, on portal.uop.edu.pk/Student/Students) reads every student from
   the portal and sends the rows here with the sync key. They are kept in the sheet "Portal Data" (a full copy,
   replaced at every sync). Then Transform-ALL-COLLEGE-STRENGTH gets ONLY ITS EMPTY CELLS filled:
   nothing that already has a value is changed, missing columns are added at the right end, and a column that
   is made by a formula is never written. The health check compares both sheets.
   ===================================================================================================== */
var TRANSFORM_SHEET = 'Transform-ALL-COLLEGE-STRENGTH';
var PORTAL_SHEET = 'Portal Data', PORTAL_STAGE = 'Portal Data (receiving)';
var PORTAL_SUBS = (function () { var a = []; ['1st', '2nd'].forEach(function (y) { for (var i = 1; i <= 7; i++) a.push('sub' + i + ' ' + y + ' year'); }); return a; })();
var PORTAL_DETAIL = ['DOB', 'StudentContact', 'EmailAddress', 'PhotoName', 'Board / University', 'Passing Year', 'Roll.No.', 'Obtain Marks/CGPA',
                     'Total Marks/CGPA', 'Division', '%age Marks', 'Challan Details', 'HQ'].concat(PORTAL_SUBS);   // v7.5: HQ = Hafiz-e-Quran
var PORTAL_COLS = ['Student Id', 'Enrollment#', 'Reg#', 'Batch', 'Session', 'Student Name', 'Father Name', 'Category', 'Program', 'Class', 'Section', 'Status']
                  .concat(PORTAL_DETAIL, ['Load Status', 'Details synced']);
/* portal field -> Transform headers that may already hold it (spaces and case ignored) -> header of a NEW column (null = never add) */
var FILL_MAP = [
  ['Student Name', ['NAME', 'STUDENTNAME', 'STUDENT', 'FULLNAME'], null],
  ['Father Name', ['FATHERNAME', 'FATHER', "FATHER'SNAME", 'FATHER/GUARDIAN', 'GUARDIANNAME'], null],
  ['Reg#', ['REG#', 'REGNO', 'REG.NO', 'REGISTRATION#', 'REGISTRATIONNO'], 'REG#'],
  ['Batch', ['BATCH'], 'BATCH'],
  ['Session', ['SESSION'], 'SESSION'],
  ['Program', ['PROGRAM', 'PROGRAMME'], 'PROGRAM'],
  ['Category', ['QUOTA', 'CATEGORY', 'ADMISSIONQUOTA'], 'QUOTA'],
  ['Status', ['STATUS', 'PORTALSTATUS', 'STUDENTSTATUS'], 'STATUS'],
  ['DOB', ['DOB', 'DATEOFBIRTH', 'D.O.B'], 'DOB'],
  ['StudentContact', ['STUDENTCONTACT', 'CONTACT', 'MOBILE', 'MOBILENO', 'PHONE', 'CONTACTNO', 'CELL'], 'STUDENT CONTACT'],
  ['EmailAddress', ['EMAIL', 'EMAILADDRESS', 'E-MAIL'], 'EMAIL'],
  ['PhotoName', ['PHOTO', 'PHOTONAME', 'PIC', 'PICTURE', 'PHOTOURL'], 'PHOTO'],
  ['Board / University', ['BOARD', 'BOARD/UNIVERSITY', 'BOARD/UNI'], 'BOARD'],
  ['Passing Year', ['PASSINGYEAR', 'YEAROFPASSING'], 'PASSING YEAR'],
  ['Roll.No.', ['MATRICROLLNO', 'SSCROLLNO', 'BOARDROLLNO', 'ROLL.NO.'], 'MATRIC ROLL NO'],
  ['Obtain Marks/CGPA', ['OBTAINEDMARKS', 'OBTAINMARKS', 'OBTAINMARKS/CGPA', 'MATRICMARKS'], 'OBTAINED MARKS'],
  ['Total Marks/CGPA', ['TOTALMARKS', 'TOTALMARKS/CGPA'], 'TOTAL MARKS'],
  ['%age Marks', ['PERCENTAGE', '%AGE', '%AGEMARKS', '%'], 'PERCENTAGE'],
  ['Challan Details', ['CHALLANDETAILS', 'CHALLAN', 'FEECHALLAN'], 'CHALLAN DETAILS'],
  ['HQ', ['HQ', 'HAFIZ', 'HAFIZEQURAN', 'HAFIZ-E-QURAN', 'HIFZ'], null],
  ['Student Id', ['PORTALID', 'STUDENTID'], 'PORTAL ID']
];
var ENROLL_HEADS = ['ENROLL#', 'ENROLLMENT#', 'ENROLLNO', 'ENROLLMENTNO'];
var CLASS_HEADS = ['CLASS-SEC', 'CLASSSEC', 'CLASS-SECTION', 'PCS', 'SECTION'];
var FORMULA_RE = /QUERY|ARRAYFORMULA|IMPORTRANGE|FILTER|SORT|UNIQUE|VLOOKUP|XLOOKUP|INDEX|\{/i;

/** "1090.00" -> "1090" (same number, no decimals); other values unchanged */
function intText_(v) { var t = String(v == null ? '' : v).trim(); return /^\d+\.0+$/.test(t) ? t.replace(/\.0+$/, '') : t; }
/** a matric roll number of 0 / 00 / 000 is no roll number */
function realRoll_(v) { var t = String(v == null ? '' : v).trim(); return /^0*$/.test(t.replace(/[^\d]/g, '')) && !/[A-Za-z]/.test(t) ? '' : t; }
function portalClean_(r, C) {
  if (C['Obtain Marks/CGPA'] != null) r[C['Obtain Marks/CGPA']] = intText_(r[C['Obtain Marks/CGPA']]);
  if (C['Total Marks/CGPA'] != null) r[C['Total Marks/CGPA']] = intText_(r[C['Total Marks/CGPA']]);
  if (C['Roll.No.'] != null) r[C['Roll.No.']] = realRoll_(intText_(r[C['Roll.No.']]));
  if (C['Reg#'] != null && blank_(r[C['Reg#']]) && C['Enrollment#'] != null) r[C['Reg#']] = r[C['Enrollment#']];   // the enrollment no. is also the registration no.
}
function squash_(s) { return String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function enr_(s) { return String(s || '').toUpperCase().replace(/\s+/g, ''); }
function blank_(v) { return v === '' || v == null || (typeof v === 'string' && !v.trim()); }

/* ---------- key for the sync button (admin, principal) ---------- */
function portalKey_(user, renew) {
  var scope = scopeOf_(user);
  if (!scope || !scope.all) return { ok: false, error: 'Only the admin and the principal can set up the portal sync.' };
  var p = PropertiesService.getScriptProperties(), k = p.getProperty('portalKey');
  if (!k || renew) { k = 'JCW-' + Utilities.getUuid().replace(/-/g, '').slice(0, 20).toUpperCase(); p.setProperty('portalKey', k); }
  return { ok: true, key: k };
}
function portalInfo_(user) {
  var scope = scopeOf_(user);
  if (!scope || !scope.all) return { ok: false, error: 'NOT_ALLOWED' };
  var p = PropertiesService.getScriptProperties(), last = null;
  try { last = JSON.parse(p.getProperty('portalLast') || 'null'); } catch (e) {}
  return { ok: true, hasKey: !!p.getProperty('portalKey'), last: last };
}

/* ---------- receiving: { key, run, part, parts, mode:'full'|'quick', rows:[{column: value}] } ---------- */
function portalImport_(req) {
  var p = PropertiesService.getScriptProperties(), key = p.getProperty('portalKey');
  if (!key || String(req.key || '') !== key) return { ok: false, error: 'The sync key is not correct. Get the key again in the app: College > UOP portal sync.' };
  var run = String(req.run || ''), part = Number(req.part) || 0, parts = Number(req.parts) || 0, rows = req.rows || [];
  if (!run || !parts || part < 1 || part > parts) return { ok: false, error: 'Bad request.' };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var book = SpreadsheetApp.openById(USERS_SS_ID), st = book.getSheetByName(PORTAL_STAGE);
    if (part === 1) {
      if (!st) { st = book.insertSheet(PORTAL_STAGE); st.hideSheet(); }
      st.clearContents();
      st.getRange(1, 1, 1, PORTAL_COLS.length).setValues([PORTAL_COLS]);
      p.setProperty('portalRun', JSON.stringify({ run: run, got: 0, parts: parts, mode: req.mode === 'quick' ? 'quick' : 'full' }));
    }
    var R = JSON.parse(p.getProperty('portalRun') || '{}');
    if (R.run !== run || !st) return { ok: false, error: 'This sync was replaced by a newer one. Start again.' };
    if (part !== R.got + 1) return part <= R.got ? { ok: true, part: part, dup: true } : { ok: false, error: 'Part ' + (R.got + 1) + ' is missing. Start again.' };
    if (rows.length) {
      var vals = rows.map(function (o) { return PORTAL_COLS.map(function (c) { var v = o[c]; return v == null ? '' : String(v).slice(0, 2000); }); });
      var at = st.getLastRow() + 1;
      st.getRange(at, 1, vals.length, PORTAL_COLS.length).setNumberFormat('@');
      st.getRange(at, 1, vals.length, PORTAL_COLS.length).setValues(vals);
    }
    R.got = part; p.setProperty('portalRun', JSON.stringify(R));
    if (part < parts) return { ok: true, part: part };
        var res = portalFinish_(book, st, R.mode);
    p.deleteProperty('portalRun');
    lock.releaseLock();                                   // the move below takes the lock itself
     res.left = portalLeftPlan_();
    return res;
  } finally { lock.releaseLock(); }
}
/** Last part received: Portal Data is replaced (a quick sync keeps the details of the last full sync), then Transform blanks are filled. */
function portalFinish_(book, st, mode) {
  var n = st.getLastRow() - 1, rows = n > 0 ? st.getRange(2, 1, n, PORTAL_COLS.length).getValues() : [];
  var C = {}; PORTAL_COLS.forEach(function (c, i) { C[c] = i; });
  var stamp = Utilities.formatDate(new Date(), TZ_(), 'd MMM yyyy HH:mm');
  rows.forEach(function (r) { portalClean_(r, C); });                   // v7.1: 1090.00 -> 1090, roll no 0 -> empty, Reg# = Enrollment# when empty
  var pd = book.getSheetByName(PORTAL_SHEET);
  if (mode === 'quick' && pd && pd.getLastRow() > 1) {                 // keep the details of the last full sync
    var w = Math.min(pd.getLastColumn(), PORTAL_COLS.length), old = {}, oh = pd.getRange(1, 1, 1, w).getValues()[0];
    pd.getRange(2, 1, pd.getLastRow() - 1, w).getValues().forEach(function (r) { var k = enr_(r[oh.indexOf('Enrollment#')]) || ('ID' + r[oh.indexOf('Student Id')]); if (k) old[k] = r; });
    rows.forEach(function (r) {
      var o = old[enr_(r[C['Enrollment#']])] || old['ID' + r[C['Student Id']]]; if (!o) return;
      PORTAL_DETAIL.concat(['Details synced']).forEach(function (c) { var j = oh.indexOf(c); if (j >= 0 && blank_(r[C[c]])) r[C[c]] = o[j]; });
    });
  } else rows.forEach(function (r) { if (/^OK/.test(String(r[C['Load Status']]))) r[C['Details synced']] = stamp; });   // only students whose details were read
  if (!pd) pd = book.insertSheet(PORTAL_SHEET);
  pd.clearContents();
  pd.getRange(1, 1, 1, PORTAL_COLS.length).setValues([PORTAL_COLS]).setFontWeight('bold').setBackground('#D9D9D9');
  pd.setFrozenRows(1);
  if (rows.length) { pd.getRange(2, 1, rows.length, PORTAL_COLS.length).setNumberFormat('@'); pd.getRange(2, 1, rows.length, PORTAL_COLS.length).setValues(rows); }
  st.clearContents();
  var fill = { skipped: 'Transform sheet not found' };
  try { fill = fillTransform_(rows, C); } catch (e) { fill = { error: String(e.message || e) }; }
  cacheClear_('master_att3');
  var last = { at: stamp, mode: mode, rows: rows.length, fill: fill };
  PropertiesService.getScriptProperties().setProperty('portalLast', JSON.stringify(last));
  return { ok: true, done: true, last: last };
}

/** Header map of a sheet: normalized header -> column index (0-based). */
function headIdx_(hdr) { var m = {}; hdr.forEach(function (h, i) { var k = norm_(h); if (k && m[k] == null) m[k] = i; }); return m; }
function pickCol_(H, list) { for (var i = 0; i < list.length; i++) if (H[list[i]] != null) return H[list[i]]; return -1; }

/** Fills ONLY the empty cells of Transform-ALL-COLLEGE-STRENGTH from the portal rows. Returns what was done. */
function fillTransform_(rows, C) {
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(TRANSFORM_SHEET);
  if (!sh) return { skipped: 'Sheet ' + TRANSFORM_SHEET + ' not found.' };
  var lastR = sh.getLastRow(), lastC = sh.getLastColumn();
  if (lastR < 2) return { skipped: 'The strength sheet is empty.' };
  var rg = sh.getRange(1, 1, lastR, lastC), V = rg.getValues(), F = rg.getFormulas();
  if (F[0].some(function (f) { return FORMULA_RE.test(f); }) || (F[1] && F[1].some(function (f) { return /QUERY|IMPORTRANGE|=\s*\{/i.test(f); })))
    return { skipped: 'The strength sheet is made by a formula (QUERY / IMPORTRANGE …), so nothing was written in it. The portal data is in the sheet "Portal Data".' };
  var H = headIdx_(V[0]), ec = pickCol_(H, ENROLL_HEADS);
  if (ec < 0) return { skipped: 'No ENROLL# column in the strength sheet.' };
  var nameC = pickCol_(H, FILL_MAP[0][1]), fatherC = pickCol_(H, FILL_MAP[1][1]);
  var colHasFormula = function (c) { for (var r = 1; r < V.length; r++) if (F[r][c]) return true; return false; };

  // portal rows by enrollment, and by name + father (for strength rows without a roll number)
  var byE = {}, byNF = {};
  rows.forEach(function (r) {
    var e = enr_(r[C['Enrollment#']]); if (e) byE[e] = r;
    var nf = squash_(r[C['Student Name']]) + '|' + squash_(r[C['Father Name']]);
    if (nf !== '|') (byNF[nf] = byNF[nf] || []).push(r);
  });
  var res = { matched: 0, notOnPortal: 0, rollsFilled: 0, cells: {}, added: [], formulaCols: [] };

  // 1) missing roll numbers: only when exactly ONE enrolled portal student has that name + father name
  if (!colHasFormula(ec) && nameC >= 0) {
    for (var r = 1; r < V.length; r++) {
      if (!blank_(V[r][ec]) || blank_(V[r][nameC])) continue;
      var list = (byNF[squash_(V[r][nameC]) + '|' + squash_(fatherC >= 0 ? V[r][fatherC] : '')] || [])
        .filter(function (x) { return /^enrolled$/i.test(String(x[C['Status']]).trim()) && x[C['Enrollment#']]; });
      if (list.length === 1 && !V.some(function (row) { return enr_(row[ec]) === enr_(list[0][C['Enrollment#']]); })) {
        V[r][ec] = String(list[0][C['Enrollment#']]).trim(); res.rollsFilled++;
        sh.getRange(r + 1, ec + 1).setNumberFormat('@').setValue(V[r][ec]);
      }
    }
  } else if (colHasFormula(ec)) res.formulaCols.push(String(V[0][ec]));

  // 2) the columns to fill (missing ones are added at the right end)
  var targets = [];
  FILL_MAP.forEach(function (m) {
    var c = pickCol_(H, m[1].map(norm_));
    if (c < 0 && m[2]) { c = V[0].length; V.forEach(function (row, i) { row.push(i ? '' : m[2]); }); F.forEach(function (row) { row.push(''); }); H[norm_(m[2])] = c; res.added.push(m[2]); }
    if (c < 0) return;
    if (colHasFormula(c)) { res.formulaCols.push(String(V[0][c])); return; }
    targets.push({ f: m[0], c: c });
  });
  if (res.added.length) {
    var need = V[0].length;
    if (sh.getMaxColumns() < need) sh.insertColumnsAfter(sh.getMaxColumns(), need - sh.getMaxColumns());
    sh.getRange(1, lastC + 1, 1, res.added.length).setValues([res.added]).setFontWeight('bold');
  }
  var subC = ['SUB1', 'SUB2', 'SUB3', 'SUB4', 'SUB5', 'SUB6', 'SUB7'].map(function (h) { return H[h]; }).filter(function (c) { return c != null && !colHasFormula(c); });

  // 3) fill the empty cells, column by column, writing only the cells that change
  var changed = {};                                                   // col -> { row: value }
  res.tidied = 0;                                                      // v7.1: 1090.00 -> 1090 and matric roll no 0 -> empty (the value itself is not changed)
  targets.forEach(function (t) {
    if (['Obtain Marks/CGPA', 'Total Marks/CGPA', 'Roll.No.'].indexOf(t.f) === -1) return;
    for (var i = 1; i < V.length; i++) {
      var v = V[i][t.c]; if (blank_(v)) continue;
      var nv = t.f === 'Roll.No.' ? realRoll_(intText_(v)) : intText_(v);
      if (typeof v === 'number' && t.f !== 'Roll.No.') continue;          // a real number cell shows no decimals by itself
      if (String(nv) !== String(v)) { (changed[t.c] = changed[t.c] || {})[i] = nv; V[i][t.c] = nv; res.tidied++; }
    }
  });
  for (var i = 1; i < V.length; i++) {
    var pr = byE[enr_(V[i][ec])];
    if (!pr) { if (!blank_(V[i][ec])) res.notOnPortal++; continue; }
    res.matched++;
    targets.forEach(function (t) {
      var v = pr[C[t.f]];
      if (blank_(v) || !blank_(V[i][t.c])) return;
      if (t.f === 'Status' || t.f === 'Category') v = String(v).trim();
      (changed[t.c] = changed[t.c] || {})[i] = v; V[i][t.c] = v;
    });
    if (subC.length && subC.every(function (c) { return blank_(V[i][c]); })) {        // subjects only when ALL subject cells are empty
      var yr = /2nd|second/i.test(String(pr[C['Class']])) ? '2nd' : '1st';
      var subs = []; for (var k = 1; k <= 7; k++) { var sv = String(pr[C['sub' + k + ' ' + yr + ' year']] || '').trim(); if (sv) subs.push(sv.toUpperCase()); }
      subs.slice(0, subC.length).forEach(function (sv, k) { (changed[subC[k]] = changed[subC[k]] || {})[i] = sv; V[i][subC[k]] = sv; });
    }
  }
  Object.keys(changed).forEach(function (c) {
    c = Number(c);
    var rowsC = Object.keys(changed[c]).map(Number).sort(function (a, b) { return a - b; }), j = 0;
    res.cells[String(V[0][c])] = rowsC.length;
    while (j < rowsC.length) {                                         // one write per run of neighbouring rows
      var a = rowsC[j], b = a; while (j + 1 < rowsC.length && rowsC[j + 1] === b + 1) { j++; b++; }
      var out = []; for (var x = a; x <= b; x++) out.push([V[x][c]]);
      sh.getRange(a + 1, c + 1, out.length, 1).setNumberFormat('@').setValues(out);
      j++;
    }
  });
  SpreadsheetApp.flush();
  return res;
}

/* ---------- v6.8 health check: the student data (strength sheet + Portal Data) ---------- */
function studentChecks_(add, A) {
  var book = SpreadsheetApp.openById(USERS_SS_ID), ts = book.getSheetByName(TRANSFORM_SHEET);
  if (!ts || ts.getLastRow() < 2) return;
  var V = ts.getRange(1, 1, ts.getLastRow(), ts.getLastColumn()).getValues(), H = headIdx_(V[0]);
  var col = function (list) { return pickCol_(H, list.map(norm_)); };
  var fm = {}; FILL_MAP.forEach(function (m) { fm[m[0]] = col(m[1].concat(m[2] ? [m[2]] : [])); });
  var ec = col(ENROLL_HEADS), cc = col(CLASS_HEADS), nc = fm['Student Name'], fc = fm['Father Name'];
  var subC = ['SUB1', 'SUB2', 'SUB3', 'SUB4', 'SUB5', 'SUB6', 'SUB7'].map(function (h) { return H[h]; }).filter(function (c) { return c != null; });

  // Portal Data (latest sync)
  var P = { rows: [], byE: {}, at: '' }, pd = book.getSheetByName(PORTAL_SHEET);
  if (pd && pd.getLastRow() > 1) {
    var pv = pd.getRange(1, 1, pd.getLastRow(), Math.min(pd.getLastColumn(), PORTAL_COLS.length)).getValues(), ph = pv[0];
    P.C = {}; ph.forEach(function (h, i) { P.C[h] = i; });
    P.rows = pv.slice(1);
    P.rows.forEach(function (r) { var e = enr_(r[P.C['Enrollment#']]); if (e) P.byE[e] = r; });
  }
  var last = null; try { last = JSON.parse(PropertiesService.getScriptProperties().getProperty('portalLast') || 'null'); } catch (e) {}
  var pget = function (e, f) { var r = P.byE[enr_(e)]; return r && P.C[f] != null ? r[P.C[f]] : ''; };
  var get = function (r, f) { var c = fm[f]; var v = c >= 0 ? r[c] : ''; return blank_(v) ? pget(r[ec], f) : v; };   // strength sheet first, then the portal
  var who = function (r) { return (String(r[ec] || '').trim() || 'no roll no') + ' · ' + String(nc >= 0 ? r[nc] : '').trim() + (cc >= 0 && r[cc] ? ' · ' + r[cc] : ''); };
  var W = function (r, d) {                                            // one student line: text for the screen + columns for Excel
    var roll = String(r[ec] || '').trim(), nm = String(nc >= 0 ? r[nc] : '').trim(), fa = String(fc >= 0 ? r[fc] : '').trim(), pcs = cc >= 0 ? String(r[cc] || '').trim() : '';
    return { t: (roll || 'no roll no') + ' · ' + nm + (pcs ? ' · ' + pcs : '') + (d ? ' · ' + d : ''), r: [roll, nm, fa, pcs, d || ''] };
  };
  var rows = V.slice(1).filter(function (r) { return (ec >= 0 && !blank_(r[ec])) || (nc >= 0 && !blank_(r[nc])); });
  var withRoll = rows.filter(function (r) { return ec >= 0 && !blank_(r[ec]); });
  var L = { noroll: [], noquota: [], nopic: [], nodob: [], noqual: [], badqual: [], nocontact: [], badcontact: [], noreg: [], nofather: [], nochallan: [], unpaid: [], nosubportal: [], status: [], notportal: [] };

  rows.forEach(function (r) {
    if (ec < 0 || blank_(r[ec])) { L.noroll.push(W(r, 'no roll number')); return; }
    var onP = !!P.byE[enr_(r[ec])];
    var known = !P.rows.length || !onP || !!pget(r[ec], 'Details synced');   // details read on the portal (else only the sheet is checked)
    if (blank_(get(r, 'Category'))) L.noquota.push(W(r, 'no quota'));
    if (blank_(get(r, 'PhotoName')) && known) L.nopic.push(W(r, 'no picture'));
    var dob = get(r, 'DOB'), ds = dob instanceof Date ? Utilities.formatDate(dob, TZ_(), 'dd/MM/yyyy') : String(dob).trim();
    if (!ds && known) L.nodob.push(W(r, 'no date of birth'));
    else if (/^0?1[\/\-.]0?1[\/\-.]2000$|^2000-0?1-0?1/.test(ds)) L.nodob.push(W(r, ds + ' (placeholder date)'));
    var mr = get(r, 'Roll.No.'), ob = get(r, 'Obtain Marks/CGPA'), tt = get(r, 'Total Marks/CGPA'), miss = [];
    if (blank_(realRoll_(mr))) miss.push('matric roll no'); if (blank_(ob)) miss.push('obtained marks'); if (blank_(tt)) miss.push('total marks');
    if (miss.length && known) L.noqual.push(W(r, 'no ' + miss.join(', ')));
    else if (Number(String(ob).replace(/[^\d.]/g, '')) > Number(String(tt).replace(/[^\d.]/g, ''))) L.badqual.push(W(r, ob + ' / ' + tt));
    var ph = String(get(r, 'StudentContact')).replace(/\D/g, '');
    if (!ph) { if (known) L.nocontact.push(W(r, 'no contact')); } else if (!/^(03\d{9}|923\d{9}|3\d{9})$/.test(ph)) L.badcontact.push(W(r, String(get(r, 'StudentContact'))));
    if (fc < 0 || blank_(r[fc])) L.nofather.push(W(r, 'no father name'));
    var ch = String(get(r, 'Challan Details')).trim();
    if (!ch) { if (known) L.nochallan.push(W(r, 'no challan')); } else if (/(^|,\s*)U-/.test(ch)) L.unpaid.push(W(r, ch));
    if (onP) {
      var st = String(pget(r[ec], 'Status')).trim();
      if (st && !/^enrolled$/i.test(st)) L.status.push(W(r, st));
      var yr = /2nd|second/i.test(String(pget(r[ec], 'Class'))) ? '2nd' : '1st', any = false;
      for (var k = 1; k <= 7; k++) if (!blank_(pget(r[ec], 'sub' + k + ' ' + yr + ' year'))) any = true;
      if (!any && pget(r[ec], 'Details synced')) L.nosubportal.push(W(r, 'no ' + yr + ' year subjects on the portal'));
    } else if (P.rows.length) L.notportal.push(W(r, 'roll number not found on the portal'));
  });

  // in two programs: the same student (name + father name) with two DIFFERENT roll numbers.
  // "Pre-Medical (JCW)" and "Pre-Medical (Jinnah College)" are the same program, so only the roll numbers decide.
  var two = [], grp = {}, info = {};
  var prog_ = function (p) { return String(p || '').replace(/\(.*?\)/g, '').replace(/\s+/g, ' ').trim(); };
  var addG = function (name, father, enroll, x) {
    var k = squash_(name) + '|' + squash_(father), e = enr_(enroll); if (!e || k.split('|')[0].length < 2 || !k.split('|')[1]) return;
    var g = grp[k] = grp[k] || { n: name, f: father, e: {} }; g.e[e] = 1;
    var o = info[e] = info[e] || { roll: String(enroll).trim(), pcs: '', cls: '', prog: '', st: '' };
    ['pcs', 'cls', 'prog', 'st'].forEach(function (q) { if (!o[q] && x[q]) o[q] = String(x[q]).trim(); });
  };
  withRoll.forEach(function (r) { addG(r[nc], fc >= 0 ? r[fc] : '', r[ec], { pcs: cc >= 0 ? r[cc] : '', prog: prog_(get(r, 'Program')), cls: pget(r[ec], 'Class'), st: pget(r[ec], 'Status') }); });
  P.rows.forEach(function (r) {
    if (!/^enrolled$/i.test(String(r[P.C['Status']]).trim())) return;
    addG(r[P.C['Student Name']], r[P.C['Father Name']], r[P.C['Enrollment#']], { prog: prog_(r[P.C['Program']]), cls: r[P.C['Class']], st: 'Enrolled' });
  });
  var desc = function (o) { return [o.pcs || 'not in the strength sheet', o.cls, o.prog].filter(Boolean).join(', '); };
  Object.keys(grp).forEach(function (k) {
    var g = grp[k], es = Object.keys(g.e); if (es.length < 2) return;
    es.forEach(function (e) {
      var o = info[e], others = es.filter(function (x) { return x !== e; }).map(function (x) { return info[x].roll + ' (' + desc(info[x]) + ')'; }).join('; ');
      var d = (o.cls || '?') + (o.prog ? ' · ' + o.prog : '') + ' · also ' + others;
      two.push({ t: o.roll + ' · ' + g.n + (o.pcs ? ' · ' + o.pcs : '') + ' · ' + d, r: [o.roll, g.n, String(g.f || ''), o.pcs, d] });
    });
  });

  // roll number range: wrong form, or far outside the block of her own section
  var range = [], bySec = {};
  withRoll.forEach(function (r) {
    var e = String(r[ec]).trim();
    if (!/^JCW-[A-Z]\d{2}-\d{1,5}$/i.test(e)) { range.push(W(r, 'not in the form JCW-F26-1234')); return; }
    var sec = cc >= 0 ? String(r[cc] || '').trim() : ''; if (!sec) return;
    var pre = e.toUpperCase().split('-').slice(0, 2).join('-');
    (bySec[sec + '|' + pre] = bySec[sec + '|' + pre] || []).push({ n: enrollNum_(e), r: r });
  });
  Object.keys(bySec).forEach(function (k) {
    var xs = bySec[k].sort(function (a, b) { return a.n - b.n; }); if (xs.length < 5) return;
    var best = [0, 0], s0 = 0;                                         // the longest run of numbers without a big gap = the section's block
    for (var i = 1; i <= xs.length; i++) { if (i === xs.length || xs[i].n - xs[i - 1].n > 15) { if (i - s0 > best[1] - best[0]) best = [s0, i]; s0 = i; } }
    var lo = xs[best[0]].n, hi = xs[best[1] - 1].n;
    xs.forEach(function (x, i) { if (i < best[0] || i >= best[1]) range.push(W(x.r, 'outside ' + k.split('|')[0] + ' range ' + lo + '–' + hi)); });
  });

  // enrolled on the portal but missing from the strength sheet
  var inSheet = {}; withRoll.forEach(function (r) { inSheet[enr_(r[ec])] = 1; });
  var missing = P.rows.filter(function (r) { return /^enrolled$/i.test(String(r[P.C['Status']]).trim()) && r[P.C['Enrollment#']] && !inSheet[enr_(r[P.C['Enrollment#']])]; })
    .map(function (r) { var pc = [r[P.C['Program']], r[P.C['Class']], r[P.C['Section']]].join(' ').replace(/\s+/g, ' ').trim();
      return { t: r[P.C['Enrollment#']] + ' · ' + r[P.C['Student Name']] + ' · ' + pc, r: [r[P.C['Enrollment#']], r[P.C['Student Name']], r[P.C['Father Name']], pc, 'Enrolled on the portal'] }; });

  // rooms in the timetable
  var room = [];
  getTimetable_().rows.forEach(function (r) { if (r[0] && r[3] && r[4] && !String(r[9] || '').trim()) room.push(secRow_({ cls: r[3], subject: r[4], teacher: r[0] }, 'period ' + r[5] + ', no room')); });

  var lv = function (list, bad) { return list.length ? (bad ? 'bad' : 'warn') : 'ok'; };
  var noP = P.rows.length ? '' : ' (Sync the UOP portal first: College > UOP portal sync.)';
  add('stu_noroll', 'Students without roll number', lv(L.noroll, true), L.noroll, 'Rows with a name but no ENROLL#. They are in NO attendance sheet. The portal sync fills the roll number when exactly one enrolled portal student has the same name and father name.');
  add('stu_status', 'Cancelled, migrated, disenrolled or transferred', lv(L.status, true), L.status, 'Still in the strength sheet (so still in attendance) but no longer Enrolled on the portal. Use "Move them to Left Students": they go to the sheet Left Students, are struck through in this month\'s attendance and removed from later months.' + noP);
  add('stu_missing', 'Enrolled on the portal but not in the strength sheet', lv(missing, true), missing, 'They are in no attendance sheet. Add them to the strength sheet with their Class-Sec and subjects.' + noP);
  add('stu_notportal', 'In the strength sheet but not on the portal', lv(L.notportal), L.notportal, 'The roll number was not found on the portal: check the roll number.');
  add('stu_two', 'In two programs (same student twice)', lv(two, true), two, 'Same name and father name with two roll numbers. Keep one admission.');
  add('stu_range', 'Roll number out of range', lv(range), range, 'Roll numbers that are not in the JCW-F26-1234 form, or far away from the roll numbers of the rest of the section.');
  add('stu_quota', 'No quota', lv(L.noquota), L.noquota, 'Quota / category is empty in the strength sheet and on the portal.' + noP);
  add('stu_pic', 'No picture', lv(L.nopic), L.nopic, 'No photo on the portal.' + noP);
  add('stu_dob', 'No date of birth, or 1/1/2000', lv(L.nodob), L.nodob, '01/01/2000 is a placeholder: enter the real date of birth on the portal.' + noP);
  add('stu_qual', 'Matric roll no / obtained / total marks missing', lv(L.noqual), L.noqual, 'From the qualification on the portal.' + noP);
  add('stu_qualbad', 'Obtained marks more than total marks', lv(L.badqual, true), L.badqual, 'Typing mistake in the qualification.');
  add('stu_subject', 'Subjects missing on the portal', lv(L.nosubportal), L.nosubportal, 'No courses on the portal for her current year.' + noP);
  add('stu_contact', 'No student contact', lv(L.nocontact), L.nocontact, 'No mobile number.' + noP);
  add('stu_contactbad', 'Contact number looks wrong', lv(L.badcontact), L.badcontact, 'Not an 11-digit mobile number (03xxxxxxxxx).');
  add('stu_challan', 'Challan details missing', lv(L.nochallan), L.nochallan, 'No fee challan on the portal.' + noP);
  add('stu_unpaid', 'Unpaid challan', lv(L.unpaid), L.unpaid, 'U- = unpaid challan (P- = paid).' + noP);
  add('stu_father', 'No father name', lv(L.nofather), L.nofather, '');
  add('room', 'Timetable classes without a room', lv(room), room, 'Fill the Room column of the Timetable.', SEC_COLS);
  var days = last && last.at ? 0 : 999;
  if (last && last.at) { var d = new Date(last.at.replace(/(\d+) (\w+) (\d+) .*/, '$2 $1, $3')); if (!isNaN(d)) days = Math.floor((Date.now() - d.getTime()) / 864e5); }
  add('portal', 'UOP portal data', days > 7 ? 'warn' : 'ok', days > 7 ? [last && last.at ? 'Last synced ' + last.at + ' (' + days + ' days ago)' : 'Never synced'] : [], 'Run "JCW Portal Sync" on the portal Students page (College > UOP portal sync).');
}

/* =====================================================================================================
   v6.9  MONTH SHEETS STATUS, DUTY RESULTS AND FAIR-DUTY SETTINGS
   ===================================================================================================== */
/** Progress of the background sheet creation (kept in a script property, read by the app every few seconds). */
function sheetJob_(tok, list, have) {
  var need = list.filter(function (a) { return !have[makeName_(a.cls, tok, a.subject)]; }).length, p = PropertiesService.getScriptProperties(), old = null;
  try { old = JSON.parse(p.getProperty('sheetJob') || 'null'); } catch (e) {}
  var job = old && old.tok === tok && old.running && !old.starting ? old : { tok: tok, made: 0, failed: [], total: need, start: Date.now() };
  job.running = true; job.starting = false; job.left = need; job.total = Math.max(job.total || 0, job.made + need);
  job.saveAt = 0;
  return job;
}
function sheetJobSave_(job, finished) {
  if (!finished && Date.now() - (job.saveAt || 0) < 5000) return;      // at most every 5 seconds
  job.saveAt = Date.now(); job.at = Date.now(); job.running = !finished; job.failed = job.failed.slice(-30);
  PropertiesService.getScriptProperties().setProperty('sheetJob', JSON.stringify(job));
}
/** { cur:{tok,label,have,need,missing[]}, next:{…} or null, job:{…} } for the Month sheets card (admin, principal). */
function sheetsStatus_(user) {
  var scope = scopeOf_(user);
  if (!scope || !scope.all) return { ok: false, error: 'NOT_ALLOWED' };
  var A = getAssignments_({ teacher: '', username: ADMIN_USER });
  var one = function (tok) {
    if (!tok) return null;
    var have = {}; bookFor_(tok).getSheets().forEach(function (s) { have[s.getName()] = true; });
    var miss = A.filter(function (a) { return !have[makeName_(a.cls, tok, a.subject)]; }).map(function (a) { return a.cls + ' · ' + titleCase_(a.subject); });
    var my = monthYear_(tok);
    return { tok: tok, label: titleCase_(MONTH_NAMES[my.m]) + ' ' + my.y, need: A.length, have: A.length - miss.length, missing: miss.slice(0, 300) };
  };
  var job = null; try { job = JSON.parse(PropertiesService.getScriptProperties().getProperty('sheetJob') || 'null'); } catch (e) {}
  if (job && job.running && Date.now() - (job.at || 0) > 15 * 60000) job.stalled = true;   // no news for 15 minutes
  return { ok: true, cur: one(currentSeasonToken_()), next: one(nextSeasonToken_()), job: job, auto: hasTrigger_('createCurrentMonthSheets') };
}

/** Chief Proctor / admin marks a duty Done or Not done (or back to Assigned). */
function subResult_(user, id, result) {
  if (!canAssign_(user)) return { ok: false, error: 'NOT_ALLOWED' };
  if (['Done', 'Not done', 'Assigned'].indexOf(result) === -1) return { ok: false, error: 'Unknown result.' };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = subsSheet_(), n = sh.getLastRow() - 1;
    if (n < 1) return { ok: false, error: 'Not found.' };
    var ids = sh.getRange(2, 1, n, 1).getValues();
    for (var i = 0; i < n; i++) {
      if (String(ids[i][0]) !== String(id)) continue;
      sh.getRange(i + 2, 14, 1, 3).setValues([[result, result === 'Assigned' ? '' : user.teacher, result === 'Assigned' ? '' : new Date()]]);
      return { ok: true, id: id, result: result };
    }
    return { ok: false, error: 'Not found.' };
  } finally { lock.releaseLock(); }
}

/** Fair-duty settings, shared by every Chief Proctor: { maxDay, exempt:[names], owedFirst, fit } */
function dutySettings_() {
  var d = { maxDay: 1, exempt: [], owedFirst: true, fit: true };
  try { var s = JSON.parse(PropertiesService.getScriptProperties().getProperty('dutySettings') || '{}'); Object.keys(d).forEach(function (k) { if (s[k] !== undefined) d[k] = s[k]; }); } catch (e) {}
  return d;
}
function dutySave_(user, st) {
  var scope = scopeOf_(user), cp = false; try { cp = isCP_(user); } catch (e) {}
  if (!cp && !(scope && scope.all)) return { ok: false, error: 'Only the Chief Proctor, the admin and the principal can change the duty settings.' };
  var d = { maxDay: Math.max(0, Math.min(6, Number(st.maxDay) || 0)), exempt: (Array.isArray(st.exempt) ? st.exempt : []).map(function (x) { return String(x).slice(0, 80); }).slice(0, 200),
            owedFirst: st.owedFirst !== false, fit: st.fit !== false, by: user.teacher, at: Utilities.formatDate(new Date(), TZ_(), 'd MMM yyyy HH:mm') };
  PropertiesService.getScriptProperties().setProperty('dutySettings', JSON.stringify(d));
  return { ok: true, settings: d };
}

/* =====================================================================================================
   v7.0  STUDENTS WHO LEFT (Migrated, Cancelled, Disenrolled, Transfer on the portal)
   "Move" takes their rows out of Transform-ALL-COLLEGE-STRENGTH into the sheet "Left Students" (nothing is lost:
   the whole row is kept, with the portal status, date and who moved it). Then, in the background:
     - months before this one: nothing changes (her attendance stays as it was);
     - this month: her row is struck through and switched off ("LEFT COLLEGE" note), so it is not counted any more;
     - later months that already have sheets: her row is removed (or struck through if marks were already entered there).
   Sheets made later never include her, because she is no longer in the strength sheet.
   ===================================================================================================== */
var LEFT_SHEET = 'Left Students';
var LEFT_EXTRA = ['MOVED ON', 'MOVED BY', 'PORTAL STATUS', 'LEFT IN MONTH'];
var LEFT_STATUSES = ['Cancelled', 'Migrated', 'Disenrolled', 'Transfer'];
/* ---- v7.0 Left Students functions (put back) + v7.3 sync buttons. Paste right after the line  var LEFT_STATUSES = [...]  ---- */
function portalStatusMap_() {
  var pd = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(PORTAL_SHEET), out = {};
  if (!pd || pd.getLastRow() < 2) return out;
  var v = pd.getRange(1, 1, pd.getLastRow(), Math.min(pd.getLastColumn(), PORTAL_COLS.length)).getValues(), h = v[0], e = h.indexOf('Enrollment#'), st = h.indexOf('Status');
  for (var i = 1; i < v.length; i++) if (v[i][e]) out[enr_(v[i][e])] = String(v[i][st] || '').trim();
  return out;
}
function transformRead_() {
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(TRANSFORM_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;
  var rg = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()), V = rg.getValues(), F = rg.getFormulas(), H = headIdx_(V[0]);
  var T = { sh: sh, V: V, F: F, H: H, ec: pickCol_(H, ENROLL_HEADS), nc: pickCol_(H, FILL_MAP[0][1]), fc: pickCol_(H, FILL_MAP[1][1]), cc: pickCol_(H, CLASS_HEADS) };
  T.formula = F[0].some(function (f) { return FORMULA_RE.test(f); }) || (F[1] && F[1].some(function (f) { return /QUERY|IMPORTRANGE|=\s*\{/i.test(f); }));
  return T;
}
/** The students of the strength sheet whose portal status is one of `statuses`. */
function leftPreview_(user, statuses) {
  var scope = scopeOf_(user);
  if (!scope || !scope.all) return { ok: false, error: 'Only the admin and the principal can move students.' };
  var want = {}; (statuses && statuses.length ? statuses : LEFT_STATUSES).forEach(function (x) { want[String(x).toUpperCase()] = 1; });
  var T = transformRead_(); if (!T) return { ok: false, error: 'The strength sheet is empty.' };
  if (T.ec < 0) return { ok: false, error: 'No ENROLL# column in the strength sheet.' };
  var PS = portalStatusMap_(); if (!Object.keys(PS).length) return { ok: false, error: 'No portal data yet: run the UOP portal sync first.' };
  var list = [];
  for (var i = 1; i < T.V.length; i++) {
    var r = T.V[i], e = enr_(r[T.ec]), st = PS[e];
    if (!e || !st || !want[st.toUpperCase()]) continue;
    list.push({ roll: String(r[T.ec]).trim(), name: T.nc >= 0 ? String(r[T.nc]) : '', father: T.fc >= 0 ? String(r[T.fc]) : '', pcs: T.cc >= 0 ? String(r[T.cc]) : '', status: st });
  }
  return { ok: true, list: list, formula: !!T.formula, statuses: LEFT_STATUSES };
}
function leftSheet_(thead) {
  var book = SpreadsheetApp.openById(USERS_SS_ID), sh = book.getSheetByName(LEFT_SHEET);
  var want = LEFT_EXTRA.concat(thead.filter(function (h) { return String(h).trim(); }));
  if (!sh) {
    sh = book.insertSheet(LEFT_SHEET);
    sh.getRange(1, 1, 1, want.length).setValues([want]).setFontWeight('bold').setBackground('#F4C7C3');
    sh.setFrozenRows(1);
    return { sh: sh, head: want };
  }
  var head = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(String), have = headIdx_(head), add = [];
  want.forEach(function (h) { if (have[norm_(h)] == null) add.push(h); });
  if (add.length) {
    if (sh.getMaxColumns() < head.length + add.length) sh.insertColumnsAfter(sh.getMaxColumns(), head.length + add.length - sh.getMaxColumns());
    sh.getRange(1, head.length + 1, 1, add.length).setValues([add]).setFontWeight('bold').setBackground('#F4C7C3');
    head = head.concat(add);
  }
  return { sh: sh, head: head };
}
/** Moves these roll numbers (must still be not Enrolled on the portal) to "Left Students" and starts the sheet clean-up. */
function leftMove_(user, rolls) {
  var scope = user === PORTAL_SYNC_USER ? { all: true } : scopeOf_(user);
  if (!scope || !scope.all) return { ok: false, error: 'Only the admin and the principal can move students.' };
  var want = {}; rolls.forEach(function (r) { want[enr_(r)] = 1; });
  if (!Object.keys(want).length) return { ok: false, error: 'Nobody chosen.' };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var T = transformRead_(); if (!T) return { ok: false, error: 'The strength sheet is empty.' };
    if (T.formula) return { ok: false, error: 'The strength sheet is made by a formula (QUERY / IMPORTRANGE): remove the students in its source sheet instead.' };
    var PS = portalStatusMap_(), tok = currentSeasonToken_() || SEASON[SEASON.length - 1];
    var now = Utilities.formatDate(new Date(), TZ_(), 'd MMM yyyy HH:mm'), moved = [], skipped = [], idx = [];
    for (var i = 1; i < T.V.length; i++) {
      var e = enr_(T.V[i][T.ec]); if (!e || !want[e]) continue;
      var st = PS[e] || '';
      if (!st || /^enrolled$/i.test(st)) { skipped.push(String(T.V[i][T.ec]) + ' (Enrolled on the portal now)'); continue; }
      idx.push(i); moved.push({ roll: String(T.V[i][T.ec]).trim(), status: st, row: T.V[i] });
    }
    if (!moved.length) return { ok: true, moved: 0, skipped: skipped };
    var L = leftSheet_(T.V[0]), LH = headIdx_(L.head);
    var out = moved.map(function (m) {
      var row = L.head.map(function () { return ''; });
      row[0] = now; row[1] = user.teacher; row[2] = m.status; row[3] = titleCase_(MONTH_NAMES[monthYear_(tok).m]) + ' ' + monthYear_(tok).y;
      T.V[0].forEach(function (h, c) { var k = LH[norm_(h)]; if (k != null && k >= LEFT_EXTRA.length) row[k] = m.row[c]; });
      return row;
    });
    var at = L.sh.getLastRow() + 1;
    L.sh.getRange(at, 1, out.length, L.head.length).setNumberFormat('@');
    L.sh.getRange(at, 1, out.length, L.head.length).setValues(out.map(function (r) { return r.map(function (x) { return x instanceof Date ? Utilities.formatDate(x, TZ_(), 'dd/MM/yyyy') : x; }); }));
    SpreadsheetApp.flush();
    // delete from the strength sheet, bottom up, neighbouring rows together
    idx.sort(function (a, b) { return b - a; });
    for (var k = 0; k < idx.length;) {
      var top = idx[k], n = 1; while (k + n < idx.length && idx[k + n] === top - n) n++;
      T.sh.deleteRows(top - n + 2, n); k += n;
    }
    SpreadsheetApp.flush();
    cacheClear_('master_att3');
    // background: strike through this month, remove from later months
    var p = PropertiesService.getScriptProperties(), job = {}; try { job = JSON.parse(p.getProperty('leftClean') || '{}'); } catch (e2) {}
    job.rolls = job.rolls || {}; moved.forEach(function (m) { job.rolls[enr_(m.roll)] = m.status; });
    job.tok = tok; job.running = true; job.struck = job.struck || 0; job.removed = job.removed || 0; job.at = Date.now();
    p.setProperty('leftClean', JSON.stringify(job));
    clearContinuation_('cleanLeftStudents');
    ScriptApp.newTrigger('cleanLeftStudents').timeBased().after(30 * 1000).create();
    return { ok: true, moved: moved.length, skipped: skipped };
  } finally { lock.releaseLock(); }
}

var PORTAL_SYNC_USER = { teacher: 'PORTAL SYNC', username: '' };
var LEFT_ADD_MAP = FILL_MAP.concat([['Enrollment#', ENROLL_HEADS, null], ['Class', ['CLASS'], null], ['Section', ['SECTION'], null]]);
function batchOf_(e) { var m = String(e || '').toUpperCase().match(/-F(\d{2})-/); return m ? m[1] : ''; }
function isLeftStatus_(st) { return LEFT_STATUSES.some(function (x) { return x.toUpperCase() === String(st || '').trim().toUpperCase(); }); }
function portalRows_() {
  var pd = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(PORTAL_SHEET);
  if (!pd || pd.getLastRow() < 2) return null;
  var v = pd.getRange(1, 1, pd.getLastRow(), Math.min(pd.getLastColumn(), PORTAL_COLS.length)).getValues(), C = {};
  v[0].forEach(function (h, i) { C[h] = i; });
  return { C: C, rows: v.slice(1) };
}
function leftRolls_() {
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(LEFT_SHEET), out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues(), ec = pickCol_(headIdx_(v[0]), ENROLL_HEADS);
  if (ec >= 0) for (var i = 1; i < v.length; i++) { var e = enr_(v[i][ec]); if (e) out[e] = 1; }
  return out;
}
/** v7.3: after a sync NOTHING is moved. Counts for this year's two batches + who could go to Left Students. */
function portalLeftPlan_() {
  try {
    var P = portalRows_(); if (!P) return { note: 'No portal data.' };
    var C = P.C, seen = {};
    P.rows.forEach(function (r) { if (/^enrolled$/i.test(String(r[C['Status']]).trim())) { var b = batchOf_(r[C['Enrollment#']]); if (b) seen[b] = (seen[b] || 0) + 1; } });
    var batches = Object.keys(seen).filter(function (b) { return seen[b] >= 10; }).sort().reverse().slice(0, 2), inB = {};
    batches.forEach(function (b) { inB[b] = 1; });
    var counts = { Enrolled: 0 }; LEFT_STATUSES.forEach(function (s) { counts[s] = 0; });
    var T = transformRead_(), inT = {}, L = leftRolls_(), move = [], add = [], already = 0;
    if (T && T.ec >= 0) for (var i = 1; i < T.V.length; i++) { var te = enr_(T.V[i][T.ec]); if (te) inT[te] = i; }
    P.rows.forEach(function (r) {
      var e = enr_(r[C['Enrollment#']]), st = String(r[C['Status']]).trim(); if (!e) return;
      var cur = !!inB[batchOf_(e)], left = isLeftStatus_(st);
      if (cur) { if (/^enrolled$/i.test(st)) counts.Enrolled++; else LEFT_STATUSES.forEach(function (s) { if (s.toUpperCase() === st.toUpperCase()) counts[s]++; }); }
      if (!left) return;
      var x = { roll: String(r[C['Enrollment#']]).trim(), name: String(r[C['Student Name']]).trim(), status: st };
      if (inT[e]) move.push(x);                                         // still in the strength sheet (any batch)
      else if (cur && !L[e]) add.push(x);                               // this year's batches, never in the strength sheet
      else if (cur) already++;
    });
    return { batches: batches.map(function (b) { return 'F' + b; }), counts: counts, move: move, add: add, already: already };
  } catch (e) { return { note: String(e && e.message || e) }; }
}
/** The sync panel's buttons: { key, move:[rolls] } or { key, add:[rolls] }. */
function portalLeftDo_(req) {
  var key = PropertiesService.getScriptProperties().getProperty('portalKey');
  if (!key || String(req.key || '') !== key) return { ok: false, error: 'The sync key is not correct. Get the key again in the app: College > UOP portal sync.' };
  var move = (req.move || []).map(String).slice(0, 500), add = (req.add || []).map(String).slice(0, 1000), out = { ok: true, moved: 0, added: 0, skipped: [] };
  if (move.length) { var r = leftMove_(PORTAL_SYNC_USER, move); if (!r.ok) return r; out.moved = r.moved; out.skipped = r.skipped || []; }
  if (add.length) out.added = leftAddFromPortal_(add);
  return out;
}
/** Students of the portal who left but were never in the strength sheet: one row each in Left Students (attendance untouched). */
function leftAddFromPortal_(rolls) {
  var want = {}; rolls.forEach(function (r) { want[enr_(r)] = 1; });
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var P = portalRows_(), T = transformRead_(); if (!P || !T) return 0;
    var C = P.C, have = leftRolls_(), inT = {};
    for (var i = 1; i < T.V.length; i++) { var te = enr_(T.V[i][T.ec]); if (te) inT[te] = 1; }
    var L = leftSheet_(T.V[0]), now = Utilities.formatDate(new Date(), TZ_(), 'd MMM yyyy HH:mm');
    var field = L.head.map(function (h, k) {
      if (k < LEFT_EXTRA.length) return null;
      var n = norm_(h);
      for (var j = 0; j < LEFT_ADD_MAP.length; j++) { var m = LEFT_ADD_MAP[j]; if (m[1].map(norm_).indexOf(n) >= 0 || (m[2] && norm_(m[2]) === n)) return m[0]; }
      return null;
    });
    var out = [];
    P.rows.forEach(function (r) {
      var e = enr_(r[C['Enrollment#']]), st = String(r[C['Status']]).trim();
      if (!e || !want[e] || have[e] || inT[e] || !isLeftStatus_(st)) return;
      have[e] = 1;
      var row = L.head.map(function (h, k) { var f = field[k]; return f && C[f] != null ? String(r[C[f]]) : ''; });
      row[0] = now; row[1] = 'PORTAL SYNC (not in strength sheet)'; row[2] = st; row[3] = '';
      out.push(row);
    });
    if (out.length) {
      var at = L.sh.getLastRow() + 1;
      L.sh.getRange(at, 1, out.length, L.head.length).setNumberFormat('@');
      L.sh.getRange(at, 1, out.length, L.head.length).setValues(out);
      SpreadsheetApp.flush();
    }
    return out.length;
  } finally { lock.releaseLock(); }
}




/** Background job (continues by itself): this month = strike through + off, later months = remove (or strike if marks exist). */
function cleanLeftStudents() {
  var p = PropertiesService.getScriptProperties(), job = null;
  try { job = JSON.parse(p.getProperty('leftClean') || 'null'); } catch (e) {}
  if (!job || (!Object.keys(job.rolls || {}).length && !Object.keys(job.restore || {}).length)) return;
  job.rolls = job.rolls || {};
  var cur = SEASON.indexOf(job.tok); if (cur < 0) cur = 0;
  var toks = SEASON.slice(cur), metas = allowed_({ username: ADMIN_USER, teacher: '' }), today = Utilities.formatDate(new Date(), TZ_(), 'd MMM yyyy');
  var restore = job.restore || {};
  var done = batchSheets_('cleanLeftStudents', function (sh) {
    var L = readLayout_(sh), n = sh.getLastRow() - FIRST_ROW + 1; if (n < 1) return;
    var meta = metas[sh.getName()], tokOf = String(sh.getName()).toUpperCase().split(/\s+/).filter(function (w) { return SEASON.indexOf(w) !== -1; })[0];
    var mi = meta ? SEASON.indexOf(meta.mon) : tokOf ? SEASON.indexOf(tokOf) : cur;          // later months are not in the teachers' list yet: the month comes from the sheet name
    var rg = sh.getRange(FIRST_ROW, L.rollCol, n, 1), rolls = rg.getValues(), notes = rg.getNotes(), del = [], changed = false, block = null;
    for (var i = 0; i < n; i++) {
      var e = enr_(rolls[i][0]); if (!e) continue;
      var row = FIRST_ROW + i, note = String(notes[i][0] || '');
      if (restore[e] && note.indexOf(LEFT_TAG) === 0) {                   // brought back: switch on again
        sh.getRange(row, L.rollCol).clearNote(); sh.getRange(row, 1, 1, 6).setBackground(null).setFontColor('#000000'); sh.getRange(row, 4).setFontLine('none'); changed = true; continue;
      }
      var st = job.rolls[e]; if (!st) continue;
      if (mi > cur) {                                                      // a later month: remove her if nothing was entered for her there
        if (!block) { var rc = dayColRange_(L); block = rc ? sh.getRange(FIRST_ROW, rc.c1, n, rc.c2 - rc.c1 + 1).getValues() : []; }
        var marks = (block[i] || []).some(function (x) { return String(x).trim(); }) || (L.mtCol > 0 && String(sh.getRange(row, L.mtCol).getValue()).trim());
        if (!marks) { del.push(row); continue; }
      }
      if (note.indexOf(LEFT_TAG) === 0) continue;                          // already struck through
      sh.getRange(row, L.rollCol).setNote(LEFT_TAG + ' - ' + st + ' on the portal (' + today + ')');
      sh.getRange(row, 1, 1, 6).setBackground('#F4C7C3').setFontColor('#999999');
      sh.getRange(row, 4).setFontLine('line-through');
      job.struck++; changed = true;
    }
    for (var k = del.length - 1; k >= 0; k--) { sh.deleteRow(del[k]); job.removed++; changed = true; }
    if (changed) { p.setProperty('leftClean', JSON.stringify(job)); if (meta) try { indexUpdate_(sh, meta); } catch (e3) {} }
  }, toks);
  if (done) {                                                              // finished: forget what was handled (keep anything added meanwhile)
    var now = null; try { now = JSON.parse(p.getProperty('leftClean') || 'null'); } catch (e4) {}
    if (now) {
      Object.keys(job.rolls).forEach(function (k) { if (now.rolls[k] === job.rolls[k]) delete now.rolls[k]; });
      Object.keys(restore).forEach(function (k) { if (now.restore) delete now.restore[k]; });
      now.running = Object.keys(now.rolls).length > 0 || Object.keys(now.restore || {}).length > 0;
      now.struck = job.struck; now.removed = job.removed; now.finished = Utilities.formatDate(new Date(), TZ_(), 'd MMM yyyy HH:mm');
      p.setProperty('leftClean', JSON.stringify(now));
      if (now.running) ScriptApp.newTrigger('cleanLeftStudents').timeBased().after(60 * 1000).create();
    }
  }
}
/** The Left Students sheet for the app (admin, principal) + the clean-up job. */
function leftList_(user) {
  var scope = scopeOf_(user);
  if (!scope || !scope.all) return { ok: false, error: 'NOT_ALLOWED' };
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(LEFT_SHEET), out = [];
  if (sh && sh.getLastRow() > 1) {
    var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues(), H = headIdx_(v[0]);
    var c = function (list) { return pickCol_(H, list.map(norm_)); }, ec = c(ENROLL_HEADS), nc = c(FILL_MAP[0][1]), fc = c(FILL_MAP[1][1]), cc = c(CLASS_HEADS);
    for (var i = 1; i < v.length; i++) out.push({ roll: ec >= 0 ? String(v[i][ec]) : '', name: nc >= 0 ? String(v[i][nc]) : '', father: fc >= 0 ? String(v[i][fc]) : '', pcs: cc >= 0 ? String(v[i][cc]) : '',
      status: String(v[i][2]), on: String(v[i][0]), by: String(v[i][1]), month: String(v[i][3]) });
  }
  var job = null; try { job = JSON.parse(PropertiesService.getScriptProperties().getProperty('leftClean') || 'null'); } catch (e) {}
  return { ok: true, list: out.reverse(), job: job };
}
/** Puts one student back into the strength sheet (moved by mistake) and switches her on again in this month's sheets. */
function leftRestore_(user, roll) {
  var scope = scopeOf_(user);
  if (!scope || !scope.all) return { ok: false, error: 'NOT_ALLOWED' };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var L = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(LEFT_SHEET), T = transformRead_();
    if (!L || !T) return { ok: false, error: 'Not found.' };
    var v = L.getRange(1, 1, L.getLastRow(), L.getLastColumn()).getValues(), LH = headIdx_(v[0]), lec = pickCol_(LH, ENROLL_HEADS), at = -1;
    for (var i = v.length - 1; i >= 1; i--) if (enr_(v[i][lec]) === enr_(roll)) { at = i; break; }
    if (at < 0) return { ok: false, error: 'Not found in Left Students.' };
    var row = T.V[0].map(function (h) { var k = LH[norm_(h)]; return k != null && k >= LEFT_EXTRA.length ? v[at][k] : ''; });
    var to = T.sh.getLastRow() + 1;
    T.sh.getRange(to, 1, 1, row.length).setNumberFormat('@').setValues([row]);
    L.deleteRow(at + 1);
    cacheClear_('master_att3');
    var p = PropertiesService.getScriptProperties(), job = {}; try { job = JSON.parse(p.getProperty('leftClean') || '{}'); } catch (e) {}
    job.rolls = job.rolls || {}; delete job.rolls[enr_(roll)]; job.restore = job.restore || {}; job.restore[enr_(roll)] = 1;
    job.tok = currentSeasonToken_() || job.tok; job.running = true; job.struck = job.struck || 0; job.removed = job.removed || 0;
    p.setProperty('leftClean', JSON.stringify(job));
    clearContinuation_('cleanLeftStudents');
    ScriptApp.newTrigger('cleanLeftStudents').timeBased().after(30 * 1000).create();
    return { ok: true, roll: roll };
  } finally { lock.releaseLock(); }
}

/* v7.2 student page: the day-by-day attendance of one month (P / A / L / B per subject and date), read when she taps a month */
function studentDays_(roll, code, mon) {
  var chk = studentReport_(roll, code);                                // same roll no + code check (and wrong-code limit)
  if (!chk.ok) return chk;
  mon = String(mon || '').toUpperCase(); if (SEASON.indexOf(mon) === -1) return { ok: false, error: 'Choose a month.' };
  var key = enr_(chk.roll), out = [], found = 0, errs = [], seen = {}, sq = function (x) { return String(x || '').replace(/\s+/g, '').toUpperCase(); };
  // v7.3.2: her sheets are found DIRECTLY from her section + her subjects (no need for the Attendance Index) …
  var subs = {}, mcls = chk.cls; (chk.rows || []).forEach(function (x) { if (x[2] === mon) subs[String(x[0]).toUpperCase()] = 1; });
  // her subjects and section from the strength sheet too (the result page may be missing a month)
  getMaster_().rows.forEach(function (m) { if (enr_(m[0]) === key) { mcls = m[2] || mcls; (m[3] || []).forEach(function (x) { subs[String(x).toUpperCase()] = 1; }); } });
  var cand = getAssignments_({ teacher: '', username: ADMIN_USER }).filter(function (a) {
    var b = sq(baseClass_(a.cls)), c = sq(mcls);
    return subs[String(a.subject).toUpperCase()] && (b === c || c.indexOf(b) === 0);
  }).map(function (a) { return { name: makeName_(a.cls, mon, a.subject), mon: mon, cls: a.cls, sub: a.subject, teacher: a.teacher }; });
  // … plus any sheet the index lists her in (a guest in another section)
  regIndex_().forEach(function (r) { if (r.mon === mon && r.json.toUpperCase().replace(/\s+/g, '').indexOf('"' + key + '"') !== -1) cand.push(r); });
  cand.forEach(function (r) {
    if (seen[r.name]) return; seen[r.name] = 1;
    found++;
    try { var x = regSheetDays_(r, chk.roll); if (x && x.st && x.st.length) out.push({ sub: x.sub, teacher: x.teacher, m: x.m, y: x.y, days: x.days, held: x.held, cells: x.st[0][2], off: x.st[0][3] }); else errs.push(r.name + ': ' + (x && x.missing ? 'sheet not found' : 'roll no. not found in the sheet')); }
    catch (e) { errs.push(r.name + ': ' + (e.message || e)); }
  });
  return { ok: true, mon: mon, subjects: out, note: out.length ? '' : (found ? found + ' sheet(s) checked, she is not in them: ' + errs.slice(0, 3).join('; ') : 'No sheet found for section ' + chk.cls + ' in ' + mon + ' (subjects: ' + Object.keys(subs).join(', ') + ').') };
}

/* =====================================================================================================
   v7.3  SHEETS vs STRENGTH SHEET
   Compares every attendance sheet of a month with Transform-ALL-COLLEGE-STRENGTH (same rule as when a sheet is made:
   students of that Class-Sec who take that subject; guests are allowed). Nothing changes until the admin presses
   Add / Remove / Switch on for a student.
   ===================================================================================================== */
var ROSTER_MS = 240000;                                                // one call works at most 4 minutes, then the app asks for the rest

/** roll -> master row, and "class|subject" -> [rolls], from the strength sheet */
function rosterModel_() {
  var M = getMaster_(), byRoll = {}, sq = function (x) { return String(x || '').replace(/\s+/g, '').toUpperCase(); };
  M.rows.forEach(function (r) { byRoll[enr_(r[0])] = r; });
  var want = function (cls, sub) {
    var base = sq(baseClass_(cls)), opt = base !== sq(cls), out = [];
    M.rows.forEach(function (r) { var c = sq(r[2]); if ((opt ? (c === base || c.indexOf(base) === 0) : c === base) && r[3].indexOf(String(sub).toUpperCase()) !== -1) out.push(r); });
    return out;
  };
  return { byRoll: byRoll, want: want };
}
/** The differences of one sheet: missing / extra / switched off but in the strength sheet */
function rosterDiff_(sh, meta, RM) {
  var L = readLayout_(sh), n = sh.getLastRow() - FIRST_ROW + 1;
  var rolls = n > 0 ? sh.getRange(FIRST_ROW, L.rollCol, n, 1).getValues() : [], notes = n > 0 ? sh.getRange(FIRST_ROW, L.rollCol, n, 1).getNotes() : [];
  var names = n > 0 && L.nameCol > 0 ? sh.getRange(FIRST_ROW, L.nameCol, n, 1).getValues() : [];
  var rc = dayColRange_(L), block = n > 0 && rc ? sh.getRange(FIRST_ROW, rc.c1, n, rc.c2 - rc.c1 + 1).getValues() : [];
  var mts = n > 0 && L.mtCol > 0 ? sh.getRange(FIRST_ROW, L.mtCol, n, 1).getValues() : [];
  var want = RM.want(meta.cls, meta.sub), wantSet = {}, inSheet = {};
  want.forEach(function (r) { wantSet[enr_(r[0])] = r; });
  var guests = {}; try { activeGuests_(meta.cls, meta.sub, meta.mon).forEach(function (g) { guests[enr_(g.roll)] = 1; }); } catch (e) {}
  var d = { sheet: meta.name, mon: meta.mon, cls: meta.cls, sub: meta.sub, teacher: meta.teacher || '', missing: [], extra: [], off: [] };
  for (var i = 0; i < n; i++) {
    var e = enr_(rolls[i][0]); if (!e) continue;
    var note = String(notes[i][0] || ''), off = isOffNote_(note), guest = note.indexOf('GUEST|') === 0;
    var marks = (block[i] || []).some(function (x) { return String(x).trim(); }) || (mts[i] && String(mts[i][0]).trim() !== '');
    inSheet[e] = 1;
    var m = RM.byRoll[e], who = { roll: String(rolls[i][0]).trim(), name: String((names[i] || [])[0] || (m ? m[1] : '')).trim(), father: m ? m[4] || '' : '', pcs: m ? m[2] : '', marks: !!marks };
    if (wantSet[e]) { if (off && note.indexOf('guest removed') === -1) d.off.push(Object.assign(who, { why: note.split(' (')[0] })); }
    else if (!off && !guest && !guests[e]) d.extra.push(Object.assign(who, { why: m ? 'in the strength sheet as ' + m[2] + (m[3].length ? ' (' + m[3].join(', ') + ')' : '') : 'not in the strength sheet' }));
  }
  want.forEach(function (r) { if (!inSheet[enr_(r[0])]) d.missing.push({ roll: r[0], name: r[1], father: r[4] || '', pcs: r[2] }); });
  return d;
}
/** mon = 'OCT' … ; cursor = where the last call stopped. Returns only the sheets with differences. */
function rosterCheck_(user, mon, cursor) {
  var scope = scopeOf_(user);
  if (!scope || !scope.all) return { ok: false, error: 'Only the admin and the principal can do this.' };
  mon = String(mon || currentSeasonToken_() || '').toUpperCase();
  if (SEASON.indexOf(mon) === -1) return { ok: false, error: 'Choose a month.' };
  var start = Date.now(), book = bookFor_(mon), RM = rosterModel_();
  var A = getAssignments_({ teacher: '', username: ADMIN_USER }).map(function (a) { return { name: makeName_(a.cls, mon, a.subject), cls: a.cls, sub: a.subject, mon: mon, teacher: a.teacher }; });
  if (!Number(cursor)) {                                                 // v7.4 fast: whole month in 2 calls (Sheets advanced service)
    try { var fast = rosterCheckFast_(book, A, RM, mon); if (fast) return fast; }
    catch (e) { Logger.log('Fast roster check not used: ' + e.message); }
  }
  var i = Number(cursor) || 0, diffs = [], made = 0;
  for (; i < A.length; i++) {
    if (Date.now() - start > ROSTER_MS) break;
    var sh = book.getSheetByName(A[i].name); if (!sh) continue;
    made++;
    try { var d = rosterDiff_(sh, A[i], RM); if (d.missing.length || d.extra.length || d.off.length) diffs.push(d); }
    catch (e) { diffs.push({ sheet: A[i].name, mon: mon, cls: A[i].cls, sub: A[i].sub, teacher: A[i].teacher, missing: [], extra: [], off: [], error: String(e.message || e) }); }
  }
  var my = monthYear_(mon);
  return { ok: true, mon: mon, label: titleCase_(MONTH_NAMES[my.m]) + ' ' + my.y, total: A.length, cursor: i, done: i >= A.length, checked: made, diffs: diffs };
}
/** v7.4: every sheet of the month read with 2 Sheets API calls per 60 sheets (values + Roll No notes), compared in memory.
 *  Also lists rows whose Student Name / Father Name differ from the strength sheet (d.wrong). Returns null if a sheet looks different. */
function rosterCheckFast_(book, A, RM, mon) {
  var have = {}; book.getSheets().forEach(function (s) { have[s.getName()] = true; });
  var list = A.filter(function (a) { return have[a.name]; }), id = book.getId(), V = {}, N = {};
  var q = function (n) { return "'" + n.replace(/'/g, "''") + "'"; };
  for (var k = 0; k < list.length; k += 60) {
    var part = list.slice(k, k + 60).map(function (a) { return a.name; });
    var res = Sheets.Spreadsheets.Values.batchGet(id, { ranges: part.map(function (n) { return q(n) + '!A1:AP'; }) });
    (res.valueRanges || []).forEach(function (vr, j) { V[part[j]] = vr.values || []; });
    var g = Sheets.Spreadsheets.get(id, { ranges: part.map(function (n) { return q(n) + '!B' + FIRST_ROW + ':B'; }), fields: 'sheets(properties(title),data(rowData(values(note))))' });
    (g.sheets || []).forEach(function (s) {
      var rd = (s.data && s.data[0] && s.data[0].rowData) || [];
      N[s.properties.title] = rd.map(function (r) { return (r.values && r.values[0] && r.values[0].note) || ''; });
    });
  }
  var diffs = [], made = 0;
  for (var i = 0; i < list.length; i++) {
    var d = rosterDiffV_(V[list[i].name] || [], N[list[i].name] || [], list[i], RM);
    if (!d) return null;                                                 // unusual layout: let the old way do it
    made++;
    if (d.missing.length || d.extra.length || d.off.length || d.wrong.length) diffs.push(d);
  }
  var my = monthYear_(mon);
  return { ok: true, mon: mon, label: titleCase_(MONTH_NAMES[my.m]) + ' ' + my.y, total: A.length, cursor: A.length, done: true, checked: made, diffs: diffs, names: true };
}
/** rosterDiff_ from values already read (v = sheet from A1, notes = Roll No notes from row FIRST_ROW) */
function rosterDiffV_(v, notes, meta, RM) {
  var hdr = v[HEADER_ROW - 1] || [], rollCol = -1, nameCol = -1, fatherCol = -1, mtCol = -1, dayIdx = [];
  for (var c = 0; c < hdr.length; c++) {
    var t = norm_(hdr[c]);
    if (t === 'ROLLNO' && rollCol < 0) rollCol = c;
    else if (t === 'STUDENTNAME' && nameCol < 0) nameCol = c;
    else if (t === 'FATHERNAME' && fatherCol < 0) fatherCol = c;
    else if ((t === 'M.T' || t === 'MT') && mtCol < 0) mtCol = c;
    else if (/^\d+$/.test(t) && Number(t) >= 1 && Number(t) <= 31) dayIdx.push(c);
  }
  if (rollCol !== 1) return null;                                       // the notes were read from column B only
  var nk = function (s) { return String(s == null ? '' : s).toLowerCase().replace(/[^a-z]/g, ''); };
  var want = RM.want(meta.cls, meta.sub), wantSet = {}, inSheet = {};
  want.forEach(function (r) { wantSet[enr_(r[0])] = r; });
  var guests = {}; try { activeGuests_(meta.cls, meta.sub, meta.mon).forEach(function (g) { guests[enr_(g.roll)] = 1; }); } catch (e) {}
  var d = { sheet: meta.name, mon: meta.mon, cls: meta.cls, sub: meta.sub, teacher: meta.teacher || '', missing: [], extra: [], off: [], wrong: [] };
  for (var i = FIRST_ROW - 1; i < v.length; i++) {
    var r = v[i] || [], e = enr_(r[rollCol]); if (!e) continue;
    var note = String(notes[i - FIRST_ROW + 1] || ''), off = isOffNote_(note), guest = note.indexOf('GUEST|') === 0;
    var marks = dayIdx.some(function (c) { return String(r[c] == null ? '' : r[c]).trim(); }) || (mtCol >= 0 && String(r[mtCol] == null ? '' : r[mtCol]).trim() !== '');
    inSheet[e] = 1;
    var m = RM.byRoll[e], nameNow = String(nameCol >= 0 && r[nameCol] != null ? r[nameCol] : '').trim();
    var who = { roll: String(r[rollCol]).trim(), name: nameNow || (m ? m[1] : ''), father: m ? m[4] || '' : '', pcs: m ? m[2] : '', marks: !!marks };
    if (wantSet[e]) {
      if (off && note.indexOf('guest removed') === -1) d.off.push(Object.assign(who, { why: note.split(' (')[0] }));
      else if (!off && m) {
        var fNow = String(fatherCol >= 0 && r[fatherCol] != null ? r[fatherCol] : '').trim();
        var badN = nk(nameNow) !== nk(m[1]), badF = fatherCol >= 0 && String(m[4] || '').trim() && nk(fNow) !== nk(m[4]);
        if (badN || badF) d.wrong.push({ roll: who.roll, name: m[1], father: m[4] || '', pcs: m[2], nameNow: nameNow, fatherNow: fNow, badName: badN, badFather: !!badF });
      }
    }
    else if (!off && !guest && !guests[e]) d.extra.push(Object.assign(who, { why: m ? 'in the strength sheet as ' + m[2] + (m[3].length ? ' (' + m[3].join(', ') + ')' : '') : 'not in the strength sheet' }));
  }
  want.forEach(function (r) { if (!inSheet[enr_(r[0])]) d.missing.push({ roll: r[0], name: r[1], father: r[4] || '', pcs: r[2] }); });
  return d;
}
/** op: 'add' (missing student), 'remove' (extra student: row deleted when she has no marks, otherwise struck through), 'on' (switch on again) */
function rosterFix_(user, name, roll, op) {
  var scope = scopeOf_(user);
  if (!scope || !scope.all) return { ok: false, error: 'Only the admin and the principal can do this.' };
  var meta = null; SEASON.forEach(function (t) { if (!meta && String(name).indexOf(' ' + t + ' ') !== -1) meta = t; });
  var A = getAssignments_({ teacher: '', username: ADMIN_USER }).filter(function (a) { return makeName_(a.cls, meta, a.subject) === name; })[0];
  if (!meta || !A) return { ok: false, error: 'Sheet not found in the timetable: ' + name };
  meta = { name: name, cls: A.cls, sub: A.subject, mon: meta, teacher: A.teacher };
  var sh = bookFor_(meta.mon).getSheetByName(name); if (!sh) return { ok: false, error: 'Sheet not found: ' + name };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var L = readLayout_(sh), n = sh.getLastRow() - FIRST_ROW + 1, rolls = n > 0 ? sh.getRange(FIRST_ROW, L.rollCol, n, 1).getValues() : [], row = 0;
    for (var i = 0; i < n; i++) if (enr_(rolls[i][0]) === enr_(roll)) { row = FIRST_ROW + i; break; }
    var today = Utilities.formatDate(new Date(), TZ_(), 'd MMM yyyy'), res = '';
    if (op === 'add') {
      var m = rosterModel_().byRoll[enr_(roll)]; if (!m) return { ok: false, error: roll + ' is not in the strength sheet.' };
      putGuestRow_(sh, String(m[0]).trim(), m, user.teacher, today);                // adds the row (same look and formula) …
      var L2 = readLayout_(sh), n2 = sh.getLastRow() - FIRST_ROW + 1, r2 = sh.getRange(FIRST_ROW, L2.rollCol, n2, 1).getValues();
      for (var k = 0; k < n2; k++) if (enr_(r2[k][0]) === enr_(roll)) { sh.getRange(FIRST_ROW + k, L2.rollCol).clearNote(); break; }   // … as a normal student, not a guest
      res = 'added';
    } else if (op === 'on') {
      if (!row) return { ok: false, error: 'Not in this sheet.' };
      sh.getRange(row, L.rollCol).clearNote(); sh.getRange(row, 1, 1, 6).setBackground(null).setFontColor('#000000'); sh.getRange(row, 4).setFontLine('none');
      res = 'switched on';
    } else if (op === 'remove') {
      if (!row) return { ok: false, error: 'Not in this sheet.' };
      var rc = dayColRange_(L), marks = rc && sh.getRange(row, rc.c1, 1, rc.c2 - rc.c1 + 1).getValues()[0].some(function (x) { return String(x).trim(); });
      if (!marks && L.mtCol > 0 && String(sh.getRange(row, L.mtCol).getValue()).trim() !== '') marks = true;
      if (marks) {
        sh.getRange(row, L.rollCol).setNote(OFF_TAG + ' - not in the strength sheet (' + today + ')');
        sh.getRange(row, 1, 1, 6).setBackground('#EFEFEF').setFontColor('#999999'); sh.getRange(row, 4).setFontLine('line-through');
        res = 'struck through (she has marks this month)';
      } else { sh.deleteRow(row); res = 'removed'; }
    } else return { ok: false, error: 'Unknown action.' };
    SpreadsheetApp.flush();
    try { updateLectures_(sh, readLayout_(sh), sh.getLastRow() - FIRST_ROW + 1); indexUpdate_(sh, meta); } catch (e) {}
    try { AUDIT_.push([new Date(), user.teacher, user.username, name, meta.mon, meta.cls, meta.sub, '', '', String(roll), '', '', res, 'Sheet roster']); auditFlush_(); } catch (e) {}
    return { ok: true, sheet: name, roll: roll, op: op, result: res };
  } finally { lock.releaseLock(); }
}