/**
 * Leaves.gs — Approved leaves (Leave In-Charge) inside the attendance app (page leaves.html).
 *
 * This is the old stand-alone "JCW Approved Leaves" backend moved into the attendance project, so the
 * Leave In-Charge signs in with her normal attendance-app account (no second login, no second deployment).
 * Leaves are still written to the same "ApprovedLeavesStore" sheet of each month workbook (LL_STORES in
 * Leavelink.gs), so nothing changes in the sheets, the shortage reports or the old web page.
 * After every save the Leave Link cache of that month is cleared, so teachers see the new L at once.
 *
 * Who can use it: the admin, and anyone whose row in the Users sheet has the role
 * Admin / Administrator / Leave In-Charge (any spelling with or without the dash).
 *
 * INSTALL:
 *   a) + > Script > name it "Leaves" > paste this file.
 *   b) doPost, next to the bnk_ line (once):
 *        if (/^lv_/.test(String(req.action || ''))) return json_(lvRoute_(user, req));
 *   c) Run lvAuthorize once (allows Drive for the application photos).
 *   d) Deploy > Manage deployments > Edit > New version > Deploy.
 */
var LV_CFG = {
  BOOK: '1ARU8zMO2daJUfU_SIIo4baeSOG7qBPuni5TYRIRLBNY',        // Users workbook (Users + Transform-ALL-COLLEGE-STRENGTH + Portal backup)
  STUDENTS: 'Transform-ALL-COLLEGE-STRENGTH',
  STORE: 'ApprovedLeavesStore',
  CONTACT_SHEET: 'Portal backup',
  CONTACT_HEADERS: ['StudentContact'],                       // [] = every column whose header says mobile / phone / contact
  PROOF_FOLDER_ID: '',                                       // '' = folder "JCW Leave Proofs" in the owner's Drive
  PROOF_FOLDER_NAME: 'JCW Leave Proofs',
  NO_NUMBER_TYPES: ['Emergency Leave', 'Sick Leave']
};
var LV_MONTHS = ['September', 'October', 'November', 'December', 'January'];
var LV_TOK = { September: 'SEPT', October: 'OCT', November: 'NOV', December: 'DEC', January: 'JAN' };
var LV_ROLES = ['admin', 'administrator', 'leave in-charge', 'leave incharge', 'leave-in-charge', 'leave in charge'];
var LV_COLS = 12;   // A S.No, B Enroll#, C Student, D Father, E Class-Sec, F Approved Leaves, G Teacher, H Department,
                    // I username, J Leave Type (type \n short dates), K Leave Dates, L Proof links

var LV_MEMO_ = { store: {}, heads: {}, rows: null, map: null, lastRow: 0 };

function lvAuthorize() {
  SpreadsheetApp.openById(LV_CFG.BOOK).getName();
  lvProofFolder_('September').getName();
  LV_MONTHS.forEach(function (m) { SpreadsheetApp.openById(LL_STORES[LV_TOK[m]]).getName(); });
  Logger.log('Approved leaves: permissions OK. Now Deploy → Manage deployments → Edit → New version.');
}

function lvRoute_(user, req) {
  try {
    var a = String(req.action || '');
    var op = lvOperator_(user);
    if (a === 'lv_whoami') return { ok: true, allowed: !!op, teacher: user.teacher || '', username: user.username || '' };
    if (!op) return { ok: false, code: 'DENIED', error: 'Approved leaves are only for the Admin and the Leave In-Charge.' };
    if (a === 'lv_data') return lvData_(op, req);
    if (a === 'lv_sync') return lvSync_(op, req.ops || []);
    if (a === 'lv_attendance') return { ok: true, at: Date.now(), months: lvAttendance_(req.enroll, req.classSec, req.months) };
    return { ok: false, error: 'UNKNOWN_ACTION' };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}

/** {teacher, username, department} if this user may enter leaves, else null (cached 10 min). */
function lvOperator_(user) {
  var u = String(user && user.username || '').trim().toLowerCase(); if (!u) return null;
  var cache = CacheService.getScriptCache(), ck = 'lvop_' + u, hit = cache.get(ck);
  if (hit) return hit === '-' ? null : JSON.parse(hit);
  var out = null, admin = false;
  try { admin = !!isAdmin_(user); } catch (e) {}
  var v = SpreadsheetApp.openById(LV_CFG.BOOK).getSheetByName('Users').getDataRange().getDisplayValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][4] || '').trim().toLowerCase() !== u) continue;
    var role = v[i].slice(0, 12).some(function (x, k) {
      return k !== 4 && k !== 5 && LV_ROLES.indexOf(String(x || '').trim().toLowerCase()) > -1;
    });
    if (admin || role) out = { teacher: String(v[i][0] || user.teacher || '').trim(), username: String(v[i][4]).trim(), department: String(v[i][3] || '').trim() };
    break;
  }
  if (!out && admin) out = { teacher: user.teacher || user.username, username: user.username, department: '' };
  cache.put(ck, out ? JSON.stringify(out) : '-', 600);
  return out;
}

/* ================= DATA FOR THE PHONE ================= */
function lvData_(op, req) {
  var out = { ok: true, at: Date.now(), teacher: op.teacher, username: op.username, months: {} };
  if (req.students) out.students = lvAllStudents_();
  (req.months || LV_MONTHS).forEach(function (m) { if (LV_MONTHS.indexOf(m) > -1) out.months[m] = lvMonthRows_(m); });
  return out;
}

function lvKey_(h) { return String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }
function lvFindSheet_(ss, name) {
  var s = ss.getSheetByName(name); if (s) return s;
  var want = lvKey_(name), all = ss.getSheets();
  for (var i = 0; i < all.length; i++) if (lvKey_(all[i].getName()) === want) return all[i];
  return null;
}

/** Transform-ALL-COLLEGE-STRENGTH, columns B..I: r[0] Enroll, r[1] Student, r[2] Father, r[7] Class-Sec */
function lvStudentRows_() {
  if (LV_MEMO_.rows) return LV_MEMO_.rows;
  var sh = SpreadsheetApp.openById(LV_CFG.BOOK).getSheetByName(LV_CFG.STUDENTS);
  if (!sh) throw new Error('Sheet "' + LV_CFG.STUDENTS + '" not found.');
  var last = sh.getLastRow();
  LV_MEMO_.rows = last < 2 ? [] : sh.getRange(2, 2, last - 1, 8).getDisplayValues();
  return LV_MEMO_.rows;
}
function lvStudentMap_() {
  if (LV_MEMO_.map) return LV_MEMO_.map;
  var m = {};
  lvStudentRows_().forEach(function (r) {
    var e = String(r[0] || '').trim();
    if (e && !m[e]) m[e] = { name: r[1], father: r[2], classSec: r[7] };
  });
  LV_MEMO_.map = m;
  return m;
}

/** "3146616133" -> "0314-6616133"; several numbers in one cell are each fixed. */
function lvFixNumbers_(cell) {
  return String(cell == null ? '' : cell).split(/\s*[\/,;|]\s*/).map(function (part) {
    var d = part.replace(/\D/g, '');
    if (!d) return '';
    if (d.indexOf('0092') === 0) d = d.slice(4);
    else if (d.indexOf('92') === 0 && d.length === 12) d = d.slice(2);
    if (d.length === 10 && d.charAt(0) === '3') d = '0' + d;
    if (d.length === 11 && d.indexOf('03') === 0) return d.slice(0, 4) + '-' + d.slice(4);
    return part.trim();
  }).filter(function (x) { return x && x.replace(/\D/g, '').length >= 7; }).join(' / ');
}

function lvContacts_() {
  var out = { labels: [], map: {}, note: '' };
  try {
    var sh = lvFindSheet_(SpreadsheetApp.openById(LV_CFG.BOOK), LV_CFG.CONTACT_SHEET);
    if (!sh) { out.note = 'Tab "' + LV_CFG.CONTACT_SHEET + '" was not found.'; return out; }
    var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow < 2) { out.note = 'Tab "' + sh.getName() + '" is empty.'; return out; }
    var top = sh.getRange(1, 1, Math.min(5, lastRow), lastCol).getDisplayValues();
    var hr = top.findIndex(function (row) { return row.some(function (h) { return /enrol/i.test(h); }); });
    if (hr < 0) hr = 0;
    var head = top[hr].map(function (h) { return String(h || '').trim(); });
    var ec = head.findIndex(function (h) { return /enrol/i.test(h); }); if (ec < 0) ec = 1;
    var wanted = LV_CFG.CONTACT_HEADERS.map(lvKey_), cols = [];
    head.forEach(function (h, i) {
      if (!h || i === ec) return;
      if (wanted.length ? wanted.indexOf(lvKey_(h)) > -1 : /(mobile|phone|contact|cell|whats\s*app)/i.test(h)) cols.push(i);
    });
    if (!cols.length) { out.note = 'No contact column in tab "' + sh.getName() + '".'; return out; }
    var n = lastRow - hr - 1; if (n < 1) return out;
    var enr = sh.getRange(hr + 2, ec + 1, n, 1).getDisplayValues();
    var vals = cols.map(function (c) { return sh.getRange(hr + 2, c + 1, n, 1).getDisplayValues(); });
    out.labels = cols.map(function (i) { return head[i].replace(/([a-z])([A-Z])/g, '$1 $2'); });
    for (var r = 0; r < n; r++) {
      var e = String(enr[r][0] || '').trim();
      if (!e || out.map[e]) continue;
      var nums = vals.map(function (v) { return lvFixNumbers_(v[r][0]); });
      if (nums.some(Boolean)) out.map[e] = nums;
    }
  } catch (err) { out.note = 'Could not read contacts: ' + (err && err.message || err); }
  return out;
}

/** { labels, list: [[enroll, student, father, classSec, [numbers]]], note } */
function lvAllStudents_() {
  var ct = lvContacts_(), list = [];
  lvStudentRows_().forEach(function (r) {
    var e = String(r[0] || '').trim(); if (!e) return;
    list.push([e, r[1], r[2], r[7], ct.map[e] || []]);
  });
  return { labels: ct.labels, list: list, note: ct.note };
}

/* ================= MONTH STORES ================= */
function lvStore_(month, create) {
  if (LV_MEMO_.store[month]) return LV_MEMO_.store[month];
  var id = LL_STORES[LV_TOK[month]];
  if (!id) throw new Error('Unknown month: ' + month);
  var ss = SpreadsheetApp.openById(id), sh = ss.getSheetByName(LV_CFG.STORE);
  if (!sh && create) sh = ss.insertSheet(LV_CFG.STORE);
  if (sh) LV_MEMO_.store[month] = sh;
  return sh;
}
function lvHeaders_(sh, month) {
  if (LV_MEMO_.heads[month]) return;
  if (sh.getMaxColumns() < LV_COLS) sh.insertColumnsAfter(sh.getMaxColumns(), LV_COLS - sh.getMaxColumns());
  sh.getRange(1, 1, 1, LV_COLS).setValues([['S.No', 'Enroll#', 'Student Name', 'Father Name', 'Class-Sec',
    'Approved Leaves', 'Teacher', 'Department', 'username', 'Leave Type', 'Leave Dates', 'Proof']]);
  sh.getRange(1, 10, Math.max(sh.getMaxRows(), 2), 2).setWrap(true);
  LV_MEMO_.heads[month] = true;
}

function lvMonthRows_(month) {
  var sh = lvStore_(month, false);
  if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, Math.min(LV_COLS, sh.getMaxColumns())).getDisplayValues(), out = [];
  rows.forEach(function (r) {
    var enroll = String(r[1] || '').trim(); if (!enroll) return;
    var lines = String(r[9] || '').split(/\r?\n/), type = (lines[0] || '').trim(), short = (lines[1] || '').trim();
    var full = String(r[10] || '').trim();
    if (!short) { var m = /^(.*?)\s*\((\d[^)]*)\)$/.exec(type); if (m) { type = m[1].trim(); short = m[2].trim(); } }
    out.push({ enroll: enroll, student: r[2], father: r[3], classSec: r[4], leaves: r[5], teacher: r[6], department: r[7],
      leaveType: type, datesShort: short || full, datesFull: full || short, datesK: full,
      proofs: String(r[11] || '').split(/\r?\n/).map(function (x) { return x.trim(); }).filter(Boolean) });
  });
  out.sort(function (a, b) { return String(a.classSec).localeCompare(String(b.classSec)) || String(a.student).localeCompare(String(b.student)); });
  return out;
}

/* ================= SAVE / REMOVE (queue from the phone) ================= */
function lvSync_(op, ops) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  var results = [], touched = {};
  try {
    ops.slice(0, 40).forEach(function (o) {
      try {
        if (o.type === 'delete') lvDelete_(o.enroll, o.month);
        else {
          lvSave_(op, o.enroll, o.month, o.leaves, o.leaveType, o.datesFull, o.datesShort);
          if (o.photos && o.photos.length) {
            var st = lvStudentMap_()[String(o.enroll).trim()] || {};
            lvAttachProofs_(o.month, LV_MEMO_.lastRow, o.enroll, st.name, o.photos, o.id);
          }
        }
        touched[o.month] = 1;
        results.push({ id: o.id, ok: true });
      } catch (e) {
        results.push({ id: o.id, ok: false, error: String(e && e.message || e) });
      }
    });
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  // teachers' attendance app shows the new L at once (Leavelink.gs cache)
  Object.keys(touched).forEach(function (m) { try { cacheClear_('ll2_' + LV_TOK[m]); } catch (e) {} });
  return { ok: true, results: results };
}

function lvSave_(op, enroll, month, approved, leaveType, dates, datesShort) {
  enroll = String(enroll || '').trim();
  approved = String(approved == null ? '' : approved).trim();
  leaveType = String(leaveType == null ? '' : leaveType).trim();
  dates = String(dates == null ? '' : dates).trim();
  datesShort = String(datesShort == null ? '' : datesShort).trim();
  if (!enroll) throw new Error('Enrollment number is required.');
  if (LV_MONTHS.indexOf(month) === -1) throw new Error('Please select a month.');
  if (!leaveType) throw new Error('Leave type is required.');
  if (approved !== '' && !/^\d+(\.\d+)?$/.test(approved)) throw new Error('Approved leaves must be a number.');
  if (approved === '' && !dates && LV_CFG.NO_NUMBER_TYPES.indexOf(leaveType) === -1) throw new Error('Number of approved leaves is required for ' + leaveType + '.');
  var st = lvStudentMap_()[enroll];
  if (!st) throw new Error('Enroll# ' + enroll + ' not found in ' + LV_CFG.STUDENTS);
  var typeLabel = datesShort ? leaveType + '\n' + datesShort : leaveType;
  var sh = lvStore_(month, true);
  lvHeaders_(sh, month);
  var last = sh.getLastRow();
  if (last > 1) {
    var en = sh.getRange(2, 2, last - 1, 1).getDisplayValues();
    for (var i = 0; i < en.length; i++) {
      if (String(en[i][0]).trim() !== enroll) continue;
      var row = i + 2;
      sh.getRange(row, 6, 1, 3).setValues([[approved, op.teacher, op.department]]);
      sh.getRange(row, 9, 1, 3).setNumberFormat('@').setValues([[op.username, typeLabel, dates]]);
      LV_MEMO_.lastRow = row;
      return;
    }
  }
  var r = last + 1;
  sh.getRange(r, 9, 1, 3).setNumberFormat('@');
  sh.getRange(r, 1, 1, LV_COLS).setValues([[Math.max(last, 1), enroll, st.name, st.father, st.classSec,
    approved, op.teacher, op.department, op.username, typeLabel, dates, '']]);
  LV_MEMO_.lastRow = r;
}

/** Removes the student's row(s) of that month and renumbers S.No. Already removed = fine (resends are harmless). */
function lvDelete_(enroll, month) {
  enroll = String(enroll || '').trim();
  if (!enroll) throw new Error('Enrollment number is required.');
  if (LV_MONTHS.indexOf(month) === -1) throw new Error('Please select a month.');
  var sh = lvStore_(month, false);
  if (!sh || sh.getLastRow() < 2) return;
  var vals = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getDisplayValues(), removed = 0;
  for (var i = vals.length - 1; i >= 0; i--) if (String(vals[i][0]).trim() === enroll) { sh.deleteRow(i + 2); removed++; }
  if (!removed) return;
  var nl = sh.getLastRow();
  if (nl > 1) { var nums = []; for (var n = 1; n < nl; n++) nums.push([n]); sh.getRange(2, 1, nl - 1, 1).setValues(nums); }
}

/* ================= APPLICATION PHOTOS -> Drive, links in column L ================= */
function lvProofFolder_(month) {
  var root;
  if (LV_CFG.PROOF_FOLDER_ID) root = DriveApp.getFolderById(LV_CFG.PROOF_FOLDER_ID);
  else { var it = DriveApp.getFoldersByName(LV_CFG.PROOF_FOLDER_NAME); root = it.hasNext() ? it.next() : DriveApp.createFolder(LV_CFG.PROOF_FOLDER_NAME); }
  var sub = root.getFoldersByName(month);
  return sub.hasNext() ? sub.next() : root.createFolder(month);
}
function lvAttachProofs_(month, row, enroll, name, photos, opId) {
  if (!row) return;
  var sh = lvStore_(month, false), folder = lvProofFolder_(month), cell = sh.getRange(row, 12);
  var links = String(cell.getDisplayValue() || '').split(/\r?\n/).filter(Boolean);
  photos.slice(0, 5).forEach(function (p, i) {
    var fn = enroll + ' ' + String(name || '').replace(/[\\/:*?"<>|]/g, '') + ' - ' + month + ' - ' + opId + '-' + (i + 1) + '.jpg';
    var same = folder.getFilesByName(fn), file;
    if (same.hasNext()) file = same.next();
    else {
      file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(p.data), 'image/jpeg', fn));
      try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    }
    var url = file.getUrl();
    if (links.indexOf(url) === -1) links.push(url);
  });
  cell.setNumberFormat('@').setValue(links.join('\n'));
}

/* ================= ATTENDANCE OF ONE STUDENT (shortage check) =================
   Section sheet named like her Class-Sec in each month workbook (LL_STORES): Enroll# column,
   subject names one row above the header, "Total Lectures" / "Attended Lecture" per subject. */
function lvAttendance_(enroll, classSec, months) {
  enroll = String(enroll || '').trim();
  var num = function (v) { var n = parseFloat(String(v).replace(/[^\d.]/g, '')); return isNaN(n) ? 0 : n; };
  return (months || LV_MONTHS).filter(function (m) { return LV_MONTHS.indexOf(m) > -1; }).map(function (m) {
    var res = { month: m, sheet: '', subjects: [], note: '' };
    try {
      var sh = lvFindSheet_(SpreadsheetApp.openById(LL_STORES[LV_TOK[m]]), classSec);
      if (!sh) { res.note = 'No sheet "' + classSec + '"'; return res; }
      res.sheet = sh.getName();
      var data = sh.getDataRange().getDisplayValues(), h = -1;
      for (var i = 0; i < Math.min(10, data.length) && h < 0; i++) if (data[i].some(function (c) { return /total\s*lec/i.test(c); })) h = i;
      if (h < 1) { res.note = 'Attendance columns not found'; return res; }
      var ec = -1;
      for (var j = 0; j <= h && ec < 0; j++) ec = data[j].findIndex(function (c) { return /enrol|roll/i.test(c); });
      if (ec < 0) ec = 1;
      var row = data.slice(h + 1).find(function (r) { return String(r[ec]).trim() === enroll; });
      if (!row) { res.note = 'Student not in ' + res.sheet; return res; }
      data[h].forEach(function (c, k) {
        if (!/total\s*lec/i.test(c)) return;
        var subj = '';
        for (var q = k; q >= Math.max(0, k - 2) && !subj; q--) subj = String(data[h - 1][q] || '').trim();
        var ac = -1;
        for (var q2 = k + 1; q2 <= k + 2 && ac < 0; q2++) if (/attend/i.test(data[h][q2] || '')) ac = q2;
        var t = num(row[k]), a = ac > -1 ? num(row[ac]) : 0;
        if (t > 0) res.subjects.push({ s: subj || ('Subject ' + (res.subjects.length + 1)), t: t, a: a });
      });
      if (!res.subjects.length) res.note = 'No lectures recorded yet';
    } catch (err) { res.note = 'Could not read: ' + (err && err.message || err); }
    return res;
  });
}
