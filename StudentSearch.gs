/**
 * STUDENT SEARCH  (needs MonthlyBooks.gs, GrandTest.gs, Detention.gs and WorkbookReports.gs in the same project)
 *
 * The app's "Student search" screen: find any student and see her whole record:
 *   profile (strength sheet), attendance of every month (monthly attendance workbooks + approved leaves),
 *   day-by-day attendance as a calendar (daily attendance workbooks), monthly tests, grand test, detention.
 * Admin, principal and Chief Proctor: every student, every personal detail and the photo.
 * Teachers (and HOD): students of the sections they teach; personal details limited to Father name, Enroll#,
 *   Program/Class/Section, DOB, Appeared in, Quota and Subjects (nothing else is sent to their phones).
 *
 * Install (attendance Apps Script project):
 *   a) + > Script > name it "StudentSearch" > paste this file.
 *   b) doPost, under the wr_ line:
 *        if (/^st_/.test(String(req.action || ''))) return json_(stRoute_(user, req));
 *   c) Deploy > Manage deployments > Edit > New version > Deploy.
 */
var ST_CFG = {
  STRENGTH_SHEET: 'Transform-ALL-COLLEGE-STRENGTH',
  HIDE: /^(S\.?NO|PIC ?URL|PHOTO|CHALLAN.*|PORTAL ?ID|SUB\d)$/i,          // never listed as a detail (photo is shown as a picture)
  TEACHER: [['Father Name', 'Father Name'], ['Enroll#', 'Enroll #'], ['PCS', 'Program/Class/Section (PCS)'],
            ['DOB', 'DOB'], ['Appeared in', 'Appeared In'], ['Quota', 'Quota']]
};

function stRoute_(user, req) {
  try {
    switch (String(req.action || '')) {
      case 'st_find':    return stFind_(user, req.q);
      case 'st_profile': return stProfile_(user, req.roll);
      case 'st_month':   return stMonth_(user, req.roll, req.mon);
      case 'st_tests':   return stTests_(user, req.roll);
    }
    return { ok: false, error: 'UNKNOWN_ACTION' };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}

/* ================= who may see what ================= */
function stFull_(user) {
  if (gtAll_(user)) return true;
  try { return !!isCP_(user); } catch (e) { return false; }
}
/** function(classSec) -> true when this user may open students of that section */
function stSecOk_(user) {
  if (stFull_(user)) return function () { return true; };
  var list = wrSectionList_(user).map(gtK_);
  try { if (scopeOf_(user)) return function () { return true; }; } catch (e) {}      // HOD: all students, limited details
  return function (cls) { return list.indexOf(gtK_(cls)) >= 0; };
}
function stMaster_(roll) {
  var k = gtK_(roll), hit = null;
  getMaster_().rows.forEach(function (r) { if (!hit && gtK_(r[0]) === k) hit = r; });
  return hit;
}
function stCheck_(user, roll) {
  var m = stMaster_(roll);
  if (!m) throw new Error('Student ' + roll + ' is not in the strength sheet.');
  if (!stSecOk_(user)(m[2])) throw new Error('NOT_ALLOWED: this student is not in your sections.');
  return m;
}

/* ================= search ================= */
function stFind_(user, q) {
  q = gtS_(q).toLowerCase();
  if (q.length < 2) return { ok: true, list: [] };
  var ok = stSecOk_(user), out = [];
  getMaster_().rows.forEach(function (r) {
    if (out.length >= 50 || !ok(r[2])) return;
    if (String(r[0]).toLowerCase().indexOf(q) < 0 && String(r[1]).toLowerCase().indexOf(q) < 0 && String(r[4] || '').toLowerCase().indexOf(q) < 0) return;
    out.push({ roll: gtS_(r[0]), name: gtS_(r[1]), father: gtS_(r[4]), cls: gtS_(r[2]) });
  });
  return { ok: true, list: out, full: stFull_(user) };
}

/* ================= profile ================= */
function stProfile_(user, roll) {
  var m = stCheck_(user, roll), full = stFull_(user);
  var sh = SpreadsheetApp.openById(USERS_SS_ID).getSheetByName(ST_CFG.STRENGTH_SHEET);
  var v = sh.getDataRange().getDisplayValues(), head = v[0], en = head.map(gtK_).indexOf('ENROLL'), row = null;
  for (var i = 1; i < v.length && !row; i++) if (gtK_(v[i][en]) === gtK_(roll)) row = v[i];
  if (!row) row = head.map(function () { return ''; });
  var col = function (name) { var k = gtK_(name); return head.map(gtK_).indexOf(k); };
  var subjects = (m[3] || []).map(function (x) { return gtS_(x).toUpperCase(); }).filter(Boolean);
  var fields = [];
  if (full) {
    head.forEach(function (h, c) {
      var val = gtS_(row[c]);
      if (!val || !gtS_(h) || ST_CFG.HIDE.test(gtS_(h)) || /^(STUDENTNAME|ENROLL)$/.test(gtK_(h))) return;
      fields.push([gtS_(h), val]);
    });
  } else {
    ST_CFG.TEACHER.forEach(function (f) { var c = col(f[0]); fields.push([f[1], c >= 0 ? gtS_(row[c]) : (f[0] === 'PCS' ? gtS_(m[2]) : '')]); });
  }
  var out = { ok: true, full: full, roll: gtS_(m[0]), name: gtS_(m[1]), cls: gtS_(m[2]), father: gtS_(m[4]), subjects: subjects, fields: fields };
  if (full) out.photo = stPhoto_(row[col('Pic URL')] || row[col('PHOTO')]);
  return out;
}
/** a Drive link -> a small picture the app can show (data URL); anything else is returned as it is */
function stPhoto_(url) {
  url = gtS_(url); if (!url) return '';
  var id = (url.match(/[-\w]{25,}/) || [])[0];
  if (!id || !/drive|docs\.google|googleusercontent/i.test(url)) return url;
  try {
    var blob = DriveApp.getFileById(id).getBlob();
    if (blob.getBytes().length > 600000) return '';
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (e) { return ''; }
}

/* ================= one month: monthly attendance + approved leaves + day-by-day ================= */
/** her own section sheet and the optional-subject sheet of a workbook, with her row in each: [{ T, row }] */
function stRows_(book, cls, roll) {
  var out = [], k = gtK_(roll), seen = {};
  [gtFindSheet_(book, cls), wrOptSheet_(book, cls)].forEach(function (sh) {
    if (!sh || seen[sh.getName()]) return; seen[sh.getName()] = 1;
    var T = wrRead_(sh); if (!T) return;
    var row = null; T.rows.forEach(function (r) { if (!row && gtK_(r[T.en]) === k) row = r; });
    if (!row) return;
    wrGroupRows_(T, [], {});
    out.push({ T: T, row: row });
  });
  return out;
}
function stMonth_(user, roll, mon) {
  var m = stCheck_(user, roll), cls = gtS_(m[2]);
  mon = String(mon || '').toUpperCase();
  var my = monthYear_(mon), res = { ok: true, mon: mon, m: my.m, y: my.y, subs: [], leaves: '', daily: [] };
  // monthly attendance workbook
  if (MA_BOOKS[mon]) {
    try {
      var book = SpreadsheetApp.openById(MA_BOOKS[mon]);
      stRows_(book, cls, roll).forEach(function (x) {
        var subs = {};
        for (var c = x.T.idc; c < x.T.lc; c++) {
          var g = x.T['g' + c] || {}, h = gtK_(x.T.head[c]);
          if (!g.label || typeof g.id !== 'number' || /total|overall/i.test(g.label)) continue;
          var S = subs[g.label] = subs[g.label] || {};
          if (h === 'TL' || h === 'TOTALLECTURES') S.tl = c; else if (h === 'AL' || /^ATTENDEDLECTURES?$/.test(h)) S.al = c;
        }
        Object.keys(subs).forEach(function (lab) {
          var S = subs[lab]; if (S.tl == null) return;
          var d = Number(gtS_(x.row[S.tl])) || 0; if (d <= 0) return;
          res.subs.push({ sub: lab.toUpperCase(), p: Math.min(d, Number(gtS_(x.row[S.al])) || 0), d: d });
        });
      });
      var lsh = book.getSheets().filter(function (s) { return s.getName().replace(/[\s_\- ​-‍﻿]/g, '').toLowerCase() === WR_CFG.LEAVES_SHEET; })[0];
      if (lsh) {
        var lv = lsh.getDataRange().getDisplayValues(), e = -1, a = -1;
        for (var r = 0; r < Math.min(lv.length, 10) && (e < 0 || a < 0); r++) {
          e = -1; a = -1; lv[r].forEach(function (x2, c2) { var kk = gtK_(x2); if (kk === 'ENROLL') e = c2; if (kk === 'APPROVEDLEAVES') a = c2; });
          if (e >= 0 && a >= 0) { for (var i = r + 1; i < lv.length; i++) if (gtK_(lv[i][e]) === gtK_(roll)) { res.leaves = gtS_(lv[i][a]); break; } }
        }
      }
    } catch (err) { res.maError = String(err.message || err); }
  }
  // daily attendance workbook: sheets "<SECTION> <MON> <SUBJECT>"
  if (typeof BOOK_IDS !== 'undefined' && BOOK_IDS[mon]) {
    try { res.daily = stDaily_(SpreadsheetApp.openById(BOOK_IDS[mon]), cls, roll, mon); } catch (err) { res.dailyError = String(err.message || err); }
  }
  return res;
}
function stDaily_(book, cls, roll, mon) {
  var N = function (s) { return String(s || '').trim().toUpperCase().replace(/\s+/g, ' '); };
  var me = N(cls), base = N(baseClass_(cls)), yr = (me.match(/^(.*?-\d)/) || [])[1] || '', marker = ' ' + mon + ' ', target = gtK_(roll), out = [];
  book.getSheets().forEach(function (sh) {
    var name = N(sh.getName()), at = name.indexOf(marker);
    if (at < 0) return;
    var sec = name.slice(0, at).trim(), subject = name.slice(at + marker.length).trim();
    var mine = sec === me || sec === base || (yr && sec.indexOf(yr) === 0 && /OPTIONAL/.test(sec)) || (yr && sec === yr);
    if (!mine || !subject) return;
    var lr = sh.getLastRow(), lc = sh.getLastColumn(); if (lr < 4 || lc < 5) return;
    var v = sh.getRange(1, 1, lr, lc).getDisplayValues(), hdr = -1, rc = -1;
    for (var r = 0; r < Math.min(10, v.length) && hdr < 0; r++) for (var c = 0; c < v[r].length; c++) if (/^ROLLNO|^ENROLL/.test(gtK_(v[r][c]))) { hdr = r; rc = c; break; }
    if (hdr < 1) return;
    var row = -1; for (var r2 = hdr + 1; r2 < v.length; r2++) if (gtK_(v[r2][rc]) === target) { row = r2; break; }
    if (row < 0) return;
    var head = v[hdr], ta = head.map(gtK_).indexOf('TA'), days = [];
    head.forEach(function (h, c3) {
      if (!/^\d{1,2}$/.test(gtS_(h))) return;
      var held = false;
      for (var r3 = hdr + 1; r3 < v.length && !held; r3++) if (gtS_(v[r3][rc]) && gtS_(v[r3][c3])) held = true;
      if (held) days.push([Number(gtS_(h)), gtS_(v[row][c3]).toUpperCase()]);
    });
    if (!days.length) return;
    out.push({ sub: subject, teacher: gtS_(v[1][0]).replace(/^teacher\s*:?\s*/i, ''), days: days,
               ta: ta >= 0 ? gtS_(v[row][ta]) : '', tl: ta >= 0 ? gtS_(v[hdr - 1][ta]) : '' });
  });
  out.sort(function (a, b) { return a.sub.localeCompare(b.sub); });
  return out;
}

/* ================= monthly tests, grand test, detention ================= */
function stMarks_(book, cls, roll, kind) {
  var out = {};
  stRows_(book, cls, roll).forEach(function (x) {
    var subs = {};
    for (var c = x.T.idc; c < x.T.lc; c++) {
      var g = x.T['g' + c] || {}, h = gtK_(x.T.head[c]);
      if (kind === 'dt' && WR_SUMMARY.test(h)) { out['#' + gtS_(x.T.head[c])] = gtS_(x.row[c]); continue; }   // TOTAL (600), PERCENTAGE, POSITION
      if (!g.label || typeof g.id !== 'number' || /total/i.test(g.label)) continue;
      var S = subs[g.label] = subs[g.label] || {};
      if (kind === 'dt') { if (h === 'THEORY') S.th = c; else if (h === 'PRACTICAL') S.pr = c; else if (h === 'TOTAL') S.tot = c; }
      else if (h === 'OBTAINEDMARKS' || h === 'OM' || h === 'OBTAINEDGTMARKS') S.obt = c;
      else if (h === 'TOTALMARKS' || h === 'TM' || h === 'TOTALGTMARKS') S.max = c;
    }
    Object.keys(subs).forEach(function (lab) {
      var S = subs[lab], key = lab.toUpperCase();
      if (kind === 'dt') {
        if (S.tot == null) return;
        var L = dtLimits_(lab), tot = gtS_(x.row[S.tot]);
        if (tot === '' && (S.th == null || gtS_(x.row[S.th]) === '')) return;
        out[key] = { th: S.th != null ? gtS_(x.row[S.th]) : '', pr: S.pr != null ? gtS_(x.row[S.pr]) : '', tot: tot, max: L.st, thMax: L.th, prMax: L.pr };
      } else {
        if (S.obt == null) return;
        var v = gtS_(x.row[S.obt]); if (v === '') return;
        var mx = S.max != null ? Number(gtS_(x.row[S.max])) || 0 : 0;
        if (!mx && kind === 'gt') { try { mx = gtApp_().total; } catch (e) {} }
        if (!mx && S.max != null) for (var i = 0; i < x.T.rows.length && !mx; i++) mx = Number(gtS_(x.T.rows[i][S.max])) || 0;
        out[key] = [v, mx];
      }
    });
  });
  return out;
}
function stTests_(user, roll) {
  var m = stCheck_(user, roll), cls = gtS_(m[2]), res = { ok: true, mt: {}, gt: {}, dt: {} };
  SEASON.forEach(function (t) {
    try { var b = mtBook_(t); if (b) { var x = stMarks_(b, cls, roll, 'mt'); if (Object.keys(x).length) res.mt[t] = x; } } catch (e) {}
  });
  try { res.gt = stMarks_(gtBook_(), cls, roll, 'gt'); } catch (e) { res.gtError = String(e.message || e); }
  try { res.dt = stMarks_(dtBook_(), cls, roll, 'dt'); } catch (e) { res.dtError = String(e.message || e); }
  return res;
}
