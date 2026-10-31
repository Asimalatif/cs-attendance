/**
 * STUDENT SEARCH  (needs MonthlyBooks.gs, GrandTest.gs, Detention.gs and WorkbookReports.gs in the same project)
 *
 * The app's "Student search" screen: find any student and see her whole record:
 *   profile (strength sheet), attendance of every month (monthly attendance workbooks + approved leaves),
 *   day-by-day attendance as a calendar (daily attendance workbooks), monthly tests, grand test, detention.
 * Admin, principal and Chief Proctor: every student, every personal detail and the photo.
 * Teachers (and HOD): any student of the college; everything (photo, board result too) except the
 *   student's contact number, email and address (these are never sent to their phones).
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
  PRIVATE: /CONTACT|MOBILE|PHONE|WHATSAPP|EMAIL|ADDRESS/i,          // shown to admin, principal and Chief Proctor only
  // 1st year board result of 2nd year students (Provisional Certificates workbook)
  BOARD_BOOK: '1C2tdr2yXAja0_C5Nk4HIk2m-71Kc4I-FS5VQUhjUtHQ',
  BOARD_BY_ENROLL: 'boardResult1STYEAR',     // has Enroll#, Roll No, Registration, Marks
  BOARD_LIST: '1styeardata2026',             // the board's list: Roll No, Name, Father Name, Registration #, Marks, Grade / Remarks
                                            // (a sheet made by board.html, e.g. 'HSSC 1st year 2026', also works: only Jinnah College rows are used)
  BOARD_COLLEGE: /JINNAH COLLEGE FOR WOMEN/i,
  BOARD_LINKS: 'App Board Links',            // choices confirmed in the app (made by itself)
  MATCH_MIN: 0.86, MATCH_GAP: 0.03,
  BOARD_TOTAL: 600,                         // total marks of the 1st year board exam
  DISTINCTION: 20,                          // Peshawar Board positions 1-20 of the 1st year result (Position column) are distinctions
  // strength-sheet columns that are about MATRIC: shown with a clear name
  LABELS: { PERCENTAGE: 'Matric percentage', OBTAINEDMARKS: 'Matric obtained marks', TOTALMARKS: 'Matric total marks', BOARD: 'Matric board',
            PASSINGYEAR: 'Matric passing year', METRICROLLNO: 'Matric Roll No', MATRICROLLNO: 'Matric Roll No' }
};

function stRoute_(user, req) {
  try {
    switch (String(req.action || '')) {
      case 'st_find':    return stFind_(user, req.q);
      case 'st_profile': return stProfile_(user, req.roll);
      case 'st_month':   return stMonth_(user, req.roll, req.mon);
      case 'st_tests':   return stTests_(user, req.roll);
      case 'st_board':   return stBoard_(user, req.roll);
      case 'st_board_link': return stBoardLink_(user, req.roll, req.board);
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
/** every signed-in teacher may open any student of the college (contact, email and address stay hidden: see PRIVATE) */
function stSecOk_(user) {
  return function () { return true; };
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
  head.forEach(function (h, c) {
    var val = gtS_(row[c]);
    if (!val || !gtS_(h) || ST_CFG.HIDE.test(gtS_(h)) || /^(STUDENTNAME|ENROLL)$/.test(gtK_(h))) return;
    if (!full && ST_CFG.PRIVATE.test(gtS_(h))) return;                // contact, email, address: admin / principal / CP only
    fields.push([ST_CFG.LABELS[gtK_(h)] || gtS_(h), val]);
  });
  fields = stTidy_(fields, roll);
  var out = { ok: true, full: full, roll: gtS_(m[0]), name: gtS_(m[1]), cls: gtS_(m[2]), father: gtS_(m[4]), subjects: subjects, fields: fields };
  out.photo = full ? stPhoto_(row[col('Pic URL')] || row[col('PHOTO')]) : '';   // photo: admin / principal / Chief Proctor only
  return out;
}
/** no repetitions: PCS already says program / class / section, REG# = Enroll#, NULL / 0 are empty; matric marks on one line */
function stTidy_(fields, roll) {
  var get = function (lab) { var f = fields.filter(function (x) { return x[0] === lab; })[0]; return f ? f[1] : ''; };
  var empty = function (v) { return /^(NULL|N\/?A|-|0+(\.0+)?|FALSE)$/i.test(gtS_(v)); };
  var drop = { PROGRAMCLASSSECTION: 1, SECTION: 1, GROUPONLY: 1 };
  var ob = get('Matric obtained marks'), tt = get('Matric total marks'), pc = get('Matric percentage'), out = [];
  fields.forEach(function (f) {
    var k = gtK_(f[0]);
    if (drop[k] || empty(f[1])) return;
    if ((k === 'REG' || k === 'REGNO') && gtK_(f[1]) === gtK_(roll)) return;
    if (k === 'MATRICTOTALMARKS' || k === 'MATRICPERCENTAGE') return;
    if (k === 'MATRICOBTAINEDMARKS') {
      var o = Number(ob), t = Number(tt), p = Number(pc) || (o && t ? Math.round(o / t * 1000) / 10 : 0);
      out.push(['Matric marks', String(o || ob) + (t ? ' / ' + t : '') + (p ? ' (' + p + '%)' : '')]); return;
    }
    out.push(f);
  });
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

/* ================= 1st YEAR BOARD RESULT of 2nd year students =================
   1) a link confirmed in the app ("App Board Links" sheet)  2) boardResult1STYEAR by Enroll#
   3) name + father name matched in 1styeardata2026, spelling tolerant (Muhammad / Mohd, double letters, ee / i,
      word order, an extra Khan ...). Only a clear best match is taken; otherwise admin / CP choose from the closest 5. */
function stBoardPct_(m) { var n = Number(gtS_(m)); return n > 0 ? String(Math.round(n / ST_CFG.BOARD_TOTAL * 1000) / 10) : ''; }
function stIsSecond_(cls) { var u = gtS_(cls).toUpperCase().replace(/\s+/g, ''); return /^(FCS|PRE-?ENG|F\.?A)-?2/.test(u) || /^[A-Z]2$/.test(u); }
function stBoardData_() {
  var c = wrCacheGet_('st_board_v3'); if (c && (!c.list.length || c.list[0].length >= 7)) return c;   // an old copy without Position is read again
  var book = SpreadsheetApp.openById(ST_CFG.BOARD_BOOK), out = { enr: {}, list: [], links: {} };
  var col = function (head, names) { var h = head.map(gtK_); for (var i = 0; i < names.length; i++) { var x = h.indexOf(gtK_(names[i])); if (x >= 0) return x; } return -1; };
  var sh = book.getSheetByName(ST_CFG.BOARD_LIST);
  if (sh) {
    var v = sh.getDataRange().getDisplayValues(), h = v[0];
    var ir = col(h, ['Roll No']), inm = col(h, ['Name']), ifa = col(h, ['Father Name']), ig = col(h, ['Registration #', 'Registration']), im = col(h, ['Marks']), igr = col(h, ['Grade', 'Remarks']), ip = col(h, ['Position']), ic = col(h, ['College Name']);
    for (var i = 1; i < v.length; i++) if (gtS_(v[i][ir]) && (ic < 0 || !gtS_(v[i][ic]) || ST_CFG.BOARD_COLLEGE.test(v[i][ic]))) out.list.push([gtS_(v[i][ir]), gtS_(v[i][inm]), gtS_(v[i][ifa]), ig >= 0 ? gtS_(v[i][ig]) : '', im >= 0 ? gtS_(v[i][im]) : '', igr >= 0 ? gtS_(v[i][igr]) : '', ip >= 0 ? gtS_(v[i][ip]) : '']);
  }
  sh = book.getSheetByName(ST_CFG.BOARD_BY_ENROLL);
  if (sh) {
    var w = sh.getDataRange().getDisplayValues(), hh = w[0];
    var e = col(hh, ['Enroll#']), r2 = col(hh, ['Roll No']), g2 = col(hh, ['Registration']), m2 = col(hh, ['Marks']), gr2 = col(hh, ['Grade']), pc = col(hh, ['percentage']);
    for (var j = 1; j < w.length; j++) if (gtS_(w[j][e]) && gtS_(w[j][r2])) out.enr[gtK_(w[j][e])] = [gtS_(w[j][r2]), gtS_(w[j][g2]), gtS_(w[j][m2]), gr2 >= 0 ? gtS_(w[j][gr2]) : '', pc >= 0 ? gtS_(w[j][pc]) : ''];
  }
  sh = book.getSheetByName(ST_CFG.BOARD_LINKS);
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues().forEach(function (r) { if (r[0] && r[1]) out.links[gtK_(r[0])] = gtS_(r[1]); });
  wrCachePut_('st_board_v3', out, 21600);
  return out;
}
/* ---- spelling-tolerant name comparison ---- */
function stNmKey_(s) {
  s = String(s || '').toUpperCase().replace(/[^A-Z ]/g, ' ');
  s = s.replace(/\b(MOHAMMAD|MOHAMMED|MUHAMMED|MUHAMAD|MOHAMAD|MOHD|MUHD|MHD|MUHAMMAD|MUHAMMMAD|M)\b/g, 'MUHAMMAD');
  s = s.replace(/\b(SYEDA|SAYEDA|SAYYEDA|SYED|SAYED|SAYYED|SAYYAD|SYYED)\b/g, 'SYED').replace(/\b(BIBI|BB|MISS|MS|MR|DR|HAJI|ENGR|MALIK|MIAN|SAHIBZADA|SAHIBZADI|QAZI|HAFIZ)\b/g, ' ');
  return s.split(/\s+/).filter(Boolean).map(function (w) {
    return w.replace(/(.)\1+/g, '$1').replace(/EE/g, 'I').replace(/OO/g, 'U').replace(/Y/g, 'I').replace(/^KH/, 'K').replace(/H(?=[^AEIOU]|$)/g, '').replace(/[AEIOU]+$/, 'A');
  });
}
function stJw_(a, b) {
  if (a === b) return 1; var la = a.length, lb = b.length; if (!la || !lb) return 0;
  var md = Math.max(0, Math.floor(Math.max(la, lb) / 2) - 1), am = [], bm = [], m = 0, i, j;
  for (i = 0; i < la; i++) for (j = Math.max(0, i - md); j < Math.min(lb, i + md + 1); j++) if (!bm[j] && a[i] === b[j]) { am[i] = bm[j] = true; m++; break; }
  if (!m) return 0; var t = 0, k = 0;
  for (i = 0; i < la; i++) if (am[i]) { while (!bm[k]) k++; if (a[i] !== b[k]) t++; k++; }
  var jj = (m / la + m / lb + (m - t / 2) / m) / 3, p = 0; while (p < 4 && a[p] === b[p]) p++;
  return jj + p * 0.1 * (1 - jj);
}
function stNmSim_(x, y) {
  var A = stNmKey_(x), B = stNmKey_(y); if (!A.length || !B.length) return 0;
  var side = function (P, Q) { return P.reduce(function (s, w) { return s + Math.max.apply(null, Q.map(function (v) { return stJw_(w, v); })); }, 0) / P.length; };
  var s1 = side(A, B), s2 = side(B, A); return Math.max(Math.min(s1, s2), (s1 + s2) / 2 - 0.04);
}
var ST_COMMON_ = null;
/** the own (not common) words of a name must really match: stops sisters with the same father, or two girls
 *  with the same common name (Alishba Khan, Eman) but different fathers, being mixed up */
function stOwnOk_(a, b, notIn) {
  if (!ST_COMMON_) { ST_COMMON_ = {}; stNmKey_('MUHAMMAD SYED KHAN ALI SHAH GUL BEGUM NOOR ZADA JAN ULLAH UDDIN UR REHMAN').forEach(function (w) { ST_COMMON_[w] = 1; }); }
  var fam = {}; stNmKey_(notIn || '').forEach(function (w) { fam[w] = 1; });               // a family name from the father does not count
  var A = stNmKey_(a).filter(function (w) { return !ST_COMMON_[w] && !fam[w]; }), B0 = stNmKey_(b);
  var B = B0.filter(function (w) { return !ST_COMMON_[w]; }); if (!B.length) B = B0;          // 'Alam' must not match the common 'Ali'
  if (!A.length) return true;
  return Math.max.apply(null, A.map(function (w) { return Math.max.apply(null, B.map(function (v) { return stJw_(w, v); })); })) >= 0.9;
}
function stBoard_(user, roll) {
  var m = stCheck_(user, roll);
  if (!stIsSecond_(m[2])) return { ok: true, show: false };
  var D = stBoardData_(), k = gtK_(roll), byRoll = {};
  D.list.forEach(function (x) { byRoll[x[0]] = x; });
  var out = function (rn, how, extra) {
    var x = byRoll[rn] || [rn, '', '', '', '', '', ''];
    var e = D.enr[k] && D.enr[k][0] === rn ? D.enr[k] : null;
    return Object.assign({ ok: true, show: true, found: true, how: how, rollNo: rn, reg: x[3] || (e ? e[1] : ''), marks: x[4] || (e ? e[2] : ''),
                           grade: x[5] || (e ? e[3] : ''), pct: stBoardPct_(x[4] || (e ? e[2] : '')), total: ST_CFG.BOARD_TOTAL, position: x[6] || '', distinction: Number(x[6]) > 0 && Number(x[6]) <= ST_CFG.DISTINCTION, boardName: x[1], boardFather: x[2] }, extra || {});
  };
  if (D.links[k] && D.links[k] !== '-') return out(D.links[k], 'confirmed');
  if (D.enr[k] && D.links[k] !== '-') return out(D.enr[k][0], 'enroll');
  var used = {};                                                   // board records that already belong to another student
  Object.keys(D.enr).forEach(function (e) { if (e !== k) used[D.enr[e][0]] = 1; });
  Object.keys(D.links).forEach(function (e) { if (e !== k) used[D.links[e]] = 1; });
  var sc = D.list.filter(function (x) { return !used[x[0]]; }).map(function (x) { return { x: x, s: stNmSim_(m[1], x[1]) * 0.6 + stNmSim_(m[4], x[2]) * 0.4 }; }).sort(function (a, b) { return b.s - a.s; });
  var best = sc[0], gap = best ? best.s - (sc[1] ? sc[1].s : 0) : 0;
  if (D.links[k] !== '-' && best && best.s >= ST_CFG.MATCH_MIN && gap >= ST_CFG.MATCH_GAP && stOwnOk_(m[1], best.x[1], m[4]) && stOwnOk_(m[4], best.x[2])) return out(best.x[0], 'name', { score: Math.round(best.s * 100) });
  return { ok: true, show: true, found: false, rejected: D.links[k] === '-', candidates: sc.slice(0, 5).map(function (c) { return { rollNo: c.x[0], name: c.x[1], father: c.x[2], marks: c.x[4], score: Math.round(c.s * 100) }; }) };
}
/** admin / CP confirm which board record belongs to the student (board = board Roll No, '-' = not the one found by name) */
function stBoardLink_(user, roll, board) {
  if (!stFull_(user)) return { ok: false, error: 'NOT_ALLOWED' };
  var m = stCheck_(user, roll), book = SpreadsheetApp.openById(ST_CFG.BOARD_BOOK), sh = book.getSheetByName(ST_CFG.BOARD_LINKS);
  if (!sh) { sh = book.insertSheet(ST_CFG.BOARD_LINKS); sh.getRange(1, 1, 1, 5).setValues([['Enroll#', 'Board Roll No', 'Student', 'Confirmed by', 'Date']]).setFontWeight('bold'); }
  var lr = sh.getLastRow(), row = 0;
  if (lr > 1) sh.getRange(2, 1, lr - 1, 1).getDisplayValues().forEach(function (r, i) { if (!row && gtK_(r[0]) === gtK_(roll)) row = i + 2; });
  var vals = [[gtS_(m[0]), gtS_(board), gtS_(m[1]), user.username || '', new Date()]];
  if (row) sh.getRange(row, 1, 1, 5).setValues(vals); else sh.getRange(lr + 1, 1, 1, 5).setValues(vals);
  CacheService.getScriptCache().remove('st_board_v3_n');
  return stBoard_(user, roll);
}

