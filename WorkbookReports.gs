/**
 * OFFICIAL SHEET PDFs  (needs MonthlyBooks.gs, GrandTest.gs and Detention.gs in the same project)
 *
 * The section-wise PDFs of the old portals, now from the app, for every month:
 *   - Attendance shortage   (monthly attendance workbook of the month, Leaves from "ApprovedLeavesStore", overall %)
 *   - Monthly test list     (Monthly Test workbook of the month, highest per subject, top 3, absent in pink)
 *   - Grand test award list (Grand Test workbook, teacher names from the timetable)
 *   - Detention award list  (Detention workbook, theory / practical / total, fail in red)
 * The app shows the PDF with Share (WhatsApp, Gmail...), Download and Open / Print.
 *
 * F.A sections: "F.A-1 A" is joined with "F.A-1 Optional Subjects" every time (only the optional subjects
 * the girls of that section take), so the PDF is always up to date. Students struck through (left) are left out.
 *
 * Install (attendance Apps Script project):
 *   a) + > Script > name it "WorkbookReports" > paste this file.
 *   b) doPost, under the dt_ line:
 *        if (/^wr_/.test(String(req.action || ''))) return json_(wrRoute_(user, req));
 *   c) Deploy > Manage deployments > Edit > New version > Deploy.
 */
var WR_CFG = {
  THRESHOLD: 0,                 // attendance shortage limit in %; 0 = the app's limit (73)
  LEAVES_SHEET: 'approvedleavesstore',
  PRINT_WIDTH: 1060             // A4 landscape, px
};

function wrRoute_(user, req) {
  try {
    switch (String(req.action || '')) {
      case 'wr_sections': return wrSections_(user);
      case 'wr_pdf':      return wrPdf_(user, req);
    }
    return { ok: false, error: 'UNKNOWN_ACTION' };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}

/* ================= who may make which section ================= */
function wrAll_(user) {
  if (gtAll_(user)) return true;
  try { return !!scopeOf_(user); } catch (e) { return false; }
}
/** section sheets this user may print: all for admin / principal / HOD, else the sections she teaches
 *  (an optional-subject teacher of F.A-1 gets F.A-1 A and F.A-1 B) */
function wrSectionList_(user) {
  var every = bkClasses_().filter(function (c) { return !/optional/i.test(c); });
  if (wrAll_(user)) return every;
  var mine = {};
  getAssignments_(user).forEach(function (a) {
    var k = gtK_(a.cls), base = gtK_(baseClass_(a.cls)), opt = base !== k;
    every.forEach(function (c) { var ck = gtK_(c); if (ck === k || (opt && ck.indexOf(base) === 0)) mine[c] = 1; });
  });
  return every.filter(function (c) { return mine[c]; });
}
function wrSections_(user) {
  var list = wrSectionList_(user).sort(function (a, b) { return a.localeCompare(b, undefined, { numeric: true }); });
  var att = SEASON.filter(function (t) { return !!MA_BOOKS[t]; });
  return { ok: true, sections: list, att: att, all: wrAll_(user) };
}

/* ================= reading a section sheet (also F.A combined) ================= */
/** { teach, subj, head, rows:[[...]], idc } ; header row = the row with Enroll# */
function wrRead_(sh) {
  var lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2 || lc < 3) return null;
  var v = sh.getRange(1, 1, lr, lc).getDisplayValues(), hi = -1;
  for (var i = 0; i < Math.min(v.length, 10) && hi < 0; i++) {
    for (var c = 0; c < v[i].length; c++) if (gtK_(v[i][c]) === 'ENROLL') { hi = i; break; }
  }
  if (hi < 0) return null;
  var head = v[hi].map(gtK_), find = function (n) { for (var k = 0; k < n.length; k++) { var x = head.indexOf(n[k]); if (x >= 0) return x; } return -1; };
  var en = head.indexOf('ENROLL'), nm = find(['STUDENTNAME', 'NAME']), fa = find(['FATHERNAME', 'FATHER']), cs = find(['CLASSSEC', 'CLASSSECTION']);
  var idc = Math.max(en, nm, fa, cs, 0) + 1;
  var struck = {};
  try {
    if (lr > hi + 1) {
      var fl = sh.getRange(hi + 2, (nm >= 0 ? nm : en) + 1, lr - hi - 1, 1).getFontLines();
      for (var j = 0; j < fl.length; j++) if (fl[j][0] === 'line-through') struck[j] = true;
    }
  } catch (e) {}
  var rows = [];
  for (var r = hi + 1; r < v.length; r++) if (gtS_(v[r][en]) && !struck[r - hi - 1]) rows.push(v[r]);
  return { teach: hi >= 2 ? v[hi - 2] : v[hi].map(function () { return ''; }), subj: hi >= 1 ? v[hi - 1] : v[hi].map(function () { return ''; }),
           head: v[hi], rows: rows, idc: idc, en: en, lc: lc };
}
/** "F.A-1 A" -> the "F.A-1 Optional Subjects" sheet of the same workbook, or null */
function wrOptSheet_(book, cls) {
  var parts = gtS_(cls).split(/\s+/); if (parts.length < 2) return null;
  var pre = gtK_(parts.slice(0, -1).join(' '));
  return book.getSheets().filter(function (s) { var k = gtK_(s.getName()); return /optional/i.test(s.getName()) && k.indexOf(pre) === 0; })[0] || null;
}
/** the table of one section: its own sheet, joined with the optional-subject sheet when there is one */
function wrTable_(book, cls) {
  var sh = gtFindSheet_(book, cls), opt = wrOptSheet_(book, cls);
  if (!sh) sh = gtFindSheet_(book, cls + ' combine');
  if (!sh) throw new Error('"' + book.getName() + '" has no sheet called "' + cls + '".');
  var B = wrRead_(sh);
  if (!B) throw new Error('Sheet "' + sh.getName() + '" has no Enroll# header row.');
  B.name = sh.getName().replace(/\s*combine\s*$/i, '').trim();
  if (!opt || /combine/i.test(sh.getName())) return B;
  var O = wrRead_(opt); if (!O) return B;
  var byRoll = {}; O.rows.forEach(function (r) { byRoll[gtK_(r[O.en])] = r; });
  var mine = B.rows.map(function (r) { return byRoll[gtK_(r[B.en])] || null; });
  // groups of optional columns (a group starts where the subject row has a name); keep groups this section takes
  var groups = [], g = null;
  for (var c = O.idc; c < O.lc; c++) {
    if (gtS_(O.subj[c]) || !g) { g = { cols: [], used: false }; groups.push(g); }
    g.cols.push(c);
  }
  groups.forEach(function (gr) {
    gr.cols.forEach(function (c) {
      var h = gtK_(O.head[c]);
      if (/^(TL|TOTALLECTURES|TM|TOTALMARKS|TOTALGTMARKS|TOTAL)$/.test(h)) return;   // filled for everyone
      mine.forEach(function (r) { if (r && gtS_(r[c])) gr.used = true; });
    });
  });
  var keep = []; groups.forEach(function (gr) { if (gr.used) keep = keep.concat(gr.cols); });
  if (!keep.length) return B;
  var add = function (a, b) { return a.slice(0, B.lc).concat(keep.map(function (c) { return b[c]; })); };
  return { teach: add(B.teach, O.teach), subj: add(B.subj, O.subj), head: add(B.head, O.head), idc: B.idc, en: B.en, lc: B.lc + keep.length, name: B.name,
           rows: B.rows.map(function (r, i) { var o = mine[i]; return r.slice(0, B.lc).concat(keep.map(function (c) { return o ? o[c] : ''; })); }), optCls: opt.getName() };
}

/* ================= helpers ================= */
var WR_SUMMARY = /^(TOTAL\d+|GRANDTOTAL|OVERALLTOTAL|PERCENT|PERCENTAGE|POSITION|RANK|RESULT|REMARKS?|GRADE)/;
function wrEsc_(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
function wrAbsent_(v) { var t = gtS_(v).toUpperCase().replace(/\./g, ''); return t === 'A' || t === 'AB' || t === 'ABS' || t === 'ABSENT'; }
/** teacher of each subject of this section, from the timetable assignments: { SUBJECTKEY: 'MS. X' } */
function wrTeachers_(T, cls) {
  var keys = [gtK_(cls), gtK_(T.optCls || '')], out = {};
  getAssignments_({ teacher: '', username: ADMIN_USER }).forEach(function (a) {
    if (keys.indexOf(gtK_(a.cls)) >= 0 && a.teacher && !out[gtK_(a.subject)]) out[gtK_(a.subject)] = a.teacher;
  });
  return out;
}
function wrTeacherFor_(map, subject) {
  var k = gtK_(subject); if (!k) return '';
  if (map[k]) return map[k];
  var hit = Object.keys(map).filter(function (x) { return x.indexOf(k) === 0 || k.indexOf(x) === 0; });
  return hit.length === 1 ? map[hit[0]] : '';
}
function wrPdfOut_(html, name) {
  var blob = HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf').setName(name);
  return { ok: true, name: name, pdf: Utilities.base64Encode(blob.getBytes()) };
}
function wrNow_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy hh:mm a'); }

/* ================= the PDF ================= */
function wrPdf_(user, req) {
  var kind = String(req.kind || ''), mon = String(req.mon || '').toUpperCase(), cls = gtS_(req.cls);
  if (!cls) return { ok: false, error: 'Choose a section.' };
  var allowed = wrSectionList_(user).some(function (c) { return gtK_(c) === gtK_(cls); });
  if (!allowed) return { ok: false, error: 'NOT_ALLOWED: ' + cls };
  var book, label;
  if (kind === 'att') {
    if (!MA_BOOKS[mon]) return { ok: false, error: 'No monthly attendance workbook for ' + mtMonthName_(mon) + '.' };
    book = SpreadsheetApp.openById(MA_BOOKS[mon]); label = mtMonthName_(mon).toUpperCase();
  } else if (kind === 'mt') {
    book = mtBook_(mon); if (!book) return { ok: false, error: 'No Monthly Test workbook for ' + mtMonthName_(mon) + ' yet.' };
    label = mtMonthName_(mon).toUpperCase();
  } else if (kind === 'gt') { book = gtBook_(); label = 'GRAND TEST'; }
  else if (kind === 'dt') { book = dtBook_(); label = 'DETENTION'; }
  else return { ok: false, error: 'Unknown report.' };
  var T = wrTable_(book, cls);
  if (!T.rows.length) return { ok: false, error: 'Sheet "' + T.name + '" has no students yet.' };
  return kind === 'att' ? wrAttPdf_(book, T, label, req) : wrMarksPdf_(T, kind, label, cls);
}

/** header cells of the teacher / subject rows, one per subject group of the kept columns */
function wrGroupRows_(T, cols, tmap) {
  var spans = [], last = null, gid = -1, label = '';
  for (var c = T.idc; c < T.lc; c++) {
    if (gtS_(T.subj[c])) { gid = c; label = gtS_(T.subj[c]); }
    else if (WR_SUMMARY.test(gtK_(T.head[c])) && label !== '') { gid = 'S' + c; label = ''; }   // overall total / % / position: no subject
    T['g' + c] = { id: gid, label: label, teach: typeof gid === 'number' && gid >= 0 ? T.teach[gid] : '' };
  }
  cols.forEach(function (c) {
    var g = T['g' + c];
    if (last && last.id === g.id) last.n++;
    else { last = { id: g.id, label: g.label, teach: g.teach, n: 1 }; spans.push(last); }
  });
  var t = '', s = '';
  spans.forEach(function (sp) {
    var who = /total/i.test(sp.label) ? '' : (wrTeacherFor_(tmap, sp.label) || gtS_(sp.teach));
    t += '<td colspan="' + sp.n + '" class="tch">' + wrEsc_(who) + '</td>';
    s += '<td colspan="' + sp.n + '" class="sub">' + wrEsc_(sp.label) + '</td>';
  });
  return { t: t, s: s };
}
function wrCss_(ncols) {
  var fs = ncols > 45 ? 6.5 : ncols > 30 ? 7 : 8.5;
  return '<style>@page{size:A4 landscape;margin:5mm}body{margin:0;font-family:Arial,Helvetica,sans-serif;width:' + WR_CFG.PRINT_WIDTH + 'px}' +
    'table{border-collapse:collapse;table-layout:fixed;width:' + WR_CFG.PRINT_WIDTH + 'px}tr{page-break-inside:avoid}' +
    'td{border:1px solid #000;padding:2px 1px;text-align:center;vertical-align:middle;font-size:' + fs + 'px;line-height:1.1;overflow:hidden;word-break:break-word;-webkit-print-color-adjust:exact;print-color-adjust:exact}' +
    'td.l{text-align:left;white-space:nowrap;text-overflow:ellipsis;padding-left:3px}td.gen{border:none;text-align:right;font-size:8px}' +
    'td.ttl{font-size:17px;font-weight:bold;padding:5px;line-height:1.3}td.tch{font-size:' + (fs - 0.5) + 'px}td.sub{font-weight:bold;font-size:' + (fs + 0.5) + 'px}' +
    'td.hd{font-weight:bold;background:#eee}tr.z td{background:#f7f7f7}.hi{background:#ccc!important;font-weight:bold}.ab{background:#ffc0cb!important}' +
    '.low{background:#ff0!important;color:#f00;font-weight:bold}.g1{background:#ffd700!important;font-weight:bold}.g2{background:#c0c0c0!important;font-weight:bold}.g3{background:#cd7f32!important;font-weight:bold}' +
    '.note{margin-top:10px;font-size:9px;line-height:1.5}.sum{margin-top:8px;font-size:10px;font-weight:bold}</style>';
}
function wrColgroup_(n, idc, extra) {
  var idw = [18, 62, 105, 95, 52].slice(0, idc), used = idw.reduce(function (a, b) { return a + b; }, 0);
  var rest = Math.max(12, (WR_CFG.PRINT_WIDTH - used) / Math.max(1, n - idc));
  var h = '<colgroup>'; idw.forEach(function (w) { h += '<col style="width:' + w + 'px">'; });
  for (var i = idc; i < n; i++) h += '<col style="width:' + (extra && extra[i - idc] ? extra[i - idc] : rest).toFixed(1) + 'px">';
  return h + '</colgroup>';
}

/* ---------- attendance shortage ---------- */
function wrAttPdf_(book, T, label, req) {
  var limit = WR_CFG.THRESHOLD > 0 ? WR_CFG.THRESHOLD : (Number(req.below) > 0 ? Number(req.below) : 73);
  var type = function (h) { h = gtK_(h); return (h === 'TL' || h === 'TOTALLECTURES') ? 'TL' : (h === 'AL' || h === 'ATTENDEDLECTURE' || h === 'ATTENDEDLECTURES') ? 'AL' : (gtS_(h) === '' ? '' : /^(PERCENT|PERCENTAGE|ATTENDANCEPERCENT|ATTENDANCE)$/.test(h) ? 'PCT' : 'OTHER'); };
  var isPct = function (raw) { var s = gtS_(raw).replace(/\s+/g, '').toUpperCase(); return s === '%' || s === 'PERCENT' || s === 'PERCENTAGE' || s === 'ATTENDANCE%' || s === 'ATTENDANCEPERCENT'; };
  // approved leaves
  var leaves = {}, lsh = book.getSheets().filter(function (s) { return s.getName().replace(/[\s_\- ​-‍﻿]/g, '').toLowerCase() === WR_CFG.LEAVES_SHEET; })[0];
  if (lsh) {
    var lv = lsh.getDataRange().getDisplayValues(), e = -1, a = -1, hr = -1;
    for (var r = 0; r < Math.min(lv.length, 10) && hr < 0; r++) {
      e = -1; a = -1;
      lv[r].forEach(function (x, c) { var k = gtK_(x); if (k === 'ENROLL') e = c; if (k === 'APPROVEDLEAVES') a = c; });
      if (e >= 0 && a >= 0) hr = r;
    }
    if (hr >= 0) for (var i = hr + 1; i < lv.length; i++) { var k = gtK_(lv[i][e]); if (k) leaves[k] = lv[i][a]; }
  }
  var cols = []; for (var c = 0; c < T.lc; c++) cols.push(c);
  var sc = cols.slice(T.idc), widths = sc.map(function (c) { return isPct(T.head[c]) ? 26 : 20; });
  var used = [18, 62, 105, 95, 52].slice(0, T.idc).reduce(function (a2, b) { return a2 + b; }, 0);
  var room = WR_CFG.PRINT_WIDTH - used - 24 - 30, sum = widths.reduce(function (a2, b) { return a2 + b; }, 0) || 1;
  widths = widths.map(function (w) { return Math.max(11, w * room / sum); }).concat([24, 30]);
  var n = T.lc + 2, G = wrGroupRows_(T, sc, wrTeachers_(T, T.name)), short = 0;
  var h = '<html><head><meta charset="UTF-8">' + wrCss_(n) + '</head><body><table>' + wrColgroup_(n, T.idc, widths) +
    '<thead><tr><td class="gen" colspan="' + n + '"><b>Generated on:</b> ' + wrNow_() + '</td></tr>' +
    '<tr><td class="ttl" colspan="' + n + '">' + label + ' - JCW ATTENDANCE SHORTAGE - SESSION ' + gtSession_() + '<br>SECTION: ' + wrEsc_(T.name) + '</td></tr>' +
    '<tr><td colspan="' + T.idc + '" rowspan="2"></td>' + G.t + '<td rowspan="2" colspan="2"></td></tr><tr>' + G.s + '</tr><tr>';
  for (c = 0; c < T.lc; c++) {
    var hd = gtS_(T.head[c]), hk = gtK_(hd);
    if (hk === 'TOTALLECTURES') hd = 'T.L'; if (hk === 'ATTENDEDLECTURE' || hk === 'ATTENDEDLECTURES') hd = 'A.L';
    h += '<td class="hd">' + wrEsc_(hd) + '</td>';
  }
  h += '<td class="hd">Leaves</td><td class="hd">%AGE</td></tr></thead><tbody>';
  T.rows.forEach(function (row, ri) {
    h += '<tr' + (ri % 2 ? ' class="z"' : '') + '>';
    var tl = 0, al = 0;
    for (var c2 = 0; c2 < T.lc; c2++) {
      var val = gtS_(row[c2]), cl = c2 === 2 || c2 === 3 ? 'l' : '';
      if (c2 >= T.idc) {
        var ty = type(T.head[c2]), num = parseFloat(val);
        if (!isNaN(num)) { if (ty === 'TL') tl += num; if (ty === 'AL') al += num; }
        if (isPct(T.head[c2])) {
          if (gtS_(row[c2 - 1]) === '') val = '';
          else { var p = parseFloat(val.replace('%', '')); if (!isNaN(p)) { val = Math.round(p) + '%'; if (p < limit) cl = 'low'; } }
        }
      }
      h += '<td' + (cl ? ' class="' + cl + '"' : '') + '>' + wrEsc_(val) + '</td>';
    }
    var lvv = leaves[gtK_(row[T.en])], lnum = parseFloat(lvv), ov = '', ocl = '';
    if (!isNaN(lnum)) al += lnum;
    if (tl > 0) { var op = Math.min(100, al / tl * 100); ov = Math.round(op) + '%'; if (op < limit) { ocl = ' class="low"'; short++; } }
    h += '<td>' + wrEsc_(lvv == null ? '' : lvv) + '</td><td' + ocl + '>' + ov + '</td></tr>';
  });
  h += '</tbody></table><div class="sum">Students: ' + T.rows.length + ' | Below ' + limit + '% overall: ' + short + '</div>' +
    '<div class="note"><b>Note:</b><br>1. <b>Attendance Threshold:</b> ' + limit + '% (2 casual leaves already adjusted in threshold).<br>' +
    '2. In case of any error in marked attendance contact to subject teacher.<br>3. Leave-related queries should be addressed to the Class Teacher.' +
    (lsh ? '' : '<br><i>Approved leaves not added: this workbook has no ApprovedLeavesStore sheet.</i>') + '</div></body></html>';
  return wrPdfOut_(h, T.name + '_' + label + '_Attendance_Shortage_' + gtSession_() + '.pdf');
}

/* ---------- monthly test / grand test / detention ---------- */
function wrMarksPdf_(T, kind, label, cls) {
  var dt = kind === 'dt';
  var head = T.head.map(function (x) {
    var t = gtS_(x).toLowerCase();
    if (t.indexOf('total gt marks') >= 0) return kind === 'gt' ? 'TGM' : 'TMM';
    if (t.indexOf('obtained gt marks') >= 0) return kind === 'gt' ? 'OGM' : 'OMM';
    if (t.indexOf('total marks') >= 0) return 'T.M';
    if (t.indexOf('obtained marks') >= 0) return 'O.M';
    return gtS_(x);
  });
  // columns: drop the "total" column of the subject row (the old award lists did the same); detention keeps all
  var cols = [];
  for (var c = 0; c < T.lc; c++) if (dt || c < T.idc || !/total/i.test(gtS_(T.subj[c]))) cols.push(c);
  var G = wrGroupRows_(T, cols.filter(function (x) { return x >= T.idc; }), wrTeachers_(T, cls));
  // highest per subject, ranking of the grand total
  var obt = [], rankCol = -1;
  cols.forEach(function (c2) {
    var hd = head[c2], k = gtK_(hd);
    if (hd === 'O.M' || (dt && k === 'TOTAL' && gtS_(T['g' + c2] && T['g' + c2].label) && !/total/i.test(T['g' + c2].label))) obt.push(c2);
    if (hd === 'OGM' || hd === 'OMM' || /^(TOTAL\d+|GRANDTOTAL|OVERALLTOTAL)$/.test(k) || (dt && k === 'TOTAL' && /total/i.test((T['g' + c2] || {}).label || ''))) rankCol = c2;
  });
  var hi = {};
  obt.forEach(function (c2) { var m = 0; T.rows.forEach(function (r) { var x = parseFloat(r[c2]); if (!isNaN(x) && x > m) m = x; }); hi[c2] = m; });
  var medal = {};
  if (rankCol >= 0) {
    var vals = T.rows.map(function (r, i) { return { i: i, v: parseFloat(r[rankCol]) || 0 }; }).sort(function (a, b) { return b.v - a.v; });
    var rank = 0, lastV = null;
    vals.forEach(function (o) { if (o.v !== lastV) rank++; lastV = o.v; if (rank <= 3 && o.v > 0) medal[o.i] = 'g' + rank; });
  }
  // detention: below pass mark in red
  var failMax = {};
  if (dt && typeof dtLimits_ === 'function') cols.forEach(function (c2) { var g = T['g' + c2]; if (g && gtK_(head[c2]) === 'TOTAL' && !/total/i.test(g.label)) failMax[c2] = dtLimits_(g.label).st; });
  var n = cols.length, pres = 0, abs = 0;
  var h = '<html><head><meta charset="UTF-8">' + wrCss_(n) + '</head><body><table>' + wrColgroup_(n, T.idc) +
    '<thead><tr><td class="gen" colspan="' + n + '"><b>Generated on:</b> ' + wrNow_() + '</td></tr>' +
    '<tr><td class="ttl" colspan="' + n + '">' + (kind === 'mt' ? label + ' MONTHLY TEST' : kind === 'gt' ? 'JCW GRAND TEST' : 'JCW DETENTION EXAM') +
    ' - SESSION ' + gtSession_() + ' &nbsp; SECTION: ' + wrEsc_(T.name) + (kind === 'mt' ? '' : '<br>AWARD LIST') + '</td></tr>' +
    '<tr><td colspan="' + T.idc + '" rowspan="2"></td>' + G.t + '</tr><tr>' + G.s + '</tr><tr>';
  cols.forEach(function (c2) { h += '<td class="hd">' + wrEsc_(head[c2]) + '</td>'; });
  h += '</tr></thead><tbody>';
  T.rows.forEach(function (row, ri) {
    var any = false, allAbs = true;
    h += '<tr' + (ri % 2 ? ' class="z"' : '') + '>';
    cols.forEach(function (c2) {
      var val = gtS_(row[c2]), cl = c2 === 2 || c2 === 3 ? 'l' : '';
      if (c2 === 2 || c2 === 3) val = val.toUpperCase();
      if (c2 >= T.idc) {
        var x = parseFloat(val);
        if (hi[c2] && x === hi[c2]) cl = 'hi';
        if (failMax[c2] && !isNaN(x) && x < failMax[c2] * DT_CFG.PASS) cl = 'low';
        if (c2 === rankCol && medal[ri]) cl = medal[ri];
        if (wrAbsent_(val)) cl = 'ab';
        if (val) { any = true; if (!wrAbsent_(val) && obt.indexOf(c2) >= 0) allAbs = false; }
      }
      h += '<td' + (cl ? ' class="' + cl + '"' : '') + '>' + wrEsc_(val) + '</td>';
    });
    h += '</tr>';
    if (any) { if (allAbs && obt.length) abs++; else pres++; }
  });
  h += '</tbody></table><div class="sum">Students: ' + T.rows.length + (obt.length ? ' | Appeared: ' + pres + ' | Absent in all: ' + abs : '') + '</div>' +
    '<div class="note">Grey = highest in the subject' + (rankCol >= 0 ? ' · Gold / silver / bronze = 1st, 2nd, 3rd overall' : '') +
    ' · Pink = absent' + (dt ? ' · Red = below ' + Math.round(GT_CFG.PASS * 100) + '%' : '') + '</div></body></html>';
  var file = kind === 'mt' ? T.name + '_' + label + '_MonthlyTest_' + gtSession_()
    : kind === 'gt' ? T.name + '_GrandTest_AwardList_' + gtSession_() : T.name + '_Detention_AwardList_' + gtSession_();
  return wrPdfOut_(h, file.replace(/\s+/g, ' ') + '.pdf');
}
