/**
 * AdminTools.gs — JCW ADMIN TOOLS inside the attendance app (page admintools.html).
 *   1. Auto timetable          (Timetable  -> "AutoTimetable" sheet)          + app-style PDF / Excel
 *   2. Teacher workload        (Science Timetable + ARTS TIMETABLE + DropdownData -> 3 workload sheets
 *                               + "WORKLOAD dd-MMM-yyyy.xlsx" in Drive)       + app-style PDF / Excel
 *   3. Workload Excel export   (the 3 workload sheets -> Drive .xlsx, print-ready)
 *   4. Stopping merit          (Portal backup -> "Stopping Merit" sheet)       + app-style PDF / Excel
 *   5. Admission slips         (Transform-ALL-COLLEGE-STRENGTH, made as PDF in the app, 4 per A4 landscape page)
 *   6. Orientation lists       (all programs: Program -> PCS -> Room, or one program + chosen sections)
 *   7. Fill room from timetable into every sheet of the destination workbook
 * The original scripts are below the router, unchanged except: they open the workbook by id (no "active"
 * spreadsheet in a web app), the toasts are skipped, and they return what they made.
 *
 * Who can open it: admin / principal / HOD see and print; only the admin can run the generators.
 *
 * INSTALL:
 *   a) + > Script > name it "AdminTools" > paste this file.
 *      If the old scripts (AUTO TIMETABLE GENERATOR, generateTeacherWorkloadByStatus, exportJCWWorkloadToExcel,
 *      generateJCWStoppingMeritReport, fillRoomFromTimetableIntoAllSheets) are ALSO in this project, delete them.
 *   b) doPost, next to the ntt_ line (once):
 *        if (/^at_/.test(String(req.action || ''))) return json_(atRoute_(user, req));
 *   c) Run atTest once from the editor (asks for Drive / Sheets / external request permission).
 *   d) Deploy > Manage deployments > Edit > New version > Deploy.
 */
var AT_CFG = {
  SS_ID: '1ARU8zMO2daJUfU_SIIo4baeSOG7qBPuni5TYRIRLBNY',   // workbook with Timetable, Science Timetable, ARTS TIMETABLE,
                                                          // DropdownData, Portal backup, Transform-ALL-COLLEGE-STRENGTH
  STRENGTH: 'Transform-ALL-COLLEGE-STRENGTH',
  WORKLOAD_SHEETS: ['Permanent Workload', 'Morning Visiting Workload', 'Evening Visiting Workload'],
  FOLDER: 'JCW Admin Reports'
};
var AT_SS_MEMO = null;
function at_ss_() { return AT_SS_MEMO || (AT_SS_MEMO = SpreadsheetApp.openById(AT_CFG.SS_ID)); }
function at_toast_() {}                                   // spreadsheet toasts cannot show in a web app
var AT_AUTO_OVERRIDE = null;                              // programs / stream label chosen on the page

function atTest() {
  Logger.log(at_ss_().getName() + ' · sheets: ' + at_ss_().getSheets().map(function (s) { return s.getName(); }).join(', '));
  Logger.log(DriveApp.getRootFolder().getName());
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
}

function atRoute_(user, req) {
  try {
    var see = false, admin = false;
    try { admin = !!isAdmin_(user); } catch (e) {}
    try { see = admin || gtAll_(user) || wrAll_(user); } catch (e) { see = admin; }
    if (!see) return { ok: false, error: 'The admin tools are for the admin, principal and HODs.' };
    var a = String(req.action || '');
    var mustAdmin = { at_run: 1, at_fillroom: 1, at_savestu: 1 };
    if (mustAdmin[a] && !admin) return { ok: false, error: 'Only the admin can run this (it writes to the workbook).' };
    switch (a) {
      case 'at_meta':     return { ok: true, data: at_meta_(admin) };
      case 'at_sheet':    return { ok: true, data: { sheets: at_readMany_(req.names || [req.name]) } };
      case 'at_strength': return { ok: true, data: at_strength_() };
      case 'at_save':     return { ok: true, data: at_save_(req) };
      case 'at_fillroom': return { ok: true, data: { message: fillRoomFromTimetableIntoAllSheets() } };
      case 'at_savestu':  return { ok: true, data: at_saveStudent_(req) };
      case 'at_roomtags': return { ok: true, data: at_roomTags_() };
      case 'at_run':      return { ok: true, data: at_run_(req) };
    }
    return { ok: false, error: 'UNKNOWN_ACTION' };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}

/** current session, e.g. 2026-27 (the one the admin set on the Shortage page, else from today's date) */
function at_session_() {
  try { return gtSession_(); } catch (e) {}
  var d = new Date(), y = d.getMonth() >= 7 ? d.getFullYear() : d.getFullYear() - 1;
  return y + '-' + String(y + 1).slice(-2);
}
function at_meta_(admin) {
  var ss = at_ss_(), names = ss.getSheets().map(function (s) { return s.getName(); });
  var progs = [];
  var tt = ss.getSheetByName(AUTO_CFG.sourceSheet);
  if (tt && tt.getLastRow() > 1) {
    var h = tt.getRange(1, 1, 1, tt.getLastColumn()).getDisplayValues()[0].map(function (x) { return String(x).trim(); });
    var ci = h.indexOf('Program');
    if (ci >= 0) {
      var seen = {};
      tt.getRange(2, ci + 1, tt.getLastRow() - 1, 1).getDisplayValues().forEach(function (r) { var p = String(r[0]).trim(); if (p && !seen[p]) { seen[p] = 1; progs.push(p); } });
    }
  }
  return {
    admin: admin, url: ss.getUrl(), workbook: ss.getName(), sheets: names, programs: progs.sort(),
    auto: { programs: AUTO_CFG.programs, streamLabel: AUTO_CFG.streamLabel, target: AUTO_CFG.targetSheet },
    workloadSheets: AT_CFG.WORKLOAD_SHEETS, session: at_session_()
  };
}

/** runs one of the original generators, then sends back the sheet(s) it wrote */
function at_run_(req) {
  var tool = String(req.tool || ''), out = {};
  if (tool === 'auto') {
    AT_AUTO_OVERRIDE = {};
    if (Array.isArray(req.programs) && req.programs.length) AT_AUTO_OVERRIDE.programs = req.programs.map(String);
    if (req.streamLabel) AT_AUTO_OVERRIDE.streamLabel = String(req.streamLabel);
    AT_AUTO_OVERRIDE.session = String(req.session || at_session_());
    generateAutoTimetable();
    out.sheets = at_readMany_([AUTO_CFG.targetSheet]);
  } else if (tool === 'workload') {
    out.excel = generateTeacherWorkloadByStatus();
    out.sheets = at_readMany_(AT_CFG.WORKLOAD_SHEETS);
  } else if (tool === 'workloadxlsx') {
    out.excel = exportJCWWorkloadToExcel();
  } else if (tool === 'merit') {
    generateJCWStoppingMeritReport();
    out.sheets = at_readMany_(['Stopping Merit']);
  } else {
    throw new Error('Unknown tool: ' + tool);
  }
  return out;
}

/* ---------- read a sheet back with its look (values, merges, colours, bold, sizes, alignment, widths) ---------- */
function at_readMany_(names) {
  return (names || []).map(function (n) { return at_read_(String(n)); });
}
function at_read_(name) {
  var sh = at_ss_().getSheetByName(name);
  if (!sh) return { name: name, missing: true };
  var nr = sh.getLastRow(), nc = sh.getLastColumn();
  if (nr < 1 || nc < 1) return { name: name, values: [], n: 0 };
  var rg = sh.getRange(1, 1, nr, nc);
  var bg = rg.getBackgrounds().map(function (r) { return r.map(function (c) { c = String(c || '').toLowerCase(); return c === '#ffffff' || c === 'white' ? '' : c; }); });
  var bold = rg.getFontWeights().map(function (r) { return r.map(function (w) { return w === 'bold' ? 1 : 0; }); });
  var size = rg.getFontSizes();
  var al = rg.getHorizontalAlignments().map(function (r) { return r.map(function (a) { a = String(a || '').toLowerCase(); return /center/.test(a) ? 'c' : /right/.test(a) ? 'r' : ''; }); });
  var merges = rg.getMergedRanges().map(function (m) { return [m.getRow() - 1, m.getColumn() - 1, m.getNumRows(), m.getNumColumns()]; });
  var widths = [];
  for (var c = 1; c <= nc; c++) widths.push(sh.getColumnWidth(c));
  return { name: name, n: nc, values: rg.getDisplayValues(), bg: bg, bold: bold, size: size, align: al, merges: merges, widths: widths, frozen: sh.getFrozenRows() };
}

/* ---------- students for admission slips / orientation lists ---------- */
/* columns of Transform-ALL-COLLEGE-STRENGTH (by header, with the old fixed positions as fallback) */
function at_cols_(head) {
  var h = head.map(function (x) { return String(x).trim().toLowerCase(); });
  function col(names, fb) { for (var i = 0; i < names.length; i++) { var k = h.indexOf(names[i]); if (k >= 0) return k; } return fb; }
  var C = {
    e: col(['enroll#', 'enroll no', 'enrollmentno', 'enrollment no', 'enrollment#'], 1),
    n: col(['student name', 'studentname'], 2), f: col(['father name', 'fathername'], 3),
    p: col(['program'], 5), c: col(['class'], 6), s: col(['section'], 7), pcs: col(['pcs'], 8),
    room: col(['room'], -1), q: col(['quota'], -1), sub: []
  };
  for (var k = 1; k <= 7; k++) C.sub.push(col(['sub' + k], -1));
  return C;
}
function at_strength_() {
  var sh = at_ss_().getSheetByName(AT_CFG.STRENGTH);
  if (!sh) throw new Error("Sheet '" + AT_CFG.STRENGTH + "' not found.");
  var v = sh.getDataRange().getDisplayValues();
  if (v.length < 2) throw new Error('No student data found.');
  var C = at_cols_(v[0]), TT = at_ttRooms_();
  var g = function (r, i) { return i >= 0 && i < r.length ? String(r[i]).trim() : ''; };
  var rows = [];
  for (var r = 1; r < v.length; r++) {
    var row = v[r], o = { row: r + 1, e: g(row, C.e), n: g(row, C.n), f: g(row, C.f), p: g(row, C.p), c: g(row, C.c),
      s: g(row, C.s), pcs: g(row, C.pcs), room: g(row, C.room), q: g(row, C.q), sub: C.sub.map(function (i) { return g(row, i); }) };
    if (!o.e && !o.n) continue;
    o.tt = at_ttRoomOf_(TT, o);                             // room of this section in the Timetable sheet
    rows.push(o);
  }
  return { rows: rows, hasRoom: C.room >= 0, hasQuota: C.q >= 0, hasSubs: C.sub.some(function (i) { return i >= 0; }), ttSheet: TT.sheet, ttFound: TT.found, faCombos: at_faCombos_() };
}

/* ---------- F.A optional subject combinations: the "FA optional subjects" sheet (gid 444973916), one combination per row ---------- */
var AT_FA_COMBO_GID = 444973916;
function at_faCombos_() {
  try {
    var ss = at_ss_(), sh = ss.getSheets().filter(function (x) { return x.getSheetId() === AT_FA_COMBO_GID; })[0]
      || ss.getSheets().filter(function (x) { return /optional/i.test(x.getName()) && /\bF\.?\s*A\b/i.test(x.getName()); })[0];
    if (!sh || sh.getLastRow() < 1) return { sheet: '', rows: [] };
    return { sheet: sh.getName(), rows: sh.getDataRange().getDisplayValues().filter(function (r) { return r.some(function (c) { return String(c).trim(); }); }) };
  } catch (e) { return { sheet: '', rows: [], error: String(e.message || e) }; }
}

/* ---------- stopping merit: only OPEN QUOTA students. Every quota / category column of Portal backup must say "open",
   and the student's Quota in Transform-ALL-COLLEGE-STRENGTH (by Enroll#) must also be open when she is found there ---------- */
var AT_SM_MEMO = null;
function at_smIsOpen_(row, headers, enrollment) {
  if (!AT_SM_MEMO) {
    var qc = [];
    headers.forEach(function (h, i) { if (/quota/i.test(String(h)) || /^\s*category\s*$/i.test(String(h))) qc.push(i); });
    var tq = {};
    try {
      var sh = at_ss_().getSheetByName(AT_CFG.STRENGTH), v = sh.getDataRange().getDisplayValues(), C = at_cols_(v[0]);
      if (C.q >= 0) for (var r = 1; r < v.length; r++) { var e = String(v[r][C.e] || '').trim().toUpperCase(); if (e) tq[e] = String(v[r][C.q] || '').trim(); }
    } catch (e) {}
    AT_SM_MEMO = { qc: qc, tq: tq, seen: {} };
  }
  var vals = AT_SM_MEMO.qc.map(function (i) { return String(row[i] == null ? '' : row[i]).trim(); }).filter(String);
  var t = AT_SM_MEMO.tq[String(enrollment || '').trim().toUpperCase()];
  if (t) vals.push(t);
  var q = vals[0] || '(blank)'; AT_SM_MEMO.seen[q] = (AT_SM_MEMO.seen[q] || 0) + 1;
  /* strictly "Open Quota" / "Open Merit" (any other quota, even one with the word open in it, is left out) */
  return vals.length > 0 && vals.every(function (x) { var k = String(x).toUpperCase().replace(/[^A-Z]+/g, ' ').trim(); return k === 'OPEN QUOTA' || k === 'OPEN MERIT' || k === 'OPEN'; });
}
/* row 3 of the Stopping Merit sheet: which quotas the 1st-year students have, and which one was used */
function at_smNote_(sh) {
  if (!AT_SM_MEMO) return;
  var seen = AT_SM_MEMO.seen, txt = Object.keys(seen).sort(function (a, b) { return seen[b] - seen[a]; }).map(function (k) { return k + ' ' + seen[k]; }).join('  \u00b7  ');
  sh.getRange(3, 1, 1, sh.getLastColumn()).merge().setValue('Only OPEN QUOTA students counted.  Quotas found (1st Year): ' + txt)
    .setFontSize(9).setFontStyle('italic').setHorizontalAlignment('center').setWrap(true);
}

/* ---------- room of every section from the Timetable sheet (the room used most by that section) ---------- */
function at_key_(v) { return String(v == null ? '' : v).toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function at_ttRooms_() {
  var out = { byPcs: {}, byPcs2: {}, byPs: {}, sheet: AUTO_CFG.sourceSheet, found: false };
  var sh = at_ss_().getSheetByName(AUTO_CFG.sourceSheet);
  if (!sh || sh.getLastRow() < 2) return out;
  out.found = true;
  var v = sh.getDataRange().getDisplayValues();
  var h = v[0].map(function (x) { return String(x).trim().toLowerCase().replace(/\s+/g, ''); });
  var iP = h.indexOf('program'), iC = h.indexOf('class'), iPcs = h.indexOf('pcs'), iS = h.indexOf('section'), iR = h.indexOf('room');
  if (iR < 0) return out;
  var cnt = {};
  function add(map, k, room) { if (!k) return; var m = cnt[map + k] || (cnt[map + k] = {}); m[room] = (m[room] || 0) + 1; }
  for (var r = 1; r < v.length; r++) {
    var room = String(v[r][iR] || '').trim(); if (!room) continue;
    var prog = iP >= 0 ? at_key_(normalizeProgramForRoom_(v[r][iP])) : '', sec = iS >= 0 ? at_key_(v[r][iS]) : '', cls = iC >= 0 ? at_key_(v[r][iC]) : '';
    if (iPcs >= 0) add('pcs:', at_key_(v[r][iPcs]), room);
    if (iPcs >= 0 && prog) add('pcs2:', prog + '|' + at_key_(v[r][iPcs]), room);
    if (prog && sec) add('ps:', prog + '|' + cls + '|' + sec, room);
  }
  Object.keys(cnt).forEach(function (k) {
    var m = cnt[k], best = Object.keys(m).sort(function (a, b) { return m[b] - m[a]; })[0];
    var p = k.indexOf(':'), map = k.slice(0, p), key = k.slice(p + 1);
    (map === 'pcs' ? out.byPcs : map === 'pcs2' ? out.byPcs2 : out.byPs)[key] = best;
  });
  return out;
}
function at_ttRoomOf_(TT, o) {
  var prog = at_key_(normalizeProgramForRoom_(o.p));
  return TT.byPcs2[prog + '|' + at_key_(o.pcs)] || TT.byPs[prog + '|' + at_key_(o.c) + '|' + at_key_(o.s)] || TT.byPcs[at_key_(o.pcs)] || '';
}

/* ---------- edit one student (admission slip) in Transform-ALL-COLLEGE-STRENGTH ---------- */
function at_saveStudent_(req) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = at_ss_().getSheetByName(AT_CFG.STRENGTH);
    if (!sh) throw new Error("Sheet '" + AT_CFG.STRENGTH + "' not found.");
    var row = Number(req.row), nc = sh.getLastColumn();
    if (!(row >= 2 && row <= sh.getLastRow())) throw new Error('Row not found. Reload the students.');
    var C = at_cols_(sh.getRange(1, 1, 1, nc).getDisplayValues()[0]);
    var rg = sh.getRange(row, 1, 1, nc), cur = rg.getDisplayValues()[0], fx = rg.getFormulas()[0];
    if (String(cur[C.e] || '').trim() !== String(req.oldE || '').trim() || String(cur[C.n] || '').trim() !== String(req.oldN || '').trim())
      throw new Error('This row was changed in the sheet since you loaded it. Reload the students and try again.');
    var ch = req.changes || {}, writes = [], log = [];
    ['e', 'n', 'f', 'p', 'c', 's', 'pcs', 'q'].forEach(function (k) { if (k in ch && C[k] >= 0) writes.push([C[k], ch[k]]); });
    if (Array.isArray(ch.sub)) ch.sub.forEach(function (v, i) { if (C.sub[i] >= 0) writes.push([C.sub[i], v]); });
    writes.forEach(function (w) {
      var nv = String(w[1] == null ? '' : w[1]).trim();
      if (nv === String(cur[w[0]] || '').trim()) return;
      if (fx[w[0]]) throw new Error('Column "' + sh.getRange(1, w[0] + 1).getDisplayValue() + '" is a formula in the sheet, so it cannot be edited here.');
      sh.getRange(row, w[0] + 1).setValue(nv);
      log.push(sh.getRange(1, w[0] + 1).getDisplayValue() + ': ' + cur[w[0]] + ' -> ' + nv);
    });
    SpreadsheetApp.flush();
    return { row: row, changed: log };
  } finally { lock.releaseLock(); }
}

/* ---------- ROOM TAGS (Science Timetable = FSC, ARTS TIMETABLE = FA): one schedule per room ---------- */
function at_roomTags_() {
  var out = { fsc: [], fa: [], notes: [] };
  try { out.fsc = at_tagsFSC_(); } catch (e) { out.notes.push('FSC: ' + e.message); }
  try { out.fa = at_tagsFA_(); } catch (e) { out.notes.push('FA: ' + e.message); }
  return out;
}
function at_isRoman_(s) { return /^(I|II|III|IV|V|VI|VII|VIII|IX|X)$/i.test(String(s || '').trim()); }
function at_roman_(s) { var m = { I: 1, V: 5, X: 10, L: 50, C: 100 }, n = 0, p = 0; s = String(s || '').trim().toUpperCase(); for (var i = s.length - 1; i >= 0; i--) { var c = m[s[i]] || 0; n += c < p ? -c : c; p = c; } return n; }
function at_roomSplit_(raw) {                                // "R # 06/55" -> {room:"06", strength:"55"}
  var rest = String(raw || '').trim().toUpperCase().replace(/^R\s*#\s*/i, '').trim(), strength = '', room = rest;
  if (rest.indexOf('/') >= 0) { var parts = rest.split('/'), last = parts[parts.length - 1].trim(); if (/^\d+$/.test(last)) { strength = last; parts.pop(); room = parts.join('/').trim(); } }
  else { var pm = rest.match(/\((\d+)\)\s*$/); if (pm) { strength = pm[1]; room = rest.replace(/\(\d+\)\s*$/, '').trim(); } }
  return { room: room || rest, strength: strength };
}
function at_tagsFSC_() {                                     // same reading as generateRoomTagsFSC
  var sh = at_ss_().getSheetByName('Science Timetable');
  if (!sh) throw new Error("sheet 'Science Timetable' not found");
  var data = sh.getDataRange().getDisplayValues();
  if (data.length < 7) return [];
  var per = [], tim = data[5] || [];
  for (var i = 2; i < (data[4] || []).length; i++) { var p = String(data[4][i]).trim().toUpperCase(); if (!at_isRoman_(p)) break; per.push(p); }
  if (!per.length) throw new Error('no Roman period numbers in row 5');
  var merge = {};
  sh.getRange(1, 1, data.length, 1).getMergedRanges().forEach(function (m) { merge[m.getRow()] = m.getNumRows(); });
  var rooms = [], r = 6;
  while (r < data.length) {
    var cell = String(data[r][0] || '').trim();
    if (!/^R\s*#/i.test(cell)) { r++; continue; }
    var size = Math.max(3, merge[r + 1] || 3), rs = at_roomSplit_(cell), cls = String(data[r][1] || '').trim().toUpperCase(), rows = [];
    for (var off = 0; off < size; off += 3) {
      var s = data[r + off], t = data[r + off + 1] || [], d = data[r + off + 2] || [];
      if (!s) continue;
      for (var k = 0; k < per.length; k++) {
        var subj = String(s[k + 2] || '').trim().toUpperCase(), tea = String(t[k + 2] || '').trim().toUpperCase(), day = String(d[k + 2] || '').trim().toUpperCase() || 'M-S';
        if (!subj && !tea) continue;
        if (subj === 'SPORTS') tea = '---';
        rows.push([per[k], String(tim[k + 2] || '').trim().toUpperCase(), cls, tea, subj, day]);
      }
    }
    rows.sort(function (a, b) { return at_roman_(a[0]) - at_roman_(b[0]); });
    rooms.push({ room: rs.room, strength: rs.strength, cls: cls, rows: rows });
    r += size;
  }
  return rooms;
}
function at_tagsFA_() {                                      // same result as generateRoomTagsFA
  var sh = at_ss_().getSheetByName('ARTS TIMETABLE');
  if (!sh) throw new Error("sheet 'ARTS TIMETABLE' not found");
  var data = sh.getDataRange().getDisplayValues();
  if (data.length < 10) return [];
  var per = [], tim = [], pcol = {};
  for (var c = 2; c < data[4].length; c += 2) { var p = String(data[4][c] || '').trim().toUpperCase(); if (p) { pcol[c] = per.length; per.push(p); tim.push(String(data[5][c] || '').trim().toUpperCase()); } }
  var get = function (r, c) { return String(((data[r] || [])[c]) || '').trim().toUpperCase(); };
  var roomLike = function (x) { return /^R\s*#/i.test(x) || /\/\s*\d+\s*$/.test(x); };
  var cells = [];                                            // [subjectRow, col] (0-based)
  if (typeof buildBlockCellAddresses === 'function') {        // the original block map, when it is in this project
    var b = buildBlockCellAddresses(8, 49), rc = function (a) { var m = String(a || '').match(/([A-Za-z]+)\$?(\d+)/); if (!m) return null; var cc = m[1].toUpperCase().split('').reduce(function (x, ch) { return x * 26 + ch.charCodeAt(0) - 64; }, 0); return [Number(m[2]) - 1, cc - 1]; };
    for (var i = 0; i < b.subjectCells.length; i++) {
      var S = rc(b.subjectCells[i]), T = rc(b.teacherCells[i]), G = rc(b.groupCells[i]), R = rc(b.roomCells[i]), D = rc(b.dayCells[i]), K = rc(b.classCells[i]);
      cells.push({ r: S[0], c: S[1], subj: get(S[0], S[1]), tea: T ? get(T[0], T[1]) : '', grp: G ? get(G[0], G[1]) : '', room: R ? get(R[0], R[1]) : '', day: D ? get(D[0], D[1]) : '', cls: K ? get(K[0], K[1]) : '' });
    }
  } else {                                                   // block = subject (group beside it) / teacher / room (days beside it)
    for (var r = 7; r <= Math.min(48, data.length - 3); r++) for (var c2 = 2; c2 < data[r].length; c2 += 2) {
      var subj = get(r, c2), room = get(r + 2, c2);
      if (!subj || !room || roomLike(subj) || !roomLike(room) || roomLike(get(r + 1, c2))) continue;
      var cls = ''; for (var u = r; u >= Math.max(7, r - 2); u--) { var x = get(u, 1); if (x && !roomLike(x)) { cls = x; break; } }
      cells.push({ r: r, c: c2, subj: subj, tea: get(r + 1, c2), grp: get(r, c2 + 1), room: room, day: get(r + 2, c2 + 1), cls: cls });
    }
  }
  var map = {}, order = [];
  cells.forEach(function (x) {
    if (!x.subj || !x.room) return;
    var rv = x.room, name = rv, strength = '';
    if (rv.indexOf('/') >= 0) { var parts = rv.split('/'), last = parts.pop().trim(); if (/^\d+$/.test(last)) strength = last; name = parts.join('/').trim(); }
    name = (name || rv).replace(/^R\s*#\s*/i, '').trim();
    var prog = (x.r + 1 >= 8 && x.r + 1 <= 28) ? 'F.A-I' : (x.r + 1 >= 29 && x.r + 1 <= 49) ? 'F.A-II' : '';
    var cls = (x.cls || prog) + (x.grp ? ' ' + x.grp : '');
    var k = Math.floor((x.c - 2) / 2);
    if (!map[name]) { map[name] = { room: name, strength: strength, cls: cls, rows: [] }; order.push(name); }
    if (!map[name].strength && strength) map[name].strength = strength;
    map[name].rows.push([per[k] || '', tim[k] || '', cls, x.subj === 'SPORTS' ? '---' : x.tea, x.subj, x.day || 'M-S']);
  });
  return order.map(function (n) { map[n].rows.sort(function (a, b) { return at_roman_(a[0]) - at_roman_(b[0]); }); return map[n]; });
}

/* ---------- save a PDF / Excel made in the app to Drive ---------- */
function at_save_(req) {
  var name = String(req.name || 'JCW report').replace(/[\\\/]/g, '-');
  var blob = Utilities.newBlob(Utilities.base64Decode(String(req.data || '')), String(req.mime || 'application/pdf'), name);
  var fname = String(req.folder || AT_CFG.FOLDER);
  var it = DriveApp.getFoldersByName(fname), folder = it.hasNext() ? it.next() : DriveApp.createFolder(fname);
  var old = folder.getFilesByName(name);
  while (old.hasNext()) old.next().setTrashed(true);
  var f = folder.createFile(blob);
  return { name: f.getName(), url: f.getUrl(), folder: fname };
}



/* ====================================================================================================
   1. AUTO TIMETABLE GENERATOR (original)
   ==================================================================================================== */

// ====================================================================
//  AUTO TIMETABLE GENERATOR
//  Builds the printable "AutoTimetable" sheet from the "Timetable" sheet
//  (same shape as Timetable_2026.xlsx: Room | Section | Period I..V,
//   each section = blocks of 3 rows: Subject / Teacher / Days)
// ====================================================================

const AUTO_CFG = {

  sourceSheet: 'Timetable',
  targetSheet: 'AutoTimetable',

  college: 'JINNAH COLLEGE FOR WOMEN, UNIVERSITY OF PESHAWAR',
  streamLabel: 'F.SC',

  // Programs (Program column) that belong to this printed timetable.
  // To print Arts later, copy this function's call with another config.
  programs: [
    'Pre-Medical (Jinnah College)',
    'Pre-Engineering (Jinnah College)',
    'FCS (Jinnah College)',
    
  ],

  // Period timings (column headers). Extra periods beyond this list get blank time.
  times: [
    '08:00 – 09:00',
    '09:00 – 10:00',
    '10:00 – 11:00',
    '11:00 – 12:00',
    '12:00 – 01:00'
  ],

  roman: ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X'],

  classRoman: { '1st Year': 'I', '2nd Year': 'II' },

  // false = period header rows only on the first table (as in your Excel shape)
  repeatHeader: false,

  missingTeacherColor: '#F4CCCC'
};


function generateAutoTimetable() {

  const cfg = Object.assign({}, AUTO_CFG, AT_AUTO_OVERRIDE || {});   // programs / stream label from the admin tools page
  const ss = at_ss_();

  // ================= READ SOURCE =================
  const src = ss.getSheetByName(cfg.sourceSheet);
  if (!src) throw new Error("Sheet '" + cfg.sourceSheet + "' was not found.");

  const data = src.getDataRange().getValues();
  if (data.length < 2) throw new Error("No data found in '" + cfg.sourceSheet + "'.");

  const headers = data[0].map(h => String(h).trim());

  const col = name => {
    const i = headers.indexOf(name);
    if (i === -1) throw new Error("Required column '" + name + "' was not found.");
    return i;
  };

  const cSubject   = col('Subject');
  const cTeacher   = col('Teacher');
  const cProgram   = col('Program');
  const cClass     = col('Class');
  const cPCS       = col('PCS');
  const cSession   = col('Session');
  const cRoom      = col('Room');
  const cTheory    = col('Theory');
  const cPractical = col('Practical');
  const cPeriod    = col('Period');

  const str = v => (v === null || v === undefined) ? '' : String(v).trim();

  // ================= GROUP: class -> section (PCS) -> period -> entries =================
  const byClass = {};
  let session = cfg.session || '';                      // admin tools page: the current session
  let maxPeriod = 0;

  for (let r = 1; r < data.length; r++) {

    const row = data[r];

    if (!cfg.programs.includes(str(row[cProgram]))) continue;

    const subject = str(row[cSubject]);
    const period = parseInt(row[cPeriod], 10);

    if (!subject || !period) continue;

    const cls = str(row[cClass]);
    const pcs = str(row[cPCS]);

    if (!session) session = str(row[cSession]);
    if (period > maxPeriod) maxPeriod = period;

    if (!byClass[cls]) byClass[cls] = { order: [], map: {} };

    const grp = byClass[cls];

    if (!grp.map[pcs]) {
      grp.map[pcs] = { pcs: pcs, rooms: [], periods: {} };
      grp.order.push(pcs);
    }

    const sec = grp.map[pcs];

    const room = str(row[cRoom]);
    if (room && !sec.rooms.includes(room)) sec.rooms.push(room);

    if (!sec.periods[period]) sec.periods[period] = [];

    sec.periods[period].push({
      subject: subject,
      teacher: str(row[cTeacher]),
      days: buildDaysText_(str(row[cTheory]), str(row[cPractical]))
    });
  }

  const classNames = Object.keys(byClass).sort();

  if (!classNames.length) {
    throw new Error("No rows matched the configured programs.");
  }

  // ================= BUILD OUTPUT ARRAYS =================
  const numPeriods = Math.max(cfg.times.length, maxPeriod);
  const W = 2 + numPeriods;

  const values = [];
  const bgs = [];
  const ops = { titles: [], romanRows: [], timeRows: [], blocks: [], tables: [] };

  const blankRow = () => new Array(W).fill('');
  const whiteRow = () => new Array(W).fill('#ffffff');

  function pushRow(vals, bg) {
    values.push(vals);
    bgs.push(bg || whiteRow());
    return values.length;               // 1-based row number
  }

  classNames.forEach((cls, ti) => {

    const roman = cfg.classRoman[cls] || cls;

    // ---- titles ----
    let r = pushRow(padTitle_(cfg.college, W));
    ops.titles.push(r);

    pushRow(blankRow());

    r = pushRow(padTitle_(
      cfg.streamLabel + ' - ' + roman + ' TIME TABLE (' + session + ')', W
    ));
    ops.titles.push(r);

    // ---- period header (first table only, unless repeatHeader) ----
    if (ti === 0 || cfg.repeatHeader) {

      pushRow(blankRow());

      const romanVals = blankRow();
      const timeVals = blankRow();
      const timeBg = whiteRow();

      for (let p = 0; p < numPeriods; p++) {
        romanVals[2 + p] = cfg.roman[p] || String(p + 1);
        timeVals[2 + p] = cfg.times[p] || '';
        timeBg[2 + p] = '#999999';
      }

      ops.romanRows.push(pushRow(romanVals));
      ops.timeRows.push(pushRow(timeVals, timeBg));
    }

    const tableStart = values.length + 1;

    // ---- sections ----
    const grp = byClass[cls];

    grp.order.forEach(pcs => {

      const sec = grp.map[pcs];

      let slots = 1;
      for (let p = 1; p <= numPeriods; p++) {
        const n = (sec.periods[p] || []).length;
        if (n > slots) slots = n;
      }

      const blockStart = values.length + 1;

      for (let k = 0; k < slots; k++) {

        const rSub = blankRow(), rTea = blankRow(), rDay = blankRow();
        const bTea = whiteRow();

        if (k === 0) {
          rSub[0] = sec.rooms.join(' / ');
          rSub[1] = sec.pcs;
        }

        for (let p = 1; p <= numPeriods; p++) {

          const e = (sec.periods[p] || [])[k];
          if (!e) continue;

          rSub[1 + p] = e.subject;
          rTea[1 + p] = e.teacher;
          rDay[1 + p] = e.days;

          if (!e.teacher) bTea[1 + p] = cfg.missingTeacherColor;
        }

        pushRow(rSub);
        pushRow(rTea, bTea);
        pushRow(rDay);
      }

      ops.blocks.push({ start: blockStart, rows: slots * 3 });
    });

    ops.tables.push({ start: tableStart, rows: values.length - tableStart + 1 });
  });

  const n = values.length;

  // ================= PREPARE TARGET SHEET =================
  let sheet = ss.getSheetByName(cfg.targetSheet);
  if (!sheet) sheet = ss.insertSheet(cfg.targetSheet);

  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clear();

  if (sheet.getMaxColumns() < W) sheet.insertColumnsAfter(sheet.getMaxColumns(), W - sheet.getMaxColumns());
  if (sheet.getMaxRows() < n) sheet.insertRowsAfter(sheet.getMaxRows(), n - sheet.getMaxRows());

  // ================= WRITE (single call each) =================
  const all = sheet.getRange(1, 1, n, W);

  all.setValues(values);
  all.setBackgrounds(bgs);
  all.setFontFamily('Arial')
     .setFontSize(14)
     .setFontWeight('normal')
     .setHorizontalAlignment('left')
     .setVerticalAlignment('top')
     .setWrap(true);

  // ================= SIZES =================
  sheet.setRowHeights(1, n, 26);

  sheet.setColumnWidth(1, 190);
  sheet.setColumnWidth(2, 130);
  for (let c = 3; c <= W; c++) sheet.setColumnWidth(c, 260);

  const MED = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
  const THICK = SpreadsheetApp.BorderStyle.SOLID;
  const THICKEST = SpreadsheetApp.BorderStyle.SOLID_THICK;

  // ================= TITLES =================
  ops.titles.forEach(r => {
    const rng = sheet.getRange(r, 1, 1, W);
    rng.merge()
       .setFontSize(18)
       .setFontWeight('bold')
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle')
       .setBorder(true, true, true, true, false, false, '#000000', MED);
    sheet.setRowHeight(r, 34);
  });

  // ================= PERIOD HEADER =================
  ops.romanRows.forEach(r => {
    sheet.getRange(r, 3, 1, numPeriods)
         .setFontWeight('bold')
         .setHorizontalAlignment('center')
         .setBorder(true, true, true, true, true, true, '#000000', MED);
  });

  ops.timeRows.forEach(r => {
    sheet.getRange(r, 3, 1, numPeriods)
         .setHorizontalAlignment('center')
         .setBorder(true, true, true, true, true, true, '#000000', MED);
  });

  // ================= TABLE GRID (medium) =================
  ops.tables.forEach(t => {
    sheet.getRange(t.start, 1, t.rows, W)
         .setBorder(true, true, true, true, true, true, '#000000', MED);
  });

  // ================= SECTION BLOCKS =================
  ops.blocks.forEach(b => {

    const roomCell = sheet.getRange(b.start, 1, b.rows, 1);
    const pcsCell = sheet.getRange(b.start, 2, b.rows, 1);

    roomCell.merge();
    pcsCell.merge();

    sheet.getRange(b.start, 1, b.rows, 2)
         .setFontWeight('bold')
         .setVerticalAlignment('middle')
         .setBorder(true, true, true, true, false, false, '#000000', THICKEST);

    // thick outline around the whole section block
    sheet.getRange(b.start, 1, b.rows, W)
         .setBorder(true, true, true, true, false, false, '#000000', THICKEST);
  });

  SpreadsheetApp.flush();
  // ss.setActiveSheet(sheet);   (no open spreadsheet in a web app)

  Logger.log(
    '✅ AutoTimetable generated: ' +
    ops.blocks.length + ' sections, ' + n + ' rows'
  );
}


// ====================================================================
//  HELPERS
// ====================================================================

// Theory + Practical columns -> single days text.
//   theory "M-TH", practical "F"    -> "M-TH,F(P)"
//   theory "",     practical "S"    -> "S(P)"
//   theory "M,TU", practical ""     -> "M,TU"
function buildDaysText_(theory, practical) {

  const parts = [];

  if (theory) parts.push(theory);

  if (practical) {
    practical
      .split(',')
      .map(d => d.trim())
      .filter(Boolean)
      .forEach(d => parts.push(d + '(P)'));
  }

  return parts.join(',');
}

function padTitle_(text, width) {
  const arr = new Array(width).fill('');
  arr[0] = text;
  return arr;
}

/* ====================================================================================================
   2. TEACHER WORKLOAD BY STATUS (original)
   ==================================================================================================== */

function generateTeacherWorkloadByStatus() {

  const startTime = new Date();

  const ss = at_ss_();

  const scienceSheet = ss.getSheetByName("Science Timetable");
  const artsSheet = ss.getSheetByName("ARTS TIMETABLE");
  const dropdownSheet = ss.getSheetByName("DropdownData");

  if (!scienceSheet) {
    throw new Error("Sheet 'Science Timetable' not found!");
  }

  if (!artsSheet) {
    throw new Error("Sheet 'ARTS TIMETABLE' not found!");
  }

  if (!dropdownSheet) {
    throw new Error("Sheet 'DropdownData' not found!");
  }

  const scienceData = scienceSheet.getDataRange().getValues();
  const artsData = artsSheet.getDataRange().getValues();
  const dropdownData = dropdownSheet.getDataRange().getValues();

  // ============================================================
  // EVENING SCIENCE SECTIONS
  // ============================================================
  //
  // These sections are ALWAYS treated as EVENING.
  //
  // IMPORTANT:
  // A teacher's DropdownData status is NOT used to decide
  // whether an evening assignment is Visiting.
  //
  // Therefore:
  //
  // Permanent teacher + L2
  //       -> Evening Visiting Workload
  //
  // Visiting teacher + L2
  //       -> Evening Visiting Workload
  //
  // ============================================================

  const EVENING_SCIENCE_SECTIONS = [
    "L2",
    "M2",
    "N2",
    "O2",
    "P2",
    "S1",
    "T1"
  ];

  // ============================================================
  // NORMALIZATION FUNCTIONS
  // ============================================================

  function normalizeName(value) {

    if (value === null || value === undefined) {
      return "";
    }

    return value
      .toString()
      .replace(/\s+/g, " ")
      .trim()
      .toUpperCase();
  }

  function cleanText(value) {

    if (value === null || value === undefined) {
      return "";
    }

    return value
      .toString()
      .replace(/\s+/g, " ")
      .trim();
  }

  // ============================================================
  // CHECK EVENING SCIENCE SECTION
  // ============================================================

  function isEveningScienceSection(className) {

    const text = cleanText(className)
      .toUpperCase()
      .replace(/\s+/g, "");

    if (!text) {
      return false;
    }

    // Exact match
    if (EVENING_SCIENCE_SECTIONS.indexOf(text) !== -1) {
      return true;
    }

    // Also handle text such as:
    // Pre-Medical L2
    // Pre-Engineering L2
    // etc.
    for (const section of EVENING_SCIENCE_SECTIONS) {

      if (text.endsWith(section)) {
        return true;
      }
    }

    return false;
  }

  // ============================================================
  // CHECK ARTS EVENING
  // ============================================================
  //
  // Arts currently has NO evening classes.
  //
  // This is retained only as a safety check in case the
  // timetable contains the word EVENING in future.
  // ============================================================

  function isEveningArtsAssignment(room, className, days) {

    const combined = (
      cleanText(room) +
      " " +
      cleanText(className) +
      " " +
      cleanText(days)
    ).toUpperCase();

    return combined.indexOf("EVENING") !== -1;
  }

  // ============================================================
  // TEACHER INFORMATION FROM DROPDOWNDATA
  //
  // IMPORTANT:
  // A teacher can accidentally occur more than once in
  // DropdownData.
  //
  // DO NOT allow a later duplicate row to overwrite the
  // teacher's actual department.
  //
  // The record with the LOWEST S.No is treated as the
  // teacher's primary/canonical record.
  //
  // Columns:
  // D = Teacher
  // E = Designation
  // F = Status
  // L = Department
  // M = S.No
  // ============================================================

  const teacherInfoMap = {};

  for (let i = 1; i < dropdownData.length; i++) {

    const teacher =
      normalizeName(dropdownData[i][3]);

    if (!teacher) {
      continue;
    }

    const designation =
      cleanText(dropdownData[i][4]);

    const status =
      cleanText(dropdownData[i][5]);

    const rawDepartment =
      cleanText(dropdownData[i][11]);

    const department =
      rawDepartment
        ? "DEPARTMENT OF " + rawDepartment
        : "UNKNOWN DEPARTMENT";

    const snValue =
      Number(dropdownData[i][12]);

    const sn =
      !isNaN(snValue) && snValue > 0
        ? snValue
        : 9999;

    // ----------------------------------------------------------
    // FIRST RECORD
    // ----------------------------------------------------------

    if (!teacherInfoMap[teacher]) {

      teacherInfoMap[teacher] = {
        designation: designation,
        status: status,
        department: department,
        sn: sn
      };

      continue;
    }

    // ----------------------------------------------------------
    // DUPLICATE TEACHER
    //
    // Keep the record having the LOWEST S.No.
    // ----------------------------------------------------------

    const existing =
      teacherInfoMap[teacher];

    if (sn < existing.sn) {

      Logger.log(
        "PRIMARY TEACHER RECORD UPDATED: " +
        teacher +
        " | OLD DEPT=" +
        existing.department +
        " | NEW DEPT=" +
        department +
        " | OLD S.No=" +
        existing.sn +
        " | NEW S.No=" +
        sn
      );

      teacherInfoMap[teacher] = {
        designation: designation,
        status: status,
        department: department,
        sn: sn
      };

    } else {

      Logger.log(
        "DUPLICATE TEACHER IGNORED: " +
        teacher +
        " | DUPLICATE DEPT=" +
        department +
        " | DUPLICATE S.No=" +
        sn +
        " | PRIMARY DEPT=" +
        existing.department +
        " | PRIMARY S.No=" +
        existing.sn
      );
    }
  }

  Logger.log(
    "Teachers loaded: " +
    Object.keys(teacherInfoMap).length
  );

  // ============================================================
  // CREATE TEACHER ENTRY
  // ============================================================

  function getTeacherEntry(
    workloadMap,
    teacherName,
    info
  ) {

    const dept =
      info.department;

    if (!workloadMap[dept]) {
      workloadMap[dept] = {};
    }

    if (!workloadMap[dept][teacherName]) {

      workloadMap[dept][teacherName] = {

        classes: [],

        workload: 0,

        designation:
          info.designation,

        sn:
          info.sn
      };
    }

    return workloadMap[dept][teacherName];
  }

  // ============================================================
  // NORMALIZE DAY EXPRESSION
  //
  // IMPORTANT:
  //   F(P)            -> F
  //   F (P)           -> F
  //   M(Group a,b)    -> M
  //   M (Group A,B)   -> M
  //
  // Anything inside parentheses is NOT a day.
  // ============================================================

  function normalizeDayExpression(value) {

    if (value === null || value === undefined) {
      return "";
    }

    let text =
      value
        .toString()
        .toUpperCase()
        .trim();

    if (!text) {
      return "";
    }

    // Remove everything inside parentheses
    text =
      text.replace(/\([^)]*\)/g, "");

    text =
      text.replace(/\s+/g, "")
        .trim();

    text =
      text
        .replace(/THURSDAY/g, "TH")
        .replace(/THURS/g, "TH")
        .replace(/THU/g, "TH")
        .replace(/TUESDAY/g, "TU")
        .replace(/TUES/g, "TU")
        .replace(/TUE/g, "TU")
        .replace(/MONDAY/g, "M")
        .replace(/MON/g, "M")
        .replace(/WEDNESDAY/g, "W")
        .replace(/WED/g, "W")
        .replace(/FRIDAY/g, "F")
        .replace(/FRI/g, "F")
        .replace(/SATURDAY/g, "S")
        .replace(/SAT/g, "S")
        .replace(/YU/g, "TU");

    return text;
  }

  // ============================================================
  // EXPAND DAY TEXT
  // ============================================================

  function expandDays(dayText) {

    const text =
      normalizeDayExpression(dayText);

    if (!text) {
      return [];
    }

    const dayIndex = {
      M: 0,
      TU: 1,
      W: 2,
      TH: 3,
      F: 4,
      S: 5
    };

    const dayNames = [
      "M",
      "TU",
      "W",
      "TH",
      "F",
      "S"
    ];

    function expandToken(token) {

      token =
        token.trim();

      if (!token) {
        return [];
      }

      // --------------------------------------------------------
      // RANGE
      // --------------------------------------------------------

      if (token.includes("-")) {

        const parts =
          token.split("-");

        if (parts.length === 2) {

          const start =
            parts[0];

          const end =
            parts[1];

          if (
            dayIndex[start] !== undefined &&
            dayIndex[end] !== undefined
          ) {

            const result = [];

            let i =
              dayIndex[start];

            const endIndex =
              dayIndex[end];

            while (true) {

              result.push(
                dayNames[i]
              );

              if (i === endIndex) {
                break;
              }

              i =
                (i + 1) % 6;
            }

            return result;
          }
        }

        return [];
      }

      // --------------------------------------------------------
      // SINGLE DAY
      // --------------------------------------------------------

      if (
        dayIndex[token] !== undefined
      ) {

        return [token];
      }

      return [];
    }

    // ----------------------------------------------------------
    // COMMA SEPARATED
    // ----------------------------------------------------------

    if (text.includes(",")) {

      const result = [];

      const parts =
        text
          .split(",")
          .filter(Boolean);

      for (const part of parts) {

        result.push(
          ...expandToken(part)
        );
      }

      return result;
    }

    return expandToken(text);
  }

  // ============================================================
  // COUNT DAYS
  // ============================================================

  function countDays(dayText) {

    const text =
      normalizeDayExpression(dayText);

    if (!text) {
      return 0;
    }

    const dayIndex = {
      M: 0,
      TU: 1,
      W: 2,
      TH: 3,
      F: 4,
      S: 5
    };

    function countToken(token) {

      token =
        token.trim();

      if (!token) {
        return 0;
      }

      // --------------------------------------------------------
      // RANGE
      // --------------------------------------------------------

      if (token.indexOf("-") !== -1) {

        const parts =
          token.split("-");

        if (parts.length !== 2) {
          return 0;
        }

        const start =
          parts[0];

        const end =
          parts[1];

        if (
          dayIndex[start] === undefined ||
          dayIndex[end] === undefined
        ) {
          return 0;
        }

        const startIndex =
          dayIndex[start];

        const endIndex =
          dayIndex[end];

        if (endIndex >= startIndex) {

          return (
            endIndex -
            startIndex +
            1
          );
        }

        // Wrapped range
        return (
          (6 - startIndex) +
          endIndex +
          1
        );
      }

      // --------------------------------------------------------
      // SINGLE DAY
      // --------------------------------------------------------

      if (
        dayIndex[token] !== undefined
      ) {
        return 1;
      }

      return 0;
    }

    // ----------------------------------------------------------
    // COMMA SEPARATED
    // ----------------------------------------------------------

    if (text.indexOf(",") !== -1) {

      return text
        .split(",")
        .filter(Boolean)
        .reduce(
          (total, token) =>
            total + countToken(token),
          0
        );
    }

    return countToken(text);
  }

  // ============================================================
  // FORMAT COMBINED DAYS
  // ============================================================

  function formatCombinedDays(dayValues) {

    if (
      !dayValues ||
      dayValues.length === 0
    ) {
      return "";
    }

    const order = [
      "M",
      "TU",
      "W",
      "TH",
      "F",
      "S"
    ];

    const normal = [];
    const practical = [];

    // ----------------------------------------------------------
    // NORMALIZE DISPLAY EXPRESSION
    // ----------------------------------------------------------

    function normalizeDisplayDay(value) {

      let text =
        cleanText(value);

      if (!text) {
        return "";
      }

      // Group annotation = practical
      if (
        /\(\s*GROUP\b[^)]*\)/i.test(text)
      ) {

        text =
          text.replace(
            /\s*\(\s*GROUP\b[^)]*\)/gi,
            ""
          );

        text =
          text.trim();

        return text + "(P)";
      }

      // Existing practical marker
      if (
        /\(\s*P\s*\)/i.test(text)
      ) {

        text =
          text.replace(
            /\s*\(\s*P\s*\)/gi,
            "(P)"
          );

        return text;
      }

      return text;
    }

    // ----------------------------------------------------------
    // SEPARATE NORMAL / PRACTICAL
    // ----------------------------------------------------------

    for (const rawValue of dayValues) {

      const original =
        cleanText(rawValue);

      if (!original) {
        continue;
      }

      const displayValue =
        normalizeDisplayDay(original);

      if (!displayValue) {
        continue;
      }

      if (
        displayValue.includes("(P)")
      ) {

        practical.push(
          displayValue
        );

      } else {

        normal.push(
          displayValue
        );
      }
    }

    // ----------------------------------------------------------
    // COUNT NORMAL DAYS
    // ----------------------------------------------------------

    const counts = {
      M: 0,
      TU: 0,
      W: 0,
      TH: 0,
      F: 0,
      S: 0
    };

    for (const value of normal) {

      const expanded =
        expandDays(value);

      for (const day of expanded) {

        if (
          counts[day] !== undefined
        ) {
          counts[day]++;
        }
      }
    }

    // ----------------------------------------------------------
    // BUILD NORMAL DAY PARTS
    // ----------------------------------------------------------

    const normalParts = [];

    for (const day of order) {

      const count =
        counts[day];

      if (count === 0) {
        continue;
      }

      if (count === 1) {

        normalParts.push(day);

      } else {

        normalParts.push(
          count + day
        );
      }
    }

    // ----------------------------------------------------------
    // COMPACT NORMAL DAYS
    // ----------------------------------------------------------

    let normalResult = "";

    if (
      normalParts.length === 1
    ) {

      normalResult =
        normalParts[0];

    } else if (
      normalParts.length > 1
    ) {

      const rebuilt = [];

      let start = 0;

      while (
        start < normalParts.length
      ) {

        let end =
          start;

        while (
          end + 1 <
          normalParts.length
        ) {

          const currentDay =
            normalParts[end]
              .replace(/^\d+/, "");

          const nextDay =
            normalParts[end + 1]
              .replace(/^\d+/, "");

          const currentIndex =
            order.indexOf(currentDay);

          const nextIndex =
            order.indexOf(nextDay);

          if (
            currentIndex >= 0 &&
            nextIndex ===
              currentIndex + 1
          ) {

            end++;

          } else {

            break;
          }
        }

        const segment =
          normalParts.slice(
            start,
            end + 1
          );

        if (
          segment.length >= 2
        ) {

          const first =
            segment[0]
              .replace(/^\d+/, "");

          const last =
            segment[
              segment.length - 1
            ];

          rebuilt.push(
            first +
            "-" +
            last
          );

        } else {

          rebuilt.push(
            ...segment
          );
        }

        start =
          end + 1;
      }

      normalResult =
        rebuilt.join(",");
    }

    // ----------------------------------------------------------
    // COMBINE
    // ----------------------------------------------------------

    const result = [];

    if (normalResult) {
      result.push(
        normalResult
      );
    }

    practical.forEach(
      value => result.push(value)
    );

    return result.join(", ");
  }

  // ============================================================
  // ADD / MERGE CLASS + DAYS
  // ============================================================

  function addWorkload(
    workloadMap,
    teacherName,
    info,
    className,
    days
  ) {

    className =
      cleanText(className);

    days =
      cleanText(days);

    if (!className) {
      return;
    }

    const entry =
      getTeacherEntry(
        workloadMap,
        teacherName,
        info
      );

    // ----------------------------------------------------------
    // FIND SAME CLASS
    // ----------------------------------------------------------

    let existingClass = null;

    for (const item of entry.classes) {

      if (
        item &&
        typeof item === "object" &&
        item.className === className
      ) {

        existingClass =
          item;

        break;
      }
    }

    // ----------------------------------------------------------
    // FIRST OCCURRENCE
    // ----------------------------------------------------------

    if (!existingClass) {

      existingClass = {
        className: className,
        dayValues: []
      };

      entry.classes.push(
        existingClass
      );
    }

    // ----------------------------------------------------------
    // STORE ORIGINAL DAY EXPRESSION
    // ----------------------------------------------------------

    if (days) {

      existingClass.dayValues.push(
        days
      );
    }

    // ----------------------------------------------------------
    // WORKLOAD
    // ----------------------------------------------------------

    const dayCount =
      countDays(days);

    entry.workload +=
      dayCount;

    Logger.log(
      "WORKLOAD ADD: " +
      teacherName +
      " | CLASS=" +
      className +
      " | DAYS=" +
      days +
      " | +" +
      dayCount +
      " | TOTAL=" +
      entry.workload
    );
  }

  // ============================================================
  // SCIENCE TIMETABLE
  // ============================================================
  //
  // NEW PARAMETER:
  //
  // targetStatus:
  //   Permanent / Visiting / ANY
  //
  // mode:
  //   PERMANENT
  //   MORNING_VISITING
  //   EVENING_VISITING
  //
  // ============================================================

  function processScienceTimetable(
    targetStatus,
    mode
  ) {

    const map = {};

    const PERIOD_START_COL = 2; // C
    const PERIOD_END_COL = 8;   // I

    Logger.log(
      "======================================"
    );

    Logger.log(
      "PROCESSING SCIENCE: " +
      mode
    );

    Logger.log(
      "======================================"
    );

    for (
      let r = 0;
      r < scienceData.length;
      r++
    ) {

      for (
        let c = PERIOD_START_COL;
        c <= PERIOD_END_COL;
        c++
      ) {

        // ------------------------------------------------------
        // CURRENT CELL = TEACHER
        // ------------------------------------------------------

        const teacherCell =
          cleanText(
            scienceData[r][c]
          );

        if (!teacherCell) {
          continue;
        }

        const teacherName =
          normalizeName(
            teacherCell
          );

        const info =
          teacherInfoMap[
            teacherName
          ];

        if (!info) {
          continue;
        }

        // ------------------------------------------------------
        // SUBJECT
        // ------------------------------------------------------

        let subject = "";

        if (r - 1 >= 0) {

          const immediateSubject =
            cleanText(
              scienceData[r - 1][c]
            );

          if (
            immediateSubject &&
            !immediateSubject
              .toUpperCase()
              .startsWith("R #")
          ) {

            subject =
              immediateSubject;
          }
        }

        // ------------------------------------------------------
        // FALLBACK SUBJECT SEARCH
        // ------------------------------------------------------

        if (!subject) {

          for (
            let rr = r - 2;
            rr >= Math.max(0, r - 5);
            rr--
          ) {

            const candidate =
              cleanText(
                scienceData[rr][c]
              );

            if (!candidate) {
              continue;
            }

            if (
              candidate
                .toUpperCase()
                .startsWith("R #")
            ) {
              continue;
            }

            if (
              countDays(candidate) > 0
            ) {
              continue;
            }

            subject =
              candidate;

            break;
          }
        }

        // ------------------------------------------------------
        // DAYS
        // ------------------------------------------------------

        let days = "";

        if (
          r + 1 <
          scienceData.length
        ) {

          const immediateDays =
            cleanText(
              scienceData[r + 1][c]
            );

          if (
            immediateDays &&
            countDays(immediateDays) > 0
          ) {

            days =
              immediateDays;
          }
        }

        // ------------------------------------------------------
        // FALLBACK DAYS SEARCH
        // ------------------------------------------------------

        if (!days) {

          for (
            let rr = r + 2;
            rr <= Math.min(
              scienceData.length - 1,
              r + 5
            );
            rr++
          ) {

            const candidate =
              cleanText(
                scienceData[rr][c]
              );

            if (!candidate) {
              continue;
            }

            if (
              countDays(candidate) > 0
            ) {

              days =
                candidate;

              break;
            }
          }
        }

        // ------------------------------------------------------
        // CLASS NAME
        // ------------------------------------------------------

        let className = "";

        for (
          let rr = r;
          rr >= Math.max(0, r - 10);
          rr--
        ) {

          const candidate =
            cleanText(
              scienceData[rr][1]
            );

          if (!candidate) {
            continue;
          }

          if (
            candidate
              .toUpperCase()
              .startsWith("R #")
          ) {
            continue;
          }

          className =
            candidate;

          break;
        }

        // ------------------------------------------------------
        // VALIDATION
        // ------------------------------------------------------

        if (
          !subject ||
          !days ||
          !className
        ) {

          Logger.log(
            "WARNING: Incomplete Science record"
          );

          Logger.log(
            "Teacher = " +
            teacherName
          );

          Logger.log(
            "Row = " +
            (r + 1)
          );

          Logger.log(
            "Col = " +
            (c + 1)
          );

          continue;
        }

        // ======================================================
        // EVENING DETECTION
        // ======================================================

        const isEvening =
          isEveningScienceSection(
            className
          );

        // ======================================================
        // ASSIGNMENT FILTER
        // ======================================================
        //
        // PERMANENT:
        //   Permanent + MORNING only
        //
        // MORNING VISITING:
        //   Visiting + MORNING only
        //
        // EVENING VISITING:
        //   ANY teacher + EVENING
        //
        // ======================================================

        if (mode === "PERMANENT") {

          if (
            info.status !== "Permanent"
          ) {
            continue;
          }

          if (isEvening) {

            Logger.log(
              "PERMANENT TEACHER EVENING " +
              "ASSIGNMENT MOVED TO " +
              "EVENING VISITING: " +
              teacherName +
              " | " +
              className
            );

            continue;
          }
        }

        if (
          mode === "MORNING_VISITING"
        ) {

          if (
            info.status !== "Visiting"
          ) {
            continue;
          }

          if (isEvening) {
            continue;
          }
        }

        if (
          mode === "EVENING_VISITING"
        ) {

          // IMPORTANT:
          // DO NOT CHECK info.status.
          //
          // Permanent teachers can be here.
          //
          if (!isEvening) {
            continue;
          }
        }

        // ------------------------------------------------------
        // SUBJECT NORMALIZATION
        // ------------------------------------------------------

        const subjectUpper =
          subject
            .toString()
            .replace(/\u00A0/g, " ")
            .replace(/\s+/g, " ")
            .trim()
            .toUpperCase();

        Logger.log(
          "--------------------------------------"
        );

        Logger.log(
          "SCIENCE RECORD"
        );

        Logger.log(
          "Teacher = " +
          teacherName
        );

        Logger.log(
          "Class   = " +
          className
        );

        Logger.log(
          "Subject = " +
          subject
        );

        Logger.log(
          "Days    = " +
          days
        );

        Logger.log(
          "Count   = " +
          countDays(days)
        );

        // ------------------------------------------------------
        // ISL/MQ
        // ------------------------------------------------------

        if (
          subjectUpper === "ISL/MQ" ||
          subjectUpper === "ISL / MQ" ||
          subjectUpper === "ISL-MQ" ||
          subjectUpper === "ISL MQ"
        ) {

          addWorkload(
            map,
            teacherName,
            info,
            className + " (ISL)",
            days
          );

          addWorkload(
            map,
            teacherName,
            info,
            className + " (MQ)",
            days
          );

          continue;
        }

        // ------------------------------------------------------
        // MUTALIA QURAN
        // ------------------------------------------------------

        if (
          subjectUpper === "MUTALIA QURAN" ||
          subjectUpper === "MUTALIA-QURAN" ||
          subjectUpper === "MUTALIAQURAN"
        ) {

          addWorkload(
            map,
            teacherName,
            info,
            className + " (MQ)",
            days
          );

          continue;
        }

        // ------------------------------------------------------
        // ISLAMIYAT
        // ------------------------------------------------------

        if (
          subjectUpper === "ISLAMIYAT" ||
          subjectUpper === "ISLAMIAT"
        ) {

          addWorkload(
            map,
            teacherName,
            info,
            className + " (ISL)",
            days
          );

          continue;
        }

        // ------------------------------------------------------
        // ALL OTHER SCIENCE SUBJECTS
        // ------------------------------------------------------

        addWorkload(
          map,
          teacherName,
          info,
          className,
          days
        );
      }
    }

    return map;
  }

  // ============================================================
  // ARTS TIMETABLE
  // ============================================================

  function processArtsTimetable(
    targetStatus,
    mode
  ) {

    const map = {};

    const artsData =
      artsSheet
        .getDataRange()
        .getValues();

    // ----------------------------------------------------------
    // NORMALIZE GROUP
    // ----------------------------------------------------------

    function normalizeGroup(value) {

      let text =
        cleanText(value)
          .toUpperCase()
          .replace(/\s+/g, "");

      if (!text) {
        return "";
      }

      text =
        text
          .replace(
            /^F\.A-II/,
            "F.A-2"
          )
          .replace(
            /^F\.A-I(?!I)/,
            "F.A-1"
          );

      return text;
    }

    // ----------------------------------------------------------
    // VALID FA GROUP
    // ----------------------------------------------------------

    function isFAGroup(value) {

      const group =
        normalizeGroup(value);

      return /^F\.A-[12][A-D]$/.test(
        group
      );
    }

    // ----------------------------------------------------------
    // HEADERS
    // ----------------------------------------------------------

    function isFAOneHeader(text) {

      return /F\.?\s*A\s*-\s*I\s*TIME\s*TABLE/i
        .test(text);
    }

    function isFATwoHeader(text) {

      return /F\.?\s*A\s*-\s*II\s*TIME\s*TABLE/i
        .test(text);
    }

    // ----------------------------------------------------------
    // OPTIONAL SUBJECT SHORT NAMES
    // ----------------------------------------------------------

    function getOptionalSubjectShortName(
      subject
    ) {

      const s =
        cleanText(subject)
          .toUpperCase();

      if (!s) {
        return "";
      }

      if (
        s === "ENGLISH ADVANCE" ||
        s.includes("ENGLISH ADVANCE")
      ) {
        return "ENG ADV";
      }

      if (
        s === "URDU ADVANCE" ||
        s.includes("URDU ADVANCE")
      ) {
        return "URD ADV";
      }

      if (
        s === "ISLAMIYAT ELE" ||
        s.includes("ISLAMIYAT ELECTIVE")
      ) {
        return "ISL ELE";
      }

      if (
        s === "GENERAL MATHEMATICS" ||
        s.includes("GENERAL MATHEMATICS")
      ) {
        return "G.MATHS";
      }

      return "";
    }

    // ----------------------------------------------------------
    // CURRENT SECTION
    // ----------------------------------------------------------

    let currentSection =
      "F.A-1";

    // ----------------------------------------------------------
    // PROCESS
    // ----------------------------------------------------------

    for (
      let r = 0;
      r < artsData.length;
      r++
    ) {

      // --------------------------------------------------------
      // Detect FA section
      // --------------------------------------------------------

      for (
        let hc = 0;
        hc < artsData[r].length;
        hc++
      ) {

        const headerText =
          cleanText(
            artsData[r][hc]
          );

        if (!headerText) {
          continue;
        }

        if (
          isFATwoHeader(headerText)
        ) {

          currentSection =
            "F.A-2";

          break;
        }

        if (
          isFAOneHeader(headerText)
        ) {

          currentSection =
            "F.A-1";

          break;
        }
      }

      // --------------------------------------------------------
      // Process teacher cells
      // --------------------------------------------------------

      for (
        let c = 0;
        c < artsData[r].length;
        c++
      ) {

        const cell =
          cleanText(
            artsData[r][c]
          );

        if (!cell) {
          continue;
        }

        const teacherName =
          normalizeName(cell);

        const info =
          teacherInfoMap[
            teacherName
          ];

        if (!info) {
          continue;
        }

        // ------------------------------------------------------
        // SIX-CELL BLOCK
        // ------------------------------------------------------

        const subject =
          r - 1 >= 0
            ? cleanText(
                artsData[r - 1][c]
              )
            : "";

        const pcs =
          r - 1 >= 0 &&
          c + 1 <
            artsData[r - 1].length
            ? cleanText(
                artsData[r - 1][c + 1]
              )
            : "";

        const room =
          r + 1 <
            artsData.length
            ? cleanText(
                artsData[r + 1][c]
              )
            : "";

        const days =
          r + 1 <
            artsData.length &&
          c + 1 <
            artsData[r + 1].length
            ? cleanText(
                artsData[r + 1][c + 1]
              )
            : "";

        if (!subject) {
          continue;
        }

        if (!room || !days) {
          continue;
        }

        // ------------------------------------------------------
        // DETERMINE CLASS
        // ------------------------------------------------------

        let className = "";

        if (
          isFAGroup(pcs)
        ) {

          className =
            normalizeGroup(pcs);

        } else {

          const optionalSubject =
            getOptionalSubjectShortName(
              subject
            );

          if (optionalSubject) {

            className =
              currentSection +
              " " +
              optionalSubject;

          } else {

            className =
              currentSection;
          }
        }

        // ------------------------------------------------------
        // ARTS EVENING DETECTION
        // ------------------------------------------------------

        const isEvening =
          isEveningArtsAssignment(
            room,
            className,
            days
          );

        // ------------------------------------------------------
        // ASSIGNMENT FILTER
        // ------------------------------------------------------

        if (
          mode === "PERMANENT"
        ) {

          if (
            info.status !== "Permanent"
          ) {
            continue;
          }

          if (isEvening) {
            continue;
          }
        }

        if (
          mode === "MORNING_VISITING"
        ) {

          if (
            info.status !== "Visiting"
          ) {
            continue;
          }

          if (isEvening) {
            continue;
          }
        }

        if (
          mode === "EVENING_VISITING"
        ) {

          // Do not check teacher status.
          if (!isEvening) {
            continue;
          }
        }

        // ------------------------------------------------------
        // ADD WORKLOAD
        // ------------------------------------------------------

        addWorkload(
          map,
          teacherName,
          info,
          className,
          days
        );
      }
    }

    return map;
  }

  // ============================================================
  // MERGE SCIENCE + ARTS
  // ============================================================

  function mergeWorkloadMaps(
    destination,
    source
  ) {

    for (
      const dept of Object.keys(source)
    ) {

      if (!destination[dept]) {
        destination[dept] = {};
      }

      for (
        const teacherName of
        Object.keys(source[dept])
      ) {

        const sourceInfo =
          source[dept][teacherName];

        if (
          !destination[dept][teacherName]
        ) {

          destination[dept][teacherName] = {

            classes: [],

            workload: 0,

            designation:
              sourceInfo.designation,

            sn:
              sourceInfo.sn
          };
        }

        const destinationInfo =
          destination[dept][teacherName];

        for (
          const sourceClass of
          sourceInfo.classes
        ) {

          if (
            sourceClass &&
            typeof sourceClass === "object"
          ) {

            let existingClass =
              null;

            for (
              const destinationClass of
              destinationInfo.classes
            ) {

              if (
                destinationClass &&
                typeof destinationClass === "object" &&
                destinationClass.className ===
                  sourceClass.className
              ) {

                existingClass =
                  destinationClass;

                break;
              }
            }

            if (!existingClass) {

              destinationInfo.classes.push({

                className:
                  sourceClass.className,

                dayValues:
                  Array.isArray(
                    sourceClass.dayValues
                  )
                    ? [
                        ...sourceClass.dayValues
                      ]
                    : []
              });

            } else if (
              Array.isArray(
                sourceClass.dayValues
              )
            ) {

              existingClass.dayValues.push(
                ...sourceClass.dayValues
              );
            }

          } else if (
            typeof sourceClass === "string"
          ) {

            // Backward compatibility
            destinationInfo.classes.push({

              className:
                sourceClass,

              dayValues: []
            });
          }
        }

        destinationInfo.workload +=
          sourceInfo.workload;
      }
    }

    return destination;
  }

  // ============================================================
  // WRITE WORKLOAD SHEET
  //
  // THIS SECTION IS KEPT EXACTLY IN THE SAME LAYOUT
  // AS YOUR ORIGINAL CODE.
  // ============================================================

  function writeWorkloadSheet(
    sheetName,
    workloadMap
  ) {

    Logger.log(
      "Writing " + sheetName
    );

    const existing =
      ss.getSheetByName(sheetName);

    if (existing) {
      ss.deleteSheet(existing);
    }

    const dest =
      ss.insertSheet(sheetName);

    let row = 1;

    const year =
      new Date().getFullYear();

    // ==========================================================
    // TITLE
    // ==========================================================

    dest
      .getRange(row, 1, 1, 5)
      .merge()
      .setValue(
        sheetName.toUpperCase() +
        " – " +
        year
      )
      .setFontWeight("bold")
      .setFontSize(22)
      .setFontFamily("Georgia")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBackground("#B4C6E7");

    row += 2;

    // ==========================================================
    // DEPARTMENTS
    // ==========================================================

    const sortedDepartments =
      Object.keys(workloadMap).sort();

    for (
      const dept of sortedDepartments
    ) {

      // --------------------------------------------------------
      // Department heading
      // --------------------------------------------------------

      dest
        .getRange(row, 1, 1, 5)
        .merge()
        .setValue(
          dept.toString().toUpperCase()
        )
        .setFontWeight("bold")
        .setFontSize(16)
        .setFontFamily("Georgia")
        .setBackground("#D9E1F2")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");

      row++;

      // --------------------------------------------------------
      // HEADER
      // --------------------------------------------------------

      dest
        .getRange(row, 1, 1, 5)
        .setValues([
          [
            "S.No",
            "Teacher",
            "Designation",
            "Classes",
            "WORKLOAD"
          ]
        ])
        .setFontWeight("bold")
        .setFontSize(14)
        .setFontFamily("Calibri")
        .setBackground("#F2F2F2")
        .setHorizontalAlignment("center");

      row++;

      // --------------------------------------------------------
      // SORT TEACHERS BY S.NO
      // --------------------------------------------------------

      const teachers =
        workloadMap[dept];

      const sortedTeachers =
        Object.entries(
          teachers
        ).sort(
          (a, b) =>
            a[1].sn -
            b[1].sn
        );

      // IMPORTANT:
      // S.No starts from 1 separately for every department.
      let sno = 1;

      // --------------------------------------------------------
      // TEACHERS
      // --------------------------------------------------------

      for (
        const [teacher, info] of
        sortedTeachers
      ) {

        dest
          .getRange(
            row,
            1,
            1,
            5
          )
          .setValues([
            [
              sno,

              teacher,

              info.designation,

              info.classes
                .map(item => {

                  const formattedDays =
                    formatCombinedDays(
                      item.dayValues
                    );

                  return formattedDays
                    ? item.className +
                      " (" +
                      formattedDays +
                      ")"
                    : item.className;

                })
                .join(", "),

              info.workload
            ]
          ])
          .setFontSize(14)
          .setFontFamily("Cambria")
          .setVerticalAlignment("middle");

        // ------------------------------------------------------
        // BORDER
        // ------------------------------------------------------

        dest
          .getRange(
            row,
            1,
            1,
            5
          )
          .setBorder(
            true,
            true,
            true,
            true,
            true,
            true
          );

        // ------------------------------------------------------
        // WORKLOAD CENTERED
        // ------------------------------------------------------

        dest
          .getRange(row, 5)
          .setHorizontalAlignment(
            "center"
          )
          .setFontWeight("bold");

        row++;

        sno++;
      }

      // Space after department
      row++;
    }

    // ==========================================================
    // COLUMN WIDTHS
    // EXACTLY AS ORIGINAL
    // ==========================================================

    dest.setColumnWidth(1, 55);
    dest.setColumnWidth(2, 220);
    dest.setColumnWidth(3, 180);
    dest.setColumnWidth(4, 500);
    dest.setColumnWidth(5, 100);

    // ==========================================================
    // FREEZE
    // EXACTLY AS ORIGINAL
    // ==========================================================

    dest.setFrozenRows(2);
  }

  // ============================================================
  // 1. PERMANENT WORKLOAD
  // ============================================================
  //
  // Permanent teachers:
  //
  //   Morning assignment -> Permanent Workload
  //
  //   Evening assignment -> NOT HERE
  //
  // ============================================================

  Logger.log(
    "======================================"
  );

  Logger.log(
    "PERMANENT WORKLOAD"
  );

  Logger.log(
    "======================================"
  );

  const permanentScience =
    processScienceTimetable(
      "Permanent",
      "PERMANENT"
    );

  const permanentArts =
    processArtsTimetable(
      "Permanent",
      "PERMANENT"
    );

  const permanentCombined =
    mergeWorkloadMaps(
      permanentScience,
      permanentArts
    );

  writeWorkloadSheet(
    "Permanent Workload",
    permanentCombined
  );

  // ============================================================
  // 2. MORNING VISITING WORKLOAD
  // ============================================================
  //
  // Visiting teachers:
  //
  //   Morning assignment -> Morning Visiting Workload
  //
  //   Evening assignment -> NOT HERE
  //
  // ============================================================

  Logger.log(
    "======================================"
  );

  Logger.log(
    "MORNING VISITING WORKLOAD"
  );

  Logger.log(
    "======================================"
  );

  const morningVisitingScience =
    processScienceTimetable(
      "Visiting",
      "MORNING_VISITING"
    );

  const morningVisitingArts =
    processArtsTimetable(
      "Visiting",
      "MORNING_VISITING"
    );

  const morningVisitingCombined =
    mergeWorkloadMaps(
      morningVisitingScience,
      morningVisitingArts
    );

  writeWorkloadSheet(
    "Morning Visiting Workload",
    morningVisitingCombined
  );

  // ============================================================
  // 3. EVENING VISITING WORKLOAD
  // ============================================================
  //
  // VERY IMPORTANT:
  //
  // targetStatus = "ANY"
  //
  // because evening classes may be taught by:
  //
  //   - Visiting teachers
  //   - Permanent teachers working as Visiting Faculty
  //
  // Therefore STATUS IS NOT CHECKED for this sheet.
  //
  // Science evening sections:
  //
  //   L2
  //   M2
  //   N2
  //   O2
  //   P2
  //   S1
  //   T1
  //
  // ============================================================

  Logger.log(
    "======================================"
  );

  Logger.log(
    "EVENING VISITING WORKLOAD"
  );

  Logger.log(
    "======================================"
  );

  const eveningVisitingScience =
    processScienceTimetable(
      "ANY",
      "EVENING_VISITING"
    );

  const eveningVisitingArts =
    processArtsTimetable(
      "ANY",
      "EVENING_VISITING"
    );

  const eveningVisitingCombined =
    mergeWorkloadMaps(
      eveningVisitingScience,
      eveningVisitingArts
    );

  writeWorkloadSheet(
    "Evening Visiting Workload",
    eveningVisitingCombined
  );

  // ============================================================
  // FINISH
  // ============================================================

  SpreadsheetApp.flush();

  const endTime =
    new Date();

  Logger.log(
    "======================================"
  );

  Logger.log(
    "WORKLOAD GENERATION COMPLETED"
  );

  Logger.log(
    "Permanent teachers: " +
    Object.keys(
      permanentCombined
    ).length +
    " departments"
  );

  Logger.log(
    "Morning Visiting teachers: " +
    Object.keys(
      morningVisitingCombined
    ).length +
    " departments"
  );

  Logger.log(
    "Evening Visiting teachers: " +
    Object.keys(
      eveningVisitingCombined
    ).length +
    " departments"
  );

  Logger.log(
    "Execution time: " +
    (
      (
        endTime.getTime() -
        startTime.getTime()
      ) / 1000
    ).toFixed(2) +
    " seconds"
  );

  Logger.log(
    "======================================"
  );

  at_toast_(
    "Permanent + Morning Visiting + Evening Visiting Workload generated successfully.",
    "JCW Workload",
    8);
    return exportJCWWorkloadToExcel();
}

/* ====================================================================================================
   3. WORKLOAD -> PRINT-READY EXCEL IN DRIVE (original)
   ==================================================================================================== */

/**
 * ============================================================
 * JCW WORKLOAD → PRINT-READY EXCEL
 * ============================================================
 *
 * SOURCE SHEETS:
 *   1. Permanent Workload
 *   2. Morning Visiting Workload
 *   3. Evening Visiting Workload
 *
 * OUTPUT:
 *   WORKLOAD DD-MMM-YYYY.xlsx
 *
 * The Excel file is created as a separate file in Google Drive.
 * The original workbook is NOT modified.
 *
 * ============================================================
 */

function exportJCWWorkloadToExcel() {

  const startTime = new Date();

  const ss =
    at_ss_();

  const sourceSheets = [
    "Permanent Workload",
    "Morning Visiting Workload",
    "Evening Visiting Workload"
  ];

  // ============================================================
  // DATE
  // ============================================================

  const timeZone =
    ss.getSpreadsheetTimeZone() ||
    Session.getScriptTimeZone() ||
    "Asia/Karachi";

  const today =
    Utilities.formatDate(
      new Date(),
      timeZone,
      "dd-MMM-yyyy"
    );

  const workbookName =
    "WORKLOAD " + today;

  const excelFileName =
    workbookName + ".xlsx";

  Logger.log(
    "=========================================="
  );

  Logger.log(
    "JCW WORKLOAD EXCEL EXPORT"
  );

  Logger.log(
    "Workbook: " + workbookName
  );

  Logger.log(
    "=========================================="
  );

  // ============================================================
  // CHECK SOURCE SHEETS
  // ============================================================

  const missingSheets = [];

  sourceSheets.forEach(function(sheetName) {

    if (!ss.getSheetByName(sheetName)) {

      missingSheets.push(sheetName);

    }

  });

  if (missingSheets.length > 0) {

    throw new Error(
      "The following workload sheets are missing:\n\n" +
      missingSheets.join("\n")
    );
  }

  // ============================================================
  // REMOVE OLD EXPORT WITH SAME NAME
  // ============================================================
  //
  // This prevents Drive from accumulating multiple copies
  // when the function is run more than once on the same day.
  //
  // ============================================================

  const sourceFile =
    DriveApp.getFileById(
      ss.getId()
    );

  const parentFolders =
    sourceFile.getParents();

  let outputFolder = null;

  if (parentFolders.hasNext()) {

    outputFolder =
      parentFolders.next();

  } else {

    outputFolder =
      DriveApp.getRootFolder();

  }

  Logger.log(
    "Output folder: " +
    outputFolder.getName()
  );

  // Delete previous XLSX having exactly the same name.
  const oldFiles =
    outputFolder.getFilesByName(
      excelFileName
    );

  while (oldFiles.hasNext()) {

    const oldFile =
      oldFiles.next();

    Logger.log(
      "Deleting previous export: " +
      oldFile.getName()
    );

    oldFile.setTrashed(true);
  }

  // ============================================================
  // CREATE TEMPORARY GOOGLE SPREADSHEET
  // ============================================================

  Logger.log(
    "Creating temporary workbook..."
  );

  const tempSpreadsheet =
    SpreadsheetApp.create(
      workbookName
    );

  const tempId =
    tempSpreadsheet.getId();

  Logger.log(
    "Temporary Spreadsheet ID: " +
    tempId
  );

  // ============================================================
  // GET DEFAULT SHEET
  // ============================================================

  const defaultSheet =
    tempSpreadsheet.getSheets()[0];

  // ============================================================
  // COPY THREE WORKLOAD SHEETS
  // ============================================================

  const copiedSheets = [];

  sourceSheets.forEach(function(sheetName) {

    const sourceSheet =
      ss.getSheetByName(
        sheetName
      );

    Logger.log(
      "Copying sheet: " +
      sheetName
    );

    const copiedSheet =
      sourceSheet.copyTo(
        tempSpreadsheet
      );

    copiedSheet.setName(
      sheetName
    );

    copiedSheets.push(
      copiedSheet
    );

  });

  // ============================================================
  // REMOVE DEFAULT SHEET
  // ============================================================

  if (
    tempSpreadsheet.getSheets().length > 1
  ) {

    tempSpreadsheet.deleteSheet(
      defaultSheet
    );

  }

  // ============================================================
  // PRINT-READY FORMATTING
  // ============================================================
  //
  // We preserve your existing layout and additionally make
  // the workbook clean for Excel printing.
  //
  // ============================================================

  copiedSheets.forEach(function(sheet) {

    prepareWorkloadSheetForExcel(
      sheet
    );

  });

  // ============================================================
  // FLUSH
  // ============================================================

  SpreadsheetApp.flush();

  Utilities.sleep(2000);

  // ============================================================
  // EXPORT GOOGLE SPREADSHEET AS XLSX
  // ============================================================

  Logger.log(
    "Exporting XLSX..."
  );

  const exportUrl =
    "https://docs.google.com/spreadsheets/d/" +
    tempId +
    "/export?format=xlsx";

  const oauthToken =
    ScriptApp.getOAuthToken();

  const response =
    UrlFetchApp.fetch(
      exportUrl,
      {
        method: "get",

        headers: {
          Authorization:
            "Bearer " +
            oauthToken
        },

        muteHttpExceptions: true
      }
    );

  const responseCode =
    response.getResponseCode();

  Logger.log(
    "Export response code: " +
    responseCode
  );

  if (responseCode !== 200) {

    // Clean temporary spreadsheet
    try {

      DriveApp
        .getFileById(tempId)
        .setTrashed(true);

    } catch (cleanupError) {

      Logger.log(
        "Temporary file cleanup error: " +
        cleanupError
      );

    }

    throw new Error(
      "Excel export failed. HTTP response code: " +
      responseCode +
      "\n\n" +
      response
        .getContentText()
        .substring(0, 1000)
    );
  }

  // ============================================================
  // CREATE EXCEL BLOB
  // ============================================================

  const excelBlob =
    response
      .getBlob()
      .setName(
        excelFileName
      );

  // ============================================================
  // SAVE EXCEL FILE
  // ============================================================

  Logger.log(
    "Saving Excel file..."
  );

  const excelFile =
    outputFolder.createFile(
      excelBlob
    );

  // ============================================================
  // TEMPORARY GOOGLE SPREADSHEET
  // ============================================================
  //
  // The temporary Google Spreadsheet is no longer required.
  //
  // The exported XLSX remains safely in Drive.
  //
  // ============================================================

  try {

    DriveApp
      .getFileById(tempId)
      .setTrashed(true);

  } catch (cleanupError) {

    Logger.log(
      "Temporary workbook could not be trashed: " +
      cleanupError
    );

  }

  // ============================================================
  // FILE INFORMATION
  // ============================================================

  const fileId =
    excelFile.getId();

  const fileUrl =
    excelFile.getUrl();

  const downloadUrl =
    "https://drive.google.com/uc?export=download&id=" +
    fileId;

  // ============================================================
  // EXECUTION TIME
  // ============================================================

  const endTime =
    new Date();

  const seconds =
    (
      endTime.getTime() -
      startTime.getTime()
    ) / 1000;

  // ============================================================
  // LOG RESULT
  // ============================================================

  Logger.log(
    "=========================================="
  );

  Logger.log(
    "EXCEL EXPORT COMPLETED"
  );

  Logger.log(
    "File name: " +
    excelFile.getName()
  );

  Logger.log(
    "File ID: " +
    fileId
  );

  Logger.log(
    "File URL: " +
    fileUrl
  );

  Logger.log(
    "Download URL: " +
    downloadUrl
  );

  Logger.log(
    "Execution time: " +
    seconds.toFixed(2) +
    " seconds"
  );

  Logger.log(
    "=========================================="
  );

  // ============================================================
  // TOAST
  // ============================================================

  at_toast_(
    "Excel workload created successfully.\n\n" +
    excelFileName +
    "\n\n" +
    "Open the execution log to get the download link.",
    "JCW Workload Excel",
    10
  );

  return {
    fileName: excelFileName,
    fileId: fileId,
    fileUrl: fileUrl,
    downloadUrl: downloadUrl,
    executionTime: seconds
  };
}


/**
 * ============================================================
 * PREPARE WORKLOAD SHEET FOR EXCEL / PRINTING
 * ============================================================
 *
 * This does NOT change your existing source sheets.
 *
 * It only prepares the copied sheets inside the temporary
 * workbook before XLSX export.
 *
 * ============================================================
 */

function prepareWorkloadSheetForExcel(sheet) {

  const lastRow =
    sheet.getLastRow();

  const lastColumn =
    sheet.getLastColumn();

  if (
    lastRow === 0 ||
    lastColumn === 0
  ) {
    return;
  }

  // ============================================================
  // GENERAL SHEET SETTINGS
  // ============================================================

  sheet.setFrozenRows(2);

  // ============================================================
  // KEEP TEXT VERTICAL ALIGNMENT
  // ============================================================

  sheet
    .getRange(
      1,
      1,
      lastRow,
      lastColumn
    )
    .setVerticalAlignment(
      "middle"
    );

  // ============================================================
  // COLUMN WIDTHS
  // ============================================================
  //
  // Same widths as your workload generator.
  //
  // ============================================================

  sheet.setColumnWidth(1, 55);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 500);
  sheet.setColumnWidth(5, 100);

  // ============================================================
  // CENTER S.NO
  // ============================================================

  sheet
    .getRange(
      1,
      1,
      lastRow,
      1
    )
    .setHorizontalAlignment(
      "center"
    );

  // ============================================================
  // CENTER WORKLOAD
  // ============================================================

  sheet
    .getRange(
      1,
      5,
      lastRow,
      1
    )
    .setHorizontalAlignment(
      "center"
    );

  // ============================================================
  // WRAP CLASSES COLUMN
  // ============================================================

  sheet
    .getRange(
      1,
      4,
      lastRow,
      1
    )
    .setWrap(
      true
    );

  // ============================================================
  // TEACHER / DESIGNATION
  // ============================================================

  sheet
    .getRange(
      1,
      2,
      lastRow,
      2
    )
    .setWrap(
      true
    );

  // ============================================================
  // AUTO ROW HEIGHT
  // ============================================================

  sheet.autoResizeRows(
    1,
    lastRow
  );

  // ============================================================
  // MINIMUM ROW HEIGHT
  // ============================================================

  for (
    let r = 1;
    r <= lastRow;
    r++
  ) {

    const currentHeight =
      sheet.getRowHeight(r);

    if (currentHeight < 24) {

      sheet.setRowHeight(
        r,
        24
      );

    }
  }

  // ============================================================
  // TITLE ROW
  // ============================================================

  sheet
    .getRange(
      1,
      1,
      1,
      Math.min(5, lastColumn)
    )
    .setHorizontalAlignment(
      "center"
    )
    .setVerticalAlignment(
      "middle"
    );

  sheet.setRowHeight(
    1,
    38
  );

  // ============================================================
  // DETECT DEPARTMENT / HEADER ROWS
  // ============================================================
  //
  // This preserves the visual hierarchy when opened in Excel.
  //
  // ============================================================

  for (
    let r = 1;
    r <= lastRow;
    r++
  ) {

    const firstCell =
      cleanWorkloadExportText(
        sheet.getRange(r, 1).getDisplayValue()
      );

    const secondCell =
      cleanWorkloadExportText(
        sheet.getRange(r, 2).getDisplayValue()
      );

    const fifthCell =
      cleanWorkloadExportText(
        sheet.getRange(r, 5).getDisplayValue()
      );

    // ----------------------------------------------------------
    // HEADER ROW
    // ----------------------------------------------------------

    if (
      firstCell === "S.NO" &&
      secondCell === "TEACHER" &&
      fifthCell === "WORKLOAD"
    ) {

      sheet.setRowHeight(
        r,
        30
      );

      sheet
        .getRange(
          r,
          1,
          1,
          Math.min(5, lastColumn)
        )
        .setHorizontalAlignment(
          "center"
        )
        .setVerticalAlignment(
          "middle"
        );

      continue;
    }

    // ----------------------------------------------------------
    // DEPARTMENT ROW
    // ----------------------------------------------------------
    //
    // Department rows are merged A:E in your source.
    //
    // ----------------------------------------------------------

    const mergedRanges =
      sheet.getRange(
        r,
        1,
        1,
        Math.min(5, lastColumn)
      ).getMergedRanges();

    if (
      mergedRanges.length > 0
    ) {

      sheet.setRowHeight(
        r,
        30
      );

      sheet
        .getRange(
          r,
          1,
          1,
          Math.min(5, lastColumn)
        )
        .setHorizontalAlignment(
          "center"
        )
        .setVerticalAlignment(
          "middle"
        );
    }
  }

  // ============================================================
  // PRINT AREA SUPPORT
  // ============================================================
  //
  // Remove completely unused rows below the workload.
  // This helps Excel open with a clean printable worksheet.
  //
  // ============================================================

  const maxRows =
    sheet.getMaxRows();

  if (
    maxRows > lastRow
  ) {

    sheet.deleteRows(
      lastRow + 1,
      maxRows - lastRow
    );

  }

  // ============================================================
  // REMOVE UNUSED COLUMNS
  // ============================================================

  const maxColumns =
    sheet.getMaxColumns();

  if (
    maxColumns > lastColumn
  ) {

    sheet.deleteColumns(
      lastColumn + 1,
      maxColumns - lastColumn
    );
  }
}


/**
 * ============================================================
 * SMALL TEXT CLEANER
 * ============================================================
 */

function cleanWorkloadExportText(value) {

  if (
    value === null ||
    value === undefined
  ) {
    return "";
  }

  return value
    .toString()
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}


/**
 * ============================================================
 * OPTIONAL MENU
 * ============================================================
 *
 * If you already have an onOpen() function in your project,
 * DO NOT create another onOpen().
 *
 * Instead, add this item to your existing menu:
 *
 * .addItem(
 *   "📊 Export Workload to Excel",
 *   "exportJCWWorkloadToExcel"
 * )
 *
 * ============================================================
 */


/**
 * ============================================================
 * OPTIONAL STANDALONE MENU
 * ============================================================
 *
 * Use this ONLY if you do not already have an onOpen().
 *
 * ============================================================
 */

function onOpenWorkloadExportMenu() {

  SpreadsheetApp
    .getUi()
    .createMenu("📊 WORKLOAD EXPORT")
    .addItem(
      "📥 Download Workload Excel",
      "exportJCWWorkloadToExcel"
    )
    .addToUi();
}


/* ====================================================================================================
   4. STOPPING MERIT REPORT (original)
   ==================================================================================================== */


/**
 * =====================================================================
 * JINNAH COLLEGE FOR WOMEN
 * STOPPING MERIT REPORT
 * SESSION 2026-27
 * =====================================================================
 *
 * SOURCE:
 *   Portal backup
 *
 * IMPORTANT:
 *   STOPPING MERIT IS CALCULATED FROM:
 *
 *   Adjusted Obtained Marks / Original Total Marks
 *
 *   where:
 *
 *   Adjusted Obtained Marks =
 *       Original Obtained Marks + 20
 *       IF HafizQuran = Yes
 *
 * Example:
 *
 *   HafizQuran = Yes
 *   Obtained = 1080
 *   Total = 1100
 *
 *   Adjusted Obtained = 1080 + 20 = 1100
 *
 *   Merit % = 1100 / 1100 × 100 = 100%
 *
 * If HafizQuran = No:
 *
 *   1080 / 1100 × 100 = 98.18%
 *
 *
 * IMPORTANT:
 *   ONLY OPEN QUOTA STUDENTS ARE USED.
 *
 *   ONLY 1ST YEAR IS USED because stopping merit is an
 *   admission cutoff.
 *
 * COMMON 1200 SCALE:
 *
 *   Adjusted Obtained / 1200 =
 *       Merit Percentage / 100 × 1200
 *
 * =====================================================================
 */


function generateJCWStoppingMeritReport() {

  const ss = at_ss_();

  const SOURCE_SHEET_NAME = AT_CFG.STRENGTH;            // admin tools: Transform-ALL-COLLEGE-STRENGTH (was "Portal backup")
  const REPORT_SHEET_NAME = "Stopping Merit";

  const TARGET_CLASS = "1st Year";

  const STANDARD_TOTAL_MARKS = 1200;

  const HAFIZ_BONUS_MARKS = 20;


  // -------------------------------------------------------------------
  // SOURCE SHEET
  // -------------------------------------------------------------------

  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);

  if (!sourceSheet) {

    throw new Error(
      'Source sheet "' + SOURCE_SHEET_NAME + '" was not found.'
    );

  }


  // -------------------------------------------------------------------
  // READ DATA
  // -------------------------------------------------------------------

  const lastRow = sourceSheet.getLastRow();
  const lastColumn = sourceSheet.getLastColumn();

  if (lastRow < 2) {

    throw new Error(
      '"' + SOURCE_SHEET_NAME + '" does not contain student data.'
    );

  }


  const data = sourceSheet
    .getRange(1, 1, lastRow, lastColumn)
    .getValues();


  const headers = data[0];


  // -------------------------------------------------------------------
  // HEADER MAP
  // -------------------------------------------------------------------

  const col = {};


  col.enrollment = jcwSMFindColumn_(
    headers,
    [
      "Enrollment#",
      "Enrollment #",
      "Enroll#",
      "Enrollment No",
      "Enrollment No."
    ]
  );


  col.studentName = jcwSMFindColumn_(
    headers,
    [
      "Student Name",
      "StudentName"
    ]
  );


  col.fatherName = jcwSMFindColumn_(
    headers,
    [
      "Father Name",
      "FatherName"
    ]
  );


  col.program = jcwSMFindColumn_(
    headers,
    [
      "Program",
      "Program Name",
      "ProgramName"
    ]
  );


  col.className = jcwSMFindColumn_(
    headers,
    [
      "Class"
    ]
  );


  col.section = jcwSMFindColumn_(
    headers,
    [
      "Section"
    ]
  );


  col.category = jcwSMFindColumn_(
    headers,
    [
      "Category"
    ],
    false
  );


  col.obtained = jcwSMFindColumn_(
    headers,
    [
      "Obtain Marks/CGPA",
      "Obtained Marks/CGPA",
      "Obtain Marks",
      "Obtained Marks",
      "Marks Obtained",
      "Obtained"
    ]
  );


  col.total = jcwSMFindColumn_(
    headers,
    [
      "Total Marks/CGPA",
      "Total Marks",
      "Total Marks / CGPA",
      "Total"
    ]
  );


  col.division = jcwSMFindColumn_(
    headers,
    [
      "Division"
    ],
    false
  );


  col.percentage = jcwSMFindColumn_(
    headers,
    [
      "%age Marks",
      "% Age Marks",
      "Percentage Marks",
      "Percentage",
      "%age"
    ],
    false
  );


  // -------------------------------------------------------------------
  // HAFIZ QURAN COLUMN
  // -------------------------------------------------------------------
  //
  // This column is optional so that the script does not fail if it
  // is temporarily absent.
  //
  // If HafizQuran = YES:
  //     +20 marks
  //
  // Otherwise:
  //     +0 marks
  //
  // -------------------------------------------------------------------

  col.hafizQuran = jcwSMFindColumn_(
    headers,
    [
      "HafizQuran",
      "Hafiz Quran",
      "Hafiz Quran?",
      "Hafiz-e-Quran",
      "Hafiz E Quran",
      "HQ",
      "Hifz",
      "Hafiz"
    ],
    false
  );


  // -------------------------------------------------------------------
  // LOG HEADER INFORMATION
  // -------------------------------------------------------------------

  Logger.log("==================================================");
  Logger.log("JCW STOPPING MERIT REPORT");
  Logger.log("==================================================");

  Logger.log(
    "Enrollment# column: " +
    (col.enrollment + 1)
  );

  Logger.log(
    "Student Name column: " +
    (col.studentName + 1)
  );

  Logger.log(
    "Father Name column: " +
    (col.fatherName + 1)
  );

  Logger.log(
    "Program column: " +
    (col.program + 1)
  );

  Logger.log(
    "Class column: " +
    (col.className + 1)
  );

  Logger.log(
    "Section column: " +
    (col.section + 1)
  );

  Logger.log(
    "Obtain Marks/CGPA column: " +
    (col.obtained + 1)
  );

  Logger.log(
    "Total Marks/CGPA column: " +
    (col.total + 1)
  );

  Logger.log(
    "%age Marks column: " +
    (col.percentage + 1)
  );

  Logger.log(
    "HafizQuran column: " +
    (
      col.hafizQuran >= 0
        ? (col.hafizQuran + 1)
        : "NOT FOUND"
    )
  );


  // -------------------------------------------------------------------
  // GROUPS
  // -------------------------------------------------------------------

  const groups = {

    FA: {
      key: "FA",
      program: "F.A",
      category: "SUBSIDIZED (M)",
      sections: "All Sections",
      records: []
    },

    FCS: {
      key: "FCS",
      program: "FCS",
      category: "SUBSIDIZED (M)",
      sections: "All Sections",
      records: []
    },

    PRE_MED_SUB: {
      key: "PRE_MED_SUB",
      program: "Pre-Medical",
      category: "SUBSIDIZED",
      sections: "A-C",
      records: []
    },

    PRE_MED_M: {
      key: "PRE_MED_M",
      program: "Pre-Medical",
      category: "NON-SUBSIDIZED (M)",
      sections: "E-R",
      records: []
    },

    PRE_MED_E: {
      key: "PRE_MED_E",
      program: "Pre-Medical",
      category: "NON-SUBSIDIZED (E)",
      sections: "S-T",
      records: []
    },

    PRE_ENG: {
      key: "PRE_ENG",
      program: "Pre-Engineering",
      category: "SUBSIDIZED (M)",
      sections: "All Sections",
      records: []
    }

  };


  // -------------------------------------------------------------------
  // DUPLICATE ENROLLMENT PROTECTION
  // -------------------------------------------------------------------

  const seenEnrollmentNumbers = new Set();

  let duplicateEnrollmentCount = 0;
  let skippedRows = 0;
  let validRows = 0;


  // -------------------------------------------------------------------
  // PROCESS STUDENTS
  // -------------------------------------------------------------------

  for (let r = 1; r < data.length; r++) {

    const row = data[r];


    // ---------------------------------------------------------------
    // ENROLLMENT#
    // ---------------------------------------------------------------

    const enrollment =
      jcwSMClean_(row[col.enrollment]);

    if (!enrollment) {

      skippedRows++;
      continue;

    }


    // ---------------------------------------------------------------
    // DUPLICATE ENROLLMENT#
    // ---------------------------------------------------------------

    const enrollmentKey =
      enrollment.toUpperCase();

    if (
      seenEnrollmentNumbers.has(enrollmentKey)
    ) {

      duplicateEnrollmentCount++;

      Logger.log(
        "Duplicate Enrollment# skipped: " +
        enrollment
      );

      continue;

    }

    seenEnrollmentNumbers.add(
      enrollmentKey
    );


    // ---------------------------------------------------------------
    // CLASS
    // ---------------------------------------------------------------

    const className =
      jcwSMNormalizeClass_(
        row[col.className]
      );

    if (className !== TARGET_CLASS) {

      continue;

    }


    // ---------------------------------------------------------------
    // CATEGORY
    // ---------------------------------------------------------------

    const category =
      col.category >= 0
        ? jcwSMClean_(row[col.category])
        : "";


    // ---------------------------------------------------------------
    // ONLY OPEN QUOTA STUDENTS
    // ---------------------------------------------------------------

    if (
      !at_smIsOpen_(row, headers, enrollment)        // admin tools: open quota only (Portal backup + Transform quota)
    ) {

      continue;

    }


    // ---------------------------------------------------------------
    // BASIC INFORMATION
    // ---------------------------------------------------------------

    const program =
      jcwSMClean_(row[col.program]);

    const section =
      jcwSMClean_(row[col.section]);

    const studentName =
      jcwSMClean_(row[col.studentName]);

    const fatherName =
      jcwSMClean_(row[col.fatherName]);


    // ---------------------------------------------------------------
    // DETERMINE DISCIPLINE
    // ---------------------------------------------------------------

    const groupKey =
      jcwSMDetermineGroup_(
        program,
        section
      );

    if (!groupKey) {

      continue;

    }


    // ---------------------------------------------------------------
    // ORIGINAL MARKS
    // ---------------------------------------------------------------

    const obtainedRaw =
      row[col.obtained];

    const totalRaw =
      row[col.total];


    const obtained =
      jcwSMParseNumber_(
        obtainedRaw
      );

    const total =
      jcwSMParseNumber_(
        totalRaw
      );


    // ---------------------------------------------------------------
    // HAFIZ QURAN
    // ---------------------------------------------------------------

    const hafizQuran =
      col.hafizQuran >= 0
        ? jcwSMClean_(
            row[col.hafizQuran]
          )
        : "";


    const isHafizQuran =
      jcwSMNormalizeText_(
        hafizQuran
      ) === "YES" || /^(Y|TRUE|HAFIZ|HAFIZA)$/i.test(String(hafizQuran).trim());


    const hafizBonus =
      isHafizQuran
        ? HAFIZ_BONUS_MARKS
        : 0;


    // ---------------------------------------------------------------
    // ADJUSTED OBTAINED MARKS
    // ---------------------------------------------------------------
    //
    // Example:
    //
    // Hafiz = Yes
    // Obtained = 1080
    // Bonus = 20
    // Adjusted = 1100
    //
    // Hafiz = No
    // Obtained = 1080
    // Adjusted = 1080
    //
    // ---------------------------------------------------------------

    let adjustedObtained = null;

    if (obtained !== null) {

      adjustedObtained =
        obtained + hafizBonus;

    }


    // ---------------------------------------------------------------
    // PORTAL %AGE
    // ---------------------------------------------------------------

    const portalPercentage =
      jcwSMParsePercentage_(
        row[col.percentage]
      );


    // ---------------------------------------------------------------
    // CALCULATED ORIGINAL %AGE
    // ---------------------------------------------------------------

    let calculatedPercentage = null;

    if (
      obtained !== null &&
      total !== null &&
      total > 0
    ) {

      calculatedPercentage =
        (obtained / total) * 100;

    }


    // ---------------------------------------------------------------
    // ADJUSTED %AGE
    // ---------------------------------------------------------------
    //
    // THIS IS NOW THE PRIMARY MERIT VALUE.
    //
    // Hafiz students receive +20 before percentage calculation.
    //
    // ---------------------------------------------------------------

    let adjustedPercentage = null;

    if (
      adjustedObtained !== null &&
      total !== null &&
      total > 0
    ) {

      adjustedPercentage =
        (adjustedObtained / total) * 100;

    }


    // ---------------------------------------------------------------
    // MERIT PERCENTAGE
    // ---------------------------------------------------------------
    //
    // Priority:
    //
    // 1. Adjusted marks / total marks
    // 2. Portal %age Marks as fallback
    //
    // This means HafizQuran bonus is actually reflected in merit.
    //
    // ---------------------------------------------------------------

    let meritPercentage = null;

    if (
      adjustedPercentage !== null
    ) {

      meritPercentage =
        adjustedPercentage;

    } else if (
      portalPercentage !== null
    ) {

      meritPercentage =
        portalPercentage;

    }


    // ---------------------------------------------------------------
    // NO USABLE MERIT PERCENTAGE
    // ---------------------------------------------------------------

    if (meritPercentage === null) {

      Logger.log(
        "Skipped - no valid percentage: " +
        enrollment
      );

      skippedRows++;
      continue;

    }


    // ---------------------------------------------------------------
    // COMMON 1200 SCALE
    // ---------------------------------------------------------------

    const adjustedObtained1200 =
      (meritPercentage / 100) *
      STANDARD_TOTAL_MARKS;


    // ---------------------------------------------------------------
    // DIVISION
    // ---------------------------------------------------------------

    const division =
      col.division >= 0
        ? jcwSMClean_(
            row[col.division]
          )
        : "";


    // ---------------------------------------------------------------
    // SAVE RECORD
    // ---------------------------------------------------------------

    groups[groupKey].records.push({

      enrollment: enrollment,

      studentName: studentName,

      fatherName: fatherName,

      program: program,

      className: className,

      section: section,

      category: category,

      hafizQuran: hafizQuran,

      hafizBonus: hafizBonus,

      obtained: obtainedRaw,

      total: totalRaw,

      obtainedNumeric: obtained,

      totalNumeric: total,

      adjustedObtained: adjustedObtained,

      division: division,

      portalPercentage: portalPercentage,

      calculatedPercentage: calculatedPercentage,

      adjustedPercentage: adjustedPercentage,

      meritPercentage: meritPercentage,

      adjustedObtained1200:
        adjustedObtained1200

    });


    validRows++;

  }


  // -------------------------------------------------------------------
  // CREATE / CLEAR REPORT SHEET
  // -------------------------------------------------------------------

  let reportSheet =
    ss.getSheetByName(
      REPORT_SHEET_NAME
    );


  if (!reportSheet) {

    reportSheet =
      ss.insertSheet(
        REPORT_SHEET_NAME
      );

  } else {

    reportSheet.clear();

    reportSheet.clearFormats();

    reportSheet
      .getRange(
        1,
        1,
        Math.max(
          reportSheet.getMaxRows(),
          1
        ),
        Math.max(
          reportSheet.getMaxColumns(),
          1
        )
      )
      .breakApart();

  }


  // -------------------------------------------------------------------
  // REPORT TITLE
  // -------------------------------------------------------------------

  reportSheet
    .getRange(1, 1, 1, 16)
    .merge();

  reportSheet
    .getRange(1, 1)
    .setValue(
      "JINNAH COLLEGE FOR WOMEN - STOPPING MERIT REPORT"
    )
    .setFontSize(16)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");


  // -------------------------------------------------------------------
  // REPORT SUBTITLE
  // -------------------------------------------------------------------

  reportSheet
    .getRange(2, 1, 1, 16)
    .merge();

  reportSheet
    .getRange(2, 1)
    .setValue(
      "Session 2026-27 | Class: 1st Year | " +
      "Merit Basis: Adjusted Obtained Marks | " +
      "HafizQuran Bonus: +20 | " +
      "Standardized Scale: 1200 Marks"
    )
    .setFontSize(10)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");


  // -------------------------------------------------------------------
  // HEADERS
  // -------------------------------------------------------------------

  const headersOutput = [

    "PROGRAM",

    "CATEGORY / DISCIPLINE",

    "SECTION(S)",

    "STUDENTS CONSIDERED",

    "STOPPING MERIT %",

    "STOPPING MERIT / 1200",

    "ORIGINAL OBTAINED",

    "HAFIZ BONUS",

    "ADJUSTED OBTAINED",

    "ORIGINAL TOTAL",

    "DIVISION",

    "ENROLLMENT#",

    "STUDENT NAME",

    "FATHER NAME",

    "PORTAL %age Marks",

    "CALCULATED / ADJUSTED %age"

  ];


  reportSheet
    .getRange(
      4,
      1,
      1,
      headersOutput.length
    )
    .setValues(
      [headersOutput]
    )
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);


  // -------------------------------------------------------------------
  // GROUP ORDER
  // -------------------------------------------------------------------

  const groupOrder = [

    "FA",
    "FCS",
    "PRE_MED_SUB",
    "PRE_MED_M",
    "PRE_MED_E",
    "PRE_ENG"

  ];


  // -------------------------------------------------------------------
  // BUILD SUMMARY
  // -------------------------------------------------------------------

  const output = [];


  groupOrder.forEach(
    function(groupKey) {

      const group =
        groups[groupKey];


      if (!group.records.length) {

        output.push([

          group.program,

          group.category,

          group.sections,

          0,

          "N/A",

          "N/A",

          "",

          "",

          "",

          "",

          "",

          "",

          "No valid student record",

          "",

          "",

          ""

        ]);

        return;

      }


      // -------------------------------------------------------------
      // SORT BY MERIT PERCENTAGE ASCENDING
      // LOWEST VALID %AGE = STOPPING MERIT
      // -------------------------------------------------------------

      group.records.sort(
        function(a, b) {

          return (
            a.meritPercentage -
            b.meritPercentage
          );

        }
      );


      const stoppingPercentage =
        group.records[0]
          .meritPercentage;


      // -------------------------------------------------------------
      // FIND ALL STUDENTS AT EXACT STOPPING MERIT
      // -------------------------------------------------------------

      const stoppingStudents =
        group.records.filter(
          function(record) {

            return Math.abs(
              record.meritPercentage -
              stoppingPercentage
            ) < 0.000001;

          }
        );


      // -------------------------------------------------------------
      // JOIN MULTIPLE TIED STUDENTS
      // -------------------------------------------------------------

      const stoppingEnrollment =
        stoppingStudents
          .map(
            function(record) {

              return record.enrollment;

            }
          )
          .join(", ");


      const stoppingNames =
        stoppingStudents
          .map(
            function(record) {

              return record.studentName;

            }
          )
          .join(", ");


      const stoppingFatherNames =
        stoppingStudents
          .map(
            function(record) {

              return record.fatherName;

            }
          )
          .join(", ");


      const stoppingOriginalObtained =
        stoppingStudents
          .map(
            function(record) {

              return jcwSMFormatOriginalValue_(
                record.obtained
              );

            }
          )
          .join(", ");


      const stoppingHafizBonus =
        stoppingStudents
          .map(
            function(record) {

              return record.hafizBonus;

            }
          )
          .join(", ");


      const stoppingAdjustedObtained =
        stoppingStudents
          .map(
            function(record) {

              return jcwSMFormatNumber_(
                record.adjustedObtained
              );

            }
          )
          .join(", ");


      const stoppingOriginalTotal =
        stoppingStudents
          .map(
            function(record) {

              return jcwSMFormatOriginalValue_(
                record.total
              );

            }
          )
          .join(", ");


      const stoppingDivisions =
        stoppingStudents
          .map(
            function(record) {

              return record.division;

            }
          )
          .join(", ");


      const stoppingPortalPercentages =
        stoppingStudents
          .map(
            function(record) {

              if (
                record.portalPercentage !== null
              ) {

                return jcwSMFormatPercent_(
                  record.portalPercentage
                );

              }

              return "";

            }
          )
          .join(", ");


      const stoppingAdjustedPercentages =
        stoppingStudents
          .map(
            function(record) {

              if (
                record.adjustedPercentage !== null
              ) {

                return jcwSMFormatPercent_(
                  record.adjustedPercentage
                );

              }

              return "";

            }
          )
          .join(", ");


      // -------------------------------------------------------------
      // COMMON 1200 SCALE
      // -------------------------------------------------------------

      const stopping1200 =
        (
          stoppingPercentage / 100
        ) *
        STANDARD_TOTAL_MARKS;


      // -------------------------------------------------------------
      // SUMMARY ROW
      // -------------------------------------------------------------

      // v7.6.1: one row per student at the stopping merit (ties no longer mixed in one cell);
      // program / category / sections / merit are merged over her rows
      SM_MERGES_.push([output.length, stoppingStudents.length]);
      stoppingStudents.forEach(function (record, k) {
        output.push([
          k ? '' : group.program,
          k ? '' : group.category,
          k ? '' : group.sections,
          k ? '' : group.records.length,
          k ? '' : jcwSMFormatPercent_(stoppingPercentage),
          k ? '' : jcwSMFormatNumber_(stopping1200) + " / 1200",
          jcwSMFormatOriginalValue_(record.obtained),
          record.hafizBonus,
          jcwSMFormatNumber_(record.adjustedObtained),
          jcwSMFormatOriginalValue_(record.total),
          record.division,
          record.enrollment,
          record.studentName,
          record.fatherName,
          record.portalPercentage !== null ? jcwSMFormatPercent_(record.portalPercentage) : "",
          record.adjustedPercentage !== null ? jcwSMFormatPercent_(record.adjustedPercentage) : ""
        ]);
      });

    }
  );


  // -------------------------------------------------------------------
  // WRITE OUTPUT
  // -------------------------------------------------------------------

  if (output.length) {

    reportSheet
      .getRange(
        5,
        1,
        output.length,
        headersOutput.length
      )
      .setValues(output);

  }
  SM_MERGES_.forEach(function (mg) { if (mg[1] > 1) for (var c = 1; c <= 6; c++) reportSheet.getRange(5 + mg[0], c, mg[1], 1).merge(); });
  SM_MERGES_ = [];


  // -------------------------------------------------------------------
  // FORMATTING
  // -------------------------------------------------------------------

  reportSheet
    .getRange(
      4,
      1,
      Math.max(
        output.length + 1,
        1
      ),
      headersOutput.length
    )
    .setVerticalAlignment("middle")
    .setWrap(true);


  if (output.length > 0) {

    reportSheet
      .getRange(
        5,
        4,
        output.length,
        1
      )
      .setHorizontalAlignment(
        "center"
      );


    // STOPPING %
    reportSheet
      .getRange(
        5,
        5,
        output.length,
        2
      )
      .setHorizontalAlignment(
        "center"
      )
      .setFontWeight("bold")
      .setFontSize(11);


    // MARKS / BONUS
    reportSheet
      .getRange(
        5,
        7,
        output.length,
        4
      )
      .setHorizontalAlignment(
        "center"
      );


    // OTHER DATA
    reportSheet
      .getRange(
        5,
        11,
        output.length,
        6
      )
      .setHorizontalAlignment(
        "center"
      );

  }


  // -------------------------------------------------------------------
  // BORDERS
  // -------------------------------------------------------------------

  reportSheet
    .getRange(
      4,
      1,
      output.length + 1,
      headersOutput.length
    )
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true
    );


  // -------------------------------------------------------------------
  // COLUMN WIDTHS
  // -------------------------------------------------------------------

  const widths = [

    150, // PROGRAM
    230, // CATEGORY
    90,  // SECTION
    120, // STUDENTS
    120, // STOPPING %
    145, // 1200
    120, // ORIGINAL OBTAINED
    100, // HAFIZ BONUS
    130, // ADJUSTED OBTAINED
    120, // ORIGINAL TOTAL
    90,  // DIVISION
    150, // ENROLLMENT
    180, // STUDENT
    180, // FATHER
    120, // PORTAL %
    145  // ADJUSTED %

  ];


  widths.forEach(
    function(width, index) {

      reportSheet.setColumnWidth(
        index + 1,
        width
      );

    }
  );


  // -------------------------------------------------------------------
  // ROW HEIGHTS
  // -------------------------------------------------------------------

  reportSheet.setRowHeight(
    1,
    30
  );

  reportSheet.setRowHeight(
    2,
    28
  );

  reportSheet.setRowHeight(
    4,
    48
  );


  // -------------------------------------------------------------------
  // FREEZE
  // -------------------------------------------------------------------

  reportSheet.setFrozenRows(4);


  // -------------------------------------------------------------------
  // VIEW
  // -------------------------------------------------------------------

  reportSheet.setHiddenGridlines(true);


  // -------------------------------------------------------------------
  // FILTER
  // -------------------------------------------------------------------

  if (reportSheet.getFilter()) {

    reportSheet
      .getFilter()
      .remove();

  }


  reportSheet
    .getRange(
      4,
      1,
      output.length + 1,
      headersOutput.length
    )
    .createFilter();

  reportSheet.deleteColumn(11);                        // admin tools: no DIVISION column
  at_smNote_(reportSheet);                             // admin tools: quotas found / used


  // -------------------------------------------------------------------
  // SUMMARY LOG
  // -------------------------------------------------------------------

  Logger.log(
    "=================================================="
  );

  Logger.log(
    "STOPPING MERIT REPORT COMPLETED"
  );

  Logger.log(
    "=================================================="
  );

  Logger.log(
    "Valid student records used: " +
    validRows
  );

  Logger.log(
    "Skipped rows: " +
    skippedRows
  );

  Logger.log(
    "Duplicate Enrollment# skipped: " +
    duplicateEnrollmentCount
  );


  groupOrder.forEach(
    function(groupKey) {

      const group =
        groups[groupKey];


      if (!group.records.length) {

        Logger.log(
          group.category +
          ": NO VALID RECORDS"
        );

        return;

      }


      group.records.sort(
        function(a, b) {

          return (
            a.meritPercentage -
            b.meritPercentage
          );

        }
      );


      const stopping =
        group.records[0]
          .meritPercentage;


      const stopping1200 =
        (
          stopping / 100
        ) *
        STANDARD_TOTAL_MARKS;


      Logger.log(

        group.category +
        " | Students = " +
        group.records.length +
        " | Stopping Merit = " +
        jcwSMFormatPercent_(
          stopping
        ) +
        " | 1200 Scale = " +
        jcwSMFormatNumber_(
          stopping1200
        ) +
        "/1200"

      );

    }
  );


  Logger.log(
    "=================================================="
  );

}


/**
 * =====================================================================
 * DETERMINE DISCIPLINE / CATEGORY
 * =====================================================================
 */
function jcwSMDetermineGroup_(
  program,
  section
) {

  const p =
    jcwSMNormalizeText_(
      program
    );

  const s =
    jcwSMNormalizeSection_(
      section
    );


  // -------------------------------------------------------------------
  // F.A
  // -------------------------------------------------------------------

  if (

    p === "FA" ||
    p === "F.A" ||
    p.indexOf("F.A") === 0 ||
    p.indexOf("FA ") === 0

  ) {

    return "FA";

  }


  // -------------------------------------------------------------------
  // FCS
  // -------------------------------------------------------------------

  if (

    p === "FCS" ||
    p.indexOf("FCS ") === 0 ||
    p.indexOf("FCS-") === 0

  ) {

    return "FCS";

  }


  // -------------------------------------------------------------------
  // PRE-MEDICAL
  // -------------------------------------------------------------------

  if (

    p.indexOf("PRE-MEDICAL") === 0 ||
    p.indexOf("PRE MEDICAL") === 0

  ) {


    // ---------------------------------------------------------------
    // SUBSIDIZED
    // A, B, C
    // ---------------------------------------------------------------

    if (
      ["A", "B", "C"].indexOf(s) !== -1
    ) {

      return "PRE_MED_SUB";

    }


    // ---------------------------------------------------------------
    // NON-SUBSIDIZED MORNING
    // E-R
    // ---------------------------------------------------------------

    const morningSections = [

      "E",
      "F",
      "G",
      "H",
      "I",
      "J",
      "K",
      "L",
      "M",
      "N",
      "O",
      "P",
      "Q",
      "R"

    ];


    if (
      morningSections.indexOf(s) !== -1
    ) {

      return "PRE_MED_M";

    }


    // ---------------------------------------------------------------
    // NON-SUBSIDIZED EVENING
    // S-T
    // ---------------------------------------------------------------

    if (
      ["S", "T"].indexOf(s) !== -1
    ) {

      return "PRE_MED_E";

    }


    return null;

  }


  // -------------------------------------------------------------------
  // PRE-ENGINEERING
  // -------------------------------------------------------------------

  if (

    p.indexOf("PRE-ENGINEERING") === 0 ||
    p.indexOf("PRE ENGINEERING") === 0

  ) {

    return "PRE_ENG";

  }


  return null;

}


/**
 * =====================================================================
 * FIND COLUMN
 * =====================================================================
 */
function jcwSMFindColumn_(
  headers,
  possibleNames,
  required
) {

  if (required === undefined) {

    required = true;

  }


  const normalizedHeaders =
    headers.map(
      function(header) {

        return jcwSMNormalizeHeader_(
          header
        );

      }
    );


  for (
    let i = 0;
    i < possibleNames.length;
    i++
  ) {

    const wanted =
      jcwSMNormalizeHeader_(
        possibleNames[i]
      );


    const index =
      normalizedHeaders.indexOf(
        wanted
      );


    if (index !== -1) {

      return index;

    }

  }


  if (required) {

    throw new Error(

      "Required column not found. " +
      "Possible names: " +
      possibleNames.join(", ")

    );

  }


  return -1;

}


/**
 * =====================================================================
 * NORMALIZE HEADER
 * =====================================================================
 */
function jcwSMNormalizeHeader_(
  value
) {

  return String(

    value === null ||
    value === undefined
      ? ""
      : value

  )
    .trim()
    .toLowerCase()
    .replace(
      /\s+/g,
      " "
    );

}


/**
 * =====================================================================
 * NORMALIZE TEXT
 * =====================================================================
 */
function jcwSMNormalizeText_(
  value
) {

  return String(

    value === null ||
    value === undefined
      ? ""
      : value

  )
    .trim()
    .toUpperCase()
    .replace(
      /\s+/g,
      " "
    );

}


/**
 * =====================================================================
 * CLEAN TEXT
 * =====================================================================
 */
function jcwSMClean_(
  value
) {

  if (

    value === null ||
    value === undefined

  ) {

    return "";

  }


  return String(value)
    .trim()
    .replace(
      /\s+/g,
      " "
    );

}


/**
 * =====================================================================
 * NORMALIZE CLASS
 * =====================================================================
 */
function jcwSMNormalizeClass_(
  value
) {

  const text =
    jcwSMNormalizeText_(
      value
    );


  if (

    text === "1ST YEAR" ||
    text === "1ST" ||
    text === "FIRST YEAR" ||
    text === "FIRST"

  ) {

    return "1st Year";

  }


  if (

    text === "2ND YEAR" ||
    text === "2ND" ||
    text === "SECOND YEAR" ||
    text === "SECOND"

  ) {

    return "2nd Year";

  }


  return jcwSMClean_(
    value
  );

}


/**
 * =====================================================================
 * NORMALIZE SECTION
 * =====================================================================
 *
 * Supports:
 *
 *   A
 *   B
 *   A1
 *   A2
 *
 * A1 / A2 become A.
 *
 * =====================================================================
 */
function jcwSMNormalizeSection_(
  value
) {

  let section =
    jcwSMNormalizeText_(
      value
    );


  section =
    section.replace(
      /\s+/g,
      ""
    );


  if (
    /^[A-Z][12]$/.test(section)
  ) {

    section =
      section.substring(
        0,
        section.length - 1
      );

  }


  return section;

}


/**
 * =====================================================================
 * PARSE NUMBER
 * =====================================================================
 */
function jcwSMParseNumber_(
  value
) {

  if (

    value === null ||
    value === undefined ||
    value === ""

  ) {

    return null;

  }


  if (
    typeof value === "number"
  ) {

    return isFinite(value)
      ? value
      : null;

  }


  let text =
    String(value)
      .trim()
      .replace(/,/g, "")
      .replace(/%/g, "")
      .trim();


  if (!text) {

    return null;

  }


  const number =
    Number(text);


  if (!isFinite(number)) {

    return null;

  }


  return number;

}


/**
 * =====================================================================
 * PARSE PERCENTAGE
 * =====================================================================
 */
function jcwSMParsePercentage_(
  value
) {

  if (

    value === null ||
    value === undefined ||
    value === ""

  ) {

    return null;

  }


  // Numeric cell
  if (
    typeof value === "number"
  ) {

    if (!isFinite(value)) {

      return null;

    }


    if (
      value >= 0 &&
      value <= 1
    ) {

      return value * 100;

    }


    return value;

  }


  let text =
    String(value)
      .trim()
      .replace(/,/g, "")
      .replace(/%/g, "")
      .trim();


  if (!text) {

    return null;

  }


  const number =
    Number(text);


  if (!isFinite(number)) {

    return null;

  }


  if (
    number >= 0 &&
    number <= 1
  ) {

    return number * 100;

  }


  return number;

}


/**
 * =====================================================================
 * FORMAT PERCENTAGE
 * =====================================================================
 */
function jcwSMFormatPercent_(
  value
) {

  if (

    value === null ||
    value === undefined ||
    !isFinite(value)

  ) {

    return "";

  }


  return Number(value)
    .toFixed(2) + "%";

}


/**
 * =====================================================================
 * FORMAT NUMBER
 * =====================================================================
 */
function jcwSMFormatNumber_(
  value
) {

  if (

    value === null ||
    value === undefined ||
    !isFinite(value)

  ) {

    return "";

  }


  return Number(value)
    .toFixed(2);

}


/**
 * =====================================================================
 * FORMAT ORIGINAL VALUE
 * =====================================================================
 */
function jcwSMFormatOriginalValue_(
  value
) {

  if (

    value === null ||
    value === undefined

  ) {

    return "";

  }


  if (
    typeof value === "number"
  ) {

    if (
      Number.isInteger(value)
    ) {

      return String(value);

    }


    return Number(value)
      .toFixed(2);

  }


  return String(value);

}


/* ====================================================================================================
   7. FILL ROOM FROM TIMETABLE INTO ALL DESTINATION SHEETS (original)
   ==================================================================================================== */

/**
 * ============================================================
 * FILL ROOM FROM TIMETABLE INTO ALL DESTINATION SHEETS
 * ============================================================
 *
 * TIMETABLE WORKBOOK:
 * 1ARU8zMO2daJUfU_SIIo4baeSOG7qBPuni5TYRIRLBNY
 * Sheet: Timetable
 *
 * DESTINATION WORKBOOK:
 * 1cz6Wf0BUEnYRw89uhIRTxy6tYQiodp_p
 *
 * MATCHING:
 *
 * Timetable:
 *   Program + Section
 *
 * is matched with destination:
 *   ProgramName + Section
 *
 * Program names are normalized automatically.
 *
 * Example:
 *   Timetable:   F.A (JCW)
 *   Destination: F.A (Jinnah College)
 *
 * Both become:
 *   F.A
 *
 * Therefore they match.
 *
 * ============================================================
 */





function fillRoomFromTimetableIntoAllSheets() {

  // ==========================================================
  // WORKBOOKS
  // ==========================================================

  const TIMETABLE_SPREADSHEET_ID =
    '1ARU8zMO2daJUfU_SIIo4baeSOG7qBPuni5TYRIRLBNY';

  const DESTINATION_SPREADSHEET_ID =
    '1cz6Wf0BUEnYRw89uhIRTxy6tYQiodp_p';


  // ==========================================================
  // OPEN TIMETABLE
  // ==========================================================

  const timetableSS =
    SpreadsheetApp.openById(TIMETABLE_SPREADSHEET_ID);

  const timetableSheet =
    timetableSS.getSheetByName('Timetable');

  if (!timetableSheet) {
    throw new Error(
      "Sheet 'Timetable' was not found in the timetable spreadsheet."
    );
  }


  // ==========================================================
  // READ TIMETABLE
  // ==========================================================

  const timetableData =
    timetableSheet.getDataRange().getValues();

  if (timetableData.length < 2) {
    throw new Error("Timetable sheet contains no data.");
  }

  const timetableHeaders = timetableData[0];

  const timetableHeaderMap =
    createHeaderMapRoom_(timetableHeaders);


  // Required Timetable columns
  const timetableProgramCol =
    timetableHeaderMap['program'];

  const timetableSectionCol =
    timetableHeaderMap['section'];

  const timetableRoomCol =
    timetableHeaderMap['room'];


  if (timetableProgramCol === undefined) {
    throw new Error(
      "Required Timetable column 'Program' was not found."
    );
  }

  if (timetableSectionCol === undefined) {
    throw new Error(
      "Required Timetable column 'Section' was not found."
    );
  }

  if (timetableRoomCol === undefined) {
    throw new Error(
      "Required Timetable column 'Room' was not found."
    );
  }


  // ==========================================================
  // CREATE PROGRAM + SECTION → ROOM LOOKUP
  // ==========================================================

  const roomMap = {};

  // Keep track of duplicate program/section combinations
  const duplicateRooms = {};

  for (let r = 1; r < timetableData.length; r++) {

    const row = timetableData[r];

    const program =
      row[timetableProgramCol];

    const section =
      row[timetableSectionCol];

    const room =
      row[timetableRoomCol];


    // Ignore incomplete rows
    if (
      program === '' ||
      program === null ||
      section === '' ||
      section === null ||
      room === '' ||
      room === null
    ) {
      continue;
    }


    // Normalize program name
    const normalizedProgram =
      normalizeProgramForRoom_(program);

    const normalizedSection =
      normalizeSectionForRoom_(section);


    if (!normalizedProgram || !normalizedSection) {
      continue;
    }


    // Final lookup key
    const key =
      normalizedProgram + '||' + normalizedSection;


    // ----------------------------------------------------------
    // FIRST ROOM
    // ----------------------------------------------------------

    if (!roomMap[key]) {

      roomMap[key] =
        String(room).trim();

    }

    // ----------------------------------------------------------
    // DUPLICATE ROOM
    // ----------------------------------------------------------

    else {

      const existingRoom =
        String(roomMap[key]).trim();

      const newRoom =
        String(room).trim();


      if (
        newRoom &&
        existingRoom &&
        newRoom !== existingRoom
      ) {

        if (!duplicateRooms[key]) {
          duplicateRooms[key] = [existingRoom];
        }

        if (
          duplicateRooms[key].indexOf(newRoom) === -1
        ) {
          duplicateRooms[key].push(newRoom);
        }
      }
    }
  }


  // ==========================================================
  // OPEN DESTINATION WORKBOOK
  // ==========================================================

  const destinationSS =
    SpreadsheetApp.openById(DESTINATION_SPREADSHEET_ID);

  const destinationSheets =
    destinationSS.getSheets();


  let totalSheetsProcessed = 0;
  let totalRowsUpdated = 0;
  let skippedSheets = [];


  // ==========================================================
  // PROCESS EVERY DESTINATION SHEET
  // ==========================================================

  destinationSheets.forEach(function(sheet) {

    const lastRow =
      sheet.getLastRow();

    const lastColumn =
      sheet.getLastColumn();


    // Skip completely empty sheets
    if (
      lastRow < 2 ||
      lastColumn < 1
    ) {
      return;
    }


    // ========================================================
    // READ HEADERS
    // ========================================================

    const headers =
      sheet
        .getRange(1, 1, 1, lastColumn)
        .getValues()[0];


    const headerMap =
      createHeaderMapRoom_(headers);


    // ========================================================
    // REQUIRED DESTINATION COLUMNS
    // ========================================================

    const programNameCol =
      headerMap['programname'];

    const sectionCol =
      headerMap['section'];

    const roomCol =
      headerMap['room'];


    // --------------------------------------------------------
    // CHECK REQUIRED COLUMNS
    // --------------------------------------------------------

    if (programNameCol === undefined) {

      skippedSheets.push(
        sheet.getName() +
        " → ProgramName column not found"
      );

      return;
    }


    if (sectionCol === undefined) {

      skippedSheets.push(
        sheet.getName() +
        " → Section column not found"
      );

      return;
    }


    if (roomCol === undefined) {

      skippedSheets.push(
        sheet.getName() +
        " → Room column not found"
      );

      return;
    }


    // ========================================================
    // READ ONLY NECESSARY DATA
    // ========================================================

    const numRows =
      lastRow - 1;

    const programValues =
      sheet
        .getRange(2, programNameCol + 1, numRows, 1)
        .getValues();

    const sectionValues =
      sheet
        .getRange(2, sectionCol + 1, numRows, 1)
        .getValues();

    const roomRange =
      sheet
        .getRange(2, roomCol + 1, numRows, 1);

    const roomValues =
      roomRange.getValues();


    // ========================================================
    // MATCH EACH DESTINATION ROW
    // ========================================================

    let sheetUpdatedRows = 0;


    for (let i = 0; i < numRows; i++) {

      const destinationProgram =
        programValues[i][0];

      const destinationSection =
        sectionValues[i][0];


      // Ignore rows without ProgramName or Section
      if (
        destinationProgram === '' ||
        destinationProgram === null ||
        destinationSection === '' ||
        destinationSection === null
      ) {
        continue;
      }


      // ------------------------------------------------------
      // NORMALIZE PROGRAM
      // ------------------------------------------------------

      const normalizedProgram =
        normalizeProgramForRoom_(
          destinationProgram
        );


      const normalizedSection =
        normalizeSectionForRoom_(
          destinationSection
        );


      if (
        !normalizedProgram ||
        !normalizedSection
      ) {
        continue;
      }


      // ------------------------------------------------------
      // CREATE SAME LOOKUP KEY
      // ------------------------------------------------------

      const key =
        normalizedProgram +
        '||' +
        normalizedSection;


      // ------------------------------------------------------
      // FIND ROOM
      // ------------------------------------------------------

      if (
        Object.prototype.hasOwnProperty.call(
          roomMap,
          key
        )
      ) {

        const newRoom =
          roomMap[key];


        // Only write when necessary
        if (
          String(roomValues[i][0]).trim() !==
          String(newRoom).trim()
        ) {

          roomValues[i][0] =
            newRoom;

          sheetUpdatedRows++;
          totalRowsUpdated++;
        }
      }
    }


    // ========================================================
    // WRITE ONLY ROOM COLUMN
    // ========================================================

    if (sheetUpdatedRows > 0) {

      roomRange.setValues(roomValues);
    }


    totalSheetsProcessed++;

    Logger.log(
      "Sheet: " +
      sheet.getName() +
      " | Room rows updated: " +
      sheetUpdatedRows
    );

  });


  // ==========================================================
  // LOG DUPLICATE PROGRAM + SECTION ROOMS
  // ==========================================================

  const duplicateKeys =
    Object.keys(duplicateRooms);


  if (duplicateKeys.length > 0) {

    Logger.log(
      "================================================"
    );

    Logger.log(
      "WARNING: Multiple rooms found for the same"
    );

    Logger.log(
      "Program + Section combination."
    );

    Logger.log(
      "The FIRST room found was used."
    );

    Logger.log(
      "================================================"
    );


    duplicateKeys.forEach(function(key) {

      Logger.log(
        key +
        " → " +
        duplicateRooms[key].join(', ')
      );

    });
  }


  // ==========================================================
  // SUMMARY
  // ==========================================================

  let message =
    "Room filling completed successfully.\n\n" +
    "Destination sheets processed: " +
    totalSheetsProcessed +
    "\n" +
    "Rows updated: " +
    totalRowsUpdated;


  if (skippedSheets.length > 0) {

    message +=
      "\n\nSkipped sheets:\n" +
      skippedSheets.join('\n');
  }


  Logger.log(message);


  // Show message when run from spreadsheet
  try {

    Logger.log(
      message
    );

  } catch (e) {
    // Ignore UI error when run from trigger/non-UI context
  }

  return message;
}


/**
 * ============================================================
 * HEADER MAP
 * ============================================================
 *
 * Converts headers such as:
 *
 *   Program
 *   ProgramName
 *   Section
 *   Room
 *
 * into normalized keys.
 */
function createHeaderMapRoom_(headers) {

  const map = {};

  headers.forEach(function(header, index) {

    const normalized =
      normalizeHeaderRoom_(header);

    if (normalized !== '') {

      map[normalized] = index;
    }
  });

  return map;
}


/**
 * ============================================================
 * HEADER NORMALIZATION
 * ============================================================
 */
function normalizeHeaderRoom_(value) {

  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '');
}


/**
 * ============================================================
 * PROGRAM NORMALIZATION
 * ============================================================
 *
 * This is the important part.
 *
 * Examples:
 *
 * F.A (JCW)
 * F.A (Jinnah College)
 *
 * become:
 *
 * F.A
 *
 * Similarly:
 *
 * FCS (JCW)
 * FCS (Jinnah College)
 *
 * become:
 *
 * FCS
 *
 * The function also handles variations such as:
 *
 * F.A(JCW)
 * F.A (JCW)
 * F.A (Jinnah College)
 *
 * ============================================================
 */
function normalizeProgramForRoom_(value) {

  if (
    value === null ||
    value === undefined
  ) {
    return '';
  }


  let program =
    String(value)
      .trim()
      .toUpperCase();


  // Normalize spaces
  program =
    program.replace(/\s+/g, ' ');


  // Remove common college suffixes
  program =
    program.replace(
      /\s*\(\s*(JCW|JINNAH\s+COLLEGE|JINNAH\s+COLLEGE\s+FOR\s+WOMEN)\s*\)\s*$/i,
      ''
    );


  // Also remove these forms without brackets
  program =
    program.replace(
      /\s+(JCW|JINNAH\s+COLLEGE|JINNAH\s+COLLEGE\s+FOR\s+WOMEN)\s*$/i,
      ''
    );


  // Remove spaces immediately inside brackets
  program =
    program.replace(/\(\s+/g, '(');
  program =
    program.replace(/\s+\)/g, ')');


  // Final cleanup
  program =
    program.trim();


  return program;
}


/**
 * ============================================================
 * SECTION NORMALIZATION
 * ============================================================
 *
 * Example:
 *
 * A
 * a
 *  A
 *
 * all become:
 *
 * A
 * ============================================================
 */
function normalizeSectionForRoom_(value) {

  if (
    value === null ||
    value === undefined
  ) {
    return '';
  }


  return String(value)
    .trim()
    .replace(/\s+/g, ' ')
    .toUpperCase();
}
var SM_MERGES_ = [];   // v7.6.1 stopping merit: [first output row, number of tied students]
