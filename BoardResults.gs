/**
 * BoardResults.gs — Peshawar Board gazette (result PDF) -> a sheet in the Provisional Certificates workbook.
 *
 * The admin opens board.html in the app, chooses the PDF, "1st year" / "2nd year" and the year (2026).
 * The phone / laptop reads the PDF (the same steps as the Colab notebook "HSSC Part I-II Gazette to excel"):
 *   every table row of every page, the college name above each block of students, Marks as a number,
 *   Position = dense rank of Marks (highest = 1) over every college except PRIVATE CANDIDATES,
 *   sorted by Marks (highest first), and the Jinnah College top 20.
 * It then sends the rows here in parts; this file only writes them:
 *   sheet  "HSSC 1st year 2026"                 Roll No | Name | Father Name | Registration # | Marks | Remarks | College Name | Position
 *   sheet  "HSSC 1st year 2026 - JCW Top 20"    S.No. | Roll No | Name | Father Name | Marks | Position
 * Every new gazette gets its own sheet in the same workbook; old sheets are never touched unless the admin
 * chooses "Replace".
 *
 * INSTALL: add this file, and in doPost (with the other module lines):
 *     if (/^br_/.test(String(req.action || ''))) return json_(brRoute_(user, req));
 * then Deploy → Manage deployments → Edit → New version.
 */
var BR_CFG = {
  BOOK: '1C2tdr2yXAja0_C5Nk4HIk2m-71Kc4I-FS5VQUhjUtHQ',      // Provisional Certificates workbook
  TEXT_COLS: /^(ROLL ?NO|ROLLNO|REGISTRATION.*|REG.*)$/i        // kept as text: long numbers must not become 2.0251E+13
};

function brRoute_(user, req) {
  try {
    if (!isAdmin_(user)) return { ok: false, error: 'Only the admin can save board results.' };
    switch (String(req.action || '')) {
      case 'br_list':   return brList_();
      case 'br_begin':  return brBegin_(req);
      case 'br_rows':   return brRows_(req);
      case 'br_finish': return brFinish_(req);
      case 'br_get':    return brGet_(req);
    }
    return { ok: false, error: 'UNKNOWN_ACTION' };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}
function brBook_() { return SpreadsheetApp.openById(BR_CFG.BOOK); }
function brName_(n) {
  n = String(n || '').replace(/[\[\]\*\?\/\\:]/g, ' ').replace(/\s+/g, ' ').trim();
  if (!n) throw new Error('Sheet name is empty.');
  return n.slice(0, 95);
}

/** the result sheets already in the workbook */
function brList_() {
  return { ok: true, url: brBook_().getUrl(), sheets: brBook_().getSheets().map(function (sh) { return { name: sh.getName(), rows: Math.max(0, sh.getLastRow() - 1), gid: sh.getSheetId() }; }) };
}

/** v7.5: { sheet, start, n } -> rows of a result sheet (for Excel / PDF made on the phone or laptop) */
function brGet_(req) {
  var sh = brBook_().getSheetByName(String(req.sheet || ''));
  if (!sh) return { ok: false, error: 'Sheet not found: ' + req.sheet };
  var last = sh.getLastRow(), w = sh.getLastColumn(), start = Math.max(2, Number(req.start) || 2), n = Math.min(Number(req.n) || 10000, 15000);
  var head = w ? sh.getRange(1, 1, 1, w).getDisplayValues()[0] : [];
  if (req.jcw) {          // v7.6: only Jinnah College for Women (University of Peshawar) rows, with Section (PCS) and Contact from the strength sheets
    var all = last > 1 ? sh.getRange(2, 1, last - 1, w).getDisplayValues() : [], cc = -1;
    head.forEach(function (h, j) { if (cc < 0 && /college/i.test(h)) cc = j; });
    var mine = cc < 0 ? all : all.filter(function (r) { return /JINNAH\s*COLLEGE\s*FOR\s*WOMEN/i.test(r[cc]) && (!/UNIVERSITY|PESHAWAR/i.test(r[cc]) || /PESHAWAR/i.test(r[cc])); });
    var E = brEnrich_(head, mine);
    return { ok: true, header: E.head, rows: E.rows, total: E.rows.length, next: 0, matched: E.matched, contactCol: E.contactCol };
  }
  var rows = last >= start && w ? sh.getRange(start, 1, Math.min(n, last - start + 1), w).getDisplayValues() : [];
  return { ok: true, header: head, rows: rows, total: Math.max(0, last - 1), next: start + rows.length <= last ? start + rows.length : 0 };
}
/** adds Section (PCS) and Contact: student matched by Name + Father Name in Transform-ALL-COLLEGE-STRENGTH,
 *  then in older copies (e.g. "Transform-ALL-COLLEGE-STRENGTH 2025") for students who have passed out */
function brEnrich_(head, rows) {
  var nm = function (v) { return String(v || '').toUpperCase().replace(/\u00A0/g, ' ').replace(/[^A-Z ]/g, ' ').replace(/\s+/g, ' ').trim(); };
  var hk = function (h) { return String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''); };
  var ni = -1, fi = -1;
  head.forEach(function (h, j) { var k = hk(h); if (ni < 0 && (k === 'name' || k === 'studentname')) ni = j; if (fi < 0 && (k === 'fathername' || k === 'father')) fi = j; });
  if (ni < 0) return { head: head, rows: rows, matched: 0, contactCol: -1 };
  var map = {}, byName = {};
  try {
    var book = SpreadsheetApp.openById(USERS_SS_ID);
    var shs = book.getSheets().filter(function (x) { return /^Transform-ALL-COLLEGE-STRENGTH/i.test(x.getName()); })
      .sort(function (a, b) { return a.getName().length - b.getName().length || b.getName().localeCompare(a.getName()); });   // the current one first, then the newest copies
    shs.forEach(function (t) {
      if (t.getLastRow() < 2) return;
      var v = t.getDataRange().getDisplayValues(), H = v[0].map(hk), c = function (L) { for (var i = 0; i < L.length; i++) { var j = H.indexOf(L[i]); if (j >= 0) return j; } return -1; };
      var a = c(['studentname', 'name']), f = c(['fathername']), p = c(['pcs', 'section']), m = c(['studentcontact', 'contact', 'contactno', 'mobile', 'mobileno', 'phone', 'phoneno', 'cellno']);
      if (a < 0) return;
      v.slice(1).forEach(function (r) {
        var key = nm(r[a]) + '|' + (f >= 0 ? nm(r[f]) : ''), val = { pcs: p >= 0 ? r[p] : '', contact: m >= 0 ? r[m] : '' };
        if (!map[key]) map[key] = val;
        var nk = nm(r[a]); byName[nk] = byName[nk] === undefined ? val : null;     // name alone only when it is unique
      });
    });
  } catch (e) {}
  var hit = 0;
  var out = rows.map(function (r) {
    var x = map[nm(r[ni]) + '|' + (fi >= 0 ? nm(r[fi]) : '')] || byName[nm(r[ni])] || null;
    if (x) hit++;
    return r.concat([x ? x.pcs : '', x ? x.contact : '']);
  });
  return { head: head.concat(['Section (PCS)', 'Contact']), rows: out, matched: hit, contactCol: head.length + 1 };
}

/** { sheet, header:[...], replace } -> empty sheet with the header row */
function brBegin_(req) {
  var name = brName_(req.sheet), head = (req.header || []).map(String), book = brBook_(), sh = book.getSheetByName(name);
  if (!head.length) throw new Error('No header.');
  if (sh && sh.getLastRow() > 1 && !req.replace) return { ok: false, error: 'EXISTS', rows: sh.getLastRow() - 1 };
  if (!sh) sh = book.insertSheet(name);
  sh.clear();
  if (sh.getMaxColumns() < head.length) sh.insertColumnsAfter(sh.getMaxColumns(), head.length - sh.getMaxColumns());
  var need = Math.max(2, Number(req.total) || 0) + 1;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  head.forEach(function (h, c) { if (BR_CFG.TEXT_COLS.test(h.trim())) sh.getRange(1, c + 1, need, 1).setNumberFormat('@'); });
  sh.getRange(1, 1, 1, head.length).setValues([head]);
  return { ok: true, sheet: name };
}

/** { sheet, start (first data row, 2 = just under the header), rows:[[...]] } */
function brRows_(req) {
  var sh = brBook_().getSheetByName(brName_(req.sheet)), rows = req.rows || [];
  if (!sh) throw new Error('Sheet not found: ' + req.sheet);
  if (!rows.length) return { ok: true, n: 0 };
  var w = rows[0].length, start = Math.max(2, Number(req.start) || 2), end = start + rows.length - 1;
  if (sh.getMaxRows() < end) sh.insertRowsAfter(sh.getMaxRows(), end - sh.getMaxRows());
  sh.getRange(start, 1, rows.length, w).setValues(rows.map(function (r) { return r.slice(0, w).concat(new Array(Math.max(0, w - r.length)).fill('')); }));
  return { ok: true, n: rows.length };
}

/** { sheet, n, top:{ sheet, header, rows } } -> colours, frozen header, the Jinnah College top 20 sheet */
function brFinish_(req) {
  var book = brBook_(), sh = book.getSheetByName(brName_(req.sheet));
  if (!sh) throw new Error('Sheet not found: ' + req.sheet);
  var last = Number(req.n || 0) + 1, w = sh.getLastColumn();
  if (sh.getMaxRows() > last + 1) sh.deleteRows(last + 2, sh.getMaxRows() - last - 1);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, w).setBackground('#6B3423').setFontColor('#ffffff').setFontWeight('bold');
  try { sh.autoResizeColumns(1, w); } catch (e) {}
  var out = { ok: true, sheet: sh.getName(), rows: last - 1, url: book.getUrl() + '#gid=' + sh.getSheetId() };
  var T = req.top;
  if (T && T.sheet && T.header) {
    var name = brName_(T.sheet), t = book.getSheetByName(name) || book.insertSheet(name), rows = T.rows || [];
    t.clear();
    t.getRange(1, 1, 1, T.header.length).setValues([T.header]).setBackground('#6B3423').setFontColor('#ffffff').setFontWeight('bold');
    T.header.forEach(function (h, c) { if (BR_CFG.TEXT_COLS.test(String(h).trim())) t.getRange(1, c + 1, rows.length + 1, 1).setNumberFormat('@'); });
    if (rows.length) {
      t.getRange(2, 1, rows.length, T.header.length).setValues(rows);
      t.getRange(1, 1, rows.length + 1, T.header.length).setBorder(true, true, true, true, true, true, '#8A4A35', SpreadsheetApp.BorderStyle.SOLID);
      for (var i = 0; i < rows.length; i += 2) t.getRange(i + 2, 1, 1, T.header.length).setBackground('#FBEDEA');
    }
    t.setFrozenRows(1);
    try { t.autoResizeColumns(1, T.header.length); } catch (e) {}
    out.top = { sheet: name, rows: rows.length };
  }
  return out;
}
