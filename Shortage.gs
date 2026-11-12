/**
 * Shortage.gs — Attendance shortage report (page shortage.html), the old "JCW SHORTAGE REPORT — FAST PIPELINE"
 * (Shortage Summary → Fancy Report → Approved Leaves → Shortage with Leaves → Final Shortage Report) inside the app.
 *
 * Any month (Sept–Jan) or the whole session (months added up), any program / year / section / subject,
 * with a threshold you choose. Students struck through (left) and students no longer in
 * Transform-ALL-COLLEGE-STRENGTH are left out. Approved leaves come from "ApprovedLeavesStore" of the month
 * and are added to the attended lectures of every subject (as in the old pipeline).
 * Optional subjects (F.A-1 Optional Subjects …) are joined to the student's own section.
 *
 * The admin can set the SESSION (e.g. 2026-27, used in the titles of every report of the app) and the default
 * threshold, and can save the final report as a sheet "Final Shortage Report - <MONTH>" in the month workbook.
 * Needs WorkbookReports.gs, MonthlyBooks.gs, GrandTest.gs in the same project.
 *
 * INSTALL:
 *   a) + > Script > name it "Shortage" > paste this file.
 *   b) doPost, next to the wr_ line (once):
 *        if (/^shr_/.test(String(req.action || ''))) return json_(shrRoute_(user, req));
 *   c) GrandTest.gs: replace gtSession_ with the new one from the repo (it reads the session the admin sets).
 *   d) Deploy > Manage deployments > Edit > New version > Deploy.
 */
var SHR_CFG = {
  LIMIT: 73,                 // default threshold %, the admin can change it in the app
  CLASS_MAP: {               // program → year → sections (used for the Program / Year filters)
    'Pre-Medical': {
      '1st Year': ['A1','B1','C1','E1','F1','G1','H1','I1','J1','K1','L1','M1','N1','O1','P1','Q1','R1'],
      '2nd Year': ['A2','B2','C2','E2','F2','G2','H2','I2','J2','K2','L2','M2','N2','O2','P2']
    },
    'Pre-Engineering': { '1st Year': ['Pre-Eng-1 D1'], '2nd Year': ['Pre-Eng-2 D1'] },
    'FCS': { '1st Year': ['FCS-1 D2','FCS-1 D3'], '2nd Year': ['FCS-2 D2','FCS-2 D3'] },
    'FA': {
      '1st Year': ['F.A-1 A','F.A-1 B','F.A-1 Optional Subjects'],
      '2nd Year': ['F.A-2 A','F.A-2 B','F.A-2 Optional Subjects']
    }
  }
};

function shrRoute_(user, req) {
  try {
    var a = String(req.action || '');
    if (a === 'shr_leaves') return shrLeavesFor_(req.mons || []);        // every teacher (her registers add the leaves)
    if (!wrAll_(user)) return { ok: false, error: 'The shortage report is for the admin, principal and HODs.' };
    if (a === 'shr_cfg') return shrCfg_(user);
    if (a === 'shr_month') return shrMonth_(String(req.mon || '').toUpperCase(), !!req.fresh);
    if (!isAdmin_(user)) return { ok: false, error: 'Only the admin can change this.' };
    if (a === 'shr_cfg_set') return shrCfgSet_(req);
    if (a === 'shr_write') return shrWrite_(req);
    return { ok: false, error: 'UNKNOWN_ACTION' };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}

function shrProps_() { return PropertiesService.getScriptProperties(); }
function shrCfg_(user) {
  var p = shrProps_();
  return { ok: true, admin: !!isAdmin_(user), session: gtSession_(), sessionSet: p.getProperty('APP_SESSION') || '',
           limit: Number(p.getProperty('SHR_LIMIT')) || SHR_CFG.LIMIT, classMap: SHR_CFG.CLASS_MAP,
           months: SEASON.filter(function (t) { return !!MA_BOOKS[t]; }) };
}
function shrCfgSet_(req) {
  var p = shrProps_();
  if (req.session !== undefined) {
    var s = String(req.session || '').trim();
    if (s && !/^\d{4}-\d{2}(\d{2})?$/.test(s)) return { ok: false, error: 'Session must look like 2026-27.' };
    if (s) p.setProperty('APP_SESSION', s); else p.deleteProperty('APP_SESSION');
  }
  if (req.limit !== undefined) {
    var n = Number(req.limit);
    if (!(n > 0 && n <= 100)) return { ok: false, error: 'Threshold must be between 1 and 100.' };
    p.setProperty('SHR_LIMIT', String(n));
  }
  return { ok: true, session: gtSession_(), sessionSet: p.getProperty('APP_SESSION') || '', limit: Number(p.getProperty('SHR_LIMIT')) || SHR_CFG.LIMIT };
}

/** One month: { people: {roll: [name, father, section]}, subs: [[roll, subject, attended, held]], leaves: {roll: n} } */
function shrMonth_(mon, fresh) {
  if (!MA_BOOKS[mon]) return { ok: true, mon: mon, people: {}, subs: [], leaves: {} };
  var key = 'wr_rows_' + mon, data = fresh ? null : wrCacheGet_(key);
  if (!data) { data = wrRowsBuild_(mon); wrCachePut_(key, data, 1800); }
  var people = {}, subs = [];
  getMaster_().rows.forEach(function (r) { people[gtS_(r[0])] = [gtS_(r[1]), gtS_(r[4]), gtS_(r[2])]; });
  data.forEach(function (g) {
    g.st.forEach(function (s) { if (people[s[0]]) subs.push([s[0], g.sub, s[2], s[3], g.cls]); });
  });
  var used = {}; subs.forEach(function (s) { used[s[0]] = 1; });
  Object.keys(people).forEach(function (k) { if (!used[k]) delete people[k]; });
  return { ok: true, mon: mon, people: people, subs: subs, leaves: shrLeaves_(mon), at: Date.now() };
}

/** { ok, leaves: { MON: { roll: days } } } for the student registers (cached 5 minutes per month) */
function shrLeavesFor_(mons) {
  var out = {}, cache = CacheService.getScriptCache();
  mons.forEach(function (m) {
    m = String(m || '').toUpperCase(); if (!MA_BOOKS[m]) return;
    var hit = cache.get('shr_lv_' + m);
    if (hit) { try { out[m] = JSON.parse(hit); return; } catch (e) {} }
    out[m] = shrLeaves_(m);
    try { cache.put('shr_lv_' + m, JSON.stringify(out[m]), 300); } catch (e) {}
  });
  return { ok: true, leaves: out };
}

/** ApprovedLeavesStore of the month: { roll: number of approved leave days } */
function shrLeaves_(mon) {
  var out = {};
  try {
    var sh = SpreadsheetApp.openById(MA_BOOKS[mon]).getSheetByName('ApprovedLeavesStore');
    if (!sh || sh.getLastRow() < 2) return out;
    sh.getRange(2, 2, sh.getLastRow() - 1, 5).getDisplayValues().forEach(function (r) {
      var k = gtS_(r[0]), n = parseFloat(r[4]);
      if (k && !isNaN(n)) out[k] = (out[k] || 0) + n;
    });
  } catch (e) {}
  return out;
}

/** Saves the final report as a sheet in the month workbook.
 *  req: { mon, title, sheet, sections: [{ sec, rows: [[sno, roll, name, father, shortages, leaves, withLeaves]] }] } */
function shrWrite_(req) {
  var mon = String(req.mon || '').toUpperCase();
  if (!MA_BOOKS[mon]) return { ok: false, error: 'Pick one month (the whole session has no workbook of its own).' };
  var ss = SpreadsheetApp.openById(MA_BOOKS[mon]);
  var name = String(req.sheet || 'Final Shortage Report').slice(0, 90);
  var old = ss.getSheetByName(name); if (old) ss.deleteSheet(old);
  var sh = ss.insertSheet(name);
  var N = 7, grid = [], bg = [], secRows = [], headRows = [], blocks = [];
  var blank = function () { return ['', '', '', '', '', '', '']; };
  var push = function (v, c) { grid.push(v); bg.push([c, c, c, c, c, c, c]); return grid.length; };
  push([String(req.title || name), '', '', '', '', '', ''], '#ffffff');
  push(['Generated: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMMM yyyy h:mm a'), '', '', '', '', '', ''], '#ffffff');
  push(blank(), '#ffffff');
  (req.sections || []).forEach(function (s) {
    secRows.push(push(['Class-Sec: ' + s.sec, '', '', '', '', '', ''], '#b7d7f5'));
    headRows.push(push(['S.No', 'Enroll#', 'Student Name', 'Father Name', 'Shortages', 'Approved Leaves', 'Shortage with Leaves'], '#c9daf8'));
    var st = grid.length + 1;
    (s.rows || []).forEach(function (r, i) { push(r.slice(0, N).map(function (x) { return x == null ? '' : x; }), i % 2 ? '#f4f4f4' : '#ffffff'); });
    if ((s.rows || []).length) blocks.push('R' + st + 'C1:R' + (grid.length) + 'C' + N);
    push(blank(), '#ffffff');
  });
  if (sh.getMaxRows() < grid.length) sh.insertRowsAfter(sh.getMaxRows(), grid.length - sh.getMaxRows());
  sh.getRange(1, 1, grid.length, N).setNumberFormat('@').setValues(grid).setBackgrounds(bg).setWrap(true).setVerticalAlignment('top');
  sh.getRange(1, 1, 1, N).merge().setFontWeight('bold').setFontSize(18).setHorizontalAlignment('center');
  sh.getRange(2, 1, 1, N).merge().setFontStyle('italic').setHorizontalAlignment('right');
  secRows.forEach(function (r) { sh.getRange(r, 1, 1, N).merge(); });
  if (secRows.length) {
    sh.getRangeList(secRows.map(function (r) { return 'R' + r + 'C1:R' + r + 'C' + N; })).setFontWeight('bold').setFontSize(15);
    sh.getRangeList(headRows.map(function (r) { return 'R' + r + 'C1:R' + r + 'C' + N; })).setFontWeight('bold').setHorizontalAlignment('center');
  }
  if (blocks.length) sh.getRangeList(blocks).setBorder(true, true, true, true, true, true);
  sh.setColumnWidth(1, 45); sh.setColumnWidth(2, 120); sh.setColumnWidth(3, 170); sh.setColumnWidth(4, 170);
  sh.setColumnWidth(5, 360); sh.setColumnWidth(6, 80); sh.setColumnWidth(7, 360);
  sh.setFrozenRows(2);
  return { ok: true, sheet: name, book: ss.getName(), url: ss.getUrl() + '#gid=' + sh.getSheetId() };
}
