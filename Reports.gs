/**
 * Reports.gs — College reports on the admin side of the attendance app (page reports.html).
 *
 * Every report of the old "JCW Reports" web app, in one place, admin only:
 *   Dashboard, Admission Flow (Left Students sheet), Merit vs Quota, Quota comparison, HOMIC vs Regular,
 *   Program & Category, Section-wise, Top Movers, Subject-wise, Custom column, Quota Strength,
 *   PCS & Category (Fancy), JCW Strength Summary, Enrollment Report, F.A subject strength and lists,
 *   Attendance Availability (old backup files) and Fee Details (CHALLAN DETAILS column).
 * Each report is written to its own sheet in the Users workbook (same sheet names as the old app),
 * and can be downloaded as Excel (one report or all of them in one workbook).
 *
 * The session years (e.g. 2026-27 / 2025-26) are set on the page (Sessions button) and used in every report.
 * If they were never set, they follow the calendar (a new session starts in July).
 *
 * INSTALL:
 *   a) + > Script > name it "Reports" > paste this file.
 *   b) doPost, next to the dut_ line:
 *        if (/^rpt_/.test(String(req.action || ''))) return json_(rptRoute_(user, req));
 *   c) Run rptAuthorize once (Run ▸ Allow), then Deploy > Manage deployments > Edit > New version > Deploy.
 *   Everything else lives inside RPT_ below, so no name here can clash with the rest of the project.
 */
var RPT_CFG = {
  BOOK: '1ARU8zMO2daJUfU_SIIo4baeSOG7qBPuni5TYRIRLBNY',       // Users workbook (Transform sheet, CATEGORY, report sheets)
  LEFT_BOOK: '1ARU8zMO2daJUfU_SIIo4baeSOG7qBPuni5TYRIRLBNY',  // workbook with the Left Students sheet
  STRENGTH: 'Transform-ALL-COLLEGE-STRENGTH',
  CATEGORY: 'CATEGORY',
  LEFT: 'Left Students',                                      // PORTAL STATUS = Cancelled / Migrated
  PROP: 'RPT_SESSIONS'                                        // Script Properties: { cur: '2026-27', prev: '2025-26' }
};

function rptAuthorize() {
  SpreadsheetApp.openById(RPT_CFG.BOOK).getName();
  DriveApp.getRootFolder().getName();
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
  Logger.log('Reports: permissions OK. Now Deploy → Manage deployments → Edit → New version.');
}

/** Session years: saved on the page, otherwise from today's date (July starts a new session). */
function rptSessions_() {
  let s = {};
  try { s = JSON.parse(PropertiesService.getScriptProperties().getProperty(RPT_CFG.PROP) || '{}'); } catch (e) { s = {}; }
  const now = new Date(), y = now.getMonth() >= 6 ? now.getFullYear() : now.getFullYear() - 1;
  const lbl = a => a + '-' + String((a + 1) % 100).padStart(2, '0');
  return { cur: s.cur || lbl(y), prev: s.prev || lbl(y - 1), saved: !!(s.cur && s.prev) };
}

function rptRoute_(user, req) {
  try {
    if (!isAdmin_(user)) return { ok: false, error: 'Only the admin can open the college reports.' };
    const action = String(req.action || '');
    if (action === 'rpt_sessions') {
      const ok = v => /^\d{4}-\d{2}$/.test(String(v || ''));
      if (req.reset) PropertiesService.getScriptProperties().deleteProperty(RPT_CFG.PROP);
      else if (ok(req.cur) && ok(req.prev)) PropertiesService.getScriptProperties().setProperty(RPT_CFG.PROP, JSON.stringify({ cur: req.cur, prev: req.prev }));
      else return { ok: false, error: 'Write sessions like 2026-27.' };
      return { ok: true, sessions: rptSessions_() };
    }
    const s = rptSessions_();
    RPT_.setSessions(s.cur, s.prev);
    switch (action) {
      case 'rpt_meta':     return { ok: true, sessions: s, columns: RPT_.columns() };
      case 'rpt_run':      return Object.assign({ ok: true }, RPT_.run(String(req.key || ''), req.opts || {}, req.extra || {}));
      case 'rpt_xlsx':     return Object.assign({ ok: true }, RPT_.xlsx(String(req.key || '')));
      case 'rpt_xlsx_all': return Object.assign({ ok: true }, RPT_.xlsxAll(req.keys || []));
    }
    return { ok: false, error: 'UNKNOWN_ACTION' };
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}

var RPT_ = (function () {
  const SPREADSHEET_ID = RPT_CFG.BOOK;

  const CMP_CONFIG = {
    SOURCE_SHEET: RPT_CFG.STRENGTH,
    CATEGORY_SHEET: RPT_CFG.CATEGORY,
    CURRENT_CLASS: '1st Year',          // value in the "Class" column
    PREVIOUS_CLASS: '2nd Year',         // value in the "Class" column
    CURRENT_SESSION: '',               // set from the reports page (Sessions)
    PREVIOUS_SESSION: '',
    MERIT_SECTIONS: ['A', 'B', 'C', 'D'],
    NOTE: 'Previous year = students who are CURRENTLY in 2nd Year (after any cancellations/migrations), ' +
          'so it can differ slightly from the original admission figures of last year.'
  };

  const CMP_REPORTS = {
    overview:        { sheet: 'CMP Overview',          file: 'JCW_Comparison_Overview',          build: cmpBuildOverview_ },
    programCategory: { sheet: 'CMP Program-Category',  file: 'JCW_Comparison_Program_Category',  build: cmpBuildProgramCategory_ },
    sections:        { sheet: 'CMP Section-wise',      file: 'JCW_Comparison_Sections',          build: cmpBuildSections_ },
    quota:           { sheet: 'CMP Quota-wise',        file: 'JCW_Comparison_Quota',             build: cmpBuildQuota_ },
    meritQuota:      { sheet: 'CMP Merit vs Quota',    file: 'JCW_Comparison_Merit_vs_Quota',    build: cmpBuildMeritQuota_ },
    subjects:        { sheet: 'CMP Subject-wise',      file: 'JCW_Comparison_Subjects',          build: cmpBuildSubjects_ },
    seats:           { sheet: 'CMP Seat Change',       file: 'JCW_Comparison_Seat_Change',       build: cmpBuildSeats_ },
    movers:          { sheet: 'CMP Top Movers',        file: 'JCW_Comparison_Top_Movers',        build: cmpBuildMovers_ },
    admissionFlow:   { sheet: 'CMP Admission Flow',    file: 'JCW_Admission_Flow',               build: cmpBuildAdmissionFlow_ },
    homic:           { sheet: 'CMP HOMIC vs Regular',  file: 'JCW_HOMIC_vs_Regular',             build: cmpBuildHomic_ },
    fee:             { sheet: 'CMP Fee Details',       file: 'JCW_Fee_Details',                  build: cmpBuildFee_ },
    quotaStrength:   { sheet: 'CMP Quota Strength',    file: 'JCW_Quota_Strength',               build: cmpBuildQuotaStrength_ },
    faStrength:      { sheet: 'CMP FA Subject Strength', file: 'JCW_FA_Subject_Strength',        build: cmpBuildFAStrength_ },
    faLists:         { sheet: 'CMP FA Subject Lists',  file: 'JCW_FA_Subject_Lists',             build: cmpBuildFALists_ },
    custom:          { sheet: 'CMP Custom Column',     file: 'JCW_Comparison_Custom',            build: cmpBuildCustom_, needsExtra: true }
  };

  /* ======================================================================
     PUBLIC FUNCTIONS (called from the HTML page)
     ====================================================================== */

  function cmpGenerate(key, opts, extra) {
    try {
      const data = cmpLoad_(opts);
      return cmpRunOne_(key, data, opts || {}, extra || {});
    } catch (e) {
      return { success: false, error: { message: e.message } };
    }
  }

  function cmpGenerateAll(opts) {
    try {
      const data = cmpLoad_(opts);
      const results = {};
      Object.keys(CMP_REPORTS).forEach(key => {
        if (CMP_REPORTS[key].needsExtra) return;
        results[key] = cmpRunOne_(key, data, opts || {}, {});
      });
      return { success: true, results: results };
    } catch (e) {
      return { success: false, error: { message: e.message } };
    }
  }

  function cmpGetColumns() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(CMP_CONFIG.SOURCE_SHEET);
    if (!sh) return [];
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    return headers.map(h => String(h || '').trim()).filter(h => h);
  }

  function cmpDownload(key) {
    const rep = CMP_REPORTS[key];
    if (!rep) throw new Error('Unknown report: ' + key);
    return getExcelFileForSheet(rep.sheet, rep.file + '_' + cmpToday_() + '.xlsx');
  }

  function cmpDownloadAll() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const names = Object.keys(CMP_REPORTS)
      .map(k => CMP_REPORTS[k].sheet)
      .filter(n => ss.getSheetByName(n));
    if (names.length === 0) throw new Error('No comparison sheets yet. Click "Generate ALL reports" first.');
    return getExcelFileForMultipleSheets(names, 'JCW_Year_Comparison_All_Reports_' + cmpToday_() + '.xlsx');
  }

  /* ======================================================================
     INTERNAL HELPERS (names end with _ so the browser cannot call them)
     ====================================================================== */

  function cmpToday_() {
    return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy');
  }

  function cmpLabels_() {
    return {
      cur: 'Current: ' + CMP_CONFIG.CURRENT_CLASS + ' (' + CMP_CONFIG.CURRENT_SESSION + ')',
      prev: 'Previous: ' + CMP_CONFIG.PREVIOUS_CLASS + ' (' + CMP_CONFIG.PREVIOUS_SESSION + ')'
    };
  }

  function cmpSigned_(n) { return n > 0 ? '+' + n : String(n); }

  function cmpPctValue_(cur, prev) {
    if (prev > 0) return (cur - prev) / prev;
    return cur > 0 ? 'NEW' : '';
  }

  function cmpPctText_(cur, prev) {
    const p = cmpPctValue_(cur, prev);
    if (typeof p === 'number') return (p > 0 ? '+' : '') + (p * 100).toFixed(1) + '%';
    return p || '-';
  }

  function cmpRow_(labels, cur, prev) {
    const diff = cur - prev;
    const status = diff > 0 ? '▲ Increase' : diff < 0 ? '▼ Decrease' : '= No change';
    return labels.concat([cur, prev, diff, cmpPctValue_(cur, prev), status]);
  }

  function cmpProgKind_(program) {
    const p = String(program || '').toUpperCase();
    if (/^F\.?\s?A/.test(p)) return 'FA';
    if (/PRE-?\s?MED/.test(p)) return 'PREMED';
    if (/PRE-?\s?ENG/.test(p)) return 'PREENG';
    if (/FCS/.test(p)) return 'FCS';
    return 'OTHER';
  }

  function cmpCatKind_(cat) {
    const c = String(cat || '').toUpperCase();
    if (!c || c === 'UNKNOWN') return 'UNKNOWN';
    if (c.indexOf('NON') >= 0) return c.indexOf('(E)') >= 0 ? 'NONSUB_E' : 'NONSUB_M';
    if (c.indexOf('SUB') >= 0) return 'SUB';
    return 'UNKNOWN';
  }

  /* ---------- Load + normalise the student data once ---------- */

  function cmpLoad_(opts) {
    opts = opts || {};
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const src = ss.getSheetByName(CMP_CONFIG.SOURCE_SHEET);
    if (!src) throw new Error("Sheet '" + CMP_CONFIG.SOURCE_SHEET + "' not found.");

    const values = src.getDataRange().getValues();
    if (values.length < 2) throw new Error('No student data found.');

    const headers = values.shift().map(h => String(h || '').trim());
    const idx = name => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());

    const iProg = idx('Program'), iClass = idx('Class'), iSec = idx('Section'),
          iQuota = idx('Quota'), iPcs = idx('PCS'), iEnr = idx('Enroll#');
    if (iProg < 0) throw new Error("'Program' column not found.");
    if (iClass < 0) throw new Error("'Class' column not found.");

    const subIdx = [];
    for (let n = 1; n <= 7; n++) { const k = idx('Sub' + n); if (k >= 0) subIdx.push(k); }

    // PCS + Class -> Category (same lookup your Fancy Report uses)
    const catMap = {};
    const cs = ss.getSheetByName(CMP_CONFIG.CATEGORY_SHEET);
    if (cs) {
      const cv = cs.getDataRange().getValues();
      const ch = cv.shift().map(h => String(h || '').trim());
      const a = ch.indexOf('PCS'), b = ch.indexOf('Class'), c = ch.indexOf('Category');
      if (a >= 0 && b >= 0 && c >= 0) {
        cv.forEach(r => { catMap[String(r[a]).trim() + '||' + String(r[b]).trim()] = String(r[c]).trim(); });
      }
    }

    const records = [];
    values.forEach(row => {
      const cls = String(row[iClass] || '').trim();
      const period = cls === CMP_CONFIG.CURRENT_CLASS ? 'cur' : cls === CMP_CONFIG.PREVIOUS_CLASS ? 'prev' : null;
      if (!period) return;

      const enroll = iEnr >= 0 ? String(row[iEnr] || '').trim() : '';
      const isHomic = /-HC\d+/i.test(enroll);
      if (opts.excludeHomic && isHomic) return;

      const program = String(row[iProg] || '').trim() || '(blank)';
      const section = iSec >= 0 ? (String(row[iSec] || '').trim().toUpperCase() || '(blank)') : '(blank)';
      const quota = iQuota >= 0 ? (String(row[iQuota] || '').trim() || '(blank)') : '(blank)';
      const pcs = iPcs >= 0 ? String(row[iPcs] || '').trim() : '';
      const category = catMap[pcs + '||' + cls] || 'UNKNOWN';

      const subjects = [];
      subIdx.forEach(k => { const s = String(row[k] || '').trim(); if (s) subjects.push(s); });

      records.push({
        period: period, program: program, progKind: cmpProgKind_(program),
        section: section, quota: quota, pcs: pcs,
        category: category, catKind: cmpCatKind_(category),
        isMerit: quota.toLowerCase() === 'open quota' && CMP_CONFIG.MERIT_SECTIONS.indexOf(section) >= 0,
        isHomic: isHomic, subjects: subjects, row: row
      });
    });

    if (records.length === 0) throw new Error('No ' + CMP_CONFIG.CURRENT_CLASS + ' / ' + CMP_CONFIG.PREVIOUS_CLASS + ' students found. Check the Class values.');
    return { ss: ss, records: records, headers: headers };
  }

  /* ---------- Generic comparison table builder ---------- */

  function cmpTable_(title, recs, dims, o) {
    o = o || {};
    const L = cmpLabels_();
    const groups = {};
    recs.forEach(r => {
      const keys = dims.map(d => d.get(r));
      if (keys.some(k => k === null || k === undefined)) return;
      const id = keys.join('\u0001');
      if (!groups[id]) groups[id] = { keys: keys, cur: 0, prev: 0 };
      groups[id][r.period]++;
    });

    const entries = Object.keys(groups).map(k => groups[k]);
    entries.sort((a, b) => {
      for (let i = 0; i < a.keys.length; i++) {
        const c = String(a.keys[i]).localeCompare(String(b.keys[i]), undefined, { numeric: true });
        if (c) return c;
      }
      return 0;
    });

    const header = dims.map(d => d.name).concat([L.cur, L.prev, 'Difference', 'Change %', 'Status']);
    const data = [header], types = ['header'];
    const pad = n => Array(n).fill('');
    let tc = 0, tp = 0;

    if (o.subtotal && dims.length > 1) {
      let i = 0;
      while (i < entries.length) {
        const g = entries[i].keys[0];
        let sc = 0, sp = 0;
        while (i < entries.length && entries[i].keys[0] === g) {
          data.push(cmpRow_(entries[i].keys, entries[i].cur, entries[i].prev)); types.push('data');
          sc += entries[i].cur; sp += entries[i].prev; i++;
        }
        data.push(cmpRow_(['Subtotal: ' + g].concat(pad(dims.length - 1)), sc, sp)); types.push('subtotal');
        tc += sc; tp += sp;
      }
    } else {
      entries.forEach(e => {
        data.push(cmpRow_(e.keys, e.cur, e.prev)); types.push('data');
        tc += e.cur; tp += e.prev;
      });
    }

    data.push(cmpRow_(['TOTAL'].concat(pad(dims.length - 1)), tc, tp)); types.push('total');

    return {
      title: title, data: data, types: types, nKeys: dims.length,
      colTypes: dims.map(() => 'text').concat(['int', 'int', 'diff', 'pct', 'status']),
      chart: dims.length === 1 && o.chart !== false && entries.length > 0 && entries.length <= 25
    };
  }

  function cmpTableFromRows_(title, header, rows, nKeys) {
    const data = [header].concat(rows), types = ['header'].concat(rows.map(() => 'data'));
    let tc = 0, tp = 0;
    rows.forEach(r => { tc += r[nKeys]; tp += r[nKeys + 1]; });
    data.push(cmpRow_(['TOTAL'].concat(Array(nKeys - 1).fill('')), tc, tp)); types.push('total');
    return {
      title: title, data: data, types: types, nKeys: nKeys,
      colTypes: Array(nKeys).fill('text').concat(['int', 'int', 'diff', 'pct', 'status']),
      chart: false
    };
  }

  const CMP_DIM = {
    program:  { name: 'Program',        get: r => r.program },
    category: { name: 'Category',       get: r => r.category },
    quota:    { name: 'Quota',          get: r => r.quota },
    section:  { name: 'Section',        get: r => r.section },
    admType:  { name: 'Admission Type', get: r => r.isMerit ? 'Merit (Open Quota, Sections A-D)' : 'Quota / Other' }
  };

  /* ======================================================================
     REPORT BUILDERS
     ====================================================================== */

  function cmpBuildOverview_(d) {
    const L = cmpLabels_();
    const recs = d.records;
    const cur = recs.filter(r => r.period === 'cur').length;
    const prev = recs.filter(r => r.period === 'prev').length;
    const diff = cur - prev;
    const tone = diff > 0 ? 'pos' : diff < 0 ? 'neg' : '';
    const programs = {};
    recs.forEach(r => { programs[r.program] = true; });

    return {
      title: 'Overview: Current vs Previous Year Strength',
      kpis: [
        { label: L.cur, value: cur },
        { label: L.prev, value: prev },
        { label: 'Net change', value: cmpSigned_(diff), tone: tone },
        { label: 'Change %', value: cmpPctText_(cur, prev), tone: tone },
        { label: 'Programs', value: Object.keys(programs).length }
      ],
      tables: [
        cmpTable_('Strength by Program', recs, [CMP_DIM.program]),
        cmpTable_('Strength by Category', recs, [CMP_DIM.category]),
        cmpTable_('Strength by Quota', recs, [CMP_DIM.quota])
      ]
    };
  }

  function cmpBuildProgramCategory_(d) {
    return {
      title: 'Program & Category Comparison',
      tables: [
        cmpTable_('Program-wise', d.records, [CMP_DIM.program]),
        cmpTable_('Program > Category (with subtotals)', d.records, [CMP_DIM.program, CMP_DIM.category], { subtotal: true }),
        cmpTable_('Category-wise (all programs)', d.records, [CMP_DIM.category])
      ]
    };
  }

  function cmpBuildSections_(d) {
    return {
      title: 'Section-wise Comparison',
      notes: ['Sections that exist in only one year show 0 in the other year (new or closed sections).'],
      tables: [
        cmpTable_('Program > Section (with subtotals)', d.records, [CMP_DIM.program, CMP_DIM.section], { subtotal: true })
      ]
    };
  }

  function cmpBuildQuota_(d) {
    return {
      title: 'Quota-wise Comparison',
      tables: [
        cmpTable_('Quota-wise (all programs)', d.records, [CMP_DIM.quota]),
        cmpTable_('Quota > Program (with subtotals)', d.records, [CMP_DIM.quota, CMP_DIM.program], { subtotal: true })
      ]
    };
  }

  function cmpBuildMeritQuota_(d) {
    return {
      title: 'Merit vs Quota Comparison',
      notes: ['Merit = "Open Quota" students in Sections A, B, C, D. Everything else = Quota / Other (same rule as your Fancy Report).'],
      tables: [
        cmpTable_('Merit vs Quota (all programs)', d.records, [CMP_DIM.admType]),
        cmpTable_('Program > Admission Type (with subtotals)', d.records, [CMP_DIM.program, CMP_DIM.admType], { subtotal: true })
      ]
    };
  }

  function cmpBuildSubjects_(d) {
    const sr = [];
    d.records.forEach(r => r.subjects.forEach(s => sr.push({ period: r.period, program: r.program, subject: s })));
    const subj = { name: 'Subject', get: r => r.subject };
    const prog = { name: 'Program', get: r => r.program };
    return {
      title: 'Subject-wise Comparison',
      notes: ['A student is counted once for every subject she has (Sub1 to Sub7).'],
      tables: [
        cmpTable_('Subject-wise (all programs)', sr, [subj]),
        cmpTable_('Program > Subject (with subtotals)', sr, [prog, subj], { subtotal: true })
      ]
    };
  }

  function cmpBuildSeats_(d) {
    const recs = d.records;
    const L = cmpLabels_();
    const defs = [
      { label: 'FSC-I (M) SUB (Pre-Medical + FCS + Pre-Engineering)', test: r => r.progKind !== 'FA' && r.catKind === 'SUB' },
      { label: 'FSC-I (M) NON-SUB (Pre-Medical)', test: r => r.progKind === 'PREMED' && r.catKind === 'NONSUB_M' },
      { label: 'FSC-I (E) NON-SUB (Pre-Medical)', test: r => r.progKind === 'PREMED' && r.catKind === 'NONSUB_E' },
      { label: 'F.A', test: r => r.progKind === 'FA' }
    ];
    const used = r => defs.some(df => df.test(r));

    const rows = defs.map(df => {
      let c = 0, p = 0;
      recs.forEach(r => { if (df.test(r)) { if (r.period === 'cur') c++; else p++; } });
      return cmpRow_([df.label], c, p);
    });

    let uc = 0, up = 0;
    recs.forEach(r => { if (!used(r)) { if (r.period === 'cur') uc++; else up++; } });
    if (uc + up > 0) rows.push(cmpRow_(['Not classified (check CATEGORY sheet / other programs)'], uc, up));

    const header = ['Category', L.cur, L.prev, 'Difference', 'Change %', 'Status'];
    const t = cmpTableFromRows_('Increase / Decrease in Seats', ['Category'].concat(header.slice(1)), rows, 1);
    t.data[t.data.length - 1][0] = 'INCREASED SEATS (NET)';
    t.chart = true;
    return {
      title: 'Seat Increase / Decrease',
      notes: ['Uses the CATEGORY sheet (PCS + Class lookup) to split Subsidized / Non-Subsidized (M) / Non-Subsidized (E).'],
      tables: [t]
    };
  }

  function cmpBuildMovers_(d) {
    const L = cmpLabels_();
    const base = cmpTable_('', d.records, [CMP_DIM.program, CMP_DIM.section]);
    const rows = base.data.filter((r, i) => base.types[i] === 'data');
    const header = ['Program', 'Section', L.cur, L.prev, 'Difference', 'Change %', 'Status'];
    const up = rows.filter(r => r[4] > 0).sort((a, b) => b[4] - a[4]).slice(0, 10);
    const down = rows.filter(r => r[4] < 0).sort((a, b) => a[4] - b[4]).slice(0, 10);
    return {
      title: 'Top Movers (Sections with biggest change)',
      tables: [
        cmpTableFromRows_('Top 10 Increases', header, up, 2),
        cmpTableFromRows_('Top 10 Decreases', header, down, 2)
      ]
    };
  }

  function cmpBuildCustom_(d, extra) {
    const col = String((extra && extra.column) || '').trim();
    if (!col) throw new Error('Please choose a column.');
    const k = d.headers.findIndex(h => h === col);
    if (k < 0) throw new Error("Column '" + col + "' not found.");
    const dim = { name: col, get: r => String(r.row[k] === undefined || r.row[k] === null ? '' : r.row[k]).trim() || '(blank)' };
    return {
      title: 'Comparison by "' + col + '"',
      tables: [
        cmpTable_(col + '-wise (all programs)', d.records, [dim]),
        cmpTable_('Program > ' + col + ' (with subtotals)', d.records, [CMP_DIM.program, dim], { subtotal: true })
      ]
    };
  }

  /* ======================================================================
     RUN ONE REPORT: build -> write sheet -> return payload for the web page
     ====================================================================== */

  function cmpRunOne_(key, data, opts, extra) {
    const rep = CMP_REPORTS[key];
    if (!rep) throw new Error('Unknown report: ' + key);

      const built = rep.build(data, extra, opts);
    const L = cmpLabels_();
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm');
      const subtitle = (built.subtitle || (L.cur + '  vs  ' + L.prev)) + (opts.excludeHomic ? '   |   HOMIC students excluded' : '') + '   |   Generated ' + stamp;
      const notes = (built.skipDefaultNote ? [] : [CMP_CONFIG.NOTE]).concat(built.notes || []);

    const sheetInfo = cmpWriteSheet_(data.ss, rep.sheet, built.title, subtitle, notes, built.kpis || [], built.tables);

    return {
      success: true, key: key, sheetName: rep.sheet, sheetUrl: sheetInfo.url,
      title: built.title, subtitle: subtitle, notes: notes,
      kpis: built.kpis || [], tables: built.tables, collapse: !!built.collapse
    };
  }

  /* ======================================================================
     SHEET WRITER (formatting, colours, charts)
     ====================================================================== */
  function cmpWriteSheet_(ss, sheetName, title, subtitle, notes, kpis, tables) {
    let sh = ss.getSheetByName(sheetName);
    if (sh) ss.deleteSheet(sh);
    sh = ss.insertSheet(sheetName);

    const maxCols = Math.max.apply(null, [7].concat(tables.map(t => t.data[0].length)));
    const WHITE = '#ffffff';
    const grid = [], merges = [], borders = [], heights = [], tpos = [];

    const add = () => {
      const rw = {
        v: Array(maxCols).fill(''), bg: Array(maxCols).fill(WHITE), fc: Array(maxCols).fill('#000000'),
        fw: Array(maxCols).fill('normal'), nf: Array(maxCols).fill('General'), ha: Array(maxCols).fill('center')
      };
      grid.push(rw);
      return rw;
    };
    const fillRow = (rw, bg, fc, fw) => {
      for (let c = 0; c < maxCols; c++) { rw.bg[c] = bg; if (fc) rw.fc[c] = fc; if (fw) rw.fw[c] = fw; }
    };

    // Title, subtitle, notes
    let rw = add(); rw.v[0] = title; fillRow(rw, '#4A2418', '#ffffff', 'bold');
    merges.push([grid.length, maxCols]); heights.push([grid.length, 34]);

    rw = add(); rw.v[0] = subtitle; fillRow(rw, '#eaf1fb', '#000000', 'normal');
    merges.push([grid.length, maxCols]);

    notes.forEach(n => {
      rw = add(); rw.v[0] = 'Note: ' + n; fillRow(rw, WHITE, '#666666', 'normal'); rw.ha[0] = 'left';
      merges.push([grid.length, maxCols]); heights.push([grid.length, 28]);
    });
    add();

    // KPIs
    kpis.forEach(k => {
      rw = add();
      rw.v[0] = k.label; rw.v[1] = k.value;
      rw.bg[0] = rw.bg[1] = '#f0f5ff'; rw.fw[0] = rw.fw[1] = 'bold'; rw.ha[0] = 'left';
      if (k.tone === 'pos') rw.fc[1] = '#137333';
      if (k.tone === 'neg') rw.fc[1] = '#c5221f';
      borders.push([grid.length, 1, 2]);
    });
    if (kpis.length) add();

    // Tables
    tables.forEach(t => {
      const n = t.data.length, nc = t.data[0].length, nk = t.nKeys;
          const ct = t.colTypes || [];
      const dCol = ct.indexOf('diff'), pCol = ct.indexOf('pct'), sCol = ct.indexOf('status');

      rw = add(); rw.v[0] = t.title; fillRow(rw, '#f1c232', '#000000', 'bold'); rw.ha[0] = 'left';
      merges.push([grid.length, nc]);

      const top = grid.length + 1;
      t.data.forEach((row, i) => {
        const type = t.types[i];
        rw = add();
        for (let c = 0; c < nc; c++) rw.v[c] = row[c];
        for (let c = 0; c < nk; c++) rw.ha[c] = 'left';

        if (type === 'header') {
          for (let c = 0; c < nc; c++) { rw.bg[c] = '#6B3423'; rw.fc[c] = '#ffffff'; rw.fw[c] = 'bold'; rw.ha[c] = 'center'; }
          heights.push([grid.length, 44]);
          return;
        }
        if (type === 'subtotal') for (let c = 0; c < nc; c++) { rw.bg[c] = '#dce6f5'; rw.fw[c] = 'bold'; }
        if (type === 'total')    for (let c = 0; c < nc; c++) { rw.bg[c] = '#ffe9a8'; rw.fw[c] = 'bold'; }
        if (type === 'data' && i % 2 === 0) for (let c = 0; c < nc; c++) if (rw.bg[c] === WHITE) rw.bg[c] = '#FBEDEA';

              if (dCol >= 0) rw.nf[dCol] = '+0;-0;0';
        if (pCol >= 0) rw.nf[pCol] = '+0.0%;-0.0%;0.0%';

        const diff = Number(row[dCol]) || 0;
        const pct = row[pCol];
        const status = String(row[sCol] || '');
        const paint = (c, sign) => {
          if (sign > 0) { rw.fc[c] = '#137333'; if (type === 'data') rw.bg[c] = '#e6f4ea'; }
          if (sign < 0) { rw.fc[c] = '#c5221f'; if (type === 'data') rw.bg[c] = '#fce8e6'; }
        };
              if (dCol >= 0) paint(dCol, diff);
        if (pCol >= 0) paint(pCol, typeof pct === 'number' ? pct : (pct === 'NEW' ? 1 : 0));
        if (sCol >= 0) paint(sCol, status.indexOf('▲') === 0 ? 1 : status.indexOf('▼') === 0 ? -1 : 0);
      });

      borders.push([top, n, nc]);
      tpos.push({ top: top, n: n, nc: nc, chart: t.chart, title: t.title });

      const need = t.chart ? top + 17 : top + n + 1;   // leave room so charts do not overlap
      while (grid.length < need) add();
    });

    // ---- ONE bulk write per property ----
    const rng = sh.getRange(1, 1, grid.length, maxCols);
    rng.setValues(grid.map(g => g.v));
    rng.setBackgrounds(grid.map(g => g.bg));
    rng.setFontColors(grid.map(g => g.fc));
    rng.setFontWeights(grid.map(g => g.fw));
    rng.setNumberFormats(grid.map(g => g.nf));
    rng.setHorizontalAlignments(grid.map(g => g.ha));
    rng.setVerticalAlignment('middle');
    sh.getRange(1, 1).setFontSize(16);

    merges.forEach(m => sh.getRange(m[0], 1, 1, m[1]).merge());
    borders.forEach(b => {
      const r = sh.getRange(b[0], 1, b[1], b[2]);
      r.setBorder(true, true, true, true, true, true, '#94a3b8', SpreadsheetApp.BorderStyle.SOLID);
      r.setBorder(true, true, true, true, null, null, '#6B3423', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    });
    heights.forEach(h => sh.setRowHeight(h[0], h[1]));
    tpos.forEach(p => {
      sh.getRange(p.top, 1, 1, p.nc).setWrap(true);
      if (p.chart && p.n >= 3) {
        const chart = sh.newChart()
          .setChartType(Charts.ChartType.COLUMN)
          .addRange(sh.getRange(p.top, 1, p.n - 1, 3))     // header + data rows, no TOTAL row
          .setNumHeaders(1)
          .setPosition(p.top, maxCols + 2, 0, 0)
          .setOption('title', p.title)
          .setOption('width', 560).setOption('height', 300)
          .setOption('legend', { position: 'bottom' })
          .build();
        sh.insertChart(chart);
      }
    });

    sh.setColumnWidth(1, 260);
    for (let c = 2; c <= maxCols; c++) sh.setColumnWidth(c, 135);
    SpreadsheetApp.flush();

    return { url: ss.getUrl() + '#gid=' + sh.getSheetId() };
  }
  const CMP_FLOW = {
    LEFT_SPREADSHEET_ID: RPT_CFG.LEFT_BOOK,  // workbook holding the Left Students sheet
    LEFT_SHEET: RPT_CFG.LEFT                                          // PORTAL STATUS = Cancelled / Migrated
  };

  // "JCW-F25-1005" -> "25"   (admission batch year from Enroll#)
  function cmpFlowBatch_(enroll) {
    const m = String(enroll || '').toUpperCase().match(/-F(\d{2})-/);
    return m ? m[1] : '';
  }
  // "Pre-Medical (Jinnah College)" -> "Pre-Medical (JCW)" so names match the Transform sheet
  function cmpFlowProg_(p) {
    return String(p || '').replace(/\(\s*Jinnah College\s*\)/i, '(JCW)').trim();
  }

  function cmpFlowLeft_(d, opts) {
    // Students on roll + which Enroll# batch belongs to 1st Year / 2nd Year
    const iEnrT = d.headers.findIndex(h => h.toLowerCase() === 'enroll#');
    const onRoll = {}, tally = { cur: {}, prev: {} };
    if (iEnrT >= 0) d.records.forEach(r => {
      const e = String(r.row[iEnrT] || '').trim().toUpperCase();
      if (!e) return;
      onRoll[e] = true;
      const b = cmpFlowBatch_(e);
      if (b) tally[r.period][b] = (tally[r.period][b] || 0) + 1;
    });
    const top = t => Object.keys(t).sort((a, b) => t[b] - t[a])[0] || '';
    const bCur = top(tally.cur), bPrev = top(tally.prev);
    const batchCls = {};
    if (bCur) batchCls[bCur] = CMP_CONFIG.CURRENT_CLASS;
    if (bPrev && bPrev !== bCur) batchCls[bPrev] = CMP_CONFIG.PREVIOUS_CLASS;

    const classes = [CMP_CONFIG.CURRENT_CLASS, CMP_CONFIG.PREVIOUS_CLASS];
    const out = {};
    const info = { otherBatch: 0, stillOnRoll: 0, sheetsFound: [], bCur: bCur, bPrev: bPrev };

    const ss = CMP_FLOW.LEFT_SPREADSHEET_ID ? SpreadsheetApp.openById(CMP_FLOW.LEFT_SPREADSHEET_ID) : d.ss;
    const sh = ss.getSheetByName(CMP_FLOW.LEFT_SHEET);
    if (!sh) return { list: [], info: info };
    info.sheetsFound.push(CMP_FLOW.LEFT_SHEET);

    const v = sh.getDataRange().getValues();
    if (v.length < 2) return { list: [], info: info };
    const h = v.shift().map(x => String(x || '').trim().toLowerCase());
    const col = names => { for (let i = 0; i < names.length; i++) { const k = h.indexOf(names[i]); if (k >= 0) return k; } return -1; };
    const iEnr = col(['enroll#', 'enrollment#', 'enrollment no', 'enroll no']);
    const iName = col(['student name', 'name']);
    const iProg = col(['program']), iCls = col(['class']), iSec = col(['section']), iQ = col(['quota']);
    const iStatus = col(['portal status', 'status']), iLeft = col(['left in month']);
    const get = (row, i) => i >= 0 ? String(row[i] === null || row[i] === undefined ? '' : row[i]).trim() : '';
    const tz = Session.getScriptTimeZone();

    v.forEach((row, n) => {
      const enroll = get(row, iEnr);
      if (!enroll && !get(row, iName)) return;
      if (opts.excludeHomic && /-HC\d+/i.test(enroll)) return;

      // Class = the class her admission batch is in NOW (falls back to Class column if Enroll# has no -Fyy-)
      const b = cmpFlowBatch_(enroll);
      const cls = b ? batchCls[b] : get(row, iCls);
      if (!cls || classes.indexOf(cls) < 0) { info.otherBatch++; return; }
      if (enroll && onRoll[enroll.toUpperCase()]) { info.stillOnRoll++; return; }

      const st = get(row, iStatus) || row.join(' ');
      const lv = iLeft >= 0 ? row[iLeft] : '';
      const key = enroll ? enroll.toUpperCase() : get(row, iName) + '#' + n;
      out[key] = {
        status: /migrat/i.test(st) ? 'migrated' : 'cancelled',
        cls: cls, classWhenLeft: get(row, iCls), enroll: enroll, name: get(row, iName),
        leftIn: lv instanceof Date ? Utilities.formatDate(lv, tz, 'dd-MM-yyyy') : String(lv || '').trim(),
        program: cmpFlowProg_(get(row, iProg)) || '(blank)',
        section: get(row, iSec).toUpperCase() || '(blank)',
        quota: get(row, iQ) || '(blank)'
      };
    });
    return { list: Object.keys(out).map(k => out[k]), info: info };
  }

  function cmpFlowTable_(title, recs, dims, o) {
    o = o || {};
    const groups = {};
    recs.forEach(r => {
      const keys = dims.map(dm => dm.get(r));
      const id = keys.join('\u0001');
      if (!groups[id]) groups[id] = { keys: keys, present: 0, cancelled: 0, migrated: 0 };
      groups[id][r.status]++;
    });
    const entries = Object.keys(groups).map(k => groups[k]);
    entries.sort((a, b) => {
      for (let i = 0; i < a.keys.length; i++) {
        const c = String(a.keys[i]).localeCompare(String(b.keys[i]), undefined, { numeric: true });
        if (c) return c;
      }
      return 0;
    });

    const row = (labels, g) => {
      const left = g.cancelled + g.migrated, adm = g.present + left;
      return labels.concat([adm, g.present, g.cancelled, g.migrated, left ? -left : 0, (adm && left) ? -left / adm : 0]);
    };
    const add = (s, g) => { s.present += g.present; s.cancelled += g.cancelled; s.migrated += g.migrated; };
    const pad = n => Array(n).fill('');
    const header = dims.map(dm => dm.name).concat(['Total Admitted (Before)', 'Present Strength (After)', 'Cancelled', 'Migrated', 'Net Loss', 'Loss %']);
    const data = [header], types = ['header'];
    const tot = { present: 0, cancelled: 0, migrated: 0 };

    if (o.subtotal && dims.length > 1) {
      let i = 0;
      while (i < entries.length) {
        const g = entries[i].keys[0];
        const sub = { present: 0, cancelled: 0, migrated: 0 };
        while (i < entries.length && entries[i].keys[0] === g) {
          data.push(row(entries[i].keys, entries[i])); types.push('data');
          add(sub, entries[i]); i++;
        }
        data.push(row(['Subtotal: ' + g].concat(pad(dims.length - 1)), sub)); types.push('subtotal');
        add(tot, sub);
      }
    } else {
      entries.forEach(e => { data.push(row(e.keys, e)); types.push('data'); add(tot, e); });
    }
    data.push(row(['TOTAL'].concat(pad(dims.length - 1)), tot)); types.push('total');

    return {
      title: title, data: data, types: types, nKeys: dims.length,
      colTypes: dims.map(() => 'text').concat(['int', 'int', 'int', 'int', 'diff', 'pct']),
      chart: dims.length === 1 && entries.length > 0 && entries.length <= 25
    };
  }

  function cmpBuildAdmissionFlow_(d, extra, opts) {
    opts = opts || {};
    const left = cmpFlowLeft_(d, opts);
    const clsOf = r => r.period === 'cur' ? CMP_CONFIG.CURRENT_CLASS : CMP_CONFIG.PREVIOUS_CLASS;
    const recs = d.records
      .map(r => ({ status: 'present', cls: clsOf(r), program: r.program, section: r.section, quota: r.quota }))
      .concat(left.list);

    const count = s => recs.filter(r => r.status === s).length;
    const pres = count('present'), can = count('cancelled'), mig = count('migrated'), adm = pres + can + mig;

    const D = {
      cls:   { name: 'Class',   get: r => r.cls },
      prog:  { name: 'Program', get: r => r.program },
      sec:   { name: 'Section', get: r => r.section },
      quota: { name: 'Quota',   get: r => r.quota }
    };

    const tables = [
      cmpFlowTable_('Class-wise: Admitted (Before) vs Present (After)', recs, [D.cls]),
      cmpFlowTable_('Class > Program (with subtotals)', recs, [D.cls, D.prog], { subtotal: true }),
      cmpFlowTable_('Class > Program > Section (subtotal per class)', recs, [D.cls, D.prog, D.sec], { subtotal: true }),
      cmpFlowTable_('Class > Quota (with subtotals)', recs, [D.cls, D.quota], { subtotal: true })
    ];

    if (left.list.length) {
      const lst = left.list.slice().sort((a, b) =>
        (a.status + a.cls + a.program + a.section + a.name).localeCompare(b.status + b.cls + b.program + b.section + b.name, undefined, { numeric: true }));
          const data = [['#', 'Enroll#', 'Student Name', 'Batch Now In', 'Class When Left', 'Program', 'Section', 'Quota', 'Left In', 'Status']];
      lst.forEach((r, i) => data.push([i + 1, r.enroll, r.name, r.cls, r.classWhenLeft, r.program, r.section, r.quota, r.leftIn,
          r.status === 'migrated' ? 'Migrated' : 'Cancelled']));
      tables.push({
        title: 'List of Cancelled & Migrated Students', data: data,
              types: ['header'].concat(lst.map(() => 'data')), nKeys: 10,
        colTypes: Array(10).fill('text'), chart: false
      });
    }

    const notes = [
      'Total Admitted (Before) = Present Strength (After) + Cancelled + Migrated.',
       'Present = students now in the Transform sheet. Cancelled / Migrated = "' + CMP_FLOW.LEFT_SHEET + '" sheet (PORTAL STATUS column).',
      'Students are matched to a class by the batch in their Enroll# (1st Year = F' + left.info.bCur + ', 2nd Year = F' + left.info.bPrev +
        '), so a student cancelled while in 1st Year last year is counted with today\'s 2nd Year batch.'
    ];
    if (cmpFlowMissing_(left, CMP_FLOW.LEFT_SHEET)) notes.push('Sheet "' + CMP_FLOW.LEFT_SHEET + '" was not found, so Cancelled / Migrated = 0.');
    if (left.info.otherBatch) notes.push(left.info.otherBatch + ' left student(s) skipped: from older batches (already passed out).');
    if (left.info.stillOnRoll) notes.push(left.info.stillOnRoll + ' left student(s) skipped because the same Enroll# is still in the Transform sheet (re-admitted).');

    return {
      title: 'Admission Flow: Admitted (Before) → Cancelled / Migrated → Present Strength (After)',
      subtitle: 'Session ' + CMP_CONFIG.CURRENT_SESSION + '  |  ' + CMP_CONFIG.CURRENT_CLASS + ' & ' + CMP_CONFIG.PREVIOUS_CLASS,
      skipDefaultNote: true,
      notes: notes,
      kpis: [
        { label: 'Total Admitted (Before)', value: adm },
        { label: 'Cancelled', value: can, tone: can ? 'neg' : '' },
        { label: 'Migrated', value: mig, tone: mig ? 'neg' : '' },
        { label: 'Present Strength (After)', value: pres, tone: 'pos' },
        { label: 'Retention %', value: adm ? (pres / adm * 100).toFixed(1) + '%' : '-' }
      ],
      tables: tables
    };
  }

  function cmpFlowMissing_(left, name) { return left.info.sheetsFound.indexOf(name) < 0; }
  /* ======================================================================
     EXTRA REPORTS: HOMIC, Fee details, Quota strength, F.A subject strength, F.A subject lists
     ====================================================================== */

  const CMP_CLASS_DIM = { name: 'Class', get: r => r.period === 'cur' ? CMP_CONFIG.CURRENT_CLASS : CMP_CONFIG.PREVIOUS_CLASS };

  /* Count table: rows = dims, columns = the values of colDim (+ Total). Subtotals per first key. */
  function cmpCountTable_(title, recs, dims, colDim, colVals, o) {
    o = o || {};
    const groups = {};
    recs.forEach(r => {
      const keys = dims.map(dm => dm.get(r));
      const id = keys.join('\u0001');
      if (!groups[id]) groups[id] = { keys: keys, c: {} };
      const cv = colDim.get(r);
      groups[id].c[cv] = (groups[id].c[cv] || 0) + 1;
    });
    const entries = Object.keys(groups).map(k => groups[k]);
    entries.sort((a, b) => {
      for (let i = 0; i < a.keys.length; i++) {
        const c = String(a.keys[i]).localeCompare(String(b.keys[i]), undefined, { numeric: true });
        if (c) return c;
      }
      return 0;
    });
    const row = (labels, c) => { const v = colVals.map(x => c[x] || 0); return labels.concat(v, [v.reduce((a, b) => a + b, 0)]); };
    const add = (s, c) => colVals.forEach(x => { s[x] = (s[x] || 0) + (c[x] || 0); });
    const pad = n => Array(n).fill('');
    const data = [dims.map(dm => dm.name).concat(colVals, ['Total'])], types = ['header'];
    const tot = {};
    if (o.subtotal && dims.length > 1) {
      let i = 0;
      while (i < entries.length) {
        const g = entries[i].keys[0], sub = {};
        while (i < entries.length && entries[i].keys[0] === g) { data.push(row(entries[i].keys, entries[i].c)); types.push('data'); add(sub, entries[i].c); i++; }
        data.push(row(['Subtotal: ' + g].concat(pad(dims.length - 1)), sub)); types.push('subtotal'); add(tot, sub);
      }
    } else {
      entries.forEach(e => { data.push(row(e.keys, e.c)); types.push('data'); add(tot, e.c); });
    }
    data.push(row(['TOTAL'].concat(pad(dims.length - 1)), tot)); types.push('total');
    return { title: title, data: data, types: types, nKeys: dims.length,
      colTypes: dims.map(() => 'text').concat(colVals.map(() => 'int'), ['int']), chart: dims.length === 1 };
  }

  function cmpSessionLine_() {
    return 'Session ' + CMP_CONFIG.CURRENT_SESSION + ' (' + CMP_CONFIG.CURRENT_CLASS + ')  &  ' + CMP_CONFIG.PREVIOUS_SESSION + ' (' + CMP_CONFIG.PREVIOUS_CLASS + ')';
  }

  function cmpBuildHomic_(d) {
    const dim = { name: 'Student Type', get: r => r.isHomic ? 'HOMIC' : 'Regular' };
    return {
      title: 'HOMIC vs Regular Students',
      notes: ['HOMIC = Enroll# containing -HC. Switch off "Exclude HOMIC" to include them.'],
      tables: [
        cmpTable_('HOMIC vs Regular', d.records, [dim]),
        cmpTable_('Program > Student Type (with subtotals)', d.records, [CMP_DIM.program, dim], { subtotal: true })
      ]
    };
  }

  function cmpBuildQuotaStrength_(d) {
    const C = [CMP_CONFIG.CURRENT_CLASS, CMP_CONFIG.PREVIOUS_CLASS];
    const recs = d.records, quotas = {};
    recs.forEach(r => { quotas[r.quota] = true; });
    const cur = recs.filter(r => r.period === 'cur').length;
    return {
      title: 'Quota Strength Report',
      subtitle: cmpSessionLine_(), skipDefaultNote: true,
      kpis: [
        { label: 'Total Students', value: recs.length },
        { label: CMP_CONFIG.CURRENT_CLASS, value: cur },
        { label: CMP_CONFIG.PREVIOUS_CLASS, value: recs.length - cur },
        { label: 'Quotas', value: Object.keys(quotas).length }
      ],
      tables: [
        cmpCountTable_('Quota-wise Strength', recs, [CMP_DIM.quota], CMP_CLASS_DIM, C),
        cmpCountTable_('Quota > Program (with subtotals)', recs, [CMP_DIM.quota, CMP_DIM.program], CMP_CLASS_DIM, C, { subtotal: true }),
        cmpCountTable_('Program-wise Strength', recs, [CMP_DIM.program], CMP_CLASS_DIM, C)
      ]
    };
  }

  function cmpBuildFAStrength_(d) {
    const C = [CMP_CONFIG.CURRENT_CLASS, CMP_CONFIG.PREVIOUS_CLASS];
    const fa = d.records.filter(r => r.progKind === 'FA');
    if (!fa.length) throw new Error('No F.A students found.');
    const sr = [];
    fa.forEach(r => r.subjects.forEach(s => sr.push({ period: r.period, subject: s })));
    const subjects = {};
    sr.forEach(x => { subjects[x.subject] = true; });
    const cur = fa.filter(r => r.period === 'cur').length;
    return {
      title: 'F.A Subject-wise Strength',
      subtitle: cmpSessionLine_(), skipDefaultNote: true,
      notes: ['A student is counted once for every subject she has (Sub1 to Sub7).'],
      kpis: [
        { label: 'F.A students', value: fa.length },
        { label: CMP_CONFIG.CURRENT_CLASS, value: cur },
        { label: CMP_CONFIG.PREVIOUS_CLASS, value: fa.length - cur },
        { label: 'Subjects', value: Object.keys(subjects).length }
      ],
      tables: [
        cmpCountTable_('Subject-wise Strength (F.A)', sr, [{ name: 'Subject', get: r => r.subject }], CMP_CLASS_DIM, C),
        cmpCountTable_('Section-wise Strength (F.A)', fa, [CMP_DIM.section], CMP_CLASS_DIM, C)
      ]
    };
  }

  function cmpBuildFALists_(d) {
    const fa = d.records.filter(r => r.progKind === 'FA');
    if (!fa.length) throw new Error('No F.A students found.');
    const col = name => d.headers.findIndex(h => h.toLowerCase() === name);
    const iEnr = col('enroll#'), iName = col('student name'), iFather = col('father name'), iPcs = col('pcs');
    const cell = (r, i) => i >= 0 ? String(r.row[i] === null || r.row[i] === undefined ? '' : r.row[i]).trim() : '';
    const groups = {};
    fa.forEach(r => r.subjects.forEach(s => {
      const cls = CMP_CLASS_DIM.get(r), k = cls + '\u0001' + s;
      (groups[k] = groups[k] || { cls: cls, subject: s, list: [] }).list.push(r);
    }));
    const keys = Object.keys(groups).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    const tables = keys.map(k => {
      const g = groups[k];
      g.list.sort((a, b) => (a.section + cell(a, iName)).localeCompare(b.section + cell(b, iName), undefined, { numeric: true }));
      const data = [['#', 'Enroll#', 'Student Name', 'Father Name', 'Section', 'PCS']];
      g.list.forEach((r, i) => data.push([i + 1, cell(r, iEnr), cell(r, iName), cell(r, iFather), r.section, cell(r, iPcs)]));
      return { title: g.subject + ' — ' + g.cls + ' (' + g.list.length + ' students)', data: data,
        types: data.map((_, i) => i ? 'data' : 'header'), nKeys: 6, colTypes: Array(6).fill('text'), chart: false };
    });
    return {
      title: 'F.A Subject Lists',
      subtitle: cmpSessionLine_(), skipDefaultNote: true, collapse: true,
      kpis: [{ label: 'Subject lists', value: tables.length }, { label: 'F.A students', value: fa.length }],
      tables: tables
    };
  }

  function cmpBuildFee_(d) {
    const col = name => d.headers.findIndex(h => h.toLowerCase() === name);
    const iCh = col('challan details');
    if (iCh < 0) throw new Error("'CHALLAN DETAILS' column not found in the Transform sheet.");
    const iEnr = col('enroll#'), iName = col('student name'), iFather = col('father name');
    const tz = Session.getScriptTimeZone();
    const cell = (r, i) => {
      if (i < 0) return '';
      const v = r.row[i];
      if (v instanceof Date) return Utilities.formatDate(v, tz, 'dd-MM-yyyy');
      return String(v === null || v === undefined ? '' : v).trim();
    };
    const recs = d.records.map(r => ({ r: r, cls: CMP_CLASS_DIM.get(r), ch: cell(r, iCh) }));

    const g = {};
    recs.forEach(x => {
      const k = x.r.program + '\u0001' + x.cls;
      if (!g[k]) g[k] = { keys: [x.r.program, x.cls], t: 0, p: 0 };
      g[k].t++; if (x.ch) g[k].p++;
    });
    const row = (lab, t, p) => lab.concat([t, p, p - t, t ? p / t : 0]);
    const sum = [['Program', 'Class', 'Total Students', 'Challan Entered', 'Missing', 'Coverage %']], st = ['header'];
    let T = 0, P = 0;
    Object.keys(g).sort((a, b) => a.localeCompare(b, undefined, { numeric: true })).forEach(k => {
      sum.push(row(g[k].keys, g[k].t, g[k].p)); st.push('data'); T += g[k].t; P += g[k].p;
    });
    sum.push(row(['TOTAL', ''], T, P)); st.push('total');

    const sorted = recs.slice().sort((a, b) =>
      (a.cls + a.r.program + a.r.section + cell(a.r, iName)).localeCompare(b.cls + b.r.program + b.r.section + cell(b.r, iName), undefined, { numeric: true }));
    const listTable = (title, rows) => {
      const data = [['#', 'Enroll#', 'Student Name', 'Father Name', 'Program', 'Class', 'Section', 'CHALLAN DETAILS']];
      rows.forEach((x, i) => data.push([i + 1, cell(x.r, iEnr), cell(x.r, iName), cell(x.r, iFather), x.r.program, x.cls, x.r.section, x.ch || '—']));
      return { title: title, data: data, types: data.map((_, i) => i ? 'data' : 'header'), nKeys: 8, colTypes: Array(8).fill('text'), chart: false };
    };
    const missing = sorted.filter(x => !x.ch);
    const tables = [
      { title: 'Challan Status by Program & Class', data: sum, types: st, nKeys: 2, colTypes: ['text', 'text', 'int', 'int', 'diff', 'pct'], chart: false },
      listTable('All Students with Challan Details (' + sorted.length + ')', sorted)
    ];
    if (missing.length) tables.splice(1, 0, listTable('Students WITHOUT Challan Details (' + missing.length + ')', missing));

    return {
      title: 'Fee Details (CHALLAN DETAILS column of the Transform sheet)',
      subtitle: cmpSessionLine_(), skipDefaultNote: true,
      notes: ['Fee details are read live from the CHALLAN DETAILS column of "' + CMP_CONFIG.SOURCE_SHEET + '", so Class and Section are always current.'],
      kpis: [
        { label: 'Total Students', value: T },
        { label: 'Challan Entered', value: P, tone: 'pos' },
        { label: 'Missing Challan Details', value: T - P, tone: T - P ? 'neg' : '' },
        { label: 'Coverage %', value: T ? (P / T * 100).toFixed(1) + '%' : '-' }
      ],
      tables: tables
    };
  }

  /* ======================================================================
     EXCEL: several report sheets in one workbook
     ====================================================================== */
  function getExcelFileForMultipleSheets(names, downloadName) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const tempSS = SpreadsheetApp.create('TEMP_JCW_Export_' + new Date().getTime());
    const tempFile = DriveApp.getFileById(tempSS.getId());
    try {
      names.forEach(n => { const sh = ss.getSheetByName(n); if (sh) sh.copyTo(tempSS).setName(n); });
      const def = tempSS.getSheetByName('Sheet1');
      if (def && tempSS.getSheets().length > 1) tempSS.deleteSheet(def);
      SpreadsheetApp.flush();
      const url = 'https://docs.google.com/spreadsheets/d/' + tempSS.getId() + '/export?format=xlsx';
      const blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
      return { base64: Utilities.base64Encode(blob.getBytes()), filename: downloadName,
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' };
    } finally {
      tempFile.setTrashed(true);
    }
  }

  /* ======================================================================
     OFFICIAL REPORTS (built by the code from the old reports project)
     ====================================================================== */
  const RPT_OFFICIAL = {
    fancy:  { fn: () => generateFancyReportWithPCS_Fancy(),             sheet: 'Fancy Report Fancy',      file: 'JCW_Detailed_Strength_Report' },
    jcwSum: { fn: () => generateJCWStrengthSummary1(),                  sheet: 'JCW Strength Summary',    file: 'JCW_Increased_Seats' },
    enroll: { fn: () => generateSideBySideEnrollmentReport1(),          sheet: 'JCW Enrollment Report',   file: 'JCW_Enrollment_Report' },
    att:    { fn: () => generateAttendanceAvailabilityFromBackups1(),   sheet: 'ATTENDANCE AVAILABILITY', file: 'JCW_Attendance_Availability' }
  };

  function rptSheetOf_(key) {
    if (CMP_REPORTS[key]) return CMP_REPORTS[key].sheet;
    if (RPT_OFFICIAL[key]) return RPT_OFFICIAL[key].sheet;
    throw new Error('Unknown report: ' + key);
  }

  function rptRun_(key, opts, extra) {
    if (CMP_REPORTS[key]) {
      const r = cmpGenerate(key, opts || {}, extra || {});
      if (!r || r.success === false) throw new Error((r && r.error && r.error.message) || 'Report failed.');
      return r;
    }
    const o = RPT_OFFICIAL[key];
    if (!o) throw new Error('Unknown report: ' + key);
    const r = o.fn();
    if (!r || r.success === false) throw new Error((r && r.error && (r.error.message || r.error)) || 'Report failed.');
    return r;
  }

  function rptXlsx_(key) {
    const file = CMP_REPORTS[key] ? CMP_REPORTS[key].file : RPT_OFFICIAL[key] ? RPT_OFFICIAL[key].file : '';
    if (!file) throw new Error('Unknown report: ' + key);
    return getExcelFileForSheet(rptSheetOf_(key), file + '_' + cmpToday_() + '.xlsx');
  }

  function rptXlsxAll_(keys) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const names = (keys || []).map(k => { try { return rptSheetOf_(k); } catch (e) { return ''; } })
      .filter((n, i, a) => n && a.indexOf(n) === i && ss.getSheetByName(n));
    if (!names.length) throw new Error('No report sheets yet. Generate the reports first.');
    return getExcelFileForMultipleSheets(names, 'JCW_All_Reports_' + cmpToday_() + '.xlsx');
  }

  /* ======================================================================
     OFFICIAL REPORTS (code from the old reports project; ss1 → ss fixed, session labels from Sessions)
     ====================================================================== */

  function getExcelFileForSheet(sheetName, downloadName) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const targetSheet = ss.getSheetByName(sheetName);

    if (!targetSheet) {
      throw new Error(sheetName + ' sheet not found. Generate the report first.');
    }

    const tempSS = SpreadsheetApp.create('TEMP_JCW_Export_' + new Date().getTime());
    const tempFile = DriveApp.getFileById(tempSS.getId());

    targetSheet.copyTo(tempSS).setName(sheetName);
    const defaultSheet = tempSS.getSheetByName('Sheet1');
    if (defaultSheet) tempSS.deleteSheet(defaultSheet);

    SpreadsheetApp.flush();

    const url = "https://docs.google.com/spreadsheets/d/" + tempSS.getId() + "/export?format=xlsx";
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token } });

    const blob = response.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());

    tempFile.setTrashed(true);

    return {
      base64: base64,
      filename: downloadName,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    };
  }
  function generateFancyReportWithPCS_Fancy() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Transform-ALL-COLLEGE-STRENGTH');
    const data = sheet.getDataRange().getValues();

    const headers = data.shift();
    const programIndex = headers.indexOf('Program');
    const classIndex = headers.indexOf('Class');
    const quotaIndex = headers.indexOf('Quota');
    const pcsIndex = headers.indexOf('PCS');       // used for Category lookup - matches CATEGORY sheet
    const sectionIndex = headers.indexOf('Section'); // used only for Merit/Quota classification

    const categorySheet = ss.getSheetByName('CATEGORY');
    const catData = categorySheet.getDataRange().getValues();
    const catHeaders = catData.shift();
    const catPCSIndex = catHeaders.indexOf('PCS');
    const catClassIndex = catHeaders.indexOf('Class');
    const catCategoryIndex = catHeaders.indexOf('Category');

    const categoryLookup = {};
    catData.forEach(r => {
      const key = r[catPCSIndex] + "||" + r[catClassIndex];
      categoryLookup[key] = r[catCategoryIndex];
    });

    // Sections that are eligible to be counted as "Merit" (only when Quota = "Open Quota")
    const MERIT_ELIGIBLE_SECTIONS = ['a', 'b', 'c', 'd'];
    // Program + Category combo that should render as "merit+quota=total" instead of a plain number
    const MERIT_QUOTA_SPLIT_PROGRAM = 'Pre-Medical (JCW)';
    const MERIT_QUOTA_SPLIT_CATEGORY = 'SUBSIDISIZED';

    // --- Seat Increase/Decrease Summary config ---
    // Class labels as they appear in the "Class" column - update if class naming ever changes
    const FIRST_YEAR_CLASS_LABEL = '1st Year';
    const SECOND_YEAR_CLASS_LABEL = '2nd Year';
    // Programs combined into the single "SUB" row (F.A is reported separately)
    const COMBINED_SUB_PROGRAMS = ['Pre-Medical (JCW)', 'FCS (JCW)', 'Pre-Engineering (JCW)'];
    // Column headers for the two sessions being compared - UPDATE THESE EACH YEAR
    const SESSION_FIRST_YEAR_LABEL = 'SESSION (' + CMP_CONFIG.CURRENT_SESSION + ') (FIRST YEAR)';
    const SESSION_SECOND_YEAR_LABEL = 'SESSION (' + CMP_CONFIG.PREVIOUS_SESSION + ') (SECOND YEAR)';

    const report = {};
    data.forEach(row => {
      const program = row[programIndex];
      const cls = row[classIndex];
      const quota = row[quotaIndex];
      const pcs = row[pcsIndex] || "UNKNOWN";
      const key = pcs + "||" + cls;
      const category = categoryLookup[key] || "UNKNOWN";

      if (!report[program]) {
        report[program] = { total: 0, categories: {}, pcs: {}, classes: {}, classCategory: {}, classCategoryMeritQuota: {} };
      }

      report[program].total++;

      if (!report[program].categories[category]) report[program].categories[category] = 0;
      report[program].categories[category]++;

      if (!report[program].pcs[pcs]) report[program].pcs[pcs] = { class: cls, category, count: 0 };
      report[program].pcs[pcs].count++;

      if (!report[program].classes[cls]) report[program].classes[cls] = { quotas: {} };
      if (!report[program].classes[cls].quotas[quota]) report[program].classes[cls].quotas[quota] = 0;
      report[program].classes[cls].quotas[quota]++;

      if (!report[program].classCategory[cls]) report[program].classCategory[cls] = {};
      if (!report[program].classCategory[cls][category]) report[program].classCategory[cls][category] = 0;
      report[program].classCategory[cls][category]++;

      // --- Merit / Quota classification ---
      // Merit  = Quota value is "Open Quota" AND Section is one of a, b, c, d
      // Quota  = everything else (non-open quota, or open quota outside a-d)
      const section = row[sectionIndex] || "";
      const sectionLetter = String(section).trim().toLowerCase();
      const quotaValue = String(quota).trim().toLowerCase();
      const isMeritEligibleSection = MERIT_ELIGIBLE_SECTIONS.includes(sectionLetter);
      const isMerit = isMeritEligibleSection && quotaValue === 'open quota';

      if (!report[program].classCategoryMeritQuota[cls]) report[program].classCategoryMeritQuota[cls] = {};
      if (!report[program].classCategoryMeritQuota[cls][category]) {
        report[program].classCategoryMeritQuota[cls][category] = { merit: 0, quota: 0 };
      }
      if (isMerit) {
        report[program].classCategoryMeritQuota[cls][category].merit++;
      } else {
        report[program].classCategoryMeritQuota[cls][category].quota++;
      }
    });

    const categoryOrder = {
      "SUBSIDISIZED": 1,
      "NON-SUBSIDIZED (M)": 2,
      "NON-SUBSIDIZED (E)": 3
    };

    const classCategoryTable = [];
    for (let program in report) {
      for (let cls in report[program].classCategory) {
        for (let category in report[program].classCategory[cls]) {
          const mq = (report[program].classCategoryMeritQuota[cls] && report[program].classCategoryMeritQuota[cls][category])
            || { merit: 0, quota: 0 };
          classCategoryTable.push({
            program,
            class: cls,
            category,
            total: report[program].classCategory[cls][category],
            merit: mq.merit,
            quota: mq.quota
          });
        }
      }
    }

    let sheetOut = ss.getSheetByName("Fancy Report Fancy");
    if (!sheetOut) sheetOut = ss.insertSheet("Fancy Report Fancy");
    sheetOut.clear();

    let r = 1;

    sheetOut.getRange(r, 1).setValue("JCW Total Students (All Programs)").setFontWeight("bold");
    sheetOut.getRange(r, 2).setValue(data.length).setFontWeight("bold");
    r += 2;

    function writeExtendedTable() {
      // 5th header slot is left blank on purpose - it gets merged into the "Total" header cell below
      const tblHeaders = [
        "Class", "Program", "Category", "Total", "",
        "Program-wise Strength", "Class-wise Strength", "JCW Strength"
      ];

      let extRows = [];
      const programClassStrength = {};
      classCategoryTable.forEach(row => {
        const key = row.class + "|" + row.program;
        programClassStrength[key] = (programClassStrength[key] || 0) + row.total;
      });

      const classStrength = {};
      classCategoryTable.forEach(row => {
        classStrength[row.class] = (classStrength[row.class] || 0) + row.total;
      });

      const jcwStrength = Object.values(classStrength).reduce((a, b) => a + b, 0);

      // Pre-Medical NON-SUBSIDIZED (M) + NON-SUBSIDIZED (E) combined total, per class
      const nonSubTotals = {};
      classCategoryTable.forEach(row => {
        if (row.program === MERIT_QUOTA_SPLIT_PROGRAM &&
            (row.category === 'NON-SUBSIDIZED (M)' || row.category === 'NON-SUBSIDIZED (E)')) {
          const key = row.class + "|" + row.program;
          nonSubTotals[key] = (nonSubTotals[key] || 0) + row.total;
        }
      });

      classCategoryTable.forEach(row => {
        const key = row.class + "|" + row.program;
        const isNonSubRow = row.program === MERIT_QUOTA_SPLIT_PROGRAM &&
          (row.category === 'NON-SUBSIDIZED (M)' || row.category === 'NON-SUBSIDIZED (E)');
        extRows.push({
          class: row.class,
          program: row.program,
          category: row.category,
          total: row.total,
          merit: row.merit,
          quota: row.quota,
          nonSubTotal: isNonSubRow ? nonSubTotals[key] : "",
          programWise: programClassStrength[row.class + "|" + row.program] || "",
          classWise: classStrength[row.class] || "",
          jcwWise: jcwStrength
        });
      });

      extRows.sort((a, b) => {
        const c1 = a.class.localeCompare(b.class);
        if (c1 !== 0) return c1;
        const c2 = a.program.localeCompare(b.program);
        if (c2 !== 0) return c2;
        return (categoryOrder[a.category] || 999) - (categoryOrder[b.category] || 999);
      });

      sheetOut.getRange(r, 1, 1, tblHeaders.length)
        .setValues([tblHeaders]).setFontWeight("bold").setBackground("#f1c232")
        .setHorizontalAlignment("center").setVerticalAlignment("middle");

      // Merge the "Total" header cell over the Total + Non-Subsidized-Total columns
      const totalColIndex = tblHeaders.indexOf("Total") + 1;
      sheetOut.getRange(r, totalColIndex, 1, 2).merge();

      r++;

      const startRow = r;
      const outRows = [tblHeaders];

      // Track contiguous row-ranges where Non-Subsidized Total carries a value,
      // so we can merge that column vertically across just those rows.
      let nonSubMergeRanges = [];
      let currentNonSubStart = null;

      extRows.forEach(row => {
        const isMeritQuotaSplitRow =
          row.program === MERIT_QUOTA_SPLIT_PROGRAM && row.category === MERIT_QUOTA_SPLIT_CATEGORY;
        const totalDisplay = isMeritQuotaSplitRow
          ? (row.merit + "+" + row.quota + "=" + row.total)
          : row.total;
        const categoryDisplay = isMeritQuotaSplitRow
          ? (row.category + " (MERIT+ QUOTA)")
          : row.category;

        const rowArr = [row.class, row.program, categoryDisplay, totalDisplay, row.nonSubTotal, row.programWise, row.classWise, row.jcwWise];
        sheetOut.getRange(r, 1, 1, tblHeaders.length).setValues([rowArr]);
        outRows.push(rowArr);

        if (row.nonSubTotal === "") {
          // No Non-Subsidized Total on this row -> merge Total + Non-Subsidized Total cells horizontally
          sheetOut.getRange(r, totalColIndex, 1, 2).merge();
          if (currentNonSubStart !== null) {
            nonSubMergeRanges.push({ start: currentNonSubStart, end: r - 1 });
            currentNonSubStart = null;
          }
        } else {
          // This row carries a Non-Subsidized Total value -> track it for a vertical merge
          if (currentNonSubStart === null) currentNonSubStart = r;
        }

        r++;
      });
      if (currentNonSubStart !== null) {
        nonSubMergeRanges.push({ start: currentNonSubStart, end: r - 1 });
      }
      nonSubMergeRanges.forEach(range => {
        if (range.end > range.start) {
          sheetOut.getRange(range.start, totalColIndex + 1, range.end - range.start + 1, 1).merge();
        }
      });

      ["Class", "Program", "Program-wise Strength", "Class-wise Strength", "JCW Strength"].forEach(col => {
        const colIndex = tblHeaders.indexOf(col) + 1;
        let mergeStart = startRow;
        let prev = sheetOut.getRange(startRow, colIndex).getValue();

        for (let rr = startRow + 1; rr < r; rr++) {
          const curr = sheetOut.getRange(rr, colIndex).getValue();
          if (curr !== prev) {
            if (rr - mergeStart > 1) sheetOut.getRange(mergeStart, colIndex, rr - mergeStart, 1).merge();
            mergeStart = rr;
            prev = curr;
          }
        }
        if (r - mergeStart > 1) sheetOut.getRange(mergeStart, colIndex, r - mergeStart, 1).merge();
      });

      sheetOut.getRange(startRow - 1, 1, r - startRow + 1, tblHeaders.length)
        .setBorder(true, true, true, true, true, true)
        .setHorizontalAlignment("center").setVerticalAlignment("middle");

      r += 2;
      return outRows;
    }

    function writeExtendedTablePCSWithCategoryLookup() {
      const tblHeaders = [
        "Class", "Program", "Category", "Section", "Strength",
        "Program-wise Strength", "Class-wise Strength", "JCW Strength"
      ];

      let extRows = [];

      const baseData = sheet.getDataRange().getValues();
      const baseHeaders = baseData.shift();
      const baseProgramIdx = baseHeaders.indexOf('Program');
      const baseClassIdx = baseHeaders.indexOf('Class');
      const basePCSIdx = baseHeaders.indexOf('PCS');

      const catSheet = ss.getSheetByName('CATEGORY');
      const catData2 = catSheet.getDataRange().getValues();
      const catHeaders2 = catData2.shift();
      const catPCSIdx = catHeaders2.indexOf('PCS');
      const catClassIdx = catHeaders2.indexOf('Class');
      const catCategoryIdx = catHeaders2.indexOf('Category');

      const categoryLookup2 = {};
      catData2.forEach(r2 => {
        const key = r2[catPCSIdx] + "|" + r2[catClassIdx];
        categoryLookup2[key] = r2[catCategoryIdx];
      });

      const pcsStrength = {};
      const programStrength = {};
      const classStrength = {};
      let jcwStrength = 0;

      baseData.forEach(r2 => {
        const cls = r2[baseClassIdx];
        const program = r2[baseProgramIdx];
        const pcs = r2[basePCSIdx];

        const catKey = pcs + "|" + cls;
        const category = categoryLookup2[catKey] || "UNKNOWN";

        const pcsKey = cls + "|" + program + "|" + pcs + "|" + category;
        pcsStrength[pcsKey] = (pcsStrength[pcsKey] || 0) + 1;

        const progKey = cls + "|" + program + "|" + category;
        programStrength[progKey] = (programStrength[progKey] || 0) + 1;

        classStrength[cls] = (classStrength[cls] || 0) + 1;

        jcwStrength++;
      });

      Object.keys(pcsStrength).forEach(key => {
        const parts = key.split("|");
        const cls = parts[0];
        const program = parts[1];
        const pcs = parts[2];
        const category = parts[3];
        const strength = pcsStrength[key];

        const progKey = cls + "|" + program + "|" + category;
        extRows.push({
          class: cls, program: program, category: category, section: pcs, strength: strength,
          programWise: programStrength[progKey], classWise: classStrength[cls], jcwWise: jcwStrength
        });
      });

      extRows.sort((a, b) => {
        const c1 = a.class.localeCompare(b.class);
        if (c1 !== 0) return c1;
        const c2 = a.program.localeCompare(b.program);
        if (c2 !== 0) return c2;
        const catA = categoryOrder[a.category] || 999;
        const catB = categoryOrder[b.category] || 999;
        if (catA !== catB) return catA - catB;
        return a.section.localeCompare(b.section);
      });

      sheetOut.getRange(r, 1, 1, tblHeaders.length)
        .setValues([tblHeaders]).setFontWeight("bold").setBackground("#f1c232")
        .setHorizontalAlignment("center").setVerticalAlignment("middle");
      r++;

      const startRow = r;
      const outRows = [tblHeaders];

      extRows.forEach(row => {
        const rowArr = [row.class, row.program, row.category, row.section, row.strength, row.programWise, row.classWise, row.jcwWise];
        sheetOut.getRange(r, 1, 1, tblHeaders.length).setValues([rowArr]);
        outRows.push(rowArr);
        r++;
      });

      ["Class", "Program", "Category", "Program-wise Strength", "Class-wise Strength", "JCW Strength"].forEach(col => {
        const colIndex = tblHeaders.indexOf(col) + 1;
        let mergeStart = startRow;

        if (col === "Category") {
          let prevProgram = sheetOut.getRange(startRow, tblHeaders.indexOf("Program") + 1).getValue();
          let prevVal = sheetOut.getRange(startRow, colIndex).getValue();

          for (let rr = startRow + 1; rr <= r; rr++) {
            let currProgram = rr <= r - 1 ? sheetOut.getRange(rr, tblHeaders.indexOf("Program") + 1).getValue() : "__END__";
            let currVal = rr <= r - 1 ? sheetOut.getRange(rr, colIndex).getValue() : "__END__";
            if (!currProgram) currProgram = prevProgram;

            if (currProgram !== prevProgram || currVal !== prevVal || rr === r) {
              const mergeEnd = rr - 1;
              if (mergeEnd - mergeStart + 1 > 1) sheetOut.getRange(mergeStart, colIndex, mergeEnd - mergeStart + 1, 1).merge();
              mergeStart = rr;
              prevProgram = currProgram;
              prevVal = currVal;
            }
          }
        } else {
          let prevVal = sheetOut.getRange(startRow, colIndex).getValue();
          for (let rr = startRow + 1; rr <= r; rr++) {
            const curr = rr <= r - 1 ? sheetOut.getRange(rr, colIndex).getValue() : "__END__";
            if (curr !== prevVal || rr === r) {
              if (rr - mergeStart > 1) sheetOut.getRange(mergeStart, colIndex, rr - mergeStart, 1).merge();
              mergeStart = rr;
              prevVal = curr;
            }
          }
        }
      });

      sheetOut.getRange(startRow - 1, 1, r - startRow + 1, tblHeaders.length)
        .setBorder(true, true, true, true, true, true)
        .setHorizontalAlignment("center").setVerticalAlignment("middle");

      r += 2;
      return outRows;
    }

    function writeSeatIncreaseDecreaseSummary() {
      const tblHeaders = [
        "Category", "Increase/Decrease Seats", SESSION_FIRST_YEAR_LABEL, SESSION_SECOND_YEAR_LABEL
      ];

      // Look up a single program+category total for a given class
      const categoryTotal = (cls, program, category) => {
        const match = classCategoryTable.find(row => row.class === cls && row.program === program && row.category === category);
        return match ? match.total : 0;
      };
      // Build the "172+76+47=295" breakdown for the combined SUB row, for a given class
      const combinedSubBreakdown = (cls) => {
        const parts = COMBINED_SUB_PROGRAMS.map(program => categoryTotal(cls, program, 'SUBSIDISIZED'));
        const total = parts.reduce((sum, v) => sum + v, 0);
        return { total, display: parts.join('+') + '=' + total };
      };

      const firstYearSub = combinedSubBreakdown(FIRST_YEAR_CLASS_LABEL);
      const secondYearSub = combinedSubBreakdown(SECOND_YEAR_CLASS_LABEL);

      const summaryRows = [
        {
          label: "FSC-I (M) SUB (PRE-MEDICAL+FCS+PRE-ENGINEERING)",
          firstYear: firstYearSub.total,
          secondYear: secondYearSub.total,
          firstYearDisplay: firstYearSub.display,
          secondYearDisplay: secondYearSub.display
        },
        {
          label: "FSC-I (M) NON-SUB (PRE-MEDICAL)",
          firstYear: categoryTotal(FIRST_YEAR_CLASS_LABEL, MERIT_QUOTA_SPLIT_PROGRAM, 'NON-SUBSIDIZED (M)'),
          secondYear: categoryTotal(SECOND_YEAR_CLASS_LABEL, MERIT_QUOTA_SPLIT_PROGRAM, 'NON-SUBSIDIZED (M)')
        },
        {
          label: "FSC-I (E) NON-SUB (PRE-MEDICAL)",
          firstYear: categoryTotal(FIRST_YEAR_CLASS_LABEL, MERIT_QUOTA_SPLIT_PROGRAM, 'NON-SUBSIDIZED (E)'),
          secondYear: categoryTotal(SECOND_YEAR_CLASS_LABEL, MERIT_QUOTA_SPLIT_PROGRAM, 'NON-SUBSIDIZED (E)')
        },
        {
          label: "F.A",
          firstYear: categoryTotal(FIRST_YEAR_CLASS_LABEL, 'F.A (JCW)', 'SUBSIDISIZED'),
          secondYear: categoryTotal(SECOND_YEAR_CLASS_LABEL, 'F.A (JCW)', 'SUBSIDISIZED')
        }
      ];

      sheetOut.getRange(r, 1, 1, tblHeaders.length)
        .setValues([tblHeaders]).setFontWeight("bold").setBackground("#f1c232")
        .setHorizontalAlignment("center").setVerticalAlignment("middle");
      r++;

      const startRow = r;
      const outRows = [tblHeaders];
      let totalDiff = 0;

      summaryRows.forEach(row => {
        const diff = row.firstYear - row.secondYear;
        totalDiff += diff;
        const firstYearCell = row.firstYearDisplay || row.firstYear;
        const secondYearCell = row.secondYearDisplay || row.secondYear;
        const rowArr = [row.label, diff, firstYearCell, secondYearCell];
        sheetOut.getRange(r, 1, 1, tblHeaders.length).setValues([rowArr]);
        outRows.push(rowArr);
        r++;
      });

      const totalRowArr = ["INCREASED SEATS", totalDiff, "", ""];
      sheetOut.getRange(r, 1, 1, tblHeaders.length).setValues([totalRowArr]).setFontWeight("bold");
      outRows.push(totalRowArr);
      r++;

      sheetOut.getRange(startRow - 1, 1, r - startRow + 1, tblHeaders.length)
        .setBorder(true, true, true, true, true, true)
        .setHorizontalAlignment("center").setVerticalAlignment("middle");

      r += 2;
      return outRows;
    }

    const table1Rows = writeExtendedTable();
    const table2Rows = writeExtendedTablePCSWithCategoryLookup();
    const seatSummaryRows = writeSeatIncreaseDecreaseSummary();

    let nextRow = r;
    const yearWiseResult = writeProgramSummaryYearWiseWithMergedYearTotal(ss, sheetOut, nextRow);

    SpreadsheetApp.flush();

    return {
      success: true,
      sheetName: 'Fancy Report Fancy',
      sheetUrl: ss.getUrl() + '#gid=' + sheetOut.getSheetId(),
      totalStudents: data.length,
      tables: {
        classCategory: table1Rows,
        pcsSection: table2Rows,
        seatIncreaseDecrease: seatSummaryRows,
        yearWise: yearWiseResult.rows
      }
    };
  }







  function writeProgramSummaryYearWiseWithMergedYearTotal(ss, sheetOut, startRow) {
    const sheet = ss.getSheetByName('Transform-ALL-COLLEGE-STRENGTH');
    const data = sheet.getDataRange().getValues();

    const headers = data.shift();
    const programCol = headers.indexOf('Program');
    const yearCol = headers.indexOf('Class');
    const quotaCol = headers.indexOf('Quota');

    const summary = {};
    const yearTotals = {};
    let grandTotalOpen = 0;
    let grandTotalQuota = 0;

    data.forEach(row => {
      const program = row[programCol];
      const year = row[yearCol];
      const quotaRaw = String(row[quotaCol] || '').trim();
      const isOpen = quotaRaw === '' || quotaRaw.toLowerCase().includes('open quota');

      if (!summary[year]) summary[year] = {};
      if (!summary[year][program]) summary[year][program] = { open: 0, quota: 0, breakdown: {} };

      if (isOpen) {
        summary[year][program].open++;
        grandTotalOpen++;
      } else {
        summary[year][program].quota++;
        grandTotalQuota++;
        summary[year][program].breakdown[quotaRaw] = (summary[year][program].breakdown[quotaRaw] || 0) + 1;
      }
    });

    Object.keys(summary).forEach(year => {
      let total = 0;
      Object.keys(summary[year]).forEach(program => {
        const info = summary[year][program];
        total += info.open + info.quota;
      });
      yearTotals[year] = total;
    });

    const output = [];
    const outHeaders = [
      'Class', 'Program', 'Quota', 'Count', 'Year Quota', 'Year Open', 'Year Strength (open+quota)', 'Year Total'
    ];
    output.push(outHeaders);

    const yearOrder = { '1st Year': 1, '2nd Year': 2 };

    Object.keys(summary)
      .sort((a, b) => (yearOrder[a] || 99) - (yearOrder[b] || 99))
      .forEach(year => {
        const programs = summary[year];
        Object.keys(programs).sort().forEach(program => {
          const info = programs[program];
          if (Object.keys(info.breakdown).length === 0) {
            output.push([year, program, '-', '-', info.open, info.quota, info.open + info.quota, yearTotals[year]]);
          } else {
            Object.entries(info.breakdown).forEach(([quotaName, count]) => {
              output.push([year, program, quotaName, count, info.quota, info.open, info.open + info.quota, yearTotals[year]]);
            });
          }
        });
      });

    const totalGrand = grandTotalOpen + grandTotalQuota;
    output.push(['JCW Total:', '', '', '', grandTotalOpen, grandTotalQuota, '', totalGrand]);

    const range = sheetOut.getRange(startRow, 1, output.length, outHeaders.length);
    range.setValues(output);

    const numRows = output.length;

    const mergeColumns = ['Class', 'Year Open', 'Year Quota', 'Year Strength (open+quota)', 'Year Total']
      .map(h => outHeaders.indexOf(h) + 1);

    mergeColumns.forEach(colIndex => {
      let mergeStart = startRow + 1;
      let prevVal = sheetOut.getRange(mergeStart, colIndex).getValue();

      for (let rr = startRow + 2; rr <= startRow + numRows; rr++) {
        const currVal = sheetOut.getRange(rr, colIndex).getValue();
        if (currVal !== prevVal || rr === startRow + numRows) {
          const mergeEnd = rr - 1;
          if (mergeEnd - mergeStart + 1 > 1) sheetOut.getRange(mergeStart, colIndex, mergeEnd - mergeStart + 1, 1).merge();
          mergeStart = rr;
          prevVal = currVal;
        }
      }
    });

    const finalRange = sheetOut.getRange(startRow, 1, numRows, outHeaders.length);
    finalRange.setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID);
    finalRange.setVerticalAlignment('middle').setHorizontalAlignment('center');

    sheetOut.getRange(startRow, 1, 1, outHeaders.length)
      .setFontWeight('bold').setBackground('#f1c232')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');

    sheetOut.getRange(startRow + 1, 2, numRows - 2, 1).setFontWeight('bold');
    sheetOut.getRange(startRow + numRows - 1, 1, 1, outHeaders.length).setFontWeight('bold');
    sheetOut.getRange(startRow + 1, 3, numRows - 2, 1).setWrap(true);
    sheetOut.getRange(startRow + 1, 4, numRows - 2, 1).setWrap(true);

    const quotaColIndex = 3;
    for (let rr = startRow + 1; rr < startRow + numRows - 1; rr++) {
      const cell = sheetOut.getRange(rr, quotaColIndex);
      const value = cell.getValue().toString().trim();
      if (value === 'Move from CHE') {
        cell.setBackground('#d9b3ff');
      }
    }

    return { nextRow: startRow + numRows + 1, rows: output };
  }

  function generateJCWStrengthSummary1() {

    // ============================================================
    // OPEN SPREADSHEET
    // ============================================================

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const sheet = ss.getSheetByName(
      "Transform-ALL-COLLEGE-STRENGTH"
    );

    if (!sheet) {
      throw new Error(
        "Sheet 'Transform-ALL-COLLEGE-STRENGTH' not found"
      );
    }

    const data = sheet.getDataRange().getValues();


    // ============================================================
    // HELPER FUNCTIONS
    // ============================================================

    // HOMIC enrollment numbers contain -HC
    const isHomic = idStr =>
      /-HC\d+/i.test(String(idStr || ""));


    // Normalize program names
    const normProgram = progStr => {

      const p = String(progStr || "").toUpperCase();

      if (
        p.includes("PRE-MEDICAL") ||
        p.includes("PREMEDICAL")
      ) {
        return "PREMEDICAL";
      }

      if (p.includes("PRE-ENGINEERING")) {
        return "PRE-ENGINEERING";
      }

      if (p.includes("FCS")) {
        return "FCS";
      }

      if (p.includes("F.A")) {
        return "F.A";
      }

      return "OTHER";
    };


    // ============================================================
    // PRE-MEDICAL SUBSIDIZED SECTIONS
    // ============================================================

    const SUBSIDIZED_SECTIONS = [
      "A",
      "B",
      "C",
      "D"
    ];


    // ============================================================
    // PRE-MEDICAL NON-SUBSIDIZED SECTIONS
    //
    // IMPORTANT:
    //
    // 1ST YEAR:
    //   Morning = E-R
    //   Evening = S-T
    //
    // 2ND YEAR:
    //   Morning = E-K
    //   Evening = L-P
    // ============================================================


    // ----------------------------
    // 1ST YEAR
    // ----------------------------

    const NONSUB_M_SECTIONS_1ST = [
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

    const NONSUB_E_SECTIONS_1ST = [
      "S",
      "T"
    ];


    // ----------------------------
    // 2ND YEAR
    // ----------------------------

    const NONSUB_M_SECTIONS_2ND = [
      "E",
      "F",
      "G",
      "H",
      "I",
      "J",
      "K"
    ];

    const NONSUB_E_SECTIONS_2ND = [
      "L",
      "M",
      "N",
      "O",
      "P"
    ];


    // ============================================================
    // RETRY FUNCTION
    // ============================================================

    const withRetry = (
      fn,
      attempts = 3,
      delayMs = 2000
    ) => {

      for (let i = 1; i <= attempts; i++) {

        try {

          return fn();

        } catch (e) {

          if (i === attempts) {
            throw e;
          }

          Logger.log(
            "Retrying after error: " +
            e.message +
            " (attempt " +
            i +
            ")"
          );

          Utilities.sleep(delayMs);
        }
      }
    };


    // ============================================================
    // READ MASTER DATA
    // ============================================================

    const records = [];

    data.forEach(row => {

      // Column G = Class
      const year = String(row[6] || "").trim();

      // Only 1st Year and 2nd Year
      if (
        year !== "1st Year" &&
        year !== "2nd Year"
      ) {
        return;
      }

      records.push({

        // Column B = Enroll#
        homic: isHomic(row[1]),

        // Column F = Program
        program: normProgram(row[5]),

        // Column G = Class
        year: year,

        // Column H = Section
        section: String(row[7] || "")
          .trim()
          .toUpperCase()
      });
    });


    // ============================================================
    // BASIC TOTALS
    // ============================================================

    const total = records.length;

    const homicRecords = records.filter(
      r => r.homic
    );

    const jcwOnlyRecords = records.filter(
      r => !r.homic
    );

    const homicCount = homicRecords.length;

    const homicPremed = homicRecords.filter(
      r => r.program === "PREMEDICAL"
    ).length;

    const homicFcs = homicRecords.filter(
      r => r.program === "FCS"
    ).length;


    // ============================================================
    // JCW ONLY YEAR-WISE TOTALS
    // ============================================================

    const jcwOnly1st = jcwOnlyRecords.filter(
      r => r.year === "1st Year"
    ).length;

    const jcwOnly2nd = jcwOnlyRecords.filter(
      r => r.year === "2nd Year"
    ).length;


    // ============================================================
    // GENERAL COUNT FUNCTION
    // ============================================================

    const count = (
      year,
      program,
      opts = {}
    ) => {

      return records.filter(r => {

        if (r.year !== year) {
          return false;
        }

        if (r.program !== program) {
          return false;
        }

        if (
          opts.excludeHomic &&
          r.homic
        ) {
          return false;
        }

        if (
          opts.section &&
          !opts.section.includes(r.section)
        ) {
          return false;
        }

        return true;

      }).length;
    };


    // ============================================================
    // PROGRAM + YEAR BREAKDOWN
    // ============================================================

    const programYearBreakdown = program => ({

      // All 1st Year
      y1: count(
        "1st Year",
        program
      ),

      // 2nd Year HOMIC
      y2Homic:
        count(
          "2nd Year",
          program
        )
        -
        count(
          "2nd Year",
          program,
          {
            excludeHomic: true
          }
        ),

      // 2nd Year JCW
      y2Jcw:
        count(
          "2nd Year",
          program,
          {
            excludeHomic: true
          }
        )
    });


    // ============================================================
    // PROGRAM BREAKDOWNS
    // ============================================================

    const FA =
      programYearBreakdown("F.A");

    const FCS =
      programYearBreakdown("FCS");

    const PREMEDICAL =
      programYearBreakdown("PREMEDICAL");

    const PRE_ENG =
      programYearBreakdown(
        "PRE-ENGINEERING"
      );


    // ============================================================
    // PRE-MEDICAL SUBSIDIZED
    //
    // Sections A-D for BOTH years
    // ============================================================

    const SUBSIDIZED = {

      // 1st Year
      y1:
        count(
          "1st Year",
          "PREMEDICAL",
          {
            section: SUBSIDIZED_SECTIONS
          }
        ),

      // 2nd Year HOMIC
      y2Homic:
        count(
          "2nd Year",
          "PREMEDICAL",
          {
            section: SUBSIDIZED_SECTIONS
          }
        )
        -
        count(
          "2nd Year",
          "PREMEDICAL",
          {
            section: SUBSIDIZED_SECTIONS,
            excludeHomic: true
          }
        ),

      // 2nd Year JCW
      y2Jcw:
        count(
          "2nd Year",
          "PREMEDICAL",
          {
            section: SUBSIDIZED_SECTIONS,
            excludeHomic: true
          }
        )
    };


    // ============================================================
    // PRE-MEDICAL NON-SUBSIDIZED MORNING
    //
    // 1st Year = E-R
    // 2nd Year = E-K
    // ============================================================

    const NONSUB_M = {

      // 1st Year: E-R
      y1:
        count(
          "1st Year",
          "PREMEDICAL",
          {
            section: NONSUB_M_SECTIONS_1ST
          }
        ),

      // 2nd Year: E-K
      y2:
        count(
          "2nd Year",
          "PREMEDICAL",
          {
            section: NONSUB_M_SECTIONS_2ND
          }
        )
    };


    // ============================================================
    // PRE-MEDICAL NON-SUBSIDIZED EVENING
    //
    // 1st Year = S-T
    // 2nd Year = L-P
    // ============================================================

    const NONSUB_E = {

      // 1st Year: S-T
      y1:
        count(
          "1st Year",
          "PREMEDICAL",
          {
            section: NONSUB_E_SECTIONS_1ST
          }
        ),

      // 2nd Year: L-P
      y2:
        count(
          "2nd Year",
          "PREMEDICAL",
          {
            section: NONSUB_E_SECTIONS_2ND
          }
        )
    };


    // ============================================================
    // CURRENT SESSION — 1ST YEAR
    // ============================================================

    const curFscISub =
      SUBSIDIZED.y1 +
      FCS.y1 +
      PRE_ENG.y1;

    const curFscINonSubM =
      NONSUB_M.y1;

    const curFscINonSubE =
      NONSUB_E.y1;

    const curFa =
      FA.y1;


    // ============================================================
    // PREVIOUS SESSION — 2ND YEAR JCW
    // ============================================================

    const prevFscISub =
      SUBSIDIZED.y2Jcw +
      FCS.y2Jcw +
      PRE_ENG.y2Jcw;

    const prevFscINonSubM =
      NONSUB_M.y2;

    const prevFscINonSubE =
      NONSUB_E.y2;

    const prevFa =
      FA.y2Jcw;


    // ============================================================
    // JCW TOTALS
    // ============================================================

    const jcwOnlyTotal =
      jcwOnly1st +
      jcwOnly2nd;

    const jcwTotalStrength =
      total;


    // ============================================================
    // OUTPUT GRID
    // ============================================================

    const GRID_ROWS = 24;
    const GRID_COLS = 8;

    const grid =
      Array.from(
        {
          length: GRID_ROWS
        },
        () =>
          Array(GRID_COLS).fill("")
      );

    const boldGrid =
      Array.from(
        {
          length: GRID_ROWS
        },
        () =>
          Array(GRID_COLS).fill(false)
      );


    // ============================================================
    // GRID HELPERS
    // ============================================================

    const put = (
      r,
      c,
      value,
      isBold = false
    ) => {

      grid[r - 1][c - 1] = value;

      if (isBold) {
        boldGrid[r - 1][c - 1] = true;
      }
    };


    const putRow = (
      r,
      cStart,
      vals,
      isBold = false
    ) => {

      vals.forEach(
        (v, i) => {

          put(
            r,
            cStart + i,
            v,
            isBold
          );

        }
      );
    };


    // ============================================================
    // SECTION 1 — TOTAL STRENGTH
    // ============================================================

    putRow(
      1,
      1,
      [
        "JCW TOTAL STRENGTH",
        ""
      ],
      true
    );

    put(
      1,
      3,
      jcwTotalStrength,
      true
    );


    // ============================================================
    // SECTION 2 — YEAR-WISE TOTAL
    // ============================================================

    putRow(
      3,
      1,
      [
        "1ST YEAR",
        ""
      ],
      true
    );

    put(
      3,
      3,
      jcwOnly1st
    );

    put(
      3,
      4,
      "2ND YEAR",
      true
    );

    put(
      3,
      5,
      jcwOnly2nd
    );


    // ============================================================
    // JCW ONLY / HOMIC
    // ============================================================

    putRow(
      4,
      1,
      [
        "JCW ONLY",
        ""
      ],
      true
    );

    put(
      4,
      3,
      jcwOnlyTotal
    );

    put(
      4,
      5,
      "HOMIC",
      true
    );

    put(
      4,
      6,
      homicCount
    );


    // HOMIC PRE-MEDICAL
    put(
      5,
      6,
      "PRE-ME"
    );

    put(
      5,
      7,
      homicPremed
    );


    // HOMIC FCS
    put(
      6,
      6,
      "FCS"
    );

    put(
      6,
      7,
      homicFcs
    );


    // ============================================================
    // PROGRAM BREAKDOWN HEADINGS
    // ============================================================

    put(
      9,
      3,
      "1ST YEAR",
      true
    );

    put(
      9,
      4,
      "2ND YEAR",
      true
    );


    // ============================================================
    // F.A
    // ============================================================

    putRow(
      10,
      1,
      [
        "F.A",
        ""
      ],
      true
    );

    put(
      10,
      3,
      FA.y1
    );

    put(
      10,
      4,
      FA.y2Homic +
      FA.y2Jcw
    );


    // ============================================================
    // FCS
    // ============================================================

    putRow(
      11,
      1,
      [
        "FCS",
        ""
      ],
      true
    );

    put(
      11,
      3,
      FCS.y1
    );

    put(
      11,
      4,
      FCS.y2Homic +
      FCS.y2Jcw
    );

    put(
      11,
      5,
      "HOMIC"
    );

    put(
      11,
      6,
      FCS.y2Homic
    );

    put(
      11,
      7,
      "JCW"
    );

    put(
      11,
      8,
      FCS.y2Jcw
    );


    // ============================================================
    // PRE-MEDICAL
    // ============================================================

    putRow(
      12,
      1,
      [
        "PREMEDICAL",
        ""
      ],
      true
    );

    put(
      12,
      3,
      PREMEDICAL.y1,
      true
    );

    put(
      12,
      4,
      PREMEDICAL.y2Homic +
      PREMEDICAL.y2Jcw,
      true
    );

    put(
      12,
      5,
      "HOMIC"
    );

    put(
      12,
      6,
      PREMEDICAL.y2Homic
    );

    put(
      12,
      7,
      "JCW"
    );

    put(
      12,
      8,
      PREMEDICAL.y2Jcw
    );


    // ============================================================
    // PRE-MEDICAL SUBSIDIZED
    // ============================================================

    put(
      13,
      2,
      "SUBIDIZED"
    );

    put(
      13,
      3,
      SUBSIDIZED.y1
    );

    put(
      13,
      4,
      SUBSIDIZED.y2Homic +
      SUBSIDIZED.y2Jcw
    );

    put(
      13,
      5,
      "HOMIC"
    );

    put(
      13,
      6,
      SUBSIDIZED.y2Homic
    );

    put(
      13,
      7,
      "JCW"
    );

    put(
      13,
      8,
      SUBSIDIZED.y2Jcw
    );


    // ============================================================
    // PRE-MEDICAL NON-SUBSIDIZED MORNING
    // ============================================================

    put(
      14,
      2,
      "NON-SUBSIDIZED (M)"
    );

    put(
      14,
      3,
      NONSUB_M.y1
    );

    put(
      14,
      4,
      NONSUB_M.y2
    );


    // ============================================================
    // PRE-MEDICAL NON-SUBSIDIZED EVENING
    // ============================================================

    put(
      15,
      2,
      "NON-SUBSIDIZED (E)"
    );

    put(
      15,
      3,
      NONSUB_E.y1
    );

    put(
      15,
      4,
      NONSUB_E.y2
    );


    // ============================================================
    // PRE-ENGINEERING
    // ============================================================

    putRow(
      16,
      1,
      [
        "PRE-ENGINEERING",
        ""
      ],
      true
    );

    put(
      16,
      3,
      PRE_ENG.y1,
      true
    );

    put(
      16,
      4,
      PRE_ENG.y2Homic +
      PRE_ENG.y2Jcw,
      true
    );


    // ============================================================
    // INCREASE / DECREASE SECTION
    // ============================================================

    put(
      19,
      2,
      "INCREASE /DECREASE",
      true
    );

    put(
      19,
      3,
      "SESSION (" + CMP_CONFIG.CURRENT_SESSION + ")",
      true
    );

    put(
      19,
      4,
      "SESSION (" + CMP_CONFIG.PREVIOUS_SESSION + ")",
      true
    );


    // ============================================================
    // FSC-I MORNING SUBSIDIZED
    // ============================================================

    put(
      20,
      1,
      "FSC-I (M) SUB"
    );

    put(
      20,
      2,
      curFscISub -
      prevFscISub
    );

    put(
      20,
      3,
      curFscISub
    );

    put(
      20,
      4,
      prevFscISub
    );


    // ============================================================
    // FSC-I MORNING NON-SUBSIDIZED
    // ============================================================

    put(
      21,
      1,
      "FSC-I (M) NON-SUB"
    );

    put(
      21,
      2,
      curFscINonSubM -
      prevFscINonSubM
    );

    put(
      21,
      3,
      curFscINonSubM
    );

    put(
      21,
      4,
      prevFscINonSubM
    );


    // ============================================================
    // FSC-I EVENING NON-SUBSIDIZED
    // ============================================================

    put(
      22,
      1,
      "FSC-I (E) NON-SUB"
    );

    put(
      22,
      2,
      curFscINonSubE -
      prevFscINonSubE
    );

    put(
      22,
      3,
      curFscINonSubE
    );

    put(
      22,
      4,
      prevFscINonSubE
    );


    // ============================================================
    // F.A INCREASE / DECREASE
    // ============================================================

    put(
      23,
      1,
      "F.A"
    );

    put(
      23,
      2,
      curFa -
      prevFa
    );

    put(
      23,
      3,
      curFa
    );

    put(
      23,
      4,
      prevFa
    );


    // ============================================================
    // TOTAL INCREASED SEATS
    // ============================================================

    const increasedSeats =
      (curFscISub - prevFscISub) +
      (curFscINonSubM - prevFscINonSubM) +
      (curFscINonSubE - prevFscINonSubE) +
      (curFa - prevFa);


    put(
      24,
      1,
      "INCREASED SEATS",
      true
    );

    put(
      24,
      2,
      increasedSeats,
      true
    );


    // ============================================================
    // CREATE / CLEAR OUTPUT SHEET
    // ============================================================

    const outName =
      "JCW Strength Summary";

    let out =
      ss.getSheetByName(outName);


    if (!out) {

      out =
        withRetry(
          () => ss.insertSheet(outName)
        );

    } else {

      withRetry(
        () => out.clear()
      );
    }


    // ============================================================
    // WRITE GRID
    // ============================================================

    const range =
      out.getRange(
        1,
        1,
        GRID_ROWS,
        GRID_COLS
      );

    withRetry(
      () => range.setValues(grid)
    );


    // ============================================================
    // APPLY BOLD / NORMAL FONT
    // ============================================================

    withRetry(
      () =>
        range.setFontWeights(
          boldGrid.map(
            r =>
              r.map(
                b =>
                  b
                    ? "bold"
                    : "normal"
              )
          )
        )
    );


    // ============================================================
    // MERGES
    // ============================================================

    const merges = [

      [1, 1, 1, 2],

      [3, 1, 3, 2],

      [3, 5, 3, 6],

      [4, 1, 4, 2],

      [9, 4, 9, 8],

      [10, 1, 10, 2],

      [10, 4, 10, 8],

      [11, 1, 11, 2],

      [12, 1, 12, 2],

      [14, 4, 14, 8],

      [15, 4, 15, 8],

      [16, 1, 16, 2],

      [16, 4, 16, 8]
    ];


    merges.forEach(
      ([r1, c1, r2, c2]) => {

        withRetry(
          () =>
            out
              .getRange(
                r1,
                c1,
                r2 - r1 + 1,
                c2 - c1 + 1
              )
              .merge()
              .setHorizontalAlignment(
                "center"
              )
        );

      }
    );


    // ============================================================
    // BORDER BLOCKS
    // ============================================================

    const borderBlocks = [

      [1, 1, 1, 3],

      [3, 1, 6, 8],

      [9, 1, 16, 8],

      [19, 1, 24, 4]
    ];


    borderBlocks.forEach(
      ([r1, c1, r2, c2]) => {

        withRetry(
          () =>
            out
              .getRange(
                r1,
                c1,
                r2 - r1 + 1,
                c2 - c1 + 1
              )
              .setBorder(
                true,
                true,
                true,
                true,
                true,
                true
              )
        );

      }
    );


    // ============================================================
    // TOTAL STRENGTH HIGHLIGHT
    // ============================================================

    withRetry(
      () =>
        out
          .getRange(1, 3)
          .setBackground("#d9d9d9")
          .setFontSize(14)
    );


    // ============================================================
    // HIGHLIGHT IMPORTANT CELLS
    // ============================================================

    const HIGHLIGHT_COLOR =
      "#ead1dc";

    const highlightCells = [

      [11, 3],

      [11, 8],

      [13, 3],

      [13, 8],

      [16, 3]
    ];


    highlightCells.forEach(
      ([r, c]) => {

        withRetry(
          () =>
            out
              .getRange(r, c)
              .setBackground(
                HIGHLIGHT_COLOR
              )
        );

      }
    );


    // Highlight row 16, columns D-H
    withRetry(
      () =>
        out
          .getRange(
            16,
            4,
            1,
            5
          )
          .setBackground(
            HIGHLIGHT_COLOR
          )
    );


    // ============================================================
    // AUTO RESIZE
    // ============================================================

    withRetry(
      () =>
        out.autoResizeColumns(
          1,
          GRID_COLS
        )
    );


    // ============================================================
    // FLUSH
    // ============================================================

    SpreadsheetApp.flush();


    // ============================================================
    // RETURN RESULT
    // ============================================================

    return {

      success: true,

      sheetName: outName,

      sheetUrl:
        ss.getUrl() +
        "#gid=" +
        out.getSheetId(),

      rowsCount:
        GRID_ROWS,

      colsCount:
        GRID_COLS,

      grid:
        grid,

      bold:
        boldGrid,

      merges:
        merges,

      highlights:
        highlightCells.concat(
          [[16, 4]]
        ),

      bigCells:
        [[1, 3]]
    };
  }


  function generateSideBySideEnrollmentReport1() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const src = ss.getSheetByName("Transform-ALL-COLLEGE-STRENGTH");
    if (!src) throw new Error("Sheet 'Transform-ALL-COLLEGE-STRENGTH' not found!");

    const data = src.getDataRange().getValues();
    const headers = data.shift();

    const idxClass = headers.indexOf("Class");

    const yearGroups = {};
    data.forEach(row => {
      const classNameRaw = row[idxClass];
      if (!classNameRaw) return;

      let year = 'Unknown';
      if (classNameRaw.match(/1/i)) year = '1st Year';
      if (classNameRaw.match(/2/i)) year = '2nd Year';

      if (!yearGroups[year]) yearGroups[year] = {};
      if (!yearGroups[year][classNameRaw]) yearGroups[year][classNameRaw] = 0;

      yearGroups[year][classNameRaw]++;
    });

    let out = ss.getSheetByName("JCW Enrollment Report");
    if (!out) out = ss.insertSheet("JCW Enrollment Report");
    out.clear();

    out.getRange(1, 1, 1, 5).setValue("JCW STUDENTS' ENROLLMENT")
      .setFontWeight("bold")
      .setFontSize(14)
      .merge();

    out.getRange(2, 1, 1, 5).setValues([["S.No.", "Class", "Number of Students", "Class", "Number of Students"]])
      .setFontWeight("bold")
      .setBackground("#cccccc");

    const classes1 = Object.entries(yearGroups['1st Year'] || {});
    const classes2 = Object.entries(yearGroups['2nd Year'] || {});

    const maxRows = Math.max(classes1.length, classes2.length);
    const rows = [];

    let total1 = 0;
    let total2 = 0;

    for (let i = 0; i < maxRows; i++) {
      const row1 = classes1[i] || ["", ""];
      const row2 = classes2[i] || ["", ""];

      total1 += row1[1] || 0;
      total2 += row2[1] || 0;

      rows.push([i + 1, row1[0], row1[1], row2[0], row2[1]]);
    }

    out.getRange(3, 1, rows.length, 5).setValues(rows);

    const totalRowIndex = 3 + rows.length;
    out.getRange(totalRowIndex, 2, 1, 1).setValue("TOTAL").setFontWeight("bold");
    out.getRange(totalRowIndex, 3, 1, 1).setValue(total1).setFontWeight("bold");
    out.getRange(totalRowIndex, 4, 1, 1).setValue("TOTAL").setFontWeight("bold");
    out.getRange(totalRowIndex, 5, 1, 1).setValue(total2).setFontWeight("bold");

    const grandTotalRowIndex = 4 + rows.length;
    out.getRange(grandTotalRowIndex, 4, 1, 2).merge()
      .setValue(`TOTAL STRENGTH = ${total1 + total2}`).setFontWeight("bold");

    out.autoResizeColumns(1, 5);
    SpreadsheetApp.flush();

    // ---- Build grid/merge data for the web page ----
    const totalDataRows = 4 + rows.length; // title + header + data rows + total row + grand total row
    const grid = [];
    const bold = [];

    grid.push(["JCW STUDENTS' ENROLLMENT", "", "", "", ""]);
    bold.push([true, true, true, true, true]);

    grid.push(["S.No.", "Class", "Number of Students", "Class", "Number of Students"]);
    bold.push([true, true, true, true, true]);

    rows.forEach(r => {
      grid.push(r.map(v => (v === undefined ? "" : v)));
      bold.push([false, false, false, false, false]);
    });

    grid.push(["", "TOTAL", total1, "TOTAL", total2]);
    bold.push([false, true, true, true, true]);

    grid.push(["", "", "", `TOTAL STRENGTH = ${total1 + total2}`, ""]);
    bold.push([false, false, false, true, false]);

    const merges = [
      [1, 1, 1, 5],                                   // title row
      [grid.length, 4, grid.length, 5]                // grand total row (D:E)
    ];

    return {
      success: true,
      sheetName: "JCW Enrollment Report",
      sheetUrl: ss.getUrl() + '#gid=' + out.getSheetId(),
      rowsCount: grid.length,
      colsCount: 5,
      grid: grid,
      bold: bold,
      merges: merges,
      highlights: [],
      bigCells: []
    };
  }



  /* ======================================================================
     REPORT 7: ATTENDANCE AVAILABILITY (FROM DRIVE BACKUP FILES)
     ====================================================================== */

  function generateAttendanceAvailabilityFromBackups1() {
    const folderUrl = 'https://drive.google.com/drive/folders/1FNJz2Tkn1tZd6umoByvsVv8-sAYDxEaF';
    const FOLDER_ID = folderUrl.match(/[-\w]{25,}/)[0];
    const folder = DriveApp.getFolderById(FOLDER_ID);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sh = ss.getSheetByName('ATTENDANCE AVAILABILITY');
    if (sh) ss.deleteSheet(sh);

    sh = ss.insertSheet('ATTENDANCE AVAILABILITY');
    sh.getRange('A1').setValue('ATTENDANCE AVAILABILITY SUMMARY');
    sh.getRange('A2:I2').setValues([[
      'SNo', 'Teacher', 'Subject', 'Class-Sec',
      'September', 'October', 'November', 'December', 'January'
    ]]);

    const monthIndexMap = {
      September: 4,
      October: 5,
      November: 6,
      December: 7,
      January: 8
    };

    const attendanceMap = getAttendanceMap(); // must exist elsewhere in this project

    const sourceFiles = [];
    const it = folder.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      if (f.getMimeType() === MimeType.GOOGLE_SHEETS &&
          f.getName().toLowerCase().startsWith('backup-')) {
        sourceFiles.push({
          name: f.getName(),
          file: SpreadsheetApp.open(f)
        });
      }
    }

    const availabilityMap = Object.create(null);

    for (const obj of sourceFiles) {
      const cleanName = obj.name
        .toLowerCase()
        .replace(/[-_]/g, ' ')
        .replace(/\s+/g, ' ');

      const m = cleanName.match(/(january|september|october|november|december)/);
      if (!m) continue;

      const month = m[1].charAt(0).toUpperCase() + m[1].slice(1);
      const file = obj.file;

      for (const program in attendanceMap) {
        const years = attendanceMap[program];

        for (const year in years) {
          const sheets = years[year];

          for (let s = 0; s < sheets.length; s++) {
            const sheet = file.getSheetByName(sheets[s]);
            if (!sheet) continue;

            const lastRow = sheet.getLastRow();
            const lastCol = sheet.getLastColumn();
            if (lastRow < 5 || lastCol < 6) continue;

            const classSec = sheet.getName()
              .replace(/\s*Optional Subjects\s*$/i, '')
              .trim();

            const teachers = sheet.getRange(2, 6, 1, lastCol - 5).getValues()[0];
            const subjects = sheet.getRange(3, 6, 1, lastCol - 5).getValues()[0];
            const data = sheet.getRange(5, 6, lastRow - 4, lastCol - 5).getValues();

            for (let i = 0; i < teachers.length; i += 3) {
              const teacher = teachers[i];
              const subject = subjects[i];
              if (!teacher || !subject) continue;

              let present = false;
              for (let r = 0; r < data.length; r++) {
                if ((data[r][i] || 0) || (data[r][i + 1] || 0)) {
                  present = true;
                  break;
                }
              }

              const key = teacher + '|' + subject + '|' + classSec + '|' + month;

              if (present) {
                availabilityMap[key] = '✔';
              } else if (!(key in availabilityMap)) {
                availabilityMap[key] = '✖';
              }
            }
          }
        }
      }
    }

    const rowMap = new Map();
    const monthSubjectClassMap = new Map();

    for (const key in availabilityMap) {
      const [teacher, subject, classSec, month] = key.split('|');
      const subjectClassMonth = subject + '|' + classSec + '|' + month;

      if (availabilityMap[key] === '✔') {
        monthSubjectClassMap.set(subjectClassMonth, true);
      }
    }

    // Build rows WITHOUT assigning SNo yet — placeholder in slot 0
    for (const key in availabilityMap) {
      const [teacher, subject, classSec, month] = key.split('|');
      const rowKey = teacher + '|' + subject + '|' + classSec;
      const subjectClassMonth = subject + '|' + classSec + '|' + month;

      let row = rowMap.get(rowKey);
      if (!row) {
        row = [null, teacher, subject, classSec, '', '', '', '', ''];
        rowMap.set(rowKey, row);
      }

      const colIndex = monthIndexMap[month];

      if (availabilityMap[key] === '✔') {
        row[colIndex] = '✔';
      } else {
        if (!monthSubjectClassMap.has(subjectClassMonth)) {
          row[colIndex] = '✖';
        }
      }
    }

    // Sort by Teacher (column index 1), then assign SNo in that order
    const summaryRows = [...rowMap.values()].sort((a, b) =>
      String(a[1] || '').localeCompare(String(b[1] || ''), undefined, { sensitivity: 'base' })
    );

    summaryRows.forEach((row, idx) => {
      row[0] = idx + 1;
    });


    if (summaryRows.length) {
      sh.getRange(3, 1, summaryRows.length, summaryRows[0].length)
        .setValues(summaryRows)
        .setBorder(true, true, true, true, true, true);
        if (summaryRows.length) {
    sh.getRange(3, 1, summaryRows.length, summaryRows[0].length)
      .setValues(summaryRows)
      .setBorder(true, true, true, true, true, true);

    // Highlight ✖ cells in light pink, matching the web view
    const monthColStart = 5; // column E (September) — matches monthIndexMap values (4-8, 0-based -> 5-9, 1-based)
    const monthColEnd = 9;   // column I (January)

    for (let r = 0; r < summaryRows.length; r++) {
      for (let c = monthColStart - 1; c < monthColEnd; c++) {
        if (summaryRows[r][c] === '✖') {
          sh.getRange(3 + r, c + 1).setBackground('#f8d7e0');
        }
      }
    }
  }
    }

    Logger.log(`Attendance Availability generated: ${summaryRows.length} rows`);

    const outHeaders = ['SNo', 'Teacher', 'Subject', 'Class-Sec', 'September', 'October', 'November', 'December', 'January'];
    const output = [outHeaders, ...summaryRows];

    return {
      success: true,
      sheetName: 'ATTENDANCE AVAILABILITY',
      sheetUrl: ss.getUrl() + '#gid=' + sh.getSheetId(),
      rowCount: summaryRows.length,
      data: output
    };
  }


  // ==================== MODULE 3: ATTENDANCE MAP ====================
  function getAttendanceMap() {
    return {
      'Pre-Medical': {
        '1st Year': ['A1','B1','C1','E1','F1','G1','H1','I1','J1','K1','L1','M1','N1','O1','P1'],
        '2nd Year': ['A2','B2','C2','E2','F2','G2','H2','I2','J2','K2','L2','M2','N2','O2','P2']
      },
      'Pre-Engineering': {
        '1st Year': ['Pre-Eng-1 D1'],
        '2nd Year': ['Pre-Eng-2 D1']
      },
      'FCS': {
        '1st Year': ['FCS-1 D2','FCS-1 D3'],
        '2nd Year': ['FCS-2 D2']
      },
      'FA': {
        '1st Year': ['F.A-1 A','F.A-1 B','F.A-1 Optional Subjects'],
        '2nd Year': ['F.A-2 A','F.A-2 B','F.A-2 Optional Subjects']
      }
    };
  }


  return {
    setSessions: (cur, prev) => { CMP_CONFIG.CURRENT_SESSION = cur; CMP_CONFIG.PREVIOUS_SESSION = prev; },
    columns: () => cmpGetColumns(),
    run: rptRun_,
    xlsx: rptXlsx_,
    xlsxAll: rptXlsxAll_
  };
})();
