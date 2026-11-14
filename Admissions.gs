/**
 * ADMISSIONS module for the CS Attendance backend  (adds the "Admissions" page to the attendance app)
 *
 * New admissions and every change to an old admission are written DIRECTLY into
 *   Transform-ALL-COLLEGE-STRENGTH   (Users workbook 1ARU8zMO2daJUfU_SIIo4baeSOG7qBPuni5TYRIRLBNY)
 * instead of the program sheets of the admissions workbook.
 *
 * The admissions workbook (1wY8hCFzmBhlVxHjLJGLEDzP_LijbSu7aJfJ-jHhbLRI) is only READ, for
 *   - "Quota Settings"   (seats per quota, change them there as before)
 *   - "Program Sections" (sections per program and year)
 *   - Roll Nos already given there, so a number is never handed out twice.
 * Nothing is ever written into it.
 *
 * Install:
 *   1) In the attendance Apps Script project: + (Add a file) > Script > name it "Admissions" > paste this whole file.
 *   2) In doPost, right after these two lines
 *        var user = authenticate_(req.username, req.password);
 *        if (!user) return json_({ ok: false, error: 'BAD_LOGIN' });
 *      add:
 *        if (/^adm_/.test(String(req.action || ''))) return json_(admRoute_(user, req));
 *   3) Deploy > Manage deployments > Edit > Version: New version > Deploy.
 *
 * Who can use it: the admin login, users whose Role in the Users sheet is ADMIN or PRINCIPAL, plus the usernames listed in ADM_CFG.EXTRA_USERS
 * (or in the Script Property ADM_USERS, comma separated) for admission clerks.
 */

var ADM_CFG = {
  BOOK_ID: '1ARU8zMO2daJUfU_SIIo4baeSOG7qBPuni5TYRIRLBNY',          // Users workbook
  SHEET: 'Transform-ALL-COLLEGE-STRENGTH',
  LEFT_SHEET: 'Left Students',
  PORTAL_SHEET: 'Portal Data',                                          // filled by the UOP portal sync
  TIMETABLE_SHEET: 'Timetable',
  LOG_SHEET: 'Admission Log',                                           // created in the Users workbook
  SETTINGS_BOOK_ID: '1wY8hCFzmBhlVxHjLJGLEDzP_LijbSu7aJfJ-jHhbLRI',  // admissions workbook (read only)
  SETTINGS_SHEET: 'Quota Settings',
  SECTIONS_SHEET: 'Program Sections',
  EXTRA_USERS: [],                                                      // e.g. ['clerk@uop.edu.pk']

  SESSION: '2026-27',
  STATUS: 'Enrolled',
  ROLL_PREFIX: 'JCW-F26-',
  ENTRY_YEAR: '1st Year',
  MIGRATION: { YEAR: '2nd Year', QUOTA: 'Open Quota', ROLL_PREFIX: 'JCW-F25-', ROLL_START: 1279, ROLL_END: 1999 },
  MODE: 'warn',          // 'warn' = full quota needs a confirm, 'block' = full quota cannot be saved
  THRESHOLD: 5,          // "only N seats left" note
  OPEN_QUOTA: 'Open Quota',
  TOTAL_NAME: '* TOTAL *',

  COMPULSORY: ['ENGLISH', 'URDU', 'ISLAMIYAT', 'MUTALIA QURAN'],
  SECOND_YEAR_SUBJECTS: [
    { from: 'ISLAMIYAT', to: 'PAK-STUDIES', only: '' },
    { from: 'BOTANY', to: 'ZOOLOGY', only: 'Pre-Medical' }
  ],
  SUBJECT_DEFAULTS: {
    'Pre-Medical': [{ name: 'Pre-Medical', subjects: ['PHYSICS', 'BOTANY', 'CHEMISTRY'] }],
    'Pre-Engineering': [{ name: 'Pre-Engineering', subjects: ['PHYSICS', 'MATHEMATICS', 'CHEMISTRY'] }],
    'FCS': [{ name: 'FCS', subjects: ['PHYSICS', 'MATHEMATICS', 'COMPUTER SCIENCE'] }],
    'FA': [
      { name: 'ECO/STAT/G.MATH', subjects: ['ECONOMICS', 'STATISTICS', 'GENERAL MATHEMATICS'] },
      { name: 'ECO/ENG ADV/SOCIO', subjects: ['ECONOMICS', 'ENGLISH ADVANCE', 'SOCIOLOGY'] },
      { name: 'PSY/ECO/STAT', subjects: ['PSYCHOLOGY', 'ECONOMICS', 'STATISTICS'] },
      { name: 'PSY/LOG/GEO', subjects: ['PSYCHOLOGY', 'LOGIC', 'GEOGRAPHY'] },
      { name: 'GEO/LOG/SOCIO', subjects: ['GEOGRAPHY', 'LOGIC', 'SOCIOLOGY'] },
      { name: 'URD ADV/ISL ELE/HIS', subjects: ['URDU ADVANCE', 'ISLAMIYAT ELECTIVE', 'HISTORY'] },
      { name: 'PASH/ISL ELE/HIS', subjects: ['PASHTO', 'ISLAMIYAT ELECTIVE', 'HISTORY'] },
      { name: 'URD ADV/HIS/CIV', subjects: ['URDU ADVANCE', 'HISTORY', 'CIVICS'] },
      { name: 'PASH/HIS/CIV', subjects: ['PASHTO', 'HISTORY', 'CIVICS'] },
      { name: 'ARAB/ISL ELE/HIS', subjects: ['ARABIC', 'ISLAMIYAT ELECTIVE', 'HISTORY'] },
      { name: 'ARAB/ISL ELE/CIV', subjects: ['ARABIC', 'ISLAMIYAT ELECTIVE', 'CIVICS'] }
    ]
  },
  // key = name used in Quota Settings / Program Sections;  program = value of the "Program" column in Transform
  PROGRAMS: [
    { key: 'Pre Medical Open',     label: 'Pre-Medical Open (S)',     program: 'Pre-Medical (JCW)',     classProgram: 'Pre-Medical',     appeared: 'FSC (PRE-MEDICAL)',          fee: 'sub' },
    { key: 'Self Morning Medical', label: 'Self Morning Medical (M)', program: 'Pre-Medical (JCW)',     classProgram: 'Pre-Medical',     appeared: 'FSC (PRE-MEDICAL)',          fee: 'non' },
    { key: 'Evening Pre Medical',  label: 'Evening Pre-Medical (E)',  program: 'Pre-Medical (JCW)',     classProgram: 'Pre-Medical',     appeared: 'FSC (PRE-MEDICAL)',          fee: 'non' },
    { key: 'FCS',                  label: 'FCS',                      program: 'FCS (JCW)',             classProgram: 'FCS',             appeared: 'GENERAL / COMPUTER SCIENCE', fee: 'sub' },
    { key: 'FA',                   label: 'F.A',                      program: 'F.A (JCW)',             classProgram: 'FA',              appeared: 'F.A',                        fee: 'sub' },
    { key: 'Pre Engg',             label: 'Pre-Engineering',          program: 'Pre-Engineering (JCW)', classProgram: 'Pre-Engineering', appeared: 'FSC (PRE-ENGINEERING)',      fee: 'sub' }
  ],
  ROLL_RANGES: {
    'Pre Medical Open': [1, 203], 'Self Morning Medical': [204, 1099], 'Evening Pre Medical': [1100, 1999],
    'Pre Engg': [2001, 2099], 'FA': [2101, 2299], 'FCS': [2301, 2399]
  },
  SECTIONS_BY_PROGRAM: {   // used only if "Program Sections" cannot be read
    'Pre Medical Open':     { '1st Year': ['A1', 'B1', 'C1'], '2nd Year': ['A2', 'B2', 'C2'] },
    'Self Morning Medical': { '1st Year': ['E1', 'F1', 'G1', 'H1', 'I1', 'J1', 'K1', 'L1', 'M1', 'N1', 'O1', 'P1', 'Q1', 'R1'],
                              '2nd Year': ['E2', 'F2', 'G2', 'H2', 'I2', 'J2', 'K2', 'L2', 'M2', 'N2', 'O2', 'P2'] },
    'Evening Pre Medical':  { '1st Year': [], '2nd Year': [] },
    'Pre Engg':             { '1st Year': ['Pre-Eng-1 D1'], '2nd Year': ['Pre-Eng-2 D1'] },
    'FCS':                  { '1st Year': ['FCS-1 D2', 'FCS-1 D3'], '2nd Year': ['FCS-2 D2', 'FCS-2 D3'] },
    'FA':                   { '1st Year': ['F.A-1 A', 'F.A-1 B', 'F.A-1 Optional Subjects'], '2nd Year': ['F.A-2 A', 'F.A-2 B', 'F.A-2 Optional Subjects'] }
  },
  DEFAULT_QUOTAS: ['Open Quota', 'UoP Employee Quota', 'Afghan Quota', 'Agriculture University Quota', 'Disabled Quota',
    'FATA IN Quota', 'Foreign Nationals', 'Minority Quota', 'O-Level Quota', 'Pakistan Forest Institute Quota', 'Sports Quota', 'UET Quota'],
  POOLS: [
    { quota: 'Afghan Quota', members: ['Pre Medical Open', 'Pre Engg', 'FCS'] },
    { quota: 'Sports Quota', members: ['FA', 'FCS', 'Pre Engg', 'Pre Medical Open'] },
    { quota: 'UET Quota',    members: ['FA', 'FCS', 'Pre Engg', 'Pre Medical Open'] }
  ],
  POOL_KEY: 'POOL',
  QUOTA_ONLY: { 'Evening Pre Medical': ['Open Quota'] },
  SEED: {   // used only if "Quota Settings" cannot be read
    'FA': { 'Afghan Quota': 1, 'Agriculture University Quota': 1, 'Disabled Quota': 1, 'FATA IN Quota': 2, 'Foreign Nationals': 2,
            'Minority Quota': 1, 'O-Level Quota': 2, 'Open Quota': 120, 'Pakistan Forest Institute Quota': 1 },
    'FCS': { 'Open Quota': 60 }, 'Pre Engg': { 'Open Quota': 40 },
    'Pre Medical Open': { 'Agriculture University Quota': 2, 'Disabled Quota': 1, 'FATA IN Quota': 2, 'Foreign Nationals': 2,
            'Minority Quota': 1, 'O-Level Quota': 6, 'Open Quota': 85, 'Pakistan Forest Institute Quota': 2 },
    'Evening Pre Medical': { 'Open Quota': 300 },
    'Self Morning Medical': { 'Agriculture University Quota': 3, 'O-Level Quota': 8, 'Open Quota': 400, 'UET Quota': 3 },
    'POOL': { 'Afghan Quota': 5, 'Sports Quota': 6, 'UET Quota': 8 }
  }
};

/* Form field -> Transform column (header with spaces / punctuation removed, lower case) */
var ADM_FIELDS = {
  roll: 'enroll', name: 'studentname', father: 'fathername', id: 'portalid', quota: 'quota', dob: 'dob',
  metric: 'obtainedmarks', total: 'totalmarks', hq: 'hq', room: 'room', pic: 'picurl', contact: 'studentcontact',
  email: 'email', board: 'board', passYear: 'passingyear', matricRoll: 'matricrollno', challan: 'challandetails',
  groupOnly: 'grouponly', meritNo: 'meritno', batch: 'batch'
};
/* Columns that always get the same value: Enroll# = REG#, Pic URL = PHOTO, MATRIC ROLL NO = Metric RollNo.
   S.No, SESSION and STATUS are filled for a new admission; Program/Class/Section, PCS, Appeared in,
   Aggregate and PERCENTAGE are worked out from the form. So every column of Transform is covered. */
/* Fields sent to the phone for every student (offline search / edit) */
var ADM_INDEX = ['sno', 'roll', 'name', 'father', 'id', 'key', 'year', 'pcs', 'quota', 's1', 's2', 's3', 's4', 's5', 's6', 's7',
  'dob', 'metric', 'total', 'hq', 'room', 'pic', 'contact', 'email', 'board', 'passYear', 'matricRoll', 'challan',
  'groupOnly', 'meritNo', 'batch', 'session', 'status', 'aggregate', 'agewise', 'feepaid', 'receipt', 'receiptNote'];   // v7.7: + merit lists

/* ================= Router (called from doPost) ================= */
function admRoute_(user, req) {
  if (!admCanUse_(user)) return { ok: false, error: 'ADM_NO_ACCESS' };
  switch (String(req.action)) {
    case 'adm_init':      return admInit_();
    case 'adm_index':     return admIndex_();
    case 'adm_next_roll': return admNextRollFor_(req.program, req.year);
    case 'adm_portal':    return admPortal_(req.q || req.id);
    case 'adm_check':     return admCheck_(req.d || {}, req.ctx || null);
    case 'adm_save':      return admSave_(user, req.d || {}, req.ctx || null, req.override === true, String(req.cid || ''));
    case 'adm_new':       return admNew_();
    case 'adm_refresh':   return admRefresh_(user, req.fields || [], req.apply === true, req.accept || {});
    case 'adm_add_new':   return admAddNew_(user, req.items || [], req.override === true);
    case 'adm_fee_receipt': return admFeeReceipt_(user, req);
    case 'adm_oml_get':   return admOmlGet_();
    case 'adm_oml_save':  return admOmlSave_(user, req);
    case 'adm_oml_confirm': return admOmlConfirm_(user, req);
    case 'adm_last_date': {                                  // v7.7: last date of admission (merit lists: later = lost her right to a seat)
      var d = String(req.date || '').trim();
      if (d && !/^\d{4}-\d\d-\d\d$/.test(d)) return { ok: false, error: 'Date as yyyy-mm-dd' };
      PropertiesService.getScriptProperties().setProperty('ADM_LAST_DATE', d);
      admLog_(user, 'LAST DATE OF ADMISSION', '', '', d || '(none)', '');
      return { ok: true, date: d };
    }
  }
  return { ok: false, error: 'UNKNOWN_ACTION' };
}

function admCanUse_(user) {
  try { if (typeof isAdmin_ === 'function' && isAdmin_(user)) return true; } catch (e) {}
  try { var sc = typeof scopeOf_ === 'function' ? scopeOf_(user) : null; if (sc && sc.all) return true; } catch (e) {}   // Users sheet Role = ADMIN or PRINCIPAL
  var u = String(user && user.username || '').trim().toLowerCase();
  var extra = ADM_CFG.EXTRA_USERS.concat(String(PropertiesService.getScriptProperties().getProperty('ADM_USERS') || '').split(','));
  return extra.some(function (x) { return String(x).trim().toLowerCase() === u && u; });
}

/* ================= small helpers ================= */
function admQ_(s) { return String(s == null ? '' : s).toLowerCase().replace(/\s+/g, ' ').trim(); }
function admH_(s) { return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, ''); }
function admS_(v) { return String(v == null ? '' : v).trim(); }
function admProg_(key) {
  var p = ADM_CFG.PROGRAMS.filter(function (x) { return x.key === key; })[0];
  if (!p) throw new Error('Unknown program: ' + key);
  return p;
}
function admIsMig_(year) { return admQ_(year) === admQ_(ADM_CFG.MIGRATION.YEAR); }
function admYear_(year) { return admIsMig_(year) ? ADM_CFG.MIGRATION.YEAR : ADM_CFG.ENTRY_YEAR; }
/** 'A1' -> 'A', 'FCS-1 D2' -> 'D2', 'F.A-1 Optional Subjects' -> 'Optional Subjects' */
function admSecCode_(name) {
  var s = admS_(name);
  if (/^[A-Za-z]\d$/.test(s)) return s.charAt(0).toUpperCase();
  return s.replace(/^\S+\s+/, '');
}
function admNormRoll_(v) {
  var s = admS_(v).toUpperCase(), m = s.match(/^(.*?)(\d+)$/);
  return m ? m[1] + parseInt(m[2], 10) : s;
}
function admRollNum_(v, prefix) {
  var s = admS_(v).toUpperCase(), pre = prefix.toUpperCase();
  if (s.indexOf(pre) !== 0) return null;
  var m = s.substring(pre.length).match(/^(\d+)$/);
  return m ? parseInt(m[1], 10) : null;
}
function admFmtRoll_(n) { return ADM_CFG.ROLL_PREFIX + (n < 10 ? '0' + n : String(n)); }
function admNormName_(s) {
  return String(s == null ? '' : s).toLowerCase().replace(/[^a-z؀-ۿ\s]/g, ' ').split(/\s+/).filter(Boolean).sort().join(' ');
}
function admBook_() { return SpreadsheetApp.openById(ADM_CFG.BOOK_ID); }
function admNum_(v) { var s = admS_(v); return s === '' || isNaN(Number(s)) ? '' : Number(s); }
function admNow_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm'); }

/* ================= Settings (read only, from the admissions workbook) ================= */
function admSettings_() {
  var cache = CacheService.getScriptCache(), hit = cache.get('adm_set_v1');
  if (hit) return JSON.parse(hit);
  var out = { seats: null, sections: null, note: '' };
  try {
    var book = SpreadsheetApp.openById(ADM_CFG.SETTINGS_BOOK_ID);
    var qs = book.getSheetByName(ADM_CFG.SETTINGS_SHEET);
    if (qs && qs.getLastRow() > 1) {
      var seats = {};
      qs.getRange(2, 1, qs.getLastRow() - 1, 3).getValues().forEach(function (r) {
        if (!r[0] || !r[1] || r[2] === '' || isNaN(Number(r[2]))) return;
        var p = admS_(r[0]); (seats[p] = seats[p] || {})[admQ_(r[1])] = Number(r[2]);
      });
      out.seats = seats;
    }
    var ps = book.getSheetByName(ADM_CFG.SECTIONS_SHEET);
    if (ps && ps.getLastRow() > 1) {
      var secs = {};
      ADM_CFG.PROGRAMS.forEach(function (p) { secs[p.key] = { '1st Year': [], '2nd Year': [] }; });
      ps.getRange(2, 1, ps.getLastRow() - 1, 3).getValues().forEach(function (r) {
        var pk = admS_(r[0]), yr = admS_(r[1]), x = admS_(r[2]);
        if (!secs[pk] || !yr || !x) return;
        (secs[pk][yr] = secs[pk][yr] || []).push(x);
      });
      out.sections = secs;
    }
  } catch (e) { out.note = 'Could not read the admissions workbook settings (' + e.message + '), built-in seats and sections are used.'; }
  if (!out.seats) {
    out.seats = {};
    Object.keys(ADM_CFG.SEED).forEach(function (pk) {
      out.seats[pk] = {};
      Object.keys(ADM_CFG.SEED[pk]).forEach(function (q) { out.seats[pk][admQ_(q)] = ADM_CFG.SEED[pk][q]; });
    });
  }
  if (!out.sections) out.sections = JSON.parse(JSON.stringify(ADM_CFG.SECTIONS_BY_PROGRAM));
  cache.put('adm_set_v1', JSON.stringify(out), 300);
  return out;
}

/* ================= Reading Transform-ALL-COLLEGE-STRENGTH ================= */
function admRead_() {
  var sh = admBook_().getSheetByName(ADM_CFG.SHEET);
  if (!sh) throw new Error('Sheet "' + ADM_CFG.SHEET + '" was not found.');
  var lr = sh.getLastRow(), lc = sh.getLastColumn();
  var vals = sh.getRange(1, 1, Math.max(lr, 1), Math.max(lc, 1)).getDisplayValues();
  var head = vals[0], cols = {};
  head.forEach(function (h, i) { var k = admH_(h); if (k && !(k in cols)) cols[k] = i; });
  if (cols.studentname === undefined || cols.enroll === undefined) throw new Error('"' + ADM_CFG.SHEET + '" needs the columns Enroll# and Student Name in row 1.');
  return { sh: sh, head: head, cols: cols, rows: vals.slice(1), width: lc };
}
function admCell_(T, r, key) { var c = T.cols[key]; return c === undefined ? '' : admS_(r[c]); }

/** which admission program (Quota Settings key) a Transform row belongs to */
function admKeyOf_(programName, year, pcs, roll, sections) {
  var h = admH_(programName).replace(/jinnahcollege/, 'jcw');
  var fam = ADM_CFG.PROGRAMS.filter(function (p) { return admH_(p.program) === h; });
  if (!fam.length) return '';
  if (fam.length === 1) return fam[0].key;
  var yr = admYear_(year);
  for (var i = 0; i < fam.length; i++) {
    var list = (sections[fam[i].key] || {})[yr] || [];
    if (list.some(function (s) { return admQ_(s) === admQ_(pcs); })) return fam[i].key;
  }
  var n = admRollNum_(roll, ADM_CFG.ROLL_PREFIX);
  if (n !== null) for (var j = 0; j < fam.length; j++) {
    var rg = ADM_CFG.ROLL_RANGES[fam[j].key];
    if (rg && n >= rg[0] && n <= rg[1]) return fam[j].key;
  }
  return fam[0].key;
}

function admRec_(T, r, i, sections) {
  var g = function (k) { return admCell_(T, r, k); };
  var rec = { row: i + 2, sno: g('sno'), roll: g('enroll'), name: g('studentname'), father: g('fathername'), id: g('portalid'),
    programName: g('program'), year: admYear_(g('class')), pcs: g('pcs'), section: g('section'), quota: g('quota'),
    dob: g('dob'), metric: g('obtainedmarks'), total: g('totalmarks'), hq: g('hq'), room: g('room'), pic: g('picurl') || g('photo'),
    contact: g('studentcontact'), email: g('email'), board: g('board'), passYear: g('passingyear'),
    matricRoll: g('matricrollno') || g('metricrollno'), challan: g('challandetails'),
    groupOnly: g('grouponly'), meritNo: g('meritno'), batch: g('batch'), session: g('session'), status: g('status'),
    aggregate: g('aggregate'), agewise: g('agewise'), feepaid: g('feepaiddate'), receipt: g('feereceiptdate'), receiptNote: g('feedatenote') };
  if (/^null$/i.test(rec.groupOnly)) rec.groupOnly = '';
  for (var s = 1; s <= 7; s++) rec['s' + s] = g('sub' + s);
  if (!rec.pcs && rec.section) rec.pcs = rec.section;
  rec.key = admKeyOf_(rec.programName, rec.year, rec.pcs, rec.roll, sections);
  if (rec.metric === '0') rec.metric = '';
  if (rec.total === '0') rec.total = '';
  if (rec.matricRoll === '0') rec.matricRoll = '';
  return rec;
}
function admAll_(T, sections) {
  var out = [];
  T.rows.forEach(function (r, i) { if (admCell_(T, r, 'studentname')) out.push(admRec_(T, r, i, sections)); });
  return out;
}

/* ================= Seats ================= */
function admCounts_(recs) {
  var c = {};
  ADM_CFG.PROGRAMS.forEach(function (p) { c[p.key] = { filled: 0, byQuota: {}, noQuota: 0 }; });
  recs.forEach(function (r) {
    if (admIsMig_(r.year) || !c[r.key]) return;   // only 1st Year is counted against seats
    c[r.key].filled++;
    var k = admQ_(r.quota);
    if (!k) c[r.key].noQuota++; else c[r.key].byQuota[k] = (c[r.key].byQuota[k] || 0) + 1;
  });
  return c;
}
function admPool_(pk, qk) {
  return ADM_CFG.POOLS.filter(function (p) { return admQ_(p.quota) === qk && p.members.indexOf(pk) >= 0; })[0] || null;
}
function admQuotaAllowed_(pk, qk) {
  var only = ADM_CFG.QUOTA_ONLY[pk];
  return !only || only.some(function (q) { return admQ_(q) === qk; });
}
function admResolve_(pk, qk, seatMap, counts) {
  var own = counts[pk].byQuota[qk] || 0, pool = admPool_(pk, qk);
  if (pool) {
    var ps = (seatMap[ADM_CFG.POOL_KEY] || {})[qk], f = 0;
    pool.members.forEach(function (m) { f += (counts[m] && counts[m].byQuota[qk]) || 0; });
    return { pooled: true, pool: pool, seats: ps === undefined ? '' : ps, filled: f, own: own };
  }
  var s = (seatMap[pk] || {})[qk];
  return { pooled: false, pool: null, seats: s === undefined ? '' : s, filled: own, own: own };
}
function admStatus_(seats, filled) {
  if (seats === '' || seats == null) return 'unset';
  if (filled > seats) return 'over';
  if (filled === seats) return 'full';
  if (seats - filled <= ADM_CFG.THRESHOLD) return 'near';
  return 'ok';
}
function admQuotaNames_(counts) {
  var names = {};
  ADM_CFG.DEFAULT_QUOTAS.forEach(function (n) { names[admQ_(n)] = n; });
  Object.keys(counts).forEach(function (pk) {
    Object.keys(counts[pk].byQuota).forEach(function (k) {
      if (!names[k]) names[k] = k.replace(/\b\w/g, function (c) { return c.toUpperCase(); });
    });
  });
  return names;
}
function admSeatState_(recs, set) {
  var counts = admCounts_(recs), names = admQuotaNames_(counts), keys = Object.keys(names), seatMap = set.seats;
  var programs = ADM_CFG.PROGRAMS.map(function (p) {
    var c = counts[p.key], tv = (seatMap[p.key] || {})[admQ_(ADM_CFG.TOTAL_NAME)];
    var quotas = keys.filter(function (k) { return admQuotaAllowed_(p.key, k); }).map(function (k) {
      var r = admResolve_(p.key, k, seatMap, counts);
      return { name: names[k], pooled: r.pooled, seats: r.seats, filled: r.filled, own: r.own, status: admStatus_(r.seats, r.filled) };
    }).filter(function (q) { return q.seats !== '' || q.own > 0 || admQ_(q.name) === admQ_(ADM_CFG.OPEN_QUOTA); });
    return { key: p.key, filled: c.filled, noQuota: c.noQuota, total: tv === undefined ? '' : tv,
             totalStatus: admStatus_(tv === undefined ? '' : tv, c.filled), quotas: quotas };
  });
  var pools = ADM_CFG.POOLS.map(function (pl) {
    var k = admQ_(pl.quota), ps = (seatMap[ADM_CFG.POOL_KEY] || {})[k], seats = ps === undefined ? '' : ps, f = 0;
    var parts = pl.members.map(function (m) { var n = counts[m].byQuota[k] || 0; f += n; return { key: m, filled: n }; });
    return { quota: pl.quota, seats: seats, filled: f, status: admStatus_(seats, f), members: parts };
  });
  return { programs: programs, pools: pools, at: admNow_() };
}

/* ================= Rooms (Timetable sheet: PCS -> most used theory room this session) ================= */
function admRooms_() {
  var cache = CacheService.getScriptCache(), hit = cache.get('adm_rooms_v1');
  if (hit) return JSON.parse(hit);
  var map = {};
  try {
    var sh = admBook_().getSheetByName(ADM_CFG.TIMETABLE_SHEET);
    var vals = sh ? sh.getDataRange().getDisplayValues() : [];
    var c = {}; (vals[0] || []).forEach(function (h, j) { var k = admH_(h); if (k && !(k in c)) c[k] = j; });
    if (c.room !== undefined && c.pcs !== undefined) {
      var tally = {};
      for (var i = 1; i < vals.length; i++) {
        var r = vals[i], room = admS_(r[c.room]), pcs = admS_(r[c.pcs]);
        if (!room || !pcs) continue;
        var cur = c.session === undefined || admS_(r[c.session]) === ADM_CFG.SESSION;
        var lab = /lab/i.test(room) || (c.practical !== undefined && admS_(r[c.practical]) !== '');
        var tier = (cur ? 0 : 1) + (lab ? 2 : 0);
        var t = tally[pcs] = tally[pcs] || [{}, {}, {}, {}];
        t[tier][room] = (t[tier][room] || 0) + 1;
      }
      Object.keys(tally).forEach(function (pcs) {
        for (var k = 0; k < 4; k++) {
          var o = tally[pcs][k], best = Object.keys(o).sort(function (a, b) { return o[b] - o[a]; })[0];
          if (best) { map[pcs] = best; break; }
        }
      });
    }
  } catch (e) { /* rooms are only a suggestion */ }
  cache.put('adm_rooms_v1', JSON.stringify(map), 600);
  return map;
}

/** room per section: the room most students of that section already have in Transform, else the timetable room */
function admRoomMap_(recs) {
  var m = admRooms_(), t = {};
  recs.forEach(function (r) { if (r.room && r.pcs) { var o = t[r.pcs] = t[r.pcs] || {}; o[r.room] = (o[r.room] || 0) + 1; } });
  Object.keys(t).forEach(function (pcs) { m[pcs] = Object.keys(t[pcs]).sort(function (a, b) { return t[pcs][b] - t[pcs][a]; })[0]; });
  return m;
}

/* ================= Subjects ================= */
function admYearSubject_(subject, classProgram, year) {
  var s = admS_(subject).toUpperCase();
  if (!admIsMig_(year)) return s;
  var rule = ADM_CFG.SECOND_YEAR_SUBJECTS.filter(function (x) { return x.from === s && (!x.only || x.only === classProgram); })[0];
  return rule ? rule.to : s;
}
/** built-in groups + other subject combinations already used in Transform for that program and year */
function admSubjects_(recs) {
  var out = {};
  ADM_CFG.PROGRAMS.forEach(function (p) {
    out[p.key] = {};
    [ADM_CFG.ENTRY_YEAR, ADM_CFG.MIGRATION.YEAR].forEach(function (yr) {
      var y = function (x) { return admYearSubject_(x, p.classProgram, yr); };
      var groups = (ADM_CFG.SUBJECT_DEFAULTS[p.classProgram] || []).map(function (g) { return { name: g.name, subjects: g.subjects.map(y) }; });
      var seen = {}; groups.forEach(function (g) { seen[g.subjects.slice().sort().join('|')] = true; });
      var tally = {};
      recs.forEach(function (r) {
        if (r.key !== p.key || r.year !== yr) return;
        var t = [r.s5, r.s6, r.s7].map(function (x) { return admS_(x).toUpperCase(); });
        if (t.some(function (x) { return !x; })) return;
        var k = t.slice().sort().join('|');
        if (seen[k]) return;
        (tally[k] = tally[k] || { n: 0, subjects: t }).n++;
      });
      Object.keys(tally).filter(function (k) { return tally[k].n >= 2; })
        .sort(function (a, b) { return tally[b].n - tally[a].n; })
        .forEach(function (k) { groups.push({ name: tally[k].subjects.join(' / '), subjects: tally[k].subjects }); });
      out[p.key][yr] = { compulsory: ADM_CFG.COMPULSORY.map(y), groups: groups };
    });
  });
  return out;
}

/* ================= Roll numbers ================= */
/** every Roll No in use: Transform, Left Students and the old admissions workbook (cached 10 min) */
function admUsedRolls_(T) {
  var used = {};
  var add = function (v) { var s = admNormRoll_(v); if (s) used[s] = true; };
  T.rows.forEach(function (r) { add(r[T.cols.enroll]); });
  [ADM_CFG.LEFT_SHEET, ADM_CFG.PORTAL_SHEET].forEach(function (name) {
    try {
      var ls = admBook_().getSheetByName(name);
      if (!ls || ls.getLastRow() < 2) return;
      var lv = ls.getDataRange().getDisplayValues(), lc = -1;
      lv[0].forEach(function (h, j) { if (lc < 0 && /^enroll(ment)?(no)?$/.test(admH_(h))) lc = j; });
      if (lc >= 0) lv.slice(1).forEach(function (r) { add(r[lc]); });
    } catch (e) {}
  });
  var cache = CacheService.getScriptCache(), hit = cache.get('adm_oldrolls_v1');
  var old = hit ? JSON.parse(hit) : null;
  if (!old) {
    old = [];
    try {
      SpreadsheetApp.openById(ADM_CFG.SETTINGS_BOOK_ID).getSheets().forEach(function (sh) {
        if (sh.getLastRow() < 2 || sh.getLastColumn() < 1) return;
        var v = sh.getRange(1, 1, Math.min(sh.getLastRow(), 3000), sh.getLastColumn()).getDisplayValues();
        for (var i = 0; i < Math.min(v.length, 10); i++) {
          var c = -1; v[i].forEach(function (h, j) { if (c < 0 && admH_(h) === 'rollno') c = j; });
          if (c >= 0) { v.slice(i + 1).forEach(function (r) { var s = admNormRoll_(r[c]); if (/^JCW-F\d\d-\d+$/.test(s)) old.push(s); }); break; }
        }
      });
      cache.put('adm_oldrolls_v1', JSON.stringify(old), 600);
    } catch (e) {}
  }
  old.forEach(function (s) { used[s] = true; });
  return used;
}
function admNextRoll_(T, pk, year) {
  var used = admUsedRolls_(T), mx = 0, cand, pre, from, to;
  if (admIsMig_(year)) {
    pre = ADM_CFG.MIGRATION.ROLL_PREFIX; from = ADM_CFG.MIGRATION.ROLL_START; to = ADM_CFG.MIGRATION.ROLL_END;
  } else {
    pre = ADM_CFG.ROLL_PREFIX; var rg = ADM_CFG.ROLL_RANGES[pk];
    if (!rg) return { roll: '', error: 'No Roll No range is set for ' + pk + '.' };
    from = rg[0]; to = rg[1];
  }
  Object.keys(used).forEach(function (s) {
    var n = admRollNum_(s, pre);
    if (n !== null && n >= from && n <= to && n > mx) mx = n;
  });
  cand = mx ? mx + 1 : from;
  while (used[admNormRoll_(pre + cand)]) cand++;
  if (cand > to) return { roll: '', error: 'The Roll No range ' + from + ' to ' + to + ' is full.' };
  var roll = admIsMig_(year) ? pre + cand : admFmtRoll_(cand);
  return { roll: roll, number: cand, range: [from, to] };
}
function admNextRollFor_(pk, year) {
  admProg_(pk);
  var r = admNextRoll_(admRead_(), pk, year);
  r.ok = !r.error;
  return r;
}

/* ================= Init / index for the phone ================= */
function admInit_() {
  var set = admSettings_(), T = admRead_(), recs = admAll_(T, set.sections), qn = admQuotaNames_(admCounts_(recs));
  return {
    ok: true,
    programs: ADM_CFG.PROGRAMS.map(function (p) {
      return { key: p.key, label: p.label, program: p.program, classProgram: p.classProgram, fee: p.fee,
               roll: ADM_CFG.ROLL_RANGES[p.key] || ['', ''], only: ADM_CFG.QUOTA_ONLY[p.key] || null };
    }),
    sections: set.sections, quotas: Object.keys(qn).map(function (k) { return qn[k]; }),
    seats: admSeatState_(recs, set), subjects: admSubjects_(recs), rooms: admRoomMap_(recs),
    cfg: { rollPrefix: ADM_CFG.ROLL_PREFIX, entryYear: ADM_CFG.ENTRY_YEAR, session: ADM_CFG.SESSION, mode: ADM_CFG.MODE,
           threshold: ADM_CFG.THRESHOLD, mig: ADM_CFG.MIGRATION, sheet: ADM_CFG.SHEET,
           lastDate: PropertiesService.getScriptProperties().getProperty('ADM_LAST_DATE') || '' },
    note: set.note, at: admNow_()
  };
}
function admPack_(rec) { return ADM_INDEX.map(function (k) { return rec[k] == null ? '' : String(rec[k]); }); }
function admIndex_() {
  var set = admSettings_(), T = admRead_();
  return { ok: true, cols: ADM_INDEX, rows: admAll_(T, set.sections).map(admPack_), at: admNow_() };
}

/* ================= Fill from the portal copy ("Portal Data" sheet, kept fresh by the UOP portal sync) ================= */
/** Search the portal copy by name / father name (all words must match), or by Student ID / Enrollment#.
 *  Returns up to 25 students; each one says whether she is already in Transform. */
/** the "Portal Data" sheet: { v: values, g(row, key) } or null when empty */
function admPortalSheet_() {
  var sh = admBook_().getSheetByName(ADM_CFG.PORTAL_SHEET);
  if (!sh || sh.getLastRow() < 2) return null;
  var v = sh.getDataRange().getDisplayValues(), c = {};
  v[0].forEach(function (h, j) { var k = admH_(h); if (k && !(k in c)) c[k] = j; });
  return { v: v, head: v[0], g: function (r, k) { return c[k] === undefined ? '' : admS_(r[c[k]]); } };
}
/** one portal row -> the fields of the admission form */
function admPortalStu_(P, row, set) {
  var g = P.g, yr = admYear_(g(row, 'class')), sub = [];
  var num = function (x) { var n = admNum_(String(x).replace(/,/g, '')); return n === '' || n === 0 ? '' : String(Math.round(n * 100) / 100); };
  for (var s = 1; s <= 7; s++) sub.push(g(row, 'sub' + s + (admIsMig_(yr) ? '2ndyear' : '1styear')) || g(row, 'sub' + s + '1styear'));
  var progName = g(row, 'program'), en = g(row, 'enrollment');
  return {
    id: g(row, 'studentid'), portalRoll: en, name: g(row, 'studentname'), father: g(row, 'fathername'),
    quota: g(row, 'category'), dob: g(row, 'dob'), contact: g(row, 'studentcontact'), email: g(row, 'emailaddress'),
    pic: g(row, 'photoname'), board: g(row, 'boarduniversity'), passYear: g(row, 'passingyear'), matricRoll: num(g(row, 'rollno')),
    metric: num(g(row, 'obtainmarkscgpa')), total: num(g(row, 'totalmarkscgpa')), challan: g(row, 'challandetails'), batch: g(row, 'batch'),
    subjects: sub, program: admKeyOf_(progName, yr, '', en, set.sections), programName: progName, year: yr,
    section: g(row, 'section'), status: g(row, 'status'), synced: g(row, 'detailssynced'), hq: g(row, 'hq'), paidOn: g(row, 'feepaiddate') };
}

function admPortal_(query) {
  var q = admS_(query);
  if (q.length < 2) return { ok: false, error: 'Type at least 2 letters of her name (or her portal ID).' };
  var P = admPortalSheet_();
  if (!P) return { ok: false, error: 'The "' + ADM_CFG.PORTAL_SHEET + '" sheet is empty. Run the UOP portal sync first.' };
  var v = P.v, g = P.g;
  var Q = q.toUpperCase(), words = q.toLowerCase().split(/\s+/).filter(Boolean), hits = [], exact = [];
  for (var i = 1; i < v.length; i++) {
    var r = v[i], id = g(r, 'studentid'), en = g(r, 'enrollment');
    if (!g(r, 'studentname')) continue;
    if (id.toUpperCase() === Q || (en && admNormRoll_(en) === admNormRoll_(Q))) { exact.push(r); continue; }
    var hay = (g(r, 'studentname') + ' ' + g(r, 'fathername') + ' ' + en + ' ' + id).toLowerCase();
    if (words.every(function (w) { return hay.indexOf(w) >= 0; })) hits.push(r);
  }
  var all = exact.concat(hits), more = Math.max(0, all.length - 25);
  if (!all.length) return { ok: true, list: [], note: 'Nobody called "' + q + '" in "' + ADM_CFG.PORTAL_SHEET + '". If she was added on the portal today, run a quick portal sync and search again.' };
  var set = admSettings_(), T = admRead_(), byId = {}, byRoll = {};
  admAll_(T, set.sections).forEach(function (r) { if (r.id) byId[r.id] = r; if (r.roll) byRoll[admNormRoll_(r.roll)] = r; });
  var list = all.slice(0, 25).map(function (row) {
    var st = admPortalStu_(P, row, set), have = byId[st.id] || (st.portalRoll ? byRoll[admNormRoll_(st.portalRoll)] : null);
    st.inTransform = have ? { roll: have.roll, pcs: have.pcs, name: have.name } : null;
    return st;
  });
  return { ok: true, list: list, more: more };
}

/* ================= New on the portal (Enrolled there, not yet in Transform or Left Students) ================= */
/** portal IDs, Enroll#s and name+father of everyone already in Transform and Left Students */
function admKnown_(recs) {
  var k = { ids: {}, rolls: {}, names: {} };
  var add = function (id, roll, name, father) {
    if (admS_(id)) k.ids[admS_(id)] = true;
    if (admS_(roll)) k.rolls[admNormRoll_(roll)] = true;
    var n = admNormName_(name), f = admNormName_(father);
    if (n && f) k.names[n + '|' + f] = true;
  };
  recs.forEach(function (r) { add(r.id, r.roll, r.name, r.father); });
  try {
    var ls = admBook_().getSheetByName(ADM_CFG.LEFT_SHEET);
    if (ls && ls.getLastRow() > 1) {
      var v = ls.getDataRange().getDisplayValues(), c = {};
      v[0].forEach(function (h, j) { var x = admH_(h); if (x && !(x in c)) c[x] = j; });
      var g = function (r, x) { return c[x] === undefined ? '' : r[c[x]]; };
      v.slice(1).forEach(function (r) {
        add(g(r, 'portalid') || g(r, 'studentid'), g(r, 'enroll') || g(r, 'enrollment'), g(r, 'studentname'), g(r, 'fathername'));
      });
    }
  } catch (e) {}
  return k;
}
function admIsKnown_(k, s) {
  if (s.id && k.ids[s.id]) return true;
  if (s.portalRoll && k.rolls[admNormRoll_(s.portalRoll)]) return true;
  var n = admNormName_(s.name), f = admNormName_(s.father);
  return !!(n && f && k.names[n + '|' + f]);
}
/** program from the portal: Pre-Medical has 3 admission programs, so look for "self", "evening" or the section letter */
function admGuessProg_(s, set) {
  var h = admH_(s.programName).replace(/jinnahcollege/, 'jcw');
  var fam = ADM_CFG.PROGRAMS.filter(function (p) { return admH_(p.program) === h; });
  if (!fam.length) fam = ADM_CFG.PROGRAMS.filter(function (p) { return h.indexOf(admH_(p.classProgram)) === 0; });
  if (!fam.length) return { key: '', sure: false };
  if (fam.length === 1) return { key: fam[0].key, sure: true };
  var txt = admQ_(s.programName + ' ' + s.quota + ' ' + s.batch + ' ' + s.section);
  var pick = function (re) { return fam.filter(function (p) { return re.test(admQ_(p.key)); })[0]; };
  var hit = /evening/.test(txt) ? pick(/evening/) : /self|morning/.test(txt) ? pick(/self/) : null;
  if (hit) return { key: hit.key, sure: true };
  var code = admS_(s.section).toUpperCase();
  if (code) {
    var bySec = fam.filter(function (p) {
      return ((set.sections[p.key] || {})[s.year] || []).some(function (x) { return admSecCode_(x).toUpperCase() === code || admS_(x).toUpperCase() === code; });
    });
    if (bySec.length === 1) return { key: bySec[0].key, sure: false };
  }
  return { key: fam[0].key, sure: false };
}
/** portal Category -> one of the quota names, else Open Quota */
function admGuessQuota_(cat, pk, names) {
  var c = admQ_(cat).replace(/\bquota\b/, '').trim(), list = Object.keys(names).map(function (k) { return names[k]; });
  if (pk && ADM_CFG.QUOTA_ONLY[pk]) list = ADM_CFG.QUOTA_ONLY[pk];
  if (c) {
    var hit = list.filter(function (q) { var x = admQ_(q).replace(/\bquota\b/, '').trim(); return x === c; })[0]
      || list.filter(function (q) { var x = admQ_(q).replace(/\bquota\b/, '').trim(); return x && (c.indexOf(x) >= 0 || x.indexOf(c) >= 0); })[0];
    if (hit) return { quota: hit, sure: true };
  }
  return { quota: ADM_CFG.OPEN_QUOTA, sure: !c || /open|merit|general|self|finance/.test(c) };
}
/** the next n free Roll Nos of a program (continuing after the highest one in use, like adm_next_roll) */
function admFreeRolls_(used, pk, year, n) {
  var mig = admIsMig_(year), pre, from, to, mx = 0, out = [];
  if (mig) { pre = ADM_CFG.MIGRATION.ROLL_PREFIX; from = ADM_CFG.MIGRATION.ROLL_START; to = ADM_CFG.MIGRATION.ROLL_END; }
  else { var rg = ADM_CFG.ROLL_RANGES[pk]; if (!rg) return { list: [], range: null }; pre = ADM_CFG.ROLL_PREFIX; from = rg[0]; to = rg[1]; }
  Object.keys(used).forEach(function (s) { var x = admRollNum_(s, pre); if (x !== null && x >= from && x <= to && x > mx) mx = x; });
  for (var cand = mx ? mx + 1 : from; out.length < n && cand <= to; cand++) {
    if (!used[admNormRoll_(pre + cand)]) out.push(mig ? pre + cand : admFmtRoll_(cand));
  }
  return { list: out, range: [from, to] };
}

/** every student Enrolled on the portal (this year's batch, or a 2nd Year migration) who is not in Transform yet,
 *  with a guessed program / quota, plus free Roll Nos and section strengths so the page can suggest them */
function admNew_() {
  var P = admPortalSheet_();
  if (!P) return { ok: true, list: [], pools: {}, note: 'The "' + ADM_CFG.PORTAL_SHEET + '" sheet is empty. Run Quick sync on the UOP portal first.', at: admNow_() };
  var set = admSettings_(), T = admRead_(), recs = admAll_(T, set.sections), known = admKnown_(recs), g = P.g;
  var names = admQuotaNames_(admCounts_(recs));
  var yearOf = function (b) { var m = String(b || '').match(/(?:19|20)\d\d/); return m ? Number(m[0]) : 0; };
  var maxY = 0;
  P.v.slice(1).forEach(function (r) {
    if (/^enrolled$/i.test(g(r, 'status')) && !admIsMig_(admYear_(g(r, 'class')))) maxY = Math.max(maxY, yearOf(g(r, 'batch')));
  });
  var used = admUsedRolls_(T), list = [], older = 0;
  P.v.slice(1).forEach(function (r) {
    if (!g(r, 'studentname') || !/^enrolled$/i.test(g(r, 'status'))) return;
    var s = admPortalStu_(P, r, set);
    if (admIsKnown_(known, s)) return;
    if (admIsMig_(s.year) ? /^JCW-/i.test(s.portalRoll) : (maxY && yearOf(s.batch) && yearOf(s.batch) < maxY)) { older++; return; }
    var gp = admGuessProg_(s, set), gq = admGuessQuota_(s.quota, gp.key, names);
    s.program = gp.key; s.progSure = gp.sure;
    s.all = [];                                   // every filled column of her "Portal Data" row, for the details view
    P.head.forEach(function (h, j) { var x = admS_(r[j]); if (admS_(h) && x && !/^(load status|details synced)$/i.test(admS_(h))) s.all.push([admS_(h), x]); });
    s.reg = g(r, 'reg'); s.session = g(r, 'session'); s.division = g(r, 'division'); s.pct = g(r, 'agemarks') || g(r, 'percentagemarks');
    s.quotaGuess = admIsMig_(s.year) ? ADM_CFG.MIGRATION.QUOTA : gq.quota; s.quotaSure = admIsMig_(s.year) || gq.sure;
    // an Enrollment# already typed on the portal is kept when it is free and in the program's range
    s.rollFromPortal = '';
    if (s.portalRoll && gp.key && !known.rolls[admNormRoll_(s.portalRoll)]) {   // (Portal Data itself counts as "used" in admUsedRolls_)
      var pre = admIsMig_(s.year) ? ADM_CFG.MIGRATION.ROLL_PREFIX : ADM_CFG.ROLL_PREFIX, n = admRollNum_(s.portalRoll, pre);
      var rg = admIsMig_(s.year) ? [ADM_CFG.MIGRATION.ROLL_START, ADM_CFG.MIGRATION.ROLL_END] : ADM_CFG.ROLL_RANGES[gp.key];
      if (n !== null && rg && n >= rg[0] && n <= rg[1]) s.rollFromPortal = admS_(s.portalRoll).toUpperCase();
    }
    list.push(s);
  });
  var order = ADM_CFG.PROGRAMS.map(function (p) { return p.key; });
  list.sort(function (a, b) {
    var pa = order.indexOf(a.program), pb = order.indexOf(b.program);
    if (pa !== pb) return (pa < 0 ? 99 : pa) - (pb < 0 ? 99 : pb);
    if (a.year !== b.year) return a.year < b.year ? -1 : 1;
    return (Number(a.id) || 0) - (Number(b.id) || 0);      // portal ID order = order of admission on the portal
  });
  // per program and year: free Roll Nos and how many students each section already has
  var pools = {}, n = list.length + 40;
  ADM_CFG.PROGRAMS.forEach(function (p) {
    [ADM_CFG.ENTRY_YEAR, ADM_CFG.MIGRATION.YEAR].forEach(function (yr) {
      var cnt = {};
      recs.forEach(function (r) { if (r.key === p.key && r.year === yr && r.pcs) cnt[r.pcs] = (cnt[r.pcs] || 0) + 1; });
      var fr = admFreeRolls_(used, p.key, yr, n);
      pools[p.key + '|' + yr] = { free: fr.list, range: fr.range,
        secs: ((set.sections[p.key] || {})[yr] || []).map(function (x) { return { name: x, count: cnt[x] || 0 }; }) };
    });
  });
  return { ok: true, list: list, pools: pools, older: older, batch: maxY || '', at: admNow_() };
}

/** adds the chosen portal students to Transform in one go. items: [{ id, program, year, section, quota, rollNo, subjects, force }] */
function admAddNew_(user, items, override) {
  if (!items || !items.length) return { ok: false, error: 'Nobody was selected.' };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var P = admPortalSheet_();
    if (!P) return { ok: false, error: 'The "' + ADM_CFG.PORTAL_SHEET + '" sheet is empty.' };
    var byId = {};
    P.v.slice(1).forEach(function (r) { var id = P.g(r, 'studentid'); if (id) byId[id] = r; });
    var set = admSettings_(), T = admRead_(), recs = admAll_(T, set.sections), C = T.cols, sh = T.sh, rooms = admRoomMap_(recs);
    var mx = 0;
    if (C.sno !== undefined) T.rows.forEach(function (r) { var n = Number(r[C.sno]); if (n > mx) mx = n; });
    var start = sh.getLastRow() + 1, lines = [], results = [], logs = [];
    items.forEach(function (it) {
      var res = { id: admS_(it.id), name: '', ok: false, errors: [], issues: [] };
      results.push(res);
      var row = byId[res.id];
      if (!row) { res.errors.push('Portal ID ' + res.id + ' is not in "' + ADM_CFG.PORTAL_SHEET + '" any more.'); return; }
      var s = admPortalStu_(P, row, set), p;
      res.name = s.name;
      try { p = admProg_(admS_(it.program)); } catch (e) { res.errors.push('Please choose her program.'); return; }
      var d = { program: p.key, year: admYear_(it.year || s.year), section: admS_(it.section), quota: admS_(it.quota) || admGuessQuota_(s.quota, p.key, admQuotaNames_(admCounts_(recs))).quota,
        rollNo: admS_(it.rollNo).toUpperCase(), name: s.name, father: s.father, id: s.id, dob: s.dob, contact: s.contact, email: s.email, pic: s.pic,
        board: s.board, passYear: s.passYear, matricRoll: s.matricRoll, metric: s.metric, total: s.total, challan: s.challan, batch: s.batch,
        hq: /^y/i.test(s.hq || '') ? 'Yes' : 'No', groupOnly: '', meritNo: '', room: '' };
      if (admIsMig_(d.year)) d.quota = ADM_CFG.MIGRATION.QUOTA;
      var subs = (it.subjects && it.subjects.filter(Boolean).length >= 4) ? it.subjects : s.subjects;
      if (subs.filter(Boolean).length < 4) {     // nothing from the portal yet: compulsory + the program's only group
        var grp = ADM_CFG.SUBJECT_DEFAULTS[p.classProgram] || [];
        subs = ADM_CFG.COMPULSORY.concat(grp.length === 1 ? grp[0].subjects : ['', '', '']);
      }
      d.subjects = subs.slice(0, 7).map(function (x) { return admYearSubject_(x, p.classProgram, d.year); });
      while (d.subjects.length < 7) d.subjects.push('');
      if (!d.rollNo) { var nr = admNextRoll_(T, p.key, d.year); if (nr.error) { res.errors.push(nr.error); return; } d.rollNo = nr.roll; }
      d.room = rooms[d.section] || '';
      var v = admValidate_(d, null, null, recs, set);
      res.roll = d.rollNo; res.section = d.section; res.issues = v.issues; res.errors = v.errors;
      var fi = admFeeIssue_(s.challan); if (fi) { v.issues.push(fi); res.issues = v.issues; res.fee = fi.text; }
      if (v.errors.length) return;
      if (v.blocking && !override && !it.force) { res.blocked = true; res.needsConfirm = true; return; }   // v7.5: seats full = warn, admit anyway after a confirm
      if (!override && !it.force && v.issues.some(function (i) { return i.level === 'danger'; })) { res.needsConfirm = true; return; }
      var vals = admValues_(d, p);
      vals.sno = ++mx; vals.session = ADM_CFG.SESSION; vals.status = ADM_CFG.STATUS;
      if (!vals.photo) vals.photo = vals.picurl;
      var line = T.head.map(function (h) { var k = admH_(h); return vals[k] === undefined || vals[k] === null ? '' : vals[k]; });
      res.row = start + lines.length; res.ok = true;
      lines.push(line);
      T.rows.push(line.map(String));             // later students in this batch see this Roll No as used
      recs.push({ row: res.row, roll: d.rollNo, name: d.name, father: d.father, id: d.id, key: p.key, year: d.year, pcs: d.section,
                  section: admSecCode_(d.section), quota: d.quota, programName: p.program });
      logs.push([new Date(), (user.teacher || '') + ' <' + user.username + '>', admIsMig_(d.year) ? 'ADD FROM PORTAL (2nd Year migration)' : 'ADD FROM PORTAL',
        d.rollNo, d.name, p.label + ', ' + d.year + ', ' + d.section + ', ' + d.quota + ', portal ID ' + d.id + (override || it.force ? ' (saved anyway)' : ''), '']);
    });
    if (lines.length) {
      if (C.dob !== undefined) sh.getRange(start, C.dob + 1, lines.length, 1).setNumberFormat('@');
      if (C.studentcontact !== undefined) sh.getRange(start, C.studentcontact + 1, lines.length, 1).setNumberFormat('@');
      sh.getRange(start, 1, lines.length, T.head.length).setValues(lines);
      SpreadsheetApp.flush();
      try {
        var lg = admLogSheet_();
        lg.getRange(lg.getLastRow() + 1, 1, logs.length, 7).setValues(logs);
      } catch (e) { /* logging never stops a save */ }
    }
    var T2 = admRead_(), recs2 = admAll_(T2, set.sections), added = {};
    results.forEach(function (r) { if (r.ok) added[r.row] = true; });
    return { ok: true, added: lines.length, results: results, seats: admSeatState_(recs2, set),
             recs: recs2.filter(function (r) { return added[r.row]; }).map(admPack_) };
  } finally { lock.releaseLock(); }
}

/** v7.5: only the registration fee (Rs 2,600 or less) paid on the portal -> a warning before she gets her Enroll# and section */
var ADM_REG_FEE_MAX = 3000;
function admFeeIssue_(challan) {
  var f = String(challan || '').split(',').map(function (x) { var m = x.trim().match(/^([PU])-?([\d.]+)/i); return m ? { paid: /p/i.test(m[1]), amt: Number(m[2]) } : null; }).filter(Boolean);
  if (!f.length) return { level: 'danger', text: 'FEE: no fee challan on the portal copy. Ask her for the paid fee slip before giving her Enroll# and section.' };
  var paidFee = f.filter(function (x) { return x.paid && x.amt > ADM_REG_FEE_MAX; }), due = f.filter(function (x) { return !x.paid; }).reduce(function (a, x) { return a + x.amt; }, 0);
  if (paidFee.length) return null;
  var reg = f.filter(function (x) { return x.paid; }).reduce(function (a, x) { return a + x.amt; }, 0);
  return { level: 'danger', text: 'FEE NOT PAID: ' + (reg ? 'only the registration fee (Rs ' + reg + ') is paid' : 'nothing is paid') + (due ? '; Rs ' + due + ' unpaid' : '') + '. Ask her to show the paid fee slip before giving her Enroll# and section.' };
}

/** v7.7: the date on the student's own fee receipt when the portal date is later (bank sent it late).
 *  Kept in FEE RECEIPT DATE + FEE DATE NOTE (portal syncs never touch them); the merit lists use it instead of FEE PAID DATE. */
function admFeeReceipt_(user, req) {
  var roll = admS_(req.roll), d = admS_(req.date), note = admS_(req.note).slice(0, 300);
  if (!roll) return { ok: false, error: 'No Enroll#' };
  if (d && !/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(d)) return { ok: false, error: 'Date as dd/mm/yyyy' };
  if (d && !note) return { ok: false, error: 'Write why (e.g. "receipt seen, bank late")' };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var T = admRead_(), sh = T.sh, want = admNormRoll_(roll), r = -1;
    for (var i = 0; i < T.rows.length; i++) if (admNormRoll_(T.rows[i][T.cols.enroll]) === want) { r = i; break; }
    if (r < 0) return { ok: false, error: roll + ' is not in the strength sheet.' };
    var col = function (key, head) {
      if (T.cols[key] !== undefined) return T.cols[key] + 1;
      var c = sh.getLastColumn() + 1; if (sh.getMaxColumns() < c) sh.insertColumnsAfter(sh.getMaxColumns(), 1);
      sh.getRange(1, c).setValue(head).setFontWeight('bold'); T.cols[key] = c - 1; return c;
    };
    var cd = col('feereceiptdate', 'FEE RECEIPT DATE'), cn = col('feedatenote', 'FEE DATE NOTE');
    sh.getRange(r + 2, cd).setNumberFormat('@').setValue(d);
    sh.getRange(r + 2, cn).setValue(d ? note + ' (' + (user.teacher || user.name || '') + ', ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy') + ')' : '');
    admLog_(user, 'FEE RECEIPT DATE', roll, admS_(T.rows[r][T.cols.studentname]), (d || '(removed)') + (note ? ' · ' + note : ''), '');
    var T2 = admRead_(); return { ok: true, row: r + 2, rec: admPack_(admRec_(T2, T2.rows[r], r, admSettings_().sections)) };
  } finally { lock.releaseLock(); }
}

/* ================= v7.8 Official merit lists (Directorate of Admissions PDFs, read on the page) =================
   The portal shows nearly every student as Open Quota in the first days; the official final merit lists say who was selected,
   in which quota and whether Subsidized / Non-Subsidized. UoP Employee Quota applicants must be confirmed at the college. */
var ADM_OML_SHEET = 'Official Merit Lists';
var ADM_OML_HEAD = ['Program', 'List', 'S.#', 'Form#', 'Name', 'Father Name', 'Obt Marks', 'Total Marks', 'HQ', 'AGG %', 'Applied Quota', 'Merit Position',
                    'Uploaded', 'Confirmed', 'Confirmed By', 'Confirmed On', 'Note'];
function admOmlSheet_() {
  var b = admBook_(), sh = b.getSheetByName(ADM_OML_SHEET);
  if (!sh) { sh = b.insertSheet(ADM_OML_SHEET); sh.getRange(1, 1, 1, ADM_OML_HEAD.length).setValues([ADM_OML_HEAD]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function admOmlGet_() {
  var sh = admOmlSheet_(), n = sh.getLastRow() - 1;
  var v = n > 0 ? sh.getRange(2, 1, n, ADM_OML_HEAD.length).getDisplayValues() : [];
  return { ok: true, rows: v };
}
function admOmlSave_(user, req) {
  var list = admS_(req.list).slice(0, 120), rows = req.rows || [];
  if (!list || !rows.length) return { ok: false, error: 'Nothing to save.' };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = admOmlSheet_(), n = sh.getLastRow() - 1, old = n > 0 ? sh.getRange(2, 1, n, ADM_OML_HEAD.length).getValues() : [];
    var conf = {};                                                       // confirmations already given are kept when the list is read again
    old.forEach(function (r) { if (r[1] === list && r[13]) conf[r[3]] = [r[13], r[14], r[15], r[16]]; });
    var keep = old.filter(function (r) { return r[1] !== list; }), now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    var add = rows.map(function (r) { var c = conf[admS_(r.form)] || ['', '', '', ''];
      return [admS_(r.program), list, admS_(r.sno), admS_(r.form), admS_(r.name), admS_(r.father), admS_(r.obt), admS_(r.total), admS_(r.hq), admS_(r.agg), admS_(r.quota), admS_(r.pos), now].concat(c); });
    var all = keep.concat(add);
    if (n > 0) sh.getRange(2, 1, n, ADM_OML_HEAD.length).clearContent();
    if (all.length) { sh.getRange(2, 1, all.length, ADM_OML_HEAD.length).setNumberFormat('@'); sh.getRange(2, 1, all.length, ADM_OML_HEAD.length).setValues(all.map(function (r) { return r.map(String); })); }
    admLog_(user, 'OFFICIAL MERIT LIST', '', '', list + ': ' + add.length + ' rows', '');
    return { ok: true, saved: add.length, total: all.length };
  } finally { lock.releaseLock(); }
}
function admOmlConfirm_(user, req) {
  var list = admS_(req.list), form = admS_(req.form), yes = req.yes === true, note = admS_(req.note).slice(0, 200);
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = admOmlSheet_(), n = sh.getLastRow() - 1, v = n > 0 ? sh.getRange(2, 1, n, 4).getDisplayValues() : [];
    for (var i = 0; i < v.length; i++) if (v[i][1] === list && v[i][3] === form) {
      var by = user.teacher || user.name || '', on = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
      sh.getRange(i + 2, 14, 1, 4).setValues([[yes ? 'Yes' : 'No', by, on, note]]);
      admLog_(user, 'EMPLOYEE QUOTA ' + (yes ? 'CONFIRMED' : 'NOT CONFIRMED'), '', '', form + (note ? ' · ' + note : ''), '');
      return { ok: true, confirmed: yes ? 'Yes' : 'No', by: by, on: on };
    }
    return { ok: false, error: 'Form# ' + form + ' is not in ' + list };
  } finally { lock.releaseLock(); }
}

/* ================= Validation ================= */
function admEvaluate_(pk, quota, recs, set) {
  var counts = admCounts_(recs), seatMap = set.seats, k = admQ_(quota), p = admProg_(pk);
  var r = admResolve_(pk, k, seatMap, counts), cnt = counts[pk];
  var tv = (seatMap[pk] || {})[admQ_(ADM_CFG.TOTAL_NAME)], total = tv === undefined ? '' : tv;
  var issues = [], blocking = false, block = ADM_CFG.MODE === 'block';
  var lbl = function (m) { return admProg_(m).label; };
  var where = r.pooled ? 'the shared pool (' + r.pool.members.map(lbl).join(', ') + ')' : 'this program';
  issues.push({ level: 'info', text: p.label + ' is a ' + (p.fee === 'non' ? 'NON-SUBSIDIZED' : 'SUBSIDIZED') + ' program.' });
  if (r.seats === '') issues.push({ level: 'info', text: 'No seat limit is set for "' + quota + '" in ' + (r.pooled ? 'the shared pool' : 'this program') + ', so it is not checked.' });
  else if (r.filled >= r.seats) {
    issues.push({ level: 'danger', text: '"' + quota + '" is ' + (r.filled > r.seats ? 'over its limit' : 'full') + ' in ' + where + ': ' + r.filled + ' of ' + r.seats + ' seats are taken.' });
    if (block) blocking = true;
  } else if (r.seats - r.filled <= ADM_CFG.THRESHOLD) {
    var left = r.seats - r.filled;
    issues.push({ level: 'note', text: 'Only ' + left + ' seat' + (left === 1 ? '' : 's') + ' left in "' + quota + '" (' + r.filled + ' of ' + r.seats + ' taken' + (r.pooled ? ', shared across programs' : '') + ').' });
  } else issues.push({ level: 'good', text: r.filled + ' of ' + r.seats + ' seats taken in "' + quota + '"' + (r.pooled ? ' (shared pool)' : '') + '. ' + (r.seats - r.filled) + ' left.' });
  if (total !== '' && cnt.filled >= total) {
    issues.push({ level: 'danger', text: 'The whole program is ' + (cnt.filled > total ? 'over' : 'at') + ' its total capacity: ' + cnt.filled + ' of ' + total + ' seats are taken.' });
    if (block) blocking = true;
  }
  return { issues: issues, blocking: blocking };
}

/** finds the row being edited: by Enroll# (or S.No when she has no Enroll#) and her name as it was */
function admFindRow_(recs, ctx) {
  if (!ctx) return null;
  var roll = admNormRoll_(ctx.roll), sno = admS_(ctx.sno), nm = admNormName_(ctx.origName);
  var hit = recs.filter(function (r) {
    if (roll) return admNormRoll_(r.roll) === roll;
    return sno && admS_(r.sno) === sno && admNormName_(r.name) === nm;
  });
  if (hit.length > 1 && nm) hit = hit.filter(function (r) { return admNormName_(r.name) === nm; });
  return hit[0] || null;
}

function admValidate_(d, ctx, cur, recs, set) {
  var errors = [], issues = [], blocking = false;
  var p = admProg_(d.program), mig = admIsMig_(d.year), year = admYear_(d.year);
  if (!admS_(d.name)) errors.push('Student name is required.');
  if (!admS_(d.quota)) errors.push('Please choose a quota.');
  if (!admS_(d.section)) errors.push('Please choose a section.');
  var secs = (set.sections[p.key] || {})[year] || [];
  if (admS_(d.section) && !secs.some(function (s) { return admQ_(s) === admQ_(d.section); }))
    issues.push({ level: 'note', text: 'Section ' + d.section + ' is not in "Program Sections" for ' + p.label + ' ' + year + '.' });
  var roll = admS_(d.rollNo);
  if (!roll) errors.push('Roll No is required.');
  var others = recs.filter(function (r) { return !cur || r.row !== cur.row; });
  var label = function (r) { return (r.key ? admProg_(r.key).label : r.programName) + ' ' + r.year + ' ' + r.pcs; };

  if (roll) {
    var dr = others.filter(function (r) { return admNormRoll_(r.roll) === admNormRoll_(roll); })[0];
    if (dr) errors.push('Roll No ' + roll + ' is already used by ' + dr.name + ' (' + label(dr) + ', row ' + dr.row + ').');
    if (cur && admNormRoll_(cur.roll) !== admNormRoll_(roll) && cur.roll)
      issues.push({ level: 'warn', text: 'Enroll# changes from ' + cur.roll + ' to ' + roll + '. Her rows in the attendance sheets still carry the old number: run "Sheets vs strength sheet" on the College screen afterwards.' });
  }

  var moving = cur && (cur.key !== p.key || cur.year !== year);
  var sameSpot = cur && !moving && admQ_(cur.quota) === admQ_(d.quota);
  if (mig) {
    issues.push({ level: 'info', text: '2nd Year (migration) student: always ' + ADM_CFG.MIGRATION.QUOTA + ', not counted against the 1st-year seats.' });
  } else if (admS_(d.quota)) {
    if (sameSpot) issues.push({ level: 'info', text: 'Quota and program are unchanged, so seats are not re-checked.' });
    else if (!admQuotaAllowed_(p.key, admQ_(d.quota))) errors.push('"' + d.quota + '" is not available in ' + p.label + '. This program only has: ' + ADM_CFG.QUOTA_ONLY[p.key].join(', ') + '.');
    else { var ev = admEvaluate_(p.key, d.quota, others, set); issues = issues.concat(ev.issues); blocking = ev.blocking; }
  }
  if (moving) {
    issues.unshift({ level: 'warn', text: 'MOVE: ' + d.name + ' changes from ' + (cur.key ? admProg_(cur.key).label : cur.programName) + ' ' + cur.year + ' (' + cur.pcs + ') to ' + p.label + ' ' + year + ' (' + d.section + ').' });
    if (cur.key && admProg_(cur.key).fee !== p.fee)
      issues.push({ level: 'warn', text: 'Fee category changes from ' + (admProg_(cur.key).fee === 'non' ? 'non-subsidized' : 'subsidized') + ' to ' + (p.fee === 'non' ? 'non-subsidized' : 'subsidized') + '.' });
  } else if (cur && admQ_(cur.pcs) !== admQ_(d.section)) {
    issues.push({ level: 'warn', text: 'Section changes from ' + cur.pcs + ' to ' + d.section + '. Run "Sheets vs strength sheet" on the College screen so the attendance sheets follow.' });
  }

  if (roll) {
    if (mig) {
      var mn = admRollNum_(roll, ADM_CFG.MIGRATION.ROLL_PREFIX);
      if (mn === null && !cur) issues.push({ level: 'note', text: 'Roll No ' + roll + ' is not in the ' + ADM_CFG.MIGRATION.ROLL_PREFIX + 'NNNN format used for 2nd Year migration students.' });
      else if (mn !== null && mn < ADM_CFG.MIGRATION.ROLL_START && !cur) issues.push({ level: 'note', text: 'Roll No ' + roll + ' is below ' + ADM_CFG.MIGRATION.ROLL_START + ', where migration numbers start.' });
    } else {
      var n = admRollNum_(roll, ADM_CFG.ROLL_PREFIX), rg = ADM_CFG.ROLL_RANGES[p.key];
      if (n === null) issues.push({ level: 'note', text: 'Roll No ' + roll + ' is not in the ' + ADM_CFG.ROLL_PREFIX + 'NNN format.' });
      else if (rg && (n < rg[0] || n > rg[1])) issues.push({ level: 'note', text: 'Roll No ' + roll + ' is outside the range ' + rg[0] + ' to ' + rg[1] + ' for ' + p.label + '.' });
    }
  }

  var subs = (d.subjects || []).map(function (x) { return admS_(x).toUpperCase(); });
  var filled = subs.filter(Boolean);
  if (subs.slice(0, 4).filter(Boolean).length < 4) issues.push({ level: 'danger', text: 'The 4 compulsory subjects are not all filled in.' });
  else if (subs.slice(4, 7).filter(Boolean).length < 3) issues.push({ level: 'note', text: 'Optional subjects: ' + subs.slice(4, 7).filter(Boolean).length + ' of 3 filled. You can add the rest later.' });
  var dup = filled.filter(function (x, i) { return filled.indexOf(x) !== i; })[0];
  if (dup) issues.push({ level: 'danger', text: 'The subject ' + dup + ' is listed more than once.' });
  var wrongY = mig ? filled.filter(function (x) { return admYearSubject_(x, p.classProgram, year) !== x; }) : [];
  if (wrongY.length) issues.push({ level: 'danger', text: '1st Year subject in a 2nd Year record: ' + wrongY.map(function (x) { return x + ' (2nd Year: ' + admYearSubject_(x, p.classProgram, year) + ')'; }).join(', ') + '.' });

  if (admS_(d.id)) {
    var di = others.filter(function (r) { return admS_(r.id) === admS_(d.id); })[0];
    if (di) issues.push({ level: 'danger', text: 'Portal ID ' + d.id + ' already belongs to ' + di.name + ' (' + label(di) + ', ' + (di.roll || 'row ' + di.row) + ').' });
  }
  var nn = admNormName_(d.name), nf = admNormName_(d.father);
  if (nn) others.forEach(function (r) {
    if (admNormName_(r.name) !== nn) return;
    var same = nf && admNormName_(r.father) === nf;
    if (!same && nf && admNormName_(r.father)) return;
    if (same) issues.push({ level: 'danger', text: 'Possible duplicate: ' + r.name + ' (father ' + r.father + ') is already in ' + label(r) + ', ' + (r.roll || 'row ' + r.row) + '.' + (cur ? '' : ' If she is changing program or year, find and edit her instead of adding again.') });
    else issues.push({ level: 'note', text: 'Another student with the same name (' + r.name + ', father ' + (r.father || '-') + ') is in ' + label(r) + '. Check she is a different girl.' });
  });
  var m = admNum_(d.metric), t = admNum_(d.total);
  if (m !== '' && t !== '' && m > t) issues.push({ level: 'danger', text: 'Matric marks (' + m + ') are more than total marks (' + t + ').' });
  return { errors: errors, issues: issues, blocking: blocking };
}

/** prepares the record as it will be written (roll auto-assigned, subjects for the year, room) */
function admPrepare_(d, ctx, T) {
  var set = admSettings_(), recs = admAll_(T, set.sections), p = admProg_(d.program);
  d.year = admYear_(d.year);
  if (admIsMig_(d.year)) d.quota = ADM_CFG.MIGRATION.QUOTA;
  var cur = ctx ? admFindRow_(recs, ctx) : null;
  if (ctx && !cur) throw new Error('She was not found in ' + ADM_CFG.SHEET + ' any more (Enroll# ' + (ctx.roll || '-') + '). She may have been moved to Left Students. Search again.');
  var autoRoll = false;
  if (!cur && (!admS_(d.rollNo) || /^auto$/i.test(admS_(d.rollNo)))) {
    var nr = admNextRoll_(T, p.key, d.year);
    if (nr.error) throw new Error(nr.error);
    d.rollNo = nr.roll; autoRoll = true;
  }
  d.subjects = (d.subjects || []).slice(0, 7).map(function (x) { return admYearSubject_(x, p.classProgram, d.year); });
  if (!admS_(d.room) && d.section) d.room = admRoomMap_(recs)[admS_(d.section)] || '';
  return { set: set, recs: recs, p: p, cur: cur, autoRoll: autoRoll, v: admValidate_(d, ctx, cur, recs, set) };
}

function admCheck_(d, ctx) {
  var T = admRead_(), x = admPrepare_(d, ctx, T);
  return { ok: true, errors: x.v.errors, issues: x.v.issues, blocking: x.v.blocking, rollNo: d.rollNo, autoRoll: x.autoRoll,
           room: d.room, subjects: d.subjects, editing: !!x.cur };
}

/* ================= Write ================= */
/** values of every Transform column this form controls, keyed by column key */
function admValues_(d, p) {
  var v = {}, m = admNum_(d.metric), t = admNum_(d.total), hq = admS_(d.hq) || 'No';
  var year = admYear_(d.year), sec = admSecCode_(d.section), pcs = admS_(d.section);
  Object.keys(ADM_FIELDS).forEach(function (f) { if (f !== 'roll') v[ADM_FIELDS[f]] = admS_(d[f]); });
  v.enroll = admS_(d.rollNo); v.reg = admS_(d.rollNo);
  v.photo = v.picurl; v.metricrollno = v.matricrollno;
  v.obtainedmarks = m; v.totalmarks = t; v.hq = hq;
  v.program = p.program; v['class'] = year; v.section = sec; v.pcs = pcs;
  v.programclasssection = p.program + ' - ' + year + ' - ' + sec;
  v.appearedin = p.appeared;
  for (var i = 1; i <= 7; i++) v['sub' + i] = admS_((d.subjects || [])[i - 1]).toUpperCase();
  v.aggregate = admAgg_(m, t, hq);
  v.percentage = m !== '' && t ? Math.round((m / t) * 10000) / 100 : '';
  return v;
}
/** Aggregate = (Obtained Marks + 20 if Hafiz-e-Quran) x 100 / Total Marks, 2 decimals */
function admAgg_(m, t, hq) {
  m = admNum_(String(m).replace(/,/g, '')); t = admNum_(String(t).replace(/,/g, ''));
  if (m === '' || !t) return '';
  return Math.round((m + (/^y/i.test(admS_(hq)) ? 20 : 0)) * 10000 / t) / 100;
}

/* ================= Update from portal: quota, HQ, Aggregate, MeritNo, Room for every student =================
   The portal sync only fills EMPTY cells. This rewrites these five columns for everybody:
   quota  - portal Category (when it clearly names one of the quotas; 2nd Year migration stays Open Quota)
   hq     - portal "HQ" (Hafiz-e-Quran) when the portal copy has it
   aggregate - (Obtained + 20 if HQ) x 100 / Total
   merit  - 1st Year, per admission program and per quota: students by Aggregate, highest = 1
   room   - the section's room in the Timetable sheet                                                                    */
var ADM_REFRESH = { portal: 'Portal', hq: 'HQ', aggregate: 'Aggregate', merit: 'MeritNo', room: 'Room', agewise: 'Age-wise', feepaid: 'FEE PAID DATE' };
/* every Transform column that has a matching value on the portal: f = field of the record (admRec_) and of the portal student (admPortalStu_),
   t = Transform column keys written, type = how values are compared, info = shown only (the college decides it, change it in Find & edit) */
var ADM_CMP = [
  { f: 'name', label: 'Student Name', t: ['studentname'], type: 'name' },
  { f: 'father', label: 'Father Name', t: ['fathername'], type: 'name' },
  { f: 'quota', label: 'Quota', t: ['quota'], type: 'quota' },
  { f: 'dob', label: 'DOB', t: ['dob'], type: 'date' },
  { f: 'contact', label: 'Student Contact', t: ['studentcontact'], type: 'phone' },
  { f: 'email', label: 'Email', t: ['email'], type: 'name' },
  { f: 'pic', label: 'Picture', t: ['picurl', 'photo'], type: 'text' },
  { f: 'board', label: 'Board', t: ['board'], type: 'name' },
  { f: 'passYear', label: 'Passing Year', t: ['passingyear'], type: 'num' },
  { f: 'matricRoll', label: 'Matric Roll No', t: ['matricrollno', 'metricrollno'], type: 'num' },
  { f: 'metric', label: 'Obtained Marks', t: ['obtainedmarks'], type: 'num' },
  { f: 'total', label: 'Total Marks', t: ['totalmarks'], type: 'num' },
  { f: 'challan', label: 'Challan Details', t: ['challandetails'], type: 'text' },
  { f: 'batch', label: 'Batch', t: ['batch'], type: 'name' },
  { f: 'subjects', label: 'Subjects', t: ['sub1', 'sub2', 'sub3', 'sub4', 'sub5', 'sub6', 'sub7'], type: 'subjects' },
  { f: 'id', label: 'Portal ID', t: ['portalid'], type: 'text' },
  { f: 'roll', label: 'Enroll#', type: 'roll', info: true },
  { f: 'programName', label: 'Program', type: 'program', info: true },
  { f: 'year', label: 'Class', type: 'name', info: true },
  { f: 'section', label: 'Section', type: 'section', info: true },
  { f: 'status', label: 'Status', type: 'name', info: true }
];
function admYmd_(v) {
  var s = admS_(v), m;
  if ((m = s.match(/^(\d{1,2})[\/\-. ](\d{1,2})[\/\-. ](\d{4})/))) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  if ((m = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/))) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  var d = new Date(s); return isNaN(d) ? admQ_(s) : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
/** true when the two values mean the same thing */
function admSame_(type, a, b) {
  a = admS_(a); b = admS_(b);
  if (type === 'num') { var x = admNum_(a.replace(/,/g, '')), y = admNum_(b.replace(/,/g, '')); return x !== '' && y !== '' ? Math.abs(x - y) < 0.001 : admQ_(a) === admQ_(b); }
  if (type === 'date') return admYmd_(a) === admYmd_(b);
  if (type === 'phone') { var p = a.replace(/\D/g, '').slice(-10), q = b.replace(/\D/g, '').slice(-10); return p === q; }
  if (type === 'roll') return admNormRoll_(a) === admNormRoll_(b);
  if (type === 'program') return admH_(a).replace(/jinnahcollege/, 'jcw') === admH_(b).replace(/jinnahcollege/, 'jcw');
  if (type === 'section') return admSecCode_(a).toUpperCase() === admSecCode_(b).toUpperCase() || a.toUpperCase() === b.toUpperCase();
  if (type === 'text') return a.replace(/\s+/g, '') === b.replace(/\s+/g, '');
  return admH_(a) === admH_(b);
}
function admRefresh_(user, fields, apply, accept) {
  accept = accept || {};              // 'row|field' -> yes (quota: the quota you chose). Nothing taken from the portal without your answer.
  fields = (fields && fields.length ? fields : Object.keys(ADM_REFRESH)).map(function (f) { return f === 'quota' ? 'portal' : f; })
    .filter(function (f, i, a) { return ADM_REFRESH[f] && a.indexOf(f) === i; });
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var set = admSettings_(), T = admRead_(), recs = admAll_(T, set.sections), C = T.cols, sh = T.sh;
    var colOf = { portal: -1, hq: C.hq, aggregate: C.aggregate, merit: C.meritno, room: C.room, agewise: C.agewise, feepaid: C.feepaiddate };
    var notes = [], skip = {}, newAgeCol = false;
    if (fields.indexOf('merit') >= 0 && fields.indexOf('agewise') < 0) fields.push('agewise');       // v7.6.3: Age-wise goes with MeritNo
    if (fields.indexOf('agewise') >= 0 && colOf.agewise === undefined) { colOf.agewise = T.width; newAgeCol = true; notes.push('A new column "Age-wise" is added at the end of the strength sheet.'); }
    // v7.6.9: FEE PAID DATE (first real fee paid, from the portal copy) goes with "Every column vs portal"; the column is added when missing
    var newFeeCol = false;
    if (fields.indexOf('portal') >= 0 && fields.indexOf('feepaid') < 0) fields.push('feepaid');
    if (fields.indexOf('feepaid') >= 0 && colOf.feepaid === undefined) { colOf.feepaid = T.width + (newAgeCol ? 1 : 0); newFeeCol = true; notes.push('A new column "FEE PAID DATE" is added at the end of the strength sheet.'); }
    fields.forEach(function (f) { if (colOf[f] === undefined) { skip[f] = true; notes.push('No "' + ADM_REFRESH[f] + '" column in the strength sheet.'); } });
    // columns made by a formula are never overwritten
    var n = T.rows.length;
    if (n) fields.forEach(function (f) {
      if (skip[f] || colOf[f] < 0 || (f === 'agewise' && newAgeCol) || (f === 'feepaid' && newFeeCol)) return;
      var fm = sh.getRange(2, colOf[f] + 1, n, 1).getFormulas();
      if (fm.some(function (x) { return x[0]; })) { skip[f] = true; notes.push('"' + ADM_REFRESH[f] + '" is made by a formula in the sheet, so it was not changed.'); }
    });
    // portal copy by portal ID and by Enroll#
    var P = admPortalSheet_(), byId = {}, byRoll = {};
    if (P) P.v.slice(1).forEach(function (r) {
      var id = P.g(r, 'studentid'), en = P.g(r, 'enrollment');
      if (id) byId[id] = r; if (en) byRoll[admNormRoll_(en)] = r;
    });
    else notes.push('The "' + ADM_CFG.PORTAL_SHEET + '" sheet is empty: quota and HQ need a portal sync first.');
    var portalHasHq = !!P && P.head.some(function (h) { return /^(hq|hafiz.*|hifz.*)$/i.test(admS_(h).replace(/[^a-z]/gi, '')); });
    if (P && !portalHasHq && fields.indexOf('hq') >= 0) notes.push('The portal copy has no HQ (Hafiz-e-Quran) column yet: update the Tampermonkey script to 1.6 and run a Full sync. HQ already in the sheet is kept.');
    var hqOf = function (r) {
      var h = '';
      P.head.forEach(function (x, j) { if (!h && /^(hq|hafiz.*|hifz.*)$/i.test(admS_(x).replace(/[^a-z]/gi, ''))) h = admS_(r[j]); });
      if (!h) return '';
      return /^(y|yes|true|1|hafiz)/i.test(h) ? 'Yes' : /^(n|no|false|0)/i.test(h) ? 'No' : '';
    };
    // v7.6: every other strength-sheet column that the portal copy also has (same header, or "Portal <header>") is compared too
    var covered = {}; ADM_CMP.forEach(function (d) { (d.t || []).forEach(function (k) { covered[k] = 1; }); });
    var pHas = {}; if (P) P.head.forEach(function (h) { var k = admH_(h); if (k) pHas[k.replace(/^portal/, '')] = 1; });
    var CMPS = ADM_CMP.concat(Object.keys(C).filter(function (k) {
      return !covered[k] && pHas[k] && k !== 'feepaiddate' && !/^(s?no|sr|srno|serial|enroll.*|rollno|roll|program.*|class|section|pcs|status|category|session|aggregate|merit.*|room.*|quota|year|group.*|key|id|sub\d.*)$/.test(k);
    }).map(function (k) { return { f: 'x_' + k, label: admS_(T.head[C[k]]), t: [k], type: 'name', raw: true }; }));
    var names = admQuotaNames_(admCounts_(recs)), rooms = admRooms_();
    var nv = {}, changes = [], cnt = {}, unknownCat = {}, ask = [], quotaOk = {}, gw = {}, fcol = {};
    var formulaCol = function (j) {           // portal values are never written over a formula
      if (!(j in fcol)) fcol[j] = n ? sh.getRange(2, j + 1, n, 1).getFormulas().some(function (x) { return x[0]; }) : false;
      return fcol[j];
    };
    Object.keys(names).forEach(function (k) { quotaOk[admQ_(names[k])] = names[k]; });
    fields.forEach(function (f) { cnt[f] = 0; });
    var put = function (r, f, val) {
      if (skip[f]) return;
      var old = admS_(T.rows[r.row - 2][colOf[f]]), nw = admS_(val);
      if (f === 'aggregate' && old !== '' && nw !== '' && Math.abs(Number(old) - Number(nw)) < 0.005) return;
      if (old === nw) return;
      (nv[f] = nv[f] || {})[r.row] = val;
      cnt[f]++;
      if (f !== 'quota' && changes.length < 400) changes.push({ roll: r.roll, name: r.name, pcs: r.pcs, field: ADM_REFRESH[f], from: old, to: nw });
    };
    recs.forEach(function (r) {
      var pr = P ? (r.id && byId[r.id]) || (r.roll && byRoll[admNormRoll_(r.roll)]) : null;
      r._hq = r.hq;
      if (fields.indexOf('portal') >= 0 && pr) {
        var ps = admPortalStu_(P, pr, set);
        CMPS.forEach(function (d) {
          var key = r.row + '|' + d.f, now, pv, sug = '';
          if (d.raw) {
            pv = P.g(pr, d.t[0]) || P.g(pr, 'portal' + d.t[0]); if (!pv) return;
            now = admS_(T.rows[r.row - 2][C[d.t[0]]]);
            if (/^hq$/.test(d.t[0])) pv = /^(y|yes|true|1|hafiz)/i.test(pv) ? 'Yes' : /^(n|no|false|0)/i.test(pv) ? 'No' : pv;
            if (admSame_(d.type, now, pv)) return;
            sug = pv;
          } else if (d.f === 'quota') {
            if (admIsMig_(r.year)) return;
            pv = P.g(pr, 'category'); if (!pv) return;
            var gq = admGuessQuota_(pv, r.key, names);
            if (gq.sure && admQ_(gq.quota) === admQ_(r.quota)) return;
            if (!gq.sure) unknownCat[pv] = (unknownCat[pv] || 0) + 1;
            now = r.quota; sug = gq.sure ? gq.quota : '';
          } else if (d.f === 'subjects') {
            var mine = [1, 2, 3, 4, 5, 6, 7].map(function (k) { return admS_(r['s' + k]).toUpperCase(); });
            var cp = admProg_(r.key || 'FCS');
            var theirs = (ps.subjects || []).map(function (x) { return admYearSubject_(x, r.key ? cp.classProgram : '', r.year); });
            if (theirs.filter(Boolean).length < 4) return;
            var A = mine.filter(Boolean).slice().sort().join(', '), B = theirs.filter(Boolean).slice().sort().join(', ');
            if (A === B) return;
            now = mine.filter(Boolean).join(', '); pv = theirs.filter(Boolean).join(', '); sug = theirs;
          } else {
            pv = admS_(ps[d.f]); if (!pv) return;
            now = admS_(r[d.f]);
            if (d.f === 'id' && (!r.roll || byId[r.id])) return;         // only when matched by Enroll#
            if (admSame_(d.type, now, pv)) return;
            sug = pv;
          }
          var cols = (d.t || []).map(function (k) { return C[k]; }).filter(function (c) { return c !== undefined; });
          var info = d.info || !cols.length || cols.some(formulaCol);
          if (ask.length < 4000) ask.push({ key: key, row: r.row, f: d.f, label: d.label, roll: r.roll, name: r.name, pcs: r.pcs, now: now, portal: d.f === 'quota' ? pv : pv,
                     suggest: d.f === 'subjects' ? pv : sug, info: info ? (d.info ? 'set by the college: change it in Find & edit (or on the portal)' : 'no such column, or it is a formula') : '' });
          var ans = accept[key] || (accept['*|' + d.f] && sug ? (d.f === 'quota' ? sug : 'yes') : '');    // '*|field' = Accept the whole column
          if (!ans || info) return;
          if (d.f === 'quota') {
            if (!quotaOk[admQ_(ans)]) return;
            r._leftOpen = admQ_(r.quota) === admQ_(ADM_CFG.OPEN_QUOTA) && admQ_(ans) !== admQ_(ADM_CFG.OPEN_QUOTA);
            sug = quotaOk[admQ_(ans)]; r.quota = sug;
          }
          var vals = d.f === 'subjects' ? sug.concat(['', '', '', '', '', '', '']).slice(0, 7) : null;
          (d.t || []).forEach(function (k, ix) {
            if (C[k] === undefined) return;
            (gw[C[k]] = gw[C[k]] || {})[r.row] = vals ? vals[ix] : sug;
          });
          cnt.portal++;
          if (d.f === 'metric' || d.f === 'total') r[d.f] = sug;               // the aggregate uses the new marks
        });
      }
      if (fields.indexOf('feepaid') >= 0 && pr) { var fpd = P.g(pr, 'feepaiddate'); if (fpd) put(r, 'feepaid', fpd); }
      if (fields.indexOf('hq') >= 0 && pr && portalHasHq) { var h = hqOf(pr); if (h) { put(r, 'hq', h); r._hq = h; } }
      if (!r._hq && C.hq !== undefined) r._hq = r.hq;
      if (fields.indexOf('aggregate') >= 0) put(r, 'aggregate', admAgg_(r.metric, r.total, r._hq));
      r._agg = admAgg_(r.metric, r.total, r._hq);
      if (fields.indexOf('room') >= 0 && r.pcs && rooms[r.pcs]) put(r, 'room', rooms[r.pcs]);
    });
    if (fields.indexOf('merit') >= 0) {
      var groups = {};
      // v7.6.3: equal Aggregate = the same MeritNo (1, 2, 2, 2, 3 … nothing skipped); among them Age-wise 1st, 2nd … (older first, from DOB)
      // v7.6.2: merit inside every program AND every quota separately (Pre-Medical Open (S) · UET Quota, Self Morning · Open Quota …),
      // by Aggregate, highest = 1. Students with no Aggregate or no quota get no merit number.
      recs.forEach(function (r) {
        if (admIsMig_(r.year) || !r.key || !admS_(r.quota) || r._agg === '' || isNaN(Number(r._agg))) { if (r.meritNo && !admIsMig_(r.year)) put(r, 'merit', ''); return; }
        var g = r.key + '|' + admQ_(r.quota);
        (groups[g] = groups[g] || []).push(r);
      });
      var ordinal = function (n) { var t = n % 100, u = n % 10; return n + (t >= 11 && t <= 13 ? 'th' : u === 1 ? 'st' : u === 2 ? 'nd' : u === 3 ? 'rd' : 'th'); };
      var dobKey = function (r) { var d = admYmd_(r.dob); return /^\d{4}-\d\d-\d\d$/.test(d) ? d : '9999-99-99'; };     // no DOB = youngest
      var aggOf = function (r) { return Math.round(Number(r._agg) * 100); };
      Object.keys(groups).forEach(function (k) {
        var L = groups[k].sort(function (a, b) {
          return aggOf(b) - aggOf(a) || dobKey(a).localeCompare(dobKey(b)) || String(a.name).localeCompare(String(b.name));
        });
        var rank = 0, tie = 0;
        L.forEach(function (r, i) {
          if (i === 0 || aggOf(r) !== aggOf(L[i - 1])) { rank++; tie = 0; }      // dense: 1, 2, 2, 2, 3 (no number skipped)
          tie++;
          put(r, 'merit', rank);
          var same = L.filter(function (o) { return aggOf(o) === aggOf(r); }).length;   // v7.6.5: Age-wise only where students share an Aggregate
          put(r, 'agewise', same > 1 ? ordinal(tie) : '');
        });
      });
    }
    Object.keys(unknownCat).forEach(function (c) { notes.push(unknownCat[c] + ' student(s) with portal category "' + c + '" (not a known quota): choose their quota yourself or keep it.'); });
    if (apply) {
      Object.keys(gw).forEach(function (j) {
        var rng = sh.getRange(2, Number(j) + 1, n, 1), cur = rng.getValues(), txt = Number(j) === C.dob || Number(j) === C.studentcontact;
        if (txt) rng.setNumberFormat('@');
        Object.keys(gw[j]).forEach(function (row) { cur[Number(row) - 2][0] = gw[j][row]; });
        rng.setValues(cur);
      });
      if (newFeeCol && nv.feepaid) {
        if (sh.getMaxColumns() < colOf.feepaid + 1) sh.insertColumnsAfter(sh.getMaxColumns(), colOf.feepaid + 1 - sh.getMaxColumns());
        sh.getRange(1, colOf.feepaid + 1).setValue('FEE PAID DATE').setFontWeight('bold');
        sh.getRange(2, colOf.feepaid + 1, Math.max(1, n), 1).setNumberFormat('@');
      }
      if (newAgeCol && nv.agewise) {
        if (sh.getMaxColumns() < colOf.agewise + 1) sh.insertColumnsAfter(sh.getMaxColumns(), colOf.agewise + 1 - sh.getMaxColumns());
        sh.getRange(1, colOf.agewise + 1).setValue('Age-wise').setFontWeight('bold');
      }
      fields.forEach(function (f) {
        if (f === 'portal') return;
        if (skip[f] || !nv[f]) return;
        var rng = sh.getRange(2, colOf[f] + 1, n, 1), cur = rng.getValues();
        Object.keys(nv[f]).forEach(function (row) { cur[Number(row) - 2][0] = nv[f][row]; });
        rng.setValues(cur);
      });
      SpreadsheetApp.flush();
      cnt.portalCells = Object.keys(gw).reduce(function (a, j) { return a + Object.keys(gw[j]).length; }, 0);
      var total = Object.keys(cnt).reduce(function (a, f) { return a + cnt[f]; }, 0);
      admLog_(user, 'UPDATE FROM PORTAL', '', '', fields.map(function (f) { return ADM_REFRESH[f] + ' ' + cnt[f]; }).join(', ') + ' (' + total + ' cells)', '');
    }
    return { ok: true, applied: apply, counts: cnt, fields: fields, changes: changes, ask: ask,
             quotas: Object.keys(quotaOk).map(function (k) { return quotaOk[k]; }).sort(), notes: notes, skipped: Object.keys(skip), students: recs.length,
             seats: apply ? admSeatState_(admAll_(admRead_(), set.sections), set) : null };
  } finally { lock.releaseLock(); }
}
/** run once from the Apps Script editor: updates these columns every hour by itself (run again to switch it off) */
function admSetupAutoRefresh() {
  var had = ScriptApp.getProjectTriggers().filter(function (t) { return t.getHandlerFunction() === 'admAutoRefresh'; });
  had.forEach(function (t) { ScriptApp.deleteTrigger(t); });
  if (had.length) { Logger.log('Hourly update from portal switched OFF.'); return; }
  ScriptApp.newTrigger('admAutoRefresh').timeBased().everyHours(1).create();
  Logger.log('Hourly update from portal switched ON (quota, HQ, Aggregate, MeritNo, Room).');
}
function admAutoRefresh() { admRefresh_({ teacher: 'Hourly update', username: 'auto' }, ['hq', 'aggregate', 'merit', 'room'], true, {}); }   // quota only with your answer

/** which columns a change of form field rewrites */
var ADM_DERIVED = {
  rollNo: ['enroll', 'reg'], program: ['program', 'programclasssection', 'appearedin', 'section', 'pcs', 'class'],
  year: ['class', 'programclasssection', 'section', 'pcs'], section: ['section', 'pcs', 'programclasssection'],
  pic: ['picurl', 'photo'], matricRoll: ['matricrollno', 'metricrollno'],
  metric: ['obtainedmarks', 'aggregate', 'percentage'], total: ['totalmarks', 'aggregate', 'percentage'], hq: ['hq', 'aggregate']
};

function admLogSheet_() {
  var book = admBook_(), sh = book.getSheetByName(ADM_CFG.LOG_SHEET);
  if (!sh) {
    sh = book.insertSheet(ADM_CFG.LOG_SHEET);
    sh.getRange(1, 1, 1, 7).setValues([['Time', 'User', 'Action', 'Enroll#', 'Student', 'Details', 'Client ID']]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function admLog_(user, action, roll, name, details, cid) {
  try {
    admLogSheet_().appendRow([new Date(), (user.teacher || '') + ' <' + user.username + '>', action, roll, name, details, cid || '']);
  } catch (e) { /* logging never stops a save */ }
}
/** an offline entry that already reached the server (the phone lost the reply) is not saved twice */
function admSeen_(cid) {
  if (!cid) return null;
  var hit = CacheService.getScriptCache().get('admcid_' + cid);
  if (hit) return JSON.parse(hit);
  try {
    var sh = admBook_().getSheetByName(ADM_CFG.LOG_SHEET);
    if (!sh || sh.getLastRow() < 2) return null;
    var n = Math.min(sh.getLastRow() - 1, 1000);
    var v = sh.getRange(sh.getLastRow() - n + 1, 1, n, 7).getDisplayValues();
    for (var i = v.length - 1; i >= 0; i--) if (v[i][6] === cid) return { roll: v[i][3], name: v[i][4] };
  } catch (e) {}
  return null;
}

function admSave_(user, d, ctx, override, cid) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var seen = admSeen_(cid);
    if (seen) return { ok: true, duplicate: true, roll: seen.roll, message: seen.name + ' was already saved (' + seen.roll + ').' };
    var T = admRead_(), x = admPrepare_(d, ctx, T);
    if (x.v.errors.length) return { ok: false, errors: x.v.errors, issues: x.v.issues };
    if (x.v.blocking && !override) return { ok: false, blocked: true, issues: x.v.issues };
    if (!override && x.v.issues.some(function (i) { return i.level === 'danger'; })) return { ok: false, needsConfirm: true, issues: x.v.issues, rollNo: d.rollNo };
    var vals = admValues_(d, x.p), C = T.cols, sh = T.sh, row, action, details;

    if (!x.cur) {
      // NEW admission: one row appended at the bottom
      var mx = 0;
      if (C.sno !== undefined) T.rows.forEach(function (r) { var n = Number(r[C.sno]); if (n > mx) mx = n; });
      vals.sno = mx + 1; vals.session = ADM_CFG.SESSION; vals.status = ADM_CFG.STATUS;
      if (!vals.photo) vals.photo = vals.picurl;
      var line = T.head.map(function (h) { var k = admH_(h); return vals[k] === undefined || vals[k] === null ? '' : vals[k]; });
      row = sh.getLastRow() + 1;
      if (C.dob !== undefined) sh.getRange(row, C.dob + 1).setNumberFormat('@');
      if (C.studentcontact !== undefined) sh.getRange(row, C.studentcontact + 1).setNumberFormat('@');
      sh.getRange(row, 1, 1, line.length).setValues([line]);
      action = admIsMig_(d.year) ? 'ADD (2nd Year migration)' : 'ADD';
      details = x.p.label + ', ' + d.year + ', ' + d.section + ', ' + d.quota + (d.migratedFrom ? ', migrated from ' + d.migratedFrom : '') + (override ? ' (saved anyway)' : '');
    } else {
      // EDIT: only the cells whose form value changed are written, so portal-sync data is never overwritten
      row = x.cur.row;
      var o = ctx.orig || {}, cols = {}, changed = [];
      var was = function (f) { return f === 'subjects' ? (o.subjects || []).map(function (s) { return admS_(s).toUpperCase(); }).join('|') : admS_(o[f]); };
      var now = function (f) { return f === 'subjects' ? (d.subjects || []).map(function (s) { return admS_(s).toUpperCase(); }).join('|') : admS_(d[f]); };
      Object.keys(ADM_FIELDS).concat(['rollNo', 'program', 'year', 'section', 'subjects']).forEach(function (f) {
        if (f === 'roll' || was(f) === now(f)) return;
        changed.push(f);
        if (ADM_DERIVED[f]) ADM_DERIVED[f].forEach(function (k) { cols[k] = true; });
        else if (f === 'subjects') for (var i = 1; i <= 7; i++) cols['sub' + i] = true;
        else cols[ADM_FIELDS[f]] = true;
      });
      if (cols.program || cols['class']) cols.appearedin = true;
      if (!changed.length) return { ok: true, nothing: true, roll: x.cur.roll, rec: admPack_(x.cur), message: 'Nothing was changed.' };
      Object.keys(cols).forEach(function (k) {
        if (C[k] === undefined) return;
        var cell = sh.getRange(row, C[k] + 1);
        if (k === 'dob' || k === 'studentcontact') cell.setNumberFormat('@');
        cell.setValue(vals[k] === undefined || vals[k] === null ? '' : vals[k]);
      });
      action = (x.cur.key !== x.p.key || x.cur.year !== admYear_(d.year)) ? 'MOVE' : 'EDIT';
      details = 'changed: ' + changed.join(', ') + (cols.quota ? ' (quota was ' + x.cur.quota + ')' : '') + (cols.pcs ? ' (section was ' + x.cur.pcs + ')' : '') + (override ? ' (saved anyway)' : '');
    }
    SpreadsheetApp.flush();
    admLog_(user, action, d.rollNo, d.name, details, cid);
    if (cid) CacheService.getScriptCache().put('admcid_' + cid, JSON.stringify({ roll: d.rollNo, name: d.name }), 21600);

    var T2 = admRead_(), recs2 = admAll_(T2, x.set.sections);
    var rec = recs2.filter(function (r) { return r.row === row; })[0];
    return { ok: true, action: action, roll: d.rollNo, autoRoll: x.autoRoll, row: row, rec: rec ? admPack_(rec) : null,
             seats: admSeatState_(recs2, x.set), next: x.cur ? null : admNextRoll_(T2, x.p.key, d.year),
             message: d.name + (x.cur ? ' updated' : ' added') + ' in ' + ADM_CFG.SHEET + ' (' + d.rollNo + ', row ' + row + ').' };
  } finally { lock.releaseLock(); }
}
