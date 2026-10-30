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
 *   2) In doPost, right after the line  if (!user) return json_({ ok: false, error: 'BAD_LOGIN' });  add:
 *        if (/^adm_/.test(String(req.action || ''))) return json_(admRoute_(user, req));
 *   3) Deploy > Manage deployments > Edit > Version: New version > Deploy.
 *
 * Who can use it: admin logins (isAdmin_), plus the usernames listed in ADM_CFG.EXTRA_USERS
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
  email: 'email', board: 'board', passYear: 'passingyear', matricRoll: 'matricrollno', challan: 'challandetails'
};
/* Fields sent to the phone for every student (offline search / edit) */
var ADM_INDEX = ['sno', 'roll', 'name', 'father', 'id', 'key', 'year', 'pcs', 'quota', 's1', 's2', 's3', 's4', 's5', 's6', 's7',
  'dob', 'metric', 'total', 'hq', 'room', 'pic', 'contact', 'email', 'board', 'passYear', 'matricRoll', 'challan'];

/* ================= Router (called from doPost) ================= */
function admRoute_(user, req) {
  if (!admCanUse_(user)) return { ok: false, error: 'ADM_NO_ACCESS' };
  switch (String(req.action)) {
    case 'adm_init':      return admInit_();
    case 'adm_index':     return admIndex_();
    case 'adm_next_roll': return admNextRollFor_(req.program, req.year);
    case 'adm_portal':    return admPortal_(req.id);
    case 'adm_check':     return admCheck_(req.d || {}, req.ctx || null);
    case 'adm_save':      return admSave_(user, req.d || {}, req.ctx || null, req.override === true, String(req.cid || ''));
  }
  return { ok: false, error: 'UNKNOWN_ACTION' };
}

function admCanUse_(user) {
  try { if (typeof isAdmin_ === 'function' && isAdmin_(user)) return true; } catch (e) {}
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
    matricRoll: g('matricrollno'), challan: g('challandetails') };
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
           threshold: ADM_CFG.THRESHOLD, mig: ADM_CFG.MIGRATION, sheet: ADM_CFG.SHEET },
    note: set.note, at: admNow_()
  };
}
function admPack_(rec) { return ADM_INDEX.map(function (k) { return rec[k] == null ? '' : String(rec[k]); }); }
function admIndex_() {
  var set = admSettings_(), T = admRead_();
  return { ok: true, cols: ADM_INDEX, rows: admAll_(T, set.sections).map(admPack_), at: admNow_() };
}

/* ================= Fill from the portal copy ("Portal Data" sheet, kept fresh by the UOP portal sync) ================= */
function admPortal_(id) {
  var q = admS_(id).toUpperCase();
  if (!q) return { ok: false, error: 'Type the portal Student ID or Enrollment# first.' };
  var sh = admBook_().getSheetByName(ADM_CFG.PORTAL_SHEET);
  if (!sh || sh.getLastRow() < 2) return { ok: false, error: 'The "' + ADM_CFG.PORTAL_SHEET + '" sheet is empty. Run the UOP portal sync first.' };
  var v = sh.getDataRange().getDisplayValues(), c = {};
  v[0].forEach(function (h, j) { var k = admH_(h); if (k && !(k in c)) c[k] = j; });
  var g = function (r, k) { return c[k] === undefined ? '' : admS_(r[c[k]]); };
  var row = null;
  for (var i = 1; i < v.length; i++) {
    if (g(v[i], 'studentid').toUpperCase() === q || admNormRoll_(g(v[i], 'enrollment')) === admNormRoll_(q)) { row = v[i]; break; }
  }
  if (!row) return { ok: false, error: 'Student ' + id + ' is not in "' + ADM_CFG.PORTAL_SHEET + '". If she was added on the portal today, run a quick portal sync and try again.' };
  var num = function (x) { var n = admNum_(String(x).replace(/,/g, '')); return n === '' || n === 0 ? '' : String(Math.round(n * 100) / 100); };
  var yr = admYear_(g(row, 'class')), sub = [];
  for (var s = 1; s <= 7; s++) sub.push(g(row, 'sub' + s + (admIsMig_(yr) ? '2ndyear' : '1styear')) || g(row, 'sub' + s + '1styear'));
  var progName = g(row, 'program'), key = admKeyOf_(progName, yr, '', g(row, 'enrollment'), admSettings_().sections);
  return { ok: true, p: {
    id: g(row, 'studentid'), portalRoll: g(row, 'enrollment'), name: g(row, 'studentname'), father: g(row, 'fathername'),
    quota: g(row, 'category'), dob: g(row, 'dob'), contact: g(row, 'studentcontact'), email: g(row, 'emailaddress'),
    pic: g(row, 'photoname'), board: g(row, 'boarduniversity'), passYear: g(row, 'passingyear'), matricRoll: num(g(row, 'rollno')),
    metric: num(g(row, 'obtainmarkscgpa')), total: num(g(row, 'totalmarkscgpa')), challan: g(row, 'challandetails'),
    subjects: sub, program: key, programName: progName, year: yr, section: g(row, 'section'), status: g(row, 'status'),
    synced: g(row, 'detailssynced') } };
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
  v.obtainedmarks = m; v.totalmarks = t; v.hq = hq;
  v.program = p.program; v['class'] = year; v.section = sec; v.pcs = pcs;
  v.programclasssection = p.program + ' - ' + year + ' - ' + sec;
  v.appearedin = p.appeared;
  for (var i = 1; i <= 7; i++) v['sub' + i] = admS_((d.subjects || [])[i - 1]).toUpperCase();
  v.aggregate = m !== '' && t ? Math.round(((m + (hq === 'Yes' ? 20 : 0)) / t) * 100) : '';
  v.percentage = m !== '' && t ? Math.round((m / t) * 10000) / 100 : '';
  return v;
}
/** which columns a change of form field rewrites */
var ADM_DERIVED = {
  rollNo: ['enroll', 'reg'], program: ['program', 'programclasssection', 'appearedin', 'section', 'pcs', 'class'],
  year: ['class', 'programclasssection', 'section', 'pcs'], section: ['section', 'pcs', 'programclasssection'],
  metric: ['obtainedmarks', 'aggregate', 'percentage'], total: ['totalmarks', 'aggregate', 'percentage'], hq: ['hq', 'aggregate']
};

function admLog_(user, action, roll, name, details, cid) {
  try {
    var book = admBook_(), sh = book.getSheetByName(ADM_CFG.LOG_SHEET);
    if (!sh) {
      sh = book.insertSheet(ADM_CFG.LOG_SHEET);
      sh.getRange(1, 1, 1, 7).setValues([['Time', 'User', 'Action', 'Enroll#', 'Student', 'Details', 'Client ID']]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), (user.teacher || '') + ' <' + user.username + '>', action, roll, name, details, cid || '']);
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
    if (x.v.blocking) return { ok: false, blocked: true, issues: x.v.issues };
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
