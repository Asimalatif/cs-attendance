/**
 * RecordSummary.gs — JCW RECORD SUMMARY inside the attendance app (page recordsummary.html).
 * Record Table, Efficiency Charts, Not Marking and Month-wise: which section + subject has its attendance,
 * monthly tests, grand test and detention marked, read from the record workbooks (schedule-aware, cached ticks).
 * The code below the router is the record summary script unchanged; the page now calls it through doPost
 * with the app login instead of google.script.run.
 *
 * Who can open it: admin, principal, HOD and Chief Proctor (the same people who see the college overview).
 *
 * INSTALL:
 *   a) + > Script > name it "RecordSummary" > paste this file.
 *      (If the old record summary script is ALSO in this project, delete that old file: the names are the same.)
 *   b) doPost, next to the shr_ line (once):
 *        if (/^rsm_/.test(String(req.action || ''))) return json_(rsmRoute_(user, req));
 *   c) Run debugRecordFile once from the editor (asks for Drive / Sheets permission).
 *   d) Deploy > Manage deployments > Edit > New version > Deploy.
 */
function rsmRoute_(user, req) {
  try {
    var all = false;
    try { all = wrAll_(user); } catch (e) {}
    if (!all) { try { all = !!isCP_(user); } catch (e) {} }
    if (!all) return { ok: false, error: 'The record summary is for the admin, principal, HODs and Chief Proctors.' };
    switch (String(req.action || '')) {
      case 'rsm_meta':     return { ok: true, data: getRecordSummaryMeta() };
      case 'rsm_files':    return { ok: true, data: getRecordFileList(!!req.force) };
      case 'rsm_file':     return { ok: true, data: getRecordFile(String(req.id || ''), req.cached || null, Number(req.mod) || 0) };
      case 'rsm_teachers': return { ok: true, data: getRecordTeachers() };
      case 'rsm_live':     return { ok: true, data: rs_live_() };
    }
    return { ok: false, error: 'UNKNOWN_ACTION' };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}

/* =============================== JCW RECORD SUMMARY (schedule-aware, modified-time cache) =============================== */

const RS_STUDENT_SS     = "1ARU8zMO2daJUfU_SIIo4baeSOG7qBPuni5TYRIRLBNY";
const RS_STRENGTH_TAB   = "Transform-ALL-COLLEGE-STRENGTH";
const RS_GT_SS          = "1FIT93zzKQnwMMjbaGy0aqrdCMv2QcM1ZTrRBh4SLKM8";
const RS_MONTHLY_FOLDER = "15sdpVNAkiQg1-heWIvtdZCczwt253WhK";
const RS_ATT_FOLDER     = "1FNJz2Tkn1tZd6umoByvsVv8-sAYDxEaF";
const RS_DET_FOLDER     = "1INQTwJ9484bguVFlPQT4QSByqhoruu2v";

// Extra rows that are NOT in the PCS column of the strength sheet (S1 & T1 are Pre-Medical sections)
// until: last month the section existed; skip: record types it never has (att / mon / grand / det)
// those cells stay blank and are never counted as pending
const RS_EXTRA_SECTIONS = [
  { pcs: "T1", key: "T1", until: "September", skip: ["mon"], subjects: ["ENGLISH", "URDU", "ISLAMIYAT", "MUTALIA QURAN", "PHYSICS", "CHEMISTRY", "BOTANY"] },
  { pcs: "S1", key: "S1", until: "September", skip: ["mon"], subjects: ["ENGLISH", "URDU", "ISLAMIYAT", "MUTALIA QURAN", "PHYSICS", "CHEMISTRY", "BOTANY"] }
];
const RS_EXTRA_PROGRAM_FALLBACK = "Pre-Medical (JCW)";

// Combined rows: optional subjects of several sections sit together; compulsory stay in each section's row
const RS_COMPULSORY = ["ENGLISH", "URDU", "ISLAMIYAT", "MUTALIAQURAN", "PAKSTUDIES"];
const RS_COMBINED = [
  { pcs: "F.A-1", prefix: "F.A-1 ", extraKey: "FA1OPTIONALSUBJECTS" },
  { pcs: "F.A-2", prefix: "F.A-2 ", extraKey: "FA2OPTIONALSUBJECTS" }
];

const RS_COLUMNS = [
  { key: "att_September", group: "ATTENDANCE",    label: "September", type: "att", month: "September" },
  { key: "att_October",   group: "ATTENDANCE",    label: "October",   type: "att", month: "October" },
  { key: "att_November",  group: "ATTENDANCE",    label: "November",  type: "att", month: "November" },
  { key: "att_December",  group: "ATTENDANCE",    label: "December",  type: "att", month: "December" },
  { key: "att_January",   group: "ATTENDANCE",    label: "January",   type: "att", month: "January" },
  { key: "mon_September", group: "MONTHLY TESTS", label: "September", type: "mon", month: "September" },
  { key: "mon_October",   group: "MONTHLY TESTS", label: "October",   type: "mon", month: "October" },
  { key: "mon_November",  group: "MONTHLY TESTS", label: "November",  type: "mon", month: "November" },
  { key: "grand",         group: "",              label: "GRAND TEST", type: "grand" },
  { key: "det",           group: "",              label: "DETENTION",  type: "det" }
];

/* ---------- which columns are open this month (getMonth(): 0 = January) ---------- */
function rs_openKeys_(month) {
  const SEP = ["att_September", "mon_September"];
  const OCT = ["att_October",   "mon_October"];
  const NOV = ["att_November",  "mon_November"];
  const ALL = RS_COLUMNS.map(c => c.key);
  const SCHEDULE = {
    9:  [...SEP],
    10: [...SEP, ...OCT],
    11: [...SEP, ...OCT, ...NOV],
    0:  [...SEP, ...OCT, ...NOV, "att_December", "grand"],
    1:  ALL, 2: ALL, 3: ALL, 4: ALL, 5: ALL, 6: ALL
  };
  return SCHEDULE[month == null ? new Date().getMonth() : month] || [];
}

function rs_session_() {
  const t = new Date(), m = t.getMonth();
  const y = (m > 7 || (m === 7 && t.getDate() >= 15)) ? t.getFullYear() : t.getFullYear() - 1;
  return y + "-" + (y + 1);
}

const RS_SUBJECT_ALIASES = {
  ISLAMIAT: "ISLAMIYAT", ISLAMYAT: "ISLAMIYAT",
  MUTALAQURAN: "MUTALIAQURAN",
  ENG: "ENGLISH", PHY: "PHYSICS", CHEM: "CHEMISTRY", BOT: "BOTANY", ZOO: "ZOOLOGY"
};

const RS_SEC_RE  = /^(class\s*[-_\/ ]*\s*sec\w*|section|sec)$/i;
const RS_VAL_RE  = /^(obt|obtained|a\.?\s*l\b|attended|marks)/i;
const RS_SKIP_RE = /gt|grand|total|perc|%|t\.?\s*l\b/i;
const RS_ID_RE   = /^(s\.?\s*no|sr|class|section|sec|name|enroll|roll|father|reg|cnic)/i;
const RS_JUNK_RE = /^(MORNING|EVENING|TOTALSUBJECTS|CLASSSEC|CLASSSECTION|SNO)$/;

const rs_norm_ = s => String(s == null ? "" : s).toUpperCase().replace(/[^A-Z0-9]/g, "");
const rs_subj_ = s => { const n = rs_norm_(s); return RS_SUBJECT_ALIASES[n] || n; };

/* ---------- rows: every PCS with its subjects, then combined and extra sections ---------- */
function getRecordSummaryMeta() {
  const sh = SpreadsheetApp.openById(RS_STUDENT_SS).getSheetByName(RS_STRENGTH_TAB);
  if (!sh) throw new Error('Sheet "' + RS_STRENGTH_TAB + '" not found');
  const data = sh.getRange(1, 1, sh.getLastRow(), 16).getValues();

  const map = {};
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const pcs = String(r[8] || "").trim();
    if (!pcs) continue;
    const s = map[pcs] || (map[pcs] = { pcs, program: String(r[5] || "").trim(), subjects: [] });
    for (let c = 9; c < 16; c++) {
      const name = String(r[c] || "").trim().toUpperCase();
      if (name && !s.subjects.some(q => q.name === name)) s.subjects.push({ name, key: rs_subj_(name) });
    }
  }
  const sections = Object.values(map)
    .sort((a, b) => a.pcs.localeCompare(b.pcs, undefined, { numeric: true }))
    .map(s => { s.keys = [rs_norm_(s.pcs)]; return s; });

  // combined rows (F.A-1, F.A-2)
  RS_COMBINED.forEach(cb => {
    const members = sections.filter(s => s.pcs.indexOf(cb.prefix) === 0);
    if (!members.length) return;
    const optional = [];
    members.forEach(s => {
      s.subjects.forEach(sub => {
        if (!RS_COMPULSORY.includes(sub.key) && !optional.some(o => o.key === sub.key)) optional.push(sub);
      });
      s.subjects = s.subjects.filter(sub => RS_COMPULSORY.includes(sub.key));
    });
    const row = {
      pcs: cb.pcs, program: members[0].program, subjects: optional,
      keys: [cb.extraKey, rs_norm_(cb.pcs)].concat(members.map(s => s.keys[0]))
    };
    sections.splice(sections.indexOf(members[members.length - 1]) + 1, 0, row);
  });

  // extra rows (S1, T1) belong to the Pre-Medical program: reuse the exact program name used in the sheet
  const pm = sections.find(s => /pre[\s\-_]*med/i.test(s.program));
  const extraProgram = pm ? pm.program : RS_EXTRA_PROGRAM_FALLBACK;
  RS_EXTRA_SECTIONS.forEach(x => {
    if (sections.some(s => s.keys[0] === x.key)) return;
    const fixed = (x.subjects || []).map(n => ({ name: String(n).toUpperCase(), key: rs_subj_(n) }));
    sections.push({ pcs: x.pcs, program: extraProgram, subjects: fixed, keys: [x.key], dynamic: !fixed.length, until: x.until || "", skip: x.skip || [] });
  });

  // only the columns due now are sent; future months appear by themselves when their month comes
  const open = rs_openKeys_();
  const MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  const nm = (new Date().getMonth() + 1) % 12;
  const upcoming = rs_openKeys_(nm).filter(k => !open.includes(k))
    .map(k => { const c = RS_COLUMNS.find(x => x.key === k); return (c.group ? c.group + " " : "") + c.label; });
  return {
    title: "JCW RECORD SUMMARY SESSION " + rs_session_(),
    columns: RS_COLUMNS.filter(c => open.includes(c.key)).map(c => Object.assign({}, c, { open: true })),
    next: upcoming.length ? "Opens in " + MONTHS[nm] + ": " + upcoming.join(", ") : "",
    sections
  };
}

/* ---------- files for the OPEN columns only; each folder is listed once; every file carries its last-modified time ---------- */
function getRecordFileList(force) {
  const open = rs_openKeys_(), out = {}, listed = {};
  const cache = force ? {} : rs_readCache_();
  const mk = (id, name, mod) => { const o = { id, name, mod }; if (cache[id]) o.cached = cache[id]; return o; };
  const list = id => {
    if (listed[id]) return listed[id];
    const arr = [], it = DriveApp.getFolderById(id).getFilesByType(MimeType.GOOGLE_SHEETS);
    while (it.hasNext()) { const f = it.next(); arr.push(mk(f.getId(), f.getName(), f.getLastUpdated().getTime())); }
    return (listed[id] = arr);
  };
  RS_COLUMNS.forEach(col => {
    if (!open.includes(col.key)) return;
    if (col.type === "grand") {
      out[col.key] = [mk(RS_GT_SS, "Grand test", DriveApp.getFileById(RS_GT_SS).getLastUpdated().getTime())];
      return;
    }
    const files = list(col.type === "att" ? RS_ATT_FOLDER : col.type === "mon" ? RS_MONTHLY_FOLDER : RS_DET_FOLDER);
    const m = col.month ? col.month.slice(0, 3).toLowerCase() : "";
    out[col.key] = files.filter(f => !m || f.name.toLowerCase().includes(m));
  });
  return out;
}

/* ---------- v7.5: instant ticks from the attendance app (Attendance Index is updated on every save) ---------- */
function rs_live_() {
  const TOK = { SEPT: "September", SEP: "September", OCT: "October", NOV: "November", DEC: "December", JAN: "January", FEB: "February", MAR: "March" };
  const sh = SpreadsheetApp.openById(RS_STUDENT_SS).getSheetByName(typeof INDEX_SHEET !== "undefined" ? INDEX_SHEET : "Attendance Index");
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 12).getValues().forEach(x => {
    const mon = TOK[String(x[1]).toUpperCase().replace(/[^A-Z]/g, "")];
    if (!mon || !(Number(x[6]) > 0 || String(x[11]).replace(/[,\s]/g, ""))) return;
    const col = "att_" + mon, sec = rs_norm_(x[2]), sub = rs_subj_(x[3]), n = Number(x[5]) || 1;
    if (!sec || !sub) return;
    ((out[col] || (out[col] = {}))[sec] || (out[col][sec] = {}))[sub] = [n, n];
  });
  return out;
}

/* ---------- one file (page calls many of these in parallel) ---------- */
function getRecordFile(id, cached, mod) {
  const acc = {}, names = {}, errors = [];
  let sheets = null;
  try { sheets = rs_fetchGrid_(id); } catch (e) { sheets = null; }
  if (!sheets) {
    try {
      sheets = SpreadsheetApp.openById(id).getSheets()
        .map(sh => ({ name: sh.getName(), v: sh.getDataRange().getValues() }));
    } catch (e) { errors.push(e.message); return { acc, names, errors }; }
  }
  sheets.forEach(s => { try { rs_scanValues_(s.v, acc, names); } catch (e) { errors.push(s.name + ": " + e.message); } });
  if (errors.length) return { acc, names, errors };

  // stored ticks always stay ticked
  let before = 0;
  if (cached && cached.acc) {
    for (const sec in cached.acc) for (const sub in cached.acc[sec]) {
      before++;
      if (!acc[sec]) acc[sec] = {};
      if (!acc[sec][sub]) acc[sec][sub] = [0, 0];
      const e = acc[sec][sub], c = cached.acc[sec][sub];
      if (e[0] === 0) { e[0] = c[0]; if (!e[1]) e[1] = c[1]; }
    }
    Object.keys(cached.names || {}).forEach(k => { if (!names[k]) names[k] = cached.names[k]; });
  }

  // save only the ticked pairs, when something new was ticked or the file changed since the last save
  const ticked = {}; let n = 0;
  for (const sec in acc) for (const sub in acc[sec]) {
    if (acc[sec][sub][0] > 0) { (ticked[sec] || (ticked[sec] = {}))[sub] = acc[sec][sub]; n++; }
  }
  const complete = rs_isComplete_(acc);
  const modChanged = !!mod && (!cached || cached.mod !== mod);
  if (n !== before || (complete && !(cached && cached.complete)) || modChanged) {
    rs_saveCache_(id, { acc: ticked, names, complete, mod: mod || 0 });
  }
  return { acc, names, errors };
}

function rs_fetchGrid_(id) {
  const fields = "sheets(properties.title,data.rowData.values.formattedValue)";
  const url = "https://sheets.googleapis.com/v4/spreadsheets/" + id + "?includeGridData=true&fields=" + encodeURIComponent(fields);
  const r = UrlFetchApp.fetch(url, { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error("HTTP " + r.getResponseCode());
  const j = JSON.parse(r.getContentText());
  return (j.sheets || []).map(s => ({
    name: s.properties.title,
    v: ((s.data && s.data[0] && s.data[0].rowData) || [])
        .map(row => (row.values || []).map(c => c.formattedValue == null ? "" : c.formattedValue))
  }));
}

/* ---------- scanner: acc[SEC][SUBJECT] = [filledCells, totalRows]; names[SUBJECT] = original label ---------- */
function rs_scanValues_(v, acc, names) {
  if (v.length < 3) return;
  const lay = rs_layout_(v.slice(0, 15));
  if (!lay) return;
  Object.keys(lay.names).forEach(k => { if (!names[k]) names[k] = lay.names[k]; });
  const cols = Object.keys(lay.subjOf).map(Number);

  for (let r = lay.start; r < v.length; r++) {
    const row = v[r];
    if (!row) continue;
    const sec = rs_norm_(row[lay.sc]);
    if (!sec || RS_JUNK_RE.test(sec)) continue;
    const a = acc[sec] || (acc[sec] = {});
    for (let k = 0; k < cols.length; k++) {
      const c = cols[k], s = lay.subjOf[c];
      const e = a[s] || (a[s] = [0, 0]);
      e[1]++;
      const cell = String(row[c] == null ? "" : row[c]).trim();
      if (cell && cell !== "-") e[0]++;
    }
  }
}

function rs_layout_(v) {
  let hr = -1, sc = -1;
  for (let r = 0; r < v.length && hr < 0; r++)
    for (let c = 0; c < v[r].length; c++)
      if (RS_SEC_RE.test(String(v[r][c]).trim())) { hr = r; sc = c; break; }
  if (hr < 0) return null;

  const isVal = t => RS_VAL_RE.test(t) && !RS_SKIP_RE.test(t);
  let vr = hr, best = -1;
  for (let r = Math.max(0, hr - 1); r <= Math.min(v.length - 1, hr + 1); r++) {
    const n = v[r].filter(x => isVal(String(x).trim())).length;
    if (n > best) { best = n; vr = r; }
  }
  // a real mark/attendance sheet ALWAYS has a subject row above an Obtained / A.L row
  if (best < 1 || vr < 1) return null;
  const sr = vr - 1;
  const start = Math.max(hr, vr) + 1;
  const width = Math.max(v[vr].length, v[sr].length);

  const subjOf = {}, names = {};
  let cur = "";
  for (let c = 0; c < width; c++) {
    const t = String(v[sr][c] || "").trim();
    if (t) cur = RS_ID_RE.test(t) ? "" : t;
    const h = String(v[vr][c] || "").trim();
    if (cur && c !== sc && isVal(h)) { subjOf[c] = rs_subj_(cur); names[subjOf[c]] = cur.toUpperCase(); }
  }
  if (!Object.keys(subjOf).length) return null;
  return { sc, subjOf, start, names };
}

/* ---------- debug: run from the editor ---------- */
function debugRecordFile() {
  const list = getRecordFileList();
  Logger.log("Open columns: " + Object.keys(list).join(", "));
  const key = Object.keys(list)[0];
  if (!key || !list[key].length) return Logger.log("No files for " + key);
  const f = list[key][0];
  let t = new Date();
  try { rs_fetchGrid_(f.id); Logger.log("Sheets API fast path OK: " + (new Date() - t) + " ms"); }
  catch (e) { Logger.log("Sheets API fast path FAILED (" + e.message + ") -> fallback is used"); }
  t = new Date();
  const res = getRecordFile(f.id, f.cached || null, f.mod);
  Logger.log(f.name + " total " + (new Date() - t) + " ms | errors: " + JSON.stringify(res.errors));
  Logger.log("Sections: " + Object.keys(res.acc).join(", "));
}

/* ---------- debug ONE cell: run from the editor (change the 3 values below) ----------
   shows exactly what the scanner reads for that row + subject in that column's file(s) */
function debugCell() {
  const PCS = "F.A-1", SUBJECT = "LOGIC", COL = "att_October";   // <-- change these

  const sec = getRecordSummaryMeta().sections.find(s => s.pcs === PCS);
  if (!sec) return Logger.log("Row not found in the summary: " + PCS);
  const sub = rs_subj_(SUBJECT);
  Logger.log("Row " + PCS + " reads section labels " + sec.keys.join(", ") + " | subject key " + sub);

  const files = getRecordFileList(true)[COL] || [];
  if (!files.length) return Logger.log("No file found for column " + COL + " (is it open this month?)");
  files.forEach(f => {
    Logger.log("FILE: " + f.name + " | last changed " + new Date(f.mod));
    let sheets;
    try { sheets = rs_fetchGrid_(f.id); }
    catch (e) { sheets = SpreadsheetApp.openById(f.id).getSheets().map(sh => ({ name: sh.getName(), v: sh.getDataRange().getValues() })); }
    sheets.forEach(s => {
      if (s.v.length < 3) return;
      const lay = rs_layout_(s.v.slice(0, 15));
      if (!lay) return;
      const acc = {}; rs_scanValues_(s.v, acc, {});
      const hasSubj = Object.keys(lay.subjOf).some(c => lay.subjOf[c] === sub);
      const hits = sec.keys.filter(k => acc[k]);
      if (!hasSubj && !hits.length) return;
      Logger.log("  SHEET '" + s.name + "' | has " + SUBJECT + " column: " + (hasSubj ? "YES" : "NO"));
      if (!hits.length) Logger.log("    none of this row's section labels found. Labels in this sheet: " + Object.keys(acc).join(", "));
      hits.forEach(k => {
        const e = acc[k][sub];
        Logger.log("    section " + k + ": " + (e ? e[0] + " filled of " + e[1] + " rows" : "no " + SUBJECT + " entry") +
                   " | subjects here: " + Object.keys(acc[k]).join(", "));
      });
    });
  });
}

/* ---------- teachers: PCS + subject -> teacher names, from the Timetable sheet ---------- */
function getRecordTeachers() {
  const sh = SpreadsheetApp.openById(RS_STUDENT_SS).getSheetByName("Timetable");
  if (!sh) return { teachers: {}, practical: {} };
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { teachers: {}, practical: {} };
  const v = sh.getRange(1, 1, lastRow, Math.min(sh.getLastColumn(), 20)).getValues();

  const head = v[0].map(h => String(h).trim().toUpperCase());
  const col = (name, fallback) => head.indexOf(name) >= 0 ? head.indexOf(name) : fallback;
  const iSub = col("SUBJECT", 2), iTea = col("TEACHER", 3), iPcs = col("PCS", 8);
  const iTh = col("THEORY", -1), iPr = col("PRACTICAL", -1), iCourse = col("COURSE NAME", -1);
  const has = x => { const s = String(x == null ? "" : x).trim().toUpperCase(); return !!s && !/^(0|-|NO|FALSE|N\/A)$/.test(s); };

  // pcs -> subject -> teacher (upper case) -> { name, th, pr }
  const map = {};
  for (let i = 1; i < v.length; i++) {
    const pcs = rs_norm_(v[i][iPcs]), sub = rs_subj_(v[i][iSub]), t = String(v[i][iTea] || "").trim();
    if (!pcs || !sub || !t) continue;
    const course = iCourse >= 0 ? String(v[i][iCourse] || "").toUpperCase() : "";
    const prOnlyByName = /PRACTICAL|\bLAB\b|\(P\)/.test(course);
    const th = iTh >= 0 ? has(v[i][iTh]) && !prOnlyByName : !prOnlyByName;
    const pr = (iPr >= 0 && has(v[i][iPr])) || prOnlyByName;
    const S = (map[pcs] = map[pcs] || {}), T = (S[sub] = S[sub] || {}), k = t.toUpperCase();
    const e = T[k] || (T[k] = { name: t, th: false, pr: false });
    e.th = e.th || th || (!th && !pr);   // a row with neither column filled counts as theory
    e.pr = e.pr || pr;
  }

  // teachers: every teacher (shown in the table)
  // practical: teachers who ONLY take the practical of that subject in that section (not responsible for marking)
  const teachers = {}, practical = {};
  Object.keys(map).forEach(pcs => Object.keys(map[pcs]).forEach(sub => {
    const list = Object.values(map[pcs][sub]);
    (teachers[pcs] = teachers[pcs] || {})[sub] = list.map(e => e.name);
    const pOnly = list.filter(e => e.pr && !e.th).map(e => e.name);
    if (pOnly.length) (practical[pcs] = practical[pcs] || {})[sub] = pOnly;
  }));
  return { teachers, practical };
}

/* ---------- shared tick store (hidden sheet) ---------- */
const RS_CACHE_TAB = "RS_TICKS";

function rs_cacheSheet_() {
  const ss = SpreadsheetApp.openById(RS_STUDENT_SS);
  let sh = ss.getSheetByName(RS_CACHE_TAB);
  if (!sh) {
    sh = ss.insertSheet(RS_CACHE_TAB);
    sh.appendRow(["fileId", "json", "updated"]);
    sh.hideSheet();
  }
  return sh;
}

function rs_readCache_() {
  const sh = rs_cacheSheet_(), last = sh.getLastRow(), map = {};
  if (last < 2) return map;
  sh.getRange(2, 1, last - 1, 2).getValues().forEach(r => {
    try { if (r[0]) map[r[0]] = JSON.parse(r[1]); } catch (e) {}
  });
  return map;
}

// complete = at least one pair, and every section+subject that has rows has at least one filled entry
function rs_isComplete_(acc) {
  let pairs = 0;
  for (const sec in acc) for (const sub in acc[sec]) {
    const e = acc[sec][sub];
    if (e[1] > 0) { pairs++; if (e[0] === 0) return false; }
  }
  return pairs > 0;
}

function rs_saveCache_(id, payload) {
  const json = JSON.stringify(payload);
  if (json.length > 45000) return;                 // cell limit is 50,000 characters
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return; }
  try {
    const sh = rs_cacheSheet_(), last = sh.getLastRow();
    let row = -1;
    if (last >= 2) {
      const ids = sh.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) { row = i + 2; break; }
    }
    if (row < 0) row = last + 1;
    sh.getRange(row, 1, 1, 3).setValues([[id, json, new Date()]]);
  } finally { lock.releaseLock(); }
}

// run from the editor to wipe the stored ticks
function clearRecordTickStore() {
  const sh = rs_cacheSheet_();
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
}