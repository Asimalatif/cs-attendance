//const SAVE_KEY = 'JCW-D12CA0595D22478AA319';
// ==UserScript==
// @name         JCW Portal Sync
// @namespace    jcw-attendance
// @version      1.0
// @description  Reads Jinnah College students from the UOP portal (read-only) and sends them to the JCW attendance system: fills only the empty cells of Transform-ALL-COLLEGE-STRENGTH and feeds the health check.
// @match        https://portal.uop.edu.pk/Student/Students*
// @match        https://portal.uop.edu.pk/student/students*
// @grant        GM_xmlhttpRequest
// @grant        GM_getValue
// @grant        GM_setValue
// @connect      script.google.com
// @connect      script.googleusercontent.com
// @run-at       document-idle
// ==/UserScript==
/*
 * READ-ONLY on the portal: nothing is saved or changed there. It only reads the student list, each student's
 * edit page (DOB, contact, photo …), qualification, fee challans and courses, exactly like the export script.
 * Quick sync  = the student list only (status, batch, program, class, section, Reg#)   ~ 1-2 minutes
 * Full sync   = also the details of every ENROLLED student                             ~ 30-60 minutes (can be resumed)
 */
(function () {
  'use strict';
  // ===================== SETTINGS =====================
  const API = 'https://script.google.com/macros/s/AKfycbzl-DwgASIhf-pIGYi9D3M-CGOm7d7-UKg0W-OoWAiuRBHe4mjlRgMQXIHuLm6zi--i/exec';
  const STATUSES = ['Enrolled', 'Cancelled', 'Migrated', 'Disenrolled', 'Transfer'];
  const INSTITUTE_ID = '1212';                 // Jinnah College for Women
  const LIST_PAGE_SIZE = 500;
  const CONCURRENCY = 4;
  const SEND_SIZE = 250;                       // rows per message to the attendance server
  const PHOTO_BASE_URL = 'https://portal.uop.edu.pk/images/StudentPhoto/';
  const SAVE_KEY = 'JCW-D12CA0595D22478AA319';

  // ===================== SMALL HELPERS =====================
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const sp = v => String(v ?? '').replace(/\s+/g, ' ').trim();
  const clean = v => sp(String(v ?? '').replace(/[*:]/g, ''));
  const txt = el => el ? sp(el.textContent || '') : '';
  const val = el => el ? sp(el.value || '') : '';
  const num = v => { const n = parseFloat(String(v ?? '').replace(/,/g, '').replace(/[^\d.-]/g, '')); return Number.isFinite(n) ? n : 0; };
  const abs = u => { try { return new URL(u, location.href).href; } catch (e) { return u || ''; } };
  const parseHTML = html => new DOMParser().parseFromString(html, 'text/html');
  const first = (o, keys) => { for (const k of keys) if (o[k] !== undefined && String(o[k]).trim() !== '') return sp(o[k]); return ''; };

  // ===================== PROGRESS STORE (IndexedDB, survives a page reload) =====================
  const idb = (mode, fn) => new Promise((res, rej) => {
    const q = indexedDB.open('JCW_PORTAL_SYNC', 1);
    q.onupgradeneeded = e => e.target.result.createObjectStore('p', { keyPath: 'key' });
    q.onsuccess = () => { const t = q.result.transaction('p', mode).objectStore('p'), r = fn(t); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); };
    q.onerror = () => rej(q.error);
  });
  const loadSaved = () => idb('readonly', s => s.get(SAVE_KEY)).catch(() => null);
  const putSaved = o => idb('readwrite', s => s.put(o)).catch(() => null);
  const dropSaved = () => idb('readwrite', s => s.delete(SAVE_KEY)).catch(() => null);

  // ===================== PANEL =====================
  const box = document.createElement('div');
  box.id = 'jcwSync';
  box.innerHTML = `<style>
    #jcwSync{position:fixed;right:16px;bottom:16px;z-index:99999;width:300px;background:#fff;color:#101828;border:1px solid #d0d5dd;border-radius:14px;
      box-shadow:0 12px 32px -12px rgba(16,24,40,.35);font:14px/1.4 system-ui,Segoe UI,Arial,sans-serif;padding:12px}
    #jcwSync b.t{display:block;font-size:15px;margin-bottom:6px;color:#14213D}
    #jcwSync button{width:100%;height:38px;margin-top:6px;border:0;border-radius:10px;font-weight:700;cursor:pointer;background:#eef2ff;color:#14213D}
    #jcwSync button.p{background:#14213D;color:#fff}
    #jcwSync button:disabled{opacity:.5;cursor:default}
    #jcwSync .s{margin-top:8px;font-size:12.5px;color:#475467;white-space:pre-line;max-height:120px;overflow:auto}
    #jcwSync .bar{height:6px;background:#eef2f6;border-radius:6px;margin-top:6px;overflow:hidden}#jcwSync .bar i{display:block;height:100%;width:0;background:#2F55D4}
    #jcwSync .x{position:absolute;right:8px;top:6px;border:0;background:none;width:auto;height:auto;margin:0;font-size:18px;color:#98a2b3}
    #jcwSync .k{font-size:12px;color:#98a2b3;background:none;height:auto;margin-top:8px;text-decoration:underline}
    #jcwSync.min>*:not(b):not(.x){display:none}</style>
    <button class="x" id="jsMin" title="Hide">–</button><b class="t">JCW Portal Sync</b>
    <button class="p" id="jsQuick">Quick sync (list only, 1–2 min)</button>
    <button id="jsFull">Full sync (all details)</button>
    <button id="jsResume" style="display:none">Resume the full sync</button>
    <div class="bar"><i id="jsBar"></i></div><div class="s" id="jsMsg">Reads the portal only. Fills empty cells in the strength sheet; nothing already there is changed.</div>
    <button class="k" id="jsKey">Change sync key</button>`;
  document.body.appendChild(box);
  const $ = id => box.querySelector('#' + id);
  const say = (m, pct) => { $('jsMsg').textContent = m; if (pct != null) $('jsBar').style.width = Math.max(0, Math.min(100, pct)) + '%'; };
  const busy = on => ['jsQuick', 'jsFull', 'jsResume', 'jsKey'].forEach(i => { $(i).disabled = on; });
  $('jsMin').onclick = () => box.classList.toggle('min');
  $('jsKey').onclick = () => askKey(true);
  loadSaved().then(s => { if (s && s.mode === 'full' && s.items && Object.keys(s.done || {}).length < s.items.length) { $('jsResume').style.display = ''; $('jsResume').textContent = 'Resume the full sync (' + Object.keys(s.done).length + ' of ' + s.items.length + ' done)'; } });

  function askKey(force) {
    let k = GM_getValue('jcwKey', '');
    if (!k || force) {
      const n = prompt('Sync key from the JCW attendance app\n(College > UOP portal sync > Get sync key):', k || '');
      if (n == null) return k;
      k = n.trim(); GM_setValue('jcwKey', k);
    }
    return k;
  }

  // ===================== PORTAL REQUESTS (same as the export script) =====================
  const expired = h => /\/account\/login/i.test(h) || /session\s*(expired|timeout)/i.test(h) || /<title[^>]*>\s*(login|sign\s*in)/i.test(h) || (/<html/i.test(h) && /<input[^>]+type=["']password["']/i.test(h));
  async function getText(url, tries = 3) {
    let err;
    for (let a = 1; a <= tries; a++) {
      try { const r = await fetch(url, { credentials: 'include', cache: 'no-store' }); if (!r.ok) throw new Error('HTTP ' + r.status); const t = await r.text(); if (expired(t)) throw new Error('LOGIN'); return t; }
      catch (e) { err = e; if (e.message === 'LOGIN') break; await sleep(600 * a); }
    }
    throw err;
  }
  async function post(url, params, tries = 3) {
    let err;
    for (let a = 1; a <= tries; a++) {
      try {
        const r = await fetch(url, { method: 'POST', credentials: 'include', cache: 'no-store',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8', 'X-Requested-With': 'XMLHttpRequest' }, body: new URLSearchParams(params) });
        if (!r.ok) throw new Error('HTTP ' + r.status);
        const t = await r.text(); if (expired(t)) throw new Error('LOGIN'); return t;
      } catch (e) { err = e; if (e.message === 'LOGIN') break; await sleep(1000 * a); }
    }
    throw err;
  }
  function studentTable() {
    const d = document.querySelector('#tblDataStudents'); if (d) return d;
    let best = null, sc = -1;
    document.querySelectorAll('table').forEach(t => { const x = sp(t.textContent); let s = 0; if (/enrollment/i.test(x)) s += 3; if (/batch/i.test(x)) s += 2; if (/session/i.test(x)) s += 2; if (s > sc) { sc = s; best = t; } });
    return best;
  }
  const headersOf = t => { const r = t.querySelector('thead tr') || t.querySelector('tr'); return r ? [...r.children].map(c => clean(txt(c))) : []; };
  function rowInfo(tr, headers, status) {
    const cells = [...tr.children], off = Math.max(0, cells.length - headers.length), o = {};
    headers.forEach((h, i) => { if (h) o[h] = txt(cells[i + off]); });
    const links = [...tr.querySelectorAll('a[href]')].map(a => ({ text: clean(txt(a)), href: a.getAttribute('href') || '' }))
      .filter(l => l.href && !/^javascript:|^#/i.test(l.href)).map(l => ({ text: l.text, href: abs(l.href) }));
    const idFrom = s => (s.match(/[?&]id=(\d+)/i) || s.match(/(?:Students|Edit|Details|NewStudent)[^0-9]*(\d+)/i) || [])[1] || '';
    let id = ''; for (const l of links) { id = idFrom(l.href); if (id) break; }
    if (!id) id = idFrom(tr.outerHTML || '');
    return {
      id: sp(id), status: first(o, ['Status', 'StudentStatus']) || status, links,
      row: { 'Student Id': sp(id), 'Enrollment#': first(o, ['Enrollment#', 'Enrollment No', 'EnrollmentNo', 'Enrollment']),
        'Reg#': first(o, ['Reg#', 'RegistrationNo', 'Registration No', 'Registration#', 'Registration']), 'Batch': first(o, ['Batch']), 'Session': first(o, ['Session']),
        'Student Name': first(o, ['Student Name', 'StudentName', 'Name']), 'Father Name': first(o, ['Father Name', 'FatherName']), 'Category': first(o, ['Category']),
        'Program': first(o, ['Program']), 'Class': first(o, ['Class']), 'Section': first(o, ['Section']), 'Status': first(o, ['Status', 'StudentStatus']) || status, 'Load Status': 'list' }
    };
  }
  async function collectList() {
    const t = studentTable(); if (!t) throw new Error('Open the Students list page first (Student > Students).');
    const headers = headersOf(t), all = [], seen = new Set(), per = {};
    for (const status of STATUSES) {
      per[status] = 0;
      for (let page = 1; page <= 100; page++) {
        say('Reading the list: ' + status + ', page ' + page + ' … (' + all.length + ' students so far)');
        const html = await post('/Student/GetStudents', { PageNumber: String(page), SearchName: '', SortOrder: '1', PageSize: String(LIST_PAGE_SIZE), SessionId: '0', ProgramId: '0',
          ClassId: '0', CategoryId: '--All--', Gender: 'all', InstituteId: INSTITUTE_ID, SortingBy: 'Enrollment', Status: status, BatchId: '0' });
        const doc = parseHTML('<table><tbody>' + html + '</tbody></table>');
        const trs = [...doc.querySelectorAll('tbody tr')].filter(tr => tr.children.length >= headers.length - 1);
        let added = 0;
        trs.forEach(tr => { const it = rowInfo(tr, headers, status), k = it.id || it.row['Enrollment#'] || it.row['Student Name']; if (!k || seen.has(k)) return; seen.add(k); all.push(it); added++; });
        per[status] += added;
        const pages = parseInt(val(doc.querySelector('#txtTotalPages')), 10) || 0;
        if (!added || trs.length < LIST_PAGE_SIZE || (pages && page >= pages)) break;
      }
    }
    return { all, per };
  }

  // ===================== DETAILS OF ONE STUDENT =====================
  function editFields(html) {
    const doc = parseHTML(html), out = {};
    doc.querySelectorAll('input, select, textarea').forEach(el => {
      const key = el.getAttribute('name') || el.getAttribute('id'); if (!key || /^__|RequestVerificationToken|password/i.test(key)) return;
      const tag = el.tagName.toLowerCase(), type = (el.getAttribute('type') || '').toLowerCase();
      if (['button', 'submit', 'file', 'image', 'reset'].includes(type)) return;
      let v = '';
      if (tag === 'select') { const s = el.selectedOptions && el.selectedOptions[0]; v = s ? txt(s) : val(el); if (/^[-* ]*select/i.test(v)) v = ''; }
      else if (type === 'radio') { if (!el.checked) return; v = val(el); }
      else if (type === 'checkbox') v = el.checked ? 'Yes' : 'No';
      else v = val(el);
      if (v !== '' && !(key in out)) out[key] = v;
    });
    return out;
  }
  const dmy = raw => { const s = sp(raw), m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ].*)?$/); return m && Number(m[1]) >= 1000 ? `${m[3]}/${m[2]}/${m[1]}` : s; };
  const photo = p => { let v = sp(p); if (!v) return ''; if (/^https?:\/\//i.test(v)) return v; v = v.replace(/^\/+/, '').replace(/^images\/StudentPhoto\//i, '').replace(/^StudentPhoto\//i, ''); return PHOTO_BASE_URL + v; };
  const programId = h => (String(h || '').match(/function\s+GetPrograms\s*\([\s\S]*?var\s+selectedId\s*=\s*'(\d+)'/) || [])[1] || '';
  function controlValue(el) {
    if (!el) return '';
    if (el.tagName.toLowerCase() === 'select') { const s = el.selectedOptions && el.selectedOptions[0]; if (!s) return ''; const v = sp(s.value), t = sp(s.textContent); return (v === '' || v === '0' || /^[-* ]*(select|choose)/i.test(t)) ? '' : t; }
    return /input|textarea/i.test(el.tagName) ? val(el) : txt(el);
  }
  const QUAL = [['Board / University', '.board-class', /board|universit/i], ['Passing Year', '.passing-year-class', /passing|year/i], ['Roll.No.', '.roll-no-class', /roll/i],
    ['Obtain Marks/CGPA', '.obt-marks-class', /obtain|obt/i], ['Total Marks/CGPA', '.total-marks-class', /total/i], ['Division', '.division-class', /division/i], ['%age Marks', '.percentage-marks-class', /%|percent|age\s*marks/i]];
  function qualification(html) {
    const doc = parseHTML(html), t = doc.querySelector('#tblStuQualification') || doc.querySelector('table'); if (!t) return {};
    const hr = t.querySelector('thead tr'), ht = hr ? [...hr.children].map(th => clean(txt(th))) : [], hi = {};
    QUAL.forEach(([n, , re]) => { const i = ht.findIndex((h, k) => re.test(h) && !Object.values(hi).includes(k)); if (i >= 0) hi[n] = i; });
    let trs = [...t.querySelectorAll('tbody tr')]; if (!trs.length) trs = [...t.querySelectorAll('tr')].slice(hr ? 1 : 0);
    for (const tr of trs) {
      const rec = {};
      QUAL.forEach(([n, cls]) => { const el = tr.querySelector(cls); let v = ''; if (el) v = controlValue(el); else if (hi[n] !== undefined && tr.children[hi[n]]) { const c = tr.children[hi[n]], ctl = c.querySelector('select, input, textarea'); v = ctl ? controlValue(ctl) : clean(txt(c)); } rec[n] = sp(v); });
      if (Object.values(rec).some(Boolean)) return rec;            // first qualification row (matric)
    }
    return {};
  }
  function challans(text) {
    text = String(text || '').trim(); if (!text) return '';
    let recs = [];
    if (/^\s*[\[{]/.test(text)) { try { const j = JSON.parse(text); recs = Array.isArray(j) ? j : (j.data || j.rows || j.result || []); } catch (e) {} }
    if (!recs.length) {
      const H = ['SNo', 'Challan', 'Programe', 'Amount', 'Created Date', 'Due Date', 'Status', 'Paid Date', 'Remarks', 'Print'];
      parseHTML('<table><tbody>' + text + '</tbody></table>').querySelectorAll('tr').forEach(tr => {
        const v = [...tr.children].map(c => clean(txt(c))); if (v.length <= 1 || !v.some(Boolean)) return;
        const off = v.length - H.length; if (off < 0 || off > 3) return; const o = {}; H.forEach((h, i) => { o[h] = v[i + off] || ''; }); recs.push(o);
      });
    }
    return recs.map(r => { const ks = Object.keys(r), ak = ks.find(k => /amount|fee|total/i.test(k)), amt = num(ak ? r[ak] : ''); if (!amt) return '';
      const sk = ks.find(k => /status/i.test(k)), st = sk ? String(r[sk] || '').trim() : ''; return ((/^p(aid)?/i.test(st) || /1/.test(st) || /^true$/i.test(st)) ? 'P' : 'U') + '-' + amt.toFixed(2); }).filter(Boolean).join(', ');
  }
  function courses(html) {
    const doc = parseHTML(html), t = doc.querySelector('table.tblcurrentcourses') || doc.querySelector('table'), out = {}; if (!t) return out;
    let ti = 2; const hr = t.querySelector('thead tr'); if (hr) { const i = [...hr.children].findIndex(th => /course\s*title|subject/i.test(txt(th))); if (i >= 0) ti = i; }
    const yr = { '1st': [], '2nd': [] }; let lab = '';
    t.querySelectorAll('tr').forEach(tr => { const c = [...tr.children]; if (!c.length) return;
      if (c.length < 4 && /(1st|2nd|first|second)\s*year/i.test(txt(tr))) { lab = /2nd|second/i.test(txt(tr)) ? '2nd' : '1st'; return; }
      if (c.length >= 4 && /^\d+$/.test(clean(txt(c[0])))) { const s = clean(txt(c[ti])); if (s && lab) yr[lab].push(s); } });
    ['1st', '2nd'].forEach(y => { for (let i = 0; i < 7; i++) out['sub' + (i + 1) + ' ' + y + ' year'] = yr[y][i] || ''; });
    return out;
  }
  async function details(it) {
    const row = Object.assign({}, it.row), warn = [];
    const ed = it.links.find(l => /^edit$/i.test(l.text)) || it.links.find(l => /(edit|newstudent)/i.test(l.href) && /student/i.test(l.href)) || it.links.find(l => /edit/i.test(l.href));
    let eh = '';
    if (ed) {
      try {
        eh = await getText(ed.href); const f = editFields(eh);
        row['DOB'] = dmy(first(f, ['DOB', 'Dob', 'DateOfBirth', 'txtDOB']));
        row['StudentContact'] = first(f, ['StudentContact', 'studentContact', 'Contact', 'Phone', 'Mobile', 'txtStudentContact']);
        row['EmailAddress'] = first(f, ['EmailAddress', 'emailAddress', 'Email', 'txtEmailAddress']);
        row['PhotoName'] = photo(first(f, ['PhotoName', 'photoName', 'Photo', 'photo']));
        if (!row['Category']) row['Category'] = first(f, ['CategoryId', 'categoryId', 'Category']);
      } catch (e) { if (e.message === 'LOGIN') throw e; warn.push('edit page'); }
    } else warn.push('no edit link');
    if (it.id) {
      try { Object.assign(row, qualification(await post('/Student/StudentQualificationPartail', { Id: it.id, ProgramId: programId(eh) || '0' }))); } catch (e) { if (e.message === 'LOGIN') throw e; warn.push('qualification'); }
      try { row['Challan Details'] = challans(await post('/Student/GetStudentFeeChallans', { PageNumber: '1', SearchName: '', SortOrder: '1', PageSize: '500', StudentId: it.id })); } catch (e) { if (e.message === 'LOGIN') throw e; warn.push('challans'); }
    }
    const cu = it.links.find(l => /^courses$/i.test(l.text)) || it.links.find(l => /courses/i.test(l.text)) || it.links.find(l => /courses/i.test(l.href));
    if (cu) { try { Object.assign(row, courses(await getText(cu.href))); } catch (e) { if (e.message === 'LOGIN') throw e; warn.push('courses'); } }
    row['Load Status'] = warn.length ? 'OK (' + warn.join(', ') + ' failed)' : 'OK';
    return row;
  }

  // ===================== SEND TO THE ATTENDANCE SERVER =====================
  function gmPost(body) {
    return new Promise((res, rej) => GM_xmlhttpRequest({ method: 'POST', url: API, data: JSON.stringify(body), headers: { 'Content-Type': 'text/plain;charset=utf-8' }, timeout: 300000,
      onload: r => { try { res(JSON.parse(r.responseText)); } catch (e) { rej(new Error('The attendance server sent a web page instead of data.')); } },
      onerror: () => rej(new Error('No connection to the attendance server.')), ontimeout: () => rej(new Error('The attendance server took too long.')) }));
  }
  async function send(rows, mode) {
    const key = askKey(false); if (!key) throw new Error('No sync key.');
    const run = 'R' + Date.now(), parts = Math.max(1, Math.ceil(rows.length / SEND_SIZE));
    let last;
    for (let p = 1; p <= parts; p++) {
      say('Sending to the attendance system: part ' + p + ' of ' + parts + (p === parts ? ' (then filling the strength sheet, up to a minute)' : '') + ' …', p / parts * 100);
      let r, err;
      for (let a = 1; a <= 3; a++) { try { r = await gmPost({ action: 'portal_import', key, run, part: p, parts, mode, rows: rows.slice((p - 1) * SEND_SIZE, p * SEND_SIZE) }); break; } catch (e) { err = e; await sleep(2000 * a); } }
      if (!r) throw err;
      if (!r.ok) { if (/key/i.test(r.error || '')) GM_setValue('jcwKey', ''); throw new Error(r.error || 'Server error'); }
      last = r;
    }
    return last;
  }
  function report(r, per) {
    const f = (r.last && r.last.fill) || {}, L = [];
    L.push('Done: ' + r.last.rows + ' students sent (' + Object.keys(per).map(k => k + ' ' + per[k]).join(', ') + ').');
    if (f.skipped) L.push('Strength sheet NOT changed: ' + f.skipped);
    else if (f.error) L.push('Strength sheet: ' + f.error);
    else {
      const n = Object.values(f.cells || {}).reduce((a, b) => a + b, 0);
      L.push(n + ' empty cell(s) filled' + (f.rollsFilled ? ', ' + f.rollsFilled + ' roll number(s) added' : '') + '.');
      if (f.added && f.added.length) L.push('New columns: ' + f.added.join(', ') + '.');
      if (f.formulaCols && f.formulaCols.length) L.push('Not written (made by a formula): ' + f.formulaCols.join(', ') + '.');
    }
    L.push('Now run the health check in the app (College > System health check).');
    say(L.join('\n'), 100);
  }

  // ===================== RUN =====================
  async function run(mode, resume) {
    if (!askKey(false)) return;
    busy(true);
    try {
      let S = resume ? await loadSaved() : null;
      if (!S) {
        const { all, per } = await collectList();
        if (!all.length) throw new Error('No students found.');
        S = { key: SAVE_KEY, mode, per, items: all.map(it => ({ id: it.id, status: it.status, links: it.links, row: it.row })), done: {} };
        if (mode === 'full') await putSaved(S);
      }
      if (S.mode === 'full') {
        const todo = S.items.filter(it => /^enrolled$/i.test(it.status) && !S.done[it.id || it.row['Enrollment#']]);
        const total = S.items.filter(it => /^enrolled$/i.test(it.status)).length;
        let i = 0, n = total - todo.length, t0 = Date.now(), stop = null;
        const worker = async () => {
          while (!stop) {
            const it = todo[i++]; if (!it) return;
            try { const row = await details(it); S.done[it.id || row['Enrollment#']] = row; }
            catch (e) { if (e.message === 'LOGIN') { stop = 'The portal logged you out. Log in again, open Students, and press "Resume".'; return; } }
            n++;
            if (n % 10 === 0) { await putSaved(S); const left = (Date.now() - t0) / Math.max(1, n - (total - todo.length)) * (total - n); say('Details: ' + n + ' of ' + total + ' enrolled students · about ' + Math.ceil(left / 60000) + ' min left', n / total * 95); }
          }
        };
        await Promise.all(Array.from({ length: CONCURRENCY }, worker));
        await putSaved(S);
        if (stop) throw new Error(stop);
      }
      const rows = S.items.map(it => S.done[it.id || it.row['Enrollment#']] || it.row);
      const r = await send(rows, S.mode);
      if (S.mode === 'full') await dropSaved();
      $('jsResume').style.display = 'none';
      report(r, S.per);
    } catch (e) {
      say('Stopped: ' + e.message, null);
      loadSaved().then(s => { if (s && s.mode === 'full') { $('jsResume').style.display = ''; $('jsResume').textContent = 'Resume the full sync (' + Object.keys(s.done || {}).length + ' done)'; } });
    } finally { busy(false); }
  }
  $('jsQuick').onclick = () => run('quick', false);
  $('jsFull').onclick = () => { if (confirm('Full sync reads the details of every enrolled student (about 30–60 minutes). Keep this tab open. Start?')) run('full', false); };
  $('jsResume').onclick = () => run('full', true);
})();
