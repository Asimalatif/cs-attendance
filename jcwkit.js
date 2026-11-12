/* =====================================================================================================
   jcwkit.js — the JCW Daily Attendance look for the admin pages (admintools.html, …):
   - api(): calls the attendance backend with the app login (localStorage csa2_user / csa2_pass)
   - PDFs: same brown band + logo + footer + table style as the app (jsPDF + autoTable)
   - Excel: same brown header / tea-pink stripes as the app (xlsx-js-style), several sheets in one file
   - offer(): the app's "PDF ready" panel: Share · Download · Open / Print · Save to Drive
   ===================================================================================================== */
(function(){
'use strict';
const API = 'https://script.google.com/macros/s/AKfycbzl-DwgASIhf-pIGYi9D3M-CGOm7d7-UKg0W-OoWAiuRBHe4mjlRgMQXIHuLm6zi--i/exec';
const COLLEGE = 'JINNAH COLLEGE FOR WOMEN', COLLEGE_SUB = 'University of Peshawar', SYSTEM = 'JCW Daily Attendance';
const C = { INK:[43,25,18], NAVY:[107,52,35], ACC:[194,118,106], TEA:[231,183,174], BR2:[138,74,53], LINE:[226,192,184], MUTED:[124,97,89], SOFT:[251,237,234],
  RED:[217,45,32], REDBG:[253,236,234], GREEN:[18,128,92], GREENBG:[230,244,238], AMB:[183,121,31], AMBBG:[255,244,219] };
const LIBS = ['https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js', 'https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js'];
const XLSX_LIB = 'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.min.js';
const PDF_CACHE = 'jcw-pdf-tools-v1';
const $ = id => document.getElementById(id);
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]));
const ls = k => { try{ const v = localStorage.getItem('csa2_' + k); return v == null ? '' : JSON.parse(v); }catch(e){ return ''; } };
const user = () => ls('user'), teacher = () => ls('teacher') || ls('user');

/* ---------- server ---------- */
async function api(payload, ms){
  const u = ls('user'), p = ls('pass');
  if(!u || !p) throw new Error('Sign in to the attendance app first.');
  const ctl = new AbortController(), t = setTimeout(() => ctl.abort(), ms || 330000);
  try{
    const r = await fetch(API, { method:'POST', body:JSON.stringify(Object.assign({ username:u, password:p }, payload)), signal:ctl.signal, redirect:'follow' });
    let j; try{ j = JSON.parse(await r.text()); }catch(e){ throw new Error('The server sent a web page instead of data. Check the deployment.'); }
    if(j && j.error === 'UNKNOWN_ACTION') throw new Error('The server does not know this yet: paste the new .gs file, add its line to doPost and deploy a new version.');
    if(j && j.error === 'BAD_LOGIN') throw new Error('Your saved login was refused. Open the app, sign out and sign in again.');
    if(!j || j.ok === false) throw new Error((j && j.error) || 'Server error');
    return j;
  }catch(e){
    if(e.name === 'AbortError') throw new Error('The server took too long. Try again.');
    if(e instanceof TypeError) throw new Error(navigator.onLine ? 'The Apps Script server did not answer. If this keeps happening: in the Apps Script editor run atTest once and allow the permissions, make sure no old copy of these scripts is left in the project, then deploy a new version.' : 'No internet.');
    throw e;
  }finally{ clearTimeout(t); }
}

/* ---------- libraries (cached like in the app) ---------- */
function addScript(src){ return new Promise((res, rej) => { const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => { s.remove(); rej(new Error('not loaded')); }; document.head.appendChild(s); }); }
async function loadScript(src){
  try{ if(window.caches){ const r = await (await caches.open(PDF_CACHE)).match(src); if(r){ const u = URL.createObjectURL(new Blob([await r.text()], { type:'text/javascript' })); await addScript(u); return; } } }catch(e){}
  try{ await addScript(src); }catch(e){
    const local = /autotable/i.test(src) ? 'jspdf.plugin.autotable.min.js' : /jspdf/i.test(src) ? 'jspdf.umd.min.js' : '';
    if(local){ await addScript(local); return; }
    throw new Error('A tool could not be loaded (internet needed the first time).');
  }
  try{ if(window.caches) (await caches.open(PDF_CACHE)).add(src).catch(() => {}); }catch(e){}
}
const pdfReady = () => !!(window.jspdf && window.jspdf.jsPDF && window.jspdf.jsPDF.API.autoTable);
let libP = null, xP = null;
function libs(){ if(pdfReady()) return Promise.resolve(); if(!libP) libP = (window.jspdf ? Promise.resolve() : loadScript(LIBS[0])).then(() => pdfReady() ? null : loadScript(LIBS[1])).catch(e => { libP = null; throw e; }); return libP; }
function xlsxLib(){ if(window.XLSX && window.XLSX.utils) return Promise.resolve(); if(!xP) xP = loadScript(XLSX_LIB).catch(e => { xP = null; throw e; }); return xP; }
let LOGO = null;
(async () => { try{ const r = await fetch('jcw-logo.png'); if(r.ok){ const b = await r.blob(); LOGO = await new Promise(ok => { const fr = new FileReader(); fr.onload = () => ok(fr.result); fr.readAsDataURL(b); }); } }catch(e){} })();

/* ---------- PDF ---------- */
const TOP = 30, W = d => d.internal.pageSize.getWidth(), H = d => d.internal.pageSize.getHeight();
const stamp = () => new Date().toLocaleString('en-GB', { day:'numeric', month:'short', year:'numeric', hour:'2-digit', minute:'2-digit' });
const safe = s => String(s || '').replace(/[^A-Za-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
function newDoc(o){ const d = new window.jspdf.jsPDF({ unit:'mm', format:'a4', orientation:o || 'portrait' }); d.setLineHeightFactor(1.3); return d; }
function heading(doc, Y, t, sub){
  doc.setFont('helvetica', 'bold'); doc.setFontSize(16); doc.setTextColor(...C.INK); doc.text(t, 14, Y + 6);
  let extra = 0;
  if(sub){ doc.setFont('helvetica', 'normal'); doc.setFontSize(10); doc.setTextColor(...C.MUTED);
    const ls = doc.splitTextToSize(String(sub), W(doc) - 28).slice(0, 3); doc.text(ls, 14, Y + 11.5); extra = (ls.length - 1) * 4.6; }
  doc.setDrawColor(...C.ACC); doc.setLineWidth(0.8); doc.line(14, Y + 14.5 + extra, 36, Y + 14.5 + extra); doc.setLineWidth(0.2);
  return Y + 19 + extra;
}
function band(doc, Y, text, size){                              /* section title bar (like the old grey rows, now in JCW tea pink) */
  if(Y > H(doc) - 40){ doc.addPage(); Y = TOP; }
  doc.setFillColor(...C.SOFT); doc.setDrawColor(...C.LINE); doc.rect(14, Y, W(doc) - 28, (size || 12) * 0.62, 'FD');
  doc.setFont('helvetica', 'bold'); doc.setFontSize(size || 12); doc.setTextColor(...C.NAVY);
  doc.text(String(text), W(doc) / 2, Y + (size || 12) * 0.45, { align:'center' });
  return Y + (size || 12) * 0.62 + 2;
}
function table(doc, opts){
  doc.autoTable(Object.assign({
    margin:{ top:TOP, left:14, right:14, bottom:18 }, theme:'grid',
    styles:{ font:'helvetica', fontSize:8.8, cellPadding:1.8, lineColor:C.LINE, lineWidth:0.2, textColor:C.INK, valign:'middle' },
    headStyles:{ fillColor:C.NAVY, textColor:[255, 255, 255], fontStyle:'bold', halign:'center', fontSize:8.3 },
    alternateRowStyles:{ fillColor:C.SOFT }
  }, opts));
  return doc.lastAutoTable.finalY + 4;
}
function finalize(doc, opt){
  opt = opt || {};
  const n = doc.internal.getNumberOfPages(), w = W(doc), h = H(doc), st = stamp(), who = teacher();
  for(let i = 1; i <= n; i++){
    doc.setPage(i);
    if(!opt.noBand){
      doc.setFillColor(...C.NAVY); doc.rect(0, 0, w, 22, 'F'); doc.setFillColor(...C.TEA); doc.rect(0, 22, w, 1.4, 'F');
      let tx = 14;
      if(LOGO){ try{ doc.setFillColor(255, 255, 255); doc.circle(21, 11, 8, 'F'); doc.addImage(LOGO, 'PNG', 14.6, 4.6, 12.8, 12.4); tx = 32; }catch(e){} }
      doc.setFont('helvetica', 'bold'); doc.setFontSize(13); doc.setTextColor(255, 255, 255); doc.text(COLLEGE, tx, 10);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(8.5); doc.setTextColor(244, 214, 208); doc.text(COLLEGE_SUB, tx, 15.5);
      doc.setFont('helvetica', 'bold'); doc.setFontSize(9.5); doc.setTextColor(255, 255, 255); doc.text(opt.system || SYSTEM, w - 14, 10, { align:'right' });
      doc.setFont('helvetica', 'normal'); doc.setFontSize(8); doc.setTextColor(244, 214, 208); doc.text(st, w - 14, 15.5, { align:'right' });
    }
    if(!opt.noFooter){
      doc.setDrawColor(...C.LINE); doc.setLineWidth(0.2); doc.line(14, h - 13, w - 14, h - 13);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5); doc.setTextColor(...C.MUTED);
      doc.text('Computer-generated report' + (who ? ' · prepared by ' + who : '') + ' · ' + st, 14, h - 8.5);
      doc.text('Page ' + i + ' of ' + n, w - 14, h - 8.5, { align:'right' });
    }
  }
  return doc;
}

/* ---------- Excel (one or more sheets, JCW colours) ----------
   sheets: [{ name, title, rows:[[header…], [values…]…], cols:[widths] }] */
async function xlsx(sheets, fileName){
  await xlsxLib();
  const X = window.XLSX, wb = X.utils.book_new();
  const bd = { style:'thin', color:{ rgb:'E2C0B8' } }, border = { top:bd, bottom:bd, left:bd, right:bd };
  sheets.forEach(sh => {
    const rows = sh.rows || [], n = Math.max(1, ...rows.map(r => r.length));
    const num = v => { const t = String(v == null ? '' : v).trim(); return t !== '' && t.length < 15 && /^-?\d+(\.\d+)?$/.test(t) && !/^0\d/.test(t) ? Number(t) : (v == null ? '' : v); };
    const head = [[sh.title || sh.name], ['Jinnah College for Women · ' + stamp()]];
    const ws = X.utils.aoa_to_sheet(head.concat(rows.map(r => r.map(num))));
    const groups = new Set(sh.groupRows || []);                 /* rows (index in rows[]) shown as a section bar */
    for(let r = 0; r < rows.length + 2; r++) for(let c = 0; c < n; c++){
      const a = X.utils.encode_cell({ r, c }); if(!ws[a]) ws[a] = { t:'s', v:'' };
      const ri = r - 2;
      ws[a].s = r === 0 ? { font:{ bold:true, sz:14, color:{ rgb:'6B3423' } } }
        : r === 1 ? { font:{ italic:true, sz:9, color:{ rgb:'7C6159' } } }
        : groups.has(ri) ? { fill:{ patternType:'solid', fgColor:{ rgb:'E7B7AE' } }, font:{ bold:true, sz:12, color:{ rgb:'6B3423' } }, alignment:{ horizontal:'center' }, border }
        : (ri === 0 || (sh.headRows || []).includes(ri)) ? { fill:{ patternType:'solid', fgColor:{ rgb:'6B3423' } }, font:{ bold:true, color:{ rgb:'FFFFFF' } }, alignment:{ vertical:'center', horizontal:'center', wrapText:true }, border }
        : { fill:{ patternType:'solid', fgColor:{ rgb:ri % 2 ? 'FFFFFF' : 'FBEDEA' } }, font:{ color:{ rgb:'2B1912' } }, alignment:{ vertical:'top', wrapText:true }, border };
    }
    const merges = [{ s:{ r:0, c:0 }, e:{ r:0, c:n - 1 } }, { s:{ r:1, c:0 }, e:{ r:1, c:n - 1 } }];
    (sh.groupRows || []).forEach(ri => merges.push({ s:{ r:ri + 2, c:0 }, e:{ r:ri + 2, c:n - 1 } }));
    (sh.merges || []).forEach(m => merges.push({ s:{ r:m[0] + 2, c:m[1] }, e:{ r:m[2] + 2, c:m[3] } }));
    ws['!merges'] = merges;
    ws['!cols'] = sh.cols ? sh.cols.map(w => ({ wch:w })) : Array.from({ length:n }, (_, c) => ({ wch:Math.min(48, Math.max(8, ...rows.map(r => String(r[c] == null ? '' : r[c]).length + 2))) }));
    ws['!rows'] = [{ hpt:22 }, {}, { hpt:28 }];
    if(!sh.noFilter && !(sh.groupRows || []).length && !(sh.headRows || []).length) ws['!autofilter'] = { ref:X.utils.encode_range({ s:{ r:2, c:0 }, e:{ r:rows.length + 1, c:n - 1 } }) };
    X.utils.book_append_sheet(wb, ws, String(sh.name || 'Sheet').replace(/[\\/?*\[\]:]/g, ' ').slice(0, 31));
  });
  const out = X.write(wb, { bookType:'xlsx', type:'array' });
  return { blob:new Blob([out], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name:fileName };
}

/* ---------- toast + file panel ---------- */
function ensureUi(){
  if($('jkDlg')) return;
  const st = document.createElement('style');
  st.textContent = `#jkDlg{border:0;border-radius:18px;padding:18px;max-width:420px;width:92vw;box-shadow:0 20px 60px rgba(43,25,18,.35)}
#jkDlg::backdrop{background:rgba(43,25,18,.45)}#jkDlg h2{margin:0 0 4px;font:800 18px system-ui;color:#2B1912}#jkDlg .meta{color:#7C6159;font-size:13px;margin-bottom:12px;word-break:break-all}
.jkb{display:flex;flex-direction:column;gap:8px}.jkb button{display:flex;flex-direction:column;align-items:flex-start;text-align:left;border:1.5px solid #E2C0B8;background:#fff;border-radius:14px;padding:10px 14px;cursor:pointer;font:inherit}
.jkb button b{font-size:15px;color:#2B1912}.jkb button small{color:#7C6159;font-size:12px}.jkb button.p{background:linear-gradient(180deg,#8A4A35,#6B3423);border-color:#6B3423}.jkb button.p b,.jkb button.p small{color:#fff}
#jkClose{margin-top:10px;width:100%;height:40px;border-radius:12px;border:1.5px solid #E2C0B8;background:#fff;font-weight:800;cursor:pointer}
#jkToast{position:fixed;left:50%;bottom:22px;transform:translateX(-50%);background:#2B1912;color:#fff;padding:10px 16px;border-radius:12px;font:600 13px system-ui;z-index:9999;display:none;max-width:90vw;text-align:center;white-space:pre-line}`;
  document.head.appendChild(st);
  const d = document.createElement('dialog'); d.id = 'jkDlg'; d.innerHTML = '<div id="jkBody"></div>'; document.body.appendChild(d);
  const t = document.createElement('div'); t.id = 'jkToast'; document.body.appendChild(t);
}
function toast(msg, ms){ ensureUi(); const t = $('jkToast'); t.textContent = msg; t.style.display = 'block'; clearTimeout(t._h); t._h = setTimeout(() => { t.style.display = 'none'; }, ms || 3500); }
function download(blob, name){ const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 60000); }
const b64 = blob => new Promise((ok, no) => { const fr = new FileReader(); fr.onload = () => ok(String(fr.result).split(',')[1]); fr.onerror = () => no(fr.error); fr.readAsDataURL(blob); });
async function saveToDrive(blob, name, folder){
  const r = await api({ action:'at_save', name, mime:blob.type || 'application/pdf', folder:folder || 'JCW Admin Reports', data:await b64(blob) }, 330000);
  return r;
}
/** the app's panel: Share · Download · Open / Print · (Save to Drive) */
function offer(blob, name, title, opt){
  opt = opt || {};
  ensureUi();
  const d = $('jkDlg'), isPdf = /\.pdf$/i.test(name), file = new File([blob], name, { type:blob.type });
  const canShare = !!(navigator.canShare && navigator.canShare({ files:[file] })), kb = Math.max(1, Math.round(blob.size / 1024));
  const btn = (id, l, s, p) => `<button type="button" id="${id}" class="${p ? 'p' : ''}"><b>${l}</b><small>${s}</small></button>`;
  $('jkBody').innerHTML = `<h2>${isPdf ? 'PDF' : 'Excel file'} ready</h2><div class="meta">${esc(name)} · ${kb} KB</div><div class="jkb">
    ${canShare ? btn('jkShare', 'Share', 'WhatsApp, Gmail, Drive…', true) : ''}
    ${btn('jkDown', 'Download', 'Save the file on this device', !canShare)}
    ${btn('jkOpen', 'Open / Print', isPdf ? 'View it now, or print it' : 'Open it in Excel / Sheets after download', false)}
    ${opt.drive !== false ? btn('jkDrive', 'Save to Google Drive', 'Folder “' + esc(opt.folder || 'JCW Admin Reports') + '”', false) : ''}
  </div><button id="jkClose" type="button">Close</button>`;
  if(!d.open) d.showModal();
  const close = () => { if(d.open) d.close(); };
  $('jkClose').onclick = close;
  if($('jkShare')) $('jkShare').onclick = () => navigator.share({ files:[file], title:title || name }).then(close).catch(e => { if(!e || e.name !== 'AbortError') toast('Could not share. Use Download, then attach the file.'); });
  $('jkDown').onclick = () => { download(blob, name); close(); toast('Downloaded: ' + name); };
  $('jkOpen').onclick = () => { if(!isPdf){ download(blob, name); close(); return; } const w = window.open(URL.createObjectURL(blob), '_blank'); if(!w) download(blob, name); close(); };
  if($('jkDrive')) $('jkDrive').onclick = async () => {
    const b = $('jkDrive'); b.disabled = true; b.querySelector('small').textContent = 'Saving…';
    try{ const r = await saveToDrive(blob, name, opt.folder); close(); toast('✔ Saved to Drive: ' + r.name); if(r.url) window.open(r.url, '_blank'); }
    catch(e){ b.disabled = false; b.querySelector('small').textContent = e.message; }
  };
}

/* ---------- a Google Sheet read back by the server (values + merges + colours + bold + sizes + widths)
   drawn in the app style: its own colours are mapped to the JCW brown / tea-pink palette ---------- */
function hexRgb(h){ h = String(h || '').replace('#', ''); if(h.length === 3) h = h.split('').map(x => x + x).join(''); const n = parseInt(h, 16); return isNaN(n) ? null : [n >> 16 & 255, n >> 8 & 255, n & 255]; }
function hsl(rgb){ const [r, g, b] = rgb.map(x => x / 255), mx = Math.max(r, g, b), mn = Math.min(r, g, b), l = (mx + mn) / 2, d = mx - mn;
  if(!d) return [0, 0, l]; const s = l > .5 ? d / (2 - mx - mn) : d / (mx + mn);
  let h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; return [h * 60, s, l]; }
/** sheet colour -> {fill:[r,g,b]|null, white:boolean} in the JCW palette (status colours such as red / green / yellow are kept) */
function mapFill(hex){
  const rgb = hexRgb(hex); if(!rgb) return { fill:null, white:false };
  const [h, s, l] = hsl(rgb);
  if(l > 0.975) return { fill:null, white:false };
  const neutral = s < 0.3 || (h >= 185 && h <= 265);              /* greys and blues = "decoration" -> JCW colours */
  if(!neutral) return { fill:rgb, white:l < 0.45 };
  if(l < 0.68) return { fill:C.NAVY, white:true };
  if(l < 0.86) return { fill:C.TEA, white:false };
  return { fill:[246, 227, 223], white:false };
}
const toHex = rgb => rgb.map(x => ('0' + x.toString(16)).slice(-2)).join('').toUpperCase();
/** analyse a sheet: leading title rows, the table header row(s), the cells */
function sheetModel(sd){
  const V = sd.values || [], nR = V.length, n = sd.n || (V[0] || []).length;
  const cover = {}, start = {};
  (sd.merges || []).forEach(m => { start[m[0] + ',' + m[1]] = m; for(let r = m[0]; r < m[0] + m[2]; r++) for(let c = m[1]; c < m[1] + m[3]; c++) if(r !== m[0] || c !== m[1]) cover[r + ',' + c] = 1; });
  const filled = r => V[r].some(x => String(x).trim() !== '');
  const fullMerge = r => { const m = start[r + ',0']; return m && m[3] >= n && m[2] === 1; };
  const isHead = r => { const ne = V[r].map((x, c) => [String(x).trim(), c]).filter(x => x[0] !== '');
    return ne.length >= 3 && ne.every(x => sd.bold && sd.bold[r] && sd.bold[r][x[1]]) && !ne.some(x => cover[r + ',' + x[1]]); };
  const titles = []; let r0 = 0;
  /* leading full-width rows are the report titles, except a bar that sits right on top of a table header (a department bar) */
  while(r0 < nR && (!filled(r0) || (fullMerge(r0) && !(r0 + 1 < nR && isHead(r0 + 1))))){ if(filled(r0)) titles.push({ text:String(V[r0][0]).trim(), size:(sd.size && sd.size[r0] && sd.size[r0][0]) || 12 }); r0++; }
  const heads = []; for(let r = r0; r < nR; r++) if(isHead(r)) heads.push(r);
  let headRows = [];
  if(heads.length === 1 && heads[0] === r0){
    headRows = [r0];
    const nx = r0 + 1; if(nx < nR && filled(nx) && V[nx].every((x, c) => String(x).trim() === '' || (sd.bg && sd.bg[nx] && sd.bg[nx][c]))) headRows.push(nx);
  }
  return { V, nR, n, cover, start, filled, fullMerge, heads:new Set(heads), titles, r0, headRows };
}
function cellStyle(sd, r, c, k, kind){
  const st = { halign:(sd.align && sd.align[r] && sd.align[r][c]) === 'c' ? 'center' : (sd.align && sd.align[r] && sd.align[r][c]) === 'r' ? 'right' : 'left' };
  const pt = (sd.size && sd.size[r] && sd.size[r][c]) || 10;
  st.fontSize = Math.max(6.5, Math.min(15, pt * k));
  if(sd.bold && sd.bold[r] && sd.bold[r][c]) st.fontStyle = 'bold';
  const f = mapFill(sd.bg && sd.bg[r] && sd.bg[r][c]);
  if(f.fill){ st.fillColor = f.fill; if(f.white) st.textColor = [255, 255, 255]; }
  if(kind === 'head'){ st.fillColor = C.NAVY; st.textColor = [255, 255, 255]; st.fontStyle = 'bold'; st.halign = 'center'; }
  if(kind === 'band'){ st.fillColor = f.fill && !f.white ? f.fill : C.SOFT; st.textColor = C.NAVY; st.fontStyle = 'bold'; st.halign = 'center'; st.fontSize = Math.max(9, st.fontSize); }
  return st;
}
/** draws one read-back sheet into doc (new page unless first), in the app style */
function sheetPdf(doc, sd, opt){
  opt = opt || {};
  const M = sheetModel(sd), w = W(doc) - 28;
  const px = (sd.widths || []).slice(0, M.n), sumPx = px.reduce((a, b) => a + (b || 100), 0) || 1;
  const k = Math.min(1, w / (sumPx * 0.2646));
  let Y = TOP;
  const t0 = M.titles.map(t => ({ text:t.text.replace(/^JINNAH COLLEGE FOR WOMEN(,?\s*(UNIVERSITY OF PESHAWAR|UOP))?\s*[-\u2013:]?\s*/i, '') })).filter(t => t.text);
  Y = heading(doc, Y, opt.title || (t0[0] && t0[0].text) || sd.name, opt.sub || t0.slice(opt.title ? 0 : 1).map(t => t.text).join('  ·  ') || '');
  const rowCells = (r, kind) => {
    const out = [];
    for(let c = 0; c < M.n; c++){
      if(M.cover[r + ',' + c]) continue;
      const m = M.start[r + ',' + c], cell = { content:String(M.V[r][c] == null ? '' : M.V[r][c]), styles:cellStyle(sd, r, c, k, kind) };
      if(m){ if(m[3] > 1) cell.colSpan = m[3]; if(m[2] > 1 && kind !== 'head') cell.rowSpan = m[2]; }
      out.push(cell);
    }
    return out;
  };
  const head = M.headRows.map(r => rowCells(r, 'head'));
  const body = [];
  for(let r = M.r0; r < M.nR; r++){
    if(M.headRows.includes(r)) continue;
    if(!M.filled(r) && !Object.keys(M.start).some(key => { const m = M.start[key]; return r >= m[0] && r < m[0] + m[2] && m[2] > 1; })){
      body.push([{ content:'', colSpan:M.n, styles:{ minCellHeight:3, cellPadding:0.4, fillColor:[255, 255, 255], lineWidth:0 } }]); continue;
    }
    body.push(rowCells(r, M.fullMerge(r) ? 'band' : M.heads.has(r) ? 'head' : ''));
  }
  const cs = {}; px.forEach((p, i) => { cs[i] = { cellWidth:(p || 100) * 0.2646 * k }; });
  table(doc, { startY:Y, head:head.length ? head : undefined, body, columnStyles:cs, showHead:'everyPage',
    styles:{ font:'helvetica', fontSize:8, cellPadding:1.2, lineColor:C.LINE, lineWidth:0.2, textColor:C.INK, valign:'middle', overflow:'linebreak' },
    alternateRowStyles:{}, rowPageBreak:'avoid' });
  return doc;
}
/** the same sheet as an Excel sheet spec for gridXlsx (JCW colours, merges, widths kept) */
async function gridXlsx(list, fileName){
  await xlsxLib();
  const X = window.XLSX, wb = X.utils.book_new();
  const bd = { style:'thin', color:{ rgb:'E2C0B8' } }, border = { top:bd, bottom:bd, left:bd, right:bd };
  list.forEach(sd => {
    const M = sheetModel(sd), n = Math.max(1, M.n);
    const aoa = [[sd.title || sd.name], ['Jinnah College for Women · ' + stamp()]].concat(M.V.map(r => r.map(v => { const t = String(v == null ? '' : v).trim(); return t !== '' && t.length < 15 && /^-?\d+(\.\d+)?$/.test(t) && !/^0\d/.test(t) ? Number(t) : v; })));
    const ws = X.utils.aoa_to_sheet(aoa);
    for(let r = 0; r < aoa.length; r++) for(let c = 0; c < n; c++){
      const a = X.utils.encode_cell({ r, c }); if(!ws[a]) ws[a] = { t:'s', v:'' };
      if(r === 0){ ws[a].s = { font:{ bold:true, sz:14, color:{ rgb:'6B3423' } } }; continue; }
      if(r === 1){ ws[a].s = { font:{ italic:true, sz:9, color:{ rgb:'7C6159' } } }; continue; }
      const sr = r - 2, kind = M.headRows.includes(sr) || M.heads.has(sr) ? 'head' : (M.fullMerge(sr) ? 'band' : '');
      const f = mapFill(sd.bg && sd.bg[sr] && sd.bg[sr][c]), al = sd.align && sd.align[sr] && sd.align[sr][c];
      const s = { font:{ bold:!!(sd.bold && sd.bold[sr] && sd.bold[sr][c]), sz:Math.min(18, (sd.size && sd.size[sr] && sd.size[sr][c]) || 10), color:{ rgb:f.white ? 'FFFFFF' : '2B1912' } },
        alignment:{ vertical:'center', wrapText:true, horizontal:al === 'c' ? 'center' : al === 'r' ? 'right' : 'left' } };
      if(f.fill) s.fill = { patternType:'solid', fgColor:{ rgb:toHex(f.fill) } };
      if(kind === 'head'){ s.fill = { patternType:'solid', fgColor:{ rgb:'6B3423' } }; s.font.color = { rgb:'FFFFFF' }; s.font.bold = true; s.alignment.horizontal = 'center'; }
      if(kind === 'band'){ s.fill = { patternType:'solid', fgColor:{ rgb:f.fill && !f.white ? toHex(f.fill) : 'FBEDEA' } }; s.font.color = { rgb:'6B3423' }; s.font.bold = true; s.alignment.horizontal = 'center'; }
      if(M.filled(sr) || kind) s.border = border;
      ws[a].s = s;
    }
    ws['!merges'] = [{ s:{ r:0, c:0 }, e:{ r:0, c:n - 1 } }, { s:{ r:1, c:0 }, e:{ r:1, c:n - 1 } }]
      .concat((sd.merges || []).map(m => ({ s:{ r:m[0] + 2, c:m[1] }, e:{ r:m[0] + m[2] + 1, c:m[1] + m[3] - 1 } })));
    ws['!cols'] = (sd.widths || []).map(p => ({ wpx:p || 100 }));
    X.utils.book_append_sheet(wb, ws, String(sd.name || 'Sheet').replace(/[\\/?*\[\]:]/g, ' ').slice(0, 31));
  });
  const out = X.write(wb, { bookType:'xlsx', type:'array' });
  return { blob:new Blob([out], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name:fileName };
}
/** the read-back sheet as an HTML preview table (app colours) */
function sheetHtml(sd, maxRows){
  const M = sheetModel(sd), lim = Math.min(M.nR, maxRows || 400);
  let h = '<table class="sht"><tbody>';
  for(let r = 0; r < lim; r++){
    if(!M.filled(r) && !(sd.merges || []).some(m => r >= m[0] && r < m[0] + m[2] && m[2] > 1)){ h += '<tr class="gap"><td colspan="' + M.n + '"></td></tr>'; continue; }
    const kind = M.headRows.includes(r) || M.heads.has(r) ? 'head' : M.fullMerge(r) ? 'band' : '';
    h += '<tr class="' + kind + '">';
    for(let c = 0; c < M.n; c++){
      if(M.cover[r + ',' + c]) continue;
      const m = M.start[r + ',' + c], f = mapFill(sd.bg && sd.bg[r] && sd.bg[r][c]);
      const st = [];
      if(!kind && f.fill) st.push('background:#' + toHex(f.fill) + (f.white ? ';color:#fff' : ''));
      if(sd.bold && sd.bold[r] && sd.bold[r][c]) st.push('font-weight:700');
      const al = sd.align && sd.align[r] && sd.align[r][c]; if(al) st.push('text-align:' + (al === 'c' ? 'center' : 'right'));
      h += '<td' + (m && m[3] > 1 ? ' colspan="' + m[3] + '"' : '') + (m && m[2] > 1 ? ' rowspan="' + Math.min(m[2], lim - r) + '"' : '') + (st.length ? ' style="' + st.join(';') + '"' : '') + '>' + esc(M.V[r][c]).replace(/\n/g, '<br>') + '</td>';
    }
    h += '</tr>';
  }
  h += '</tbody></table>';
  if(M.nR > lim) h += '<p class="more">… ' + (M.nR - lim) + ' more rows in the PDF / Excel.</p>';
  return h;
}

window.JCW = { api, libs, xlsx, xlsxLib, sheetPdf, gridXlsx, sheetHtml, sheetModel, mapFill, newDoc, heading, band, table, finalize, offer, saveToDrive, toast, download, esc, safe, stamp, C, TOP, W, H, user, teacher, logo:() => LOGO };
})();
