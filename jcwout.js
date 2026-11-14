/* =====================================================================================================
   jcwout.js — one "file ready" panel for the whole JCW attendance software (v7.6)
   Every PDF or Excel the app makes gets the same choices:
     PDF   : Open · Print · Download · Share
     Excel : Open (table, filter, print) · Download · Share
   A PDF made of tables also offers the Excel version and an Excel / CSV file also offers the PDF version.
   - JCWOUT.offer({ blob, name, title, tables, extra }) shows the panel
   - JCWOUT.list({ title, sub, head, rows, file }) : a list (students etc.) -> the same panel
   - jsPDF doc.save() and XLSX.writeFile() are routed to the panel by themselves
   Load it after inksaver.js:  <script src="jcwout.js"></script>
   ===================================================================================================== */
(function () {
  'use strict';
  if (window.JCWOUT) return;
  var TBL = new WeakMap();                       // PDF blob -> [{ head, body }]
  var XLSX_LIB = 'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.min.js';
  var base = (function () { var s = document.currentScript && document.currentScript.src; return s ? s.replace(/[^/]*$/, '') : ''; })();
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); };
  var isPhone = /Android|iPhone|iPad|Mobile/i.test(navigator.userAgent);

  function addScript(src) { return new Promise(function (ok, no) { var s = document.createElement('script'); s.src = src; s.onload = ok; s.onerror = function () { s.remove(); no(new Error('Could not load ' + src)); }; document.head.appendChild(s); }); }
  function needPdf() {
    var j = window.jspdf;
    if (j && j.jsPDF && j.jsPDF.API.autoTable) return Promise.resolve();
    return (j && j.jsPDF ? Promise.resolve() : addScript(base + 'jspdf.umd.min.js'))
      .then(function () { return window.jspdf.jsPDF.API.autoTable ? null : addScript(base + 'jspdf.plugin.autotable.min.js'); });
  }
  function needXlsx() { return window.XLSX && window.XLSX.utils ? Promise.resolve() : addScript(XLSX_LIB); }

  /* ---------- table text out of autoTable options ---------- */
  var cellText = function (c) { return c == null ? '' : typeof c === 'object' && !(c instanceof Date) ? String(c.content != null ? c.content : c.text != null ? c.text : '') : String(c); };
  function record(doc, args) {
    try {
      var o = args[0] && typeof args[0] === 'object' && !Array.isArray(args[0]) ? args[0] : (args[2] || { columns: args[0], body: args[1] });
      var head = o.head ? (Array.isArray(o.head[0]) ? o.head[o.head.length - 1] : o.head) : o.columns ? o.columns.map(function (c) { return c && typeof c === 'object' ? (c.header || c.title || c.dataKey || '') : c; }) : [];
      var body = (o.body || []).map(function (r) { return Array.isArray(r) ? r.map(cellText) : head.map(function (h, i) { var v = r[(o.columns && o.columns[i] && o.columns[i].dataKey) || i]; return cellText(v); }); });
      if (!body.length && o.html) {
        var t = typeof o.html === 'string' ? document.querySelector(o.html) : o.html;
        if (t) { var R = [].slice.call(t.rows).map(function (tr) { return [].slice.call(tr.cells).map(function (c) { return c.innerText.trim(); }); }); head = R.shift() || []; body = R; }
      }
      if (body.length) (doc.__jcwTbl = doc.__jcwTbl || []).push({ head: head.map(cellText), body: body });
    } catch (e) {}
  }
  function patchPdf(J) {
    if (!J || J.__jcwOut || !J.API) return; J.__jcwOut = 1;
    (J.API.events = J.API.events || []).push(['initialized', function () {     // each new document: wrap its own methods
      var d = this;
      var at = d.autoTable;
      if (typeof at === 'function') d.autoTable = function () { record(d, arguments); return at.apply(d, arguments); };
      var out = d.output;
      if (typeof out === 'function') d.output = function (type) {
        var r = out.apply(d, arguments);
        if (type === 'blob' && r instanceof Blob && d.__jcwTbl) TBL.set(r, d.__jcwTbl);
        return r;
      };
      var sv = d.save;
      if (typeof sv === 'function') d.save = function (name, o) {
        if (d.__jcwDirect || (o && o.returnPromise)) return sv.apply(d, arguments);
        offer({ blob: d.output('blob'), name: name || 'document.pdf', title: String(name || 'PDF').replace(/\.pdf$/i, '') });
        return d;
      };
    }]);
  }
  function patchX(X) {
    if (!X || !X.writeFile || X.__jcwOut) return; X.__jcwOut = 1;
    var wf = X.writeFile;
    X.writeFile = function (wb, name, o) {
      if (o && o.direct) return wf.apply(this, arguments);
      try {
        var arr = X.write(wb, { bookType: 'xlsx', type: 'array' });
        var blob = new Blob([arr], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        offer({ blob: blob, name: /\.xlsx$/i.test(name) ? name : String(name || 'file').replace(/\.[^.]+$/, '') + '.xlsx', title: String(name || '').replace(/\.[^.]+$/, ''), wb: wb });
      } catch (e) { return wf.apply(this, arguments); }
    };
  }
  /* hook window.jspdf / window.XLSX whenever they are set (inksaver.js may already own the jspdf property) */
  function hook(prop, fn) {
    var d = Object.getOwnPropertyDescriptor(window, prop), cur = window[prop];
    try {
      Object.defineProperty(window, prop, {
        configurable: true,
        get: function () { var v = d && d.get ? d.get.call(window) : cur; if (v) fn(v); return v; },
        set: function (v) { if (d && d.set) d.set.call(window, v); else cur = v; if (v) fn(v); }
      });
    } catch (e) { if (cur) fn(cur); }
  }
  hook('jspdf', function (j) { if (j && j.jsPDF) patchPdf(j.jsPDF); });
  hook('XLSX', patchX);

  /* ---------- rows <-> files ---------- */
  function rowsOfWb(wb) {
    var X = window.XLSX; return wb.SheetNames.map(function (n) { return { name: n, rows: X.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: false, defval: '' }) }; });
  }
  function parseCsv(t) {
    var R = [], row = [], f = '', q = false; t = String(t).replace(/^﻿/, '');
    for (var i = 0; i < t.length; i++) {
      var c = t[i];
      if (q) { if (c === '"') { if (t[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c; }
      else if (c === '"') q = true; else if (c === ',') { row.push(f); f = ''; }
      else if (c === '\n' || c === '\r') { if (c === '\r' && t[i + 1] === '\n') i++; row.push(f); R.push(row); row = []; f = ''; }
      else f += c;
    }
    if (f !== '' || row.length) { row.push(f); R.push(row); }
    return R;
  }
  function sheetsOf(o) {                                          // -> [{ name, rows:[[...]] }] (first row = header)
    if (o._sheets) return Promise.resolve(o._sheets);
    var p;
    if (o.tables && o.tables.length) p = Promise.resolve(o.tables.map(function (t, i) { return { name: 'Table ' + (i + 1), rows: [t.head].concat(t.body) }; }));
    else if (o.rows) p = Promise.resolve([{ name: 'List', rows: [o.head || []].concat(o.rows) }]);
    else if (o.wb) p = needXlsx().then(function () { return rowsOfWb(o.wb); });
    else if (/\.csv$/i.test(o.name)) p = o.blob.text().then(function (t) { return [{ name: 'List', rows: parseCsv(t).filter(function (r) { return r.some(function (v) { return String(v).trim(); }); }) }]; });
    else if (/\.xlsx?$/i.test(o.name)) p = Promise.all([needXlsx(), o.blob.arrayBuffer()]).then(function (a) { return rowsOfWb(window.XLSX.read(a[1], { type: 'array' })); });
    else p = Promise.resolve([]);
    return p.then(function (s) { o._sheets = s.filter(function (x) { return x.rows.length; }); return o._sheets; });
  }
  function toXlsx(o) {
    return Promise.all([needXlsx(), sheetsOf(o)]).then(function (a) {
      var X = window.XLSX, wb = X.utils.book_new(), used = {};
      a[1].forEach(function (s, k) {
        var ws = X.utils.aoa_to_sheet([[o.title || o.name], ['Jinnah College for Women, Peshawar']].concat([[]]).concat(s.rows));
        var w = Math.max.apply(null, s.rows.map(function (r) { return r.length; }));
        for (var c = 0; c < w; c++) { var ad = X.utils.encode_cell({ r: 3, c: c }); if (ws[ad]) ws[ad].s = { font: { bold: true }, fill: { fgColor: { rgb: 'EEEEEE' } } }; }
        if (ws.A1) ws.A1.s = { font: { bold: true, sz: 14 } };
        ws['!cols'] = Array.from({ length: w }, function (_, c) { return { wch: Math.min(45, Math.max.apply(null, s.rows.slice(0, 300).map(function (r) { return String(r[c] == null ? '' : r[c]).length; }).concat([6])) + 2) }; });
        var nm = String(a[1].length > 1 ? s.name : (o.title || s.name)).replace(/[\[\]*?\/\\:]/g, ' ').slice(0, 28) || 'Sheet';
        while (used[nm]) nm = nm.slice(0, 26) + ' ' + (++k); used[nm] = 1;
        X.utils.book_append_sheet(wb, ws, nm);
      });
      var arr = X.write(wb, { bookType: 'xlsx', type: 'array' });
      return { blob: new Blob([arr], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name: base_(o.name) + '.xlsx' };
    });
  }
  function base_(n) { return String(n || 'JCW').replace(/\.[^.]+$/, ''); }
  function toPdf(o) {
    return Promise.all([needPdf(), sheetsOf(o)]).then(function (a) {
      var S = a[1], wide = Math.max.apply(null, S.map(function (s) { return (s.rows[0] || []).length; }).concat([1])) > 6;
      var doc = new window.jspdf.jsPDF({ unit: 'mm', format: 'a4', orientation: wide ? 'landscape' : 'portrait' }), W = doc.internal.pageSize.getWidth();
      doc.__jcwDirect = 1;
      doc.setFont('helvetica', 'bold'); doc.setFontSize(13); doc.text('JINNAH COLLEGE FOR WOMEN, UNIVERSITY OF PESHAWAR', W / 2, 12, { align: 'center' });
      doc.setFontSize(11); doc.text(String(o.title || base_(o.name)).slice(0, 110), W / 2, 18, { align: 'center' });
      if (o.sub) { doc.setFont('helvetica', 'normal'); doc.setFontSize(9); doc.text(String(o.sub).replace(/<[^>]+>/g, '').slice(0, 160), W / 2, 23, { align: 'center' }); }
      var y = o.sub ? 27 : 23;
      S.forEach(function (s, i) {
        if (i) { y = doc.lastAutoTable.finalY + 8; if (y > doc.internal.pageSize.getHeight() - 30) { doc.addPage(); y = 14; } if (S.length > 1) { doc.setFont('helvetica', 'bold'); doc.setFontSize(10); doc.text(s.name, 8, y); y += 3; } }
        doc.autoTable({ startY: y, head: [s.rows[0] || []], body: s.rows.slice(1), theme: 'grid', margin: { left: 8, right: 8, bottom: 12 },
          styles: { fontSize: s.rows.length > 1500 ? 6.5 : 8, cellPadding: 1, textColor: 0, lineColor: 120, lineWidth: 0.1, overflow: 'linebreak' },
          headStyles: { fillColor: 235, textColor: 0, fontStyle: 'bold' } });
      });
      var n = doc.internal.getNumberOfPages();
      for (var p = 1; p <= n; p++) { doc.setPage(p); doc.setFont('helvetica', 'normal'); doc.setFontSize(7); doc.text('Page ' + p + ' of ' + n + ' · ' + new Date().toLocaleString('en-GB'), W / 2, doc.internal.pageSize.getHeight() - 5, { align: 'center' }); }
      return { blob: doc.output('blob'), name: base_(o.name) + '.pdf' };
    });
  }

  /* ---------- actions ---------- */
  function download(blob, name) { var a = document.createElement('a'), u = URL.createObjectURL(blob); a.href = u; a.download = name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(function () { URL.revokeObjectURL(u); }, 120000); }
  function share(blob, name, title) {
    var f = new File([blob], name, { type: blob.type });
    if (navigator.canShare && navigator.canShare({ files: [f] })) return navigator.share({ files: [f], title: title || name }).catch(function (e) { if (e && e.name !== 'AbortError') { download(blob, name); toast('Sharing did not work here: the file was downloaded instead.'); } });
    download(blob, name); toast('This browser cannot share files: it was downloaded. Send it from your Downloads.');
    return Promise.resolve();
  }
  function openPdf(blob, name, print) {
    var u = URL.createObjectURL(blob);
    if (print && !isPhone) {
      var fr = document.createElement('iframe'); fr.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0';
      fr.src = u; document.body.appendChild(fr);
      fr.onload = function () { try { fr.contentWindow.focus(); fr.contentWindow.print(); } catch (e) { window.open(u, '_blank'); } };
      setTimeout(function () { fr.remove(); URL.revokeObjectURL(u); }, 600000);
      return;
    }
    var w = window.open(u, '_blank');
    if (!w) { download(blob, name); toast('The PDF was downloaded (this browser does not open it here).'); }
    else if (print && isPhone) toast('Use the ⋮ menu → Print (or Share → Print) in the PDF that opened.');
    setTimeout(function () { URL.revokeObjectURL(u); }, 600000);
  }
  function openTable(o, print) {
    return sheetsOf(o).then(function (S) {
      var w = window.open('', '_blank');
      if (!w) { toast('Allow pop-ups for this site to open the table.'); return; }
      var t = esc(o.title || base_(o.name));
      w.document.write('<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>' + t + '</title><style>' +
        'body{font:14px system-ui,sans-serif;margin:14px;color:#111}h1{font-size:17px;margin:0 0 2px}.s{color:#555;font-size:12px;margin-bottom:8px}' +
        'table{border-collapse:collapse;width:100%;font-size:12.5px;margin-bottom:16px}th,td{border:1px solid #999;padding:4px 6px;text-align:left;vertical-align:top}th{background:#eee;position:sticky;top:0}' +
        '.bar{display:flex;gap:8px;flex-wrap:wrap;margin:8px 0}.bar button,.bar input{height:36px;border-radius:9px;border:1.5px solid #999;background:#fff;padding:0 12px;font:inherit}.bar button{font-weight:700;cursor:pointer}' +
        '@media print{.bar{display:none}th{position:static}}</style></head><body><h1>' + t + '</h1><div class="s">Jinnah College for Women, Peshawar · ' + new Date().toLocaleString('en-GB') + '</div>' +
        '<div class="bar"><input id="q" placeholder="Filter rows…" oninput="var v=this.value.toLowerCase();[].forEach.call(document.querySelectorAll(\'tbody tr\'),function(r){r.style.display=r.textContent.toLowerCase().indexOf(v)>-1?\'\':\'none\'})"><button onclick="print()">Print</button><button onclick="close()">Close</button></div>' +
        S.map(function (s) {
          return (S.length > 1 ? '<h1 style="font-size:14px">' + esc(s.name) + '</h1>' : '') + '<table><thead><tr>' + (s.rows[0] || []).map(function (h) { return '<th>' + esc(h) + '</th>'; }).join('') + '</tr></thead><tbody>' +
            s.rows.slice(1, 20001).map(function (r) { return '<tr>' + r.map(function (v) { return '<td>' + esc(v) + '</td>'; }).join('') + '</tr>'; }).join('') + '</tbody></table>' +
            (s.rows.length > 20001 ? '<div class="s">First 20,000 rows shown: download the Excel file for all ' + (s.rows.length - 1).toLocaleString() + '.</div>' : '');
        }).join('') + (print ? '<script>setTimeout(print,400)<\/script>' : '') + '</body></html>');
      w.document.close();
    });
  }

  /* ---------- the panel ---------- */
  function ui() {
    if (document.getElementById('joDlg')) return;
    var st = document.createElement('style');
    st.textContent = '#joDlg{border:0;border-radius:18px;padding:16px;max-width:460px;width:94vw;box-shadow:0 20px 60px rgba(43,25,18,.35);font:14px system-ui,sans-serif;color:#2B1912}' +
      '#joDlg::backdrop{background:rgba(43,25,18,.45)}#joDlg h2{margin:0 0 2px;font-size:18px}#joDlg .jm{color:#7C6159;font-size:12.5px;margin-bottom:10px;word-break:break-all}' +
      '#joDlg .jg{border:1.5px solid #E2C0B8;border-radius:14px;padding:10px;margin-bottom:10px}#joDlg .jg b.t{display:block;font-size:13px;margin-bottom:6px}' +
      '#joDlg .jr{display:grid;grid-template-columns:repeat(4,1fr);gap:6px}#joDlg .jr button{height:42px;border-radius:11px;border:1.5px solid #E2C0B8;background:#fff;font:700 13px system-ui;cursor:pointer;color:#2B1912}' +
      '#joDlg .jr button.p{background:#6B3423;border-color:#6B3423;color:#fff}#joDlg .jr button:disabled{opacity:.5}#joDlg .jx{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:8px}#joDlg .jx button{flex:1;height:40px;border-radius:11px;border:1.5px solid #E2C0B8;background:#fff;font:700 13px system-ui;cursor:pointer}' +
      '#joClose{width:100%;height:42px;border-radius:12px;border:1.5px solid #E2C0B8;background:#fff;font:800 14px system-ui;cursor:pointer}#joMsg{font-size:12.5px;color:#6B3423;min-height:16px;margin-bottom:6px}' +
      '#joToast{position:fixed;left:50%;bottom:22px;transform:translateX(-50%);background:#2B1912;color:#fff;padding:10px 16px;border-radius:12px;font:600 13px system-ui;z-index:2147483647;display:none;max-width:90vw;text-align:center;white-space:pre-line}' +
      '@media (max-width:380px){#joDlg .jr{grid-template-columns:repeat(2,1fr)}}';
    document.head.appendChild(st);
    var d = document.createElement('dialog'); d.id = 'joDlg'; d.innerHTML = '<div id="joBody"></div>'; document.body.appendChild(d);
    var t = document.createElement('div'); t.id = 'joToast'; document.body.appendChild(t);
  }
  function toast(m, ms) { ui(); var t = document.getElementById('joToast'); t.textContent = m; t.style.display = 'block'; clearTimeout(t._h); t._h = setTimeout(function () { t.style.display = 'none'; }, ms || 4000); }

  /** o: { blob, name, title, tables?, wb?, rows?, head?, sub?, extra:[{ label, fn }] } — blob may be left out for a list (rows) */
  function offer(o) {
    ui();
    o = Object.assign({}, o);
    if (o.blob && !o.tables && TBL.has(o.blob)) o.tables = TBL.get(o.blob);
    var isPdf = o.blob && /\.pdf$/i.test(o.name || '') || (o.blob && o.blob.type === 'application/pdf');
    var pdf = isPdf ? { blob: o.blob, name: o.name } : null, xl = o.blob && !isPdf && /\.xlsx$/i.test(o.name) ? { blob: o.blob, name: o.name } : null;
    var canX = !o.noXl && (!isPdf || !!(o.tables && o.tables.length)), canP = !o.noPdf && !!(isPdf || o.rows || o.wb || /\.(csv|xlsx?)$/i.test(o.name || ''));
    var d = document.getElementById('joDlg'), sz = o.blob ? ' · ' + Math.max(1, Math.round(o.blob.size / 1024)) + ' KB' : o.rows ? ' · ' + o.rows.length + ' rows' : '';
    var grp = function (k, t, on) {
      return '<div class="jg"><b class="t">' + t + '</b><div class="jr">' +
        '<button type="button" data-jo="' + k + '-open"' + (on ? '' : ' disabled') + '>Open</button>' +
        '<button type="button" data-jo="' + k + '-print"' + (on ? '' : ' disabled') + '>Print</button>' +
        '<button type="button" data-jo="' + k + '-down"' + (on ? '' : ' disabled') + '>Download</button>' +
        '<button type="button" class="p" data-jo="' + k + '-share"' + (on ? '' : ' disabled') + '>Share</button></div></div>';
    };
    document.getElementById('joBody').innerHTML = '<h2>' + esc(o.title || 'File ready') + '</h2><div class="jm">' + esc(o.name || '') + sz + (o.sub ? '<br>' + o.sub : '') + '</div>' +
      (canP ? grp('pdf', 'PDF', true) : '') + (canX ? grp('xl', 'Excel', true) : '') +
      (o.extra && o.extra.length ? '<div class="jx">' + o.extra.map(function (x, i) { return '<button type="button" data-jx="' + i + '">' + esc(x.label) + '</button>'; }).join('') + '</div>' : '') +
      '<div id="joMsg"></div><button id="joClose" type="button">Close</button>';
    if (!d.open) d.showModal();
    var msg = function (m) { var e = document.getElementById('joMsg'); if (e) e.textContent = m || ''; };
    document.getElementById('joClose').onclick = function () { d.close(); };
    var getPdf = function () { return pdf ? Promise.resolve(pdf) : (msg('Making the PDF…'), toPdf(o).then(function (r) { pdf = r; msg(''); return r; })); };
    var getXl = function () { return xl ? Promise.resolve(xl) : (msg('Making the Excel file…'), toXlsx(o).then(function (r) { xl = r; msg(''); return r; })); };
    [].forEach.call(d.querySelectorAll('[data-jo]'), function (b) {
      b.onclick = function () {
        var k = b.dataset.jo.split('-'), kind = k[0], act = k[1];
        if (kind === 'xl' && (act === 'open' || act === 'print')) { openTable(o, act === 'print').catch(function (e) { msg(e.message); }); return; }
        (kind === 'pdf' ? getPdf() : getXl()).then(function (f) {
          if (act === 'open') openPdf(f.blob, f.name, false);
          else if (act === 'print') openPdf(f.blob, f.name, true);
          else if (act === 'down') { download(f.blob, f.name); toast('Downloaded: ' + f.name); }
          else return share(f.blob, f.name, o.title);
        }).catch(function (e) { msg('Could not make it: ' + (e && e.message || e)); });
      };
    });
    [].forEach.call(d.querySelectorAll('[data-jx]'), function (b) { b.onclick = function () { var x = o.extra[+b.dataset.jx]; try { x.fn(o, d); } catch (e) { msg(e.message); } }; });
  }
  /** a list on screen -> same panel (PDF / Excel made on demand) */
  function list(o) { offer({ title: o.title, sub: o.sub, head: o.head, rows: o.rows, name: (o.file || String(o.title || 'List').replace(/[\\/:*?"<>|]+/g, ' ').trim()) + '.xlsx', extra: o.extra }); }

  window.JCWOUT = { offer: offer, list: list, toast: toast, tablesOf: function (b) { return TBL.get(b); }, download: download, share: share };
})();
