/* JCW app — one spinner for every page.
   Shows a small "Please wait…" pill (with a spinning ring) whenever the page waits:
   - any fetch (server calls, files, logo …)
   - building PDFs / Excel / Word files and loading the PDF tools
   It appears only if the wait is longer than 0.3 s, so quick actions do not flicker.
   Pages can also call  jcwBusy(promise or function, 'message')  for their own long jobs. */
(function(){
  if(window.jcwBusy) return;
  let n = 0, timer = null, el = null, label = '';
  const css = `#jcwSpin{position:fixed;left:50%;top:14px;transform:translate(-50%,-8px);z-index:2147483000;display:flex;align-items:center;gap:10px;
    padding:9px 16px 9px 12px;border-radius:999px;background:rgba(40,28,20,.92);color:#fff;font:600 14px/1.2 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;
    box-shadow:0 6px 24px rgba(0,0,0,.25);opacity:0;pointer-events:none;transition:opacity .18s,transform .18s}
  #jcwSpin.on{opacity:1;transform:translate(-50%,0)}
  #jcwSpin i{width:20px;height:20px;border-radius:50%;border:3px solid rgba(255,255,255,.3);border-top-color:#f3c77a;animation:jcwSp .75s linear infinite;flex:none}
  @keyframes jcwSp{to{transform:rotate(360deg)}}
  @media print{#jcwSpin{display:none!important}}`;
  function make(){
    if(el) return el;
    const st = document.createElement('style'); st.textContent = css; document.head.appendChild(st);
    el = document.createElement('div'); el.id = 'jcwSpin'; el.setAttribute('role', 'status'); el.setAttribute('aria-live', 'polite');
    el.innerHTML = '<i></i><span>Please wait…</span>';
    document.body.appendChild(el); return el;
  }
  function show(){ if(!document.body) return; const e = make(); e.querySelector('span').textContent = label || 'Please wait…'; e.classList.add('on'); document.documentElement.style.cursor = 'progress'; }
  function hide(){ if(el) el.classList.remove('on'); document.documentElement.style.cursor = ''; label = ''; }
  function start(msg){ n++; if(msg) label = msg; if(n === 1 && !timer) timer = setTimeout(() => { timer = null; if(n > 0) show(); }, 300); else if(n > 0 && el && el.classList.contains('on') && msg) el.querySelector('span').textContent = msg; }
  function stop(){ n = Math.max(0, n - 1); if(n === 0){ if(timer){ clearTimeout(timer); timer = null; } hide(); } }

  /** jcwBusy(promise | function, message) — returns the same result */
  window.jcwBusy = function(job, msg){
    start(msg);
    let p; try { p = Promise.resolve(typeof job === 'function' ? job() : job); } catch(e){ stop(); throw e; }
    return p.finally(stop);
  };
  window.jcwBusy.start = start; window.jcwBusy.stop = stop;

  /* every fetch */
  const f0 = window.fetch && window.fetch.bind(window);
  if(f0) window.fetch = function(u, o){
    const quiet = o && o.quiet;   // fetch(url, { quiet:true }) = background, no spinner
    if(quiet) return f0(u, o);
    let msg = '';
    try { if(o && o.body && typeof o.body === 'string' && /"action"\s*:\s*"[^"]*(save|sync|upload|push|submit)/i.test(o.body)) msg = 'Saving…'; } catch(e){}
    start(msg || 'Loading…');
    return f0(u, o).finally(stop);
  };

  /* long jobs on the pages (PDF, Excel, Word, loading the PDF tools) */
  const JOBS = { buildPdf:'Making PDF…', makePdf:'Making PDF…', onePdf:'Making PDF…', allPdfs:'Making all files…', teacherPdf:'Making PDF…', onPdf:'Making PDF…',
    ssPdfCopy:'Making PDF…', makeXlsx:'Making Excel file…', makeDocx:'Making Word file…', readPdf:'Reading PDF…', pdfLib:'Loading PDF tools…', loadScript:'Loading…',
    readDocx:'Reading Word file…', printRows:'Preparing…', prepareCerts:'Preparing…', xlsx:'Making Excel file…', toXlsx:'Making Excel file…',
    printSheet:'Preparing…', ewMake:'Preparing…', rpGo:'Preparing…', tpGo:'Preparing…', wrGo:'Preparing…', regRun:'Preparing…' };
  function wrapJobs(){
    Object.keys(JOBS).forEach(k => {
      const fn = window[k]; if(typeof fn !== 'function' || fn.__jcw) return;
      const w = function(){ start(JOBS[k]); let r; try { r = fn.apply(this, arguments); } catch(e){ stop(); throw e; }
        if(r && typeof r.then === 'function') return r.finally(stop);
        stop(); return r; };
      w.__jcw = true;
      try { window[k] = w; } catch(e){}
    });
  }
  if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', wrapJobs); else wrapJobs();
  window.addEventListener('load', wrapJobs);
})();
