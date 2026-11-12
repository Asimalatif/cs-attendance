// Bump VERSION every time you change index.html so phones pick up the update.
const VERSION = 'csa-v142';

// Approved leaves (leaves.html): self.LV = storage + sending of the Leave In-Charge's waiting leaves
try { importScripts('./leavesync.js'); } catch (e) {}
const CORE = ['./', './index.html', './admission.html', './grandtest.html', './detention.html', './board.html', './boardparse.js', './spinner.js', './certificates.html', './idcards.html', './idcard/sheet.svg', './idcard/sheet-template.pdf', './idcard/Carlito-Bold.ttf', './idcard/Carlito-Bold.woff2', './idcard/principal-sign.png', './idcheck.html', './duties.html', './reports.html', './bunkers.html', './leaves.html', './shortage.html', './recordsummary.html', './timetable.html', './admintools.html', './jcwkit.js', './leavesync.js', './jspdf.umd.min.js', './jspdf.plugin.autotable.min.js', './manifest.json', './icon-192.png', './jcw-logo.png', './icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(CORE)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;                       // saves (POST) always go to the network
  const url = new URL(req.url);
  if (url.origin !== location.origin) return;             // never touch Google Apps Script calls

  // Pages: network first, fall back to the saved copy when offline
  if (req.mode === 'navigate') {
    const page = /admission\.html$/.test(url.pathname) ? './admission.html' : /grandtest\.html$/.test(url.pathname) ? './grandtest.html' : /detention\.html$/.test(url.pathname) ? './detention.html' : /student\.html$/.test(url.pathname) ? './student.html' : /board\.html$/.test(url.pathname) ? './board.html' : /certificates\.html$/.test(url.pathname) ? './certificates.html' : /idcards\.html$/.test(url.pathname) ? './idcards.html' : /idcheck\.html$/.test(url.pathname) ? './idcheck.html' : /duties\.html$/.test(url.pathname) ? './duties.html' : /reports\.html$/.test(url.pathname) ? './reports.html' : /bunkers\.html$/.test(url.pathname) ? './bunkers.html' : /leaves\.html$/.test(url.pathname) ? './leaves.html' : /shortage\.html$/.test(url.pathname) ? './shortage.html' : /recordsummary\.html$/.test(url.pathname) ? './recordsummary.html' : /timetable\.html$/.test(url.pathname) ? './timetable.html' : /admintools\.html$/.test(url.pathname) ? './admintools.html' : './index.html';
    e.respondWith(
      fetch(req).then(res => {
        const copy = res.clone();
        caches.open(VERSION).then(c => c.put(page, copy));
        return res;
      }).catch(() => caches.match(page).then(hit => hit || caches.match('./index.html')))
    );
    return;
  }
  // Other files: saved copy first, then network
  e.respondWith(caches.match(req).then(hit => hit || fetch(req)));
});

/* ===== Background sync: sends queued attendance when internet returns, even with the app closed ===== */
function kvDb() {
  return new Promise((res, rej) => {
    const r = indexedDB.open('csa_bg', 1);
    r.onupgradeneeded = () => r.result.createObjectStore('kv');
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
}
async function kvGet(k) {
  const db = await kvDb();
  return new Promise((res, rej) => {
    const r = db.transaction('kv').objectStore('kv').get(k);
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
}
async function kvSet(k, v) {
  const db = await kvDb();
  return new Promise((res, rej) => {
    const t = db.transaction('kv', 'readwrite');
    t.objectStore('kv').put(v, k);
    t.oncomplete = () => res();
    t.onerror = () => rej(t.error);
  });
}

// The browser fires this when internet returns (also retried later if the handler fails)
self.addEventListener('sync', e => {
  if (e.tag === 'leave-sync' && self.LV) { e.waitUntil(leaveBgSend()); return; }
  if (e.tag === 'csa-sync') {
    e.waitUntil((async () => {
      await log('background sync started' + (e.lastChance ? ' (last chance)' : ''));
      await bgSync();
    })());
  }
});

// The app's "Test background send" button asks the worker to run the same job right now
self.addEventListener('message', e => {
  if (e.data && e.data.type === 'run-bg') {
    e.waitUntil(bgSync().catch(err => log('test failed: ' + (err && err.message || err))));
  }
});

async function log(msg) { try { await kvSet('bglog', { at: Date.now(), msg: msg }); } catch (e) {} }

async function bgSync() {
  const job = await kvGet('job');
  if (!job || !job.queue || !job.queue.length) { await log('nothing queued'); return; }
  if (!job.user || (!job.token && !job.pass)) { await log('no login saved for the worker (open the app online once)'); return; }
  const batch = job.queue.map(q => Object.assign({}, q));
  const auth = job.token ? { username: job.user, token: job.token } : { username: job.user, password: job.pass };
  await log('sending ' + batch.length + ' record(s)...');

  // a network failure throws here, so the browser retries automatically later
  let r;
  try {
    r = await fetch(job.api, {
      method: 'POST',
      redirect: 'follow',
      body: JSON.stringify(Object.assign({}, auth, { action: 'save', records: batch }))
    });
  } catch (err) { await log('no network yet, will retry'); throw err; }
  let res;
  try { res = JSON.parse(await r.text()); } catch (err) { await log('server sent a web page, will retry'); throw new Error('bad response'); }
  if (!res.ok) { await log('server refused: ' + (res.error || 'unknown') + (res.error === 'BAD_LOGIN' ? ' (open the app and sign in again)' : '')); return; }

  const saved = res.saved || [];
  const savedBatch = batch.filter(b => saved.includes(b.id));
  if (!savedBatch.length) { await log('server saved nothing: ' + JSON.stringify(res.failed || [])); return; }
  await log('sent ' + savedBatch.length + ' record(s) OK');

  // remove saved records from the mirrored queue (keep anything newer the page added meanwhile)
  const cur = (await kvGet('job')) || job;
  cur.queue = (cur.queue || []).filter(x => !savedBatch.some(b => b.id === x.id && b.ts === x.ts));
  await kvSet('job', cur);

  // leave a note for the page to absorb when it opens
  const done = (await kvGet('done')) || [];
  done.push({ batch: savedBatch, res: { saved: saved, failed: res.failed || [], confirmed: res.confirmed || {} } });
  await kvSet('done', done);

  const cl = await self.clients.matchAll({ includeUncontrolled: true });
  cl.forEach(c => c.postMessage({ type: 'bg-synced' }));

  // tell the teacher, but only if the app is not already open on screen
  const visible = cl.some(c => c.visibilityState === 'visible');
  if (!visible && self.Notification && Notification.permission === 'granted') {
    try {
      // one readable line per record, e.g. "C2 Chemistry · Attendance 1 Oct reached the sheet"
      const labels = job.labels || {};
      const lines = savedBatch.map(b => '\u2713 ' + (labels[b.id] || b.sheet) + ' reached the sheet');
      const body = lines.slice(0, 4).join('\n') + (lines.length > 4 ? '\n+' + (lines.length - 4) + ' more' : '');
      await self.registration.showNotification('Attendance synced \u2713', {
        body: body,
        icon: './icon-192.png',
        tag: 'csa-synced'
      });
    } catch (err) {}
  }
}

self.addEventListener('notificationclick', e => {
  e.notification.close();
  if (e.notification.data && e.notification.data.url) { e.waitUntil(self.clients.openWindow(e.notification.data.url)); return; }
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(list => {
    for (const c of list) { if ('focus' in c) return c.focus(); }
    return self.clients.openWindow('./');
  }));
});

/* ===== Approved leaves: send the Leave In-Charge's waiting leaves when internet returns ===== */
async function leaveBgSend() {
  const r = await self.LV.syncQueue();          // a network error here makes the phone retry later
  if (r.busy) return;
  const wins = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
  wins.forEach(c => c.postMessage({ type: 'leaves-synced', sent: r.sent || 0, failed: r.failed || 0 }));
  const visible = wins.some(c => c.visibilityState === 'visible');
  let allowed = false;
  try { allowed = Notification.permission === 'granted'; } catch (e) {}
  if (!visible && allowed && (r.sent || r.failed || r.authError)) {
    let body;
    if (r.authError) body = 'Leaves are waiting. Open the attendance app and sign in again to send them.';
    else {
      body = r.sent ? r.sent + ' leave change' + (r.sent > 1 ? 's' : '') + ' saved to the sheets.' : '';
      if (r.failed) body += ' ' + r.failed + ' could not be saved. Open Approved leaves to see why.';
    }
    try { await self.registration.showNotification('Approved leaves', { body: body.trim(), icon: './icon-192.png', tag: 'leave-sync', data: { url: './leaves.html' } }); } catch (e) {}
  }
}
