// Bump VERSION every time you change index.html so phones pick up the update.
const VERSION = 'csa-v11';
const CORE = ['./', './index.html', './manifest.json', './icon-192.png', './icon-512.png'];

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
    e.respondWith(
      fetch(req).then(res => {
        const copy = res.clone();
        caches.open(VERSION).then(c => c.put('./index.html', copy));
        return res;
      }).catch(() => caches.match('./index.html'))
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

self.addEventListener('sync', e => {
  if (e.tag === 'csa-sync') e.waitUntil(bgSync());
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
      await self.registration.showNotification('Attendance synced \u2713', {
        body: savedBatch.length + ' record(s) reached the sheet.',
        icon: './icon-192.png',
        tag: 'csa-synced'
      });
    } catch (err) {}
  }
}

self.addEventListener('notificationclick', e => {
  e.notification.close();
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(list => {
    for (const c of list) { if ('focus' in c) return c.focus(); }
    return self.clients.openWindow('./');
  }));
});
