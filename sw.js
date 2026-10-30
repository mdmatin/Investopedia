/* Centric service worker — opens instantly and works offline.
   The app's own files (this site only) come from the phone's copy straight away,
   while a fresh copy is fetched in the background. When the fresh page differs,
   the app is told, and it switches to the new version the next time it is opened.
   Calls to Google (your data) never pass through here. */
const CACHE = 'centric-v3';   // a new name clears the old copies when this version takes over
const CORE = ['./', './index.html', './manifest.json', './icon.png', './apple-touch-icon.png', './favicon-32.png'];

self.addEventListener('install', e => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then(c => Promise.all(CORE.map(u => c.add(new Request(u, { cache: 'reload' })).catch(() => {})))));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim())
    .then(() => tell('centric-updated')));   // an open app reloads into the new copy when next brought back
});

async function tell(msg){
  const list = await self.clients.matchAll({ type: 'window' });
  list.forEach(c => c.postMessage(msg));
}

self.addEventListener('fetch', e => {
  const r = e.request;
  if (r.method !== 'GET') return;
  const u = new URL(r.url);
  if (u.origin !== self.location.origin) return;            // API and fonts go straight to the network
  const page = r.mode === 'navigate' || /\/(index\.html)?$/.test(u.pathname);
  const key = page ? './' : r;
  e.respondWith((async () => {
    const c = await caches.open(CACHE);
    const hit = await c.match(key, { ignoreSearch: true });
    const prev = hit ? hit.clone() : null;   // read before hit goes to the browser: a used body can't be cloned later (that kept the old page forever)
    // by URL: a page request can't be re-sent with options (that threw, so the copy never updated)
    const fresh = fetch(r.url, { cache: 'no-store', credentials: 'same-origin' }).then(async res => {
      if (!res || !res.ok) return res;
      if (page){
        const [a, b] = await Promise.all([res.clone().text(), prev ? prev.text() : Promise.resolve(null)]);
        await c.put(key, res.clone());
        if (b != null && a !== b) tell('centric-updated');   // a new build: the app picks it up on its next open
      } else await c.put(key, res.clone());
      return res;
    }).catch(() => null);
    if (hit){ e.waitUntil(fresh); return hit; }               // instant: the copy on the phone
    const res = await fresh;                                  // first visit: wait for the network
    return res || new Response('Offline', { status: 503, headers: { 'Content-Type': 'text/plain' } });
  })());
});

/* Draft alerts. The script sends an empty push when a bank text becomes a draft (a push
   with a message would need encryption the script can't do), so the text for the alert is
   fetched from the script here, using the address and key the app hands over (stored below).
   Tapping the alert opens Spend, or brings the open app there. */
const CFG_CACHE = 'centric-cfg', CFG_URL = './__cfg';
self.addEventListener('message', e => {
  const m = e.data || {};
  if (m.type === 'cfg' && m.url) e.waitUntil(caches.open(CFG_CACHE).then(c => c.put(CFG_URL, new Response(JSON.stringify({ url: m.url, key: m.key })))));
});
async function pushText(){
  try {
    const c = await caches.open(CFG_CACHE), r = await c.match(CFG_URL); if (!r) return null;
    const cfg = await r.json(), u = new URL(cfg.url);
    u.searchParams.set('key', cfg.key); u.searchParams.set('action', 'pushMsg');
    const j = await (await fetch(u, { redirect: 'follow', cache: 'no-store' })).json();
    return j && j.ok && j.data && j.data.title ? j.data : null;
  } catch (err) { return null; }
}
self.addEventListener('push', e => {
  e.waitUntil((async () => {
    let m = null; try { m = e.data ? e.data.json() : null; } catch (err) {}
    if (!m) m = await pushText();
    // iPhone needs every push to show something, so there is always a fallback line
    const title = (m && m.title) || 'New draft', body = (m && m.body) || 'A bank text is waiting in Spend.';
    await self.registration.showNotification(title, { body, tag: 'centric-draft', icon: './icon.png', badge: './favicon-32.png', data: { go: 'spend' } });
  })());
});
self.addEventListener('notificationclick', e => {
  e.notification.close();
  e.waitUntil((async () => {
    const list = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
    for (const c of list) { try { await c.focus(); } catch (err) {} c.postMessage({ type: 'go', view: 'spend' }); return; }
    await self.clients.openWindow('./?go=spend');
  })());
});
