/* Centric service worker — opens instantly and works offline.
   The app's own files (this site only) come from the phone's copy straight away,
   while a fresh copy is fetched in the background. When the fresh page differs,
   the app is told, and it switches to the new version the next time it is opened.
   Calls to Google (your data) never pass through here. */
const CACHE = 'centric-v1';
const CORE = ['./', './index.html', './manifest.json', './icon.png', './apple-touch-icon.png', './favicon-32.png'];

self.addEventListener('install', e => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then(c => Promise.all(CORE.map(u => c.add(u).catch(() => {})))));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
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
    const fresh = fetch(r, { cache: 'no-cache' }).then(async res => {
      if (!res || !res.ok) return res;
      if (page){
        const [a, b] = await Promise.all([res.clone().text(), hit ? hit.clone().text() : Promise.resolve(null)]);
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
