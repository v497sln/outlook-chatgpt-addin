// Cache-first service worker: the whole app works offline after the first visit.
const CACHE = 'squadsage-v1';
const FILES = ['./', 'index.html', 'styles.css', 'icon.svg', 'manifest.webmanifest',
  'js/app.js', 'js/engine.js', 'js/data.js', 'js/store.js', 'js/solver-worker.js'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  e.respondWith(caches.match(e.request).then(hit => hit || fetch(e.request).then(res => {
    const copy = res.clone();
    if (res.ok && new URL(e.request.url).origin === location.origin) caches.open(CACHE).then(c => c.put(e.request, copy));
    return res;
  })));
});
