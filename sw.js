// Keeps the Supplybot app page, its icon and manifest on the phone so it still opens offline.
const CACHE = 'supplybot-launcher-v7';
const FILES = ['./', './index.html', './config.js', './manifest.webmanifest', './icon-192.png', './icon-512.png', './apple-touch-icon.png', './supplybot.svg'];
self.addEventListener('install', e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES))); self.skipWaiting(); });
self.addEventListener('activate', e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))); self.clients.claim(); });
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (u.origin !== location.origin) return;   // Supplybot itself always comes fresh from Google
  // fresh from the network when online (so edits to config.js / index.html show up), the saved copy when offline
  e.respondWith(fetch(e.request).then(res => { const copy = res.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)).catch(() => {}); return res; })
    .catch(() => caches.match(e.request, { ignoreSearch: true })));
});
