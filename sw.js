/* Service worker Almeyra's Shop: menyimpan semua berkas aplikasi agar bisa dibuka tanpa internet.
   Naikkan nomor VERSI setiap kali berkas aplikasi diubah. */
const VERSI = 'almeyras-shop-v1';
const BERKAS = ['./', './index.html', './app.js', './report.js', './jszip.min.js', './manifest.webmanifest',
  './icons/icon-192.png', './icons/icon-512.png', './icons/icon-maskable-512.png'];

self.addEventListener('install', (e) => {
  e.waitUntil(caches.open(VERSI).then((c) => c.addAll(BERKAS)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', (e) => {
  e.waitUntil(caches.keys()
    .then((ks) => Promise.all(ks.filter((k) => k !== VERSI).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  e.respondWith(
    caches.match(req, { ignoreSearch: true }).then((hit) => hit || fetch(req).then((res) => {
      if (res.ok) { const salinan = res.clone(); caches.open(VERSI).then((c) => c.put(req, salinan)); }
      return res;
    }).catch(() => caches.match('./index.html')))
  );
});
