/* Luck Swan Sales — offline cache */
const CACHE = 'ls-sales-v29';
const ASSETS = [
  'css/app.css',
  'fonts.css',
  'fonts/BaiJamjuree-500-latin.woff2',
  'fonts/BaiJamjuree-500-thai.woff2',
  'fonts/BaiJamjuree-600-latin.woff2',
  'fonts/BaiJamjuree-600-thai.woff2',
  'fonts/BaiJamjuree-700-latin.woff2',
  'fonts/BaiJamjuree-700-thai.woff2',
  'fonts/IBMPlexSansThai-400-latin.woff2',
  'fonts/IBMPlexSansThai-400-thai.woff2',
  'fonts/IBMPlexSansThai-500-latin.woff2',
  'fonts/IBMPlexSansThai-500-thai.woff2',
  'fonts/IBMPlexSansThai-600-latin.woff2',
  'fonts/IBMPlexSansThai-600-thai.woff2',
  'icons/apple-touch-icon.png',
  'icons/favicon-32.png',
  'icons/icon-192.png',
  'icons/icon-512.png',
  'icons/maskable-192.png',
  'icons/maskable-512.png',
  'img/food1-0.jpg',
  'img/food2-0.jpg',
  'img/food2-1.jpg',
  'img/logo-wide-sm.png',
  'img/logo-wide.png',
  'img/meet-kinnaree.jpg',
  'img/meet-prayamengrai.jpg',
  'img/meet-ruenhong.jpg',
  'img/meet-thikhampon.jpg',
  'img/room-deluxe.jpg',
  'img/room-grand-deluxe.jpg',
  'img/room-junior-suite.jpg',
  'img/room-superior.jpg',
  'index.html',
  'js/api.js',
  'js/app.js',
  'js/catalog-admin.js',
  'js/data.js',
  'js/sheet.js',
  'js/util.js',
  'js/v-catalog.js',
  'js/v-dash.js',
  'js/v-ops.js',
  'js/v-quote.js',
  'manifest.webmanifest',
  './'
];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return;
  const path = new URL(req.url).pathname;
  if (path.includes('/api/') && !path.includes('/api/images/')) return; // live data: never cache (uploaded photos are cached)
  // pages, scripts, styles and data: network first so edits show up right away; cache is only the offline fallback
  if (req.mode === 'navigate' || /\.(html|js|css|webmanifest)$/.test(path) || path.endsWith('/')) {
    e.respondWith(fetch(req, { cache: 'no-cache' }).then(r => {
      if (r.ok) { const copy = r.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
      return r;
    }).catch(() => caches.match(req).then(r => r || (req.mode === 'navigate' ? caches.match('index.html') : undefined))));
    return;
  }
  // images and fonts rarely change: cache first
  e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(r => {
    if (r.ok) { const copy = r.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
    return r;
  })));
});
