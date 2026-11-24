/**
 * Service Worker — 支援賽道無網路環境離線快取 (PWA Offline-First)
 */
const CACHE_NAME = 'nfc-race-timing-v3.12.17';

const CORE_ASSETS = [
  './',
  './index.html',
  './public.html',
  './manifest.json',
  './tur-logo.jpg',
  './tur-icon.png',
  './icon.svg',
  './css/style.css',
  './js/config.js',
  './js/db.js',
  './js/timing-engine.js',
  './js/nfc.js',
  './js/sync.js',
  './js/excel.js',
  './js/certificate.js',
  './js/app.js',
  './js/public-app.js'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(CORE_ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))
      )
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // Google Apps Script API 永遠走網路，不快取
  if (url.hostname.includes('script.google.com') || url.hostname.includes('googleusercontent.com')) {
    return;
  }

  // 非 GET 請求直接放行
  if (event.request.method !== 'GET') {
    return;
  }

  // 對外部 CDN (Tailwind / SheetJS) 採 Stale-While-Revalidate 快取，離線也能開
  if (url.origin !== self.location.origin) {
    event.respondWith(
      caches.match(event.request).then((cached) => {
        const networkFetch = fetch(event.request)
          .then((resp) => {
            if (resp && resp.status === 200) {
              const copy = resp.clone();
              caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
            }
            return resp;
          })
          .catch(() => cached);
        return cached || networkFetch;
      })
    );
    return;
  }

  // 本地核心檔案採 Network-First（強制向伺服器驗證最新版本），斷網自動回退快取
  event.respondWith(
    fetch(event.request, { cache: 'no-cache' })
      .then((resp) => {
        if (resp && resp.status === 200) {
          const copy = resp.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
        }
        return resp;
      })
      .catch(() => caches.match(event.request, { ignoreSearch: true }))
  );
});
