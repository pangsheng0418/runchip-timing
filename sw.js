/**
 * Service Worker — 支援賽道無網路環境離線快取 (PWA Offline-First)
 */
const CACHE_NAME = 'nfc-race-timing-v3.13.1';

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

  // HTML 頁面採 Network-First（確保版本號更新即時生效）；其餘帶版本參數之 JS/CSS/圖片採 Cache-First 加速秒開
  const isHtml = event.request.mode === 'navigate' || url.pathname.endsWith('.html') || url.pathname.endsWith('/');
  if (!isHtml) {
    event.respondWith(
      caches.match(event.request).then((cached) => {
        if (cached) return cached;
        return fetch(event.request).then((resp) => {
          if (resp && resp.status === 200) {
            const copy = resp.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
          }
          return resp;
        });
      })
    );
    return;
  }

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
