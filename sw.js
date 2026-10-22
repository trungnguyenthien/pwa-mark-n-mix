/**
 * Service Worker cho PWA Mark 'n Mix (Công Cụ Trộn Đề Trắc Nghiệm)
 * Tuân thủ chuẩn PWA hiện đại:
 * - Scope Isolation & Safe Fetch handling (chỉ xử lý GET, bỏ qua blob:/data:/chrome-extension:)
 * - Precache 100% tài nguyên nội bộ, thư viện Word và đề mẫu
 * - Cache-First cho tệp tĩnh với fallback về mạng
 * - Hỗ trợ force update qua message 'skipWaiting'
 */

const CACHE_NAME = 'pwa-mark-n-mix-v11';

const PRECACHE_ASSETS = [
  './',
  './index.html',
  './css/style.css',
  './js/app.bundle.js',
  './js/app.js',
  './js/ui.js',
  './js/parser.js',
  './js/shuffler.js',
  './js/docx-builder.js',
  './js/answer-key.js',
  './js/docx-importer.js',
  './libs/docx.umd.js',
  './libs/jszip.min.js',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './mau_de_thi_co_marker.docx'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(PRECACHE_ASSETS);
    }).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((name) => {
          if (name !== CACHE_NAME) {
            return caches.delete(name);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  const url = new URL(req.url);

  // Chỉ chặn các yêu cầu GET trong cùng origin hoặc CDN cần thiết
  if (req.method !== 'GET') return;
  if (!url.protocol.startsWith('http')) return;

  event.respondWith(
    caches.match(req).then((cachedResponse) => {
      if (cachedResponse) {
        // Cập nhật ngầm (stale-while-revalidate) đối với các tệp nội bộ
        fetch(req).then((networkResponse) => {
          if (networkResponse && networkResponse.status === 200 && networkResponse.type === 'basic') {
            caches.open(CACHE_NAME).then((cache) => cache.put(req, networkResponse));
          }
        }).catch(() => {
          // Bỏ qua lỗi ngầm khi offline
        });
        return cachedResponse;
      }

      return fetch(req).then((networkResponse) => {
        if (!networkResponse || networkResponse.status !== 200) {
          return networkResponse;
        }
        const responseToCache = networkResponse.clone();
        caches.open(CACHE_NAME).then((cache) => {
          cache.put(req, responseToCache);
        });
        return networkResponse;
      }).catch(() => {
        // Nếu offline và yêu cầu trang HTML, trả về trang chủ đã cache
        if (req.headers.get('accept')?.includes('text/html')) {
          return caches.match('./index.html');
        }
      });
    })
  );
});

self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});
