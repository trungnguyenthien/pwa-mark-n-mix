/**
 * Service Worker cho PWA Mark 'n Mix (Trộn Đề Thi)
 * Tuân thủ 100% quy chuẩn pwa-standards.md:
 * - Scope Isolation & Fetch Safety
 * - Bỏ qua non-GET, blob:, data:, chrome-extension:
 * - Precache 100% tài nguyên và thư viện nội bộ
 * - Chiến lược Stale-While-Revalidate kết hợp Cache-First cho assets tĩnh
 * - Hỗ trợ force update và skipWaiting an toàn
 */

const CACHE_NAME = 'pwa-mark-n-mix-v1';

const PRECACHE_ASSETS = [
  './',
  './index.html',
  './index.html?v=1',
  './css/style.css',
  './css/style.css?v=1',
  './js/app.js',
  './js/app.js?v=1',
  './js/ui.js',
  './js/ui.js?v=1',
  './js/parser.js',
  './js/parser.js?v=1',
  './js/shuffler.js',
  './js/shuffler.js?v=1',
  './js/docx-builder.js',
  './js/docx-builder.js?v=1',
  './js/answer-key.js',
  './js/answer-key.js?v=1',
  './libs/docx.umd.js',
  './libs/docx.umd.js?v=1',
  './libs/jszip.min.js',
  './libs/jszip.min.js?v=1',
  './manifest.webmanifest',
  './manifest.webmanifest?v=1',
  './icons/icon-192.png',
  './icons/icon-512.png'
];

// Cài đặt Service Worker và nạp toàn bộ precache assets
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(PRECACHE_ASSETS);
    })
  );
});

// Kích hoạt Service Worker và xóa sạch mọi cache cũ khác CACHE_NAME hiện tại
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

// Lắng nghe lệnh từ ứng dụng (ví dụ: skipWaiting khi bấm Force Update)
self.addEventListener('message', (event) => {
  if (event.data && event.data.action === 'skipWaiting') {
    self.skipWaiting();
  }
});

// Xử lý sự kiện fetch với bảo vệ phạm vi (Scope Isolation)
self.addEventListener('fetch', (event) => {
  const request = event.request;

  // 1. Chỉ can thiệp phương thức GET
  if (request.method !== 'GET') {
    return;
  }

  const url = new URL(request.url);

  // 2. Bỏ qua các scheme không phải HTTP/HTTPS (blob:, data:, chrome-extension:)
  if (url.protocol !== 'http:' && url.protocol !== 'https:') {
    return;
  }

  // 3. Quy chuẩn: Scope Isolation - không can thiệp vào request của các app khác trên cùng domain
  if (url.origin === self.location.origin && !request.url.startsWith(self.registration.scope)) {
    return;
  }

  // 4. Xử lý điều hướng Navigation (SPA Fallback trong phạm vi scope)
  if (request.mode === 'navigate') {
    event.respondWith(
      fetch(request).catch(() => {
        return caches.match('./index.html', { ignoreSearch: true });
      })
    );
    return;
  }

  // 5. Cache First / Stale While Revalidate với ignoreSearch: true để hỗ trợ versioning ?v=X
  event.respondWith(
    caches.match(request, { ignoreSearch: true }).then((cachedResponse) => {
      if (cachedResponse) {
        // Cập nhật ngầm trong nền nếu có mạng
        fetch(request).then((networkResponse) => {
          if (networkResponse && networkResponse.status === 200) {
            caches.open(CACHE_NAME).then((cache) => {
              cache.put(request, networkResponse);
            });
          }
        }).catch(() => {
          // Khi offline, hoàn toàn yên tâm dùng cachedResponse
        });
        return cachedResponse;
      }

      // Chưa có trong cache: tải từ mạng và lưu vào cache
      return fetch(request).then((networkResponse) => {
        if (!networkResponse || networkResponse.status !== 200 || networkResponse.type !== 'basic') {
          return networkResponse;
        }
        const responseToCache = networkResponse.clone();
        caches.open(CACHE_NAME).then((cache) => {
          cache.put(request, responseToCache);
        });
        return networkResponse;
      });
    })
  );
});
