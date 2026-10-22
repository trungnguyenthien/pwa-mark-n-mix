/**
 * Module Chính của Ứng Dụng (App Entrypoint)
 * Tuân thủ đầy đủ quy chuẩn PWA Standards:
 * - Đăng ký Service Worker với scope tương đối './'
 * - Quản lý nút Cài đặt (#btnInstall) cho Android/Chromium & iOS Safari
 * - Quản lý nút Cập nhật ứng dụng (#btnUpdateApp) cơ chế Force Update
 * - Trạng thái mạng Online / Offline
 */

import { initUI, showSnackbar } from './ui.js';

let swRegistration = null;
let deferredInstallPrompt = null;

document.addEventListener('DOMContentLoaded', () => {
  // 1. Khởi tạo toàn bộ giao diện người dùng
  initUI();

  // 2. Quản lý trạng thái mạng Online / Offline
  initNetworkStatus();

  // 3. Khởi tạo cơ chế Cài đặt PWA (#btnInstall)
  initPWAInstallation();

  // 4. Khởi tạo cơ chế Service Worker và Nút Cập nhật App (#btnUpdateApp)
  initServiceWorker();
});

/**
 * Theo dõi kết nối mạng
 */
function initNetworkStatus() {
  const netStatus = document.getElementById('netStatus');
  if (!netStatus) return;

  function updateStatus() {
    if (navigator.onLine) {
      netStatus.className = 'net-status-pill online';
      netStatus.querySelector('.status-text').textContent = 'Online';
    } else {
      netStatus.className = 'net-status-pill offline';
      netStatus.querySelector('.status-text').textContent = 'Offline';
      showSnackbar('Bạn đang ở chế độ Offline (Ngoại tuyến)');
    }
  }

  window.addEventListener('online', () => {
    updateStatus();
    showSnackbar('Đã kết nối Internet trở lại!');
  });
  window.addEventListener('offline', updateStatus);

  updateStatus();
}

/**
 * Quản lý cài đặt PWA theo mục 6 pwa-standards.md
 */
function initPWAInstallation() {
  const btnInstall = document.getElementById('btnInstall');
  const iosInstallModal = document.getElementById('iosInstallModal');
  const btnCloseIosModal = document.getElementById('btnCloseIosModal');
  const btnGotItIos = document.getElementById('btnGotItIos');

  // Kiểm tra chế độ standalone
  const isStandalone =
    window.matchMedia('(display-mode: standalone)').matches ||
    window.navigator.standalone === true;

  // Phát hiện iOS Safari
  const isIOS =
    /iPad|iPhone|iPod/.test(navigator.userAgent) ||
    (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);

  if (isStandalone) {
    if (btnInstall) btnInstall.classList.add('hidden');
  } else if (isIOS && btnInstall) {
    // iOS Safari không hỗ trợ beforeinstallprompt, hiện nút để mở modal hướng dẫn
    btnInstall.classList.remove('hidden');
  }

  // 1. Lắng nghe beforeinstallprompt trên Chromium (Android / Desktop)
  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredInstallPrompt = e;
    if (!isStandalone && btnInstall) {
      btnInstall.classList.remove('hidden');
    }
  });

  // 2. Lắng nghe khi cài đặt hoàn tất
  window.addEventListener('appinstalled', () => {
    deferredInstallPrompt = null;
    if (btnInstall) btnInstall.classList.add('hidden');
    showSnackbar('Cài đặt ứng dụng thành công!');
  });

  // 3. Xử lý khi nhấn nút Install
  if (btnInstall) {
    btnInstall.addEventListener('click', async () => {
      if (deferredInstallPrompt) {
        // Luồng native prompt cho Android Chrome / Desktop
        deferredInstallPrompt.prompt();
        const { outcome } = await deferredInstallPrompt.userChoice;
        if (outcome === 'accepted') {
          btnInstall.classList.add('hidden');
        }
        deferredInstallPrompt = null;
      } else if (iosInstallModal) {
        // Luồng mở hướng dẫn cho iOS Safari
        iosInstallModal.classList.add('open');
      }
    });
  }

  // Đóng modal hướng dẫn iOS
  function closeIosModal() {
    if (iosInstallModal) {
      iosInstallModal.classList.remove('open');
    }
  }

  if (btnCloseIosModal) btnCloseIosModal.addEventListener('click', closeIosModal);
  if (btnGotItIos) btnGotItIos.addEventListener('click', closeIosModal);

  // Đóng bằng Esc
  window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      closeIosModal();
      const exportModal = document.getElementById('exportModal');
      if (exportModal) exportModal.classList.remove('open');
    }
  });
}

/**
 * Đăng ký Service Worker và Quản lý Force Update (#btnUpdateApp)
 * Tuân thủ mục 3, 4, 6.D của pwa-standards.md
 */
function initServiceWorker() {
  const btnUpdateApp = document.getElementById('btnUpdateApp');
  if (!btnUpdateApp) return;

  const btnLabel = btnUpdateApp.querySelector('.btn-label');
  const spinIcon = btnUpdateApp.querySelector('.spin-icon');

  if ('serviceWorker' in navigator) {
    // Đăng ký Service Worker với scope tương đối './'
    navigator.serviceWorker
      .register('./sw.js', { scope: './' })
      .then((registration) => {
        swRegistration = registration;

        // Kiểm tra nếu có worker đang chờ kích hoạt (waiting)
        if (registration.waiting) {
          notifyUpdateReady();
        }

        // Lắng nghe bản cập nhật mới được tìm thấy
        registration.addEventListener('updatefound', () => {
          const newWorker = registration.installing;
          if (!newWorker) return;

          newWorker.addEventListener('statechange', () => {
            if (newWorker.state === 'installed' && navigator.serviceWorker.controller) {
              notifyUpdateReady();
            }
          });
        });
      })
      .catch((error) => {
        console.warn('Lỗi khi đăng ký Service Worker:', error);
      });
  }

  // Thông báo bản cập nhật sẵn sàng
  function notifyUpdateReady() {
    btnUpdateApp.classList.add('update-ready');
    if (btnLabel) btnLabel.textContent = 'Update Ready';
    showSnackbar("Đã có bản cập nhật mới! Nhấn 'Update Ready' để áp dụng.");
  }

  // Xử lý khi nhấn nút Cập Nhật App
  btnUpdateApp.addEventListener('click', async () => {
    // 1. Kiểm tra trạng thái mạng
    if (!navigator.onLine) {
      showSnackbar('Không thể cập nhật khi đang offline. Hãy kiểm tra kết nối mạng!');
      return;
    }

    // 2. Chuyển sang trạng thái đang cập nhật
    btnUpdateApp.disabled = true;
    if (spinIcon) spinIcon.classList.add('spinning');
    if (btnLabel) btnLabel.textContent = 'Updating...';

    try {
      if (swRegistration) {
        // Kiểm tra phiên bản mới từ server
        await swRegistration.update();

        // Gửi thông điệp skipWaiting nếu có worker đang waiting
        if (swRegistration.waiting) {
          swRegistration.waiting.postMessage({ action: 'skipWaiting' });
        }
      }

      // Xóa sạch các cache cũ
      if ('caches' in window) {
        const cacheNames = await caches.keys();
        await Promise.all(cacheNames.map((name) => caches.delete(name)));
      }

      // Hủy đăng ký service worker cũ nếu có để nạp mới hoàn toàn
      if (swRegistration) {
        await swRegistration.unregister();
      }

      showSnackbar('Đang tải lại ứng dụng với phiên bản mới nhất...');

      // Reload ứng dụng với query string ép bypass cache
      setTimeout(() => {
        window.location.href = window.location.pathname + '?forceUpdate=' + Date.now();
      }, 600);
    } catch (err) {
      console.error('Lỗi khi cập nhật:', err);
      showSnackbar('Cập nhật thất bại. Vui lòng thử lại sau.');
      btnUpdateApp.disabled = false;
      if (spinIcon) spinIcon.classList.remove('spinning');
      if (btnLabel) btnLabel.textContent = 'Update App';
    }
  });
}
