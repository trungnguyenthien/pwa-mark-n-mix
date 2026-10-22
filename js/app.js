/**
 * PWA Mark 'n Mix - Main Application Entry
 * Khởi tạo ứng dụng, lắng nghe sự kiện, đăng ký Service Worker PWA
 */

import {
  state,
  updateExamCodes,
  handleFileLoaded,
  renderExamPreview,
  executeMixAndExport,
  downloadAllZip,
  showSnackbar
} from './ui.js';

let deferredInstallPrompt = null;

// Khởi chạy khi DOM sẵn sàng
document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  initNetworkStatus();
  initPWA();
  initEventListeners();

  // Khởi tạo 4 mã đề mặc định bắt đầu từ 101
  updateExamCodes(4, '101');

  // Nạp sẵn đề mẫu tự động khi mở trang lần đầu để giáo viên xem ngay
  loadSampleExamDocx(true);
});

/**
 * Quản lý Dark/Light Theme
 */
function initTheme() {
  const savedTheme = localStorage.getItem('pwa_mark_mix_theme') || 'dark';
  applyTheme(savedTheme);

  const btnThemeToggle = document.getElementById('btnThemeToggle');
  if (btnThemeToggle) {
    btnThemeToggle.addEventListener('click', () => {
      const currentTheme = document.documentElement.getAttribute('data-theme') || 'dark';
      const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
      applyTheme(newTheme);
      localStorage.setItem('pwa_mark_mix_theme', newTheme);
    });
  }
}

function applyTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme);
  const iconSun = document.querySelector('.icon-sun');
  const iconMoon = document.querySelector('.icon-moon');
  if (iconSun && iconMoon) {
    if (theme === 'dark') {
      iconSun.classList.remove('hidden');
      iconMoon.classList.add('hidden');
    } else {
      iconSun.classList.add('hidden');
      iconMoon.classList.remove('hidden');
    }
  }
}

/**
 * Trạng thái mạng Online / Offline
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
      netStatus.querySelector('.status-text').textContent = 'Offline (PWA)';
      showSnackbar('Ứng dụng đang hoạt động ở chế độ Offline (PWA)!');
    }
  }

  window.addEventListener('online', updateStatus);
  window.addEventListener('offline', updateStatus);
  updateStatus();
}

/**
 * Đăng ký Service Worker và Cài đặt PWA
 */
function initPWA() {
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js').then((reg) => {
      // Lắng nghe cập nhật
      reg.addEventListener('updatefound', () => {
        const newWorker = reg.installing;
        if (newWorker) {
          newWorker.addEventListener('statechange', () => {
            if (newWorker.state === 'installed' && navigator.serviceWorker.controller) {
              showSnackbar('Đã có bản cập nhật mới! Bấm nút Cập nhật để áp dụng.');
            }
          });
        }
      });
    }).catch((err) => {
      console.warn('Service Worker registration skipped or failed:', err);
    });
  }

  // Bắt sự kiện cài đặt PWA
  const btnInstall = document.getElementById('btnInstall');
  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredInstallPrompt = e;
    if (btnInstall) {
      btnInstall.classList.remove('hidden');
    }
  });

  if (btnInstall) {
    btnInstall.addEventListener('click', async () => {
      if (!deferredInstallPrompt) return;
      deferredInstallPrompt.prompt();
      const choice = await deferredInstallPrompt.userChoice;
      if (choice.outcome === 'accepted') {
        showSnackbar('Cảm ơn bạn đã cài đặt ứng dụng!');
      }
      deferredInstallPrompt = null;
      btnInstall.classList.add('hidden');
    });
  }

  // Nút ép cập nhật phiên bản mới
  const btnUpdateApp = document.getElementById('btnUpdateApp');
  if (btnUpdateApp) {
    btnUpdateApp.addEventListener('click', async () => {
      if ('serviceWorker' in navigator) {
        const reg = await navigator.serviceWorker.getRegistration();
        if (reg) {
          await reg.update();
        }
      }
      window.location.reload();
    });
  }
}

/**
 * Nạp đề thi Word mẫu từ file mau_de_thi_co_marker.docx
 */
async function loadSampleExamDocx(silent = false) {
  try {
    const res = await fetch('./mau_de_thi_co_marker.docx');
    if (!res.ok) throw new Error(`Không thể tải tệp mẫu (${res.status})`);
    const blob = await res.blob();
    const file = new File([blob], 'mau_de_thi_co_marker.docx', {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
    await handleFileLoaded(file);
    if (!silent) {
      showSnackbar('Đã nạp đề thi mẫu thử nghiệm thành công!');
    }
  } catch (err) {
    console.warn('Không thể nạp file mẫu ngầm:', err);
  }
}

/**
 * Gán toàn bộ sự kiện giao diện
 */
function initEventListeners() {
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');
  const btnLoadSample = document.getElementById('btnLoadSample');
  const btnChangeFile = document.getElementById('btnChangeFile');
  const btnToggleRawView = document.getElementById('btnToggleRawView');
  const rawViewContainer = document.getElementById('rawViewContainer');

  // Xử lý nạp file và kéo thả (Dropzone)
  if (dropzone && fileInput) {
    // Ngăn chặn sự kiện click từ input nổi bọt ngược lên label dropzone gây đệ quy
    fileInput.addEventListener('click', (e) => {
      e.stopPropagation();
    });

    // Hỗ trợ phím Enter hoặc Space khi focus bằng bàn phím vào dropzone
    dropzone.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        fileInput.click();
      }
    });

    ['dragenter', 'dragover'].forEach(name => {
      dropzone.addEventListener(name, (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach(name => {
      dropzone.addEventListener(name, (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
      });
    });

    dropzone.addEventListener('drop', (e) => {
      const files = e.dataTransfer.files;
      if (files && files.length > 0) {
        handleFileLoaded(files[0]);
      }
    });

    fileInput.addEventListener('change', () => {
      if (fileInput.files && fileInput.files.length > 0) {
        handleFileLoaded(fileInput.files[0]);
        // Reset giá trị để nếu người dùng chọn lại cùng một file vẫn nhận diện change
        fileInput.value = '';
      }
    });
  }

  // Nút Nạp Đề Mẫu
  if (btnLoadSample) {
    btnLoadSample.addEventListener('click', (e) => {
      e.preventDefault();
      loadSampleExamDocx(false);
    });
  }

  // Đổi file khác
  if (btnChangeFile && fileInput) {
    btnChangeFile.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      fileInput.click();
    });
  }

  // Bật/tắt xem văn bản thô
  if (btnToggleRawView && rawViewContainer) {
    btnToggleRawView.addEventListener('click', () => {
      const isHidden = rawViewContainer.classList.toggle('hidden');
      btnToggleRawView.textContent = isHidden ? 'Xem văn bản trích xuất' : 'Ẩn văn bản trích xuất';
    });
  }

  // Stepper số lượng đề
  const inpNumExams = document.getElementById('inpNumExams');
  const btnDecExams = document.getElementById('btnDecExams');
  const btnIncExams = document.getElementById('btnIncExams');
  const inpStartExamCode = document.getElementById('inpStartExamCode');

  function syncExamCodesFromInputs() {
    let count = parseInt(inpNumExams.value, 10);
    if (isNaN(count) || count < 1) count = 1;
    if (count > 24) count = 24;
    inpNumExams.value = count;

    const startCode = inpStartExamCode.value.trim() || '101';
    updateExamCodes(count, startCode);
  }

  if (btnDecExams) {
    btnDecExams.addEventListener('click', () => {
      const current = parseInt(inpNumExams.value, 10) || 1;
      if (current > 1) {
        inpNumExams.value = current - 1;
        syncExamCodesFromInputs();
      }
    });
  }

  if (btnIncExams) {
    btnIncExams.addEventListener('click', () => {
      const current = parseInt(inpNumExams.value, 10) || 1;
      if (current < 24) {
        inpNumExams.value = current + 1;
        syncExamCodesFromInputs();
      }
    });
  }

  if (inpNumExams) {
    inpNumExams.addEventListener('change', syncExamCodesFromInputs);
  }

  if (inpStartExamCode) {
    inpStartExamCode.addEventListener('input', () => {
      const startCode = inpStartExamCode.value.trim() || '101';
      const count = parseInt(inpNumExams.value, 10) || 4;
      updateExamCodes(count, startCode);
    });
  }

  // Cập nhật lại giao diện xem trước khi đổi kiểu nhãn đáp án
  const selDefaultLabelStyle = document.getElementById('selDefaultLabelStyle');
  if (selDefaultLabelStyle) {
    selDefaultLabelStyle.addEventListener('change', () => {
      if (state.parsedExam && state.parsedExam.valid) {
        renderExamPreview(state.parsedExam);
      }
    });
  }

  // Nút Trộn & Xuất Đề Thi
  const btnMixAndExport = document.getElementById('btnMixAndExport');
  if (btnMixAndExport) {
    btnMixAndExport.addEventListener('click', () => {
      executeMixAndExport();
    });
  }

  // Modal đóng
  const exportModal = document.getElementById('exportModal');
  const btnCloseExportModal = document.getElementById('btnCloseExportModal');
  const btnCloseModalBtn = document.getElementById('btnCloseModalBtn');

  [btnCloseExportModal, btnCloseModalBtn].forEach(btn => {
    if (btn) {
      btn.addEventListener('click', () => {
        exportModal.classList.remove('active');
      });
    }
  });

  // Tải ZIP
  const btnDownloadZip = document.getElementById('btnDownloadZip');
  if (btnDownloadZip) {
    btnDownloadZip.addEventListener('click', () => {
      downloadAllZip();
    });
  }

  // Mobile Web Share
  const btnShareFiles = document.getElementById('btnShareFiles');
  if (btnShareFiles) {
    btnShareFiles.addEventListener('click', async () => {
      if (state.generatedZipBlob && navigator.share) {
        try {
          const file = new File([state.generatedZipBlob], 'bo_de_thi_tron.zip', {
            type: 'application/zip'
          });
          await navigator.share({
            title: 'Bộ Đề Thi Trộn',
            text: 'Đã tạo xong các đề thi trắc nghiệm và bảng đáp án',
            files: [file]
          });
        } catch (e) {
          if (e.name !== 'AbortError') {
            showSnackbar('Không thể chia sẻ tệp');
          }
        }
      }
    });
  }

  // iOS Safari Modal
  const isIos = /iPad|iPhone|iPod/.test(navigator.userAgent) && !window.MSStream;
  const isStandalone = window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone;
  if (isIos && !isStandalone) {
    const iosInstallModal = document.getElementById('iosInstallModal');
    const btnCloseIosModal = document.getElementById('btnCloseIosModal');
    const btnGotItIos = document.getElementById('btnGotItIos');

    [btnCloseIosModal, btnGotItIos].forEach(b => {
      if (b) {
        b.addEventListener('click', () => {
          iosInstallModal.classList.remove('active');
        });
      }
    });
  }

  // Image Lightbox xem phóng to và kiểm tra thông số ảnh
  const imageLightboxModal = document.getElementById('imageLightboxModal');
  const btnCloseLightbox = document.getElementById('btnCloseLightbox');
  const lightboxImg = document.getElementById('lightboxImg');
  const lightboxMeta = document.getElementById('lightboxMeta');
  const previewContainer = document.getElementById('previewQuestionsContainer');

  if (previewContainer && imageLightboxModal) {
    previewContainer.addEventListener('click', (e) => {
      const imgTarget = e.target.closest('.exam-preview-img');
      if (imgTarget && lightboxImg) {
        const w = imgTarget.getAttribute('data-width') || imgTarget.naturalWidth;
        const h = imgTarget.getAttribute('data-height') || imgTarget.naturalHeight;
        const natW = imgTarget.naturalWidth;
        const natH = imgTarget.naturalHeight;
        const ratio = (w && h) ? (w / h).toFixed(2) : '1.00';

        lightboxImg.src = imgTarget.src;
        if (lightboxMeta) {
          lightboxMeta.textContent = `Frame hiển thị Word: ${w} × ${h} px • Tỉ lệ: ${ratio} • Độ phân giải ảnh: ${natW} × ${natH} px`;
        }
        imageLightboxModal.classList.add('active');
      }
    });

    if (btnCloseLightbox) {
      btnCloseLightbox.addEventListener('click', () => {
        imageLightboxModal.classList.remove('active');
      });
    }

    imageLightboxModal.addEventListener('click', (e) => {
      if (e.target === imageLightboxModal) {
        imageLightboxModal.classList.remove('active');
      }
    });

    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && imageLightboxModal.classList.contains('active')) {
        imageLightboxModal.classList.remove('active');
      }
    });
  }
}
