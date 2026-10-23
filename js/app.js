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
  initHeaderInfoCacheAndValidation();
  initNetworkStatus();
  initPWA();
  initEventListeners();

  // Khởi tạo 4 mã đề mặc định bắt đầu từ 101
  updateExamCodes(4, '101');

  // Kiểm tra nếu chuyển từ trang Convert sang thì nạp đề vừa convert, ngược lại nạp đề mẫu
  checkAndLoadTransferredExam().then((loaded) => {
    if (!loaded) {
      loadSampleExamDocx(true);
    }
  });
});

/**
 * Quản lý Giao diện (Duy nhất 1 Theme Light chuẩn Microsoft Word)
 */
function initTheme() {
  document.documentElement.setAttribute('data-theme', 'light');
  try {
    localStorage.removeItem('pwa_mark_mix_theme');
  } catch (e) {}

  const metaThemeColor = document.querySelector('meta[name="theme-color"]');
  if (metaThemeColor) {
    metaThemeColor.setAttribute('content', '#185abd');
  }
}

/**
 * Cookie Helpers & Cache Thông Tin Đầu Đề Thi
 * - Cookie lưu 365 ngày (SameSite=Lax, URI-encoded)
 * - Tự động đồng bộ localStorage để hoạt động bền vững ngay cả khi offline PWA
 */
function setCookie(name, value, days = 365) {
  try {
    const expires = new Date(Date.now() + days * 864e5).toUTCString();
    document.cookie = `${encodeURIComponent(name)}=${encodeURIComponent(value)}; expires=${expires}; path=/; SameSite=Lax`;
    localStorage.setItem(`cookie_${name}`, value);
  } catch (e) {
    console.warn('Lỗi ghi cookie/localStorage:', e);
  }
}

function getCookie(name) {
  try {
    const nameEq = encodeURIComponent(name) + '=';
    const ca = document.cookie.split(';');
    for (let c of ca) {
      c = c.trim();
      if (c.indexOf(nameEq) === 0) {
        return decodeURIComponent(c.substring(nameEq.length));
      }
    }
    const fallbackVal = localStorage.getItem(`cookie_${name}`);
    if (fallbackVal !== null && fallbackVal !== undefined) {
      return fallbackVal;
    }
  } catch (e) {
    console.warn('Lỗi đọc cookie/localStorage:', e);
  }
  return '';
}

const HEADER_FIELDS_CONFIG = [
  { id: 'inpSoGD', cookieKey: 'pwa_exam_so_gd', label: 'Sở / Phòng GD&ĐT' },
  { id: 'inpTruong', cookieKey: 'pwa_exam_truong', label: 'Tên Trường / Đơn vị' },
  { id: 'inpKiKiemTra', cookieKey: 'pwa_exam_ki_kiem_tra', label: 'Kì kiểm tra / Đợt thi' },
  { id: 'inpNamHoc', cookieKey: 'pwa_exam_nam_hoc', label: 'Năm học' },
  { id: 'inpMon', cookieKey: 'pwa_exam_mon', label: 'Môn thi' },
  { id: 'inpThoiGian', cookieKey: 'pwa_exam_thoi_gian', label: 'Thời gian làm bài' }
];

/**
 * Khởi tạo tính năng Cache Cookie và Warning cho Thông tin đầu đề thi
 * - Nếu user đã nhập trước đó: Điền sẵn từ cookie/storage và ẩn warning.
 * - Nếu user chưa nhập trước đó: Để trống với hint placeholder và hiện warning.
 * - Lắng nghe sự kiện gõ phím để cache ngay lập tức và tự động cập nhật cảnh báo real-time.
 */
function initHeaderInfoCacheAndValidation() {
  const noticeBanner = document.getElementById('headerWarningNotice');

  // 1. Nạp dữ liệu đã lưu từ cookie/storage nếu có
  HEADER_FIELDS_CONFIG.forEach(field => {
    const input = document.getElementById(field.id);
    if (!input) return;

    const cachedVal = getCookie(field.cookieKey);
    if (cachedVal && cachedVal.trim().length > 0) {
      input.value = cachedVal;
    }
  });

  // 2. Cập nhật giao diện cảnh báo (warning badge trên từng ô & banner chung)
  function updateValidationUI() {
    let emptyCount = 0;

    HEADER_FIELDS_CONFIG.forEach(field => {
      const input = document.getElementById(field.id);
      const warnTag = document.getElementById(`warn_${field.id}`);
      if (!input) return;

      const val = input.value.trim();
      if (val.length === 0) {
        emptyCount++;
        input.classList.add('input-warning');
        if (warnTag) warnTag.classList.remove('hidden');
      } else {
        input.classList.remove('input-warning');
        if (warnTag) warnTag.classList.add('hidden');
      }
    });

    if (noticeBanner) {
      if (emptyCount > 0) {
        noticeBanner.classList.remove('hidden');
      } else {
        noticeBanner.classList.add('hidden');
      }
    }
  }

  // 3. Lắng nghe real-time khi user gõ để tự động lưu cookie và tắt warning
  HEADER_FIELDS_CONFIG.forEach(field => {
    const input = document.getElementById(field.id);
    if (!input) return;

    input.addEventListener('input', () => {
      const val = input.value.trim();
      setCookie(field.cookieKey, val, 365);
      updateValidationUI();
    });

    input.addEventListener('change', () => {
      const val = input.value.trim();
      setCookie(field.cookieKey, val, 365);
      updateValidationUI();
    });
  });

  // Gán hàm cập nhật vào window để các module khác có thể kích hoạt khi cần
  window.updateHeaderValidationUI = updateValidationUI;

  // Chạy kiểm tra ngay khi nạp trang
  updateValidationUI();
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
 * Lấy tệp đã lưu trong IndexedDB khi chuyển trang từ Convert sang Mix
 */
function getBlobFromIndexedDB(key) {
  return new Promise((resolve) => {
    try {
      const request = indexedDB.open('PwaMarkMixDB', 1);
      request.onsuccess = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains('transfers')) {
          return resolve(null);
        }
        const tx = db.transaction('transfers', 'readonly');
        const store = tx.objectStore('transfers');
        const getReq = store.get(key);
        getReq.onsuccess = () => resolve(getReq.result || null);
        getReq.onerror = () => resolve(null);
      };
      request.onerror = () => resolve(null);
    } catch (err) {
      resolve(null);
    }
  });
}

/**
 * Kiểm tra và nạp đề thi vừa chuyển đổi từ Convert Marker
 */
async function checkAndLoadTransferredExam() {
  const urlParams = new URLSearchParams(window.location.search);
  if (urlParams.get('from') === 'convert') {
    try {
      const data = await getBlobFromIndexedDB('pwa_transfer_exam');
      if (data && data.buffer) {
        const file = new File([data.buffer], data.fileName || 'de_thi_da_gan_marker.docx', {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });
        await handleFileLoaded(file);
        showSnackbar('Đã tự động nạp đề thi vừa chuyển đổi từ công cụ Convert Marker!');
        return true;
      }
    } catch (e) {
      console.warn('Không thể nạp tệp chuyển tiếp:', e);
    }
  }
  return false;
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
    file.isSample = true;
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

  // Cập nhật lại giao diện xem trước ngay lập tức khi đổi Định dạng số câu
  const radioNumberings = document.querySelectorAll('input[name="radioNumbering"]');
  radioNumberings.forEach(radio => {
    radio.addEventListener('change', () => {
      if (state.parsedExam && state.parsedExam.valid) {
        renderExamPreview(state.parsedExam);
      }
    });
  });

  // Cập nhật lại giao diện xem trước ngay lập tức khi đổi Kiểu nhãn đáp án
  const selDefaultLabelStyle = document.getElementById('selDefaultLabelStyle');
  if (selDefaultLabelStyle) {
    const updatePreviewOnLabelChange = () => {
      if (state.parsedExam && state.parsedExam.valid) {
        renderExamPreview(state.parsedExam);
      }
    };
    selDefaultLabelStyle.addEventListener('change', updatePreviewOnLabelChange);
    selDefaultLabelStyle.addEventListener('input', updatePreviewOnLabelChange);
  }

  // Cập nhật lại giao diện xem trước ngay lập tức khi đổi Số câu đầu tiên của đề
  const inpStartQNum = document.getElementById('inpStartQNum');
  if (inpStartQNum) {
    const updatePreviewOnStartNumChange = () => {
      if (state.parsedExam && state.parsedExam.valid) {
        renderExamPreview(state.parsedExam);
      }
    };
    inpStartQNum.addEventListener('input', updatePreviewOnStartNumChange);
    inpStartQNum.addEventListener('change', updatePreviewOnStartNumChange);
  }

  // Cập nhật lại giao diện xem trước ngay lập tức khi chỉnh sửa Thông Tin Đầu Đề Thi
  const headerInputs = document.querySelectorAll(".header-cached-input");
  headerInputs.forEach(inp => {
    const updateHeaderPreview = () => {
      if (state.parsedExam && state.parsedExam.valid) {
        renderExamPreview(state.parsedExam);
      }
    };
    inp.addEventListener("input", updateHeaderPreview);
    inp.addEventListener("change", updateHeaderPreview);
  });

  const inpStartCode = document.getElementById("inpStartCode");
  if (inpStartCode) {
    inpStartCode.addEventListener("input", () => {
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

  const closeExportModal = () => {
    if (exportModal) {
      exportModal.classList.remove('active');
    }
  };

  [btnCloseExportModal, btnCloseModalBtn].forEach(btn => {
    if (btn) {
      btn.addEventListener('click', closeExportModal);
    }
  });

  if (exportModal) {
    exportModal.addEventListener('click', (e) => {
      if (e.target === exportModal) {
        closeExportModal();
      }
    });
  }

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
  const iosInstallModal = document.getElementById('iosInstallModal');
  if (isIos && !isStandalone && iosInstallModal) {
    const btnCloseIosModal = document.getElementById('btnCloseIosModal');
    const btnGotItIos = document.getElementById('btnGotItIos');

    const closeIosModal = () => {
      iosInstallModal.classList.remove('active');
    };

    [btnCloseIosModal, btnGotItIos].forEach(b => {
      if (b) {
        b.addEventListener('click', closeIosModal);
      }
    });

    iosInstallModal.addEventListener('click', (e) => {
      if (e.target === iosInstallModal) {
        closeIosModal();
      }
    });
  }

  // Lắng nghe phím Escape để đóng các modal đang mở
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      if (exportModal && exportModal.classList.contains('active')) {
        closeExportModal();
      }
      if (iosInstallModal && iosInstallModal.classList.contains('active')) {
        iosInstallModal.classList.remove('active');
      }
    }
  });

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
