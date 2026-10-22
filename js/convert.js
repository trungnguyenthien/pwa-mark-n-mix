/**
 * Controller Giao Diện Cho Trang Chuyển Đổi Marker (convert.html)
 * - Quản lý nạp file Word truyền thống (drag & drop / file picker)
 * - Tự động trích xuất, phân tích và chuyển đổi sang chuẩn marker
 * - Hiển thị bản xem trước trực quan với huy hiệu phân loại màu sắc
 * - Cảnh báo các câu hỏi thiếu đáp án gạch chân hoặc ảnh không phải inline
 * - Xuất và tải về file Word (.docx) đã gắn marker chuẩn Times New Roman
 */

import { extractTextFromDocx } from './docx-importer.js';
import { convertTextToMarkers, generateConvertedDocxBlob } from './marker-converter.js';

// Trạng thái cục bộ của trang chuyển đổi
const convertState = {
  file: null,
  fileName: '',
  rawText: '',
  convertedText: '',
  imageMap: {},
  nonInlineImages: [],
  stats: null,
  activeTab: 'visual' // 'visual' | 'code'
};

// Khởi chạy khi DOM sẵn sàng
document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  initNetworkStatus();
  initEventListeners();
});

/**
 * Đồng bộ theme Word Light
 */
function initTheme() {
  document.documentElement.setAttribute('data-theme', 'light');
  const metaThemeColor = document.querySelector('meta[name="theme-color"]');
  if (metaThemeColor) {
    metaThemeColor.setAttribute('content', '#185abd');
  }
}

/**
 * Trạng thái kết nối mạng (Online/Offline)
 */
function initNetworkStatus() {
  const netStatus = document.getElementById('netStatus');
  if (!netStatus) return;

  function update() {
    const isOnline = navigator.onLine;
    netStatus.className = `net-status-pill ${isOnline ? 'online' : 'offline'}`;
    const dot = netStatus.querySelector('.status-dot');
    const txt = netStatus.querySelector('.status-text');
    if (dot) dot.className = `status-dot ${isOnline ? '' : 'offline'}`;
    if (txt) txt.textContent = isOnline ? 'Online' : 'Offline';
  }

  window.addEventListener('online', update);
  window.addEventListener('offline', update);
  update();
}

/**
 * Hiển thị thông báo Toast / Snackbar
 */
function showToast(message, type = 'info', duration = 3500) {
  const snackbar = document.getElementById('snackbar');
  if (!snackbar) return;

  snackbar.textContent = message;
  snackbar.className = `snackbar show ${type}`;

  setTimeout(() => {
    snackbar.className = snackbar.className.replace('show', '').trim();
  }, duration);
}

/**
 * Gán toàn bộ sự kiện trên trang Convert
 */
function initEventListeners() {
  const dropzone = document.getElementById('dropzoneConvert');
  const fileInput = document.getElementById('fileInputConvert');
  const btnSelectFile = document.getElementById('btnSelectFile');
  const btnLoadSampleTrad = document.getElementById('btnLoadSampleTrad');
  const btnDownloadConvertedDocx = document.getElementById('btnDownloadConvertedDocx');
  const btnOpenInMixer = document.getElementById('btnOpenInMixer');
  const btnCopyMarkerText = document.getElementById('btnCopyMarkerText');
  const btnTabVisual = document.getElementById('btnTabVisual');
  const btnTabCode = document.getElementById('btnTabCode');
  const btnCloseNonInlineModal = document.getElementById('btnCloseNonInlineModal');
  const btnDismissNonInline = document.getElementById('btnDismissNonInline');
  const modalNonInline = document.getElementById('nonInlineModal');

  // Xử lý nạp file qua file input
  if (fileInput) {
    fileInput.addEventListener('change', () => {
      if (fileInput.files && fileInput.files.length > 0) {
        processUploadedFile(fileInput.files[0]);
        fileInput.value = '';
      }
    });
  }

  if (btnSelectFile && fileInput) {
    btnSelectFile.addEventListener('click', (e) => {
      e.stopPropagation();
      fileInput.click();
    });
  }

  // Xử lý Drag & Drop
  if (dropzone) {
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
        processUploadedFile(files[0]);
      }
    });

    dropzone.addEventListener('click', () => {
      if (fileInput) fileInput.click();
    });

    dropzone.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        if (fileInput) fileInput.click();
      }
    });
  }

  // Nạp đề mẫu truyền thống để thử nghiệm nhanh
  if (btnLoadSampleTrad) {
    btnLoadSampleTrad.addEventListener('click', (e) => {
      e.stopPropagation();
      loadTraditionalSample();
    });
  }

  // Nút Tải file Word đã gắn marker
  if (btnDownloadConvertedDocx) {
    btnDownloadConvertedDocx.addEventListener('click', handleDownloadDocx);
  }

  // Nút Chuyển sang trang Trộn đề
  if (btnOpenInMixer) {
    btnOpenInMixer.addEventListener('click', handleOpenInMixer);
  }

  // Nút Sao chép mã Marker
  if (btnCopyMarkerText) {
    btnCopyMarkerText.addEventListener('click', () => {
      const rawTextarea = document.getElementById('rawMarkerTextarea');
      if (rawTextarea && rawTextarea.value) {
        navigator.clipboard.writeText(rawTextarea.value).then(() => {
          showToast('Đã sao chép toàn bộ mã marker vào bộ nhớ đệm!', 'success');
        }).catch(() => {
          rawTextarea.select();
          document.execCommand('copy');
          showToast('Đã chọn và sao chép mã marker!', 'success');
        });
      }
    });
  }

  // Chuyển Tab xem trước
  if (btnTabVisual && btnTabCode) {
    btnTabVisual.addEventListener('click', () => switchTab('visual'));
    btnTabCode.addEventListener('click', () => switchTab('code'));
  }

  // Đóng Modal cảnh báo ảnh không phải InLine
  if (btnCloseNonInlineModal && modalNonInline) {
    btnCloseNonInlineModal.addEventListener('click', () => {
      modalNonInline.classList.remove('active');
    });
  }
  if (btnDismissNonInline && modalNonInline) {
    btnDismissNonInline.addEventListener('click', () => {
      modalNonInline.classList.remove('active');
    });
  }
  if (modalNonInline) {
    modalNonInline.addEventListener('click', (e) => {
      if (e.target === modalNonInline) {
        modalNonInline.classList.remove('active');
      }
    });
    window.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && modalNonInline.classList.contains('active')) {
        modalNonInline.classList.remove('active');
      }
    });
  }
}

/**
 * Chuyển đổi tab xem trước
 */
function switchTab(tab) {
  convertState.activeTab = tab;
  const btnTabVisual = document.getElementById('btnTabVisual');
  const btnTabCode = document.getElementById('btnTabCode');
  const visualContainer = document.getElementById('visualPreviewContainer');
  const codeContainer = document.getElementById('codePreviewContainer');

  if (tab === 'visual') {
    if (btnTabVisual) btnTabVisual.classList.add('active');
    if (btnTabCode) btnTabCode.classList.remove('active');
    if (visualContainer) visualContainer.classList.remove('hidden');
    if (codeContainer) codeContainer.classList.add('hidden');
  } else {
    if (btnTabCode) btnTabCode.classList.add('active');
    if (btnTabVisual) btnTabVisual.classList.remove('active');
    if (codeContainer) codeContainer.classList.remove('hidden');
    if (visualContainer) visualContainer.classList.add('hidden');
  }
}

/**
 * Xử lý tệp DOCX người dùng tải lên
 */
async function processUploadedFile(file) {
  if (!file) return;

  const fileName = file.name || 'de_thi.docx';
  if (!fileName.toLowerCase().endsWith('.docx')) {
    showToast('Vui lòng chọn tệp Word (.docx) hợp lệ!', 'danger');
    return;
  }

  setLoadingState(true, `Đang phân tích tệp "${fileName}"...`);

  try {
    // 1. Trích xuất text và ảnh từ file Word
    const extracted = await extractTextFromDocx(file);

    convertState.file = file;
    convertState.fileName = fileName;
    convertState.rawText = extracted.rawText;
    convertState.imageMap = extracted.imageMap || {};
    convertState.nonInlineImages = extracted.nonInlineImages || [];

    // 2. Chuyển đổi sang hệ thống marker
    const conversion = convertTextToMarkers(extracted.rawText);
    convertState.convertedText = conversion.convertedText;
    convertState.stats = conversion.stats;

    // 3. Hiển thị cảnh báo ảnh không phải InLine nếu có
    if (convertState.nonInlineImages.length > 0) {
      showNonInlineWarningModal(convertState.nonInlineImages);
    }

    // 4. Cập nhật giao diện kết quả
    renderConversionResults();
    showToast(`Đã chuyển đổi thành công ${conversion.stats.totalQuestions} câu hỏi!`, 'success');

  } catch (err) {
    console.error('Lỗi chuyển đổi:', err);
    showToast(`Lỗi xử lý file Word: ${err.message}`, 'danger');
  } finally {
    setLoadingState(false);
  }
}

/**
 * Nạp đề mẫu truyền thống để trải nghiệm ngay
 */
function loadTraditionalSample() {
  const sampleTraditionalText = `SỞ GIÁO DỤC VÀ ĐÀO TẠO TP HỒ CHÍ MINH
TRƯỜNG THPT CHUYÊN LÊ HỒNG PHONG
ĐỀ KIỂM TRA ĐỊNH KỲ NĂM HỌC 2026 - 2027
MÔN: VẬT LÝ 10 - THỜI GIAN: 45 PHÚT

PHẦN I. CÂU TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN
Thí sinh trả lời từ câu 1 đến câu 4. Mỗi câu hỏi chỉ chọn một phương án đúng nhất.

Câu 1. Trong cuộc Cách mạng công nghiệp 4.0, sự phát triển vượt bậc của viễn thông số, internet cáp quang và trí tuệ nhân tạo (AI) đều bắt nguồn từ nghiên cứu nền tảng của Vật lí. Đối tượng nghiên cứu chủ yếu của Vật lí học là
A. sự biến đổi chất này thành chất khác và cơ chế phản ứng hóa học.
<u>B.</u> các dạng vận động của vật chất và năng lượng trong tự nhiên.
C. quy luật tiến hóa, di truyền và sinh trưởng của các loài sinh vật.
D. cấu tạo địa chất và sự dịch chuyển của các tầng mảng kiến tạo Trái Đất.

Câu 2: Một xe máy chạy trên một đường thẳng thẳng đều với tốc độ 36 km/h. Đổi tốc độ này ra đơn vị mét trên giây (m/s) là:
<u>A. 10 m/s</u>       B. 15 m/s       C. 20 m/s       D. 25 m/s

Câu 3) Đại lượng đặc trưng cho sự biến thiên nhanh hay chậm của vận tốc theo thời gian là:
a) Quãng đường đi được.
b) Độ dịch chuyển tổng cộng.
<u>c)</u> Gia tốc của chuyển động.
d) Tốc độ tức thời.

4) Một vật chuyển động thẳng biến đổi đều từ trạng thái nghỉ với gia tốc a = 2 m/s^2. Vận tốc của vật sau 5 giây là:
A) 5 m/s       B) 8 m/s
<u>C) 10 m/s</u>       D) 20 m/s
`;

  setLoadingState(true, 'Đang nạp dữ liệu đề mẫu truyền thống...');
  setTimeout(() => {
    convertState.file = null;
    convertState.fileName = 'de_thi_truyen_thong_mau.docx';
    convertState.rawText = sampleTraditionalText;
    convertState.imageMap = {};
    convertState.nonInlineImages = [];

    const conversion = convertTextToMarkers(sampleTraditionalText);
    convertState.convertedText = conversion.convertedText;
    convertState.stats = conversion.stats;

    renderConversionResults();
    setLoadingState(false);
    showToast('Đã nạp đề mẫu truyền thống và chuyển đổi xong!', 'success');
  }, 300);
}

/**
 * Hiển thị kết quả chuyển đổi lên giao diện
 */
function renderConversionResults() {
  const resultSection = document.getElementById('resultSection');
  const emptyState = document.getElementById('emptyState');
  if (resultSection) resultSection.classList.remove('hidden');
  if (emptyState) emptyState.classList.add('hidden');

  const stats = convertState.stats || { totalQuestions: 0, totalD: 0, totalS: 0, questionsWithoutKey: [], warnings: [] };

  // 1. Cập nhật các con số thống kê
  const statQuestions = document.getElementById('statQuestions');
  const statD = document.getElementById('statD');
  const statS = document.getElementById('statS');
  const statWarningsCount = document.getElementById('statWarningsCount');

  if (statQuestions) statQuestions.textContent = stats.totalQuestions;
  if (statD) statD.textContent = stats.totalD;
  if (statS) statS.textContent = stats.totalS;
  if (statWarningsCount) statWarningsCount.textContent = (stats.questionsWithoutKey.length) + (convertState.nonInlineImages.length);

  // 2. Hiển thị cảnh báo thiếu đáp án gạch chân nếu có
  const warningContainer = document.getElementById('warningContainer');
  if (warningContainer) {
    if (stats.warnings.length > 0 || convertState.nonInlineImages.length > 0) {
      warningContainer.classList.remove('hidden');
      let html = '<div class="alert-box alert-warning">';
      html += '<div class="alert-icon">⚠️</div><div class="alert-content">';
      html += '<h4 class="alert-title">Lưu ý kiểm tra đề thi:</h4><ul class="alert-list">';

      for (const w of stats.warnings) {
        html += `<li>${escapeHtml(w)}</li>`;
      }
      if (convertState.nonInlineImages.length > 0) {
        html += `<li>Phát hiện <strong>${convertState.nonInlineImages.length}</strong> hình ảnh không phải In Line with Text (ảnh nổi). Hãy chỉnh ảnh thành In Line để hiển thị chuẩn xác nhất.</li>`;
      }
      html += '</ul></div></div>';
      warningContainer.innerHTML = html;
    } else {
      warningContainer.classList.add('hidden');
      warningContainer.innerHTML = '';
    }
  }

  // 3. Render bản xem trước trực quan (Visual Preview)
  renderVisualPreview(convertState.convertedText, convertState.imageMap);

  // 4. Render mã marker vào textarea
  const rawMarkerTextarea = document.getElementById('rawMarkerTextarea');
  if (rawMarkerTextarea) {
    rawMarkerTextarea.value = convertState.convertedText;
  }

  // 5. Cập nhật thông tin file nguồn
  const sourceFileName = document.getElementById('sourceFileName');
  if (sourceFileName) {
    sourceFileName.textContent = convertState.fileName;
  }

  // Cuộn nhẹ tới khu vực kết quả
  if (resultSection) {
    resultSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

/**
 * Render bản xem trước trực quan với huy hiệu màu sắc nổi bật
 */
function renderVisualPreview(convertedText, imageMap = {}) {
  const container = document.getElementById('visualPreviewContainer');
  if (!container) return;

  const lines = convertedText.split(/\r?\n/);
  let html = '<div class="doc-sheet">';

  let qIndex = 0;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) {
      html += '<div class="doc-line-spacer"></div>';
      continue;
    }

    if (line.startsWith('__G__') || /^__G\w*__/.test(line)) {
      html += `<div class="doc-marker-pill group-pill" title="Phân nhóm đề thi"><span class="pill-tag">${escapeHtml(line)}</span> <span class="pill-desc">Bắt đầu nhóm/phần mới</span></div>`;
      continue;
    }

    if (line === '__QUESTION__' || /^__QUESTION__\s*/.test(line)) {
      qIndex++;
      const rest = line.replace(/^__QUESTION__\s*/, '').trim();
      const renderedRest = rest ? formatContentForHtml(rest, imageMap) : '';
      html += `<div class="doc-marker-pill question-pill" title="Câu hỏi"><span class="pill-tag">__QUESTION__</span> <span class="pill-desc">Câu ${qIndex}</span></div>`;
      if (renderedRest) {
        html += `<div class="doc-content-line doc-question-text" style="font-weight: 500; margin-bottom: 6px;">${renderedRest}</div>`;
      }
      continue;
    }

    if (/^__(?:D|D#|#D)__\s*/.test(line)) {
      const markerMatch = line.match(/^__(?:D|D#|#D)__\s*/)[0];
      const rest = line.substring(markerMatch.length);
      const renderedContent = formatContentForHtml(rest, imageMap);
      html += `<div class="doc-answer-item correct-ans"><span class="ans-marker-badge correct" title="Đáp án đúng (có gạch chân ở đề gốc)">__D__ ✓</span> <span class="ans-text">${renderedContent}</span></div>`;
      continue;
    }

    if (/^__(?:S|S#|#S)__\s*/.test(line)) {
      const markerMatch = line.match(/^__(?:S|S#|#S)__\s*/)[0];
      const rest = line.substring(markerMatch.length);
      const renderedContent = formatContentForHtml(rest, imageMap);
      html += `<div class="doc-answer-item incorrect-ans"><span class="ans-marker-badge incorrect" title="Đáp án sai">__S__</span> <span class="ans-text">${renderedContent}</span></div>`;
      continue;
    }

    // Dòng văn bản bình thường (nội dung câu hỏi, tiêu đề nhóm, ảnh...)
    const rendered = formatContentForHtml(line, imageMap);
    html += `<div class="doc-content-line">${rendered}</div>`;
  }

  html += '</div>';
  container.innerHTML = html;
}

/**
 * Định dạng các thẻ HTML <b>, <i>, <u> và thay token [IMG:id] thành ảnh HTML
 */
function formatContentForHtml(text, imageMap = {}) {
  if (!text) return '';

  // Thay thế [IMG:id]
  let res = text.replace(/\[IMG:([^\]]+)\]/gi, (match, id) => {
    const img = imageMap[id];
    if (img && img.dataUrl) {
      return `<img src="${img.dataUrl}" alt="Hình ảnh đề thi" class="preview-inline-img" style="max-width: 100%; height: auto; max-height: 250px; display: inline-block; vertical-align: middle; margin: 4px 6px; border-radius: 4px;" />`;
    }
    return `<span class="img-missing-badge">[Hình ảnh: ${escapeHtml(id)}]</span>`;
  });

  return res;
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * Hiển thị / Ẩn trạng thái đang tải (Loading overlay)
 */
function setLoadingState(isLoading, message = 'Đang xử lý...') {
  const overlay = document.getElementById('loadingOverlay');
  const msgEl = document.getElementById('loadingMessage');
  if (overlay) {
    if (isLoading) {
      if (msgEl) msgEl.textContent = message;
      overlay.classList.add('active');
    } else {
      overlay.classList.remove('active');
    }
  }
}

/**
 * Tải file Word đã gắn marker về máy tính (Yêu cầu cốt lõi của người dùng)
 */
async function handleDownloadDocx() {
  if (!convertState.convertedText) {
    showToast('Chưa có nội dung đề thi đã chuyển đổi!', 'warning');
    return;
  }

  setLoadingState(true, 'Đang tạo tệp Word (.docx) chuẩn Times New Roman...');

  try {
    const blob = await generateConvertedDocxBlob(convertState.convertedText, convertState.imageMap);

    // Đặt tên file xuất: [TênGốc]_da_gan_marker.docx
    let baseName = convertState.fileName.replace(/\.docx$/i, '');
    if (!baseName) baseName = 'de_thi';
    const downloadName = `${baseName}_da_gan_marker.docx`;

    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = downloadName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 2000);

    showToast(`Đã tải xuống thành công tệp "${downloadName}"!`, 'success');
  } catch (err) {
    console.error('Lỗi xuất file Word:', err);
    showToast(`Không thể tạo file Word: ${err.message}`, 'danger');
  } finally {
    setLoadingState(false);
  }
}

/**
 * Chuyển nội dung sang trang Trộn đề (mix.html)
 */
async function handleOpenInMixer() {
  if (!convertState.convertedText) {
    showToast('Chưa có nội dung đề thi đã chuyển đổi!', 'warning');
    return;
  }

  setLoadingState(true, 'Đang chuẩn bị dữ liệu chuyển sang trang Trộn đề...');

  try {
    const blob = await generateConvertedDocxBlob(convertState.convertedText, convertState.imageMap);
    const baseName = convertState.fileName.replace(/\.docx$/i, '') || 'de_thi';
    const finalName = `${baseName}_da_gan_marker.docx`;

    // Lưu mảng buffer vào IndexedDB hoặc sessionStorage để chuyển qua mix.html
    const arrayBuffer = await blob.arrayBuffer();

    // Dùng IndexedDB để lưu buffer không bị giới hạn 5MB của sessionStorage
    await saveBlobToIndexedDB('pwa_transfer_exam', arrayBuffer, finalName);

    window.location.href = './mix.html?from=convert';
  } catch (err) {
    console.error('Lỗi chuyển trang:', err);
    showToast('Chuyển trang thất bại. Bạn có thể tải file Word về và nạp vào trang Trộn đề.', 'warning');
    setLoadingState(false);
  }
}

/**
 * Lưu tệp buffer vào IndexedDB cho việc truyền giữa các trang
 */
function saveBlobToIndexedDB(key, buffer, fileName) {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open('PwaMarkMixDB', 1);
    request.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains('transfers')) {
        db.createObjectStore('transfers');
      }
    };
    request.onsuccess = (e) => {
      const db = e.target.result;
      const tx = db.transaction('transfers', 'readwrite');
      const store = tx.objectStore('transfers');
      store.put({ buffer, fileName, time: Date.now() }, key);
      tx.oncomplete = () => resolve(true);
      tx.onerror = () => reject(tx.error);
    };
    request.onerror = () => reject(request.error);
  });
}

/**
 * Hiển thị modal cảnh báo ảnh không phải InLine
 */
function showNonInlineWarningModal(images) {
  const modal = document.getElementById('nonInlineModal');
  const countEl = document.getElementById('nonInlineCount');
  const listEl = document.getElementById('nonInlineList');
  if (!modal || !listEl) return;

  if (countEl) countEl.textContent = `${images.length} hình ảnh`;

  let html = '';
  images.forEach((img, idx) => {
    html += `
      <div class="non-inline-item">
        <div class="non-inline-num">${idx + 1}</div>
        <div class="non-inline-thumb">
          ${img.dataUrl ? `<img src="${img.dataUrl}" alt="${escapeHtml(img.name)}" />` : '<div class="no-thumb">Ảnh Word</div>'}
        </div>
        <div class="non-inline-info">
          <div class="non-inline-name">${escapeHtml(img.name)}</div>
          <div class="non-inline-meta">Kiểu bọc chữ hiện tại: <strong>${escapeHtml(img.wrapType)}</strong></div>
          <div class="non-inline-snippet">Vị trí đoạn văn: <em>"${escapeHtml(img.snippet)}"</em></div>
        </div>
      </div>
    `;
  });

  listEl.innerHTML = html;
  modal.classList.add('active');
}
