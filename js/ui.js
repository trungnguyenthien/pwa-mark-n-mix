/**
 * Module Giao Diện Người Dùng (UI Controller)
 * Quản lý tương tác nạp file Word (.docx), cấu hình đầu đề, tùy chọn trộn,
 * hiển thị xem trước trực quan và xử lý xuất file DOCX / ZIP.
 */

import { readExamDocx } from './docx-importer.js';
import { parseExam } from './parser.js';
import { shuffleAllExams, formatGroupTitle } from './shuffler.js';
import { createExamDocxBlob } from './docx-builder.js';
import { createAnswerKeyDocxBlob, buildAnswerMatrix } from './answer-key.js';

// Trạng thái ứng dụng
export const state = {
  rawText: '',
  fileName: '',
  fileSize: 0,
  imageMap: {},
  parsedExam: null,
  examCodes: ['101', '102', '103', '104'],
  generatedZipBlob: null,
  generatedFiles: [] // { name, blob, url }
};

// Hàm hiển thị thông báo Snackbar
export function showSnackbar(message, duration = 3000) {
  const snackbar = document.getElementById('snackbar');
  if (!snackbar) return;
  snackbar.textContent = message;
  snackbar.classList.add('show');
  setTimeout(() => {
    snackbar.classList.remove('show');
  }, duration);
}

// Format dung lượng tệp
function formatBytes(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

/**
 * Cập nhật danh sách mã đề hiển thị theo số lượng đề hoặc mã bắt đầu
 */
export function updateExamCodes(count, startCode) {
  const numericStart = parseInt(startCode, 10) || 101;
  const newCodes = [];

  for (let i = 0; i < count; i++) {
    // Nếu người dùng đã tự sửa mã ở vị trí i thì ưu tiên giữ lại, nếu không thì tăng dần
    if (state.examCodes[i] !== undefined) {
      newCodes.push(state.examCodes[i]);
    } else {
      newCodes.push(String(numericStart + i));
    }
  }

  state.examCodes = newCodes;
  renderExamCodesChips();
}

/**
 * Hiển thị danh sách chip mã đề thi tương tác
 */
function renderExamCodesChips() {
  const container = document.getElementById('examCodesContainer');
  if (!container) return;

  container.innerHTML = '';
  state.examCodes.forEach((code, index) => {
    const chip = document.createElement('div');
    chip.className = 'code-chip';
    chip.title = `Bấm để đổi mã đề số ${index + 1}`;
    chip.innerHTML = `<span>Mã:</span> <strong>${code}</strong>`;

    chip.addEventListener('click', () => {
      // Chuyển sang ô input để chỉnh sửa trực tiếp
      const input = document.createElement('input');
      input.type = 'text';
      input.className = 'code-chip-input';
      input.value = code;
      input.maxLength = 6;

      chip.replaceWith(input);
      input.focus();
      input.select();

      const commitChange = () => {
        const val = input.value.trim() || code;
        state.examCodes[index] = val;
        renderExamCodesChips();
      };

      input.addEventListener('blur', commitChange);
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          commitChange();
        } else if (e.key === 'Escape') {
          renderExamCodesChips();
        }
      });
    });

    container.appendChild(chip);
  });
}

/**
 * Tạo HTML xem trước khung tiêu đề bài thi (Header)
 */
export function renderExamHeaderPreview() {
  const header = getHeaderConfig();
  const inpSo = typeof document !== "undefined" ? document.getElementById("inpSoGD") : null;
  const inpTr = typeof document !== "undefined" ? document.getElementById("inpTruong") : null;
  const inpKi = typeof document !== "undefined" ? document.getElementById("inpKiKiemTra") : null;
  const inpNam = typeof document !== "undefined" ? document.getElementById("inpNamHoc") : null;
  const inpMon = typeof document !== "undefined" ? document.getElementById("inpMon") : null;
  const inpThoi = typeof document !== "undefined" ? document.getElementById("inpThoiGian") : null;
  const inpStartCode = typeof document !== "undefined" ? document.getElementById("inpStartCode") : null;

  const soGD = (header.soGD || (inpSo ? inpSo.placeholder : "") || "SỞ GIÁO DỤC VÀ ĐÀO TẠO").trim().toUpperCase();
  const truong = (header.truong || (inpTr ? inpTr.placeholder : "") || "TRƯỜNG THPT NGUYỄN HỮU HUÂN").trim().toUpperCase();
  const kiKiemTra = (header.kiKiemTra || (inpKi ? inpKi.placeholder : "") || "KIỂM TRA HỌC KỲ").trim().toUpperCase();
  const namHoc = (header.namHoc || (inpNam ? inpNam.placeholder : "") || "2026 - 2027").trim();
  const mon = (header.mon || (inpMon ? inpMon.placeholder : "") || "VẬT LÝ").trim().toUpperCase();
  const thoiGian = (header.thoiGian || (inpThoi ? inpThoi.placeholder : "") || "45 phút").trim();
  const previewExamCode = (inpStartCode && inpStartCode.value.trim()) || "101";

  return `
    <div class="exam-header-preview-card" title="Tiêu đề bài thi - Tự động đồng bộ với Thông Tin Đầu Đề Thi ở Bước 2">
      <div class="exam-header-preview-content">
        <div class="exam-header-col col-left">
          <div class="exam-header-dept">${escapeHtml(soGD)}</div>
          <div class="exam-header-school">${escapeHtml(truong)}</div>
          <div class="exam-header-dash">--------------------</div>
          <div class="exam-header-pages"><em>(Đề thi có ... trang)</em></div>
        </div>
        <div class="exam-header-col col-right">
          <div class="exam-header-session">${escapeHtml(kiKiemTra)}</div>
          <div class="exam-header-year"><strong>NĂM HỌC ${escapeHtml(namHoc)}</strong></div>
          <div class="exam-header-subject"><strong>MÔN: ${escapeHtml(mon)}</strong></div>
          <div class="exam-header-time"><em>Thời gian làm bài: <strong>${escapeHtml(thoiGian)}</strong></em></div>
          <div class="exam-header-timenote"><em>(không kể thời gian phát đề)</em></div>
        </div>
      </div>
      <div class="exam-header-candidate-line">
        <div class="candidate-meta">
          <span>Họ và tên: ................................................................</span>
          <span>Số báo danh: .......................</span>
        </div>
        <div class="candidate-code">
          <strong>Mã đề ${escapeHtml(previewExamCode)}</strong>
        </div>
      </div>
    </div>
  `;
}

/**
 * Hiển thị xem trước cấu trúc đề thi
 */
export function renderExamPreview(parsedExam) {
  const container = document.getElementById('previewQuestionsContainer');
  if (!container) return;

  if (!parsedExam || parsedExam.groups.length === 0) {
    container.innerHTML = `
      <div class="empty-state">
        <p class="empty-title">Chưa nạp đề thi</p>
        <p class="empty-desc">Kéo thả file Word (.docx) ở Bước 1 để phân tích và xem trước đề thi.</p>
      </div>
    `;
    return;
  }

  let html = renderExamHeaderPreview();
  let globalQuestionCounter = 0;
  const imageMap = parsedExam.imageMap || state.imageMap || {};

  function renderFormattedContent(text) {
    if (!text) return '';
    let formatted = escapeHtml(text);
    // Phục hồi an toàn các thẻ định dạng <b>, <i>, <u> để xem trước trực quan
    formatted = formatted
      .replace(/&lt;b&gt;/gi, '<strong class="q-bold">')
      .replace(/&lt;\/b&gt;/gi, '</strong>')
      .replace(/&lt;i&gt;/gi, '<em class="q-italic">')
      .replace(/&lt;\/i&gt;/gi, '</em>')
      .replace(/&lt;u&gt;/gi, '<span class="q-underline">')
      .replace(/&lt;\/u&gt;/gi, '</span>');

    // Giữ các dấu xuống hàng tự nhiên
    const withBreaks = formatted.replace(/\n/g, '<br/>');
    return withBreaks.replace(/\[IMG:([^\]]+)\]/g, (match, id) => {
      const img = imageMap[id];
      if (img && img.dataUrl) {
        const w = img.width || 180;
        const h = img.height || 100;
        const ratio = (w / h).toFixed(2);
        return `<span class="exam-img-frame" title="Hình ảnh từ file Word gốc: ${w} × ${h} px (Tỉ lệ: ${ratio})">` +
          `<img src="${img.dataUrl}" alt="Hình ảnh đề thi" class="exam-preview-img" ` +
          `style="width: ${w}px; max-width: 100%; aspect-ratio: ${w} / ${h};" ` +
          `data-width="${w}" data-height="${h}" data-img-id="${id}" loading="lazy" />` +
          `</span>`;
      }
      return '';
    });
  }

  const labelSelect = typeof document !== 'undefined' ? document.getElementById('selDefaultLabelStyle') : null;
  const currentLabelStyle = labelSelect ? labelSelect.value : 'ABCD';

  const radioNumbering = typeof document !== 'undefined' ? document.querySelector('input[name="radioNumbering"]:checked') : null;
  const currentNumberingFormat = radioNumbering ? radioNumbering.value : 'dot';

  const startNumInput = typeof document !== 'undefined' ? document.getElementById('inpStartQNum') : null;
  const startQNum = (startNumInput && parseInt(startNumInput.value, 10)) || 1;

  const chkPerGroup = typeof document !== 'undefined' ? document.getElementById('chkNumberingPerGroup') : null;
  const numberingScope = (chkPerGroup && chkPerGroup.checked) ? 'group' : 'all';

  const chkInsertBlank = typeof document !== 'undefined' ? document.getElementById('chkInsertBlankLine') : null;
  const insertBlankLine = chkInsertBlank ? chkInsertBlank.checked : false;

  parsedExam.groups.forEach((group, gIdx) => {
    const modeBadge = group.mode || 'G';
    const isFixedGroup = group.fixed;
    const modeDesc = {
      'G': 'Mix cả câu hỏi & đáp án',
      'GQ': 'Chỉ mix câu hỏi',
      'GA': 'Chỉ mix đáp án',
      'G0': 'Không đảo vị trí'
    }[modeBadge] || 'Mix đề';

    html += `
      <div class="preview-group-box">
        <div class="preview-group-header">
          <div style="display: inline-flex; align-items: center; gap: 8px;">
            <span class="group-mode-pill">${group.rawMarker} ${modeDesc}</span>
            ${isFixedGroup ? '<span class="badge badge-success">Vị trí cố định (#)</span>' : ''}
          </div>
          <div style="font-size: 0.8125rem; color: var(--text-muted); font-weight: 500;">
            ${group.questions.length} câu hỏi
          </div>
        </div>
        ${group.title && group.title.trim() ? `
          <div class="preview-group-title" style="padding: 10px 16px; background: #f8fafc; border-bottom: 1px solid var(--border-light); font-size: 0.9375rem; color: #1e293b; line-height: 1.5; white-space: pre-line;">
            ${renderFormattedContent(formatGroupTitle(group.title.trim(), gIdx + 1))}
          </div>
        ` : ''}
        <div class="preview-questions-list">
    `;

    let groupQuestionCounter = 0;
    group.questions.forEach((q, qIdxInGroup) => {
      globalQuestionCounter++;
      groupQuestionCounter++;
      const currentQNumber = (numberingScope === 'group')
        ? (startQNum + groupQuestionCounter - 1)
        : (startQNum + globalQuestionCounter - 1);
      let qNumDisplay = `Câu ${currentQNumber}.`;
      if (currentNumberingFormat === 'colon') {
        qNumDisplay = `Câu ${currentQNumber}:`;
      } else if (currentNumberingFormat === 'paren') {
        qNumDisplay = `Câu ${currentQNumber})`;
      }

      html += `
        <div class="question-card">
          <div class="question-title-line">
            <span class="question-qnum">${qNumDisplay}</span>
            <span>${renderFormattedContent(q.text)}</span>
          </div>
          <div class="answers-grid">
      `;

      if (q.type === 'short_answer') {
        html += `
          <div class="answer-item">
            <span class="ans-pill d">Trả lời ngắn</span>
            <span class="answer-text"><strong>Đáp án:</strong> ${renderFormattedContent(q.answers[0].text)}</span>
          </div>
        `;
      } else {
        q.answers.forEach((ans, aIdx) => {
          const isD = ans.correct;
          const isFixed = ans.fixed;

          let labelDisplay = `${String.fromCharCode(65 + (aIdx % 26))}.`;
          if (currentLabelStyle === 'ABCD_paren') {
            labelDisplay = `${String.fromCharCode(65 + (aIdx % 26))})`;
          } else if (currentLabelStyle === 'abcd') {
            labelDisplay = `${String.fromCharCode(97 + (aIdx % 26))})`;
          } else if (currentLabelStyle === '1234') {
            labelDisplay = `${aIdx + 1}.`;
          }

          html += `
            <div class="answer-item">
              <span class="ans-pill ${isD ? 'd' : 's'}" title="${isFixed ? (isD ? 'Đúng (Cố định vị trí #)' : 'Sai (Cố định vị trí #)') : (isD ? 'Đáp án Đúng' : 'Đáp án Sai')}">
                ${isD ? (isFixed ? 'D#' : 'D_') : (isFixed ? 'S#' : 'S_')}
              </span>
              <span class="answer-text"><strong>${labelDisplay}</strong> ${renderFormattedContent(ans.text)}</span>
            </div>
          `;
        });
      }

      html += `
          </div>
        </div>
      `;
      const isLastQuestionOverall = (gIdx === parsedExam.groups.length - 1 && qIdxInGroup === group.questions.length - 1);
      if (insertBlankLine && !isLastQuestionOverall) {
        html += '<div class="preview-blank-line" style="height: 16px;" title="Dòng trống giữa các câu"></div>';
      }
    });

    html += `
        </div>
      </div>
    `;
  });

  container.innerHTML = html;
}

function escapeHtml(str) {
  if (!str) return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * Xử lý khi nạp file Word thành công
 */
export async function handleFileLoaded(file) {
  try {
    showSnackbar('Đang đọc và phân tích file Word...');
    const result = await readExamDocx(file);

    state.rawText = result.text;
    state.fileName = result.fileName;
    state.fileSize = result.fileSize;
    state.imageMap = result.imageMap || {};

    // Cập nhật giao diện thông tin file
    document.getElementById('txtFileName').textContent = result.fileName;
    document.getElementById('txtFileSize').textContent = formatBytes(result.fileSize);
    document.getElementById('loadedFileCard').classList.remove('hidden');

    // Cập nhật text xem trước thô
    document.getElementById('txtRawExtracted').textContent = result.text;
    document.getElementById('rawTextMeta').textContent = `${result.text.split('\n').length} dòng`;

    // Điền trước thông tin đầu đề nếu trích xuất được từ file Word của người dùng (không áp dụng cho file mẫu)
    if (result.extractedHeader && !file.isSample) {
      const h = result.extractedHeader;
      const fillIfEmpty = (id, val) => {
        const el = document.getElementById(id);
        if (el && !el.value.trim() && val) {
          el.value = val;
        }
      };
      if (h.soGD) fillIfEmpty('inpSoGD', h.soGD);
      if (h.truong) fillIfEmpty('inpTruong', h.truong);
      if (h.kiKiemTra) fillIfEmpty('inpKiKiemTra', h.kiKiemTra);
      if (h.namHoc) fillIfEmpty('inpNamHoc', h.namHoc);
      if (h.mon) fillIfEmpty('inpMon', h.mon);
      if (h.thoiGian) fillIfEmpty('inpThoiGian', h.thoiGian);
    }

    // Cập nhật trạng thái warning và cookie cho các ô thông tin đầu đề
    if (typeof window.updateHeaderValidationUI === 'function') {
      window.updateHeaderValidationUI();
    }

    // Phân tích cú pháp Marker theo tiêu chuẩn Mục 3 & 7
    const parsed = parseExam(result.text);
    parsed.imageMap = state.imageMap;
    state.parsedExam = parsed;

    const errorBanner = document.getElementById('errorBanner');
    const errorList = document.getElementById('errorList');
    const statusBadge = document.getElementById('txtFileStatusBadge');
    const btnMixAndExport = document.getElementById('btnMixAndExport');
    const txtFileStats = document.getElementById('txtFileStats');

    // Kiểm tra định dạng hình ảnh (Chỉ hỗ trợ In Line with Text)
    const nonInlineImages = result.nonInlineImages || [];
    state.nonInlineImages = nonInlineImages;

    const imageWarningBanner = document.getElementById('imageWarningBanner');
    const imageWarningList = document.getElementById('imageWarningList');

    if (imageWarningBanner && imageWarningList) {
      if (nonInlineImages.length > 0) {
        imageWarningBanner.classList.remove('hidden');
        imageWarningList.innerHTML = nonInlineImages.map((img) => `
          <li class="image-warning-item">
            <div class="image-warning-thumb-box">
              ${img.dataUrl
                ? `<img src="${img.dataUrl}" alt="${escapeHtml(img.name)}" class="image-warning-thumb"/>`
                : `<div class="image-warning-thumb-placeholder">📷</div>`}
            </div>
            <div class="image-warning-details">
              <div class="image-warning-name-row">
                <strong class="image-warning-name">${escapeHtml(img.name)}</strong>
                <span class="badge" style="background:#fee2e2; color:#b91c1c; border:1px solid #fca5a5;">
                  Wrap Text hiện tại: ${escapeHtml(img.wrapType)}
                </span>
              </div>
              <div class="image-warning-snippet">
                <span class="snippet-label">Đoạn văn chứa ảnh:</span> "${escapeHtml(img.snippet)}"
              </div>
              <div class="image-warning-action-tip">
                👉 Vui lòng mở file Word gốc ➔ Click chuột phải vào ảnh này ➔ <strong>Wrap Text</strong> ➔ Chọn <strong>In Line with Text</strong>.
              </div>
            </div>
          </li>
        `).join('');
      } else {
        imageWarningBanner.classList.add('hidden');
        imageWarningList.innerHTML = '';
      }
    }

    if (!parsed.valid) {
      const errorMsg = nonInlineImages.length > 0
        ? `⚠ ${parsed.errors.length} lỗi cú pháp • ${nonInlineImages.length} ảnh sai InLine`
        : `⚠ ${parsed.errors.length} lỗi cú pháp`;
      statusBadge.className = 'badge badge-danger';
      statusBadge.textContent = errorMsg;
      errorBanner.classList.remove('hidden');
      errorList.innerHTML = parsed.errors.map(err => `<li>${escapeHtml(err.message)}</li>`).join('');
      btnMixAndExport.disabled = true;
      txtFileStats.textContent = 'Phát hiện lỗi cú pháp và định dạng, hãy sửa theo hướng dẫn trên';
      renderExamPreview(null);
      showSnackbar('Phát hiện lỗi cú pháp trong đề thi. Vui lòng kiểm tra!', 4000);
    } else if (nonInlineImages.length > 0) {
      statusBadge.className = 'badge badge-danger';
      statusBadge.textContent = `⚠ ${nonInlineImages.length} ảnh không phải InLine`;
      errorBanner.classList.add('hidden');
      errorList.innerHTML = '';
      btnMixAndExport.disabled = true;
      txtFileStats.textContent = `Phát hiện ${nonInlineImages.length} ảnh không phải InLine. Cần sửa file gốc sang "In Line with Text" trước khi trộn.`;
      renderExamPreview(parsed);
      showSnackbar(`⚠️ Có ${nonInlineImages.length} ảnh không phải InLine! Vui lòng sửa lại file gốc.`, 5000);
    } else {
      statusBadge.className = 'badge badge-success';
      statusBadge.textContent = '✓ Hợp lệ';
      errorBanner.classList.add('hidden');
      errorList.innerHTML = '';
      btnMixAndExport.disabled = false;

      const { totalGroups, totalQuestions, choiceCount, tfCount, shortCount } = parsed.stats;
      txtFileStats.textContent = `${totalGroups} nhóm • ${totalQuestions} câu hỏi (${choiceCount} trắc nghiệm, ${tfCount} đúng/sai, ${shortCount} trả lời ngắn)`;

      renderExamPreview(parsed);
      showSnackbar(`Đã nạp thành công đề thi: ${totalQuestions} câu hỏi!`, 3000);
    }
  } catch (err) {
    console.error('Lỗi khi nạp file Word:', err);
    showSnackbar(`Lỗi khi đọc file: ${err.message}`, 4000);
  }
}

/**
 * Lấy các thông số cấu hình đầu đề từ giao diện
 */
export function getHeaderConfig() {
  const getVal = (id) => {
    const el = document.getElementById(id);
    return el ? el.value.trim() : '';
  };

  return {
    soGD: getVal('inpSoGD'),
    truong: getVal('inpTruong'),
    kiKiemTra: getVal('inpKiKiemTra'),
    namHoc: getVal('inpNamHoc'),
    mon: getVal('inpMon'),
    thoiGian: getVal('inpThoiGian')
  };
}

/**
 * Lấy các tùy chọn trộn đề
 */
export function getMixingOptions() {
  const radioNumbering = document.querySelector('input[name="radioNumbering"]:checked');
  const numberingFormat = radioNumbering ? radioNumbering.value : 'dot';
  const chkPerGroup = document.getElementById('chkNumberingPerGroup');
  const numberingScope = (chkPerGroup && chkPerGroup.checked) ? 'group' : 'all';
  const startQNum = parseInt(document.getElementById('inpStartQNum').value, 10) || 1;
  const defaultLabelStyle = document.getElementById('selDefaultLabelStyle').value || 'ABCD';
  const shuffleGroups = document.getElementById('chkShuffleGroups') ? document.getElementById('chkShuffleGroups').checked : false;
  const insertBlankLine = document.getElementById('chkInsertBlankLine') ? document.getElementById('chkInsertBlankLine').checked : false;

  return {
    numberingFormat,
    numberingScope,
    startQuestionNumber: startQNum,
    defaultLabelStyle,
    shuffleGroups,
    insertBlankLine
  };
}

/**
 * Thực hiện Trộn & Tạo các file Word DOCX và ZIP
 */
export async function executeMixAndExport() {
  if (!state.parsedExam || !state.parsedExam.valid) {
    showSnackbar('Vui lòng nạp file đề Word hợp lệ trước khi trộn!');
    return;
  }

  const btnMix = document.getElementById('btnMixAndExport');
  const originalText = btnMix.innerHTML;

  if (state.nonInlineImages && state.nonInlineImages.length > 0) {
    showSnackbar(`⚠️ File còn ${state.nonInlineImages.length} hình ảnh chưa chuyển sang In Line with Text! Vui lòng chỉnh sửa file gốc.`, 5000);
    return;
  }

  try {
    btnMix.disabled = true;
    btnMix.innerHTML = `
      <span class="spin-icon" style="display:inline-block; margin-right: 8px;">🔄</span>
      <span>Đang tạo các file Word...</span>
    `;

    const headerInfo = getHeaderConfig();
    headerInfo.imageMap = state.imageMap || {};
    const mixingOptions = getMixingOptions();

    // Cảnh báo nếu các mục thông tin đầu đề thi chưa được nhập
    const emptyFields = [];
    if (!headerInfo.soGD) emptyFields.push('Sở/Phòng GD');
    if (!headerInfo.truong) emptyFields.push('Tên trường');
    if (!headerInfo.kiKiemTra) emptyFields.push('Kì kiểm tra');
    if (!headerInfo.namHoc) emptyFields.push('Năm học');
    if (!headerInfo.mon) emptyFields.push('Môn thi');
    if (!headerInfo.thoiGian) emptyFields.push('Thời gian');

    if (emptyFields.length > 0) {
      showSnackbar(`⚠️ Chú ý: Còn ${emptyFields.length} mục đầu đề chưa điền (${emptyFields.slice(0, 3).join(', ')}${emptyFields.length > 3 ? '...' : ''})!`, 3500);
    }

    // 1. Hoán vị tất cả các mã đề
    const shuffledExams = shuffleAllExams(state.parsedExam, state.examCodes, mixingOptions);
    shuffledExams.forEach(exam => {
      exam.imageMap = state.imageMap || {};
    });

    // 2. Tạo JSZip để đóng gói
    const JSZip = (typeof window !== 'undefined' && window.JSZip) ? window.JSZip : globalThis.JSZip;
    if (!JSZip) throw new Error('Thư viện JSZip chưa sẵn sàng!');
    const zip = new JSZip();

    state.generatedFiles = [];

    // 3. Tạo từng file Word cho mỗi mã đề (Đề thi & Đáp án chi tiết)
    for (const exam of shuffledExams) {
      // 3.1 File Đề thi (học sinh làm, không hiển thị đáp án)
      const docxBlob = await createExamDocxBlob(exam, headerInfo, { isAnswerKey: false, insertBlankLine: mixingOptions.insertBlankLine });
      const examFileName = `Đề_${exam.examCode}.docx`;
      zip.file(examFileName, docxBlob);
      const examUrl = URL.createObjectURL(docxBlob);
      state.generatedFiles.push({
        name: examFileName,
        blob: docxBlob,
        url: examUrl,
        isAnswerKey: false
      });

      // 3.2 File Đáp án chi tiết tương ứng (các câu đúng in đậm và có màu blue)
      const answerDocxBlob = await createExamDocxBlob(exam, headerInfo, { isAnswerKey: true, insertBlankLine: mixingOptions.insertBlankLine });
      const answerFileName = `Đáp_Án_${exam.examCode}.docx`;
      zip.file(answerFileName, answerDocxBlob);
      const answerUrl = URL.createObjectURL(answerDocxBlob);
      state.generatedFiles.push({
        name: answerFileName,
        blob: answerDocxBlob,
        url: answerUrl,
        isAnswerKey: true
      });
    }

    // 4. Tạo file Word Bảng Đáp Án Ma Trận (tổng hợp tất cả các mã đề)
    const answerBlob = await createAnswerKeyDocxBlob(shuffledExams, headerInfo, mixingOptions);
    const safeMon = (headerInfo.mon || 'Mon_Hoc').trim().replace(/[\s/\\:*?"<>|]+/g, '_') || 'Mon_Hoc';
    const answerFileName = `Bảng_Đáp_Án_${safeMon}.docx`;
    zip.file(answerFileName, answerBlob);
    const answerUrl = URL.createObjectURL(answerBlob);
    state.generatedFiles.push({
      name: answerFileName,
      blob: answerBlob,
      url: answerUrl,
      isAnswerKey: true
    });

    // 5. Đóng gói ZIP
    const zipBlob = await zip.generateAsync({ type: 'blob' });
    state.generatedZipBlob = zipBlob;

    // 6. Hiển thị bảng ma trận đáp án trên Modal
    const matrix = buildAnswerMatrix(shuffledExams);
    renderAnswerMatrixTable(matrix);

    // 7. Hiển thị danh sách file tải lẻ
    renderDownloadFileList();

    // 8. Mở modal kết quả
    document.getElementById('exportModal').classList.add('active');

    // Kiểm tra khả năng Web Share API
    const btnShare = document.getElementById('btnShareFiles');
    if (navigator.canShare && navigator.canShare({ files: [new File([zipBlob], 'de_thi.zip', { type: 'application/zip' })] })) {
      btnShare.classList.remove('hidden');
    } else {
      btnShare.classList.add('hidden');
    }

    showSnackbar('Đã tạo thành công bộ đề thi!', 3000);
  } catch (err) {
    console.error('Lỗi khi trộn và xuất đề:', err);
    showSnackbar(`Lỗi khi tạo file: ${err.message}`, 4000);
  } finally {
    btnMix.disabled = false;
    btnMix.innerHTML = originalText;
  }
}

/**
 * Hiển thị bảng ma trận đáp án trên Modal
 */
function renderAnswerMatrixTable(matrix) {
  const table = document.getElementById('tblAnswerPreview');
  if (!table) return;

  let html = '<thead><tr><th>Câu</th>';
  matrix.headers.forEach(code => {
    html += `<th>Mã ${code}</th>`;
  });
  html += '</tr></thead><tbody>';

  matrix.rows.forEach(row => {
    html += `<tr><td><strong>Câu ${row.questionNumber}</strong></td>`;
    matrix.headers.forEach(code => {
      html += `<td class="correct-val">${escapeHtml(row.answers[code] || '-')}</td>`;
    });
    html += '</tr>';
  });

  html += '</tbody>';
  table.innerHTML = html;
}

/**
 * Hiển thị danh sách file riêng lẻ trên Modal
 */
function renderDownloadFileList() {
  const container = document.getElementById('fileDownloadList');
  if (!container) return;

  container.innerHTML = '';
  state.generatedFiles.forEach(file => {
    const item = document.createElement('div');
    item.className = 'file-download-item';
    item.innerHTML = `
      <div class="file-download-name">
        <span>📄</span>
        <span>${escapeHtml(file.name)}</span>
        ${file.isAnswerKey ? '<span class="badge badge-success" style="font-size:0.7rem;">Đáp án</span>' : ''}
      </div>
      <a href="${file.url}" download="${file.name}" class="btn btn-secondary btn-xs">
        Tải về
      </a>
    `;
    container.appendChild(item);
  });
}

/**
 * Tải toàn bộ file ZIP
 */
export function downloadAllZip() {
  if (!state.generatedZipBlob) return;
  const headerInfo = getHeaderConfig();
  const zipName = `Bo_De_Thi_${headerInfo.mon.replace(/\s+/g, '_')}_${state.examCodes.length}_MaDe.zip`;

  const url = URL.createObjectURL(state.generatedZipBlob);
  const a = document.createElement('a');
  a.href = url;
  a.download = zipName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}
