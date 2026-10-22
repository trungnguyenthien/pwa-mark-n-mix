/**
 * Module Giao Diện Người Dùng (UI Controller)
 * Quản lý tương tác nạp file Word (.docx), cấu hình đầu đề, tùy chọn trộn,
 * hiển thị xem trước trực quan và xử lý xuất file DOCX / ZIP.
 */

import { readExamDocx } from './docx-importer.js';
import { parseExam } from './parser.js';
import { shuffleAllExams } from './shuffler.js';
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

  let html = '';
  let questionCounter = 0;
  const imageMap = parsedExam.imageMap || state.imageMap || {};

  function renderFormattedContent(text) {
    if (!text) return '';
    const safe = escapeHtml(text);
    // Giữ các dấu xuống hàng tự nhiên
    const withBreaks = safe.replace(/\n/g, '<br/>');
    return withBreaks.replace(/\[IMG:([^\]]+)\]/g, (match, id) => {
      const img = imageMap[id];
      if (img && img.dataUrl) {
        const w = img.width || 180;
        const h = img.height || 100;
        const ratio = (w / h).toFixed(2);
        return `<span class="exam-img-frame" title="Hình ảnh từ file Word gốc: ${w} × ${h} px (Tỉ lệ: ${ratio})">` +
          `<img src="${img.dataUrl}" alt="Hình ảnh đề thi" class="exam-preview-img" ` +
          `style="width: ${w}px; aspect-ratio: ${w} / ${h};" ` +
          `data-width="${w}" data-height="${h}" data-img-id="${id}" loading="lazy" />` +
          `</span>`;
      }
      return '';
    });
  }

  parsedExam.groups.forEach((group, gIdx) => {
    const modeBadge = group.mode || 'G';
    const isFixedGroup = group.fixed;
    const modeDesc = {
      'G': 'Đổi cả câu hỏi & đáp án',
      'GQ': 'Chỉ đổi câu hỏi',
      'GA': 'Chỉ đổi đáp án',
      'G0': 'Không đảo vị trí'
    }[modeBadge] || 'Hoán vị';

    html += `
      <div class="preview-group-box">
        <div class="preview-group-header">
          <div>
            <span>Nhóm ${gIdx + 1}</span>
            <span class="group-mode-pill">${group.rawMarker} - ${modeDesc}</span>
            ${isFixedGroup ? '<span class="badge badge-success" style="margin-left: 6px;">Vị trí cố định (#)</span>' : ''}
          </div>
          <div style="font-size: 0.8125rem; color: var(--text-muted); font-weight: 500;">
            ${group.questions.length} câu hỏi
          </div>
        </div>
        <div class="preview-questions-list">
    `;

    group.questions.forEach((q) => {
      questionCounter++;
      html += `
        <div class="question-card">
          <div class="question-title-line">
            <span class="question-qnum">Câu ${questionCounter}.</span>
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
          const labelTemp = String.fromCharCode(65 + (aIdx % 26));

          html += `
            <div class="answer-item">
              <span class="ans-pill ${isD ? 'd' : 's'}">
                ${isD ? '__D__' : '__S__'}
                ${isFixed ? ' <span class="lock-icon" title="Cố định vị trí (#)">🔒#</span>' : ''}
              </span>
              <span class="answer-text"><strong>${labelTemp}.</strong> ${renderFormattedContent(ans.text)}</span>
            </div>
          `;
        });
      }

      html += `
          </div>
        </div>
      `;
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

    // Điền trước thông tin đầu đề nếu trích xuất được từ file Word
    if (result.extractedHeader) {
      const h = result.extractedHeader;
      if (h.soGD) document.getElementById('inpSoGD').value = h.soGD;
      if (h.truong) document.getElementById('inpTruong').value = h.truong;
      if (h.kiKiemTra) document.getElementById('inpKiKiemTra').value = h.kiKiemTra;
      if (h.namHoc) document.getElementById('inpNamHoc').value = h.namHoc;
      if (h.mon) document.getElementById('inpMon').value = h.mon;
      if (h.thoiGian) document.getElementById('inpThoiGian').value = h.thoiGian;
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

    if (!parsed.valid) {
      statusBadge.className = 'badge badge-danger';
      statusBadge.textContent = `⚠ ${parsed.errors.length} lỗi cú pháp`;
      errorBanner.classList.remove('hidden');
      errorList.innerHTML = parsed.errors.map(err => `<li>${escapeHtml(err.message)}</li>`).join('');
      btnMixAndExport.disabled = true;
      txtFileStats.textContent = 'Phát hiện lỗi định dạng marker, hãy sửa theo hướng dẫn trên';
      renderExamPreview(null);
      showSnackbar('Phát hiện lỗi cú pháp trong đề thi. Vui lòng kiểm tra!', 4000);
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
  return {
    soGD: document.getElementById('inpSoGD').value.trim() || 'SỞ GD&ĐT TP HCM',
    truong: document.getElementById('inpTruong').value.trim() || 'TRƯỜNG THPT NGUYỄN HỮU HUÂN',
    kiKiemTra: document.getElementById('inpKiKiemTra').value.trim() || 'GIỮA HK 1',
    namHoc: document.getElementById('inpNamHoc').value.trim() || '2026 - 2027',
    mon: document.getElementById('inpMon').value.trim() || 'VẬT LÝ',
    thoiGian: document.getElementById('inpThoiGian').value.trim() || '45 phút'
  };
}

/**
 * Lấy các tùy chọn trộn đề
 */
export function getMixingOptions() {
  const radioNumbering = document.querySelector('input[name="radioNumbering"]:checked');
  const numberingFormat = radioNumbering ? radioNumbering.value : 'dot';
  const startQNum = parseInt(document.getElementById('inpStartQNum').value, 10) || 1;
  const defaultLabelStyle = document.getElementById('selDefaultLabelStyle').value || 'ABCD';
  const shuffleGroups = document.getElementById('chkShuffleGroups').checked;

  return {
    numberingFormat,
    startQuestionNumber: startQNum,
    defaultLabelStyle,
    shuffleGroups
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

  try {
    btnMix.disabled = true;
    btnMix.innerHTML = `
      <span class="spin-icon" style="display:inline-block; margin-right: 8px;">🔄</span>
      <span>Đang tạo các file Word...</span>
    `;

    const headerInfo = getHeaderConfig();
    headerInfo.imageMap = state.imageMap || {};
    const mixingOptions = getMixingOptions();

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

    // 3. Tạo từng file Word cho mỗi mã đề
    for (const exam of shuffledExams) {
      const docxBlob = await createExamDocxBlob(exam, headerInfo);
      const fileName = `De_${exam.examCode}.docx`;
      zip.file(fileName, docxBlob);
      const url = URL.createObjectURL(docxBlob);
      state.generatedFiles.push({
        name: fileName,
        blob: docxBlob,
        url,
        isAnswerKey: false
      });
    }

    // 4. Tạo file Word Bảng Đáp Án Ma Trận
    const answerBlob = await createAnswerKeyDocxBlob(shuffledExams, headerInfo);
    const answerFileName = `Bang_Dap_An_${headerInfo.mon.replace(/\s+/g, '_')}.docx`;
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
