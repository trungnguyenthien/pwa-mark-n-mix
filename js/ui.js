/**
 * Module Giao diện Người dùng (UI Controller)
 * Quản lý tương tác, cập nhật thời gian thực, tô màu cú pháp và xuất file Word
 */

import { parseExam } from './parser.js';
import { shuffleExamBatch } from './shuffler.js';
import { createExamDocxBlob } from './docx-builder.js';
import { createAnswerKeyDocxBlob, generateAnswerMatrix } from './answer-key.js';
import { readExamFile } from './docx-importer.js';

// Mẫu đề thi chuẩn theo mục 3.4 của tài liệu
export const SAMPLE_EXAM = `__G__
__QUESTION__ Hai đại lượng nào sau đây là hai đại lượng vectơ?
__S__ Quãng đường và tốc độ.
__D__ Độ dịch chuyển và vận tốc.
__S__ Quãng đường và độ dịch chuyển.
__S#__ Cả ba đáp án trên đều sai.

__G__
__QUESTION__ Khi chơi gôn, phát biểu nào đúng?
__S__ Chuyển động của quả bóng là chuyển động thẳng.
__D__ Gia tốc ở điểm cao nhất bằng gia tốc rơi tự do.
__S__ Tốc độ ở điểm cao nhất bằng không.
__D__ Tốc độ ban đầu không nhỏ hơn 50 m/s.

__GQ__
__QUESTION__ Tốc độ trung bình của vận động viên là bao nhiêu m/s?
__D__ 4`;

let currentParsedData = null;
let currentGroupStyles = {}; // id nhóm -> 'ABCD' | 'abcd' | '1234'
let generatedZipBlob = null;
let generatedFilesList = []; // { name, blob }

/**
 * Hiển thị thông báo Snackbar
 */
export function showSnackbar(message, duration = 3000) {
  const bar = document.getElementById('snackbar');
  if (!bar) return;
  bar.textContent = message;
  bar.classList.add('show');
  clearTimeout(bar._timer);
  bar._timer = setTimeout(() => {
    bar.classList.remove('show');
  }, duration);
}

/**
 * Khởi tạo toàn bộ giao diện và các event listeners
 */
export function initUI() {
  const rawEditor = document.getElementById('rawEditor');
  const lineNumbers = document.getElementById('lineNumbers');
  const charCount = document.getElementById('charCount');
  const lineCount = document.getElementById('lineCount');

  // 1. Cập nhật số dòng và đồng bộ cuộn (Line Numbers)
  function updateLineNumbers() {
    const text = rawEditor.value;
    const lines = text.split('\n');
    const totalLines = lines.length;
    let numbersHtml = '';
    for (let i = 1; i <= totalLines; i++) {
      numbersHtml += `${i}\n`;
    }
    lineNumbers.textContent = numbersHtml;
    lineCount.textContent = `${totalLines} dòng`;
    charCount.textContent = `${text.length} ký tự`;
  }

  rawEditor.addEventListener('scroll', () => {
    lineNumbers.scrollTop = rawEditor.scrollTop;
  });

  // 2. Chèn nhanh marker tại vị trí con trỏ
  const chips = document.querySelectorAll('.chip[data-insert]');
  chips.forEach((chip) => {
    chip.addEventListener('click', () => {
      const insertText = chip.getAttribute('data-insert');
      const start = rawEditor.selectionStart;
      const end = rawEditor.selectionEnd;
      const text = rawEditor.value;

      // Nếu marker nhóm, kiểm tra có cần xuống dòng trước không
      let prefix = '';
      if (insertText.startsWith('__G') && start > 0 && text[start - 1] !== '\n') {
        prefix = '\n';
      }

      const insertion = prefix + insertText;
      rawEditor.value = text.substring(0, start) + insertion + text.substring(end);
      const newPos = start + insertion.length;
      rawEditor.selectionStart = newPos;
      rawEditor.selectionEnd = newPos;
      rawEditor.focus();

      updateLineNumbers();
      triggerLiveParse();
    });
  });

  // 3. Tải đề mẫu
  const btnLoadSample = document.getElementById('btnLoadSample');
  btnLoadSample.addEventListener('click', () => {
    rawEditor.value = SAMPLE_EXAM;
    updateLineNumbers();
    triggerLiveParse();
    showSnackbar('Đã nạp đề mẫu trắc nghiệm thành công!');
  });

  // 4. Nạp file Word (.docx) hoặc file văn bản (.txt)
  const fileUploadInput = document.getElementById('fileUploadInput');
  async function handleLoadFile(file) {
    if (!file) return;
    try {
      showSnackbar(`Đang xử lý và đọc tệp '${file.name}'...`, 2000);
      const res = await readExamFile(file);
      rawEditor.value = res.text;
      updateLineNumbers();
      runParsingAndRender(); // Cập nhật bản xem trước ngay lập tức

      // Tự động điền các ô thông tin đầu đề nếu trích xuất được từ file Word
      if (res.extractedHeader) {
        if (res.extractedHeader.soGD) document.getElementById('inpSoGD').value = res.extractedHeader.soGD;
        if (res.extractedHeader.truong) document.getElementById('inpTruong').value = res.extractedHeader.truong;
        if (res.extractedHeader.mon) document.getElementById('inpMon').value = res.extractedHeader.mon;
        if (res.extractedHeader.thoiGian) document.getElementById('inpThoiGian').value = res.extractedHeader.thoiGian;
        if (res.extractedHeader.namHoc) document.getElementById('inpNamHoc').value = res.extractedHeader.namHoc;
        if (res.extractedHeader.kiKiemTra) document.getElementById('inpKiKiemTra').value = res.extractedHeader.kiKiemTra;
      }

      // Tự động chuyển sang tab Xem trước nếu trên màn hình di động để người dùng thấy ngay
      if (window.innerWidth <= 992) {
        document.getElementById('tabBtnPreview').click();
      }

      if (res.isConverted) {
        showSnackbar(`Đã nạp file Word '${file.name}' và tự động chuyển đổi sang hệ marker!`);
      } else {
        showSnackbar(`Đã nạp thành công nội dung từ '${file.name}'!`);
      }
    } catch (err) {
      console.error('Lỗi khi đọc file:', err);
      alert(`Không thể đọc file: ${err.message}`);
    }
  }

  fileUploadInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
      handleLoadFile(file);
      fileUploadInput.value = ''; // Reset input để cho phép chọn lại cùng file
    }
  });

  // Hỗ trợ Kéo & Thả file trực tiếp vào ô soạn thảo
  const editorWrapper = document.querySelector('.editor-wrapper');
  if (editorWrapper) {
    ['dragenter', 'dragover'].forEach(eventName => {
      editorWrapper.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        editorWrapper.style.outline = '2px dashed var(--primary)';
        editorWrapper.style.backgroundColor = 'var(--primary-light)';
      });
    });

    ['dragleave', 'drop'].forEach(eventName => {
      editorWrapper.addEventListener(eventName, (e) => {
        e.preventDefault();
        e.stopPropagation();
        editorWrapper.style.outline = '';
        editorWrapper.style.backgroundColor = '';
      });
    });

    editorWrapper.addEventListener('drop', (e) => {
      const dt = e.dataTransfer;
      const files = dt.files;
      if (files && files.length > 0) {
        handleLoadFile(files[0]);
      }
    });
  }

  // 5. Xóa toàn bộ
  const btnClearEditor = document.getElementById('btnClearEditor');
  btnClearEditor.addEventListener('click', () => {
    if (confirm('Bạn có chắc chắn muốn xóa toàn bộ nội dung trong ô soạn thảo?')) {
      rawEditor.value = '';
      updateLineNumbers();
      triggerLiveParse();
    }
  });

  // 6. Header Settings Card Accordion
  const headerSettingsToggle = document.getElementById('headerSettingsToggle');
  const headerSettingsCard = document.querySelector('.header-settings-card');
  headerSettingsToggle.addEventListener('click', () => {
    headerSettingsCard.classList.toggle('collapsed');
  });

  // 7. Mobile Navigation Tabs
  const tabBtnEditor = document.getElementById('tabBtnEditor');
  const tabBtnPreview = document.getElementById('tabBtnPreview');
  const panelEditor = document.getElementById('panelEditor');
  const panelPreview = document.getElementById('panelPreview');

  tabBtnEditor.addEventListener('click', () => {
    tabBtnEditor.classList.add('active');
    tabBtnEditor.setAttribute('aria-selected', 'true');
    tabBtnPreview.classList.remove('active');
    tabBtnPreview.setAttribute('aria-selected', 'false');

    panelEditor.classList.add('active');
    panelPreview.classList.remove('active');
  });

  tabBtnPreview.addEventListener('click', () => {
    tabBtnPreview.classList.add('active');
    tabBtnPreview.setAttribute('aria-selected', 'true');
    tabBtnEditor.classList.remove('active');
    tabBtnEditor.setAttribute('aria-selected', 'false');

    panelPreview.classList.add('active');
    panelEditor.classList.remove('active');
  });

  // 8. Tùy chọn Số lượng đề & Mã đề
  const inpNumExams = document.getElementById('inpNumExams');
  const btnDecExams = document.getElementById('btnDecExams');
  const btnIncExams = document.getElementById('btnIncExams');
  const inpFirstCode = document.getElementById('inpFirstCode');

  btnDecExams.addEventListener('click', () => {
    let val = parseInt(inpNumExams.value, 10) || 4;
    if (val > 1) {
      inpNumExams.value = val - 1;
      refreshExamCodeChips();
    }
  });

  btnIncExams.addEventListener('click', () => {
    let val = parseInt(inpNumExams.value, 10) || 4;
    if (val < 24) {
      inpNumExams.value = val + 1;
      refreshExamCodeChips();
    }
  });

  inpNumExams.addEventListener('input', refreshExamCodeChips);
  inpFirstCode.addEventListener('input', refreshExamCodeChips);

  // 9. Lắng nghe thay đổi văn bản để parse thời gian thực
  let debounceTimer = null;
  function triggerLiveParse() {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(() => {
      runParsingAndRender();
    }, 150);
  }

  rawEditor.addEventListener('input', () => {
    updateLineNumbers();
    triggerLiveParse();
  });

  // 10. Chuyển đổi giao diện Sáng / Tối
  const btnThemeToggle = document.getElementById('btnThemeToggle');
  const iconSun = btnThemeToggle.querySelector('.icon-sun');
  const iconMoon = btnThemeToggle.querySelector('.icon-moon');

  btnThemeToggle.addEventListener('click', () => {
    const isLight = document.documentElement.getAttribute('data-theme') === 'light';
    if (isLight) {
      document.documentElement.removeAttribute('data-theme');
      iconSun.classList.remove('hidden');
      iconMoon.classList.add('hidden');
      localStorage.setItem('theme', 'dark');
    } else {
      document.documentElement.setAttribute('data-theme', 'light');
      iconSun.classList.add('hidden');
      iconMoon.classList.remove('hidden');
      localStorage.setItem('theme', 'light');
    }
  });

  // Khôi phục theme đã lưu
  if (localStorage.getItem('theme') === 'light') {
    document.documentElement.setAttribute('data-theme', 'light');
    iconSun.classList.add('hidden');
    iconMoon.classList.remove('hidden');
  }

  // 11. Xử lý xuất file Word (.docx)
  const btnGenerateDocs = document.getElementById('btnGenerateDocs');
  btnGenerateDocs.addEventListener('click', handleExportDocs);

  // 12. Modal kết quả xuất đề
  const exportModal = document.getElementById('exportModal');
  const btnCloseExportModal = document.getElementById('btnCloseExportModal');
  const btnCloseExportDone = document.getElementById('btnCloseExportDone');
  const btnDownloadZip = document.getElementById('btnDownloadZip');
  const btnShareFiles = document.getElementById('btnShareFiles');

  [btnCloseExportModal, btnCloseExportDone].forEach((btn) => {
    btn.addEventListener('click', () => {
      exportModal.classList.remove('open');
    });
  });

  btnDownloadZip.addEventListener('click', () => {
    if (!generatedZipBlob) return;
    downloadBlob(generatedZipBlob, 'Tron_De_Trac_Nghiem_Tron_Bo.zip');
    showSnackbar('Đã bắt đầu tải file ZIP!');
  });

  if (navigator.canShare) {
    btnShareFiles.classList.remove('hidden');
    btnShareFiles.addEventListener('click', async () => {
      if (generatedFilesList.length === 0) return;
      try {
        const filesToShare = generatedFilesList.map(
          f => new File([f.blob], f.name, { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' })
        );
        if (navigator.canShare({ files: filesToShare })) {
          await navigator.share({
            files: filesToShare,
            title: 'Đề thi trắc nghiệm',
            text: 'Các mã đề thi đã được trộn và bảng đáp án.'
          });
        }
      } catch (err) {
        if (err.name !== 'AbortError') {
          showSnackbar('Trình duyệt không hỗ trợ chia sẻ tệp này.');
        }
      }
    });
  }

  // Khởi động với đề mẫu
  rawEditor.value = SAMPLE_EXAM;
  updateLineNumbers();
  refreshExamCodeChips();
  runParsingAndRender();
}

/**
 * Cập nhật danh sách các ô nhập mã đề thi (tự tăng hoặc chỉnh sửa riêng từng mã)
 */
function refreshExamCodeChips() {
  const container = document.getElementById('examCodesContainer');
  const numExams = parseInt(document.getElementById('inpNumExams').value, 10) || 4;
  const firstCode = parseInt(document.getElementById('inpFirstCode').value, 10) || 101;

  // Thu thập các giá trị cũ đã sửa nếu có
  const oldInputs = container.querySelectorAll('.exam-code-input');
  const existingValues = Array.from(oldInputs).map(inp => inp.value.trim());

  container.innerHTML = '';

  for (let i = 0; i < numExams; i++) {
    const val = existingValues[i] || String(firstCode + i);
    const badge = document.createElement('div');
    badge.className = 'exam-code-badge';
    badge.innerHTML = `
      <span style="font-size: 0.75rem; color: var(--text-muted);">Đề ${i + 1}:</span>
      <input type="text" class="exam-code-input" value="${val}" maxlength="6" aria-label="Mã đề ${i + 1}">
    `;
    container.appendChild(badge);
  }
}

/**
 * Lấy danh sách các mã đề hiện tại từ các ô nhập
 */
function getCurrentExamCodes() {
  const container = document.getElementById('examCodesContainer');
  const inputs = container.querySelectorAll('.exam-code-input');
  const codes = [];
  inputs.forEach((inp, idx) => {
    const val = inp.value.trim() || `10${idx + 1}`;
    codes.push(val);
  });
  return codes;
}

/**
 * Phân tích đề và cập nhật toàn bộ hiển thị xem trước, thống kê, bảng lỗi
 */
function runParsingAndRender() {
  const rawText = document.getElementById('rawEditor').value;
  const result = parseExam(rawText);
  currentParsedData = result;

  const badgeQuestions = document.getElementById('mobileBadgeQuestions');
  badgeQuestions.textContent = result.stats.totalQuestions;

  // Thống kê
  document.getElementById('statGroups').textContent = result.stats.totalGroups;
  document.getElementById('statQuestions').textContent = result.stats.totalQuestions;
  document.getElementById('statChoice').textContent = result.stats.choiceCount;
  document.getElementById('statTF').textContent = result.stats.tfCount;
  document.getElementById('statShort').textContent = result.stats.shortCount;

  // Trạng thái hợp lệ
  const validityBadge = document.getElementById('validityStatusBadge');
  const validityText = document.getElementById('validityStatusText');
  const iconValid = validityBadge.querySelector('.icon-valid');
  const iconInvalid = validityBadge.querySelector('.icon-invalid');
  const errorAlertBanner = document.getElementById('errorAlertBanner');
  const errorList = document.getElementById('errorList');

  if (result.valid) {
    validityBadge.className = 'validity-badge valid';
    validityText.textContent = 'Cú pháp hợp lệ';
    iconValid.classList.remove('hidden');
    iconInvalid.classList.add('hidden');
    errorAlertBanner.classList.add('hidden');
    errorList.innerHTML = '';
  } else {
    validityBadge.className = 'validity-badge invalid';
    validityText.textContent = `Phát hiện ${result.errors.length} lỗi`;
    iconValid.classList.add('hidden');
    iconInvalid.classList.remove('hidden');
    errorAlertBanner.classList.remove('hidden');

    errorList.innerHTML = '';
    result.errors.forEach((err) => {
      const li = document.createElement('li');
      li.className = 'error-item';
      li.textContent = err.message;
      li.title = 'Nhấn để cuộn tới dòng lỗi trong trình soạn thảo';
      li.addEventListener('click', () => {
        highlightEditorLine(err.line);
      });
      errorList.appendChild(li);
    });
  }

  // Render bản xem trước trực quan
  renderLivePreview(result.groups);
}

/**
 * Cuộn và bôi đen dòng lỗi trong textarea
 */
function highlightEditorLine(lineNumber) {
  const rawEditor = document.getElementById('rawEditor');
  const lines = rawEditor.value.split('\n');
  if (lineNumber > lines.length) return;

  let startPos = 0;
  for (let i = 0; i < lineNumber - 1; i++) {
    startPos += lines[i].length + 1; // +1 cho \n
  }
  const endPos = startPos + lines[lineNumber - 1].length;

  rawEditor.focus();
  rawEditor.setSelectionRange(startPos, endPos);

  // Tính toán khoảng cuộn
  const lineHeight = 19.5; // px xấp xỉ
  rawEditor.scrollTop = Math.max(0, (lineNumber - 4) * lineHeight);
}

/**
 * Render bản xem trước các nhóm và câu hỏi có tô màu
 */
function renderLivePreview(groups) {
  const container = document.getElementById('livePreviewContainer');
  container.innerHTML = '';

  if (!groups || groups.length === 0) {
    container.innerHTML = `
      <div style="text-align: center; color: var(--text-muted); padding: 40px 10px;">
        Chưa có câu hỏi nào. Hãy dán đề thi hoặc bấm nút <strong>'Tải đề mẫu'</strong> ở góc trên.
      </div>
    `;
    return;
  }

  const modeDescriptions = {
    G: 'Đổi thứ tự câu & đáp án',
    GQ: 'Đổi thứ tự câu, giữ nguyên đáp án',
    GA: 'Giữ nguyên câu, đổi đáp án',
    G0: 'Giữ nguyên cả câu & đáp án'
  };

  groups.forEach((group) => {
    const groupCard = document.createElement('div');
    groupCard.className = 'preview-group-card';

    // Lưu lại kiểu nhãn đã chọn trước đó nếu có
    const currentStyle = currentGroupStyles[group.id] || group.labelStyle || 'ABCD';
    currentGroupStyles[group.id] = currentStyle;

    const modeText = modeDescriptions[group.mode] || group.mode;
    const fixedText = group.fixed ? " • <span style='color: var(--warning);'>[Nhóm Cố Định #]</span>" : '';

    groupCard.innerHTML = `
      <div class="preview-group-header">
        <div class="group-tag">
          <span class="group-badge">__${group.mode}${group.fixed ? '#' : ''}__</span>
          <span class="group-desc">${modeText}${fixedText}</span>
        </div>
        <div class="group-style-picker">
          <label for="selGroupStyle_${group.id}">Kiểu nhãn:</label>
          <select id="selGroupStyle_${group.id}" class="select-sm">
            <option value="ABCD" ${currentStyle === 'ABCD' ? 'selected' : ''}>A. B. C. D.</option>
            <option value="abcd" ${currentStyle === 'abcd' ? 'selected' : ''}>a) b) c) d)</option>
            <option value="1234" ${currentStyle === '1234' ? 'selected' : ''}>1. 2. 3. 4.</option>
          </select>
        </div>
      </div>
      <div class="group-questions-container" id="groupQuestions_${group.id}"></div>
    `;

    container.appendChild(groupCard);

    // Lắng nghe thay đổi kiểu nhãn cho nhóm
    const selectElem = groupCard.querySelector(`#selGroupStyle_${group.id}`);
    selectElem.addEventListener('change', (e) => {
      currentGroupStyles[group.id] = e.target.value;
      group.labelStyle = e.target.value;
    });

    const questionsContainer = groupCard.querySelector(`#groupQuestions_${group.id}`);

    group.questions.forEach((q) => {
      const qCard = document.createElement('div');
      qCard.className = 'preview-question-card';

      let answersHtml = '';
      if (q.type === 'short') {
        const val = q.answers[0] ? q.answers[0].text : '';
        answersHtml = `
          <div class="short-answer-badge">
            <strong>Đáp án điền khuyết:</strong> ${escapeHtml(val)} (Ẩn trong đề thi, xuất trong bảng đáp án)
          </div>
        `;
      } else {
        answersHtml = '<div class="q-answers-list">';
        q.answers.forEach((ans) => {
          const markerClass = ans.correct ? 'is-correct' : 'is-wrong';
          const markerLabel = ans.correct ? '__D__' : '__S__';
          const fixedIcon = ans.fixed ? `<span class="fixed-icon" title="Cố định vị trí trong câu">🔒 #</span>` : '';

          answersHtml += `
            <div class="ans-row ${markerClass}">
              <span class="ans-marker-pill">${markerLabel}</span>
              <span class="ans-text">${escapeHtml(ans.text)}${fixedIcon}</span>
            </div>
          `;
        });
        answersHtml += '</div>';
      }

      qCard.innerHTML = `
        <div class="q-title">Câu ${q.id}. ${escapeHtml(q.text)}</div>
        ${answersHtml}
      `;

      questionsContainer.appendChild(qCard);
    });
  });
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
 * Xử lý sự kiện xuất file đề thi và bảng đáp án
 */
async function handleExportDocs() {
  if (!currentParsedData || !currentParsedData.valid) {
    showSnackbar('Vui lòng sửa các lỗi cú pháp trước khi xuất đề thi!');
    // Chuyển sang tab xem trước trên di động
    document.getElementById('tabBtnPreview').click();
    return;
  }

  const btnGenerateDocs = document.getElementById('btnGenerateDocs');
  const originalBtnHtml = btnGenerateDocs.innerHTML;
  btnGenerateDocs.disabled = true;
  btnGenerateDocs.innerHTML = `
    <svg class="spin-icon spinning" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
      <circle cx="12" cy="12" r="10" stroke-opacity="0.25"></circle>
      <path d="M12 2a10 10 0 0 1 10 10"></path>
    </svg>
    <span>Đang hoán vị &amp; tạo file Word...</span>
  `;

  try {
    const examCodes = getCurrentExamCodes();
    const shuffleGroups = document.getElementById('chkShuffleGroups').checked;
    const radioNumbering = document.querySelector('input[name="radioNumbering"]:checked').value;
    const startQNum = parseInt(document.getElementById('inpStartQNum').value, 10) || 1;

    const headerInfo = {
      soGD: document.getElementById('inpSoGD').value.trim(),
      truong: document.getElementById('inpTruong').value.trim(),
      kiKiemTra: document.getElementById('inpKiKiemTra').value.trim(),
      namHoc: document.getElementById('inpNamHoc').value.trim(),
      mon: document.getElementById('inpMon').value.trim(),
      thoiGian: document.getElementById('inpThoiGian').value.trim(),
    };

    // 1. Hoán vị các mã đề thi
    const shuffledExams = shuffleExamBatch(currentParsedData, examCodes, {
      shuffleGroups,
      questionNumbering: radioNumbering,
      questionStartNum: startQNum,
      groupLabelStyles: currentGroupStyles
    });

    // 2. Khởi tạo JSZip để đóng gói
    const zip = new window.JSZip();
    generatedFilesList = [];

    // 3. Tạo từng file đề thi .docx
    for (const exam of shuffledExams) {
      const fileName = `De_${exam.examCode}.docx`;
      const blob = await createExamDocxBlob(exam, headerInfo);
      zip.file(fileName, blob);
      generatedFilesList.push({ name: fileName, blob, code: exam.examCode, type: 'exam' });
    }

    // 4. Tạo file Bảng đáp án .docx
    const answerKeyBlob = await createAnswerKeyDocxBlob(shuffledExams, headerInfo);
    const ansFileName = 'Dap_an.docx';
    zip.file(ansFileName, answerKeyBlob);
    generatedFilesList.push({ name: ansFileName, blob: answerKeyBlob, type: 'answerKey' });

    // 5. Nén file ZIP
    generatedZipBlob = await zip.generateAsync({ type: 'blob' });

    // 6. Hiển thị Modal Kết quả
    renderExportModal(shuffledExams);

    showSnackbar('Trộn đề thi thành công!');
  } catch (err) {
    console.error('Lỗi khi xuất đề:', err);
    alert(`Đã xảy ra lỗi khi tạo file Word: ${err.message}`);
  } finally {
    btnGenerateDocs.disabled = false;
    btnGenerateDocs.innerHTML = originalBtnHtml;
  }
}

/**
 * Hiển thị danh sách file và bảng đáp án trong Modal
 */
function renderExportModal(shuffledExams) {
  const fileDownloadList = document.getElementById('fileDownloadList');
  fileDownloadList.innerHTML = '';

  generatedFilesList.forEach((file) => {
    const row = document.createElement('div');
    row.className = 'file-row';

    const isKey = file.type === 'answerKey';
    const iconColor = isKey ? 'var(--warning)' : 'var(--primary)';
    const labelTitle = isKey ? 'Bảng đáp án tổng hợp' : `Mã đề ${file.code}`;

    row.innerHTML = `
      <div class="file-row-info">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="${iconColor}" stroke-width="2">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
          <polyline points="14 2 14 8 20 8"></polyline>
        </svg>
        <div>
          <div>${file.name}</div>
          <div style="font-size: 0.75rem; color: var(--text-muted); font-weight: normal;">${labelTitle}</div>
        </div>
      </div>
      <button class="btn btn-secondary btn-sm btn-dl-file">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
          <polyline points="7 10 12 15 17 10"></polyline>
          <line x1="12" y1="15" x2="12" y2="3"></line>
        </svg>
        <span>Tải về</span>
      </button>
    `;

    row.querySelector('.btn-dl-file').addEventListener('click', () => {
      downloadBlob(file.blob, file.name);
    });

    fileDownloadList.appendChild(row);
  });

  // Render bảng ma trận đáp án nhanh
  const matrix = generateAnswerMatrix(shuffledExams);
  const tbl = document.getElementById('tblAnswerPreview');
  tbl.innerHTML = '';

  // Header row
  const thead = document.createElement('thead');
  const trHead = document.createElement('tr');
  matrix.headers.forEach((h) => {
    const th = document.createElement('th');
    th.textContent = h;
    trHead.appendChild(th);
  });
  thead.appendChild(trHead);
  tbl.appendChild(thead);

  // Body rows
  const tbody = document.createElement('tbody');
  matrix.rows.forEach((r) => {
    const tr = document.createElement('tr');
    r.forEach((cell, idx) => {
      const td = document.createElement('td');
      td.textContent = cell;
      if (idx > 0) {
        td.style.color = 'var(--primary)';
        td.style.fontWeight = '600';
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  tbl.appendChild(tbody);

  // Mở modal
  const exportModal = document.getElementById('exportModal');
  exportModal.classList.add('open');
}

/**
 * Tải file Blob an toàn trên di động và desktop
 */
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);

  // Thu hồi Object URL an toàn sau 60 giây theo mục 7 pwa-standards
  setTimeout(() => {
    URL.revokeObjectURL(url);
  }, 60000);
}
