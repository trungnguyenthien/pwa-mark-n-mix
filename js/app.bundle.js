(function() {
"use strict";

// --- FILE: js/docx-importer.js ---
/**
 * Module Nhập & Trích Xuất File Word (.docx) Đã Đánh Marker
 * - Sử dụng JSZip để đọc và phân tích cấu trúc word/document.xml trực tiếp trên trình duyệt
 * - Bảo toàn các dấu ngắt dòng (<w:br/>, <w:cr/>), tab và toàn bộ ký tự tiếng Việt Unicode
 * - Tự động trích xuất thông tin đầu đề (Sở GD, Trường, Kì kiểm tra, Năm học, Môn, Thời gian)
 * - Tách và chuẩn hóa các khối marker (__G__, __QUESTION__, __D__, __S__, ...)
 */

function getJSZip() {
  if (typeof window !== 'undefined' && window.JSZip) {
    return window.JSZip;
  }
  if (typeof global !== 'undefined' && global.JSZip) {
    return global.JSZip;
  }
  return null;
}

/**
 * Trích xuất toàn bộ chuỗi văn bản thuần và ngắt dòng từ tệp Word (.docx)
 * @param {Blob|File|ArrayBuffer} file
 * @returns {Promise<string>}
 */
async function extractTextFromDocx(file) {
  const JSZip = getJSZip();
  if (!JSZip) {
    throw new Error('Thư viện JSZip chưa được tải!');
  }

  let buffer;
  if (file instanceof ArrayBuffer) {
    buffer = file;
  } else if (file.arrayBuffer) {
    buffer = await file.arrayBuffer();
  } else {
    buffer = file;
  }

  const zip = await JSZip.loadAsync(buffer);
  const docXmlFile = zip.file('word/document.xml');
  if (!docXmlFile) {
    throw new Error('Tệp không đúng định dạng Word (.docx hợp lệ): thiếu word/document.xml');
  }

  const xmlText = await docXmlFile.async('text');

  // Phân tích cú pháp XML của Word, xử lý từng đoạn văn <w:p>
  const pRegex = /<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g;
  const lines = [];
  let pMatch;

  while ((pMatch = pRegex.exec(xmlText)) !== null) {
    let pContent = pMatch[1];

    // Chuyển đổi các thẻ ngắt dòng và tab
    pContent = pContent.replace(/<w:br(?:\s[^>]*)?\/>/gi, '\n');
    pContent = pContent.replace(/<w:cr(?:\s[^>]*)?\/>/gi, '\n');
    pContent = pContent.replace(/<w:tab(?:\s[^>]*)?\/>/gi, '\t');

    // Giải mã thực thể XML
    let decoded = pContent
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'");

    // Trích xuất text từ các thẻ <w:t>, bảo toàn ký tự xuống dòng \n
    let pText = '';
    const tokenRegex = /(<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|\n)/g;
    let tokMatch;
    while ((tokMatch = tokenRegex.exec(decoded)) !== null) {
      if (tokMatch[0] === '\n') {
        pText += '\n';
      } else {
        pText += tokMatch[2];
      }
    }

    if (!pText && /<w:t/i.test(decoded)) {
      pText = decoded.replace(/<[^>]+>/g, '');
    }

    if (pText.trim()) {
      lines.push(pText.trim());
    }
  }

  const rawExtracted = lines.join('\n');
  return rawExtracted;
}

/**
 * Tự động chuẩn hóa văn bản đã đánh marker từ file Word và trích xuất thông tin đầu đề nếu có
 * @param {string} rawText
 * @returns {Object} { text, extractedHeader }
 */
function normalizeMarkedText(rawText) {
  const extractedHeader = {};

  // Hàm quét thông tin đầu đề (nếu xuất hiện trước marker đầu tiên)
  function checkHeaderLine(line) {
    if (/SỞ\s+GD/i.test(line)) {
      extractedHeader.soGD = line.trim();
      return true;
    }
    if (/TRƯỜNG/i.test(line)) {
      extractedHeader.truong = line.trim();
      return true;
    }
    if (/MÔN[:\s]/i.test(line)) {
      const m = line.match(/MÔN[:\s]+(.*)/i);
      if (m && m[1]) extractedHeader.mon = m[1].trim();
      return true;
    }
    if (/Thời\s+gian/i.test(line)) {
      const m = line.match(/Thời\s+gian[^:]*:\s*(.*)/i);
      if (m && m[1]) extractedHeader.thoiGian = m[1].trim();
      return true;
    }
    if (/NĂM\s+HỌC/i.test(line)) {
      const m = line.match(/NĂM\s+HỌC[:\s]+(.*)/i);
      if (m && m[1]) extractedHeader.namHoc = m[1].trim();
      return true;
    }
    if (/GIỮA\s+HK|CUỐI\s+HK|HỌC\s+KÌ/i.test(line)) {
      extractedHeader.kiKiemTra = line.trim();
      return true;
    }
    return false;
  }

  // Nếu file có chứa các marker đề thi (__QUESTION__ hoặc __G__ hoặc __D__)
  const lines = rawText.split(/\r?\n/);
  const cleanedLines = [];
  let foundFirstMarker = false;

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) continue;

    if (!foundFirstMarker) {
      if (trimmed.startsWith('__G') || trimmed.startsWith('__QUESTION__')) {
        foundFirstMarker = true;
        cleanedLines.push(trimmed);
      } else {
        checkHeaderLine(trimmed);
      }
    } else {
      cleanedLines.push(trimmed);
    }
  }

  // Nếu không tìm thấy marker ở đầu (ví dụ các marker viết dính liền trên 1 dòng như ví dụ 3.4)
  if (cleanedLines.length === 0 && rawText.includes('__QUESTION__')) {
    // Tách các marker bị dính liền: __G____QUESTION__ -> __G__ \n __QUESTION__
    const separated = rawText
      .replace(/(__[A-Za-z0-9#_]+__)/g, '\n$1 ')
      .split('\n')
      .map(s => s.trim())
      .filter(Boolean);
    return {
      text: separated.join('\n'),
      extractedHeader
    };
  }

  return {
    text: cleanedLines.length > 0 ? cleanedLines.join('\n') : rawText,
    extractedHeader
  };
}

/**
 * Đọc file tải lên (.docx)
 * @param {File|Blob} file
 * @returns {Promise<Object>} { text, extractedHeader, fileName, fileSize }
 */
async function readExamDocx(file) {
  const fileName = (file && file.name) ? file.name : 'De_Thi.docx';
  const fileSize = file.size || 0;

  const rawText = await extractTextFromDocx(file);
  const normalized = normalizeMarkedText(rawText);

  return {
    text: normalized.text,
    extractedHeader: normalized.extractedHeader || {},
    fileName,
    fileSize
  };
}


// --- FILE: js/parser.js ---
/**
 * Module Phân Tích Cú Pháp Marker Đề Thi (Parser)
 * Tuân thủ quy định tại mục 3 và mục 7 của tài liệu:
 * "Công cụ trộn đề trắc nghiệm Yêu cầu và giải pháp.docx"
 */

// Danh sách marker nhóm hợp lệ: __G__, __G0__, __GQ__, __GA__ và biến thể cố định #
const GROUP_MARKER_REGEX = /^__(G|G0|GQ|GA)(#?)__$/;
const ALT_GROUP_MARKER_REGEX = /^__(#?)(G|G0|GQ|GA)__$/;

// Danh sách marker đáp án hợp lệ: __D__, __S__, __D#__, __#D__, __S#__, __#S__
const ANSWER_MARKER_REGEX = /^__(D|S)(#?)__$/;
const ALT_ANSWER_MARKER_REGEX = /^__(#?)(D|S)__$/;

/**
 * Kiểm tra và phân loại marker token
 */
function parseMarkerToken(token) {
  // Kiểm tra lỗi chữ O thay cho số 0
  if (token === '__GO__' || token === '__GO#__' || token === '__#GO__') {
    return { error: `Sử dụng marker sai '${token}' (chữ O). Hãy sửa thành '__G0__' (số không).` };
  }

  // Nhóm
  let m = token.match(GROUP_MARKER_REGEX);
  if (m) {
    return { type: 'group', mode: m[1], fixed: m[2] === '#' };
  }
  m = token.match(ALT_GROUP_MARKER_REGEX);
  if (m) {
    return { type: 'group', mode: m[2], fixed: m[1] === '#' };
  }

  // Câu hỏi
  if (token === '__QUESTION__') {
    return { type: 'question' };
  }

  // Đáp án
  m = token.match(ANSWER_MARKER_REGEX);
  if (m) {
    return { type: 'answer', correct: m[1] === 'D', fixed: m[2] === '#' };
  }
  m = token.match(ALT_ANSWER_MARKER_REGEX);
  if (m) {
    return { type: 'answer', correct: m[2] === 'D', fixed: m[1] === '#' };
  }

  return { unknown: true };
}

/**
 * Phân tích văn bản chứa marker thành cấu trúc đề thi
 * @param {string} rawText
 * @returns {Object} { valid: boolean, errors: Array, groups: Array, stats: Object }
 */
function parseExam(rawText) {
  const errors = [];
  const lines = rawText.split(/\r?\n/);

  const groups = [];
  let currentGroup = null;
  let currentQuestion = null;
  let currentAnswer = null;

  let hasMetFirstQuestion = false;
  let questionCounter = 0;

  for (let i = 0; i < lines.length; i++) {
    const lineNumber = i + 1;
    const rawLine = lines[i];
    const trimmedLine = rawLine.trim();

    // Bỏ qua dòng trống
    if (!trimmedLine) continue;

    // Kiểm tra tất cả marker trên dòng
    const markerMatches = [...rawLine.matchAll(/__([A-Za-z0-9#_]+)__/g)];

    if (markerMatches.length > 0) {
      const firstMatch = markerMatches[0];
      const matchIndex = firstMatch.index;
      const leadingPart = rawLine.substring(0, matchIndex);

      // 1. Marker không nằm ở đầu dòng
      if (leadingPart.trim().length > 0) {
        errors.push({
          line: lineNumber,
          message: `Dòng ${lineNumber}: Marker '${firstMatch[0]}' không nằm ở đầu dòng. Marker phải bắt đầu từ đầu dòng.`
        });
        continue;
      }

      // Marker phải có ít nhất 1 khoảng trắng theo sau hoặc là hết dòng
      const afterMarkerIndex = matchIndex + firstMatch[0].length;
      if (afterMarkerIndex < rawLine.length && !/\s/.test(rawLine[afterMarkerIndex])) {
        errors.push({
          line: lineNumber,
          message: `Dòng ${lineNumber}: Marker '${firstMatch[0]}' phải có ít nhất một khoảng trắng theo sau.`
        });
        continue;
      }

      // Nhiều hơn 1 marker trên cùng dòng
      if (markerMatches.length > 1) {
        errors.push({
          line: lineNumber,
          message: `Dòng ${lineNumber}: Phát hiện nhiều marker trên cùng một dòng ('${markerMatches.map(m => m[0]).join("', '")}'). Mỗi marker phải nằm trên một dòng riêng biệt.`
        });
        continue;
      }
    }

    // Dòng bắt đầu bằng marker
    const lineMarkerMatch = rawLine.match(/^\s*(__[A-Za-z0-9#_]+__)(.*)$/);
    if (lineMarkerMatch) {
      const token = lineMarkerMatch[1];
      const content = lineMarkerMatch[2].trim();
      const parsedToken = parseMarkerToken(token);

      if (parsedToken.error) {
        errors.push({
          line: lineNumber,
          message: `Dòng ${lineNumber}: ${parsedToken.error}`
        });
        continue;
      }

      if (parsedToken.unknown) {
        errors.push({
          line: lineNumber,
          message: `Dòng ${lineNumber}: Marker '${token}' không hợp lệ. Các marker được hỗ trợ: __QUESTION__, __D__, __S__, __D#__, __S#__, __G__, __GQ__, __GA__, __G0__ (và các biến thể có '#').`
        });
        continue;
      }

      // Xử lý Marker NHÓM
      if (parsedToken.type === 'group') {
        currentAnswer = null;
        currentQuestion = null;

        currentGroup = {
          id: groups.length + 1,
          mode: parsedToken.mode,
          fixed: parsedToken.fixed,
          rawMarker: token,
          startLine: lineNumber,
          labelStyle: 'ABCD', // Mặc định A. B. C. D.
          questions: []
        };
        groups.push(currentGroup);
        continue;
      }

      // Xử lý Marker CÂU HỎI (__QUESTION__)
      if (parsedToken.type === 'question') {
        hasMetFirstQuestion = true;

        // Nếu chưa có nhóm nào, tạo nhóm mặc định __G__
        if (!currentGroup) {
          currentGroup = {
            id: groups.length + 1,
            mode: 'G',
            fixed: false,
            rawMarker: '__G__',
            startLine: lineNumber,
            labelStyle: 'ABCD',
            questions: []
          };
          groups.push(currentGroup);
        }

        questionCounter++;
        currentAnswer = null;

        currentQuestion = {
          id: questionCounter,
          globalIndex: questionCounter - 1,
          startLine: lineNumber,
          text: content,
          answers: []
        };
        currentGroup.questions.push(currentQuestion);
        continue;
      }

      // Xử lý Marker ĐÁP ÁN (__D__, __S__, __D#__, __S#__)
      if (parsedToken.type === 'answer') {
        if (!currentQuestion) {
          errors.push({
            line: lineNumber,
            message: `Dòng ${lineNumber}: Đáp án '${token}' xuất hiện khi chưa bắt đầu câu hỏi (__QUESTION__).`
          });
          continue;
        }

        currentAnswer = {
          id: currentQuestion.answers.length + 1,
          line: lineNumber,
          correct: parsedToken.correct,
          fixed: parsedToken.fixed,
          rawMarker: token,
          text: content
        };
        currentQuestion.answers.push(currentAnswer);
        continue;
      }
    }

    // Dòng không có marker
    // Kiểm tra: Có nội dung trước __QUESTION__ đầu tiên mà không phải marker nhóm
    if (!hasMetFirstQuestion) {
      errors.push({
        line: lineNumber,
        message: `Dòng ${lineNumber}: Nội dung '${trimmedLine.substring(0, 50)}${trimmedLine.length > 50 ? '...' : ''}' xuất hiện trước câu hỏi (__QUESTION__) đầu tiên mà không phải marker nhóm.`
      });
      continue;
    }

    // Dòng nối tiếp của đáp án hoặc câu hỏi đang mở
    if (currentAnswer) {
      currentAnswer.text = currentAnswer.text ? `${currentAnswer.text}\n${trimmedLine}` : trimmedLine;
    } else if (currentQuestion) {
      currentQuestion.text = currentQuestion.text ? `${currentQuestion.text}\n${trimmedLine}` : trimmedLine;
    }
  }

  // Kiểm tra tính toàn vẹn của từng câu hỏi
  let allQuestions = [];
  let choiceCount = 0;
  let tfCount = 0;
  let shortCount = 0;

  for (const group of groups) {
    for (const q of group.questions) {
      allQuestions.push(q);
      const totalAnswers = q.answers.length;
      const correctAnswers = q.answers.filter(a => a.correct);
      const wrongAnswers = q.answers.filter(a => !a.correct);

      // Câu không có đáp án nào
      if (totalAnswers === 0) {
        errors.push({
          line: q.startLine,
          message: `Câu ${q.id} (dòng ${q.startLine}) chưa có đáp án nào. Hãy thêm các phương án đáp án cho câu hỏi.`
        });
        continue;
      }

      // Câu không có đáp án đúng (__D__)
      if (correctAnswers.length === 0) {
        errors.push({
          line: q.startLine,
          message: `Câu ${q.id} (dòng ${q.startLine}) chưa có đáp án đúng. Hãy thêm __D__ trước đáp án đúng.`
        });
        continue;
      }

      // Phân loại dạng câu hỏi
      if (totalAnswers === 1 && correctAnswers.length === 1) {
        // Dạng trả lời ngắn: 1 đáp án duy nhất là __D__
        q.type = 'short_answer';
        shortCount++;
      } else if (wrongAnswers.length === 0 && totalAnswers > 1) {
        // Có nhiều đáp án nhưng toàn là D mà không có S (hoặc trường hợp trả lời ngắn có S đi kèm)
        q.type = 'true_false';
        tfCount++;
      } else if (correctAnswers.length > 1) {
        // Đúng sai hoặc nhiều đáp án đúng
        q.type = 'true_false';
        tfCount++;
      } else {
        // Lựa chọn đơn truyền thống (1 D, còn lại S)
        q.type = 'multiple_choice';
        choiceCount++;
      }
    }
  }

  // Đảm bảo có ít nhất 1 câu hỏi
  if (allQuestions.length === 0 && errors.length === 0) {
    errors.push({
      line: 1,
      message: 'Không tìm thấy câu hỏi nào trong đề thi. Hãy bắt đầu câu hỏi bằng marker __QUESTION__.'
    });
  }

  return {
    valid: errors.length === 0,
    errors,
    groups,
    stats: {
      totalGroups: groups.length,
      totalQuestions: allQuestions.length,
      choiceCount,
      tfCount,
      shortCount
    }
  };
}


// --- FILE: js/shuffler.js ---
/**
 * Module Hoán Vị Đề Thi (Shuffler)
 * Tuân thủ quy định tại mục 4 của tài liệu:
 * "Công cụ trộn đề trắc nghiệm Yêu cầu và giải pháp.docx"
 *
 * - Thuật toán Fisher-Yates với PRNG có seed (Mulberry32).
 * - Cùng đầu vào và cùng seed cho cùng một kết quả chính xác (deterministic).
 * - Đáp án cố định (#) giữ đúng vị trí ban đầu trong câu.
 * - Nhóm cố định (#) giữ đúng vị trí ban đầu trong đề.
 * - Kiểm tra và loại bỏ trùng lặp mã đề trong cùng một đợt sinh.
 */

/**
 * Bộ sinh số giả ngẫu nhiên Mulberry32 (32-bit PRNG)
 * @param {number} seed Số nguyên 32-bit
 * @returns {Function} Hàm trả về số thực trong khoảng [0, 1)
 */
function createRNG(seed) {
  let s = Math.floor(seed) >>> 0;
  return function() {
    let t = (s += 0x6D2B79F5) >>> 0;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/**
 * Tạo số seed nguyên từ chuỗi (mã đề + chuỗi muối)
 */
function hashStringToSeed(str) {
  let hash = 2166136261;
  for (let i = 0; i < str.length; i++) {
    hash ^= str.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return hash >>> 0;
}

/**
 * Hoán vị một mảng giữ nguyên vị trí của các phần tử có cờ fixed === true
 * @param {Array} array
 * @param {Function} rng
 * @returns {Array} Mảng mới đã được hoán vị
 */
function shufflePreservingFixed(array, rng) {
  const n = array.length;
  if (n <= 1) return [...array];

  const result = new Array(n);
  const nonFixedItems = [];
  const nonFixedIndices = [];

  for (let i = 0; i < n; i++) {
    if (array[i].fixed) {
      result[i] = array[i];
    } else {
      nonFixedItems.push(array[i]);
      nonFixedIndices.push(i);
    }
  }

  // Fisher-Yates trên danh sách không cố định
  for (let i = nonFixedItems.length - 1; i > 0; i--) {
    const j = Math.floor(rng() * (i + 1));
    const temp = nonFixedItems[i];
    nonFixedItems[i] = nonFixedItems[j];
    nonFixedItems[j] = temp;
  }

  // Đặt lại vào các vị trí trống ban đầu
  for (let k = 0; k < nonFixedIndices.length; k++) {
    result[nonFixedIndices[k]] = nonFixedItems[k];
  }

  return result;
}

/**
 * Sinh nhãn hiển thị cho đáp án theo kiểu đã chọn
 * @param {number} index Thứ tự hiển thị (0 -> A, 1 -> B...)
 * @param {string} style 'ABCD' | 'abcd' | '1234'
 */
function getAnswerLabel(index, style = 'ABCD') {
  if (style === 'abcd') {
    const char = String.fromCharCode(97 + (index % 26)); // a, b, c, d...
    return {
      raw: char,
      display: `${char})`
    };
  }

  if (style === '1234') {
    const num = index + 1;
    return {
      raw: String(num),
      display: `${num}.`
    };
  }

  // Mặc định: 'ABCD' -> A., B., C., D.
  const char = String.fromCharCode(65 + (index % 26)); // A, B, C, D...
  return {
    raw: char,
    display: `${char}.`
  };
}

/**
 * Hoán vị một đề thi với mã đề và seed cụ thể
 * @param {Object} parsedExam Kết quả từ parser
 * @param {string} examCode Mã đề (ví dụ '101')
 * @param {Object} options Tùy chọn cấu hình trộn
 * @returns {Object} Cấu trúc đề thi đã hoán vị kèm nhãn
 */
function shuffleExamForCode(parsedExam, examCode, options = {}) {
  const {
    shuffleGroups = false,
    defaultLabelStyle = 'ABCD',
    numberingFormat = 'dot', // 'dot' (Câu 1.) hoặc 'colon' (Câu 1:)
    startQuestionNumber = 1,
    baseSeed = 1337
  } = options;

  const seed = hashStringToSeed(`${examCode}-${baseSeed}`);
  const rng = createRNG(seed);

  // 1. Hoán vị các nhóm (nếu bật shuffleGroups)
  let groupsToProcess = parsedExam.groups.map(g => ({
    ...g,
    questions: [...g.questions]
  }));

  if (shuffleGroups && groupsToProcess.length > 1) {
    groupsToProcess = shufflePreservingFixed(groupsToProcess, rng);
  }

  let globalQuestionIndex = 0;
  const shuffledGroups = [];
  const flatQuestions = [];

  for (const group of groupsToProcess) {
    const mode = group.mode || 'G';
    const labelStyle = group.labelStyle || defaultLabelStyle;

    // Xác định cờ hoán vị theo chế độ nhóm:
    // G0: không đổi gì
    // GQ: đổi câu hỏi, giữ đáp án
    // GA: giữ câu hỏi, đổi đáp án
    // G: đổi cả hai
    const canShuffleQuestions = (mode === 'G' || mode === 'GQ');
    const canShuffleAnswers = (mode === 'G' || mode === 'GA');

    let groupQuestions = group.questions.map(q => ({
      ...q,
      answers: [...q.answers]
    }));

    if (canShuffleQuestions && groupQuestions.length > 1) {
      groupQuestions = shufflePreservingFixed(groupQuestions, rng);
    }

    // Hoán vị đáp án bên trong từng câu
    const finalGroupQuestions = [];
    for (const q of groupQuestions) {
      globalQuestionIndex++;
      const currentQNumber = startQuestionNumber + globalQuestionIndex - 1;
      const qPrefix = numberingFormat === 'colon' ? `Câu ${currentQNumber}:` : `Câu ${currentQNumber}.`;

      let processedAnswers = [...q.answers];

      // Nếu câu có nhiều hơn 1 đáp án và nhóm cho phép đảo đáp án
      if (canShuffleAnswers && processedAnswers.length > 1 && q.type !== 'short_answer') {
        processedAnswers = shufflePreservingFixed(processedAnswers, rng);
      }

      // Gán nhãn đáp án theo thứ tự hiển thị sau khi đã hoán vị
      const answersWithLabels = processedAnswers.map((ans, idx) => {
        const label = getAnswerLabel(idx, labelStyle);
        return {
          ...ans,
          displayIndex: idx,
          labelRaw: label.raw,
          labelDisplay: label.display
        };
      });

      const processedQuestion = {
        ...q,
        displayNumber: currentQNumber,
        displayPrefix: qPrefix,
        labelStyle,
        answers: answersWithLabels
      };

      finalGroupQuestions.push(processedQuestion);
      flatQuestions.push(processedQuestion);
    }

    shuffledGroups.push({
      ...group,
      labelStyle,
      questions: finalGroupQuestions
    });
  }

  return {
    examCode,
    groups: shuffledGroups,
    flatQuestions,
    totalQuestions: flatQuestions.length
  };
}

/**
 * Sinh dấu vân tay (fingerprint) của đề thi để kiểm tra trùng lặp
 */
function getExamFingerprint(shuffledExam) {
  return shuffledExam.flatQuestions.map(q => {
    const ansOrder = q.answers.map(a => `${a.id}:${a.correct ? 'D' : 'S'}`).join(',');
    return `Q${q.id}[${ansOrder}]`;
  }).join('|');
}

/**
 * Sinh toàn bộ các mã đề thi và đảm bảo không có 2 mã đề nào bị trùng nhau
 * @param {Object} parsedExam
 * @param {Array<string>} examCodes
 * @param {Object} options
 * @returns {Array<Object>} Danh sách các đề thi đã hoán vị
 */
function shuffleAllExams(parsedExam, examCodes, options = {}) {
  const result = [];
  const seenFingerprints = new Set();

  for (let i = 0; i < examCodes.length; i++) {
    const code = String(examCodes[i]).trim();
    let currentSeedOffset = 0;
    let shuffled;
    let fingerprint;

    do {
      shuffled = shuffleExamForCode(parsedExam, code, {
        ...options,
        baseSeed: (options.baseSeed || 1337) + currentSeedOffset
      });
      fingerprint = getExamFingerprint(shuffled);
      currentSeedOffset += 997; // Tăng seed nếu bị trùng
    } while (seenFingerprints.has(fingerprint) && currentSeedOffset < 100000);

    seenFingerprints.add(fingerprint);
    result.push(shuffled);
  }

  return result;
}


// --- FILE: js/docx-builder.js ---
/**
 * Module Dựng File Word Đề Thi (Docx Builder)
 * Tuân thủ quy định tại mục 5.1 và 5.2 của tài liệu:
 * "Công cụ trộn đề trắc nghiệm Yêu cầu và giải pháp.docx"
 *
 * - Font: Times New Roman toàn bộ văn bản.
 * - Đầu đề: Bảng 2 cột không viền căn giữa, phối màu đỏ/đen đậm chuẩn mực theo mẫu duyệt.
 * - Dòng "Họ và tên...", "Số báo danh...", "Mã đề [101]" kèm vạch kẻ ngang.
 * - Đánh số câu liên tục theo lựa chọn người dùng ("Câu 1." hoặc "Câu 1:").
 * - Nhãn đáp án theo định dạng đã chọn (A. B. C. D. / a) b) c) d) / 1. 2. 3. 4.).
 * - Câu trả lời ngắn (chỉ có __D__) không hiển thị đáp án trong đề thi.
 */

function getDocx() {
  if (typeof window !== 'undefined' && window.docx) {
    return window.docx;
  }
  if (typeof global !== 'undefined' && global.docx) {
    return global.docx;
  }
  return null;
}

/**
 * Tạo Blob file Word (.docx) cho một mã đề thi cụ thể
 * @param {Object} shuffledExam Đề thi đã hoán vị (từ shuffler)
 * @param {Object} headerInfo Thông tin đầu đề (7 ô nhập)
 * @returns {Promise<Blob>}
 */
async function createExamDocxBlob(shuffledExam, headerInfo = {}) {
  const docx = getDocx();
  if (!docx) {
    throw new Error('Thư viện docx chưa được nạp!');
  }

  const {
    Document,
    Paragraph,
    TextRun,
    Table,
    TableRow,
    TableCell,
    WidthType,
    BorderStyle,
    AlignmentType,
    PageNumber,
    Packer
  } = docx;

  const {
    soGD = 'SỞ GD&ĐT TP HCM',
    truong = 'TRƯỜNG THPT NGUYỄN HỮU HUÂN',
    kiKiemTra = 'GIỮA HK 1',
    namHoc = '2026 - 2027',
    mon = 'VẬT LÝ',
    thoiGian = '45 phút'
  } = headerInfo;

  const examCode = shuffledExam.examCode;
  const FONT_FAMILY = 'Times New Roman';
  const COLOR_RED = 'FF0000';
  const COLOR_BLACK = '000000';

  // 1. Bảng đầu đề thi (1 hàng, 2 cột không viền)
  const headerTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.NONE },
      bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE }
    },
    rows: [
      new TableRow({
        children: [
          // Cột trái: Sở GD&ĐT, Tên trường (đỏ, đậm); vạch phân cách; "(Đề thi có ___ trang)"
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE }
            },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 30 },
                children: [
                  new TextRun({
                    text: (soGD || '').toUpperCase(),
                    bold: true,
                    color: COLOR_RED,
                    font: FONT_FAMILY,
                    size: 24 // 12pt
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 30 },
                children: [
                  new TextRun({
                    text: (truong || '').toUpperCase(),
                    bold: true,
                    color: COLOR_RED,
                    font: FONT_FAMILY,
                    size: 24
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 30 },
                children: [
                  new TextRun({
                    text: '--------------------',
                    font: FONT_FAMILY,
                    size: 20
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 60 },
                children: [
                  new TextRun({
                    text: '(Đề thi có ',
                    italics: true,
                    font: FONT_FAMILY,
                    size: 22 // 11pt
                  }),
                  new TextRun({
                    children: [PageNumber.TOTAL_PAGES],
                    italics: true,
                    font: FONT_FAMILY,
                    size: 22
                  }),
                  new TextRun({
                    text: ' trang)',
                    italics: true,
                    font: FONT_FAMILY,
                    size: 22
                  })
                ]
              })
            ]
          }),

          // Cột phải: Kì kiểm tra (đỏ, đậm); Năm học; Môn; Thời gian làm bài; "(không kể thời gian phát đề)"
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE }
            },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 30 },
                children: [
                  new TextRun({
                    text: (kiKiemTra || '').toUpperCase(),
                    bold: true,
                    color: COLOR_RED,
                    font: FONT_FAMILY,
                    size: 24
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 30 },
                children: [
                  new TextRun({
                    text: 'NĂM HỌC ',
                    bold: true,
                    color: COLOR_BLACK,
                    font: FONT_FAMILY,
                    size: 22
                  }),
                  new TextRun({
                    text: (namHoc || '').trim(),
                    bold: true,
                    color: COLOR_RED,
                    font: FONT_FAMILY,
                    size: 22
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 30 },
                children: [
                  new TextRun({
                    text: 'MÔN: ',
                    bold: true,
                    color: COLOR_BLACK,
                    font: FONT_FAMILY,
                    size: 24
                  }),
                  new TextRun({
                    text: (mon || '').toUpperCase(),
                    bold: true,
                    color: COLOR_RED,
                    font: FONT_FAMILY,
                    size: 24
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 20 },
                children: [
                  new TextRun({
                    text: 'Thời gian làm bài: ',
                    italics: true,
                    color: COLOR_BLACK,
                    font: FONT_FAMILY,
                    size: 22
                  }),
                  new TextRun({
                    text: (thoiGian || '').trim(),
                    bold: true,
                    color: COLOR_RED,
                    font: FONT_FAMILY,
                    size: 22
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 60 },
                children: [
                  new TextRun({
                    text: '(không kể thời gian phát đề)',
                    italics: true,
                    color: COLOR_BLACK,
                    font: FONT_FAMILY,
                    size: 20
                  })
                ]
              })
            ]
          })
        ]
      })
    ]
  });

  // 2. Dòng Họ tên, Số báo danh, Mã đề kèm vạch kẻ ngang
  const candidateInfoParagraph = new Paragraph({
    spacing: { before: 100, after: 120 },
    border: {
      bottom: {
        color: '333333',
        space: 6,
        style: BorderStyle.SINGLE,
        size: 10
      }
    },
    children: [
      new TextRun({
        text: 'Họ và tên: ................................................................  ',
        font: FONT_FAMILY,
        size: 22
      }),
      new TextRun({
        text: 'Số báo danh: .......................  ',
        font: FONT_FAMILY,
        size: 22
      }),
      new TextRun({
        text: 'Mã đề ',
        bold: true,
        font: FONT_FAMILY,
        size: 24
      }),
      new TextRun({
        text: String(examCode),
        bold: true,
        color: COLOR_RED,
        font: FONT_FAMILY,
        size: 26
      })
    ]
  });

  // 3. Nội dung phần thân đề (Danh sách câu hỏi & đáp án)
  const contentParagraphs = [];

  for (const q of shuffledExam.flatQuestions) {
    // Đoạn câu hỏi
    const questionRuns = [
      new TextRun({
        text: `${q.displayPrefix} `,
        bold: true,
        font: FONT_FAMILY,
        size: 24
      })
    ];

    // Xử lý nội dung câu hỏi (có thể có nhiều dòng)
    const qLines = (q.text || '').split('\n');
    qLines.forEach((line, lIdx) => {
      if (lIdx > 0) {
        questionRuns.push(new TextRun({ break: 1 }));
      }
      questionRuns.push(new TextRun({
        text: line,
        font: FONT_FAMILY,
        size: 24
      }));
    });

    contentParagraphs.push(new Paragraph({
      spacing: { before: 120, after: 60 },
      children: questionRuns
    }));

    // Đoạn đáp án (Câu trả lời ngắn không hiển thị đáp án trong đề thi theo mục 5.2)
    if (q.type !== 'short_answer') {
      for (const ans of q.answers) {
        const answerRuns = [
          new TextRun({
            text: `${ans.labelDisplay} `,
            bold: true,
            font: FONT_FAMILY,
            size: 24
          })
        ];

        const aLines = (ans.text || '').split('\n');
        aLines.forEach((aLine, aIdx) => {
          if (aIdx > 0) {
            answerRuns.push(new TextRun({ break: 1 }));
          }
          answerRuns.push(new TextRun({
            text: aLine,
            font: FONT_FAMILY,
            size: 24
          }));
        });

        contentParagraphs.push(new Paragraph({
          indent: { left: 400 }, // Thụt đầu dòng cho các phương án đáp án
          spacing: { before: 30, after: 30 },
          children: answerRuns
        }));
      }
    }
  }

  // Đoạn kết thúc đề
  contentParagraphs.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 240, after: 100 },
    children: [
      new TextRun({
        text: '----------------------- HẾT -----------------------',
        bold: true,
        font: FONT_FAMILY,
        size: 22
      })
    ]
  }));

  // Tạo tài liệu Word
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: FONT_FAMILY,
            size: 24
          }
        }
      }
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 1000,
              bottom: 1000,
              left: 1200,
              right: 1200
            }
          }
        },
        children: [
          headerTable,
          candidateInfoParagraph,
          ...contentParagraphs
        ]
      }
    ]
  });

  return await Packer.toBlob(doc);
}


// --- FILE: js/answer-key.js ---
/**
 * Module Tạo Bảng Đáp Án (Answer Key)
 * Tuân thủ quy định tại mục 5.4 của tài liệu:
 * "Công cụ trộn đề trắc nghiệm Yêu cầu và giải pháp.docx"
 *
 * - Tự suy ra từ dữ liệu:
 *   + Một __D__: ghi nhãn (ví dụ C).
 *   + Nhiều __D__: ghi tập hợp (ví dụ A, C); nếu là nhãn a) b) c) d) ghi từng ý (ví dụ a-Đ, b-S, c-Đ, d-S).
 *   + Câu trả lời ngắn (chỉ có __D__): ghi nguyên giá trị chuỗi đáp án (ví dụ 4).
 * - Bảng ma trận: mỗi hàng là một câu, mỗi cột là một mã đề thi.
 * - Xuất thành một file .docx riêng biệt.
 */

function getDocx() {
  if (typeof window !== 'undefined' && window.docx) {
    return window.docx;
  }
  if (typeof global !== 'undefined' && global.docx) {
    return global.docx;
  }
  return null;
}

/**
 * Trích xuất giá trị đáp án đúng của một câu hỏi đã hoán vị
 * @param {Object} q Câu hỏi đã hoán vị
 * @returns {string} Chuỗi hiển thị đáp án
 */
function extractQuestionAnswerKey(q) {
  const correctAnswers = q.answers.filter(a => a.correct);

  // 1. Câu trả lời ngắn (chỉ có 1 __D__ và không có __S__)
  if (q.type === 'short_answer' || (q.answers.length === 1 && correctAnswers.length === 1)) {
    return correctAnswers[0].text.trim();
  }

  // 2. Định dạng nhãn a) b) c) d) cho câu đúng/sai nhiều ý
  if (q.labelStyle === 'abcd') {
    return q.answers.map(ans => `${ans.labelRaw}-${ans.correct ? 'Đ' : 'S'}`).join(', ');
  }

  // 3. Một đáp án đúng duy nhất (__D__)
  if (correctAnswers.length === 1) {
    return correctAnswers[0].labelRaw;
  }

  // 4. Nhiều đáp án đúng (__D__)
  return correctAnswers.map(a => a.labelRaw).join(', ');
}

/**
 * Sinh cấu trúc ma trận bảng đáp án (để xem trước trên UI hoặc xuất DOCX)
 * @param {Array<Object>} allShuffledExams Danh sách tất cả các mã đề thi đã hoán vị
 * @returns {Object} { headers: Array<string>, rows: Array<Object> }
 */
function buildAnswerMatrix(allShuffledExams) {
  if (!allShuffledExams || allShuffledExams.length === 0) {
    return { headers: [], rows: [] };
  }

  const examCodes = allShuffledExams.map(e => e.examCode);
  const totalQuestions = allShuffledExams[0].flatQuestions.length;
  const rows = [];

  for (let i = 0; i < totalQuestions; i++) {
    const qNum = allShuffledExams[0].flatQuestions[i].displayNumber || (i + 1);
    const row = {
      questionNumber: qNum,
      answers: {}
    };

    allShuffledExams.forEach(exam => {
      const q = exam.flatQuestions[i];
      row.answers[exam.examCode] = extractQuestionAnswerKey(q);
    });

    rows.push(row);
  }

  return {
    headers: examCodes,
    rows
  };
}

/**
 * Tạo Blob file Word (.docx) chứa Bảng Đáp Án Ma Trận
 * @param {Array<Object>} allShuffledExams
 * @param {Object} headerInfo
 * @returns {Promise<Blob>}
 */
async function createAnswerKeyDocxBlob(allShuffledExams, headerInfo = {}) {
  const docx = getDocx();
  if (!docx) {
    throw new Error('Thư viện docx chưa được nạp!');
  }

  const {
    Document,
    Paragraph,
    TextRun,
    Table,
    TableRow,
    TableCell,
    WidthType,
    BorderStyle,
    AlignmentType,
    ShadingType,
    Packer
  } = docx;

  const {
    soGD = 'SỞ GD&ĐT TP HCM',
    truong = 'TRƯỜNG THPT NGUYỄN HỮU HUÂN',
    kiKiemTra = 'GIỮA HK 1',
    namHoc = '2026 - 2027',
    mon = 'VẬT LÝ'
  } = headerInfo;

  const FONT_FAMILY = 'Times New Roman';
  const matrix = buildAnswerMatrix(allShuffledExams);

  // Tiêu đề trang đáp án
  const titleParagraphs = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 40 },
      children: [
        new TextRun({
          text: `${(soGD || '').toUpperCase()} - ${(truong || '').toUpperCase()}`,
          bold: true,
          font: FONT_FAMILY,
          size: 22
        })
      ]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 40 },
      children: [
        new TextRun({
          text: `BẢNG ĐÁP ÁN ${kiKiemTra.toUpperCase()} - NĂM HỌC ${namHoc}`,
          bold: true,
          color: 'FF0000',
          font: FONT_FAMILY,
          size: 26
        })
      ]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 140 },
      children: [
        new TextRun({
          text: `MÔN: ${(mon || '').toUpperCase()}`,
          bold: true,
          font: FONT_FAMILY,
          size: 24
        })
      ]
    })
  ];

  // Bảng đáp án
  const tableBorder = {
    color: '555555',
    size: 4,
    style: BorderStyle.SINGLE
  };

  const bordersConfig = {
    top: tableBorder,
    bottom: tableBorder,
    left: tableBorder,
    right: tableBorder,
    insideHorizontal: tableBorder,
    insideVertical: tableBorder
  };

  // Hàng tiêu đề bảng: Câu | Mã 101 | Mã 102 | ...
  const headerCells = [
    new TableCell({
      width: { size: 15, type: WidthType.PERCENTAGE },
      shading: { type: ShadingType.CLEAR, fill: 'EEEEEE' },
      borders: bordersConfig,
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text: 'Câu',
              bold: true,
              font: FONT_FAMILY,
              size: 22
            })
          ]
        })
      ]
    })
  ];

  const colWidthPercent = 85 / (matrix.headers.length || 1);

  matrix.headers.forEach(code => {
    headerCells.push(
      new TableCell({
        width: { size: colWidthPercent, type: WidthType.PERCENTAGE },
        shading: { type: ShadingType.CLEAR, fill: 'EEEEEE' },
        borders: bordersConfig,
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: `Mã ${code}`,
                bold: true,
                color: 'FF0000',
                font: FONT_FAMILY,
                size: 22
              })
            ]
          })
        ]
      })
    );
  });

  const tableRows = [
    new TableRow({
      tableHeader: true,
      children: headerCells
    })
  ];

  // Các hàng dữ liệu: từng câu hỏi
  matrix.rows.forEach((row, rIdx) => {
    const isAlt = (rIdx % 2 === 1);
    const fillBg = isAlt ? 'F9FAFB' : 'FFFFFF';

    const rowCells = [
      new TableCell({
        width: { size: 15, type: WidthType.PERCENTAGE },
        shading: { type: ShadingType.CLEAR, fill: fillBg },
        borders: bordersConfig,
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: `Câu ${row.questionNumber}`,
                bold: true,
                font: FONT_FAMILY,
                size: 22
              })
            ]
          })
        ]
      })
    ];

    matrix.headers.forEach(code => {
      rowCells.push(
        new TableCell({
          width: { size: colWidthPercent, type: WidthType.PERCENTAGE },
          shading: { type: ShadingType.CLEAR, fill: fillBg },
          borders: bordersConfig,
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: row.answers[code] || '',
                  bold: true,
                  font: FONT_FAMILY,
                  size: 22
                })
              ]
            })
          ]
        })
      );
    });

    tableRows.push(new TableRow({ children: rowCells }));
  });

  const answerTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: bordersConfig,
    rows: tableRows
  });

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: FONT_FAMILY,
            size: 22
          }
        }
      }
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 1000,
              bottom: 1000,
              left: 1200,
              right: 1200
            }
          }
        },
        children: [
          ...titleParagraphs,
          answerTable
        ]
      }
    ]
  });

  return await Packer.toBlob(doc);
}


// --- FILE: js/ui.js ---
/**
 * Module Giao Diện Người Dùng (UI Controller)
 * Quản lý tương tác nạp file Word (.docx), cấu hình đầu đề, tùy chọn trộn,
 * hiển thị xem trước trực quan và xử lý xuất file DOCX / ZIP.
 */






// Trạng thái ứng dụng
const state = {
  rawText: '',
  fileName: '',
  fileSize: 0,
  parsedExam: null,
  examCodes: ['101', '102', '103', '104'],
  generatedZipBlob: null,
  generatedFiles: [] // { name, blob, url }
};

// Hàm hiển thị thông báo Snackbar
function showSnackbar(message, duration = 3000) {
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
function updateExamCodes(count, startCode) {
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
function renderExamPreview(parsedExam) {
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
            <span>${escapeHtml(q.text)}</span>
          </div>
          <div class="answers-grid">
      `;

      if (q.type === 'short_answer') {
        html += `
          <div class="answer-item">
            <span class="ans-pill d">Trả lời ngắn</span>
            <span class="answer-text"><strong>Đáp án:</strong> ${escapeHtml(q.answers[0].text)}</span>
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
              <span class="answer-text"><strong>${labelTemp}.</strong> ${escapeHtml(ans.text)}</span>
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
async function handleFileLoaded(file) {
  try {
    showSnackbar('Đang đọc và phân tích file Word...');
    const result = await readExamDocx(file);

    state.rawText = result.text;
    state.fileName = result.fileName;
    state.fileSize = result.fileSize;

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
function getHeaderConfig() {
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
function getMixingOptions() {
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
async function executeMixAndExport() {
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
    const mixingOptions = getMixingOptions();

    // 1. Hoán vị tất cả các mã đề
    const shuffledExams = shuffleAllExams(state.parsedExam, state.examCodes, mixingOptions);

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
function downloadAllZip() {
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


// --- FILE: js/app.js ---
/**
 * PWA Mark 'n Mix - Main Application Entry
 * Khởi tạo ứng dụng, lắng nghe sự kiện, đăng ký Service Worker PWA
 */


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
}


})();
