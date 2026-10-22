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
export function createRNG(seed) {
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
export function hashStringToSeed(str) {
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
export function shufflePreservingFixed(array, rng) {
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
export function getAnswerLabel(index, style = 'ABCD') {
  if (style === 'ABCD_paren') {
    const char = String.fromCharCode(65 + (index % 26)); // A, B, C, D...
    return {
      raw: char,
      display: `${char})`
    };
  }

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
 * Danh sách 20 chữ số La Mã cho tối đa 20 phần
 */
export const ROMAN_NUMERALS = [
  'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X',
  'XI', 'XII', 'XIII', 'XIV', 'XV', 'XVI', 'XVII', 'XVIII', 'XIX', 'XX'
];

/**
 * Chuyển đổi số thứ tự (1-based) sang số La Mã (I, II, ..., XX)
 * @param {number} n
 * @returns {string}
 */
export function toRoman(n) {
  if (n >= 1 && n <= 20) return ROMAN_NUMERALS[n - 1];
  return String(n);
}

/**
 * Nhận diện dòng tiêu đề là IN HOA toàn bộ hay viết bình thường
 * @param {string} text
 * @returns {boolean} true nếu là IN HOA toàn bộ, false nếu viết bình thường
 */
export function isTitleUpperCase(text) {
  if (!text) return true;

  // 1. Nếu có tiền tố PHẦN / Phần sẵn có, ưu tiên theo case của tiền tố đó
  const prefixMatch = text.match(/(?:^|\s|<[^>]+>)(PHẦN|Phần|phần)\b/);
  if (prefixMatch) {
    return prefixMatch[1] === prefixMatch[1].toUpperCase();
  }

  // 2. Lấy nội dung chữ thuần (loại bỏ toàn bộ thẻ HTML và token ảnh)
  const clean = text.replace(/<[^>]+>/g, '').replace(/\[IMG:[^\]]+\]/g, '').trim();
  if (!clean) return true;

  // 3. Trích xuất tất cả chữ cái (hỗ trợ cả tiếng Việt có dấu)
  const allLetters = clean.replace(/[^a-zA-Zà-ỹÀ-Ỹ]/g, '');
  if (!allLetters) return true;
  if (allLetters === allLetters.toUpperCase()) return true;

  // 4. Nếu toàn bộ dòng có lẫn chữ thường (ví dụ: "CÂU TRẮC NGHIỆM: Thí sinh chọn..."),
  // kiểm tra cụm từ tiêu đề chính (trước dấu hai chấm hoặc dấu gạch ngang)
  const headingPart = clean.split(/[:\-\–\—]/)[0].trim();
  const headingLetters = headingPart.replace(/[^a-zA-Zà-ỹÀ-Ỹ]/g, '');
  if (headingLetters && headingLetters.length >= 3 && headingLetters === headingLetters.toUpperCase()) {
    return true;
  }

  return false;
}

/**
 * Định dạng dòng đầu tiên của tiêu đề nhóm, tự động thêm hoặc cập nhật tiền tố "PHẦN I." / "Phần I."
 * Tự động nhận diện chính xác kiểu chữ (IN HOA vs Viết thường) và giữ nguyên các định dạng (bold <b>, italic <i>, underline <u>).
 *
 * @param {string} line Dòng đầu tiên của tiêu đề nhóm
 * @param {string} roman Chữ số La Mã (I, II, III, ..., XX)
 * @returns {string} Dòng tiêu đề đã được thêm hoặc cập nhật tiền tố
 */
export function formatGroupTitleFirstLine(line, roman) {
  if (!line || !line.trim()) return `PHẦN ${roman}.`;
  const trimmed = line.trim();

  // Xác định tiền tố là "PHẦN" (IN HOA) hay "Phần" (Viết bình thường)
  const isUpper = isTitleUpperCase(trimmed);
  const prefixWord = isUpper ? 'PHẦN' : 'Phần';
  const targetPrefix = `${prefixWord} ${roman}.`;

  // TH1: Tiền tố PHẦN nằm độc lập trong các thẻ định dạng riêng ở đầu dòng
  // Ví dụ: <b>PHẦN I.</b> Tiêu đề... hoặc <b><i>PHẦN I:</i></b> ... hoặc <u><b>PHẦN 1</b></u>: ...
  const isolatedPrefixRegex = /^((?:<[a-z0-9]+[^>]*>\s*)+)\s*(?:PHẦN|Phần|phần)\s+(?:[IVXLCDM]+|[0-9]+)[\.\:\-]?\s*((?:<\/[a-z0-9]+>\s*)+)(?:[\.\:\-]?\s*)(.*)$/i;
  const isolatedMatch = trimmed.match(isolatedPrefixRegex);
  if (isolatedMatch) {
    const openTags = isolatedMatch[1].trim();
    const closeTags = isolatedMatch[2].trim();
    const rest = isolatedMatch[3].trim();
    return rest
      ? `${openTags}${targetPrefix}${closeTags} ${rest}`
      : `${openTags}${targetPrefix}${closeTags}`;
  }

  // TH2: Tiền tố PHẦN nằm chung thẻ với nội dung phía sau
  // Ví dụ: <b>PHẦN I. CÂU TRẮC NGHIỆM...</b> hoặc <b><i>PHẦN I: CÂU TRẮC NGHIỆM...</i></b>
  const sharedPrefixRegex = /^((?:<[a-z0-9]+[^>]*>\s*)+)\s*(?:PHẦN|Phần|phần)\s+(?:[IVXLCDM]+|[0-9]+)[\.\:\-]?\s*(.*)$/i;
  const sharedMatch = trimmed.match(sharedPrefixRegex);
  if (sharedMatch) {
    const openTags = sharedMatch[1];
    const rest = sharedMatch[2].trim();
    return rest
      ? `${openTags}${targetPrefix} ${rest}`
      : `${openTags}${targetPrefix}`;
  }

  // TH3: Tiền tố PHẦN dạng văn bản thuần ở đầu dòng (không có thẻ bao quanh tiền tố)
  // Ví dụ: 'PHẦN I. CÂU TRẮC NGHIỆM...' hoặc 'Phần 1: Câu trắc nghiệm...'
  const plainPrefixRegex = /^(?:PHẦN|Phần|phần)\s+(?:[IVXLCDM]+|[0-9]+)[\.\:\-]?\s*(.*)$/i;
  const plainMatch = trimmed.match(plainPrefixRegex);
  if (plainMatch) {
    const rest = plainMatch[1].trim();
    return rest ? `${targetPrefix} ${rest}` : targetPrefix;
  }

  // TH4: Trong nội dung gốc CHƯA CÓ "PHẦN I." -> Tự động chèn thêm
  // 4a. Dòng bắt đầu bằng các thẻ định dạng (bold <b>, italic <i>, underline <u> hoặc lồng nhau)
  // Ví dụ: <b>CÂU TRẮC NGHIỆM...</b> -> <b>PHẦN I. CÂU TRẮC NGHIỆM...</b>
  // Ví dụ: <b><i><u>CÂU HỎI...</u></i></b> -> <b><i><u>PHẦN I. CÂU HỎI...</u></i></b>
  const openTagsMatch = trimmed.match(/^((?:<[a-z0-9]+[^>]*>\s*)+)(.*)$/i);
  if (openTagsMatch) {
    const openTags = openTagsMatch[1];
    const rest = openTagsMatch[2].trim();
    return rest
      ? `${openTags}${targetPrefix} ${rest}`
      : `${openTags}${targetPrefix}`;
  }

  // 4b. Dòng văn bản thuần không có thẻ ở đầu
  // Ví dụ: 'CÂU TRẮC NGHIỆM...' -> 'PHẦN I. CÂU TRẮC NGHIỆM...'
  // Ví dụ: 'Câu hỏi trắc nghiệm...' -> 'Phần I. Câu hỏi trắc nghiệm...'
  return `${targetPrefix} ${trimmed}`;
}

/**
 * Định dạng toàn bộ tiêu đề nhóm, gắn tiền tố "PHẦN {roman}." vào dòng đầu tiên
 * @param {string} title Nội dung tiêu đề nhóm
 * @param {number} groupIndex Thứ tự nhóm (1-based: 1, 2, 3...)
 * @returns {string} Tiêu đề đã được định dạng
 */
export function formatGroupTitle(title, groupIndex) {
  if (!title || !title.trim()) return '';
  const roman = toRoman(groupIndex);
  const lines = title.trim().split(/\r?\n/);
  lines[0] = formatGroupTitleFirstLine(lines[0], roman);
  return lines.join('\n');
}

/**
 * Hoán vị một đề thi với mã đề và seed cụ thể
 * @param {Object} parsedExam Kết quả từ parser
 * @param {string} examCode Mã đề (ví dụ '101')
 * @param {Object} options Tùy chọn cấu hình trộn
 * @returns {Object} Cấu trúc đề thi đã hoán vị kèm nhãn
 */
export function shuffleExamForCode(parsedExam, examCode, options = {}) {
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
      const qPrefix = numberingFormat === 'colon'
        ? `Câu ${currentQNumber}:`
        : (numberingFormat === 'paren' ? `Câu ${currentQNumber})` : `Câu ${currentQNumber}.`);

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

    let groupTitle = group.title;
    if (groupTitle && groupTitle.trim()) {
      groupTitle = formatGroupTitle(groupTitle, shuffledGroups.length + 1);
    }

    shuffledGroups.push({
      ...group,
      title: groupTitle,
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
export function shuffleAllExams(parsedExam, examCodes, options = {}) {
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
