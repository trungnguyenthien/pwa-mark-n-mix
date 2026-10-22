/**
 * Module Hoán vị Đề thi (Shuffler)
 * Tuân thủ quy định tại mục 4 của tài liệu:
 * "Công cụ trộn đề trắc nghiệm Yêu cầu và giải pháp.docx"
 *
 * - Thuật toán Fisher-Yates với PRNG có seed (Mulberry32).
 * - Cùng đầu vào và cùng seed cho cùng một kết quả chính xác.
 * - Đáp án cố định (#) giữ đúng vị trí ban đầu trong câu.
 * - Nhóm cố định (#) giữ đúng vị trí ban đầu trong đề.
 * - Kiểm tra trùng lặp mã đề trong cùng một đợt sinh.
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
 * Tạo số seed nguyên từ chuỗi (chuỗi mã đề + base seed)
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
 * @param {number} index Thứ tự hiển thị (0-indexed: 0 -> A, 1 -> B...)
 * @param {string} style 'ABCD' | 'abcd' | '1234'
 */
export function getAnswerLabel(index, style = 'ABCD') {
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

  // Mặc định 'ABCD'
  const char = String.fromCharCode(65 + (index % 26)); // A, B, C, D...
  return {
    raw: char,
    display: `${char}.`
  };
}

/**
 * Tạo chữ ký (signature) cho một đề thi đã hoán vị để kiểm tra trùng lặp
 */
function getExamSignature(shuffledExam) {
  const parts = [];
  for (const group of shuffledExam.groups) {
    parts.push(`G:${group.id}`);
    for (const q of group.questions) {
      parts.push(`Q:${q.originalId || q.id}`);
      for (const a of q.answers) {
        parts.push(`A:${a.id}`);
      }
    }
  }
  return parts.join('|');
}

/**
 * Hoán vị một đề thi cho một mã đề cụ thể
 * @param {Object} parsedExam Kết quả từ parseExam
 * @param {string|number} examCode Mã đề (ví dụ "101")
 * @param {Object} options Tùy chọn hoán vị
 * @returns {Object} Đề thi đã hoán vị cho mã đề này
 */
export function shuffleExam(parsedExam, examCode, options = {}) {
  const {
    baseSeed = 20261022,
    shuffleGroups = false,
    salt = 0,
    groupLabelStyles = {},
    questionStartNum = 1,
    questionNumbering = 'dot' // 'dot' -> "Câu 1.", 'colon' -> "Câu 1:"
  } = options;

  const combinedSeed = (hashStringToSeed(`${baseSeed}_${examCode}`) + salt) >>> 0;
  const rng = createRNG(combinedSeed);

  // 1. Sao chép và chuẩn bị nhóm
  let groups = parsedExam.groups.map(g => ({
    ...g,
    labelStyle: groupLabelStyles[g.id] || g.labelStyle || 'ABCD',
    questions: g.questions.map(q => ({
      ...q,
      originalId: q.id,
      answers: q.answers.map(a => ({ ...a }))
    }))
  }));

  // 2. Hoán vị thứ tự các nhóm (nếu được bật)
  if (shuffleGroups && groups.length > 1) {
    groups = shufflePreservingFixed(groups, rng);
  }

  // 3. Hoán vị câu hỏi và đáp án trong từng nhóm
  let currentQNum = questionStartNum;
  const allOrderedQuestions = [];

  for (const group of groups) {
    const mode = group.mode; // 'G' | 'GQ' | 'GA' | 'G0'
    const shouldShuffleQuestions = mode === 'G' || mode === 'GQ';
    const shouldShuffleAnswers = mode === 'G' || mode === 'GA';

    // Hoán vị câu hỏi trong nhóm
    if (shouldShuffleQuestions && group.questions.length > 1) {
      group.questions = shufflePreservingFixed(group.questions, rng);
    }

    // Hoán vị đáp án trong từng câu
    for (const question of group.questions) {
      // Câu trả lời ngắn (chỉ có 1 __D__) không hoán vị đáp án
      if (shouldShuffleAnswers && question.answers.length > 1 && question.type !== 'short') {
        question.answers = shufflePreservingFixed(question.answers, rng);
      }

      // Đánh số câu hỏi liên tục
      question.displayNumber = currentQNum;
      question.displayPrefix = questionNumbering === 'colon'
        ? `Câu ${currentQNum}:`
        : `Câu ${currentQNum}.`;
      currentQNum++;

      // Gán nhãn đáp án sau khi hoán vị theo thứ tự hiển thị
      const labelStyle = group.labelStyle || 'ABCD';
      question.answers.forEach((ans, idx) => {
        const labelObj = getAnswerLabel(idx, labelStyle);
        ans.displayIndex = idx;
        ans.label = labelObj.display;
        ans.rawLabel = labelObj.raw;
      });

      allOrderedQuestions.push(question);
    }
  }

  const result = {
    examCode: String(examCode),
    groups,
    questions: allOrderedQuestions,
    seed: combinedSeed,
    signature: ''
  };

  result.signature = getExamSignature(result);
  return result;
}

/**
 * Sinh hàng loạt đề thi cho danh sách các mã đề
 * Tự động kiểm tra và sinh lại nếu phát hiện hai mã đề trùng hoàn toàn
 * @param {Object} parsedExam
 * @param {Array<string|number>} examCodes Danh sách mã đề (ví dụ ['101', '102', '103', '104'])
 * @param {Object} options
 * @returns {Array<Object>} Danh sách các đề thi đã hoán vị
 */
export function shuffleExamBatch(parsedExam, examCodes, options = {}) {
  const existingSignatures = new Set();
  const results = [];

  for (const code of examCodes) {
    let salt = 0;
    let candidate = null;
    let attempts = 0;

    // Thử sinh cho đến khi signature không trùng với các mã trước
    do {
      candidate = shuffleExam(parsedExam, code, { ...options, salt });
      attempts++;
      salt += 1337;
    } while (existingSignatures.has(candidate.signature) && attempts < 20);

    existingSignatures.add(candidate.signature);
    results.push(candidate);
  }

  return results;
}
