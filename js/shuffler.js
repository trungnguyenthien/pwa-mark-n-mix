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
