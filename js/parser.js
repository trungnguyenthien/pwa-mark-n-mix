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
  // Bỏ qua đường kẻ ngang hoặc chỗ điền khuyết như '____', '______'
  if (/^_{3,}$/.test(token)) {
    return { isBlankFill: true };
  }

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
export function parseExam(rawText) {
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

    // Tìm tất cả marker có dạng __...__ trên dòng (marker chỉ gồm chữ, số và ký tự #)
    const markerMatches = [...rawLine.matchAll(/__([A-Za-z0-9#]+)__/g)];

    if (markerMatches.length === 0) {
      // Dòng không có marker
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
      continue;
    }

    // Có ít nhất 1 marker trên dòng
    // 1. Nếu có phần văn bản đứng trước marker đầu tiên:
    const firstMatch = markerMatches[0];
    const leadingText = rawLine.substring(0, firstMatch.index).trim();
    if (leadingText) {
      if (currentAnswer) {
        currentAnswer.text = currentAnswer.text ? `${currentAnswer.text}\n${leadingText}` : leadingText;
      } else if (currentQuestion) {
        currentQuestion.text = currentQuestion.text ? `${currentQuestion.text}\n${leadingText}` : leadingText;
      } else if (!hasMetFirstQuestion) {
        errors.push({
          line: lineNumber,
          message: `Dòng ${lineNumber}: Nội dung '${leadingText.substring(0, 50)}${leadingText.length > 50 ? '...' : ''}' xuất hiện trước câu hỏi (__QUESTION__) đầu tiên mà không phải marker nhóm.`
        });
      }
    }

    // 2. Duyệt qua từng marker trên dòng theo thứ tự từ trái sang phải
    for (let k = 0; k < markerMatches.length; k++) {
      const match = markerMatches[k];
      const token = match[0];
      const startIndex = match.index + token.length;
      const endIndex = (k + 1 < markerMatches.length) ? markerMatches[k + 1].index : rawLine.length;
      const content = rawLine.substring(startIndex, endIndex).trim();

      const parsedToken = parseMarkerToken(token);

      if (parsedToken.isBlankFill) {
        continue;
      }

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
