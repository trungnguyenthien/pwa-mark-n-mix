/**
 * Module phân tích cú pháp đề thi (Parser)
 * Tuân thủ quy định tại mục 3 và mục 7 của tài liệu:
 * "Công cụ trộn đề trắc nghiệm Yêu cầu và giải pháp.docx"
 */

// Danh sách marker nhóm hợp lệ
const GROUP_MARKER_REGEX = /^__(G|G0|GQ|GA)(#?)__$/;
const ALT_GROUP_MARKER_REGEX = /^__(#?)(G|G0|GQ|GA)__$/;

// Danh sách marker đáp án hợp lệ
const ANSWER_MARKER_REGEX = /^__(D|S)(#?)__$/;
const ALT_ANSWER_MARKER_REGEX = /^__(#?)(D|S)__$/;

/**
 * Phân tích và kiểm tra tính hợp lệ của chuỗi marker
 */
function parseMarkerToken(token) {
  // Kiểm tra lỗi chữ O thay cho số 0
  if (token === '__GO__' || token === '__GO#__' || token === '__#GO__') {
    return { error: "Sử dụng marker sai '" + token + "' (chữ O). Hãy sửa thành '__G0__' (số không)." };
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
 * Phân tích toàn bộ đề thi dạng văn bản thuần
 * @param {string} rawText
 * @returns {Object} { valid, errors, groups, stats }
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

  // Quét từng dòng
  for (let i = 0; i < lines.length; i++) {
    const lineNumber = i + 1;
    const rawLine = lines[i];
    const trimmedLine = rawLine.trim();

    // Bỏ qua dòng trống
    if (!trimmedLine) {
      continue;
    }

    // Kiểm tra các chuỗi có dạng marker __...__ trên dòng
    const markerMatches = [...rawLine.matchAll(/__([A-Za-z0-9#_]+)__/g)];

    // 1. Kiểm tra nếu có marker nhưng không nằm ở đầu dòng
    if (markerMatches.length > 0) {
      const firstMatch = markerMatches[0];
      const matchIndex = firstMatch.index;
      const leadingPart = rawLine.substring(0, matchIndex);
      
      // Nếu có ký tự khác khoảng trắng đứng trước marker
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

      // Kiểm tra nếu có nhiều hơn 1 marker trên cùng một dòng
      if (markerMatches.length > 1) {
        errors.push({
          line: lineNumber,
          message: `Dòng ${lineNumber}: Phát hiện nhiều marker trên cùng một dòng ('${markerMatches.map(m => m[0]).join("', '")}'). Mỗi marker phải nằm trên một dòng riêng.`
        });
        continue;
      }
    }

    // 2. Nếu dòng bắt đầu bằng marker hợp lệ
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
          message: `Dòng ${lineNumber}: Marker '${token}' không hợp lệ. Các marker được hỗ trợ: __QUESTION__, __D__, __S__, __D#__, __S#__, __G__, __GQ__, __GA__, __G0__ và các biến thể cố định có '#'.`
        });
        continue;
      }

      // Xử lý marker NHÓM
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

      // Xử lý marker CÂU HỎI (__QUESTION__)
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

      // Xử lý marker ĐÁP ÁN (__D__, __S__, __D#__, __S#__)
      if (parsedToken.type === 'answer') {
        if (!currentQuestion) {
          errors.push({
            line: lineNumber,
            message: `Dòng ${lineNumber}: Đáp án '${token}' xuất hiện khi chưa có câu hỏi (__QUESTION__).`
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

    // 3. Dòng không có marker
    // Kiểm tra: "Có nội dung trước __QUESTION__ đầu tiên mà không phải marker nhóm"
    if (!hasMetFirstQuestion) {
      errors.push({
        line: lineNumber,
        message: `Dòng ${lineNumber}: Nội dung '${trimmedLine.substring(0, 40)}${trimmedLine.length > 40 ? '...' : ''}' xuất hiện trước câu hỏi (__QUESTION__) đầu tiên mà không phải marker nhóm.`
      });
      continue;
    }

    // Nối tiếp nội dung vào đáp án hoặc câu hỏi đang mở
    if (currentAnswer) {
      currentAnswer.text = currentAnswer.text ? `${currentAnswer.text}\n${trimmedLine}` : trimmedLine;
    } else if (currentQuestion) {
      currentQuestion.text = currentQuestion.text ? `${currentQuestion.text}\n${trimmedLine}` : trimmedLine;
    }
  }

  // 4. Kiểm tra logic từng câu hỏi sau khi phân tích
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
          message: `Câu ${q.id} (dòng ${q.startLine}) chưa có đáp án nào. Hãy thêm đáp án cho câu hỏi.`
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
        // Trả lời ngắn: chỉ có 1 __D__, không có __S__
        q.type = 'short';
        shortCount++;
        // Mục 10: câu chỉ có __D__ tự coi như __GQ__
        q.effectiveMode = 'GQ';
      } else if (correctAnswers.length > 1) {
        // Đúng / Sai nhiều ý
        q.type = 'tf';
        tfCount++;
      } else {
        // Trắc nghiệm một lựa chọn
        q.type = 'choice';
        choiceCount++;
      }

      // Kiểm tra nội dung rỗng
      if (!q.text.trim()) {
        errors.push({
          line: q.startLine,
          message: `Câu ${q.id} (dòng ${q.startLine}) chưa có nội dung câu hỏi.`
        });
      }
    }
  }

  if (allQuestions.length === 0 && errors.length === 0) {
    errors.push({
      line: 1,
      message: 'Chưa tìm thấy câu hỏi nào. Hãy thêm ít nhất một câu hỏi với marker __QUESTION__.'
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
