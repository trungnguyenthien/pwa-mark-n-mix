/**
 * Module Chuyển Đổi Đề Thi Truyền Thống Sang Chuẩn Marker (Marker Converter)
 * - Nhận diện định dạng câu hỏi phổ biến:
 *     Câu 1., Câu 1:, Câu 1), 1), 1., Bài 1., Bài 1:, Bài 1)
 *     (kể cả khi in đậm <b>, in nghiêng <i>, gạch chân <u>)
 * - Nhận diện định dạng đáp án:
 *     A., A), a., a), B., B), b., b), C., C), c., c), D., D), d., d)
 *     (hỗ trợ cả các phương án viết thẳng hàng 2 đến 4 cột hoặc mỗi phương án một dòng)
 * - Tự động phát hiện gạch chân (underline) ở nhãn hoặc nội dung đáp án:
 *     Có gạch chân <u> -> chuyển thành __D__ (Đáp án đúng)
 *     Không có gạch chân -> chuyển thành __S__ (Đáp án sai)
 * - Gỡ nhãn chữ cái thừa (A., A), a., a)...) và gỡ gạch chân đánh dấu để văn bản sạch đẹp
 * - Bảo toàn 100% hình ảnh inline [IMG:instanceId] và các thẻ định dạng <b>, <i>, <u>
 * - Dựng và xuất file Word (.docx) sạch, chuẩn Times New Roman 12pt sẵn sàng tải về máy
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
 * Phân tích và chuyển đổi chuỗi rawText trích xuất từ DOCX sang cú pháp Marker chuẩn
 * @param {string} rawText Chuỗi văn bản trích xuất từ extractTextFromDocx
 * @returns {{ convertedText: string, stats: Object }}
 */
export function convertTextToMarkers(rawText) {
  if (!rawText || !rawText.trim()) {
    return {
      convertedText: '',
      stats: {
        totalQuestions: 0,
        totalD: 0,
        totalS: 0,
        questionsWithoutKey: [],
        questionsWithMultipleKeys: [],
        warnings: ['Nội dung rỗng hoặc không có dữ liệu']
      }
    };
  }

  const lines = rawText.split(/\r?\n/);
  const outputLines = [];

  // Regex nhận diện tiêu đề câu hỏi: Câu 1., Câu 1:, Câu 1), 1), 1., Bài 1., Bài 1:, Bài 1)...
  // Hỗ trợ Câu hỏi 1, Câu số 1, Bài 1, Câu 1a, Câu 1.1, Câu I, Câu II...
  // Hỗ trợ đầy đủ mọi kiểu bọc thẻ in đậm <b>, in nghiêng <i>, gạch chân <u> xung quanh hoặc bên trong dấu chấm/hai chấm/điểm số
  const qNumPattern = "(?:\\d+(?:[a-zA-Z]|\\.\\d+)?|[IVXLCDM]+)";
  const qLabelPattern = "(?:câu\\s*(?:hỏi|số)?|bài)";
  const qRegex = new RegExp(
    "^\\s*(?:<[biu]>\\s*)*(?:" +
      "(?:" + qLabelPattern + ")\\s*(?:<[biu]>\\s*)*" + qNumPattern + "(?:\\s*<\\/[biu]>)*(?:\\s*[\\(\\[][^\\)\\]]+[\\)\\]])?(?:\\s*<[biu]>)*\\s*[:.)\\-–—/]?(?:\\s*<\\/[biu]>)*" +
      "|" +
      "(?:\\d+)(?:\\s*<\\/[biu]>)*\\s*[:.)\\-–—/](?:\\s*<\\/[biu]>)*" +
    ")(?:\\s*[\\(\\[][^\\)\\]]+[\\)\\]])?(?:\\s*<\\/[biu]>)*\\s*",
    "i"
  );

  // Regex nhận diện phân chia Phần / Chương / Mục: PHẦN I, PHẦN 1, CHƯƠNG I, I. PHẦN TRẮC NGHIỆM...
  const sectionRegex = /^\s*(?:<[biu]>\s*)*(?:(?:phần|chương|mục)\s+([IVXLCDM\d]+)|(?:[IVXLCDM]+|[A-D])\.\s*(?:phần|chương|mục|trắc nghiệm|tự luận))[\s.:]/i;

  // Regex nhận diện nhãn đáp án trên dòng: A., A), a., a), B., B), b., b)...
  // Có thể đứng ở đầu dòng hoặc ngăn cách bởi ít nhất 2 khoảng trắng hoặc ký tự tab \t
  const choiceTokenRegex = /(?:^|[\t\s]{2,}|\t)\s*((?:<[biu]>\s*)*[A-Da-d](?:\s*<\/[biu]>)*(?:\s*<[biu]>)*\s*[.:)\-–—/](?:\s*<\/[biu]>)*)\s*/g;

  let currentBlock = null; // 'HEADER', 'G', 'QUESTION', 'CHOICE'
  let hasGroup = false;
  let questionCount = 0;
  let dCount = 0;
  let sCount = 0;

  let currentQNumber = 0;
  let currentQChoices = [];
  const questionsInfo = [];

  function finishCurrentQuestion() {
    if (currentQNumber > 0) {
      const correctCount = currentQChoices.filter(c => c.isCorrect).length;
      questionsInfo.push({
        qNum: currentQNumber,
        choicesCount: currentQChoices.length,
        correctCount,
        hasKey: correctCount > 0,
        hasMultipleKeys: correctCount > 1
      });
      currentQChoices = [];
    }
  }

  for (let i = 0; i < lines.length; i++) {
    const rawLine = lines[i];
    const line = rawLine.trim();
    if (!line) continue;

    // 1. Nếu dòng đã chứa sẵn marker chuẩn (__G__, __QUESTION__, __D__, __S__...)
    if (/^__(?:G|QUESTION|D|S|D#|S#|#D|#S)__/.test(line)) {
      if (line.startsWith('__QUESTION__')) {
        finishCurrentQuestion();
        questionCount++;
        currentQNumber = questionCount;
        currentBlock = 'QUESTION';
      } else if (line.startsWith('__G')) {
        finishCurrentQuestion();
        hasGroup = true;
        currentBlock = 'G';
      } else if (line.startsWith('__D') || line.startsWith('__S')) {
        currentBlock = 'CHOICE';
        const isD = line.startsWith('__D');
        if (isD) dCount++; else sCount++;
        currentQChoices.push({ isCorrect: isD, text: line });
      }
      outputLines.push(line);
      continue;
    }

    // 2. Nhận diện tiêu đề nhóm / phần (ví dụ PHẦN I. CÂU TRẮC NGHIỆM...)
    if (sectionRegex.test(line)) {
      finishCurrentQuestion();
      outputLines.push('__G__');
      outputLines.push(line);
      hasGroup = true;
      currentBlock = 'G';
      continue;
    }

    // 3. Nhận diện bắt đầu câu hỏi (Câu 1., Câu 1:, 1)...)
    const qMatch = line.match(qRegex);
    if (qMatch) {
      finishCurrentQuestion();
      questionCount++;
      currentQNumber = questionCount;
      currentBlock = 'QUESTION';

      // Tách bỏ nhãn câu hỏi (ví dụ "Câu 1. ", "<b>Câu 1:</b> ", "1) ")
      const matchedStr = qMatch[0];
      const rest = line.substring(matchedStr.length);

      // Cân bằng các thẻ định dạng chưa đóng trong phần tiền tố bị cắt
      const opened = [];
      const tagRegex = /<\/?([biu])>/gi;
      let tm;
      while ((tm = tagRegex.exec(matchedStr)) !== null) {
        if (tm[0].startsWith('</')) {
          const idx = opened.lastIndexOf(tm[1].toLowerCase());
          if (idx !== -1) opened.splice(idx, 1);
        } else {
          opened.push(tm[1].toLowerCase());
        }
      }
      const prefixToRest = opened.map(t => '<' + t + '>').join('');
      let cleanRest = (prefixToRest + rest).trim();

      // Nếu dòng này chỉ chứa nhãn câu hỏi ("Câu 1.", "<b>Câu 1:</b>") mà nội dung nằm ở dòng tiếp theo:
      // Tự động gộp nội dung dòng kế tiếp vào cùng một dòng với __QUESTION__
      if (!cleanRest) {
        let nextIndex = i + 1;
        while (nextIndex < lines.length && !lines[nextIndex].trim()) {
          nextIndex++;
        }
        if (nextIndex < lines.length) {
          const candidate = lines[nextIndex].trim();
          choiceTokenRegex.lastIndex = 0;
          const isNextQ = qRegex.test(candidate);
          const isNextSection = sectionRegex.test(candidate);
          const isNextChoice = choiceTokenRegex.test(candidate);
          const isNextMarker = /^__(?:G|QUESTION|D|S)__/.test(candidate);

          if (!isNextQ && !isNextSection && !isNextChoice && !isNextMarker) {
            cleanRest = candidate;
            i = nextIndex; // Nhảy qua dòng tiếp theo đã được gộp
          }
        }
      }

      // Xuất marker __QUESTION__ trực tiếp thay thế vị trí của "Câu N"
      if (cleanRest) {
        outputLines.push(`__QUESTION__ ${cleanRest}`);
      } else {
        outputLines.push('__QUESTION__');
      }
      continue;
    }

    // 4. Nhận diện các phương án đáp án (A., B., C., D. hoặc a), b), c), d)...)
    choiceTokenRegex.lastIndex = 0;
    const choiceMatches = [];
    let cm;
    while ((cm = choiceTokenRegex.exec(line)) !== null) {
      choiceMatches.push({
        fullMatch: cm[0],
        labelWithTags: cm[1],
        index: cm.index
      });
    }

    // Kiểm tra tính hợp lệ của dòng đáp án:
    // Phải nằm trong phạm vi câu hỏi HOẶC nhãn đầu tiên là A/a
    const firstLabelClean = choiceMatches.length > 0
      ? choiceMatches[0].labelWithTags.replace(/<[^>]+>/g, '').trim()
      : '';
    const isFirstChoiceA = /^[Aa][.:)]/.test(firstLabelClean);

    if (choiceMatches.length > 0 && (currentBlock === 'QUESTION' || currentBlock === 'CHOICE' || isFirstChoiceA)) {
      currentBlock = 'CHOICE';

      for (let c = 0; c < choiceMatches.length; c++) {
        const cur = choiceMatches[c];
        const nxt = choiceMatches[c + 1];
        const segEnd = nxt ? nxt.index : line.length;
        const seg = line.substring(cur.index, segEnd).trim();

        // Kiểm tra xem phương án có được gạch chân (underline) hay không
        const isUnderlined = /<u[\s>]/i.test(seg);

        // Tách nhãn đáp án (A., A), a., a)...) khỏi nội dung
        const labelRegex = /^\s*(?:<([biu])>\s*)*([A-Da-d])(?:\s*<\/\1>)*(?:\s*<([biu])>\s*)*\s*[.:)\-–—/]\s*(?:<\/[biu]>\s*)*/i;
        let content = seg.replace(labelRegex, '').trim();

        // Gỡ bỏ thẻ <u> bao bọc bên ngoài nội dung phương án nếu có
        // (Do gạch chân chỉ dùng để đánh dấu đáp án đúng, không in ra gạch chân cho thí sinh)
        let openU = (content.match(/<u[\s>]/gi) || []).length;
        let closeU = (content.match(/<\/u>/gi) || []).length;
        if (closeU > openU) {
          content = content.replace(/<\/u>\s*$/i, '').trim();
        } else if (openU > closeU) {
          content = content + '</u>';
        }
        const outerUMatch = content.match(/^<u>([\s\S]*?)<\/u>$/i);
        if (outerUMatch) {
          content = outerUMatch[1].trim();
        }

        const marker = isUnderlined ? '__D__' : '__S__';
        if (isUnderlined) dCount++; else sCount++;

        currentQChoices.push({ isCorrect: isUnderlined, text: content });
        outputLines.push(`${marker} ${content}`);
      }
      continue;
    }

    // 5. Các dòng nội dung khác (đoạn nối câu hỏi, mô tả bảng, công thức, ảnh...)
    outputLines.push(line);
  }

  finishCurrentQuestion();

  // Đảm bảo luôn có __G__ trước câu hỏi đầu tiên nếu đề chưa chia nhóm/phần
  if (!hasGroup) {
    const firstQIdx = outputLines.findIndex(l => l.startsWith('__QUESTION__'));
    if (firstQIdx !== -1) {
      outputLines.splice(firstQIdx, 0, '__G__');
    } else {
      outputLines.unshift('__G__');
    }
  }

  // Thu thập các cảnh báo lỗi logic
  const questionsWithoutKey = questionsInfo.filter(q => !q.hasKey).map(q => q.qNum);
  const questionsWithMultipleKeys = questionsInfo.filter(q => q.hasMultipleKeys).map(q => q.qNum);
  const warnings = [];

  if (questionCount === 0) {
    warnings.push('Không phát hiện được câu hỏi nào theo chuẩn (Câu 1., Câu 1:, 1)...). Vui lòng kiểm tra lại file gốc.');
  }

  if (questionsWithoutKey.length > 0) {
    warnings.push(`Có ${questionsWithoutKey.length} câu hỏi chưa có đáp án gạch chân (__D__): Câu ${questionsWithoutKey.join(', ')}.`);
  }

  if (questionsWithMultipleKeys.length > 0) {
    warnings.push(`Có ${questionsWithMultipleKeys.length} câu hỏi có nhiều hơn 1 đáp án gạch chân: Câu ${questionsWithMultipleKeys.join(', ')}.`);
  }

  return {
    convertedText: outputLines.join('\n'),
    stats: {
      totalQuestions: questionCount,
      totalD: dCount,
      totalS: sCount,
      questionsWithoutKey,
      questionsWithMultipleKeys,
      warnings
    }
  };
}

/**
 * Tạo Blob file Word (.docx) chứa nội dung đã gắn marker chuẩn
 * - Font: Times New Roman
 * - Cỡ chữ: 12pt (size: 24)
 * - Căn lề chuẩn A4: Trái 2.5cm, Phải 2.0cm, Trên 2.0cm, Dưới 2.0cm
 * - Bảo toàn toàn bộ ảnh inline [IMG:instanceId]
 * @param {string} convertedText Chuỗi đề thi đã gắn marker
 * @param {Object} imageMap Bản đồ hình ảnh trích xuất từ DOCX gốc
 * @returns {Promise<Blob>}
 */
export async function generateConvertedDocxBlob(convertedText, imageMap = {}) {
  const docx = getDocx();
  if (!docx) {
    throw new Error('Thư viện docx chưa được nạp!');
  }

  const {
    Document,
    Paragraph,
    TextRun,
    ImageRun,
    UnderlineType,
    Packer
  } = docx;

  const FONT_FAMILY = 'Times New Roman';
  const COLOR_PRIMARY = '185ABD'; // Xanh Word cho __QUESTION__
  const COLOR_CORRECT = '107C41'; // Xanh lá cho __D__
  const COLOR_INCORRECT = '333333'; // Xám đen cho __S__
  const COLOR_GROUP = '7C3AED';   // Tím cho __G__
  const COLOR_BLACK = '000000';

  // Hàm phụ trợ tách đoạn văn bản kèm định dạng <b>, <i>, <u>, [IMG:...]
  function parseContentRuns(text) {
    const runs = [];
    if (!text) return runs;

    const tokenRegex = /(<br\s*\/?>|\n|\[IMG:[^\]]+\]|<\/?(?:b|i|u)>)/gi;
    const parts = text.split(tokenRegex);

    let isBold = false;
    let isItalic = false;
    let isUnderline = false;

    for (const part of parts) {
      if (!part) continue;
      const lower = part.toLowerCase();
      if (lower === '<b>') {
        isBold = true;
      } else if (lower === '</b>') {
        isBold = false;
      } else if (lower === '<i>') {
        isItalic = true;
      } else if (lower === '</i>') {
        isItalic = false;
      } else if (lower === '<u>') {
        isUnderline = true;
      } else if (lower === '</u>') {
        isUnderline = false;
      } else if (lower === '\n' || lower === '<br>' || lower === '<br/>') {
        runs.push(new TextRun({ break: 1, font: FONT_FAMILY, size: 24 }));
      } else if (/^\[IMG:([^\]]+)\]$/i.test(part)) {
        const imgMatch = part.match(/^\[IMG:([^\]]+)\]$/i);
        const imgId = imgMatch[1];
        const imgInfo = imageMap[imgId];
        if (imgInfo && imgInfo.buffer) {
          let w = imgInfo.width || 180;
          let h = imgInfo.height || 100;
          const MAX_WIDTH = 500;
          const MAX_HEIGHT = 450;
          if (w > MAX_WIDTH) {
            h = Math.round(h * (MAX_WIDTH / w));
            w = MAX_WIDTH;
          }
          if (h > MAX_HEIGHT) {
            w = Math.round(w * (MAX_HEIGHT / h));
            h = MAX_HEIGHT;
          }
          runs.push(new ImageRun({
            data: imgInfo.buffer,
            transformation: { width: w, height: h }
          }));
        }
      } else {
        runs.push(new TextRun({
          text: part,
          font: FONT_FAMILY,
          size: 24,
          bold: isBold,
          italics: isItalic,
          underline: isUnderline ? { type: UnderlineType.SINGLE } : undefined,
          color: COLOR_BLACK
        }));
      }
    }
    return runs;
  }

  const lines = convertedText.split(/\r?\n/);
  const docChildren = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) {
      // Dòng trống cách đoạn nhẹ nhàng
      docChildren.push(new Paragraph({
        spacing: { before: 60, after: 60 },
        children: []
      }));
      continue;
    }

    // 1. Dòng marker __G__ hoặc các biến thể nhóm
    if (/^__G\w*__/.test(line)) {
      docChildren.push(new Paragraph({
        spacing: { before: 180, after: 80 },
        children: [
          new TextRun({
            text: line,
            bold: true,
            font: FONT_FAMILY,
            size: 24,
            color: COLOR_GROUP
          })
        ]
      }));
      continue;
    }

    // 2. Dòng marker __QUESTION__ (đứng riêng hoặc cùng dòng với câu hỏi)
    if (line === '__QUESTION__' || /^__QUESTION__\s*/.test(line)) {
      const markerMatch = line.match(/^__QUESTION__\s*/)[0];
      const rest = line.substring(markerMatch.length).trim();
      const children = [
        new TextRun({
          text: '__QUESTION__ ',
          bold: true,
          font: FONT_FAMILY,
          size: 24,
          color: COLOR_PRIMARY
        })
      ];
      if (rest) {
        children.push(...parseContentRuns(rest));
      }
      docChildren.push(new Paragraph({
        spacing: { before: 140, after: 60 },
        children
      }));
      continue;
    }

    // 3. Dòng đáp án đúng (__D__, __D#__...)
    if (/^__(?:D|D#|#D)__\s*/.test(line)) {
      const markerMatch = line.match(/^__(?:D|D#|#D)__\s*/)[0];
      const rest = line.substring(markerMatch.length);
      const markerText = markerMatch.trim() + ' ';

      const children = [
        new TextRun({
          text: markerText,
          bold: true,
          font: FONT_FAMILY,
          size: 24,
          color: COLOR_CORRECT
        }),
        ...parseContentRuns(rest)
      ];

      docChildren.push(new Paragraph({
        spacing: { before: 40, after: 40 },
        children
      }));
      continue;
    }

    // 4. Dòng đáp án sai (__S__, __S#__...)
    if (/^__(?:S|S#|#S)__\s*/.test(line)) {
      const markerMatch = line.match(/^__(?:S|S#|#S)__\s*/)[0];
      const rest = line.substring(markerMatch.length);
      const markerText = markerMatch.trim() + ' ';

      const children = [
        new TextRun({
          text: markerText,
          bold: true,
          font: FONT_FAMILY,
          size: 24,
          color: COLOR_INCORRECT
        }),
        ...parseContentRuns(rest)
      ];

      docChildren.push(new Paragraph({
        spacing: { before: 40, after: 40 },
        children
      }));
      continue;
    }

    // 5. Các dòng văn bản nội dung thông thường
    const runs = parseContentRuns(line);
    docChildren.push(new Paragraph({
      spacing: { before: 60, after: 60 },
      children: runs
    }));
  }

  // Khởi tạo Document
  const doc = new Document({
    sections: [{
      properties: {
        page: {
          margin: {
            top: 1134,    // 2.0 cm
            bottom: 1134, // 2.0 cm
            left: 1417,   // 2.5 cm
            right: 1134   // 2.0 cm
          }
        }
      },
      children: docChildren
    }]
  });

  return await Packer.toBlob(doc);
}
