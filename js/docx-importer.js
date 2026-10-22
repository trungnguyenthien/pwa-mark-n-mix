/**
 * Module Nhập File Word (.docx) & Tệp Văn Bản (.txt)
 * - Đọc trực tiếp tệp Word (.docx) trên trình duyệt bằng JSZip
 * - Phân tích cấu trúc đoạn (w:p), khối chữ (w:r), ngắt dòng (w:br)
 * - Tự động trích xuất thông tin đầu đề (Sở, Trường, Môn, Thời gian)
 * - Tự động nhận diện định dạng marker (__QUESTION__, __D__, __S__)
 * - Tự động chuyển đổi đề thi Word truyền thống (Câu 1: A. B. C. D. có gạch chân/đậm/đỏ/dấu * thành __D__)
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
 * Trích xuất toàn bộ văn bản từ tệp Word (.docx)
 * @param {Blob|File} file
 * @returns {Promise<string>}
 */
export async function extractTextFromDocx(file) {
  const JSZip = getJSZip();
  if (!JSZip) {
    throw new Error('Thư viện JSZip chưa được tải!');
  }

  const zip = await JSZip.loadAsync(file);
  const docXmlFile = zip.file('word/document.xml');
  if (!docXmlFile) {
    throw new Error('Tệp không đúng định dạng Word (.docx hợp lệ): thiếu word/document.xml');
  }

  const xmlText = await docXmlFile.async('text');

  // Phân tích cú pháp XML của Word
  const pRegex = /<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g;
  const rRegex = /<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/g;
  const tRegex = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
  const brRegex = /<w:br(?:\s[^>]*)?\/>/g;
  const tabRegex = /<w:tab(?:\s[^>]*)?\/>/g;

  const paragraphs = [];
  let pMatch;

  while ((pMatch = pRegex.exec(xmlText)) !== null) {
    const pXml = pMatch[1];
    
    // Thay thế ngắt dòng và tab
    const pNormalized = pXml.replace(brRegex, '\n').replace(tabRegex, '\t');
    
    let pText = '';
    let rMatch;

    while ((rMatch = rRegex.exec(pNormalized)) !== null) {
      const rXml = rMatch[1];

      // Kiểm tra xem run có được in đậm, gạch chân hoặc màu đỏ không (dấu hiệu đáp án đúng trong đề Word)
      const isUnderline = /<w:u(?:\s[^>]*)?\/>/.test(rXml) || /<w:u\s+[^>]*w:val="(single|words|thick|double)"/i.test(rXml);
      const isRed = /<w:color\s+[^>]*w:val="(FF0000|ED1C24|red)"/i.test(rXml);

      let rText = '';
      let tMatch;
      while ((tMatch = tRegex.exec(rXml)) !== null) {
        let textPiece = tMatch[1]
          .replace(/&amp;/g, '&')
          .replace(/&lt;/g, '<')
          .replace(/&gt;/g, '>')
          .replace(/&quot;/g, '"')
          .replace(/&apos;/g, "'");
        rText += textPiece;
      }

      // Nếu run có định dạng đáp án đúng truyền thống (gạch chân hoặc chữ đỏ)
      // gắn cờ dấu sao '*' để bộ nhận diện chuyển thành __D__
      if ((isUnderline || isRed) && /^\s*([A-D]|[a-d])[\.)]/.test(rText)) {
        rText = '*' + rText.trimStart();
      }

      pText += rText;
    }

    // Dự phòng nếu không trích được qua run
    if (!pText) {
      let tMatch;
      while ((tMatch = tRegex.exec(pNormalized)) !== null) {
        pText += tMatch[1];
      }
    }

    const trimmed = pText.trim();
    if (trimmed) {
      paragraphs.push(trimmed);
    }
  }

  return paragraphs.join('\n');
}

/**
 * Tự động trích xuất thông tin đầu đề và chuyển đổi đề thi truyền thống (Câu 1: A. B. C. D.) sang hệ marker
 * @param {string} rawText
 * @returns {Object} { text, isConverted, extractedHeader }
 */
export function normalizeExamText(rawText) {
  // Nếu đã chứa các marker chuẩn, giữ nguyên
  if (
    rawText.includes('__QUESTION__') ||
    rawText.includes('__D__') ||
    rawText.includes('__S__')
  ) {
    return { text: rawText, isConverted: false, extractedHeader: {} };
  }

  // Kiểm tra xem có cấu trúc "Câu 1", "Câu 2"... không
  if (!/Câu\s+\d+[:.]/i.test(rawText)) {
    return { text: rawText, isConverted: false, extractedHeader: {} };
  }

  const lines = rawText.split('\n');
  const resultLines = ['__G__'];
  let currentAnswers = [];
  let foundFirstQuestion = false;
  const extractedHeader = {};

  function flushAnswers() {
    if (currentAnswers.length > 0) {
      const hasCorrect = currentAnswers.some(a => a.correct);
      currentAnswers.forEach((ans, idx) => {
        // Nếu có dấu * thì là __D__, nếu không thì tạm gán đáp án đầu tiên là __D__
        const isD = hasCorrect ? ans.correct : (idx === 0);
        const marker = isD ? '__D__' : '__S__';
        resultLines.push(`${marker} ${ans.text}`);
      });
      currentAnswers = [];
    }
  }

  for (let i = 0; i < lines.length; i++) {
    let line = lines[i].trim();
    if (!line) continue;

    // Trước câu hỏi đầu tiên: trích xuất thông tin đầu đề và không đưa vào thân đề thi
    if (!foundFirstQuestion) {
      const qMatch = line.match(/^Câu\s+\d+[:.]\s*(.*)$/i);
      if (qMatch) {
        foundFirstQuestion = true;
        resultLines.push(`__QUESTION__ ${qMatch[1]}`);
        continue;
      }

      // Trích xuất thông tin header nếu có
      if (/SỞ\s+GD/i.test(line)) {
        extractedHeader.soGD = line;
      } else if (/TRƯỜNG/i.test(line)) {
        extractedHeader.truong = line;
      } else if (/MÔN[:\s]/i.test(line)) {
        const m = line.match(/MÔN[:\s]+(.*)/i);
        if (m) extractedHeader.mon = m[1].trim();
      } else if (/Thời\s+gian/i.test(line)) {
        const m = line.match(/Thời\s+gian[^:]*:\s*(.*)/i);
        if (m) extractedHeader.thoiGian = m[1].trim();
      } else if (/NĂM\s+HỌC/i.test(line)) {
        const m = line.match(/NĂM\s+HỌC[:\s]+(.*)/i);
        if (m) extractedHeader.namHoc = m[1].trim();
      } else if (/GIỮA\s+HK|CUỐI\s+HK|HỌC\s+KÌ/i.test(line)) {
        extractedHeader.kiKiemTra = line;
      }
      continue;
    }

    // Từ câu hỏi đầu tiên trở đi
    const qMatch = line.match(/^Câu\s+\d+[:.]\s*(.*)$/i);
    if (qMatch) {
      flushAnswers();
      resultLines.push(`__QUESTION__ ${qMatch[1]}`);
      continue;
    }

    // Kiểm tra nếu các đáp án A. B. C. D. nằm trên cùng một dòng
    const multiAnsMatch = [...line.matchAll(/(?:\b|(?<=\s))(\*?)([A-D]|[a-d])[\.)]\s+([^A-D\n]+?)(?=(?:\s+(\*?)[A-D][\.)]|$))/g)];
    if (multiAnsMatch && multiAnsMatch.length >= 2) {
      multiAnsMatch.forEach((m) => {
        currentAnswers.push({
          correct: m[1] === '*',
          text: m[3].trim()
        });
      });
      continue;
    }

    // Phát hiện từng đáp án trên dòng riêng
    const singleAnsMatch = line.match(/^(\*?)\s*([A-D]|[a-d])[\.)]\s*(.*)$/);
    if (singleAnsMatch) {
      currentAnswers.push({
        correct: singleAnsMatch[1] === '*',
        text: singleAnsMatch[3].trim()
      });
      continue;
    }

    // Dòng nối tiếp
    if (currentAnswers.length > 0) {
      currentAnswers[currentAnswers.length - 1].text += ' ' + line;
    } else {
      resultLines.push(line);
    }
  }

  flushAnswers();

  return {
    text: resultLines.join('\n'),
    isConverted: true,
    extractedHeader
  };
}

/**
 * Đọc file tải lên (Word .docx hoặc Text .txt)
 * @param {File} file
 * @returns {Promise<Object>} { text, isConverted, isDocx, extractedHeader, fileName }
 */
export async function readExamFile(file) {
  const isDocx = file.name.endsWith('.docx') || file.type.includes('wordprocessingml');

  if (isDocx) {
    const docxText = await extractTextFromDocx(file);
    const normalized = normalizeExamText(docxText);
    return {
      text: normalized.text,
      isConverted: normalized.isConverted,
      extractedHeader: normalized.extractedHeader || {},
      isDocx: true,
      fileName: file.name
    };
  }

  // Tệp văn bản thuần .txt
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const normalized = normalizeExamText(e.target.result);
      resolve({
        text: normalized.text,
        isConverted: normalized.isConverted,
        extractedHeader: normalized.extractedHeader || {},
        isDocx: false,
        fileName: file.name
      });
    };
    reader.onerror = () => reject(new Error('Lỗi khi đọc file text.'));
    reader.readAsText(file, 'UTF-8');
  });
}
