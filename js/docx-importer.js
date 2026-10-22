/**
 * Module Nhập File Word (.docx) & Tệp Văn Bản (.txt)
 * - Đọc trực tiếp tệp Word (.docx) trên trình duyệt bằng JSZip qua ArrayBuffer
 * - Bảo toàn 100% ngắt dòng (w:br, w:cr) và dấu tiếng Việt
 * - Tự động trích xuất thông tin đầu đề (Sở, Trường, Môn, Thời gian)
 * - Tự động nhận diện định dạng marker (__QUESTION__, __D__, __S__)
 * - Tự động chuyển đổi đề thi Word truyền thống (Câu 1: A. B. C. D. có gạch chân/đậm/đỏ/dấu * thành __D__)
 * - Lọc bỏ phần văn bản rác hoặc phần giải thích trước câu hỏi đầu tiên
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
 * Trích xuất toàn bộ văn bản và định dạng từ tệp Word (.docx)
 * @param {Blob|File|ArrayBuffer} file
 * @returns {Promise<string>}
 */
export async function extractTextFromDocx(file) {
  const JSZip = getJSZip();
  if (!JSZip) {
    throw new Error('Thư viện JSZip chưa được tải!');
  }

  // Chuyển sang ArrayBuffer để tương thích tối đa trên mọi trình duyệt (Safari, Chrome, Mobile)
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

  // Phân tích cú pháp XML của Word bảo toàn <w:br/>, <w:cr/>, <w:tab/>
  const pRegex = /<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g;
  const lines = [];
  let pMatch;

  while ((pMatch = pRegex.exec(xmlText)) !== null) {
    let pContent = pMatch[1];

    // 1. Chuyển đổi ngắt dòng nội bộ
    pContent = pContent.replace(/<w:br(?:\s[^>]*)?\/>/gi, '\n');
    pContent = pContent.replace(/<w:cr(?:\s[^>]*)?\/>/gi, '\n');
    pContent = pContent.replace(/<w:tab(?:\s[^>]*)?\/>/gi, '\t');

    // 2. Giải mã thực thể XML
    let decoded = pContent
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'");

    // 3. Quét định dạng đáp án đúng truyền thống (gạch chân hoặc chữ đỏ)
    decoded = decoded.replace(/<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/gi, (match, rInner) => {
      const isUnderline = /<w:u(?:\s[^>]*)?\/>/i.test(rInner) || /<w:u\s+[^>]*w:val="(single|words|thick|double)"/i.test(rInner);
      const isRed = /<w:color\s+[^>]*w:val="(FF0000|ED1C24|red)"/i.test(rInner);
      if (isUnderline || isRed) {
        // Gắn cờ dấu * trước chữ cái đáp án
        return rInner.replace(/(<w:t(?:\s[^>]*)?>\s*)([A-Da-d][\.)])/g, '$1*$2');
      }
      return rInner;
    });

    // 4. Trích xuất text từ các thẻ <w:t>, bảo toàn ký tự xuống dòng \n
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

    // Nếu không có thẻ w:t nào (đoạn trống hoặc chỉ có ký tự)
    if (!pText && /<w:t/i.test(decoded)) {
      pText = decoded.replace(/<[^>]+>/g, '');
    }

    if (pText.trim()) {
      lines.push(pText.trim());
    }
  }

  const rawExtracted = lines.join('\n');

  // Nếu file Word là tài liệu đặc tả có mục "Ví dụ đầu vào" (mục 3.4)
  const sampleMatch = rawExtracted.match(/(?:Ví dụ đầu vào|Ví dụ đề thi)[\s\S]*?(__G[\s\S]*?)(?:\n\d+\.|\nQuy tắc trộn|$)/i);
  if (sampleMatch) {
    return sampleMatch[1].trim();
  }

  return rawExtracted;
}

/**
 * Tự động chuẩn hóa văn bản đề thi, tách thông tin đầu đề và chuyển đổi sang marker
 * @param {string} rawText
 * @returns {Object} { text, isConverted, extractedHeader }
 */
export function normalizeExamText(rawText) {
  const extractedHeader = {};

  // Hàm quét thông tin đầu đề
  function checkHeaderLine(line) {
    if (/SỞ\s+GD/i.test(line)) {
      extractedHeader.soGD = line;
      return true;
    }
    if (/TRƯỜNG/i.test(line)) {
      extractedHeader.truong = line;
      return true;
    }
    if (/MÔN[:\s]/i.test(line)) {
      const m = line.match(/MÔN[:\s]+(.*)/i);
      if (m) extractedHeader.mon = m[1].trim();
      return true;
    }
    if (/Thời\s+gian/i.test(line)) {
      const m = line.match(/Thời\s+gian[^:]*:\s*(.*)/i);
      if (m) extractedHeader.thoiGian = m[1].trim();
      return true;
    }
    if (/NĂM\s+HỌC/i.test(line)) {
      const m = line.match(/NĂM\s+HỌC[:\s]+(.*)/i);
      if (m) extractedHeader.namHoc = m[1].trim();
      return true;
    }
    if (/GIỮA\s+HK|CUỐI\s+HK|HỌC\s+KÌ/i.test(line)) {
      extractedHeader.kiKiemTra = line;
      return true;
    }
    return false;
  }

  // TRƯỜNG HỢP 1: Văn bản đã có marker (__QUESTION__ hoặc __G__)
  if (rawText.includes('__QUESTION__') || rawText.includes('__D__')) {
    const lines = rawText.split('\n');
    const cleanedLines = [];
    let foundFirstMarker = false;

    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed) continue;

      // Tìm vị trí marker đầu tiên
      if (!foundFirstMarker) {
        if (trimmed.startsWith('__G') || trimmed.startsWith('__QUESTION__')) {
          foundFirstMarker = true;
          cleanedLines.push(trimmed);
        } else {
          // Trích xuất thông tin đầu đề nếu có trước marker đầu tiên
          checkHeaderLine(trimmed);
        }
      } else {
        cleanedLines.push(trimmed);
      }
    }

    return {
      text: cleanedLines.join('\n'),
      isConverted: false,
      extractedHeader
    };
  }

  // TRƯỜNG HỢP 2: Đề thi Word truyền thống (Câu 1: A. B. C. D.)
  const lines = rawText.split('\n');
  const resultLines = ['__G__'];
  let currentAnswers = [];
  let foundFirstQuestion = false;

  function flushAnswers() {
    if (currentAnswers.length > 0) {
      const hasCorrect = currentAnswers.some(a => a.correct);
      currentAnswers.forEach((ans, idx) => {
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

    // Trước câu hỏi đầu tiên
    if (!foundFirstQuestion) {
      const qMatch = line.match(/^(?:Câu|Bài|Question)\s*(\d+)[:.\-\/\s]\s*(.*)$/i);
      if (qMatch) {
        foundFirstQuestion = true;
        resultLines.push(`__QUESTION__ ${qMatch[2] || line}`);
        continue;
      }
      checkHeaderLine(line);
      continue;
    }

    // Từ câu hỏi đầu tiên trở đi
    const qMatch = line.match(/^(?:Câu|Bài|Question)\s*(\d+)[:.\-\/\s]\s*(.*)$/i);
    if (qMatch) {
      flushAnswers();
      resultLines.push(`__QUESTION__ ${qMatch[2] || line}`);
      continue;
    }

    // Nhận diện nếu các đáp án A. B. C. D. nằm trên cùng một dòng
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

    // Nhận diện từng đáp án trên dòng riêng (A. / B. / *C. / a) / b))
    const singleAnsMatch = line.match(/^(\*?)\s*([A-D]|[a-d])[\.)]\s*(.*)$/);
    if (singleAnsMatch) {
      currentAnswers.push({
        correct: singleAnsMatch[1] === '*',
        text: singleAnsMatch[3].trim()
      });
      continue;
    }

    // Dòng nối tiếp của câu hỏi hoặc đáp án
    if (currentAnswers.length > 0) {
      currentAnswers[currentAnswers.length - 1].text += ' ' + line;
    } else {
      resultLines.push(line);
    }
  }

  flushAnswers();

  if (!foundFirstQuestion) {
    // Không tìm thấy dạng câu hỏi, giữ nguyên nội dung
    return {
      text: rawText,
      isConverted: false,
      extractedHeader
    };
  }

  return {
    text: resultLines.join('\n'),
    isConverted: true,
    extractedHeader
  };
}

/**
 * Đọc file tải lên (Word .docx hoặc Text .txt)
 * @param {File|Blob} file
 * @returns {Promise<Object>} { text, isConverted, isDocx, extractedHeader, fileName }
 */
export async function readExamFile(file) {
  const fileName = (file && file.name) ? file.name.toLowerCase() : '';
  const fileType = (file && file.type) ? file.type.toLowerCase() : '';
  const isDocx = fileName.endsWith('.docx') || fileType.includes('wordprocessingml') || fileType === 'application/zip';

  if (isDocx) {
    const docxText = await extractTextFromDocx(file);
    const normalized = normalizeExamText(docxText);
    return {
      text: normalized.text,
      isConverted: normalized.isConverted,
      extractedHeader: normalized.extractedHeader || {},
      isDocx: true,
      fileName: file.name || 'document.docx'
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
        fileName: file.name || 'document.txt'
      });
    };
    reader.onerror = () => reject(new Error('Lỗi khi đọc file text.'));
    reader.readAsText(file, 'UTF-8');
  });
}
