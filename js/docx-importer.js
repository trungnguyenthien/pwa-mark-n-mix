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
export async function extractTextFromDocx(file) {
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
export function normalizeMarkedText(rawText) {
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
export async function readExamDocx(file) {
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
