/**
 * Module Nhập & Trích Xuất File Word (.docx) Đã Đánh Marker
 * - Sử dụng JSZip để đọc và phân tích cấu trúc word/document.xml trực tiếp trên trình duyệt
 * - Trích xuất và bảo toàn toàn bộ ảnh inline (<w:drawing>, <w:pict>) trong câu hỏi và đáp án
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
 * Trích xuất toàn bộ chuỗi văn bản, thẻ ảnh và ngắt dòng từ tệp Word (.docx)
 * @param {Blob|File|ArrayBuffer} file
 * @returns {Promise<{ rawText: string, imageMap: Object }>}
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

  // 1. Phân tích các mối liên kết quan hệ trong word/_rels/document.xml.rels để tìm tệp media (ảnh)
  const imageMap = {};
  const relsFile = zip.file('word/_rels/document.xml.rels');
  if (relsFile) {
    const relsXml = await relsFile.async('text');
    const relMatches = [...relsXml.matchAll(/<Relationship\s+([^>]*)\/>/gi)];
    for (const m of relMatches) {
      const attrs = m[1];
      const idMatch = attrs.match(/Id="([^"]+)"/i);
      const targetMatch = attrs.match(/Target="([^"]+)"/i);
      const typeMatch = attrs.match(/Type="([^"]+)"/i);

      if (idMatch && targetMatch) {
        const rId = idMatch[1];
        let target = targetMatch[1];
        const isImage = (typeMatch && typeMatch[1].endsWith('/image')) || /\.(png|jpe?g|gif|bmp|webp|svg)$/i.test(target);

        if (isImage) {
          if (!target.startsWith('word/')) {
            target = 'word/' + target.replace(/^(\.\.\/)+/, '');
          }
          const mediaFile = zip.file(target);
          if (mediaFile) {
            const ext = target.split('.').pop().toLowerCase();
            const base64 = await mediaFile.async('base64');
            const uint8 = await mediaFile.async('uint8array');
            const mime = ext === 'jpg' ? 'image/jpeg' : `image/${ext}`;
            const dataUrl = `data:${mime};base64,${base64}`;

            let naturalWidth = 0;
            let naturalHeight = 0;
            if (ext === 'png' && uint8.length >= 24) {
              naturalWidth = (uint8[16] << 24) | (uint8[17] << 16) | (uint8[18] << 8) | uint8[19];
              naturalHeight = (uint8[20] << 24) | (uint8[21] << 16) | (uint8[22] << 8) | uint8[23];
            } else if ((ext === 'jpg' || ext === 'jpeg') && uint8.length >= 4) {
              let offset = 2;
              while (offset < uint8.length - 8) {
                if (uint8[offset] === 0xFF && (uint8[offset + 1] >= 0xC0 && uint8[offset + 1] <= 0xC3)) {
                  naturalHeight = (uint8[offset + 5] << 8) | uint8[offset + 6];
                  naturalWidth = (uint8[offset + 7] << 8) | uint8[offset + 8];
                  break;
                }
                const segLen = (uint8[offset + 2] << 8) | uint8[offset + 3];
                if (segLen <= 0) break;
                offset += 2 + segLen;
              }
            } else if (ext === 'gif' && uint8.length >= 10) {
              naturalWidth = uint8[6] | (uint8[7] << 8);
              naturalHeight = uint8[8] | (uint8[9] << 8);
            }

            imageMap[rId] = {
              id: rId,
              path: target,
              mime,
              dataUrl,
              buffer: uint8,
              naturalWidth: naturalWidth || 180,
              naturalHeight: naturalHeight || 100,
              width: naturalWidth || 180,
              height: naturalHeight || 100
            };
          }
        }
      }
    }
  }

  const xmlText = await docXmlFile.async('text');

  // 2. Phân tích cú pháp XML của Word, xử lý từng đoạn văn <w:p>
  const pRegex = /<w:p(?:\s[^>]*)?>([\s\S]*?)<\/w:p>/g;
  const lines = [];
  let pMatch;
  let imgCounter = 0;

  while ((pMatch = pRegex.exec(xmlText)) !== null) {
    let pContent = pMatch[1];

    // Trích xuất hình ảnh DrawingML (<w:drawing>) và thay thế bằng token [IMG:instanceId]
    pContent = pContent.replace(/<w:drawing[\s\S]*?<\/w:drawing>/gi, (drawingXml) => {
      const blipMatch = drawingXml.match(/r:(?:embed|link)="([^"]+)"/i);
      if (blipMatch) {
        const rId = blipMatch[1];
        const extMatch = drawingXml.match(/<wp:extent\s+cx="(\d+)"\s+cy="(\d+)"/i) || drawingXml.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"/i);
        let w = 0, h = 0;
        if (extMatch) {
          w = Math.round(parseInt(extMatch[1], 10) / 9525);
          h = Math.round(parseInt(extMatch[2], 10) / 9525);
        }

        imgCounter++;
        const instanceId = `img_${imgCounter}`;
        if (imageMap[rId]) {
          const baseImg = imageMap[rId];
          const finalW = w || baseImg.width || 180;
          const finalH = h || baseImg.height || 100;
          imageMap[instanceId] = {
            ...baseImg,
            id: instanceId,
            width: finalW,
            height: finalH,
            originalRId: rId
          };
          if (!imageMap[rId].width || w) {
            imageMap[rId].width = finalW;
            imageMap[rId].height = finalH;
          }
        }
        return ` [IMG:${instanceId}] `;
      }
      return '';
    });

    // Trích xuất hình ảnh VML (<w:pict>) và thay thế bằng token [IMG:instanceId]
    pContent = pContent.replace(/<w:pict[\s\S]*?<\/w:pict>/gi, (pictXml) => {
      const idMatch = pictXml.match(/r:id="([^"]+)"/i);
      if (idMatch) {
        const rId = idMatch[1];
        let w = 0, h = 0;
        const styleMatch = pictXml.match(/style="([^"]+)"/i);
        if (styleMatch) {
          const wMatch = styleMatch[1].match(/width:\s*([\d.]+)(pt|px|in|cm)?/i);
          const hMatch = styleMatch[1].match(/height:\s*([\d.]+)(pt|px|in|cm)?/i);
          if (wMatch) {
            const val = parseFloat(wMatch[1]);
            const unit = (wMatch[2] || 'pt').toLowerCase();
            if (unit === 'pt') w = Math.round(val * 96 / 72);
            else if (unit === 'in') w = Math.round(val * 96);
            else if (unit === 'cm') w = Math.round(val / 2.54 * 96);
            else w = Math.round(val);
          }
          if (hMatch) {
            const val = parseFloat(hMatch[1]);
            const unit = (hMatch[2] || 'pt').toLowerCase();
            if (unit === 'pt') h = Math.round(val * 96 / 72);
            else if (unit === 'in') h = Math.round(val * 96);
            else if (unit === 'cm') h = Math.round(val / 2.54 * 96);
            else h = Math.round(val);
          }
        }

        imgCounter++;
        const instanceId = `img_${imgCounter}`;
        if (imageMap[rId]) {
          const baseImg = imageMap[rId];
          const finalW = w || baseImg.width || 180;
          const finalH = h || baseImg.height || 100;
          imageMap[instanceId] = {
            ...baseImg,
            id: instanceId,
            width: finalW,
            height: finalH,
            originalRId: rId
          };
          if (!imageMap[rId].width || w) {
            imageMap[rId].width = finalW;
            imageMap[rId].height = finalH;
          }
        }
        return ` [IMG:${instanceId}] `;
      }
      return '';
    });

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

    // Trích xuất text từ các thẻ <w:t> và giữ nguyên token ảnh [IMG:...]
    let pText = '';
    const tokenRegex = /(<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|\n|\[IMG:[^\]]+\])/g;
    let tokMatch;
    while ((tokMatch = tokenRegex.exec(decoded)) !== null) {
      if (tokMatch[0] === '\n') {
        pText += '\n';
      } else if (tokMatch[0].startsWith('[IMG:')) {
        pText += ' ' + tokMatch[0] + ' ';
      } else {
        pText += tokMatch[2];
      }
    }

    // Nếu đoạn không có thẻ w:t nhưng có chứa ảnh
    if (!pText && decoded.includes('[IMG:')) {
      const m = decoded.match(/\[IMG:[^\]]+\]/g);
      if (m) pText = m.join(' ');
    }

    if (pText.trim()) {
      lines.push(pText.trim());
    }
  }

  const rawExtracted = lines.join('\n');
  return {
    rawText: rawExtracted,
    imageMap
  };
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

  // Nếu không tìm thấy marker ở đầu (ví dụ các marker viết dính liền trên 1 dòng)
  if (cleanedLines.length === 0 && rawText.includes('__QUESTION__')) {
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
 * @returns {Promise<Object>} { text, extractedHeader, fileName, fileSize, imageMap }
 */
export async function readExamDocx(file) {
  const fileName = (file && file.name) ? file.name : 'De_Thi.docx';
  const fileSize = file.size || 0;

  const { rawText, imageMap } = await extractTextFromDocx(file);
  const normalized = normalizeMarkedText(rawText);

  return {
    text: normalized.text,
    extractedHeader: normalized.extractedHeader || {},
    fileName,
    fileSize,
    imageMap: imageMap || {}
  };
}
