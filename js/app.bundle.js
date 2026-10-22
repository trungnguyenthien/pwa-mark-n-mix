(function() {
"use strict";

// --- FILE: js/docx-importer.js ---
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

function decodeXml(str) {
  if (!str) return '';
  return str
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function wrapRunFormatting(text, b, i, u) {
  if (!text || (!b && !i && !u)) return text;
  const startSpaceMatch = text.match(/^(\s*)/);
  const endSpaceMatch = text.match(/(\s*)$/);
  const lead = startSpaceMatch ? startSpaceMatch[1] : '';
  const trail = endSpaceMatch ? endSpaceMatch[1] : '';
  const core = text.substring(lead.length, text.length - trail.length);
  if (!core) return text;
  let res = core;
  if (u) res = `<u>${res}</u>`;
  if (i) res = `<i>${res}</i>`;
  if (b) res = `<b>${res}</b>`;
  return lead + res + trail;
}

/**
 * Trích xuất toàn bộ chuỗi văn bản, thẻ ảnh và ngắt dòng từ tệp Word (.docx)
 * @param {Blob|File|ArrayBuffer} file
 * @returns {Promise<{ rawText: string, imageMap: Object }>}
 */
async function extractTextFromDocx(file) {
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
  const nonInlineImages = [];
  let pMatch;
  let imgCounter = 0;

  while ((pMatch = pRegex.exec(xmlText)) !== null) {
    let pContent = pMatch[1];

    // Trích xuất hình ảnh DrawingML (<w:drawing>) và thay thế bằng token [IMG:instanceId]
    pContent = pContent.replace(/<w:drawing[\s\S]*?<\/w:drawing>/gi, (drawingXml) => {
      const isInline = /<wp:inline[\s>]/i.test(drawingXml);
      const isAnchor = /<wp:anchor[\s>]/i.test(drawingXml);
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

        // Kiểm tra xem ảnh có phải là In Line with Text hay không
        if (isAnchor || !isInline) {
          let wrapType = 'Nổi tự do (Floating)';
          if (/<wp:wrapSquare/i.test(drawingXml)) {
            wrapType = 'Square (Vuông bao quanh)';
          } else if (/<wp:wrapTight/i.test(drawingXml)) {
            wrapType = 'Tight (Bao sát văn bản)';
          } else if (/<wp:wrapThrough/i.test(drawingXml)) {
            wrapType = 'Through (Xuyên qua)';
          } else if (/<wp:wrapTopAndBottom/i.test(drawingXml)) {
            wrapType = 'Top and Bottom (Trên và dưới)';
          } else if (/<wp:wrapNone/i.test(drawingXml) || /behindDoc="1"/i.test(drawingXml)) {
            const behind = /behindDoc="1"/i.test(drawingXml);
            wrapType = behind ? 'Behind Text (Nằm dưới chữ)' : 'In Front of Text (Nằm trên chữ)';
          }

          const docPrName = drawingXml.match(/<wp:docPr\s+[^>]*name="([^"]+)"/i);
          const docPrTitle = drawingXml.match(/<wp:docPr\s+[^>]*title="([^"]+)"/i);
          const docPrDescr = drawingXml.match(/<wp:docPr\s+[^>]*descr="([^"]+)"/i);
          const imgName = (docPrTitle && docPrTitle[1]) || (docPrName && docPrName[1]) || (docPrDescr && docPrDescr[1]) || `Hình ảnh ${imgCounter}`;

          const rawTextInParagraph = pContent.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
          const snippet = rawTextInParagraph.length > 70
            ? rawTextInParagraph.substring(0, 70) + '...'
            : (rawTextInParagraph || 'Ảnh nằm trên đoạn văn bản riêng');

          nonInlineImages.push({
            id: instanceId,
            name: imgName,
            rId,
            wrapType,
            snippet,
            dataUrl: (imageMap[rId] && imageMap[rId].dataUrl) ? imageMap[rId].dataUrl : ''
          });
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
        // Kiểm tra nếu pict là dạng nổi (VML floating)
        const isFloatingVml = /position\s*:\s*absolute/i.test(pictXml) || /mso-wrap-style/i.test(pictXml) || /mso-position/i.test(pictXml);
        if (isFloatingVml) {
          const rawTextInParagraph = pContent.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
          const snippet = rawTextInParagraph.length > 70
            ? rawTextInParagraph.substring(0, 70) + '...'
            : (rawTextInParagraph || 'Ảnh nằm trên đoạn văn bản riêng');
          nonInlineImages.push({
            id: instanceId,
            name: `Ảnh vẽ VML ${imgCounter}`,
            rId,
            wrapType: 'VML Floating (Ảnh vẽ nổi tự do)',
            snippet,
            dataUrl: (imageMap[rId] && imageMap[rId].dataUrl) ? imageMap[rId].dataUrl : ''
          });
        }

        return ` [IMG:${instanceId}] `;
      }
      return '';
    });

    // Trích xuất tuần tự các run <w:r> kèm định dạng bold, italic, underline và ngắt dòng
    const tokenRegex = /(<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>|<w:br(?:\s[^>]*)?\/>|<w:cr(?:\s[^>]*)?\/>|<w:tab(?:\s[^>]*)?\/>|\[IMG:[^\]]+\])/gi;
    let tokMatch;
    let pText = '';

    while ((tokMatch = tokenRegex.exec(pContent)) !== null) {
      const tok = tokMatch[0];
      if (tok.startsWith('[IMG:')) {
        pText += ' ' + tok + ' ';
      } else if (/<w:(?:br|cr)/i.test(tok)) {
        pText += '\n';
      } else if (/<w:tab/i.test(tok)) {
        pText += '\t';
      } else if (tok.startsWith('<w:r')) {
        const rPrMatch = tok.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/i);
        const rPr = rPrMatch ? rPrMatch[1] : '';

        const isBold = /<w:b(?:\s+(?!w:val="(?:0|false)")[^>]*)?\/>|<w:b\s+w:val="(?:1|true)"/i.test(rPr);
        const isItalic = /<w:i(?:\s+(?!w:val="(?:0|false)")[^>]*)?\/>|<w:i\s+w:val="(?:1|true)"/i.test(rPr);
        const isUnderline = /<w:u(?:\s+(?!w:val="(?:none|false|0)")[^>]*)?\/>|<w:u\s+w:val="(?!none|false|0)[^"]+"/i.test(rPr);

        const innerRegex = /(<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:br(?:\s[^>]*)?\/>|<w:cr(?:\s[^>]*)?\/>|<w:tab(?:\s[^>]*)?\/>|\[IMG:[^\]]+\])/gi;
        let inMatch;
        while ((inMatch = innerRegex.exec(tok)) !== null) {
          const inTok = inMatch[0];
          if (inTok.startsWith('[IMG:')) {
            pText += ' ' + inTok + ' ';
          } else if (/<w:(?:br|cr)/i.test(inTok)) {
            pText += '\n';
          } else if (/<w:tab/i.test(inTok)) {
            pText += '\t';
          } else if (inMatch[2] !== undefined) {
            const decoded = decodeXml(inMatch[2]);
            pText += wrapRunFormatting(decoded, isBold, isItalic, isUnderline);
          }
        }
      }
    }

    // Nếu không khớp run nào nhưng có text trực tiếp
    if (!pText) {
      const rawTextMatches = [...pContent.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/gi)];
      if (rawTextMatches.length > 0) {
        pText = rawTextMatches.map(m => decodeXml(m[1])).join('');
      } else if (pContent.includes('[IMG:')) {
        const m = pContent.match(/\[IMG:[^\]]+\]/g);
        if (m) pText = m.join(' ');
      }
    }

    // 1. Gộp các cặp thẻ định dạng liền kề cùng loại do Word tách run: </b><b>, </i><i>, </u><u>
    let prevMerge;
    do {
      prevMerge = pText;
      pText = pText.replace(/<\/([biu])>\s*<\1>/gi, '');
    } while (pText !== prevMerge);

    // 2. Gỡ bỏ tag format nếu marker bị bọc trong bold/italic
    pText = pText.replace(/<([biu])>\s*(__[A-Za-z0-9#]+__)\s*<\/\1>/gi, '$2');
    pText = pText.replace(/<([biu])>\s*(<([biu])>\s*(__[A-Za-z0-9#]+__)\s*<\/\3>)\s*<\/\1>/gi, '$4');

    // 3. Nếu marker nằm dính ở đầu hoặc cuối thẻ định dạng, tách marker ra ngoài để đảm bảo tính độc lập
    pText = pText.replace(/<([biu])>(\s*)(__[A-Za-z0-9#]+__)/gi, '$3<$1>$2');
    pText = pText.replace(/(__[A-Za-z0-9#]+__)(\s*)<\/([biu])>/gi, '</$3>$2$1');
    pText = pText.replace(/<([biu])>\s*<\/\1>/gi, '');

    if (pText.trim()) {
      lines.push(pText.trim());
    }
  }

  const rawExtracted = lines.join('\n');
  return {
    rawText: rawExtracted,
    imageMap,
    nonInlineImages
  };
}

/**
 * Kiểm tra xem một dòng có phải là dòng phân cách kết thúc đề thi / lời dặn cuối đề hay không
 * Ví dụ: "----------------------- HẾT ĐỀ KIỂM TRA -----------------------", "--- HẾT ---", "HẾT."
 */
function isEndOfExamLine(line) {
  if (!line) return false;
  const plain = line.replace(/<[^>]+>/g, '').trim();
  if (!plain) return false;
  // Dòng kẻ gạch ngang: ------------------, ==================
  if (/^[-–—=_*~]{3,}$/.test(plain)) return true;
  // Dòng chứa chữ HẾT: "--- HẾT ---", "HẾT ĐỀ KIỂM TRA", "- HẾT -", "HẾT", "HẾT ĐỀ THI", "HẾT."
  if (/^[-–—=_*~.\s]*(?:HẾT|THE\s+END)(?:\s+(?:ĐỀ|BÀI)?\s*(?:THI|KIỂM\s+TRA)?)?[-–—=_*~.\s]*$/i.test(plain)) return true;
  // Lời dặn giám thị / thí sinh ở cuối đề: "Cán bộ coi thi không giải thích gì thêm", "Giám thị không giải thích gì thêm", "Thí sinh không được sử dụng tài liệu"
  if (/^(?:Ghi\s+chú\s*:\s*)?(?:Thí\s+sinh\s+không\s+được|Cán\s+bộ\s+coi\s+thi|Giám\s+thị\s+không)/i.test(plain)) return true;
  return false;
}

/**
 * Tự động chuẩn hóa văn bản đã đánh marker từ file Word và trích xuất thông tin đầu đề nếu có
 * @param {string} rawText
 * @returns {Object} { text, extractedHeader }
 */
function normalizeMarkedText(rawText) {
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
      const markersInLine = [...trimmed.matchAll(/__([A-Za-z0-9#]+)__/g)];
      if (markersInLine.length > 0) {
        foundFirstMarker = true;
        cleanedLines.push(trimmed);
      } else {
        checkHeaderLine(trimmed);
      }
    } else {
      cleanedLines.push(trimmed);
    }
  }

  // Loại bỏ các dòng kết thúc đề thi ở cuối file (ví dụ "--- HẾT ---", "HẾT ĐỀ KIỂM TRA")
  while (cleanedLines.length > 0 && isEndOfExamLine(cleanedLines[cleanedLines.length - 1])) {
    cleanedLines.pop();
  }

  // Nếu không tìm thấy marker ở đầu (ví dụ các marker viết dính liền trên 1 dòng)
  if (cleanedLines.length === 0 && rawText.includes('__QUESTION__')) {
    const separated = rawText
      .replace(/(__[A-Za-z0-9#]+__)/g, '\n$1 ')
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
async function readExamDocx(file) {
  const fileName = (file && file.name) ? file.name : 'De_Thi.docx';
  const fileSize = file.size || 0;

  const { rawText, imageMap, nonInlineImages } = await extractTextFromDocx(file);
  const normalized = normalizeMarkedText(rawText);

  return {
    text: normalized.text,
    extractedHeader: normalized.extractedHeader || {},
    fileName,
    fileSize,
    imageMap: imageMap || {},
    nonInlineImages: nonInlineImages || []
  };
}

// --- FILE: js/parser.js ---
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
 * Kiểm tra xem một dòng có phải là dòng phân cách kết thúc đề thi / lời dặn cuối đề hay không
 * Ví dụ: "----------------------- HẾT ĐỀ KIỂM TRA -----------------------", "--- HẾT ---", "HẾT."
 */
function isEndOfExamLine(line) {
  if (!line) return false;
  const plain = line.replace(/<[^>]+>/g, '').trim();
  if (!plain) return false;
  // Dòng kẻ gạch ngang: ------------------, ==================
  if (/^[-–—=_*~]{3,}$/.test(plain)) return true;
  // Dòng chứa chữ HẾT: "--- HẾT ---", "HẾT ĐỀ KIỂM TRA", "- HẾT -", "HẾT", "HẾT ĐỀ THI", "HẾT."
  if (/^[-–—=_*~.\s]*(?:HẾT|THE\s+END)(?:\s+(?:ĐỀ|BÀI)?\s*(?:THI|KIỂM\s+TRA)?)?[-–—=_*~.\s]*$/i.test(plain)) return true;
  // Lời dặn giám thị / thí sinh ở cuối đề: "Cán bộ coi thi không giải thích gì thêm", "Giám thị không giải thích gì thêm", "Thí sinh không được sử dụng tài liệu"
  if (/^(?:Ghi\s+chú\s*:\s*)?(?:Thí\s+sinh\s+không\s+được|Cán\s+bộ\s+coi\s+thi|Giám\s+thị\s+không)/i.test(plain)) return true;
  return false;
}

/**
 * Phân tích văn bản chứa marker thành cấu trúc đề thi
 * @param {string} rawText
 * @returns {Object} { valid: boolean, errors: Array, groups: Array, stats: Object }
 */
function parseExam(rawText) {
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

    // Bỏ qua các dòng phân cách kết thúc đề thi (ví dụ "--- HẾT ---", "HẾT ĐỀ KIỂM TRA")
    if (hasMetFirstQuestion && isEndOfExamLine(trimmedLine)) {
      continue;
    }

    // Tìm tất cả marker có dạng __...__ trên dòng (marker chỉ gồm chữ, số và ký tự #)
    const markerMatches = [...rawLine.matchAll(/__([A-Za-z0-9#]+)__/g)];

    if (markerMatches.length === 0) {
      // Dòng không có marker
      // Nếu đang trong một nhóm nhưng chưa bắt đầu câu hỏi nào -> nội dung thuộc tiêu đề của nhóm đó
      if (currentGroup && !currentQuestion) {
        currentGroup.title = currentGroup.title ? `${currentGroup.title}\n${trimmedLine}` : trimmedLine;
        continue;
      }

      if (!hasMetFirstQuestion) {
        errors.push({
          line: lineNumber,
          message: `Dòng ${lineNumber}: Nội dung '${trimmedLine.substring(0, 50)}${trimmedLine.length > 50 ? '...' : ''}' xuất hiện trước câu hỏi (__QUESTION__) đầu tiên mà không thuộc nhóm nào.`
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
      } else if (currentGroup && !currentQuestion) {
        currentGroup.title = currentGroup.title ? `${currentGroup.title}\n${leadingText}` : leadingText;
      } else if (!hasMetFirstQuestion) {
        errors.push({
          line: lineNumber,
          message: `Dòng ${lineNumber}: Nội dung '${leadingText.substring(0, 50)}${leadingText.length > 50 ? '...' : ''}' xuất hiện trước câu hỏi (__QUESTION__) đầu tiên mà không thuộc nhóm nào.`
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
          title: content || '',
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

  // Hậu kiểm làm sạch các câu hỏi và đáp án nếu có dòng kết thúc đề bám vào
  for (const g of groups) {
    for (const q of g.questions) {
      if (q.text) {
        const qLines = q.text.split(/\r?\n/).filter(l => !isEndOfExamLine(l));
        q.text = qLines.join('\n').trim();
      }
      for (const a of q.answers) {
        if (a.text) {
          const aLines = a.text.split(/\r?\n/).filter(l => !isEndOfExamLine(l));
          a.text = aLines.join('\n').trim();
        }
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

// --- FILE: js/shuffler.js ---
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
function createRNG(seed) {
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
function hashStringToSeed(str) {
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
function shufflePreservingFixed(array, rng) {
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
function getAnswerLabel(index, style = 'ABCD') {
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
const ROMAN_NUMERALS = [
  'I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X',
  'XI', 'XII', 'XIII', 'XIV', 'XV', 'XVI', 'XVII', 'XVIII', 'XIX', 'XX'
];

/**
 * Chuyển đổi số thứ tự (1-based) sang số La Mã (I, II, ..., XX)
 * @param {number} n
 * @returns {string}
 */
function toRoman(n) {
  if (n >= 1 && n <= 20) return ROMAN_NUMERALS[n - 1];
  return String(n);
}

/**
 * Nhận diện dòng tiêu đề là IN HOA toàn bộ hay viết bình thường
 * @param {string} text
 * @returns {boolean} true nếu là IN HOA toàn bộ, false nếu viết bình thường
 */
function isTitleUpperCase(text) {
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
function formatGroupTitleFirstLine(line, roman) {
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
function formatGroupTitle(title, groupIndex) {
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
function shuffleExamForCode(parsedExam, examCode, options = {}) {
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
function shuffleAllExams(parsedExam, examCodes, options = {}) {
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

// --- FILE: js/docx-builder.js ---
/**
 * Module Dựng File Word Đề Thi (Docx Builder)
 * Tuân thủ quy định tại mục 5.1 và 5.2 của tài liệu:
 * "Công cụ trộn đề trắc nghiệm Yêu cầu và giải pháp.docx"
 *
 * - Font: Times New Roman toàn bộ văn bản.
 * - Đầu đề: Bảng 2 cột không viền căn giữa, phối màu đỏ/đen đậm chuẩn mực theo mẫu duyệt.
 * - Dòng "Họ và tên...", "Số báo danh...", "Mã đề [101]" kèm vạch kẻ ngang.
 * - Đánh số câu liên tục theo lựa chọn người dùng ("Câu 1." hoặc "Câu 1:").
 * - Nhãn đáp án theo định dạng đã chọn (A. B. C. D. / a) b) c) d) / 1. 2. 3. 4.).
 * - Câu trả lời ngắn (chỉ có __D__) không hiển thị đáp án trong đề thi.
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
 * Tạo Blob file Word (.docx) cho một mã đề thi cụ thể (hoặc file đáp án chi tiết tương ứng)
 * @param {Object} shuffledExam Đề thi đã hoán vị (từ shuffler)
 * @param {Object} headerInfo Thông tin đầu đề (7 ô nhập)
 * @param {Object} options Tùy chọn xuất đề: { isAnswerKey: boolean }
 * @returns {Promise<Blob>}
 */
async function createExamDocxBlob(shuffledExam, headerInfo = {}, options = {}) {
  const docx = getDocx();
  if (!docx) {
    throw new Error('Thư viện docx chưa được nạp!');
  }

  const {
    Document,
    Paragraph,
    TextRun,
    ImageRun,
    Table,
    TableRow,
    TableCell,
    WidthType,
    BorderStyle,
    AlignmentType,
    PageNumber,
    UnderlineType,
    Packer
  } = docx;

  const {
    soGD = 'SỞ GD&ĐT TP HCM',
    truong = 'TRƯỜNG THPT NGUYỄN HỮU HUÂN',
    kiKiemTra = 'GIỮA HK 1',
    namHoc = '2026 - 2027',
    mon = 'VẬT LÝ',
    thoiGian = '45 phút'
  } = headerInfo;

  const isAnswerKey = !!options.isAnswerKey;
  const examCode = shuffledExam.examCode;
  const FONT_FAMILY = 'Times New Roman';
  const COLOR_RED = 'FF0000';
  const COLOR_BLACK = '000000';
  const COLOR_BLUE = '0000FF';

  // 1. Bảng đầu đề thi (1 hàng, 2 cột không viền)
  const headerTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.NONE },
      bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE }
    },
    rows: [
      new TableRow({
        children: [
          // Cột trái: Sở GD&ĐT, Tên trường (đỏ, đậm); vạch phân cách; "(Đề thi có ___ trang)"
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE }
            },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 30 },
                children: [
                  new TextRun({
                    text: (soGD || '').toUpperCase(),
                    bold: true,
                    color: COLOR_BLACK,
                    font: FONT_FAMILY,
                    size: 24 // 12pt
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 30 },
                children: [
                  new TextRun({
                    text: (truong || '').toUpperCase(),
                    bold: true,
                    color: COLOR_BLACK,
                    font: FONT_FAMILY,
                    size: 24
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 30 },
                children: [
                  new TextRun({
                    text: '--------------------',
                    font: FONT_FAMILY,
                    size: 20
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 60 },
                children: [
                  new TextRun({
                    text: '(Đề thi có ',
                    italics: true,
                    font: FONT_FAMILY,
                    size: 22 // 11pt
                  }),
                  new TextRun({
                    children: [PageNumber.TOTAL_PAGES],
                    italics: true,
                    font: FONT_FAMILY,
                    size: 22
                  }),
                  new TextRun({
                    text: ' trang)',
                    italics: true,
                    font: FONT_FAMILY,
                    size: 22
                  })
                ]
              })
            ]
          }),

          // Cột phải: Kì kiểm tra (đen, đậm); Năm học; Môn; Thời gian làm bài; "(không kể thời gian phát đề)"
          new TableCell({
            width: { size: 50, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.NONE },
              bottom: { style: BorderStyle.NONE },
              left: { style: BorderStyle.NONE },
              right: { style: BorderStyle.NONE }
            },
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 30 },
                children: [
                  new TextRun({
                    text: (kiKiemTra || '').toUpperCase(),
                    bold: true,
                    color: COLOR_BLACK,
                    font: FONT_FAMILY,
                    size: 24
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 30 },
                children: [
                  new TextRun({
                    text: 'NĂM HỌC ',
                    bold: true,
                    color: COLOR_BLACK,
                    font: FONT_FAMILY,
                    size: 22
                  }),
                  new TextRun({
                    text: (namHoc || '').trim(),
                    bold: true,
                    color: COLOR_BLACK,
                    font: FONT_FAMILY,
                    size: 22
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 30 },
                children: [
                  new TextRun({
                    text: 'MÔN: ',
                    bold: true,
                    color: COLOR_BLACK,
                    font: FONT_FAMILY,
                    size: 24
                  }),
                  new TextRun({
                    text: (mon || '').toUpperCase(),
                    bold: true,
                    color: COLOR_BLACK,
                    font: FONT_FAMILY,
                    size: 24
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 20 },
                children: [
                  new TextRun({
                    text: 'Thời gian làm bài: ',
                    italics: true,
                    color: COLOR_BLACK,
                    font: FONT_FAMILY,
                    size: 22
                  }),
                  new TextRun({
                    text: (thoiGian || '').trim(),
                    bold: true,
                    color: COLOR_BLACK,
                    font: FONT_FAMILY,
                    size: 22
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 60 },
                children: [
                  new TextRun({
                    text: '(không kể thời gian phát đề)',
                    italics: true,
                    color: COLOR_BLACK,
                    font: FONT_FAMILY,
                    size: 20
                  })
                ]
              })
            ]
          })
        ]
      })
    ]
  });

  // 2. Dòng Họ tên, Số báo danh, Mã đề kèm vạch kẻ ngang
  const candidateInfoParagraph = new Paragraph({
    spacing: { before: 100, after: 120 },
    border: {
      bottom: {
        color: '333333',
        space: 6,
        style: BorderStyle.SINGLE,
        size: 10
      }
    },
    children: isAnswerKey ? [
      new TextRun({
        text: 'Họ và tên: ................................................................  ',
        font: FONT_FAMILY,
        size: 24
      }),
      new TextRun({
        text: 'Số báo danh: .......................  ',
        font: FONT_FAMILY,
        size: 24
      }),
      new TextRun({
        text: 'Mã đề ',
        bold: true,
        font: FONT_FAMILY,
        size: 24
      }),
      new TextRun({
        text: `${String(examCode)} (ĐÁP ÁN)`,
        bold: true,
        color: COLOR_BLUE,
        font: FONT_FAMILY,
        size: 24
      })
    ] : [
      new TextRun({
        text: 'Họ và tên: ................................................................  ',
        font: FONT_FAMILY,
        size: 24
      }),
      new TextRun({
        text: 'Số báo danh: .......................  ',
        font: FONT_FAMILY,
        size: 24
      }),
      new TextRun({
        text: 'Mã đề ',
        bold: true,
        font: FONT_FAMILY,
        size: 24
      }),
      new TextRun({
        text: String(examCode),
        bold: true,
        color: COLOR_BLACK,
        font: FONT_FAMILY,
        size: 24
      })
    ]
  });

  // 3. Nội dung phần thân đề (Danh sách câu hỏi & đáp án)
  const contentParagraphs = [];
  const imageMap = shuffledExam.imageMap || headerInfo.imageMap || {};

  // Hàm hỗ trợ tách văn bản kèm định dạng <b>, <i>, <u>, [IMG:...] và ngắt dòng
  function appendContentRuns(targetRuns, text, forceBold = false, forceColor = COLOR_BLACK) {
    if (!text) return;
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
        targetRuns.push(new TextRun({ break: 1, font: FONT_FAMILY, size: 24 }));
      } else if (/^\[IMG:([^\]]+)\]$/i.test(part)) {
        const imgMatch = part.match(/^\[IMG:([^\]]+)\]$/i);
        const imgId = imgMatch[1];
        const imgInfo = imageMap[imgId];
        if (imgInfo && imgInfo.buffer) {
          let w = imgInfo.width || 180;
          let h = imgInfo.height || 100;
          const MAX_DOCX_WIDTH = 500;
          const MAX_DOCX_HEIGHT = 450;
          if (w > MAX_DOCX_WIDTH) {
            const scale = MAX_DOCX_WIDTH / w;
            w = MAX_DOCX_WIDTH;
            h = Math.max(10, Math.round(h * scale));
          }
          if (h > MAX_DOCX_HEIGHT) {
            const scale = MAX_DOCX_HEIGHT / h;
            h = MAX_DOCX_HEIGHT;
            w = Math.max(10, Math.round(w * scale));
          }

          targetRuns.push(new ImageRun({
            data: imgInfo.buffer,
            transformation: {
              width: w,
              height: h
            }
          }));
        }
      } else {
        targetRuns.push(new TextRun({
          text: part,
          font: FONT_FAMILY, // Times New Roman
          size: 24,          // 12pt
          bold: forceBold || isBold,
          color: forceColor,
          italics: isItalic,
          underline: isUnderline ? { type: UnderlineType.SINGLE } : undefined
        }));
      }
    }
  }

  // Duyệt theo từng nhóm để đảm bảo có khoảng trống rõ ràng giữa các câu hỏi và các nhóm
  const groups = (shuffledExam.groups && shuffledExam.groups.length > 0)
    ? shuffledExam.groups
    : [{ questions: shuffledExam.flatQuestions || [] }];

  groups.forEach((group, gIdx) => {
    const isFirstGroup = (gIdx === 0);

    // Xuất phần tiêu đề nhóm (nằm giữa marker nhóm và __QUESTION__) nếu có
    if (group.title && group.title.trim()) {
      const roman = toRoman(gIdx + 1);
      const titleLines = group.title.trim().split(/\r?\n/);
      titleLines.forEach((tLine, tIdx) => {
        let trimmedTLine = tLine.trim();
        if (!trimmedTLine) return;
        if (tIdx === 0) {
          trimmedTLine = formatGroupTitleFirstLine(trimmedTLine, roman);
        }
        const titleRuns = [];
        appendContentRuns(titleRuns, trimmedTLine, false, COLOR_BLACK);

        contentParagraphs.push(new Paragraph({
          spacing: { before: (isFirstGroup && tIdx === 0) ? 140 : 240, after: 60 },
          children: titleRuns
        }));
      });
    }

    group.questions.forEach((q, qIdxInGroup) => {
      const isFirstQInGroup = (qIdxInGroup === 0);

      // Đoạn câu hỏi với "Câu 1.", "Câu 1:" được IN ĐẬM và chuẩn 12pt Times New Roman
      const questionRuns = [
        new TextRun({
          text: `${q.displayPrefix} `,
          bold: true, // IN ĐẬM "Câu 1.", "Câu 2."
          font: FONT_FAMILY,
          size: 24 // 12pt
        })
      ];

      appendContentRuns(questionRuns, q.text, false, COLOR_BLACK);

      // Khoảng cách: Nếu nhóm chưa có tiêu đề thì trước câu đầu tiên của nhóm mới (từ nhóm 2 trở đi) có khoảng trống 260 twips
      // Nếu nhóm đã có tiêu đề thì câu đầu tiên giãn 140 twips chuẩn
      const hasGroupTitle = (group.title && group.title.trim());
      const beforeSpacing = (!isFirstGroup && isFirstQInGroup && !hasGroupTitle) ? 260 : 140;

      contentParagraphs.push(new Paragraph({
        spacing: { before: beforeSpacing, after: 60 },
        children: questionRuns
      }));

      // Đoạn đáp án
      if (q.type === 'short_answer') {
        // Trong đề thi học sinh: câu trả lời ngắn không hiển thị đáp án (mục 5.2)
        // Trong file đáp án: hiển thị đáp án đúng (in đậm và có màu blue)
        if (isAnswerKey) {
          const correctAns = q.answers.find(a => a.correct) || q.answers[0];
          if (correctAns) {
            const answerRuns = [
              new TextRun({
                text: 'Đáp án: ',
                bold: true,
                color: COLOR_BLUE,
                font: FONT_FAMILY,
                size: 24
              })
            ];
            appendContentRuns(answerRuns, correctAns.text, true, COLOR_BLUE);

            contentParagraphs.push(new Paragraph({
              indent: { left: 400 },
              spacing: { before: 30, after: 30 },
              children: answerRuns
            }));
          }
        }
      } else {
        for (const ans of q.answers) {
          const isCorrectChoice = isAnswerKey && !!ans.correct;
          const choiceColor = isCorrectChoice ? COLOR_BLUE : COLOR_BLACK;

          const answerRuns = [
            new TextRun({
              text: `${ans.labelDisplay} `,
              bold: true, // IN ĐẬM "A.", "B.", "A)", "B)"
              color: choiceColor,
              font: FONT_FAMILY,
              size: 24 // 12pt
            })
          ];

          appendContentRuns(answerRuns, ans.text, isCorrectChoice, choiceColor);

          contentParagraphs.push(new Paragraph({
            indent: { left: 400 }, // Thụt đầu dòng cho các phương án đáp án
            spacing: { before: 30, after: 30 },
            children: answerRuns
          }));
        }
      }
    });
  });

  // Đoạn kết thúc đề
  contentParagraphs.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 240, after: 100 },
    children: [
      new TextRun({
        text: isAnswerKey
          ? `----------------------- HẾT (ĐÁP ÁN MÃ ĐỀ ${examCode}) -----------------------`
          : '----------------------- HẾT -----------------------',
        bold: true,
        color: isAnswerKey ? COLOR_BLUE : COLOR_BLACK,
        font: FONT_FAMILY,
        size: 24 // 12pt
      })
    ]
  }));

  // Tạo tài liệu Word
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: FONT_FAMILY,
            size: 24
          }
        }
      }
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 1000,
              bottom: 1000,
              left: 1200,
              right: 1200
            }
          }
        },
        children: [
          headerTable,
          candidateInfoParagraph,
          ...contentParagraphs
        ]
      }
    ]
  });

  return await Packer.toBlob(doc);
}

// --- FILE: js/answer-key.js ---
/**
 * Module Tạo Bảng Đáp Án (Answer Key)
 * Tuân thủ quy định tại mục 5.4 của tài liệu:
 * "Công cụ trộn đề trắc nghiệm Yêu cầu và giải pháp.docx"
 *
 * - Tự suy ra từ dữ liệu:
 *   + Một __D__: ghi nhãn (ví dụ C).
 *   + Nhiều __D__: ghi tập hợp (ví dụ A, C); nếu là nhãn a) b) c) d) ghi từng ý (ví dụ a-Đ, b-S, c-Đ, d-S).
 *   + Câu trả lời ngắn (chỉ có __D__): ghi nguyên giá trị chuỗi đáp án (ví dụ 4).
 * - Bảng ma trận: mỗi hàng là một câu, mỗi cột là một mã đề thi.
 * - Xuất thành một file .docx riêng biệt.
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
 * Làm sạch văn bản đáp án hiển thị trong Bảng Đáp Án Ma Trận
 * - Bỏ toàn bộ thẻ HTML như <b>, </b>, <i>, </i>, <u>, </u>
 * - Bỏ các dòng kết thúc đề thi nếu có
 */
function cleanAnswerText(text) {
  if (!text) return '';
  let cleaned = text;
  // Bỏ các dòng kết đề nếu có
  cleaned = cleaned.replace(/[-–—=_*~.\s]*(?:HẾT|THE\s+END)(?:\s+(?:ĐỀ|BÀI)?\s*(?:THI|KIỂM\s+TRA)?)?[-–—=_*~.\s]*/gi, '');
  // Loại bỏ các thẻ HTML như <b>, </b>, <i>, </i>, <u>, </u>
  cleaned = cleaned.replace(/<[^>]+>/g, '').trim();
  return cleaned;
}

/**
 * Trích xuất giá trị đáp án đúng của một câu hỏi đã hoán vị
 * @param {Object} q Câu hỏi đã hoán vị
 * @returns {string} Chuỗi hiển thị đáp án
 */
function extractQuestionAnswerKey(q) {
  const correctAnswers = q.answers.filter(a => a.correct);

  // 1. Câu trả lời ngắn (chỉ có 1 __D__ và không có __S__)
  if (q.type === 'short_answer' || (q.answers.length === 1 && correctAnswers.length === 1)) {
    return cleanAnswerText(correctAnswers[0].text);
  }

  // 2. Định dạng nhãn a) b) c) d) cho câu đúng/sai nhiều ý
  if (q.labelStyle === 'abcd') {
    return q.answers.map(ans => `${ans.labelRaw}-${ans.correct ? 'Đ' : 'S'}`).join(', ');
  }

  // 3. Một đáp án đúng duy nhất (__D__)
  if (correctAnswers.length === 1) {
    return correctAnswers[0].labelRaw;
  }

  // 4. Nhiều đáp án đúng (__D__)
  return correctAnswers.map(a => a.labelRaw).join(', ');
}

/**
 * Sinh cấu trúc ma trận bảng đáp án (để xem trước trên UI hoặc xuất DOCX)
 * @param {Array<Object>} allShuffledExams Danh sách tất cả các mã đề thi đã hoán vị
 * @returns {Object} { headers: Array<string>, rows: Array<Object> }
 */
function buildAnswerMatrix(allShuffledExams) {
  if (!allShuffledExams || allShuffledExams.length === 0) {
    return { headers: [], rows: [] };
  }

  const examCodes = allShuffledExams.map(e => e.examCode);
  const totalQuestions = allShuffledExams[0].flatQuestions.length;
  const rows = [];

  for (let i = 0; i < totalQuestions; i++) {
    const qNum = allShuffledExams[0].flatQuestions[i].displayNumber || (i + 1);
    const row = {
      questionNumber: qNum,
      answers: {}
    };

    allShuffledExams.forEach(exam => {
      const q = exam.flatQuestions[i];
      row.answers[exam.examCode] = extractQuestionAnswerKey(q);
    });

    rows.push(row);
  }

  return {
    headers: examCodes,
    rows
  };
}

/**
 * Tạo Blob file Word (.docx) chứa Bảng Đáp Án Ma Trận
 * @param {Array<Object>} allShuffledExams
 * @param {Object} headerInfo
 * @returns {Promise<Blob>}
 */
async function createAnswerKeyDocxBlob(allShuffledExams, headerInfo = {}) {
  const docx = getDocx();
  if (!docx) {
    throw new Error('Thư viện docx chưa được nạp!');
  }

  const {
    Document,
    Paragraph,
    TextRun,
    Table,
    TableRow,
    TableCell,
    WidthType,
    BorderStyle,
    AlignmentType,
    ShadingType,
    Packer
  } = docx;

  const {
    soGD = 'SỞ GD&ĐT TP HCM',
    truong = 'TRƯỜNG THPT NGUYỄN HỮU HUÂN',
    kiKiemTra = 'GIỮA HK 1',
    namHoc = '2026 - 2027',
    mon = 'VẬT LÝ'
  } = headerInfo;

  const FONT_FAMILY = 'Times New Roman';
  const matrix = buildAnswerMatrix(allShuffledExams);

  // Tiêu đề trang đáp án
  const titleParagraphs = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 40 },
      children: [
        new TextRun({
          text: `${(soGD || '').toUpperCase()} - ${(truong || '').toUpperCase()}`,
          bold: true,
          font: FONT_FAMILY,
          size: 22
        })
      ]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 40 },
      children: [
        new TextRun({
          text: `BẢNG ĐÁP ÁN ${(kiKiemTra || '').toUpperCase()} - NĂM HỌC ${(namHoc || '')}`.trim(),
          bold: true,
          color: '000000',
          font: FONT_FAMILY,
          size: 26
        })
      ]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 140 },
      children: [
        new TextRun({
          text: `MÔN: ${(mon || '').toUpperCase()}`,
          bold: true,
          font: FONT_FAMILY,
          size: 24
        })
      ]
    })
  ];

  // Bảng đáp án
  const tableBorder = {
    color: '555555',
    size: 4,
    style: BorderStyle.SINGLE
  };

  const bordersConfig = {
    top: tableBorder,
    bottom: tableBorder,
    left: tableBorder,
    right: tableBorder,
    insideHorizontal: tableBorder,
    insideVertical: tableBorder
  };

  // Hàng tiêu đề bảng: Câu | Mã 101 | Mã 102 | ...
  const headerCells = [
    new TableCell({
      width: { size: 15, type: WidthType.PERCENTAGE },
      shading: { type: ShadingType.CLEAR, fill: 'EEEEEE' },
      borders: bordersConfig,
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              text: 'Câu',
              bold: true,
              font: FONT_FAMILY,
              size: 22
            })
          ]
        })
      ]
    })
  ];

  const colWidthPercent = 85 / (matrix.headers.length || 1);

  matrix.headers.forEach(code => {
    headerCells.push(
      new TableCell({
        width: { size: colWidthPercent, type: WidthType.PERCENTAGE },
        shading: { type: ShadingType.CLEAR, fill: 'EEEEEE' },
        borders: bordersConfig,
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: `Mã ${code}`,
                bold: true,
                color: '000000',
                font: FONT_FAMILY,
                size: 22
              })
            ]
          })
        ]
      })
    );
  });

  const tableRows = [
    new TableRow({
      tableHeader: true,
      children: headerCells
    })
  ];

  // Các hàng dữ liệu: từng câu hỏi
  matrix.rows.forEach((row, rIdx) => {
    const isAlt = (rIdx % 2 === 1);
    const fillBg = isAlt ? 'F9FAFB' : 'FFFFFF';

    const rowCells = [
      new TableCell({
        width: { size: 15, type: WidthType.PERCENTAGE },
        shading: { type: ShadingType.CLEAR, fill: fillBg },
        borders: bordersConfig,
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: `Câu ${row.questionNumber}`,
                bold: true,
                font: FONT_FAMILY,
                size: 22
              })
            ]
          })
        ]
      })
    ];

    matrix.headers.forEach(code => {
      rowCells.push(
        new TableCell({
          width: { size: colWidthPercent, type: WidthType.PERCENTAGE },
          shading: { type: ShadingType.CLEAR, fill: fillBg },
          borders: bordersConfig,
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: row.answers[code] || '',
                  bold: true,
                  font: FONT_FAMILY,
                  size: 22
                })
              ]
            })
          ]
        })
      );
    });

    tableRows.push(new TableRow({ children: rowCells }));
  });

  const answerTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: bordersConfig,
    rows: tableRows
  });

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: FONT_FAMILY,
            size: 22
          }
        }
      }
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 1000,
              bottom: 1000,
              left: 1200,
              right: 1200
            }
          }
        },
        children: [
          ...titleParagraphs,
          answerTable
        ]
      }
    ]
  });

  return await Packer.toBlob(doc);
}

// --- FILE: js/ui.js ---
/**
 * Module Giao Diện Người Dùng (UI Controller)
 * Quản lý tương tác nạp file Word (.docx), cấu hình đầu đề, tùy chọn trộn,
 * hiển thị xem trước trực quan và xử lý xuất file DOCX / ZIP.
 */







// Trạng thái ứng dụng
const state = {
  rawText: '',
  fileName: '',
  fileSize: 0,
  imageMap: {},
  parsedExam: null,
  examCodes: ['101', '102', '103', '104'],
  generatedZipBlob: null,
  generatedFiles: [] // { name, blob, url }
};

// Hàm hiển thị thông báo Snackbar
function showSnackbar(message, duration = 3000) {
  const snackbar = document.getElementById('snackbar');
  if (!snackbar) return;
  snackbar.textContent = message;
  snackbar.classList.add('show');
  setTimeout(() => {
    snackbar.classList.remove('show');
  }, duration);
}

// Format dung lượng tệp
function formatBytes(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

/**
 * Cập nhật danh sách mã đề hiển thị theo số lượng đề hoặc mã bắt đầu
 */
function updateExamCodes(count, startCode) {
  const numericStart = parseInt(startCode, 10) || 101;
  const newCodes = [];

  for (let i = 0; i < count; i++) {
    // Nếu người dùng đã tự sửa mã ở vị trí i thì ưu tiên giữ lại, nếu không thì tăng dần
    if (state.examCodes[i] !== undefined) {
      newCodes.push(state.examCodes[i]);
    } else {
      newCodes.push(String(numericStart + i));
    }
  }

  state.examCodes = newCodes;
  renderExamCodesChips();
}

/**
 * Hiển thị danh sách chip mã đề thi tương tác
 */
function renderExamCodesChips() {
  const container = document.getElementById('examCodesContainer');
  if (!container) return;

  container.innerHTML = '';
  state.examCodes.forEach((code, index) => {
    const chip = document.createElement('div');
    chip.className = 'code-chip';
    chip.title = `Bấm để đổi mã đề số ${index + 1}`;
    chip.innerHTML = `<span>Mã:</span> <strong>${code}</strong>`;

    chip.addEventListener('click', () => {
      // Chuyển sang ô input để chỉnh sửa trực tiếp
      const input = document.createElement('input');
      input.type = 'text';
      input.className = 'code-chip-input';
      input.value = code;
      input.maxLength = 6;

      chip.replaceWith(input);
      input.focus();
      input.select();

      const commitChange = () => {
        const val = input.value.trim() || code;
        state.examCodes[index] = val;
        renderExamCodesChips();
      };

      input.addEventListener('blur', commitChange);
      input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          commitChange();
        } else if (e.key === 'Escape') {
          renderExamCodesChips();
        }
      });
    });

    container.appendChild(chip);
  });
}

/**
 * Hiển thị xem trước cấu trúc đề thi
 */
function renderExamPreview(parsedExam) {
  const container = document.getElementById('previewQuestionsContainer');
  if (!container) return;

  if (!parsedExam || parsedExam.groups.length === 0) {
    container.innerHTML = `
      <div class="empty-state">
        <p class="empty-title">Chưa nạp đề thi</p>
        <p class="empty-desc">Kéo thả file Word (.docx) ở Bước 1 để phân tích và xem trước đề thi.</p>
      </div>
    `;
    return;
  }

  let html = '';
  let questionCounter = 0;
  const imageMap = parsedExam.imageMap || state.imageMap || {};

  function renderFormattedContent(text) {
    if (!text) return '';
    let formatted = escapeHtml(text);
    // Phục hồi an toàn các thẻ định dạng <b>, <i>, <u> để xem trước trực quan
    formatted = formatted
      .replace(/&lt;b&gt;/gi, '<strong class="q-bold">')
      .replace(/&lt;\/b&gt;/gi, '</strong>')
      .replace(/&lt;i&gt;/gi, '<em class="q-italic">')
      .replace(/&lt;\/i&gt;/gi, '</em>')
      .replace(/&lt;u&gt;/gi, '<span class="q-underline">')
      .replace(/&lt;\/u&gt;/gi, '</span>');

    // Giữ các dấu xuống hàng tự nhiên
    const withBreaks = formatted.replace(/\n/g, '<br/>');
    return withBreaks.replace(/\[IMG:([^\]]+)\]/g, (match, id) => {
      const img = imageMap[id];
      if (img && img.dataUrl) {
        const w = img.width || 180;
        const h = img.height || 100;
        const ratio = (w / h).toFixed(2);
        return `<span class="exam-img-frame" title="Hình ảnh từ file Word gốc: ${w} × ${h} px (Tỉ lệ: ${ratio})">` +
          `<img src="${img.dataUrl}" alt="Hình ảnh đề thi" class="exam-preview-img" ` +
          `style="width: ${w}px; max-width: 100%; aspect-ratio: ${w} / ${h};" ` +
          `data-width="${w}" data-height="${h}" data-img-id="${id}" loading="lazy" />` +
          `</span>`;
      }
      return '';
    });
  }

  const labelSelect = typeof document !== 'undefined' ? document.getElementById('selDefaultLabelStyle') : null;
  const currentLabelStyle = labelSelect ? labelSelect.value : 'ABCD';

  const radioNumbering = typeof document !== 'undefined' ? document.querySelector('input[name="radioNumbering"]:checked') : null;
  const currentNumberingFormat = radioNumbering ? radioNumbering.value : 'dot';

  const startNumInput = typeof document !== 'undefined' ? document.getElementById('inpStartQNum') : null;
  const startQNum = (startNumInput && parseInt(startNumInput.value, 10)) || 1;

  parsedExam.groups.forEach((group, gIdx) => {
    const modeBadge = group.mode || 'G';
    const isFixedGroup = group.fixed;
    const modeDesc = {
      'G': 'Mix cả câu hỏi & đáp án',
      'GQ': 'Chỉ mix câu hỏi',
      'GA': 'Chỉ mix đáp án',
      'G0': 'Không đảo vị trí'
    }[modeBadge] || 'Mix đề';

    html += `
      <div class="preview-group-box">
        <div class="preview-group-header">
          <div>
            <span>Nhóm ${gIdx + 1}</span>
            <span class="group-mode-pill">${group.rawMarker} - ${modeDesc}</span>
            ${isFixedGroup ? '<span class="badge badge-success" style="margin-left: 6px;">Vị trí cố định (#)</span>' : ''}
          </div>
          <div style="font-size: 0.8125rem; color: var(--text-muted); font-weight: 500;">
            ${group.questions.length} câu hỏi
          </div>
        </div>
        ${group.title && group.title.trim() ? `
          <div class="preview-group-title" style="padding: 10px 16px; background: #f8fafc; border-bottom: 1px solid var(--border-light); font-size: 0.9375rem; color: #1e293b; line-height: 1.5; white-space: pre-line;">
            ${renderFormattedContent(formatGroupTitle(group.title.trim(), gIdx + 1))}
          </div>
        ` : ''}
        <div class="preview-questions-list">
    `;

    group.questions.forEach((q) => {
      questionCounter++;
      const currentQNumber = startQNum + questionCounter - 1;
      let qNumDisplay = `Câu ${currentQNumber}.`;
      if (currentNumberingFormat === 'colon') {
        qNumDisplay = `Câu ${currentQNumber}:`;
      } else if (currentNumberingFormat === 'paren') {
        qNumDisplay = `Câu ${currentQNumber})`;
      }

      html += `
        <div class="question-card">
          <div class="question-title-line">
            <span class="question-qnum"><strong>${qNumDisplay}</strong></span>
            <span>${renderFormattedContent(q.text)}</span>
          </div>
          <div class="answers-grid">
      `;

      if (q.type === 'short_answer') {
        html += `
          <div class="answer-item">
            <span class="ans-pill d">Trả lời ngắn</span>
            <span class="answer-text"><strong>Đáp án:</strong> ${renderFormattedContent(q.answers[0].text)}</span>
          </div>
        `;
      } else {
        q.answers.forEach((ans, aIdx) => {
          const isD = ans.correct;
          const isFixed = ans.fixed;

          let labelDisplay = `${String.fromCharCode(65 + (aIdx % 26))}.`;
          if (currentLabelStyle === 'ABCD_paren') {
            labelDisplay = `${String.fromCharCode(65 + (aIdx % 26))})`;
          } else if (currentLabelStyle === 'abcd') {
            labelDisplay = `${String.fromCharCode(97 + (aIdx % 26))})`;
          } else if (currentLabelStyle === '1234') {
            labelDisplay = `${aIdx + 1}.`;
          }

          html += `
            <div class="answer-item">
              <span class="ans-pill ${isD ? 'd' : 's'}">
                ${isD ? '__D__' : '__S__'}
                ${isFixed ? ' <span class="lock-icon" title="Cố định vị trí (#)">🔒#</span>' : ''}
              </span>
              <span class="answer-text"><strong>${labelDisplay}</strong> ${renderFormattedContent(ans.text)}</span>
            </div>
          `;
        });
      }

      html += `
          </div>
        </div>
      `;
    });

    html += `
        </div>
      </div>
    `;
  });

  container.innerHTML = html;
}

function escapeHtml(str) {
  if (!str) return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * Xử lý khi nạp file Word thành công
 */
async function handleFileLoaded(file) {
  try {
    showSnackbar('Đang đọc và phân tích file Word...');
    const result = await readExamDocx(file);

    state.rawText = result.text;
    state.fileName = result.fileName;
    state.fileSize = result.fileSize;
    state.imageMap = result.imageMap || {};

    // Cập nhật giao diện thông tin file
    document.getElementById('txtFileName').textContent = result.fileName;
    document.getElementById('txtFileSize').textContent = formatBytes(result.fileSize);
    document.getElementById('loadedFileCard').classList.remove('hidden');

    // Cập nhật text xem trước thô
    document.getElementById('txtRawExtracted').textContent = result.text;
    document.getElementById('rawTextMeta').textContent = `${result.text.split('\n').length} dòng`;

    // Điền trước thông tin đầu đề nếu trích xuất được từ file Word của người dùng (không áp dụng cho file mẫu)
    if (result.extractedHeader && !file.isSample) {
      const h = result.extractedHeader;
      const fillIfEmpty = (id, val) => {
        const el = document.getElementById(id);
        if (el && !el.value.trim() && val) {
          el.value = val;
        }
      };
      if (h.soGD) fillIfEmpty('inpSoGD', h.soGD);
      if (h.truong) fillIfEmpty('inpTruong', h.truong);
      if (h.kiKiemTra) fillIfEmpty('inpKiKiemTra', h.kiKiemTra);
      if (h.namHoc) fillIfEmpty('inpNamHoc', h.namHoc);
      if (h.mon) fillIfEmpty('inpMon', h.mon);
      if (h.thoiGian) fillIfEmpty('inpThoiGian', h.thoiGian);
    }

    // Cập nhật trạng thái warning và cookie cho các ô thông tin đầu đề
    if (typeof window.updateHeaderValidationUI === 'function') {
      window.updateHeaderValidationUI();
    }

    // Phân tích cú pháp Marker theo tiêu chuẩn Mục 3 & 7
    const parsed = parseExam(result.text);
    parsed.imageMap = state.imageMap;
    state.parsedExam = parsed;

    const errorBanner = document.getElementById('errorBanner');
    const errorList = document.getElementById('errorList');
    const statusBadge = document.getElementById('txtFileStatusBadge');
    const btnMixAndExport = document.getElementById('btnMixAndExport');
    const txtFileStats = document.getElementById('txtFileStats');

    // Kiểm tra định dạng hình ảnh (Chỉ hỗ trợ In Line with Text)
    const nonInlineImages = result.nonInlineImages || [];
    state.nonInlineImages = nonInlineImages;

    const imageWarningBanner = document.getElementById('imageWarningBanner');
    const imageWarningList = document.getElementById('imageWarningList');

    if (imageWarningBanner && imageWarningList) {
      if (nonInlineImages.length > 0) {
        imageWarningBanner.classList.remove('hidden');
        imageWarningList.innerHTML = nonInlineImages.map((img) => `
          <li class="image-warning-item">
            <div class="image-warning-thumb-box">
              ${img.dataUrl
                ? `<img src="${img.dataUrl}" alt="${escapeHtml(img.name)}" class="image-warning-thumb"/>`
                : `<div class="image-warning-thumb-placeholder">📷</div>`}
            </div>
            <div class="image-warning-details">
              <div class="image-warning-name-row">
                <strong class="image-warning-name">${escapeHtml(img.name)}</strong>
                <span class="badge" style="background:#fee2e2; color:#b91c1c; border:1px solid #fca5a5;">
                  Wrap Text hiện tại: ${escapeHtml(img.wrapType)}
                </span>
              </div>
              <div class="image-warning-snippet">
                <span class="snippet-label">Đoạn văn chứa ảnh:</span> "${escapeHtml(img.snippet)}"
              </div>
              <div class="image-warning-action-tip">
                👉 Vui lòng mở file Word gốc ➔ Click chuột phải vào ảnh này ➔ <strong>Wrap Text</strong> ➔ Chọn <strong>In Line with Text</strong>.
              </div>
            </div>
          </li>
        `).join('');
      } else {
        imageWarningBanner.classList.add('hidden');
        imageWarningList.innerHTML = '';
      }
    }

    if (!parsed.valid) {
      const errorMsg = nonInlineImages.length > 0
        ? `⚠ ${parsed.errors.length} lỗi cú pháp • ${nonInlineImages.length} ảnh sai InLine`
        : `⚠ ${parsed.errors.length} lỗi cú pháp`;
      statusBadge.className = 'badge badge-danger';
      statusBadge.textContent = errorMsg;
      errorBanner.classList.remove('hidden');
      errorList.innerHTML = parsed.errors.map(err => `<li>${escapeHtml(err.message)}</li>`).join('');
      btnMixAndExport.disabled = true;
      txtFileStats.textContent = 'Phát hiện lỗi cú pháp và định dạng, hãy sửa theo hướng dẫn trên';
      renderExamPreview(null);
      showSnackbar('Phát hiện lỗi cú pháp trong đề thi. Vui lòng kiểm tra!', 4000);
    } else if (nonInlineImages.length > 0) {
      statusBadge.className = 'badge badge-danger';
      statusBadge.textContent = `⚠ ${nonInlineImages.length} ảnh không phải InLine`;
      errorBanner.classList.add('hidden');
      errorList.innerHTML = '';
      btnMixAndExport.disabled = true;
      txtFileStats.textContent = `Phát hiện ${nonInlineImages.length} ảnh không phải InLine. Cần sửa file gốc sang "In Line with Text" trước khi trộn.`;
      renderExamPreview(parsed);
      showSnackbar(`⚠️ Có ${nonInlineImages.length} ảnh không phải InLine! Vui lòng sửa lại file gốc.`, 5000);
    } else {
      statusBadge.className = 'badge badge-success';
      statusBadge.textContent = '✓ Hợp lệ';
      errorBanner.classList.add('hidden');
      errorList.innerHTML = '';
      btnMixAndExport.disabled = false;

      const { totalGroups, totalQuestions, choiceCount, tfCount, shortCount } = parsed.stats;
      txtFileStats.textContent = `${totalGroups} nhóm • ${totalQuestions} câu hỏi (${choiceCount} trắc nghiệm, ${tfCount} đúng/sai, ${shortCount} trả lời ngắn)`;

      renderExamPreview(parsed);
      showSnackbar(`Đã nạp thành công đề thi: ${totalQuestions} câu hỏi!`, 3000);
    }
  } catch (err) {
    console.error('Lỗi khi nạp file Word:', err);
    showSnackbar(`Lỗi khi đọc file: ${err.message}`, 4000);
  }
}

/**
 * Lấy các thông số cấu hình đầu đề từ giao diện
 */
function getHeaderConfig() {
  const getVal = (id) => {
    const el = document.getElementById(id);
    return el ? el.value.trim() : '';
  };

  return {
    soGD: getVal('inpSoGD'),
    truong: getVal('inpTruong'),
    kiKiemTra: getVal('inpKiKiemTra'),
    namHoc: getVal('inpNamHoc'),
    mon: getVal('inpMon'),
    thoiGian: getVal('inpThoiGian')
  };
}

/**
 * Lấy các tùy chọn trộn đề
 */
function getMixingOptions() {
  const radioNumbering = document.querySelector('input[name="radioNumbering"]:checked');
  const numberingFormat = radioNumbering ? radioNumbering.value : 'dot';
  const startQNum = parseInt(document.getElementById('inpStartQNum').value, 10) || 1;
  const defaultLabelStyle = document.getElementById('selDefaultLabelStyle').value || 'ABCD';
  const shuffleGroups = document.getElementById('chkShuffleGroups').checked;

  return {
    numberingFormat,
    startQuestionNumber: startQNum,
    defaultLabelStyle,
    shuffleGroups
  };
}

/**
 * Thực hiện Trộn & Tạo các file Word DOCX và ZIP
 */
async function executeMixAndExport() {
  if (!state.parsedExam || !state.parsedExam.valid) {
    showSnackbar('Vui lòng nạp file đề Word hợp lệ trước khi trộn!');
    return;
  }

  const btnMix = document.getElementById('btnMixAndExport');
  const originalText = btnMix.innerHTML;

  if (state.nonInlineImages && state.nonInlineImages.length > 0) {
    showSnackbar(`⚠️ File còn ${state.nonInlineImages.length} hình ảnh chưa chuyển sang In Line with Text! Vui lòng chỉnh sửa file gốc.`, 5000);
    return;
  }

  try {
    btnMix.disabled = true;
    btnMix.innerHTML = `
      <span class="spin-icon" style="display:inline-block; margin-right: 8px;">🔄</span>
      <span>Đang tạo các file Word...</span>
    `;

    const headerInfo = getHeaderConfig();
    headerInfo.imageMap = state.imageMap || {};
    const mixingOptions = getMixingOptions();

    // Cảnh báo nếu các mục thông tin đầu đề thi chưa được nhập
    const emptyFields = [];
    if (!headerInfo.soGD) emptyFields.push('Sở/Phòng GD');
    if (!headerInfo.truong) emptyFields.push('Tên trường');
    if (!headerInfo.kiKiemTra) emptyFields.push('Kì kiểm tra');
    if (!headerInfo.namHoc) emptyFields.push('Năm học');
    if (!headerInfo.mon) emptyFields.push('Môn thi');
    if (!headerInfo.thoiGian) emptyFields.push('Thời gian');

    if (emptyFields.length > 0) {
      showSnackbar(`⚠️ Chú ý: Còn ${emptyFields.length} mục đầu đề chưa điền (${emptyFields.slice(0, 3).join(', ')}${emptyFields.length > 3 ? '...' : ''})!`, 3500);
    }

    // 1. Hoán vị tất cả các mã đề
    const shuffledExams = shuffleAllExams(state.parsedExam, state.examCodes, mixingOptions);
    shuffledExams.forEach(exam => {
      exam.imageMap = state.imageMap || {};
    });

    // 2. Tạo JSZip để đóng gói
    const JSZip = (typeof window !== 'undefined' && window.JSZip) ? window.JSZip : globalThis.JSZip;
    if (!JSZip) throw new Error('Thư viện JSZip chưa sẵn sàng!');
    const zip = new JSZip();

    state.generatedFiles = [];

    // 3. Tạo từng file Word cho mỗi mã đề (Đề thi & Đáp án chi tiết)
    for (const exam of shuffledExams) {
      // 3.1 File Đề thi (học sinh làm, không hiển thị đáp án)
      const docxBlob = await createExamDocxBlob(exam, headerInfo, { isAnswerKey: false });
      const examFileName = `Đề_${exam.examCode}.docx`;
      zip.file(examFileName, docxBlob);
      const examUrl = URL.createObjectURL(docxBlob);
      state.generatedFiles.push({
        name: examFileName,
        blob: docxBlob,
        url: examUrl,
        isAnswerKey: false
      });

      // 3.2 File Đáp án chi tiết tương ứng (các câu đúng in đậm và có màu blue)
      const answerDocxBlob = await createExamDocxBlob(exam, headerInfo, { isAnswerKey: true });
      const answerFileName = `Đáp_Án_${exam.examCode}.docx`;
      zip.file(answerFileName, answerDocxBlob);
      const answerUrl = URL.createObjectURL(answerDocxBlob);
      state.generatedFiles.push({
        name: answerFileName,
        blob: answerDocxBlob,
        url: answerUrl,
        isAnswerKey: true
      });
    }

    // 4. Tạo file Word Bảng Đáp Án Ma Trận (tổng hợp tất cả các mã đề)
    const answerBlob = await createAnswerKeyDocxBlob(shuffledExams, headerInfo);
    const safeMon = (headerInfo.mon || 'Mon_Hoc').trim().replace(/[\s/\\:*?"<>|]+/g, '_') || 'Mon_Hoc';
    const answerFileName = `Bảng_Đáp_Án_${safeMon}.docx`;
    zip.file(answerFileName, answerBlob);
    const answerUrl = URL.createObjectURL(answerBlob);
    state.generatedFiles.push({
      name: answerFileName,
      blob: answerBlob,
      url: answerUrl,
      isAnswerKey: true
    });

    // 5. Đóng gói ZIP
    const zipBlob = await zip.generateAsync({ type: 'blob' });
    state.generatedZipBlob = zipBlob;

    // 6. Hiển thị bảng ma trận đáp án trên Modal
    const matrix = buildAnswerMatrix(shuffledExams);
    renderAnswerMatrixTable(matrix);

    // 7. Hiển thị danh sách file tải lẻ
    renderDownloadFileList();

    // 8. Mở modal kết quả
    document.getElementById('exportModal').classList.add('active');

    // Kiểm tra khả năng Web Share API
    const btnShare = document.getElementById('btnShareFiles');
    if (navigator.canShare && navigator.canShare({ files: [new File([zipBlob], 'de_thi.zip', { type: 'application/zip' })] })) {
      btnShare.classList.remove('hidden');
    } else {
      btnShare.classList.add('hidden');
    }

    showSnackbar('Đã tạo thành công bộ đề thi!', 3000);
  } catch (err) {
    console.error('Lỗi khi trộn và xuất đề:', err);
    showSnackbar(`Lỗi khi tạo file: ${err.message}`, 4000);
  } finally {
    btnMix.disabled = false;
    btnMix.innerHTML = originalText;
  }
}

/**
 * Hiển thị bảng ma trận đáp án trên Modal
 */
function renderAnswerMatrixTable(matrix) {
  const table = document.getElementById('tblAnswerPreview');
  if (!table) return;

  let html = '<thead><tr><th>Câu</th>';
  matrix.headers.forEach(code => {
    html += `<th>Mã ${code}</th>`;
  });
  html += '</tr></thead><tbody>';

  matrix.rows.forEach(row => {
    html += `<tr><td><strong>Câu ${row.questionNumber}</strong></td>`;
    matrix.headers.forEach(code => {
      html += `<td class="correct-val">${escapeHtml(row.answers[code] || '-')}</td>`;
    });
    html += '</tr>';
  });

  html += '</tbody>';
  table.innerHTML = html;
}

/**
 * Hiển thị danh sách file riêng lẻ trên Modal
 */
function renderDownloadFileList() {
  const container = document.getElementById('fileDownloadList');
  if (!container) return;

  container.innerHTML = '';
  state.generatedFiles.forEach(file => {
    const item = document.createElement('div');
    item.className = 'file-download-item';
    item.innerHTML = `
      <div class="file-download-name">
        <span>📄</span>
        <span>${escapeHtml(file.name)}</span>
        ${file.isAnswerKey ? '<span class="badge badge-success" style="font-size:0.7rem;">Đáp án</span>' : ''}
      </div>
      <a href="${file.url}" download="${file.name}" class="btn btn-secondary btn-xs">
        Tải về
      </a>
    `;
    container.appendChild(item);
  });
}

/**
 * Tải toàn bộ file ZIP
 */
function downloadAllZip() {
  if (!state.generatedZipBlob) return;
  const headerInfo = getHeaderConfig();
  const zipName = `Bo_De_Thi_${headerInfo.mon.replace(/\s+/g, '_')}_${state.examCodes.length}_MaDe.zip`;

  const url = URL.createObjectURL(state.generatedZipBlob);
  const a = document.createElement('a');
  a.href = url;
  a.download = zipName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

// --- FILE: js/app.js ---
/**
 * PWA Mark 'n Mix - Main Application Entry
 * Khởi tạo ứng dụng, lắng nghe sự kiện, đăng ký Service Worker PWA
 */



let deferredInstallPrompt = null;

// Khởi chạy khi DOM sẵn sàng
document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  initHeaderInfoCacheAndValidation();
  initNetworkStatus();
  initPWA();
  initEventListeners();

  // Khởi tạo 4 mã đề mặc định bắt đầu từ 101
  updateExamCodes(4, '101');

  // Nạp sẵn đề mẫu tự động khi mở trang lần đầu để giáo viên xem ngay
  loadSampleExamDocx(true);
});

/**
 * Quản lý Giao diện (Duy nhất 1 Theme Light chuẩn Microsoft Word)
 */
function initTheme() {
  document.documentElement.setAttribute('data-theme', 'light');
  try {
    localStorage.removeItem('pwa_mark_mix_theme');
  } catch (e) {}

  const metaThemeColor = document.querySelector('meta[name="theme-color"]');
  if (metaThemeColor) {
    metaThemeColor.setAttribute('content', '#185abd');
  }
}

/**
 * Cookie Helpers & Cache Thông Tin Đầu Đề Thi
 * - Cookie lưu 365 ngày (SameSite=Lax, URI-encoded)
 * - Tự động đồng bộ localStorage để hoạt động bền vững ngay cả khi offline PWA
 */
function setCookie(name, value, days = 365) {
  try {
    const expires = new Date(Date.now() + days * 864e5).toUTCString();
    document.cookie = `${encodeURIComponent(name)}=${encodeURIComponent(value)}; expires=${expires}; path=/; SameSite=Lax`;
    localStorage.setItem(`cookie_${name}`, value);
  } catch (e) {
    console.warn('Lỗi ghi cookie/localStorage:', e);
  }
}

function getCookie(name) {
  try {
    const nameEq = encodeURIComponent(name) + '=';
    const ca = document.cookie.split(';');
    for (let c of ca) {
      c = c.trim();
      if (c.indexOf(nameEq) === 0) {
        return decodeURIComponent(c.substring(nameEq.length));
      }
    }
    const fallbackVal = localStorage.getItem(`cookie_${name}`);
    if (fallbackVal !== null && fallbackVal !== undefined) {
      return fallbackVal;
    }
  } catch (e) {
    console.warn('Lỗi đọc cookie/localStorage:', e);
  }
  return '';
}

const HEADER_FIELDS_CONFIG = [
  { id: 'inpSoGD', cookieKey: 'pwa_exam_so_gd', label: 'Sở / Phòng GD&ĐT' },
  { id: 'inpTruong', cookieKey: 'pwa_exam_truong', label: 'Tên Trường / Đơn vị' },
  { id: 'inpKiKiemTra', cookieKey: 'pwa_exam_ki_kiem_tra', label: 'Kì kiểm tra / Đợt thi' },
  { id: 'inpNamHoc', cookieKey: 'pwa_exam_nam_hoc', label: 'Năm học' },
  { id: 'inpMon', cookieKey: 'pwa_exam_mon', label: 'Môn thi' },
  { id: 'inpThoiGian', cookieKey: 'pwa_exam_thoi_gian', label: 'Thời gian làm bài' }
];

/**
 * Khởi tạo tính năng Cache Cookie và Warning cho Thông tin đầu đề thi
 * - Nếu user đã nhập trước đó: Điền sẵn từ cookie/storage và ẩn warning.
 * - Nếu user chưa nhập trước đó: Để trống với hint placeholder và hiện warning.
 * - Lắng nghe sự kiện gõ phím để cache ngay lập tức và tự động cập nhật cảnh báo real-time.
 */
function initHeaderInfoCacheAndValidation() {
  const noticeBanner = document.getElementById('headerWarningNotice');

  // 1. Nạp dữ liệu đã lưu từ cookie/storage nếu có
  HEADER_FIELDS_CONFIG.forEach(field => {
    const input = document.getElementById(field.id);
    if (!input) return;

    const cachedVal = getCookie(field.cookieKey);
    if (cachedVal && cachedVal.trim().length > 0) {
      input.value = cachedVal;
    }
  });

  // 2. Cập nhật giao diện cảnh báo (warning badge trên từng ô & banner chung)
  function updateValidationUI() {
    let emptyCount = 0;

    HEADER_FIELDS_CONFIG.forEach(field => {
      const input = document.getElementById(field.id);
      const warnTag = document.getElementById(`warn_${field.id}`);
      if (!input) return;

      const val = input.value.trim();
      if (val.length === 0) {
        emptyCount++;
        input.classList.add('input-warning');
        if (warnTag) warnTag.classList.remove('hidden');
      } else {
        input.classList.remove('input-warning');
        if (warnTag) warnTag.classList.add('hidden');
      }
    });

    if (noticeBanner) {
      if (emptyCount > 0) {
        noticeBanner.classList.remove('hidden');
      } else {
        noticeBanner.classList.add('hidden');
      }
    }
  }

  // 3. Lắng nghe real-time khi user gõ để tự động lưu cookie và tắt warning
  HEADER_FIELDS_CONFIG.forEach(field => {
    const input = document.getElementById(field.id);
    if (!input) return;

    input.addEventListener('input', () => {
      const val = input.value.trim();
      setCookie(field.cookieKey, val, 365);
      updateValidationUI();
    });

    input.addEventListener('change', () => {
      const val = input.value.trim();
      setCookie(field.cookieKey, val, 365);
      updateValidationUI();
    });
  });

  // Gán hàm cập nhật vào window để các module khác có thể kích hoạt khi cần
  window.updateHeaderValidationUI = updateValidationUI;

  // Chạy kiểm tra ngay khi nạp trang
  updateValidationUI();
}

/**
 * Trạng thái mạng Online / Offline
 */
function initNetworkStatus() {
  const netStatus = document.getElementById('netStatus');
  if (!netStatus) return;

  function updateStatus() {
    if (navigator.onLine) {
      netStatus.className = 'net-status-pill online';
      netStatus.querySelector('.status-text').textContent = 'Online';
    } else {
      netStatus.className = 'net-status-pill offline';
      netStatus.querySelector('.status-text').textContent = 'Offline (PWA)';
      showSnackbar('Ứng dụng đang hoạt động ở chế độ Offline (PWA)!');
    }
  }

  window.addEventListener('online', updateStatus);
  window.addEventListener('offline', updateStatus);
  updateStatus();
}

/**
 * Đăng ký Service Worker và Cài đặt PWA
 */
function initPWA() {
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js').then((reg) => {
      // Lắng nghe cập nhật
      reg.addEventListener('updatefound', () => {
        const newWorker = reg.installing;
        if (newWorker) {
          newWorker.addEventListener('statechange', () => {
            if (newWorker.state === 'installed' && navigator.serviceWorker.controller) {
              showSnackbar('Đã có bản cập nhật mới! Bấm nút Cập nhật để áp dụng.');
            }
          });
        }
      });
    }).catch((err) => {
      console.warn('Service Worker registration skipped or failed:', err);
    });
  }

  // Bắt sự kiện cài đặt PWA
  const btnInstall = document.getElementById('btnInstall');
  window.addEventListener('beforeinstallprompt', (e) => {
    e.preventDefault();
    deferredInstallPrompt = e;
    if (btnInstall) {
      btnInstall.classList.remove('hidden');
    }
  });

  if (btnInstall) {
    btnInstall.addEventListener('click', async () => {
      if (!deferredInstallPrompt) return;
      deferredInstallPrompt.prompt();
      const choice = await deferredInstallPrompt.userChoice;
      if (choice.outcome === 'accepted') {
        showSnackbar('Cảm ơn bạn đã cài đặt ứng dụng!');
      }
      deferredInstallPrompt = null;
      btnInstall.classList.add('hidden');
    });
  }

  // Nút ép cập nhật phiên bản mới
  const btnUpdateApp = document.getElementById('btnUpdateApp');
  if (btnUpdateApp) {
    btnUpdateApp.addEventListener('click', async () => {
      if ('serviceWorker' in navigator) {
        const reg = await navigator.serviceWorker.getRegistration();
        if (reg) {
          await reg.update();
        }
      }
      window.location.reload();
    });
  }
}

/**
 * Nạp đề thi Word mẫu từ file mau_de_thi_co_marker.docx
 */
async function loadSampleExamDocx(silent = false) {
  try {
    const res = await fetch('./mau_de_thi_co_marker.docx');
    if (!res.ok) throw new Error(`Không thể tải tệp mẫu (${res.status})`);
    const blob = await res.blob();
    const file = new File([blob], 'mau_de_thi_co_marker.docx', {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });
    file.isSample = true;
    await handleFileLoaded(file);
    if (!silent) {
      showSnackbar('Đã nạp đề thi mẫu thử nghiệm thành công!');
    }
  } catch (err) {
    console.warn('Không thể nạp file mẫu ngầm:', err);
  }
}

/**
 * Gán toàn bộ sự kiện giao diện
 */
function initEventListeners() {
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');
  const btnLoadSample = document.getElementById('btnLoadSample');
  const btnChangeFile = document.getElementById('btnChangeFile');
  const btnToggleRawView = document.getElementById('btnToggleRawView');
  const rawViewContainer = document.getElementById('rawViewContainer');

  // Xử lý nạp file và kéo thả (Dropzone)
  if (dropzone && fileInput) {
    // Ngăn chặn sự kiện click từ input nổi bọt ngược lên label dropzone gây đệ quy
    fileInput.addEventListener('click', (e) => {
      e.stopPropagation();
    });

    // Hỗ trợ phím Enter hoặc Space khi focus bằng bàn phím vào dropzone
    dropzone.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        fileInput.click();
      }
    });

    ['dragenter', 'dragover'].forEach(name => {
      dropzone.addEventListener(name, (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach(name => {
      dropzone.addEventListener(name, (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
      });
    });

    dropzone.addEventListener('drop', (e) => {
      const files = e.dataTransfer.files;
      if (files && files.length > 0) {
        handleFileLoaded(files[0]);
      }
    });

    fileInput.addEventListener('change', () => {
      if (fileInput.files && fileInput.files.length > 0) {
        handleFileLoaded(fileInput.files[0]);
        // Reset giá trị để nếu người dùng chọn lại cùng một file vẫn nhận diện change
        fileInput.value = '';
      }
    });
  }

  // Nút Nạp Đề Mẫu
  if (btnLoadSample) {
    btnLoadSample.addEventListener('click', (e) => {
      e.preventDefault();
      loadSampleExamDocx(false);
    });
  }

  // Đổi file khác
  if (btnChangeFile && fileInput) {
    btnChangeFile.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      fileInput.click();
    });
  }

  // Bật/tắt xem văn bản thô
  if (btnToggleRawView && rawViewContainer) {
    btnToggleRawView.addEventListener('click', () => {
      const isHidden = rawViewContainer.classList.toggle('hidden');
      btnToggleRawView.textContent = isHidden ? 'Xem văn bản trích xuất' : 'Ẩn văn bản trích xuất';
    });
  }

  // Stepper số lượng đề
  const inpNumExams = document.getElementById('inpNumExams');
  const btnDecExams = document.getElementById('btnDecExams');
  const btnIncExams = document.getElementById('btnIncExams');
  const inpStartExamCode = document.getElementById('inpStartExamCode');

  function syncExamCodesFromInputs() {
    let count = parseInt(inpNumExams.value, 10);
    if (isNaN(count) || count < 1) count = 1;
    if (count > 24) count = 24;
    inpNumExams.value = count;

    const startCode = inpStartExamCode.value.trim() || '101';
    updateExamCodes(count, startCode);
  }

  if (btnDecExams) {
    btnDecExams.addEventListener('click', () => {
      const current = parseInt(inpNumExams.value, 10) || 1;
      if (current > 1) {
        inpNumExams.value = current - 1;
        syncExamCodesFromInputs();
      }
    });
  }

  if (btnIncExams) {
    btnIncExams.addEventListener('click', () => {
      const current = parseInt(inpNumExams.value, 10) || 1;
      if (current < 24) {
        inpNumExams.value = current + 1;
        syncExamCodesFromInputs();
      }
    });
  }

  if (inpNumExams) {
    inpNumExams.addEventListener('change', syncExamCodesFromInputs);
  }

  if (inpStartExamCode) {
    inpStartExamCode.addEventListener('input', () => {
      const startCode = inpStartExamCode.value.trim() || '101';
      const count = parseInt(inpNumExams.value, 10) || 4;
      updateExamCodes(count, startCode);
    });
  }

  // Cập nhật lại giao diện xem trước ngay lập tức khi đổi Định dạng số câu
  const radioNumberings = document.querySelectorAll('input[name="radioNumbering"]');
  radioNumberings.forEach(radio => {
    radio.addEventListener('change', () => {
      if (state.parsedExam && state.parsedExam.valid) {
        renderExamPreview(state.parsedExam);
      }
    });
  });

  // Cập nhật lại giao diện xem trước ngay lập tức khi đổi Kiểu nhãn đáp án
  const selDefaultLabelStyle = document.getElementById('selDefaultLabelStyle');
  if (selDefaultLabelStyle) {
    const updatePreviewOnLabelChange = () => {
      if (state.parsedExam && state.parsedExam.valid) {
        renderExamPreview(state.parsedExam);
      }
    };
    selDefaultLabelStyle.addEventListener('change', updatePreviewOnLabelChange);
    selDefaultLabelStyle.addEventListener('input', updatePreviewOnLabelChange);
  }

  // Cập nhật lại giao diện xem trước ngay lập tức khi đổi Số câu đầu tiên của đề
  const inpStartQNum = document.getElementById('inpStartQNum');
  if (inpStartQNum) {
    const updatePreviewOnStartNumChange = () => {
      if (state.parsedExam && state.parsedExam.valid) {
        renderExamPreview(state.parsedExam);
      }
    };
    inpStartQNum.addEventListener('input', updatePreviewOnStartNumChange);
    inpStartQNum.addEventListener('change', updatePreviewOnStartNumChange);
  }

  // Nút Trộn & Xuất Đề Thi
  const btnMixAndExport = document.getElementById('btnMixAndExport');
  if (btnMixAndExport) {
    btnMixAndExport.addEventListener('click', () => {
      executeMixAndExport();
    });
  }

  // Modal đóng
  const exportModal = document.getElementById('exportModal');
  const btnCloseExportModal = document.getElementById('btnCloseExportModal');
  const btnCloseModalBtn = document.getElementById('btnCloseModalBtn');

  const closeExportModal = () => {
    if (exportModal) {
      exportModal.classList.remove('active');
    }
  };

  [btnCloseExportModal, btnCloseModalBtn].forEach(btn => {
    if (btn) {
      btn.addEventListener('click', closeExportModal);
    }
  });

  if (exportModal) {
    exportModal.addEventListener('click', (e) => {
      if (e.target === exportModal) {
        closeExportModal();
      }
    });
  }

  // Tải ZIP
  const btnDownloadZip = document.getElementById('btnDownloadZip');
  if (btnDownloadZip) {
    btnDownloadZip.addEventListener('click', () => {
      downloadAllZip();
    });
  }

  // Mobile Web Share
  const btnShareFiles = document.getElementById('btnShareFiles');
  if (btnShareFiles) {
    btnShareFiles.addEventListener('click', async () => {
      if (state.generatedZipBlob && navigator.share) {
        try {
          const file = new File([state.generatedZipBlob], 'bo_de_thi_tron.zip', {
            type: 'application/zip'
          });
          await navigator.share({
            title: 'Bộ Đề Thi Trộn',
            text: 'Đã tạo xong các đề thi trắc nghiệm và bảng đáp án',
            files: [file]
          });
        } catch (e) {
          if (e.name !== 'AbortError') {
            showSnackbar('Không thể chia sẻ tệp');
          }
        }
      }
    });
  }

  // iOS Safari Modal
  const isIos = /iPad|iPhone|iPod/.test(navigator.userAgent) && !window.MSStream;
  const isStandalone = window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone;
  const iosInstallModal = document.getElementById('iosInstallModal');
  if (isIos && !isStandalone && iosInstallModal) {
    const btnCloseIosModal = document.getElementById('btnCloseIosModal');
    const btnGotItIos = document.getElementById('btnGotItIos');

    const closeIosModal = () => {
      iosInstallModal.classList.remove('active');
    };

    [btnCloseIosModal, btnGotItIos].forEach(b => {
      if (b) {
        b.addEventListener('click', closeIosModal);
      }
    });

    iosInstallModal.addEventListener('click', (e) => {
      if (e.target === iosInstallModal) {
        closeIosModal();
      }
    });
  }

  // Lắng nghe phím Escape để đóng các modal đang mở
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      if (exportModal && exportModal.classList.contains('active')) {
        closeExportModal();
      }
      if (iosInstallModal && iosInstallModal.classList.contains('active')) {
        iosInstallModal.classList.remove('active');
      }
    }
  });

  // Image Lightbox xem phóng to và kiểm tra thông số ảnh
  const imageLightboxModal = document.getElementById('imageLightboxModal');
  const btnCloseLightbox = document.getElementById('btnCloseLightbox');
  const lightboxImg = document.getElementById('lightboxImg');
  const lightboxMeta = document.getElementById('lightboxMeta');
  const previewContainer = document.getElementById('previewQuestionsContainer');

  if (previewContainer && imageLightboxModal) {
    previewContainer.addEventListener('click', (e) => {
      const imgTarget = e.target.closest('.exam-preview-img');
      if (imgTarget && lightboxImg) {
        const w = imgTarget.getAttribute('data-width') || imgTarget.naturalWidth;
        const h = imgTarget.getAttribute('data-height') || imgTarget.naturalHeight;
        const natW = imgTarget.naturalWidth;
        const natH = imgTarget.naturalHeight;
        const ratio = (w && h) ? (w / h).toFixed(2) : '1.00';

        lightboxImg.src = imgTarget.src;
        if (lightboxMeta) {
          lightboxMeta.textContent = `Frame hiển thị Word: ${w} × ${h} px • Tỉ lệ: ${ratio} • Độ phân giải ảnh: ${natW} × ${natH} px`;
        }
        imageLightboxModal.classList.add('active');
      }
    });

    if (btnCloseLightbox) {
      btnCloseLightbox.addEventListener('click', () => {
        imageLightboxModal.classList.remove('active');
      });
    }

    imageLightboxModal.addEventListener('click', (e) => {
      if (e.target === imageLightboxModal) {
        imageLightboxModal.classList.remove('active');
      }
    });

    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && imageLightboxModal.classList.contains('active')) {
        imageLightboxModal.classList.remove('active');
      }
    });
  }
}


})();
