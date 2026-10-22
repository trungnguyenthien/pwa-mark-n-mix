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

// --- FILE: js/marker-converter.js ---
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
function convertTextToMarkers(rawText) {
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
async function generateConvertedDocxBlob(convertedText, imageMap = {}) {
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

// --- FILE: js/convert.js ---
/**
 * Controller Giao Diện Cho Trang Chuyển Đổi Marker (convert.html)
 * - Quản lý nạp file Word truyền thống (drag & drop / file picker)
 * - Tự động trích xuất, phân tích và chuyển đổi sang chuẩn marker
 * - Hiển thị bản xem trước trực quan với huy hiệu phân loại màu sắc
 * - Cảnh báo các câu hỏi thiếu đáp án gạch chân hoặc ảnh không phải inline
 * - Xuất và tải về file Word (.docx) đã gắn marker chuẩn Times New Roman
 */




// Trạng thái cục bộ của trang chuyển đổi
const convertState = {
  file: null,
  fileName: '',
  rawText: '',
  convertedText: '',
  imageMap: {},
  nonInlineImages: [],
  stats: null,
  activeTab: 'visual' // 'visual' | 'code'
};

// Khởi chạy khi DOM sẵn sàng
document.addEventListener('DOMContentLoaded', () => {
  initTheme();
  initNetworkStatus();
  initEventListeners();
});

/**
 * Đồng bộ theme Word Light
 */
function initTheme() {
  document.documentElement.setAttribute('data-theme', 'light');
  const metaThemeColor = document.querySelector('meta[name="theme-color"]');
  if (metaThemeColor) {
    metaThemeColor.setAttribute('content', '#185abd');
  }
}

/**
 * Trạng thái kết nối mạng (Online/Offline)
 */
function initNetworkStatus() {
  const netStatus = document.getElementById('netStatus');
  if (!netStatus) return;

  function update() {
    const isOnline = navigator.onLine;
    netStatus.className = `net-status-pill ${isOnline ? 'online' : 'offline'}`;
    const dot = netStatus.querySelector('.status-dot');
    const txt = netStatus.querySelector('.status-text');
    if (dot) dot.className = `status-dot ${isOnline ? '' : 'offline'}`;
    if (txt) txt.textContent = isOnline ? 'Online' : 'Offline';
  }

  window.addEventListener('online', update);
  window.addEventListener('offline', update);
  update();
}

/**
 * Hiển thị thông báo Toast / Snackbar
 */
function showToast(message, type = 'info', duration = 3500) {
  const snackbar = document.getElementById('snackbar');
  if (!snackbar) return;

  snackbar.textContent = message;
  snackbar.className = `snackbar show ${type}`;

  setTimeout(() => {
    snackbar.className = snackbar.className.replace('show', '').trim();
  }, duration);
}

/**
 * Gán toàn bộ sự kiện trên trang Convert
 */
function initEventListeners() {
  const dropzone = document.getElementById('dropzoneConvert');
  const fileInput = document.getElementById('fileInputConvert');
  const btnSelectFile = document.getElementById('btnSelectFile');
  const btnLoadSampleTrad = document.getElementById('btnLoadSampleTrad');
  const btnDownloadConvertedDocx = document.getElementById('btnDownloadConvertedDocx');
  const btnOpenInMixer = document.getElementById('btnOpenInMixer');
  const btnCopyMarkerText = document.getElementById('btnCopyMarkerText');
  const btnTabVisual = document.getElementById('btnTabVisual');
  const btnTabCode = document.getElementById('btnTabCode');
  const btnCloseNonInlineModal = document.getElementById('btnCloseNonInlineModal');
  const btnDismissNonInline = document.getElementById('btnDismissNonInline');
  const modalNonInline = document.getElementById('nonInlineModal');

  // Xử lý nạp file qua file input
  if (fileInput) {
    fileInput.addEventListener('change', () => {
      if (fileInput.files && fileInput.files.length > 0) {
        processUploadedFile(fileInput.files[0]);
        fileInput.value = '';
      }
    });
  }

  if (btnSelectFile && fileInput) {
    btnSelectFile.addEventListener('click', (e) => {
      e.stopPropagation();
      fileInput.click();
    });
  }

  // Xử lý Drag & Drop
  if (dropzone) {
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
        processUploadedFile(files[0]);
      }
    });

    dropzone.addEventListener('click', () => {
      if (fileInput) fileInput.click();
    });

    dropzone.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        if (fileInput) fileInput.click();
      }
    });
  }

  // Nạp đề mẫu truyền thống để thử nghiệm nhanh
  if (btnLoadSampleTrad) {
    btnLoadSampleTrad.addEventListener('click', (e) => {
      e.stopPropagation();
      loadTraditionalSample();
    });
  }

  // Nút Tải file Word đã gắn marker
  if (btnDownloadConvertedDocx) {
    btnDownloadConvertedDocx.addEventListener('click', handleDownloadDocx);
  }

  // Nút Chuyển sang trang Trộn đề
  if (btnOpenInMixer) {
    btnOpenInMixer.addEventListener('click', handleOpenInMixer);
  }

  // Nút Sao chép mã Marker
  if (btnCopyMarkerText) {
    btnCopyMarkerText.addEventListener('click', () => {
      const rawTextarea = document.getElementById('rawMarkerTextarea');
      if (rawTextarea && rawTextarea.value) {
        navigator.clipboard.writeText(rawTextarea.value).then(() => {
          showToast('Đã sao chép toàn bộ mã marker vào bộ nhớ đệm!', 'success');
        }).catch(() => {
          rawTextarea.select();
          document.execCommand('copy');
          showToast('Đã chọn và sao chép mã marker!', 'success');
        });
      }
    });
  }

  // Chuyển Tab xem trước
  if (btnTabVisual && btnTabCode) {
    btnTabVisual.addEventListener('click', () => switchTab('visual'));
    btnTabCode.addEventListener('click', () => switchTab('code'));
  }

  // Đóng Modal cảnh báo ảnh không phải InLine
  if (btnCloseNonInlineModal && modalNonInline) {
    btnCloseNonInlineModal.addEventListener('click', () => {
      modalNonInline.classList.remove('active');
    });
  }
  if (btnDismissNonInline && modalNonInline) {
    btnDismissNonInline.addEventListener('click', () => {
      modalNonInline.classList.remove('active');
    });
  }
  if (modalNonInline) {
    modalNonInline.addEventListener('click', (e) => {
      if (e.target === modalNonInline) {
        modalNonInline.classList.remove('active');
      }
    });
    window.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && modalNonInline.classList.contains('active')) {
        modalNonInline.classList.remove('active');
      }
    });
  }
}

/**
 * Chuyển đổi tab xem trước
 */
function switchTab(tab) {
  convertState.activeTab = tab;
  const btnTabVisual = document.getElementById('btnTabVisual');
  const btnTabCode = document.getElementById('btnTabCode');
  const visualContainer = document.getElementById('visualPreviewContainer');
  const codeContainer = document.getElementById('codePreviewContainer');

  if (tab === 'visual') {
    if (btnTabVisual) btnTabVisual.classList.add('active');
    if (btnTabCode) btnTabCode.classList.remove('active');
    if (visualContainer) visualContainer.classList.remove('hidden');
    if (codeContainer) codeContainer.classList.add('hidden');
  } else {
    if (btnTabCode) btnTabCode.classList.add('active');
    if (btnTabVisual) btnTabVisual.classList.remove('active');
    if (codeContainer) codeContainer.classList.remove('hidden');
    if (visualContainer) visualContainer.classList.add('hidden');
  }
}

/**
 * Xử lý tệp DOCX người dùng tải lên
 */
async function processUploadedFile(file) {
  if (!file) return;

  const fileName = file.name || 'de_thi.docx';
  if (!fileName.toLowerCase().endsWith('.docx')) {
    showToast('Vui lòng chọn tệp Word (.docx) hợp lệ!', 'danger');
    return;
  }

  setLoadingState(true, `Đang phân tích tệp "${fileName}"...`);

  try {
    // 1. Trích xuất text và ảnh từ file Word
    const extracted = await extractTextFromDocx(file);

    convertState.file = file;
    convertState.fileName = fileName;
    convertState.rawText = extracted.rawText;
    convertState.imageMap = extracted.imageMap || {};
    convertState.nonInlineImages = extracted.nonInlineImages || [];

    // 2. Chuyển đổi sang hệ thống marker
    const conversion = convertTextToMarkers(extracted.rawText);
    convertState.convertedText = conversion.convertedText;
    convertState.stats = conversion.stats;

    // 3. Hiển thị cảnh báo ảnh không phải InLine nếu có
    if (convertState.nonInlineImages.length > 0) {
      showNonInlineWarningModal(convertState.nonInlineImages);
    }

    // 4. Cập nhật giao diện kết quả
    renderConversionResults();
    showToast(`Đã chuyển đổi thành công ${conversion.stats.totalQuestions} câu hỏi!`, 'success');

  } catch (err) {
    console.error('Lỗi chuyển đổi:', err);
    showToast(`Lỗi xử lý file Word: ${err.message}`, 'danger');
  } finally {
    setLoadingState(false);
  }
}

/**
 * Nạp đề mẫu truyền thống để trải nghiệm ngay
 */
function loadTraditionalSample() {
  const sampleTraditionalText = `SỞ GIÁO DỤC VÀ ĐÀO TẠO TP HỒ CHÍ MINH
TRƯỜNG THPT CHUYÊN LÊ HỒNG PHONG
ĐỀ KIỂM TRA ĐỊNH KỲ NĂM HỌC 2026 - 2027
MÔN: VẬT LÝ 10 - THỜI GIAN: 45 PHÚT

PHẦN I. CÂU TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN
Thí sinh trả lời từ câu 1 đến câu 4. Mỗi câu hỏi chỉ chọn một phương án đúng nhất.

Câu 1. Trong cuộc Cách mạng công nghiệp 4.0, sự phát triển vượt bậc của viễn thông số, internet cáp quang và trí tuệ nhân tạo (AI) đều bắt nguồn từ nghiên cứu nền tảng của Vật lí. Đối tượng nghiên cứu chủ yếu của Vật lí học là
A. sự biến đổi chất này thành chất khác và cơ chế phản ứng hóa học.
<u>B.</u> các dạng vận động của vật chất và năng lượng trong tự nhiên.
C. quy luật tiến hóa, di truyền và sinh trưởng của các loài sinh vật.
D. cấu tạo địa chất và sự dịch chuyển của các tầng mảng kiến tạo Trái Đất.

Câu 2: Một xe máy chạy trên một đường thẳng thẳng đều với tốc độ 36 km/h. Đổi tốc độ này ra đơn vị mét trên giây (m/s) là:
<u>A. 10 m/s</u>       B. 15 m/s       C. 20 m/s       D. 25 m/s

Câu 3) Đại lượng đặc trưng cho sự biến thiên nhanh hay chậm của vận tốc theo thời gian là:
a) Quãng đường đi được.
b) Độ dịch chuyển tổng cộng.
<u>c)</u> Gia tốc của chuyển động.
d) Tốc độ tức thời.

4) Một vật chuyển động thẳng biến đổi đều từ trạng thái nghỉ với gia tốc a = 2 m/s^2. Vận tốc của vật sau 5 giây là:
A) 5 m/s       B) 8 m/s
<u>C) 10 m/s</u>       D) 20 m/s
`;

  setLoadingState(true, 'Đang nạp dữ liệu đề mẫu truyền thống...');
  setTimeout(() => {
    convertState.file = null;
    convertState.fileName = 'de_thi_truyen_thong_mau.docx';
    convertState.rawText = sampleTraditionalText;
    convertState.imageMap = {};
    convertState.nonInlineImages = [];

    const conversion = convertTextToMarkers(sampleTraditionalText);
    convertState.convertedText = conversion.convertedText;
    convertState.stats = conversion.stats;

    renderConversionResults();
    setLoadingState(false);
    showToast('Đã nạp đề mẫu truyền thống và chuyển đổi xong!', 'success');
  }, 300);
}

/**
 * Hiển thị kết quả chuyển đổi lên giao diện
 */
function renderConversionResults() {
  const resultSection = document.getElementById('resultSection');
  const emptyState = document.getElementById('emptyState');
  if (resultSection) resultSection.classList.remove('hidden');
  if (emptyState) emptyState.classList.add('hidden');

  const stats = convertState.stats || { totalQuestions: 0, totalD: 0, totalS: 0, questionsWithoutKey: [], warnings: [] };

  // 1. Cập nhật các con số thống kê
  const statQuestions = document.getElementById('statQuestions');
  const statD = document.getElementById('statD');
  const statS = document.getElementById('statS');
  const statWarningsCount = document.getElementById('statWarningsCount');

  if (statQuestions) statQuestions.textContent = stats.totalQuestions;
  if (statD) statD.textContent = stats.totalD;
  if (statS) statS.textContent = stats.totalS;
  if (statWarningsCount) statWarningsCount.textContent = (stats.questionsWithoutKey.length) + (convertState.nonInlineImages.length);

  // 2. Hiển thị cảnh báo thiếu đáp án gạch chân nếu có
  const warningContainer = document.getElementById('warningContainer');
  if (warningContainer) {
    if (stats.warnings.length > 0 || convertState.nonInlineImages.length > 0) {
      warningContainer.classList.remove('hidden');
      let html = '<div class="alert-box alert-warning">';
      html += '<div class="alert-icon">⚠️</div><div class="alert-content">';
      html += '<h4 class="alert-title">Lưu ý kiểm tra đề thi:</h4><ul class="alert-list">';

      for (const w of stats.warnings) {
        html += `<li>${escapeHtml(w)}</li>`;
      }
      if (convertState.nonInlineImages.length > 0) {
        html += `<li>Phát hiện <strong>${convertState.nonInlineImages.length}</strong> hình ảnh không phải In Line with Text (ảnh nổi). Hãy chỉnh ảnh thành In Line để hiển thị chuẩn xác nhất.</li>`;
      }
      html += '</ul></div></div>';
      warningContainer.innerHTML = html;
    } else {
      warningContainer.classList.add('hidden');
      warningContainer.innerHTML = '';
    }
  }

  // 3. Render bản xem trước trực quan (Visual Preview)
  renderVisualPreview(convertState.convertedText, convertState.imageMap);

  // 4. Render mã marker vào textarea
  const rawMarkerTextarea = document.getElementById('rawMarkerTextarea');
  if (rawMarkerTextarea) {
    rawMarkerTextarea.value = convertState.convertedText;
  }

  // 5. Cập nhật thông tin file nguồn
  const sourceFileName = document.getElementById('sourceFileName');
  if (sourceFileName) {
    sourceFileName.textContent = convertState.fileName;
  }

  // Cuộn nhẹ tới khu vực kết quả
  if (resultSection) {
    resultSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

/**
 * Render bản xem trước trực quan với huy hiệu màu sắc nổi bật
 */
function renderVisualPreview(convertedText, imageMap = {}) {
  const container = document.getElementById('visualPreviewContainer');
  if (!container) return;

  const lines = convertedText.split(/\r?\n/);
  let html = '<div class="doc-sheet">';

  let qIndex = 0;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) {
      html += '<div class="doc-line-spacer"></div>';
      continue;
    }

    if (line.startsWith('__G__') || /^__G\w*__/.test(line)) {
      html += `<div class="doc-marker-pill group-pill" title="Phân nhóm đề thi"><span class="pill-tag">${escapeHtml(line)}</span> <span class="pill-desc">Bắt đầu nhóm/phần mới</span></div>`;
      continue;
    }

    if (line === '__QUESTION__' || /^__QUESTION__\s*/.test(line)) {
      qIndex++;
      const rest = line.replace(/^__QUESTION__\s*/, '').trim();
      const renderedRest = rest ? formatContentForHtml(rest, imageMap) : '';
      html += `<div class="doc-marker-pill question-pill" title="Câu hỏi"><span class="pill-tag">__QUESTION__</span> <span class="pill-desc">Câu ${qIndex}</span></div>`;
      if (renderedRest) {
        html += `<div class="doc-content-line doc-question-text" style="font-weight: 500; margin-bottom: 6px;">${renderedRest}</div>`;
      }
      continue;
    }

    if (/^__(?:D|D#|#D)__\s*/.test(line)) {
      const markerMatch = line.match(/^__(?:D|D#|#D)__\s*/)[0];
      const rest = line.substring(markerMatch.length);
      const renderedContent = formatContentForHtml(rest, imageMap);
      html += `<div class="doc-answer-item correct-ans"><span class="ans-marker-badge correct" title="Đáp án đúng (có gạch chân ở đề gốc)">__D__ ✓</span> <span class="ans-text">${renderedContent}</span></div>`;
      continue;
    }

    if (/^__(?:S|S#|#S)__\s*/.test(line)) {
      const markerMatch = line.match(/^__(?:S|S#|#S)__\s*/)[0];
      const rest = line.substring(markerMatch.length);
      const renderedContent = formatContentForHtml(rest, imageMap);
      html += `<div class="doc-answer-item incorrect-ans"><span class="ans-marker-badge incorrect" title="Đáp án sai">__S__</span> <span class="ans-text">${renderedContent}</span></div>`;
      continue;
    }

    // Dòng văn bản bình thường (nội dung câu hỏi, tiêu đề nhóm, ảnh...)
    const rendered = formatContentForHtml(line, imageMap);
    html += `<div class="doc-content-line">${rendered}</div>`;
  }

  html += '</div>';
  container.innerHTML = html;
}

/**
 * Định dạng các thẻ HTML <b>, <i>, <u> và thay token [IMG:id] thành ảnh HTML
 */
function formatContentForHtml(text, imageMap = {}) {
  if (!text) return '';

  // Thay thế [IMG:id]
  let res = text.replace(/\[IMG:([^\]]+)\]/gi, (match, id) => {
    const img = imageMap[id];
    if (img && img.dataUrl) {
      return `<img src="${img.dataUrl}" alt="Hình ảnh đề thi" class="preview-inline-img" style="max-width: 100%; height: auto; max-height: 250px; display: inline-block; vertical-align: middle; margin: 4px 6px; border-radius: 4px;" />`;
    }
    return `<span class="img-missing-badge">[Hình ảnh: ${escapeHtml(id)}]</span>`;
  });

  return res;
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * Hiển thị / Ẩn trạng thái đang tải (Loading overlay)
 */
function setLoadingState(isLoading, message = 'Đang xử lý...') {
  const overlay = document.getElementById('loadingOverlay');
  const msgEl = document.getElementById('loadingMessage');
  if (overlay) {
    if (isLoading) {
      if (msgEl) msgEl.textContent = message;
      overlay.classList.add('active');
    } else {
      overlay.classList.remove('active');
    }
  }
}

/**
 * Tải file Word đã gắn marker về máy tính (Yêu cầu cốt lõi của người dùng)
 */
async function handleDownloadDocx() {
  if (!convertState.convertedText) {
    showToast('Chưa có nội dung đề thi đã chuyển đổi!', 'warning');
    return;
  }

  setLoadingState(true, 'Đang tạo tệp Word (.docx) chuẩn Times New Roman...');

  try {
    const blob = await generateConvertedDocxBlob(convertState.convertedText, convertState.imageMap);

    // Đặt tên file xuất: [TênGốc]_da_gan_marker.docx
    let baseName = convertState.fileName.replace(/\.docx$/i, '');
    if (!baseName) baseName = 'de_thi';
    const downloadName = `${baseName}_da_gan_marker.docx`;

    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = downloadName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 2000);

    showToast(`Đã tải xuống thành công tệp "${downloadName}"!`, 'success');
  } catch (err) {
    console.error('Lỗi xuất file Word:', err);
    showToast(`Không thể tạo file Word: ${err.message}`, 'danger');
  } finally {
    setLoadingState(false);
  }
}

/**
 * Chuyển nội dung sang trang Trộn đề (mix.html)
 */
async function handleOpenInMixer() {
  if (!convertState.convertedText) {
    showToast('Chưa có nội dung đề thi đã chuyển đổi!', 'warning');
    return;
  }

  setLoadingState(true, 'Đang chuẩn bị dữ liệu chuyển sang trang Trộn đề...');

  try {
    const blob = await generateConvertedDocxBlob(convertState.convertedText, convertState.imageMap);
    const baseName = convertState.fileName.replace(/\.docx$/i, '') || 'de_thi';
    const finalName = `${baseName}_da_gan_marker.docx`;

    // Lưu mảng buffer vào IndexedDB hoặc sessionStorage để chuyển qua mix.html
    const arrayBuffer = await blob.arrayBuffer();

    // Dùng IndexedDB để lưu buffer không bị giới hạn 5MB của sessionStorage
    await saveBlobToIndexedDB('pwa_transfer_exam', arrayBuffer, finalName);

    window.location.href = './mix.html?from=convert';
  } catch (err) {
    console.error('Lỗi chuyển trang:', err);
    showToast('Chuyển trang thất bại. Bạn có thể tải file Word về và nạp vào trang Trộn đề.', 'warning');
    setLoadingState(false);
  }
}

/**
 * Lưu tệp buffer vào IndexedDB cho việc truyền giữa các trang
 */
function saveBlobToIndexedDB(key, buffer, fileName) {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open('PwaMarkMixDB', 1);
    request.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains('transfers')) {
        db.createObjectStore('transfers');
      }
    };
    request.onsuccess = (e) => {
      const db = e.target.result;
      const tx = db.transaction('transfers', 'readwrite');
      const store = tx.objectStore('transfers');
      store.put({ buffer, fileName, time: Date.now() }, key);
      tx.oncomplete = () => resolve(true);
      tx.onerror = () => reject(tx.error);
    };
    request.onerror = () => reject(request.error);
  });
}

/**
 * Hiển thị modal cảnh báo ảnh không phải InLine
 */
function showNonInlineWarningModal(images) {
  const modal = document.getElementById('nonInlineModal');
  const countEl = document.getElementById('nonInlineCount');
  const listEl = document.getElementById('nonInlineList');
  if (!modal || !listEl) return;

  if (countEl) countEl.textContent = `${images.length} hình ảnh`;

  let html = '';
  images.forEach((img, idx) => {
    html += `
      <div class="non-inline-item">
        <div class="non-inline-num">${idx + 1}</div>
        <div class="non-inline-thumb">
          ${img.dataUrl ? `<img src="${img.dataUrl}" alt="${escapeHtml(img.name)}" />` : '<div class="no-thumb">Ảnh Word</div>'}
        </div>
        <div class="non-inline-info">
          <div class="non-inline-name">${escapeHtml(img.name)}</div>
          <div class="non-inline-meta">Kiểu bọc chữ hiện tại: <strong>${escapeHtml(img.wrapType)}</strong></div>
          <div class="non-inline-snippet">Vị trí đoạn văn: <em>"${escapeHtml(img.snippet)}"</em></div>
        </div>
      </div>
    `;
  });

  listEl.innerHTML = html;
  modal.classList.add('active');
}


})();
