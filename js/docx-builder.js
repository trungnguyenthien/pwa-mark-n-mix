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
 * Tạo Blob file Word (.docx) cho một mã đề thi cụ thể
 * @param {Object} shuffledExam Đề thi đã hoán vị (từ shuffler)
 * @param {Object} headerInfo Thông tin đầu đề (7 ô nhập)
 * @returns {Promise<Blob>}
 */
export async function createExamDocxBlob(shuffledExam, headerInfo = {}) {
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

  const examCode = shuffledExam.examCode;
  const FONT_FAMILY = 'Times New Roman';
  const COLOR_RED = 'FF0000';
  const COLOR_BLACK = '000000';

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
    children: [
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
  function appendContentRuns(targetRuns, text) {
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
          bold: isBold,
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

      appendContentRuns(questionRuns, q.text);

      // Khoảng cách: Trước câu đầu tiên của nhóm mới (từ nhóm 2 trở đi) có khoảng trống lớn hơn rõ rệt (260 twips)
      // Giữa các câu hỏi thông thường có khoảng trống 160 twips trước câu
      const beforeSpacing = (!isFirstGroup && isFirstQInGroup) ? 260 : 160;

      contentParagraphs.push(new Paragraph({
        spacing: { before: beforeSpacing, after: 60 },
        children: questionRuns
      }));

      // Đoạn đáp án (Câu trả lời ngắn không hiển thị đáp án trong đề thi theo mục 5.2)
      if (q.type !== 'short_answer') {
        for (const ans of q.answers) {
          const answerRuns = [
            new TextRun({
              text: `${ans.labelDisplay} `,
              bold: true, // IN ĐẬM "A.", "B.", "A)", "B)"
              font: FONT_FAMILY,
              size: 24 // 12pt
            })
          ];

          appendContentRuns(answerRuns, ans.text);

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
        text: '----------------------- HẾT -----------------------',
        bold: true,
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
