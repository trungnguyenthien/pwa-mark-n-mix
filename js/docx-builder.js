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
    Table,
    TableRow,
    TableCell,
    WidthType,
    BorderStyle,
    AlignmentType,
    PageNumber,
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
                    color: COLOR_RED,
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
                    color: COLOR_RED,
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

          // Cột phải: Kì kiểm tra (đỏ, đậm); Năm học; Môn; Thời gian làm bài; "(không kể thời gian phát đề)"
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
                    color: COLOR_RED,
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
                    color: COLOR_RED,
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
                    color: COLOR_RED,
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
                    color: COLOR_RED,
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
        size: 22
      }),
      new TextRun({
        text: 'Số báo danh: .......................  ',
        font: FONT_FAMILY,
        size: 22
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
        color: COLOR_RED,
        font: FONT_FAMILY,
        size: 26
      })
    ]
  });

  // 3. Nội dung phần thân đề (Danh sách câu hỏi & đáp án)
  const contentParagraphs = [];

  for (const q of shuffledExam.flatQuestions) {
    // Đoạn câu hỏi
    const questionRuns = [
      new TextRun({
        text: `${q.displayPrefix} `,
        bold: true,
        font: FONT_FAMILY,
        size: 24
      })
    ];

    // Xử lý nội dung câu hỏi (có thể có nhiều dòng)
    const qLines = (q.text || '').split('\n');
    qLines.forEach((line, lIdx) => {
      if (lIdx > 0) {
        questionRuns.push(new TextRun({ break: 1 }));
      }
      questionRuns.push(new TextRun({
        text: line,
        font: FONT_FAMILY,
        size: 24
      }));
    });

    contentParagraphs.push(new Paragraph({
      spacing: { before: 120, after: 60 },
      children: questionRuns
    }));

    // Đoạn đáp án (Câu trả lời ngắn không hiển thị đáp án trong đề thi theo mục 5.2)
    if (q.type !== 'short_answer') {
      for (const ans of q.answers) {
        const answerRuns = [
          new TextRun({
            text: `${ans.labelDisplay} `,
            bold: true,
            font: FONT_FAMILY,
            size: 24
          })
        ];

        const aLines = (ans.text || '').split('\n');
        aLines.forEach((aLine, aIdx) => {
          if (aIdx > 0) {
            answerRuns.push(new TextRun({ break: 1 }));
          }
          answerRuns.push(new TextRun({
            text: aLine,
            font: FONT_FAMILY,
            size: 24
          }));
        });

        contentParagraphs.push(new Paragraph({
          indent: { left: 400 }, // Thụt đầu dòng cho các phương án đáp án
          spacing: { before: 30, after: 30 },
          children: answerRuns
        }));
      }
    }
  }

  // Đoạn kết thúc đề
  contentParagraphs.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 240, after: 100 },
    children: [
      new TextRun({
        text: '----------------------- HẾT -----------------------',
        bold: true,
        font: FONT_FAMILY,
        size: 22
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
