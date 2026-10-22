/**
 * Module Dựng File Word Đề Thi (Docx Builder)
 * Tuân thủ quy định tại mục 5.1 và 5.2 của tài liệu:
 * "Công cụ trộn đề trắc nghiệm Yêu cầu và giải pháp.docx"
 *
 * - Font: Times New Roman toàn văn bản.
 * - Đầu đề: Bảng 2 cột không viền căn giữa, phối màu đỏ/đen đậm chuẩn mực.
 * - Dòng "Họ và tên", "Số báo danh", "Mã đề" có đường gạch ngang ngăn cách.
 * - Đánh số câu liên tục theo lựa chọn của người dùng.
 * - Nhãn đáp án theo định dạng từng nhóm.
 * - Câu trả lời ngắn không hiển thị đáp án trong đề thi.
 */

// Lấy tham chiếu docx từ window (trên trình duyệt) hoặc global (khi test)
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
 * Tạo một tài liệu docx cho một mã đề thi cụ thể
 * @param {Object} exam Đề thi đã hoán vị (từ shuffler)
 * @param {Object} headerInfo Thông tin đầu đề (7 ô nhập)
 * @returns {Promise<Blob>}
 */
export async function createExamDocxBlob(exam, headerInfo = {}) {
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
    Footer
  } = docx;

  const {
    soGD = 'SỞ GD&ĐT TP HCM',
    truong = 'TRƯỜNG THPT NGUYỄN HỮU HUÂN',
    kiKiemTra = 'GIỮA HK 1',
    namHoc = '2026 - 2027',
    mon = 'VẬT LÝ',
    thoiGian = '45 phút'
  } = headerInfo;

  const examCode = exam.examCode;
  const FONT_FAMILY = 'Times New Roman';

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
          // Cột trái: Sở GD, Trường, Vạch phân cách, Số trang
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
                spacing: { after: 40 },
                children: [
                  new TextRun({
                    text: (soGD || '').toUpperCase(),
                    bold: true,
                    color: 'FF0000',
                    font: FONT_FAMILY,
                    size: 24 // 12pt
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 40 },
                children: [
                  new TextRun({
                    text: (truong || '').toUpperCase(),
                    bold: true,
                    color: 'FF0000',
                    font: FONT_FAMILY,
                    size: 24 // 12pt
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 40 },
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
                  PageNumber.TOTAL_PAGES,
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

          // Cột phải: Kì kiểm tra, Năm học, Môn, Thời gian, Lưu ý
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
                spacing: { after: 40 },
                children: [
                  new TextRun({
                    text: (kiKiemTra || '').toUpperCase(),
                    bold: true,
                    color: 'FF0000',
                    font: FONT_FAMILY,
                    size: 24 // 12pt
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 40 },
                children: [
                  new TextRun({
                    text: 'NĂM HỌC ',
                    bold: true,
                    font: FONT_FAMILY,
                    size: 24
                  }),
                  new TextRun({
                    text: namHoc || '',
                    bold: true,
                    color: 'FF0000',
                    font: FONT_FAMILY,
                    size: 24
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 40 },
                children: [
                  new TextRun({
                    text: 'MÔN: ',
                    bold: true,
                    font: FONT_FAMILY,
                    size: 24
                  }),
                  new TextRun({
                    text: (mon || '').toUpperCase(),
                    bold: true,
                    color: 'FF0000',
                    font: FONT_FAMILY,
                    size: 24
                  })
                ]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { after: 40 },
                children: [
                  new TextRun({
                    text: 'Thời gian làm bài: ',
                    italics: true,
                    font: FONT_FAMILY,
                    size: 22
                  }),
                  new TextRun({
                    text: thoiGian || '',
                    italics: true,
                    color: 'FF0000',
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

  // 2. Dòng thông tin thí sinh và Mã đề
  const studentInfoLine = new Paragraph({
    spacing: { before: 100, after: 200 },
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 8, color: '000000' }
    },
    children: [
      new TextRun({
        text: 'Họ và tên: ............................................................................    ',
        font: FONT_FAMILY,
        size: 22
      }),
      new TextRun({
        text: 'Số báo danh: .......        ',
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
        color: 'FF0000',
        font: FONT_FAMILY,
        size: 24
      })
    ]
  });

  // 3. Nội dung các câu hỏi và đáp án
  const bodyParagraphs = [];

  for (const question of exam.questions) {
    // Đoạn văn bản câu hỏi
    const questionPrefix = question.displayPrefix || `Câu ${question.displayNumber || 1}.`;
    
    // Tách các dòng của câu hỏi (nếu có xuống dòng)
    const qLines = question.text.split('\n');
    const firstLine = qLines[0] || '';
    const otherLines = qLines.slice(1);

    const qRuns = [
      new TextRun({
        text: `${questionPrefix} `,
        bold: true,
        font: FONT_FAMILY,
        size: 24 // 12pt
      }),
      new TextRun({
        text: firstLine,
        font: FONT_FAMILY,
        size: 24
      })
    ];

    bodyParagraphs.push(
      new Paragraph({
        spacing: { before: 160, after: 60, line: 276 }, // 1.15 line spacing
        children: qRuns
      })
    );

    // Các dòng bổ sung của câu hỏi nếu có
    for (const line of otherLines) {
      bodyParagraphs.push(
        new Paragraph({
          spacing: { after: 60, line: 276 },
          indent: { left: 360 },
          children: [
            new TextRun({
              text: line,
              font: FONT_FAMILY,
              size: 24
            })
          ]
        })
      );
    }

    // Nếu không phải câu trả lời ngắn, hiển thị các đáp án
    if (question.type !== 'short') {
      for (const answer of question.answers) {
        const ansLines = answer.text.split('\n');
        const firstAnsLine = ansLines[0] || '';
        const otherAnsLines = ansLines.slice(1);

        bodyParagraphs.push(
          new Paragraph({
            spacing: { after: 40, line: 260 },
            indent: { left: 420 },
            children: [
              new TextRun({
                text: `${answer.label} `,
                bold: true,
                font: FONT_FAMILY,
                size: 24
              }),
              new TextRun({
                text: firstAnsLine,
                font: FONT_FAMILY,
                size: 24
              })
            ]
          })
        );

        for (const line of otherAnsLines) {
          bodyParagraphs.push(
            new Paragraph({
              spacing: { after: 40, line: 260 },
              indent: { left: 680 },
              children: [
                new TextRun({
                  text: line,
                  font: FONT_FAMILY,
                  size: 24
                })
              ]
            })
          );
        }
      }
    }
  }

  // Kết thúc đề thi
  const endParagraph = new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 360, after: 120 },
    children: [
      new TextRun({
        text: '---------- HẾT ----------',
        bold: true,
        font: FONT_FAMILY,
        size: 22
      })
    ]
  });

  // Footer chân trang
  const footer = new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: `Trang `,
            italics: true,
            font: FONT_FAMILY,
            size: 18 // 9pt
          }),
          PageNumber.CURRENT,
          new TextRun({
            text: ` / `,
            italics: true,
            font: FONT_FAMILY,
            size: 18
          }),
          PageNumber.TOTAL_PAGES,
          new TextRun({
            text: ` - Mã đề: ${examCode}`,
            italics: true,
            font: FONT_FAMILY,
            size: 18
          })
        ]
      })
    ]
  });

  // Tạo Section và Document
  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 1134, // ~2cm
              bottom: 1134,
              left: 1417, // ~2.5cm
              right: 1134  // ~2cm
            }
          }
        },
        footers: {
          default: footer
        },
        children: [
          headerTable,
          studentInfoLine,
          ...bodyParagraphs,
          endParagraph
        ]
      }
    ]
  });

  return await docx.Packer.toBlob(doc);
}
