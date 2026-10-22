/**
 * Module Bảng Đáp Án (Answer Key)
 * Tuân thủ quy định tại mục 5.4 của tài liệu:
 * "Công cụ trộn đề trắc nghiệm Yêu cầu và giải pháp.docx"
 *
 * - Tự suy ra từ dữ liệu mà không cần biết trước loại câu:
 *   + 1 __D__: ghi nhãn của nó (ví dụ: C).
 *   + Nhiều __D__: ghi tập hợp (ví dụ: A, C) hoặc theo từng ý (a-Đ, b-S, c-Đ, d-S).
 *   + Trả lời ngắn (chỉ có __D__): ghi nguyên giá trị đáp án (ví dụ: 4).
 * - Bảng đáp án: mỗi hàng là một câu (Câu 1, Câu 2...), mỗi cột là một mã đề.
 * - Xuất thành file .docx riêng biệt.
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
 * Trích xuất chuỗi đáp án cho một câu hỏi trong đề
 * @param {Object} question
 * @returns {string}
 */
export function getQuestionAnswerString(question) {
  if (!question || !question.answers || question.answers.length === 0) {
    return '';
  }

  // 1. Trả lời ngắn: chỉ có 1 __D__, không có __S__
  if (question.type === 'short' || (question.answers.length === 1 && question.answers[0].correct)) {
    return question.answers[0].text.trim();
  }

  const correctAnswers = question.answers.filter(a => a.correct);

  // 2. Một đáp án đúng: ghi nhãn của đáp án đó (ví dụ "B" hoặc "A")
  if (correctAnswers.length === 1) {
    return correctAnswers[0].rawLabel || correctAnswers[0].label.replace(/[^A-Za-z0-9]/g, '');
  }

  // 3. Nhiều đáp án đúng:
  // Nếu nhãn theo kiểu a) b) c) d) -> định dạng a-Đ, b-S...
  const hasParenLabel = question.answers.some(a => a.label && a.label.includes(')'));
  if (hasParenLabel) {
    return question.answers
      .map(a => `${a.rawLabel}-${a.correct ? 'Đ' : 'S'}`)
      .join(', ');
  }

  // Nếu nhãn ABCD -> ghi danh sách nhãn đúng: A, C
  return correctAnswers
    .map(a => a.rawLabel || a.label.replace(/[^A-Za-z0-9]/g, ''))
    .join(', ');
}

/**
 * Tạo ma trận đáp án cho giao diện xem trước và xuất file
 * @param {Array<Object>} shuffledExams Danh sách các đề thi đã hoán vị
 * @returns {Object} { headers: ['Câu', 'Mã 101', ...], rows: [ ['Câu 1', 'A', 'B', ...], ... ] }
 */
export function generateAnswerMatrix(shuffledExams) {
  if (!shuffledExams || shuffledExams.length === 0) {
    return { headers: ['Câu'], rows: [] };
  }

  const examCodes = shuffledExams.map(e => e.examCode);
  const headers = ['Câu', ...examCodes.map(code => `Mã đề ${code}`)];

  // Số lượng câu hỏi lấy theo đề đầu tiên
  const totalQuestions = shuffledExams[0].questions.length;
  const rows = [];

  for (let qIdx = 0; qIdx < totalQuestions; qIdx++) {
    const qNum = shuffledExams[0].questions[qIdx].displayNumber || (qIdx + 1);
    const row = [`Câu ${qNum}`];

    for (const exam of shuffledExams) {
      const q = exam.questions[qIdx];
      const ansStr = getQuestionAnswerString(q);
      row.push(ansStr);
    }
    rows.push(row);
  }

  return { headers, rows };
}

/**
 * Tạo file docx bảng đáp án cho toàn bộ các mã đề
 * @param {Array<Object>} shuffledExams Danh sách đề thi đã hoán vị
 * @param {Object} headerInfo Thông tin chung
 * @returns {Promise<Blob>}
 */
export async function createAnswerKeyDocxBlob(shuffledExams, headerInfo = {}) {
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
    mon = 'VẬT LÝ'
  } = headerInfo;

  const FONT_FAMILY = 'Times New Roman';
  const matrix = generateAnswerMatrix(shuffledExams);

  // Tính toán độ rộng cột trong bảng
  const totalCols = matrix.headers.length;
  const firstColWidth = 1400; // dxa cho cột "Câu"
  const totalTableWidth = 9600; // dxa
  const remainingColWidth = Math.floor((totalTableWidth - firstColWidth) / (totalCols - 1 || 1));

  // Tạo hàng tiêu đề bảng
  const headerRowCells = matrix.headers.map((hText, colIdx) => {
    return new TableCell({
      width: {
        size: colIdx === 0 ? firstColWidth : remainingColWidth,
        type: WidthType.DXA
      },
      shading: {
        fill: 'E2E8F0' // Xám nhạt thanh lịch
      },
      borders: {
        top: { style: BorderStyle.SINGLE, size: 4, color: '94A3B8' },
        bottom: { style: BorderStyle.SINGLE, size: 8, color: '475569' },
        left: { style: BorderStyle.SINGLE, size: 4, color: '94A3B8' },
        right: { style: BorderStyle.SINGLE, size: 4, color: '94A3B8' }
      },
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 80, after: 80 },
          children: [
            new TextRun({
              text: hText,
              bold: true,
              font: FONT_FAMILY,
              size: 22
            })
          ]
        })
      ]
    });
  });

  const tableRows = [
    new TableRow({
      tableHeader: true,
      children: headerRowCells
    })
  ];

  // Tạo các hàng dữ liệu câu hỏi
  matrix.rows.forEach((row, rowIdx) => {
    const isAlt = rowIdx % 2 === 1;
    const cells = row.map((cellText, colIdx) => {
      const isFirstCol = colIdx === 0;
      return new TableCell({
        width: {
          size: colIdx === 0 ? firstColWidth : remainingColWidth,
          type: WidthType.DXA
        },
        shading: {
          fill: isAlt ? 'F8FAFC' : 'FFFFFF'
        },
        borders: {
          top: { style: BorderStyle.SINGLE, size: 2, color: 'CBD5E1' },
          bottom: { style: BorderStyle.SINGLE, size: 2, color: 'CBD5E1' },
          left: { style: BorderStyle.SINGLE, size: 2, color: 'CBD5E1' },
          right: { style: BorderStyle.SINGLE, size: 2, color: 'CBD5E1' }
        },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 50, after: 50 },
            children: [
              new TextRun({
                text: cellText,
                bold: isFirstCol,
                color: isFirstCol ? '0F172A' : '0284C7', // Xanh nổi bật cho đáp án
                font: FONT_FAMILY,
                size: 22
              })
            ]
          })
        ]
      });
    });

    tableRows.push(
      new TableRow({
        children: cells
      })
    );
  });

  const answerTable = new Table({
    width: { size: totalTableWidth, type: WidthType.DXA },
    rows: tableRows
  });

  // Tiêu đề trang đáp án
  const titleParagraph = new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 60 },
    children: [
      new TextRun({
        text: 'BẢNG ĐÁP ÁN CÁC MÃ ĐỀ THI',
        bold: true,
        color: 'FF0000',
        font: FONT_FAMILY,
        size: 32 // 16pt
      })
    ]
  });

  const subtitleParagraph = new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
    children: [
      new TextRun({
        text: `${(truong || '').toUpperCase()} - MÔN: ${(mon || '').toUpperCase()} (${(kiKiemTra || '').toUpperCase()} - NĂM HỌC ${namHoc || ''})`,
        bold: true,
        font: FONT_FAMILY,
        size: 22
      })
    ]
  });

  const footer = new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: 'Bảng đáp án tổng hợp - Trang ',
            italics: true,
            font: FONT_FAMILY,
            size: 18
          }),
          PageNumber.CURRENT,
          new TextRun({
            text: ' / ',
            italics: true,
            font: FONT_FAMILY,
            size: 18
          }),
          PageNumber.TOTAL_PAGES
        ]
      })
    ]
  });

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 1134,
              bottom: 1134,
              left: 1134,
              right: 1134
            }
          }
        },
        footers: {
          default: footer
        },
        children: [
          titleParagraph,
          subtitleParagraph,
          answerTable
        ]
      }
    ]
  });

  return await docx.Packer.toBlob(doc);
}
