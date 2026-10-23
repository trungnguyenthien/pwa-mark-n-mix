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
export function cleanAnswerText(text) {
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
export function extractQuestionAnswerKey(q) {
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
export function buildAnswerMatrix(allShuffledExams, options = {}) {
  if (!allShuffledExams || allShuffledExams.length === 0) {
    return { headers: [], rows: [] };
  }

  const examCodes = allShuffledExams.map(e => e.examCode);
  const totalQuestions = allShuffledExams[0].flatQuestions.length;
  const rows = [];
  const hasMultipleGroups = (allShuffledExams[0].groups && allShuffledExams[0].groups.length > 1);
  const isGroupScope = (options.numberingScope === 'group');

  for (let i = 0; i < totalQuestions; i++) {
    const qSample = allShuffledExams[0].flatQuestions[i];
    const qNum = qSample.displayNumber || (i + 1);
    const questionLabel = (isGroupScope && hasMultipleGroups && qSample.groupNumber)
      ? `P${qSample.groupNumber} - Câu ${qNum}`
      : `Câu ${qNum}`;

    const row = {
      questionNumber: qNum,
      questionLabel,
      groupNumber: qSample.groupNumber,
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
export async function createAnswerKeyDocxBlob(allShuffledExams, headerInfo = {}, options = {}) {
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
  const matrix = buildAnswerMatrix(allShuffledExams, options);

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
                text: row.questionLabel || `Câu ${row.questionNumber}`,
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
