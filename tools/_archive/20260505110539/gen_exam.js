const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  HeadingLevel, LevelFormat, PageOrientation, convertMillimetersToTwip
} = require('docx');
const fs = require('fs');

// ===== 页面尺寸：严格标准A4 =====
// 1 inch = 1440 twip(DXA), 1mm = 56.69 twip
// A4: 210mm x 297mm = 11906 x 16838 twip
const A4_W = 11906;
const A4_H = 16838;
// 页边距 20mm
const MARGIN_TOP    = 1134; // 20mm
const MARGIN_BOTTOM = 1134;
const MARGIN_LEFT   = 1134;
const MARGIN_RIGHT  = 1134;
const CONTENT_W = A4_W - MARGIN_LEFT - MARGIN_RIGHT; // ~9638 twip

// ===== 无边框/无底纹的单元格公共配置 =====
const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const NO_BORDERS = { top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER };
// 注意：不设置任何 shading，避免阴影
const THIN = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const CELL_BORDERS = { top: THIN, bottom: THIN, left: THIN, right: THIN };

// ===== 工具函数 =====
function t(text, opts = {}) {
  return new TextRun({ text, font: '宋体', size: 24, ...opts });
}
function bold(text, opts = {}) {
  return new TextRun({ text, font: '宋体', size: 24, bold: true, ...opts });
}
function sectionTitle(num, title) {
  return new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text: `${num}、${title}`, font: '黑体', size: 26, bold: true })],
  });
}
function para(children, opts = {}) {
  return new Paragraph({ spacing: { before: 80, after: 80 }, children, ...opts });
}
function emptyLine() {
  return new Paragraph({ children: [new TextRun('')], spacing: { before: 50, after: 50 } });
}

// 无底纹无边框单元格（用于布局表格）
function layoutCell(children, width) {
  return new TableCell({
    borders: NO_BORDERS,
    width: { size: width, type: WidthType.DXA },
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    // 不设置 shading，完全透明
    children,
  });
}

// 有边框无底纹单元格（用于填数表格）
function borderCell(text, width) {
  return new TableCell({
    borders: CELL_BORDERS,
    width: { size: width, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 60, right: 60 },
    // 明确设置白色背景，防止默认阴影
    shading: { fill: 'FFFFFF', type: ShadingType.CLEAR, color: 'auto' },
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: String(text), font: '宋体', size: 24 })],
    })],
  });
}

// 无边框布局表格
function layoutTable(rows, colWidths) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: colWidths,
    borders: {
      top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER,
      insideH: NO_BORDER, insideV: NO_BORDER,
    },
    rows,
  });
}

// =====================
// 一、计算题
// =====================
const calcData = [
  ['4+4-6=____', '3+6-9=____', '8-3+5=____', '8-4-2=____'],
  ['9-2-3=____', '7-5+2=____', '2+2+3=____', '1+5+3=____'],
  ['5+2-4=____', '10-6+3=____', '9-1-7=____', '3+3+3=____'],
];

function makeCalc() {
  const colW = Math.floor(CONTENT_W / 4);
  const rows = calcData.map(row =>
    new TableRow({
      children: row.map(item =>
        layoutCell([para([t(item)])], colW)
      ),
    })
  );
  return layoutTable(rows, Array(4).fill(colW));
}

// =====================
// 二、填空题
// =====================
const fillData = [
  ['（  ）- 7 = 3', '10 -（  ）= 4', '（  ）- 2 = 0', '（  ）+ 2 = 2'],
  ['5 +（  ）= 8', '6 -（  ）= 6', '（  ）+ 3 = 7', '（  ）- 3 = 3'],
  ['（  ）+ 4 = 9', '（  ）- 2 = 7', '7 -（  ）= 1', '4 -（  ）= 2'],
];

function makeFill() {
  const colW = Math.floor(CONTENT_W / 4);
  const rows = fillData.map(row =>
    new TableRow({
      children: row.map(item =>
        layoutCell([para([t(item)])], colW)
      ),
    })
  );
  return layoutTable(rows, Array(4).fill(colW));
}

// =====================
// 三、比大小
// =====================
const compareData = [
  ['5-2  ○  2+2', '2+4  ○  7-2', '8-6  ○  1+3', '9-6  ○  6-3'],
  ['10-5  ○  7-2', '3+3  ○  2+4', '2+3  ○  7-1', '1+4  ○  8-6'],
];

function makeCompare() {
  const colW = Math.floor(CONTENT_W / 4);
  const rows = compareData.map(row =>
    new TableRow({
      children: row.map(item =>
        layoutCell([para([t(item)])], colW)
      ),
    })
  );
  return layoutTable(rows, Array(4).fill(colW));
}

// =====================
// 四、连线题
// =====================
function makeLianxian() {
  // 左右各一组，用分隔线
  const lW = Math.floor(CONTENT_W * 0.23);
  const sepW = Math.floor(CONTENT_W * 0.04);

  const leftL = ['5+3', '10-3', '2+7', '9-6'];
  const leftR = ['8-5', '4+5', '9-2', '2+6'];
  const rightL = ['10-8', '9-5', '4+5', '3+3'];
  const rightR = ['3+6', '4-2', '10-4', '8-4'];

  const sepBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const midBorder = {
    top: sepBorder, bottom: sepBorder,
    left: THIN, right: sepBorder,
  };

  function lxCell(text, w, borders = NO_BORDERS) {
    return new TableCell({
      borders,
      width: { size: w, type: WidthType.DXA },
      margins: { top: 60, bottom: 60, left: 80, right: 80 },
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 80, after: 80 },
        children: [t(text)],
      })],
    });
  }

  const rows = leftL.map((_, i) =>
    new TableRow({
      children: [
        lxCell(leftL[i], lW),
        lxCell(leftR[i], lW),
        lxCell('', sepW, midBorder),
        lxCell(rightL[i], lW),
        lxCell(rightR[i], lW),
      ],
    })
  );

  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [lW, lW, sepW, lW, lW],
    borders: {
      top: NO_BORDER, bottom: NO_BORDER,
      left: NO_BORDER, right: NO_BORDER,
      insideH: NO_BORDER, insideV: NO_BORDER,
    },
    rows,
  });
}

// =====================
// 五、分与合
// =====================
function makeFenhege() {
  const colW = Math.floor(CONTENT_W / 6);

  // 每个分合图用3行文字表示: [顶, 斜线, 底]
  const row1Items = [
    ['□', '╱  ╲', '4    4'],
    ['10', '╱  ╲', '3    □'],
    ['2', '╱  ╲', '□    4'],
    ['□', '╱  ╲', '□    7'],
    ['3', '╱  ╲', '□    3'],
    ['6', '╱  ╲', '□    3'],
  ];
  const row2Items = [
    ['2    3', '╲  ╱', '□'],
    ['□    1', '╲  ╱', '4'],
    ['5    □', '╲  ╱', '9'],
    ['10', '╱  ╲', '□', '╲  ╱', '2'],  // 双层
    ['□', '╱  ╲', '4    6'],
  ];

  function fhCell(lines, w) {
    return new TableCell({
      borders: NO_BORDERS,
      width: { size: w, type: WidthType.DXA },
      margins: { top: 20, bottom: 20, left: 40, right: 40 },
      children: lines.map(line => new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 28, after: 28 },
        children: [new TextRun({ text: line, font: '宋体', size: 21 })],
      })),
    });
  }

  const tableRow1 = new TableRow({
    children: row1Items.map(lines => fhCell(lines, colW)),
  });

  // row2有5个，最后一个补空格
  const row2Cells = row2Items.map(lines => fhCell(lines, colW));
  // 补一个空白格
  row2Cells.push(fhCell([''], colW));

  const tableRow2 = new TableRow({ children: row2Cells });

  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: Array(6).fill(colW),
    borders: {
      top: NO_BORDER, bottom: NO_BORDER,
      left: NO_BORDER, right: NO_BORDER,
      insideH: NO_BORDER, insideV: NO_BORDER,
    },
    rows: [tableRow1, tableRow2],
  });
}

// =====================
// 六、按顺序填数
// =====================
function makeShunxu() {
  // 有边框数字表格，明确白色背景
  function numTable(data, colW) {
    return new Table({
      width: { size: colW * data.length, type: WidthType.DXA },
      columnWidths: Array(data.length).fill(colW),
      rows: [new TableRow({
        children: data.map(v => borderCell(v || '', colW)),
      })],
    });
  }

  return [
    sectionTitle('六', '按顺序填数'),
    para([t('① 1  2（  ）（  ）5（  ）7（  ）（  ）10')]),
    para([t('② 10（  ）8（  ）（  ）5（  ）（  ）2（  ）')]),
    emptyLine(),
    numTable(['', '7', '', '5', '4'], 480),
    emptyLine(),
    numTable(['2', '', '4', '5', '', '8', '', '10'], 380),
    emptyLine(),
    numTable(['10', '', '8', '', '6'], 480),
    emptyLine(),
    numTable(['10', '', '7', '6', '', '4', '', '2'], 380),
    emptyLine(),
  ];
}

// =====================
// 七、看图列算式
// =====================
function makeTusuan() {
  const halfW = Math.floor(CONTENT_W / 2);

  function tsCell(lines, w) {
    return new TableCell({
      borders: NO_BORDERS,
      width: { size: w, type: WidthType.DXA },
      margins: { top: 40, bottom: 40, left: 60, right: 60 },
      children: lines.map(line => new Paragraph({
        spacing: { before: 60, after: 60 },
        children: [t(line)],
      })),
    });
  }

  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [halfW, halfW],
    borders: {
      top: NO_BORDER, bottom: NO_BORDER,
      left: NO_BORDER, right: NO_BORDER,
      insideH: NO_BORDER, insideV: NO_BORDER,
    },
    rows: [
      new TableRow({ children: [
        tsCell(['○○○  ○○○○○', '□ ○ □ = □'], halfW),
        tsCell(['女女女  女女  女女女女', '□ ○ □ ○ □ = □'], halfW),
      ]}),
      new TableRow({ children: [
        tsCell(['△△△△△  （划去△△△）', '□ ○ □ = □'], halfW),
        tsCell(['PPP  PP  PPPP', '□ ○ □ ○ □ = □'], halfW),
      ]}),
    ],
  });
}

// =====================
// 八、单数双数
// =====================
function makeDanShuang() {
  return [
    sectionTitle('八', '请找出下列数字中的单数和双数'),
    para([t('1  2  3  4  5  6  7  8  9  10', { size: 26, bold: true })]),
    para([t('单数：（  ）（  ）（  ）（  ）（  ）')]),
    para([t('双数：（  ）（  ）（  ）（  ）（  ）')]),
  ];
}

// =====================
// 九、从大到小
// =====================
function makePaixuDA() {
  return [
    sectionTitle('九', '将下列数字从大到小排列'),
    para([t('① 2、5、7、4、3、9')]),
    para([t('□ > □ > □ > □ > □ > □')]),
    para([t('② 8、4、2、6、10、1')]),
    para([t('□ > □ > □ > □ > □ > □')]),
  ];
}

// =====================
// 十、从小到大
// =====================
function makePaixuXiao() {
  return [
    sectionTitle('十', '将下列数从小到大排列'),
    para([t('3、7、10、5、2、6')]),
    para([t('□ < □ < □ < □ < □ < □')]),
  ];
}

// =====================
// 组装文档
// =====================
const doc = new Document({
  // 关闭自动缩放相关设置
  settings: {
    // 兼容性设置，防止Word自动调整
    compatibilityModeVersion: 15,
  },
  styles: {
    default: {
      document: {
        run: { font: '宋体', size: 24, color: '000000' },
        paragraph: { spacing: { line: 320, lineRule: 'auto' } },
      },
    },
  },
  sections: [{
    properties: {
      page: {
        // 标准A4，精确尺寸
        size: {
          width: A4_W,   // 210mm = 11906 twip
          height: A4_H,  // 297mm = 16838 twip
        },
        margin: {
          top: MARGIN_TOP,
          bottom: MARGIN_BOTTOM,
          left: MARGIN_LEFT,
          right: MARGIN_RIGHT,
          header: 709,
          footer: 709,
          gutter: 0,
        },
      },
    },
    children: [
      // ===== 标题行 =====
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 200 },
        border: {
          bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000', space: 8 },
        },
        children: [
          new TextRun({ text: '小学数学练习试卷（1～10的认识）', font: '黑体', size: 30, bold: true }),
          new TextRun({ text: '          姓名：________', font: '宋体', size: 24 }),
        ],
      }),

      emptyLine(),

      // ===== 一、计算 =====
      sectionTitle('一', '计算'),
      makeCalc(),
      emptyLine(),

      // ===== 二、填空 =====
      sectionTitle('二', '填空'),
      makeFill(),
      emptyLine(),

      // ===== 三、比大小 =====
      sectionTitle('三', '比大小（在○里填上 > < =）'),
      makeCompare(),
      emptyLine(),

      // ===== 四、连线 =====
      sectionTitle('四', '连线（将相等的算式用线连起来）'),
      makeLianxian(),
      emptyLine(),

      // ===== 五、分与合 =====
      sectionTitle('五', '分与合填数'),
      makeFenhege(),
      emptyLine(),

      // ===== 六、按顺序填数 =====
      ...makeShunxu(),

      // ===== 七、看图列算式 =====
      sectionTitle('七', '看图列算式'),
      makeTusuan(),
      emptyLine(),

      // ===== 八、单数双数 =====
      ...makeDanShuang(),
      emptyLine(),

      // ===== 九、从大到小 =====
      ...makePaixuDA(),
      emptyLine(),

      // ===== 十、从小到大 =====
      ...makePaixuXiao(),
      emptyLine(),

      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 300 },
        children: [new TextRun({ text: '—— 试卷完 ——', font: '宋体', size: 20, color: '888888' })],
      }),
    ],
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('C:/Users/12844/WorkBuddy/20260505110539/小学数学试卷.docx', buffer);
  console.log('Done!');
});
