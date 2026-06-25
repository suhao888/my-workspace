const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  PageOrientation, Header, Footer, PageNumber
} = require('docx');
const fs = require('fs');

// ========= 样式辅助 =========
const border = { style: BorderStyle.SINGLE, size: 1, color: '999999' };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorders = {
  top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
};
const cellMargins = { top: 100, bottom: 100, left: 150, right: 150 };

function headerCell(text, w, bgColor = 'D6E4FF') {
  return new TableCell({
    borders, width: { size: w, type: WidthType.DXA },
    shading: { fill: bgColor, type: ShadingType.CLEAR },
    margins: cellMargins,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold: true, size: 18, font: 'Arial' })]
    })]
  });
}

function dataCell(text, w, align = AlignmentType.CENTER, bold = false, color = null) {
  const runProps = { text: String(text ?? ''), size: 18, font: 'Arial', bold };
  if (color) runProps.color = color;
  return new TableCell({
    borders, width: { size: w, type: WidthType.DXA },
    margins: cellMargins,
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun(runProps)]
    })]
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 400, after: 200 },
    children: [new TextRun({ text, bold: true, size: 32, font: 'Arial', color: '1F3864' })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 300, after: 150 },
    children: [new TextRun({ text, bold: true, size: 26, font: 'Arial', color: '2E75B6' })]
  });
}

function h3(text) {
  return new Paragraph({
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text, bold: true, size: 22, font: 'Arial', color: 'C00000' })]
  });
}

function body(text, indent = 0) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    indent: indent ? { left: indent } : undefined,
    children: [new TextRun({ text, size: 20, font: 'Arial' })]
  });
}

function bullet(text) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    indent: { left: 360, hanging: 180 },
    children: [new TextRun({ text: '• ' + text, size: 20, font: 'Arial' })]
  });
}

function redBullet(text) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    indent: { left: 360, hanging: 180 },
    children: [new TextRun({ text: '⚠ ' + text, size: 20, font: 'Arial', color: 'C00000' })]
  });
}

function spacer() {
  return new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun('')] });
}

// ========= 文档内容 =========
const children = [];

// 封面标题
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 1200, after: 400 },
  children: [new TextRun({ text: '车辆服务中心司机差旅补助', bold: true, size: 56, font: 'Arial', color: '1F3864' })]
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 0, after: 400 },
  children: [new TextRun({ text: '审计风险分析报告', bold: true, size: 52, font: 'Arial', color: '2E75B6' })]
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 0, after: 200 },
  children: [new TextRun({ text: '（2024年12月 — 2026年3月）', size: 28, font: 'Arial', color: '595959' })]
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 0, after: 800 },
  children: [new TextRun({ text: '制表日期：2026年5月', size: 22, font: 'Arial', color: '595959' })]
}));
children.push(new Paragraph({
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: '2E75B6', space: 1 } },
  spacing: { before: 0, after: 600 },
  children: [new TextRun('')]
}));

// ============================================================
// 一、总体概况
// ============================================================
children.push(h1('一、总体概况'));
children.push(body('本次审计对车辆服务中心2024年12月至2026年3月共16个月的司机差旅费补助数据进行分析，数据来源于各月《司机差费计算表》，涵盖出车公里数补助、节日出车补助、周末出车补助、任内/任外出车补助等四类。'));
children.push(spacer());

// 总体数据汇总表
children.push(h3('数据汇总概览（2024年12月—2026年3月）'));

const overviewData = [
  ['统计维度', '数 值'],
  ['审计月份跨度', '16个月（2024.12—2026.03）'],
  ['在册司机人数', '约90~93人/月'],
  ['16个月累计总补助', '约 1,736,176 元'],
  ['月均补助总额', '约 108,511 元'],
  ['最高单月总补助', '约 147,296 元（2025年9月）'],
  ['最低单月总补助', '约 67,956 元（2025年2月，春节月）'],
  ['识别高风险记录数', '42条（出车天数超当月日历天）'],
  ['识别异常高补助记录', '73条（单人单月>3,000元）'],
];
const overviewTable = new Table({
  width: { size: 8640, type: WidthType.DXA },
  columnWidths: [4320, 4320],
  rows: overviewData.map((row, i) => new TableRow({
    children: i === 0
      ? [headerCell(row[0], 4320), headerCell(row[1], 4320)]
      : [dataCell(row[0], 4320, AlignmentType.LEFT, true), dataCell(row[1], 4320, AlignmentType.LEFT)]
  }))
});
children.push(overviewTable);
children.push(spacer());

// 各月补助金额趋势表
children.push(h3('各月补助金额汇总（元）'));
const monthlyData = [
  ['月份', '在册人数', '总实际公里(km)', '总计提公里(km)', '总补助(元)', '周末出车天次', '节日出车天次'],
  ['2024-12', '93', '300,764', '408,333', '94,079', '0', '631'],
  ['2025-01', '92', '212,296', '315,303', '87,524', '119', '11'],
  ['2025-02', '87', '129,226', '241,003', '67,956', '131', '15'],
  ['2025-03', '90', '311,278', '413,306', '120,073', '159', '0'],
  ['2025-04', '90', '332,285', '426,103', '142,593', '149', '12'],
  ['2025-05', '90', '326,392', '414,764', '135,278', '231', '15'],
  ['2025-06', '91', '352,404', '446,169', '136,268', '173', '16'],
  ['2025-07', '91', '333,068', '449,412', '141,528', '201', '0'],
  ['2025-08', '92', '350,367', '440,290', '136,294', '226', '0'],
  ['2025-09', '90', '392,499', '474,707', '147,296', '230', '0'],
  ['2025-10', '89', '265,607', '347,949', '111,285', '185', '39'],
  ['2025-11', '89', '364,043', '445,179', '135,921', '197', '0'],
  ['2025-12', '87', '328,811', '414,075', '121,585', '162', '0'],
  ['2026-01', '87', '205,815', '306,318', '84,658', '107', '11'],
  ['2026-02', '87', '162,958', '257,831', '70,779', '110', '14'],
  ['2026-03', '85', '282,223', '365,618', '103,260', '129', '0'],
  ['合  计', '—', '4,151,036', '5,765,359', '1,736,176', '2,309', '754'],
];
const colW2 = [900, 800, 1200, 1200, 1100, 1120, 1120];
const monthlyTable = new Table({
  width: { size: 9440, type: WidthType.DXA },
  columnWidths: colW2,
  rows: monthlyData.map((row, i) => new TableRow({
    children: row.map((cell, j) => {
      if (i === 0) return headerCell(cell, colW2[j]);
      const isTot = i === monthlyData.length - 1;
      return dataCell(cell, colW2[j], j === 0 ? AlignmentType.CENTER : AlignmentType.RIGHT, isTot);
    })
  }))
});
children.push(monthlyTable);
children.push(spacer());

// 说明：计提公里 vs 实际公里
children.push(body('注：计提公里＞实际公里是正常现象——月包车/日包车存在保底公里数制度。据附录C，月包车月计提下限3,000公里，日包车日计提下限144公里（约每月21工作日保底约3,024公里），体现了"保底"激励机制。审计关注点在于：保底标准的合理性，以及实际公里数与计提公里差距过大时的真实性。'));
children.push(spacer());

// ============================================================
// 二、制度文件变化分析
// ============================================================
children.push(h1('二、制度文件变化分析（风险1）'));
children.push(body('【关注点】近几年差旅费报销依据的制度文件变化过程，上的什么会、谁通过的。'));
children.push(spacer());

children.push(h2('（一）现行制度体系'));
children.push(bullet('上位制度：ZY/HBYT 040110-2024《差旅费管理办法》（中国石油华北油田公司级，以下简称"差旅费管理办法"）'));
children.push(bullet('下位制度：车辆服务中心《差旅费管理细则》（本单位实施层面细则，须经营财务室归口管理）'));
children.push(bullet('司机补助单独核算依据：《中国石油华北油田公司关于发布公司机关及直属单位公务用车与小车队内部结算价格的通知》（华北概预〔2013〕46号，以下简称"46号文"）'));
children.push(spacer());

children.push(h2('（二）制度关键内容梳理'));
const regData = [
  ['文件名称', '文号/版本', '关键内容', '审批层级', '风险提示'],
  ['差旅费管理办法（上位）', 'ZY/HBYT 040110-2024', '确立各类差旅标准框架，规定伙食补助≤100元/天（西藏等≤120元），差旅费实施细则须结合本单位实际制定', '华北油田公司级', '下位细则若超越上位标准，存在违规风险'],
  ['差旅费管理细则（现行本）', '本单位制定，版本不明', '司机补助按附录C阶梯计算；月包车保底3,000公里；日包车保底144公里/天；周末出车65元/天；节日出车300元/天（≤3天）', '本单位级，经营财务室负责', '细则版本及审批会议记录未见随文附注，需核查'],
  ['公务用车结算价格通知', '华北概预〔2013〕46号', '确立月包车3,000公里定额及日包车144公里定额，是计提公里保底数的直接来源', '华北油田公司概预（价格）部门', '文件已逾12年，市场油价/成本变化大，合理性存疑'],
];
const colW3 = [1800, 1400, 3200, 1200, 1840];
const regTable = new Table({
  width: { size: 9440, type: WidthType.DXA },
  columnWidths: colW3,
  rows: regData.map((row, i) => new TableRow({
    children: row.map((cell, j) => {
      if (i === 0) return headerCell(cell, colW3[j]);
      return dataCell(cell, colW3[j], AlignmentType.LEFT, false, j === 4 ? 'C00000' : null);
    })
  }))
});
children.push(regTable);
children.push(spacer());

children.push(h2('（三）审计关注要点'));
children.push(redBullet('制度文件历史版本缺失：现有细则未在文档中标注"依据×年×月×日职代会/党委会通过"等表述，审计无法判断历史修订节点。建议取阅历年职代会决议及经营分析会记录，确认细则每次修订时间及审批级别。'));
children.push(redBullet('46号文"历史化"风险：2013年文件沿用至今已超12年，未见后续调整文件，但补助标准（阶梯单价0.2—0.35元/公里）是否符合当前成本水平缺乏重新测算依据。建议要求被审单位提供历史版本文件目录。'));
children.push(redBullet('下位细则是否超标：细则规定"节日出车补助300元/天（出车3天以内）"，需对照上位差旅费管理办法相应条款，确认是否存在超标准发放。'));
children.push(spacer());

// ============================================================
// 三、保底公里数合理性分析
// ============================================================
children.push(h1('三、保底公里数合理性质疑（风险2）'));
children.push(body('【关注点】现有制度规定日包车每天保底144公里、月包车每月保底3,000公里（相当于每天约100公里），是否存在保底标准过高问题。'));
children.push(spacer());

children.push(h2('（一）保底标准来源与计算逻辑'));
children.push(bullet('月包车：月行驶公里≤3,000公里时，按3,000公里计提，超过则按实际公里计提。依据：华北概预〔2013〕46号。'));
children.push(bullet('日包车：日行驶公里≤144公里时，按144公里计提（144=3,000÷（月包价÷日包价）的估算值）。'));
children.push(bullet('保底折日均：月3,000公里÷20个工作日=150公里/工作日；或÷30日历天=100公里/日历天。'));
children.push(spacer());

children.push(h2('（二）实际数据对比分析'));
children.push(body('根据16个月数据统计：'));

const minKmData = [
  ['分析指标', '数 据', '说 明'],
  ['月均实际行驶总公里数', '约 259,440 公里', '16个月合计4,151,036公里÷16月'],
  ['月均计提总公里数', '约 360,335 公里', '保底机制导致计提远高于实际'],
  ['计提/实际 倍数', '约 1.39 倍', '平均超计提39%，大量补助基于"虚拟公里"'],
  ['实际行驶显著偏低的案例', '王小虎，16个月平均月行驶仅722公里', '但每月均计提3,000公里获600元保底补助'],
  ['王小虎16个月累计补助', '9,615元（补助均为600元或650元/月）', '实际年均行驶不足1,500公里，仍按3,000公里领取'],
  ['16月合计保底"虚增"公里', '约 1,614,323 公里（计提-实际）', '平均每月约100,895公里系按保底而非实际里程计提'],
];
const colW4 = [2400, 3840, 3200];
const minKmTable = new Table({
  width: { size: 9440, type: WidthType.DXA },
  columnWidths: colW4,
  rows: minKmData.map((row, i) => new TableRow({
    children: row.map((cell, j) => {
      if (i === 0) return headerCell(cell, colW4[j]);
      return dataCell(cell, colW4[j], AlignmentType.LEFT, j === 0);
    })
  }))
});
children.push(minKmTable);
children.push(spacer());

children.push(h2('（三）风险结论'));
children.push(redBullet('保底公里实质上是对"出勤不出车"司机的隐性补贴。月包车/日包车司机只要出勤、挂靠某辆车，即使实际行驶极低也能领取保底补助，缺乏与真实劳动量的对应关系。'));
children.push(redBullet('保底3,000公里/月折算至日历天约100公里，远高于普通城市用车日均30—50公里的合理区间，保底设定明显偏高。'));
children.push(redBullet('典型案例——王小虎：出车天数连续16个月均超过当月日历天数（最高47天/30天月），疑似出勤统计异常；实际行驶仅600—1,300公里/月，却长期享受3,000公里保底，涉及虚领补助风险。'));
children.push(spacer());

// ============================================================
// 四、出车天数及连续出勤异常
// ============================================================
children.push(h1('四、出车天数与连续出勤异常（风险3）'));
children.push(body('【关注点】个别司机出勤天数=31天，连续多天每天出车数百公里。'));
children.push(spacer());

children.push(h2('（一）出车天数超出当月日历天数（42条异常）'));
children.push(body('按常理，一个月最多30/31天（2月28/29天），单名司机出车天数不应超过月历天数。以下为主要异常案例（出车天数超过当月日历天数）：'));
children.push(spacer());

const overData = [
  ['月份', '司机', '出车天数', '当月日历天', '超出', '实际公里', '补助(元)', '风险说明'],
  ['2025-09', '王小虎', '47', '30', '超出17天', '1,329', '600', '极度异常：出车天数几乎为月历天2倍，路单真实性存疑'],
  ['2025-04', '王小虎', '43', '30', '超出13天', '811', '600', '出车多但里程极低，路单或存在虚报'],
  ['2025-01', '王小虎', '42', '31', '超出11天', '658', '600', ''],
  ['2025-06', '王小虎', '42', '30', '超出12天', '811', '600', ''],
  ['2025-12', '王小虎', '42', '31', '超出11天', '578', '600', ''],
  ['2025-03', '王小虎', '40', '31', '超出9天', '541', '600', ''],
  ['2025-09', '王锡堂', '35', '30', '超出5天', '3,094', '1,824', ''],
  ['2025-06', '王志军', '34', '30', '超出4天', '5,800', '2,265', ''],
  ['2025-04', '于卫东', '34', '30', '超出4天', '7,603', '2,584', ''],
  ['2025-09', '王志军', '34', '30', '超出4天', '9,126', '3,153', '单月9,126公里+34天出车，真实性需核实'],
  ['2025-11', '童庆龙', '31', '30', '超出1天', '6,872', '3,350', ''],
  ['2025-12', '郑旺', '31', '31', '等于', '8,654', '3,124', '全勤31天+8,654公里，高强度'],
  ['2025-11', '徐建军', '31', '30', '超出1天', '9,769', '4,079', '31天9,769公里，日均315公里'],
  ['2026-02', '王红平', '31', '28', '超出3天', '5,593', '2,128', '2月仅28天却记录31出车天'],
  ['2025-02', '张毅', '31', '28', '超出3天', '262', '1,591', '2月28天却出车31天，数据异常'],
];
const colW5 = [780, 780, 850, 900, 750, 880, 880, 2620];
const overTable = new Table({
  width: { size: 9440, type: WidthType.DXA },
  columnWidths: colW5,
  rows: overData.map((row, i) => new TableRow({
    children: row.map((cell, j) => {
      if (i === 0) return headerCell(cell, colW5[j]);
      const isRisk = j === 7 && cell.length > 1;
      return dataCell(cell, colW5[j], j <= 1 || j === 7 ? AlignmentType.LEFT : AlignmentType.CENTER, j === 0, isRisk ? 'C00000' : null);
    })
  }))
});
children.push(overTable);
children.push(spacer());

children.push(h2('（二）王小虎专项分析'));
children.push(body('王小虎在16个月内出车天数均超出当月日历天数（2024年12月最低也超出11天），且月行驶公里极低（最低396公里、最高1,329公里），形成以下异常组合：'));
children.push(spacer());

const wxhData = [
  ['月份', '出车天数', '当月天数', '超出天数', '实际公里', '计提公里', '月补助(元)'],
  ['2024-12', '42', '31', '+11', '629', '3,000', '600'],
  ['2025-01', '42', '31', '+11', '658', '3,000', '600'],
  ['2025-02', '36', '28', '+8', '455', '3,000', '600'],
  ['2025-03', '40', '31', '+9', '541', '3,000', '600'],
  ['2025-04', '43', '30', '+13', '811', '3,000', '600'],
  ['2025-05', '36', '31', '+5', '724', '3,000', '600'],
  ['2025-06', '42', '30', '+12', '811', '3,000', '600'],
  ['2025-07', '39', '31', '+8', '957', '3,000', '600'],
  ['2025-08', '30', '31', '-1', '1,068', '3,000', '650'],
  ['2025-09', '47', '30', '+17', '1,329', '3,000', '600'],
  ['2025-10', '36', '31', '+5', '396', '3,000', '600'],
  ['2025-11', '37', '30', '+7', '1,016', '3,000', '665'],
  ['2025-12', '42', '31', '+11', '578', '3,000', '600'],
  ['2026-01', '38', '31', '+7', '714', '3,000', '600'],
  ['2026-02', '32', '28', '+4', '651', '3,000', '600'],
  ['2026-03', '26', '31', '-5', '667', '3,000', '600'],
  ['合计', '—', '—', '—', '11,025', '48,000', '9,615'],
];
const colW6 = [780, 900, 900, 860, 1050, 1050, 1050];
const wxhTable = new Table({
  width: { size: 6590, type: WidthType.DXA },
  columnWidths: colW6,
  rows: wxhData.map((row, i) => new TableRow({
    children: row.map((cell, j) => {
      if (i === 0) return headerCell(cell, colW6[j]);
      const isTot = i === wxhData.length - 1;
      return dataCell(cell, colW6[j], AlignmentType.CENTER, isTot);
    })
  }))
});
children.push(wxhTable);
children.push(spacer());
children.push(redBullet('16个月合计领取差费补助9,615元，但全部基于3,000公里保底而非实际里程（实际仅11,025公里，平均每月约689公里）。'));
children.push(redBullet('出车天数持续超过日历天数属于逻辑不可能，除非存在"多车统计合并"情形，但合并统计亦不符合单人出勤逻辑。'));
children.push(redBullet('建议：调取王小虎对应期间的派车单、车辆行驶记录（GPS系统）及考勤记录进行比对核实。'));
children.push(spacer());

children.push(h2('（三）高单月行驶公里异常'));
children.push(body('除王小虎外，以下司机在某月出车天数超月历天数同时公里数极高，真实性存疑：'));
children.push(bullet('王红平（2025-09）：31天出车，行驶13,614公里，日均439公里，为全员最高月公里数。'));
children.push(bullet('王志军（2025-04）：33天出车（超出3天），行驶8,602公里，补助4,861元，为全员最高月补助之一。'));
children.push(bullet('童庆龙（2025-08）：22天出车，行驶12,896公里，日均586公里，创全数据集唯一"日均>500公里"记录。单月补助4,814元。'));
children.push(spacer());

// ============================================================
// 五、个人补助过高分析
// ============================================================
children.push(h1('五、个人补助过高及真实性质疑（风险4）'));
children.push(body('【关注点】个别司机补助金额过高，真实性存疑。'));
children.push(spacer());

children.push(h2('（一）累计补助TOP10（16个月合计）'));

const top10Data = [
  ['司机姓名', '出现月份', '16月实际公里', '16月计提公里', '16月累计补助(元)', '月均补助(元)', '单月最高补助(元)', '风险备注'],
  ['王红平', '16', '153,677', '155,108', '54,321', '3,395', '4,915', '长期高强度，2025-09单月13,614公里，真实性需核实'],
  ['姚东旭', '16', '134,370', '136,019', '45,529', '2,846', '4,126', ''],
  ['郑旺', '16', '119,789', '124,843', '43,245', '2,703', '4,339', ''],
  ['王志军', '16', '112,817', '120,494', '39,005', '2,438', '4,861', '2025-04月33天+4,861元'],
  ['魏福庄', '16', '117,783', '120,626', '38,245', '2,390', '5,089', '2025-04月补助5,089元，全员最高月补助记录之一'],
  ['李青松', '16', '115,177', '116,796', '37,566', '2,348', '3,513', ''],
  ['杨营波', '16', '75,955', '92,498', '36,742', '2,297', '5,125', '2025-05补助5,125元，需核实路单'],
  ['童庆龙', '16', '98,404', '99,466', '35,231', '2,202', '4,814', '2025-08日均586km，全数据集唯一超500km记录'],
  ['何利平', '16', '72,548', '94,734', '34,397', '2,150', '5,772', '2025-07计提公里10,168远超实际2,555，差距悬殊'],
  ['李滨', '16', '101,067', '106,166', '34,374', '2,148', '3,538', ''],
];
const colW7 = [780, 780, 1200, 1200, 1400, 1100, 1200, 2780];
const top10Table = new Table({
  width: { size: 10440, type: WidthType.DXA },
  columnWidths: colW7,
  rows: top10Data.map((row, i) => new TableRow({
    children: row.map((cell, j) => {
      if (i === 0) return headerCell(cell, colW7[j]);
      const isRisk = j === 7 && cell.length > 1;
      return dataCell(cell, colW7[j], j === 0 || j === 7 ? AlignmentType.LEFT : AlignmentType.RIGHT, j === 0, isRisk ? 'C00000' : null);
    })
  }))
});
children.push(top10Table);
children.push(spacer());

children.push(h2('（二）重点异常个案'));
children.push(h3('1. 何利平（2025-07）：计提公里=实际公里×3.98'));
children.push(body('2025年7月：何利平实际行驶2,555公里，计提公里达10,168公里，差额7,613公里，差距比率约3:1。月补助5,771.60元，为全员最高月补助记录（非节假日月）。计提公里与实际公里的极大差距需要确认是否存在"任外出车"叠加计算或数据录入错误。'));
children.push(spacer());
children.push(h3('2. 魏福庄（2025-04）：月补助5,089元'));
children.push(body('2025年4月：出车29天（超出月历30天-1），行驶8,548公里，计提8,664公里，周末出车5天（补助325元），总补助5,088.83元。此为全部16个月数据中最高单月个人补助。'));
children.push(spacer());
children.push(h3('3. 杨营波（2025-05）：月补助5,125元'));
children.push(body('2025年5月：出车23天，行驶10,063公里，计提10,458公里，周末出车3天，月补助5,124.60元。日均行驶437公里（远超正常区间），需提供对应路单逐日核验。'));
children.push(spacer());
children.push(h3('4. 童庆龙（2025-08）：日均行驶586公里'));
children.push(body('2025年8月：出车22天，行驶12,896公里，日均约586公里。依据中国公路规范，正常小汽车长途单日行驶极限约500—600公里，22天平均达到此上限高度不合理。月补助4,813.60元。建议调取GPS行驶记录核实。'));
children.push(spacer());

// ============================================================
// 六、周末额外补助风险
// ============================================================
children.push(h1('六、周末额外补助合规性分析（风险5）'));
children.push(body('【关注点】周末出车补助标准合理性及超标发放情况。'));
children.push(spacer());

children.push(h2('（一）制度规定'));
children.push(bullet('双休日出车：65元/天'));
children.push(bullet('法定节假日出车：300元/天（出车3天以内），超过3天的部分不再按节假日标准，归入串休日或双休日；'));
children.push(bullet('法定节假日串休日出车：65元/天；'));
children.push(bullet('京津冀范围月包车双休日补助上限：每月4天；非月包车：每月5天；'));
children.push(bullet('执行非京津冀外埠常驻任务：可据实申请双休日补助，上限每月8天。'));
children.push(spacer());

children.push(h2('（二）异常案例'));
children.push(body('以下司机出现单月周末出车天数达6—8天（超过月包车上限4天、非月包车上限5天）：'));
children.push(spacer());

const weekendData = [
  ['月份', '司机', '周末出车天数', '出车总天数', '月补助(元)', '超标情况（参考上限5天）'],
  ['2025-11', '童庆龙', '8', '31', '3,350', '超出3天，多发195元'],
  ['2025-07', '王胜路', '8', '30', '2,550', '超出3天，多发195元'],
  ['2025-08', '王胜路', '8', '31', '2,699', '超出3天，多发195元'],
  ['2025-10', '王建涛', '8', '29', '2,820', '超出3天，多发195元'],
  ['2025-11', '于善虎', '8', '30', '2,630', '超出3天，多发195元'],
  ['2025-04', '王建涛', '8', '31', '3,087', '超出3天，多发195元'],
  ['2025-11', '王胜路', '8', '31', '2,669', '超出3天，多发195元'],
  ['2025-03', '苏保国', '7', '27', '2,128', '超出2天，多发130元（若月包车上限4天则超3天）'],
  ['2025-09', '王红平', '6', '31', '4,915', '超出1天，多发65元'],
  ['2025-08', '王红平', '6', '32', '4,052', '超出1天，多发65元'],
];
const colW8 = [780, 780, 1100, 1100, 1100, 4580];
const weekendTable = new Table({
  width: { size: 9440, type: WidthType.DXA },
  columnWidths: colW8,
  rows: weekendData.map((row, i) => new TableRow({
    children: row.map((cell, j) => {
      if (i === 0) return headerCell(cell, colW8[j]);
      return dataCell(cell, colW8[j], j <= 1 || j === 5 ? AlignmentType.LEFT : AlignmentType.CENTER, false, j === 5 ? 'C00000' : null);
    })
  }))
});
children.push(weekendTable);
children.push(spacer());

children.push(h2('（三）节日补助特殊月份（2024年12月）'));
children.push(body('2024年12月节日出车天次汇总高达631天（16个月中最高），当月合计补助94,079元（亦高于2025年1月），显著高于其他节假日较少的月份。需核实：'));
children.push(bullet('2024年12月何来如此多的节日出车天次？是否包含元旦假期（1天）及其串休日，统计口径是否存在扩大解释？'));
children.push(bullet('单日300元节日补助，631天合计节日补助约18.93万元，占当月总补助的重大比例，需逐笔核验路单。'));
children.push(spacer());

// ============================================================
// 七、其他风险点
// ============================================================
children.push(h1('七、其他风险点汇总'));
children.push(spacer());

children.push(h2('（一）统计口径风险'));
children.push(bullet('王小虎"出车天数"持续超月历天（最高47天），可能存在统计口径混淆（将多辆车出车天数叠加），导致出勤天数虚高。'));
children.push(bullet('部分司机同月驾驶多辆车（路单分行列示），合计出车天数时存在重叠计算可能。'));
children.push(spacer());

children.push(h2('（二）保底计提与实际差距极大'));
children.push(bullet('全部16个月，计提公里合计5,765,359公里，实际行驶4,151,036公里，差额1,614,323公里（占计提的28%）。'));
children.push(bullet('这意味着28%的补助基础是"虚拟公里"，由保底机制产生。在制度本身合规的前提下，虚增金额约=1,614,323×0.2元/公里≈32.3万元（按最低档0.2元/公里估算）。'));
children.push(spacer());

children.push(h2('（三）数据质量风险'));
children.push(bullet('2025年7月文件含"一队(2)"和"X月计算基础表1(2)"等重复表单，存在数据重复或调整记录，需核查该月是否存在数据修正情形。'));
children.push(bullet('2024年12月附件表中存在"11月补助错误"修正列，表明历史存在补助计算错误及事后调整，需了解调整审批流程是否合规。'));
children.push(spacer());

// ============================================================
// 八、审计建议
// ============================================================
children.push(h1('八、审计建议'));
children.push(spacer());

const auditRec = [
  ['序号', '建议事项', '建议对象', '优先级'],
  ['1', '调取王小虎对应各月派车单、GPS轨迹记录及考勤表，核实出车天数是否存在虚报，并追溯2024年12月以前数据。', '经营财务室+信息部门', '高'],
  ['2', '调取何利平2025年7月路单明细，核实计提10,168公里vs实际2,555公里差异来源，排除录入错误或重复计提。', '经营财务室', '高'],
  ['3', '调取童庆龙2025年8月GPS行驶日志，核实日均586公里的真实性。', '经营财务室+信息部门', '高'],
  ['4', '对王建涛、王胜路、童庆龙、于善虎等周末出车8天的月份，核实是否属于"非京津冀外埠常驻任务"（上限8天），若为普通京津冀任务则属超标，应追缴多发补助（每超1天65元）。', '经营财务室', '高'],
  ['5', '核查差旅费管理细则的最新有效版本，确认：①本细则经何次会议审批通过（职代会/党委会/行政会）；②2013年46号文后是否有后续修订文件；③保底公里数是否经过重新测算。', '综合管理部', '中'],
  ['6', '2024年12月节日出车631天次合计数异常偏高，应逐笔核验路单及节日认定依据，重点关注节日统计口径是否合规。', '经营财务室', '中'],
  ['7', '对16个月累计补助超过3万元的司机（共10人，合计约378,978元），抽取至少3个月的路单、派车单进行穿行测试，核实出车数据真实性。', '审计组', '中'],
  ['8', '建议单位建立GPS与路单自动比对机制，将GPS实际里程作为计提公里的校验数据，从制度层面防范里程虚报。', '信息部门+经营财务室', '低'],
];
const colW9 = [400, 5400, 1900, 700];
const auditTable = new Table({
  width: { size: 8400, type: WidthType.DXA },
  columnWidths: colW9,
  rows: auditRec.map((row, i) => new TableRow({
    children: row.map((cell, j) => {
      if (i === 0) return headerCell(cell, colW9[j]);
      const isHigh = row[3] === '高' && j === 3;
      return dataCell(cell, colW9[j], j === 0 || j === 3 ? AlignmentType.CENTER : AlignmentType.LEFT, false, isHigh ? 'C00000' : null);
    })
  }))
});
children.push(auditTable);
children.push(spacer());

// 签名
children.push(spacer());
children.push(new Paragraph({
  border: { top: { style: BorderStyle.SINGLE, size: 4, color: '999999', space: 1 } },
  spacing: { before: 400, after: 200 },
  children: [new TextRun('')]
}));
children.push(new Paragraph({
  alignment: AlignmentType.RIGHT,
  spacing: { before: 100, after: 60 },
  children: [new TextRun({ text: '审 计 人：________________', size: 20, font: 'Arial' })]
}));
children.push(new Paragraph({
  alignment: AlignmentType.RIGHT,
  spacing: { before: 100, after: 60 },
  children: [new TextRun({ text: '复 核 人：________________', size: 20, font: 'Arial' })]
}));
children.push(new Paragraph({
  alignment: AlignmentType.RIGHT,
  spacing: { before: 100, after: 60 },
  children: [new TextRun({ text: '日    期：________________', size: 20, font: 'Arial' })]
}));

// ========= 生成文档 =========
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: 'Arial', size: 20 } }
    }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1080, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: '2E75B6', space: 1 } },
          children: [new TextRun({ text: '车辆服务中心司机差旅补助审计风险分析报告（2024.12—2026.03）', size: 16, font: 'Arial', color: '595959' })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: '第 ', size: 16, font: 'Arial', color: '595959' }),
            new TextRun({ children: [PageNumber.CURRENT], size: 16, font: 'Arial', color: '595959' }),
            new TextRun({ text: ' 页', size: 16, font: 'Arial', color: '595959' }),
          ]
        })]
      })
    },
    children
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('D:/Users/12844/Desktop/小车队/司机差旅补助审计风险分析报告.docx', buf);
  console.log('DONE: 报告已生成');
});
