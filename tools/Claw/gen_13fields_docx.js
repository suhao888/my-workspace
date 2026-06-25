const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType,
  HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageNumber, PageBreak
} = require("docx");

// ===== 样式工具 =====
const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 50, bottom: 50, left: 80, right: 80 };
const headerBorder = { style: BorderStyle.SINGLE, size: 1, color: "1A3C6E" };
const headerBorders = { top: headerBorder, bottom: headerBorder, left: headerBorder, right: headerBorder };

function hCell(text, width) {
  return new TableCell({
    borders: headerBorders, width: { size: width, type: WidthType.DXA },
    margins: cellMargins, shading: { fill: "1A3C6E", type: ShadingType.CLEAR },
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold: true, font: "Arial", size: 18, color: "FFFFFF" })]
    })]
  });
}

function dCell(text, width, opts = {}) {
  const c = opts.color || "000000";
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    margins: cellMargins, verticalAlign: "center",
    shading: opts.bg ? { fill: opts.bg, type: ShadingType.CLEAR } : undefined,
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      children: [new TextRun({ text: text || "", font: "Arial", size: 18, bold: !!opts.bold, color: c })]
    })]
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1, spacing: { before: 360, after: 200 },
    children: [new TextRun({ text, bold: true, font: "Arial", size: 32, color: "1A3C6E" })]
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2, spacing: { before: 280, after: 160 },
    children: [new TextRun({ text, bold: true, font: "Arial", size: 26, color: "2E75B6" })]
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3, spacing: { before: 200, after: 120 },
    children: [new TextRun({ text, bold: true, font: "Arial", size: 22, color: "333333" })]
  });
}
function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 100, line: 360 },
    alignment: opts.align || AlignmentType.LEFT,
    children: [new TextRun({ text, font: "Arial", size: 21, bold: !!opts.bold })]
  });
}
function pBr() { return new Paragraph({ spacing: { after: 60 }, children: [] }); }

const children = [];

// ===== 封面 =====
children.push(new Paragraph({ spacing: { before: 2400, after: 200 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "13家油田公司", bold: true, font: "Arial", size: 40, color: "1A3C6E" })]
}));
children.push(new Paragraph({ spacing: { after: 400 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "合同结算业务流程横向对比分析报告", bold: true, font: "Arial", size: 36, color: "1A3C6E" })]
}));
children.push(new Paragraph({ spacing: { before: 600, after: 120 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "—— 流程缺陷识别、繁琐度分析与优化建议 ——", font: "Arial", size: 22, color: "888888", italics: true })]
}));
children.push(new Paragraph({ spacing: { before: 800, after: 120 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "审计对象：大港油田、辽河油田、西南油气田、塔里木油田、浙江油田、玉门油田、青海油田、大庆油田、吉林油田、长庆油田、冀东油田、华北油田、新疆油田", font: "Arial", size: 20, color: "666666" })]
}));
children.push(new Paragraph({ spacing: { after: 120 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "报告日期：2026年5月29日", font: "Arial", size: 20, color: "666666" })]
}));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 目录概览 =====
children.push(h1("目  录"));
children.push(p("一、审计背景与方法"));
children.push(p("二、13家油田基本信息概览"));
children.push(p("三、合同类别覆盖度对比"));
children.push(p("四、履约确认系统/平台对比"));
children.push(p("五、付款审批金额档次对比"));
children.push(p("六、付款审批链深度对比"));
children.push(p("七、13家共性缺陷识别"));
children.push(p("八、各家特有缺陷详析"));
children.push(p("九、流程繁琐度评分排名"));
children.push(p("十、系统性优化建议"));
children.push(p("十一、整改优先级矩阵"));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 一、审计背景 =====
children.push(h1("一、审计背景与方法"));
children.push(p("本次审计以中石油集团旗下13家油田公司为对象，横向对比各油田在「合同签订→履约确认→发票预制→结算审批→付款」全业务链条上的流程设计与执行情况，识别共性缺陷、特有不足及流程繁琐点，提出系统性优化建议。"));
children.push(p("审计方法："));
children.push(p("• 逐家读取13份《合同结算流程调查表》，提取合同类别、操作平台、审批层级、付款金额档次等核心指标；"));
children.push(p("• 横向对比识别各环节差异与异常；"));
children.push(p("• 结合审计专业判断，识别控制缺陷与流程冗余。"));
children.push(p("• 注：浙江油田提交格式为WPS(.et)文件，无法直接读取结构化数据，本次分析覆盖其余12家。"));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 二、基本信息概览 =====
children.push(h1("二、13家油田基本信息概览"));

const overview = [
  ["大港油田", "75", "13", "A8/SAP/EPM", "4.0", "×", "×"],
  ["辽河油田", "95", "12", "A8/EPM/SAP/TR/共享平台", "3.0", "✓", "✓"],
  ["西南油气田", "19", "13(多Sheet)", "A8/SAP/合同系统/共享平台", "—", "✓", "—"],
  ["塔里木油田", "421", "12", "业务集成平台/A8/EPM/SAP/ERP", "—", "×", "×"],
  ["浙江油田", "ET格式", "—", "无法读取", "—", "—", "—"],
  ["玉门油田", "464", "13", "A8/A15/SAP/ERP/合同2.0", "4.0", "✓", "✓"],
  ["青海油田", "507", "11", "A8~A11/EPM/SAP/云梦泽", "3.9", "×", "✓"],
  ["大庆油田", "100", "13", "A8/EPM/ERP/SAP", "—", "×", "×"],
  ["吉林油田", "71", "13", "A8/ERP/EPM/财务共享", "5.6", "✓", "✓"],
  ["长庆油田", "101", "13", "A8/A9/A10/SAP/物资共享平台", "3.7", "✓", "✓"],
  ["冀东油田", "73", "13", "A5/SAP/造价系统/线下", "5.0", "✓", "✓"],
  ["华北油田", "118", "13", "制度引用而非系统名称", "—", "×", "✓"],
  ["新疆油田", "523", "13", "生产融合/ERP/A8/A5/BPM", "—", "✓", "×"],
];

const ovTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [1300, 600, 600, 2200, 700, 1200, 1200],
  rows: [
    new TableRow({ children: [
      hCell("油田", 1300), hCell("数据行", 600), hCell("合同类别", 600),
      hCell("主要系统/平台", 2200), hCell("平均审批层数", 700),
      hCell("有并行审查", 1200), hCell("关联交易特殊流程", 1200)
    ]}),
    ...overview.map(([name, rows, types, sys, depth, parallel, related]) =>
      new TableRow({ children: [
        dCell(name, 1300, { bold: true }),
        dCell(rows, 600, { align: AlignmentType.CENTER }),
        dCell(types, 600, { align: AlignmentType.CENTER }),
        dCell(sys, 2200),
        dCell(depth, 700, { align: AlignmentType.CENTER }),
        dCell(parallel, 1200, { align: AlignmentType.CENTER, color: parallel === "×" ? "CC0000" : "008800" }),
        dCell(related, 1200, { align: AlignmentType.CENTER, color: related === "×" ? "CC0000" : "008800" }),
      ]})
    )
  ]
});
children.push(ovTable);
children.push(pBr());
children.push(p("关键发现：12家中仅4家合同签订有并行审查机制；5家内部/关联交易无特殊审批流程。"));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 三、合同类别覆盖度 =====
children.push(h1("三、合同类别覆盖度对比"));
children.push(p("13家油田均使用集团统一的合同分类体系，共13个一级类别。但部分油田存在覆盖缺失："));

const catTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [2000, 7360],
  rows: [
    new TableRow({ children: [hCell("油田", 2000), hCell("缺失/异常的合同类别", 7360)]}),
    new TableRow({ children: [dCell("辽河油田", 2000, { bold: true }), dCell("缺少「其它合同」类别", 7360)]}),
    new TableRow({ children: [dCell("塔里木油田", 2000, { bold: true }), dCell("缺少「合资合作经营合同」", 7360)]}),
    new TableRow({ children: [dCell("青海油田", 2000, { bold: true }), dCell("缺少「其它合同」「合资合作经营合同」", 7360)]}),
    new TableRow({ children: [dCell("西南油气田", 2000, { bold: true }), dCell("主Sheet仅含2类，但通过多Sheet覆盖全部13类", 7360)]}),
  ]
});
children.push(catTable);
children.push(pBr());
children.push(p("不足：青海油田缺少2个类别，可能表明合资合作类合同和地区自定义合同的管控存在盲区。"));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 四、系统/平台对比 =====
children.push(h1("四、履约确认系统/平台对比"));
children.push(p("各油田使用的系统数量和集成度差异巨大，部分油田存在严重碎片化："));

const sysTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [1400, 1000, 6960],
  rows: [
    new TableRow({ children: [hCell("油田", 1400), hCell("系统数", 1000), hCell("使用系统列表", 6960)]}),
    new TableRow({ children: [dCell("塔里木油田", 1400, { bold: true }), dCell("最多", 1000, { align: AlignmentType.CENTER }), dCell("业务集成平台、A8、EPM、SAP、ERP、运输管理平台、共享平台、合同系统等8+系统", 6960)]}),
    new TableRow({ children: [dCell("长庆油田", 1400, { bold: true }), dCell("最多", 1000, { align: AlignmentType.CENTER }), dCell("A8、A9、A10、SAP、物资共享平台、设备信息平台、油田交通运输共享平台、合同系统2.0等", 6960)]}),
    new TableRow({ children: [dCell("辽河油田", 1400, { bold: true }), dCell("6", 1000, { align: AlignmentType.CENTER }), dCell("A8、EPM、SAP、TR系统、合同系统、财务共享平台、线下纸质/网银", 6960)]}),
    new TableRow({ children: [dCell("青海油田", 1400, { bold: true }), dCell("7+", 1000, { align: AlignmentType.CENTER }), dCell("A8、A9、A10、A11、SAP、EPM、云梦泽、合同管理系统、共享平台", 6960)]}),
    new TableRow({ children: [dCell("冀东油田", 1400, { bold: true }), dCell("5", 1000, { align: AlignmentType.CENTER }), dCell("A5、SAP、造价系统、线下操作、无系统", 6960)]}),
    new TableRow({ children: [dCell("大庆油田", 1400, { bold: true }), dCell("4", 1000, { align: AlignmentType.CENTER }), dCell("A8、EPM、ERP、SAP（相对精简）", 6960)]}),
    new TableRow({ children: [dCell("吉林油田", 1400, { bold: true }), dCell("4", 1000, { align: AlignmentType.CENTER }), dCell("A8、ERP、EPM、财务共享服务", 6960)]}),
    new TableRow({ children: [dCell("华北油田", 1400, { bold: true }), dCell("异常", 1000, { align: AlignmentType.CENTER }), dCell("操作平台栏填写的是制度编号而非实际系统名称，无法识别使用的技术平台", 6960)]}),
  ]
});
children.push(sysTable);
children.push(pBr());

children.push(h2("缺陷1：系统碎片化严重，缺乏统一集成平台"));
children.push(p("塔里木油田、长庆油田、青海油田涉及8+个系统，合同→履约→结算→付款需跨多个平台操作，每次系统跳转增加数据出错、传输延迟和权限失控风险。冀东油田部分流程仍依赖线下操作和纸质文档。"));

children.push(h2("缺陷2：系统版本混乱（A5~A11并存）"));
children.push(p("青海油田使用了A8~A11四个不同版本的A系统，冀东油田使用A5。不同版本的A系统接口协议、数据结构、控制规则可能存在差异，影响跨系统数据一致性和集团层面的流程标准化。"));

children.push(h2("缺陷3：华北油田调查表填写异常"));
children.push(p("华北油田在「操作平台」栏填写的是制度文件编号（如\"CXHBYT 0605-2024合同管理程序\"），而非实际使用的技术系统名称。这可能是填表理解偏差，但也反映出该油田对系统化操作的重视程度不足。"));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 五、付款审批金额档次 =====
children.push(h1("五、付款审批金额档次对比"));
children.push(p("付款审批按金额分档是集团统一要求，但各油田实际执行的档次差异巨大："));

const tierTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [1400, 800, 7160],
  rows: [
    new TableRow({ children: [hCell("油田", 1400), hCell("档次", 800), hCell("金额档次详情", 7160)]}),
    new TableRow({ children: [dCell("玉门油田", 1400, { bold: true }), dCell("7档", 800, { align: AlignmentType.CENTER, bold: true }), dCell("小于1万 / 小于30万 / ≤300万 / 300-500万 / >500万 / <5000万 / ≥5000万", 7160)]}),
    new TableRow({ children: [dCell("冀东油田", 1400, { bold: true }), dCell("7档", 800, { align: AlignmentType.CENTER, bold: true }), dCell("小于1万 / 小于30万 / 小于800万 / 800-1600万 / <5000万 / ≥5000万 / 1600万以上", 7160)]}),
    new TableRow({ children: [dCell("吉林油田", 1400, { bold: true }), dCell("8档", 800, { align: AlignmentType.CENTER, bold: true }), dCell("<100万 / ≥100万 两大类，但表述不统一出现6种写法", 7160)]}),
    new TableRow({ children: [dCell("长庆油田", 1400, { bold: true }), dCell("6档", 800, { align: AlignmentType.CENTER }), dCell("结合资金支出总量分单位类型定档（5亿以上/1-5亿），非统一金额标准", 7160)]}),
    new TableRow({ children: [dCell("塔里木油田", 1400, { bold: true }), dCell("5档", 800, { align: AlignmentType.CENTER }), dCell("<1000万 / 1000-5000万 / ≥5000万", 7160)]}),
    new TableRow({ children: [dCell("青海油田", 1400, { bold: true }), dCell("0档", 800, { align: AlignmentType.CENTER, color: "CC0000", bold: true }), dCell("⚠️ 全部标注「不限」或「无」——付款审批无金额分档", 7160)]}),
    new TableRow({ children: [dCell("辽河油田", 1400, { bold: true }), dCell("异常", 800, { align: AlignmentType.CENTER, color: "CC0000", bold: true }), dCell("⚠️ 未填金额档次，标注为「二级单位结合自身规模，自行简化审批流程」", 7160)]}),
    new TableRow({ children: [dCell("大庆油田", 1400, { bold: true }), dCell("1档", 800, { align: AlignmentType.CENTER, color: "CC0000" }), dCell("⚠️ 「各单位按照不同结算金额配置审批流程」——实际未在调查表中体现分档", 7160)]}),
    new TableRow({ children: [dCell("华北油田", 1400, { bold: true }), dCell("异常", 800, { align: AlignmentType.CENTER, color: "CC0000" }), dCell("⚠️ 付款审批金额栏填写的是所需资料清单，而非金额档次", 7160)]}),
    new TableRow({ children: [dCell("新疆油田", 1400, { bold: true }), dCell("异常", 800, { align: AlignmentType.CENTER, color: "CC0000" }), dCell("⚠️ 仅1档，且为叙述性描述而非清晰金额标准", 7160)]}),
  ]
});
children.push(tierTable);
children.push(pBr());

children.push(h2("缺陷4：付款审批金额分档严重不统一"));
children.push(p("12家油田的金额档次从0档（青海油田）到8档（吉林油田）不等。集团层面缺乏统一的付款审批金额分档标准，导致同类付款在不同油田经历完全不同的审批流程，内控强度不可比。"));

children.push(h2("缺陷5：青海油田付款审批完全无金额约束"));
children.push(p("青海油田507条数据全部标注「不限」或「无」金额档次，意味着不论付款金额大小，审批流程一致。这在超过一定金额的大额付款场景下，存在严重的资金安全风险——例如5000万付款与500元付款使用同一审批链。"));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 六、付款审批链深度 =====
children.push(h1("六、付款审批链深度对比"));
children.push(p("付款审批链长度（审批节点数）反映控制强度，但过长的链条也意味着效率低下。"));

const chainTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [1400, 1000, 1000, 5960],
  rows: [
    new TableRow({ children: [hCell("油田", 1400), hCell("平均层级", 1000), hCell("样本数", 1000), hCell("评估", 5960)]}),
    new TableRow({ children: [dCell("吉林油田", 1400, { bold: true }), dCell("5.6层", 1000, { align: AlignmentType.CENTER, bold: true, color: "CC6600" }), dCell("7", 1000, { align: AlignmentType.CENTER }), dCell("⚠️ 审批链最长，可能存在冗余节点。典型流程：经办人→部门负责人→财务部门→分管领导→主要领导", 5960)]}),
    new TableRow({ children: [dCell("冀东油田", 1400, { bold: true }), dCell("5.0层", 1000, { align: AlignmentType.CENTER, color: "CC6600" }), dCell("7", 1000, { align: AlignmentType.CENTER }), dCell("⚠️ 审批链较长", 5960)]}),
    new TableRow({ children: [dCell("玉门油田", 1400, { bold: true }), dCell("4.0层", 1000, { align: AlignmentType.CENTER }), dCell("103", 1000, { align: AlignmentType.CENTER }), dCell("样本量大、数据可信，审批深度适中", 5960)]}),
    new TableRow({ children: [dCell("大港油田", 1400, { bold: true }), dCell("4.0层", 1000, { align: AlignmentType.CENTER }), dCell("1", 1000, { align: AlignmentType.CENTER }), dCell("样本仅1条，代表性不足", 5960)]}),
    new TableRow({ children: [dCell("青海油田", 1400, { bold: true }), dCell("3.9层", 1000, { align: AlignmentType.CENTER }), dCell("495", 1000, { align: AlignmentType.CENTER }), dCell("样本量大，但结合无金额分档来看，审批深度与金额不匹配", 5960)]}),
    new TableRow({ children: [dCell("长庆油田", 1400, { bold: true }), dCell("3.7层", 1000, { align: AlignmentType.CENTER }), dCell("42", 1000, { align: AlignmentType.CENTER }), dCell("适中", 5960)]}),
    new TableRow({ children: [dCell("辽河油田", 1400, { bold: true }), dCell("3.0层", 1000, { align: AlignmentType.CENTER, color: "CC0000" }), dCell("46", 1000, { align: AlignmentType.CENTER }), dCell("⚠️ 审批链最短，且标注「自行简化」，控制最弱", 5960)]}),
    new TableRow({ children: [dCell("塔里木油田", 1400, { bold: true }), dCell("—", 1000, { align: AlignmentType.CENTER, color: "CC0000" }), dCell("0", 1000, { align: AlignmentType.CENTER }), dCell("⚠️ 421行数据中未提取到有效审批链信息", 5960)]}),
    new TableRow({ children: [dCell("大庆油田", 1400, { bold: true }), dCell("—", 1000, { align: AlignmentType.CENTER, color: "CC0000" }), dCell("0", 1000, { align: AlignmentType.CENTER }), dCell("⚠️ 100行数据中未提取到有效审批链信息", 5960)]}),
  ]
});
children.push(chainTable);
children.push(pBr());

children.push(h2("缺陷6：付款审批链长度差异达1.9倍，缺乏统一标准"));
children.push(p("吉林油田平均5.6层 vs 辽河油田3.0层。若吉林油田的审批链确实必要，则辽河油田存在控制不足；若辽河油田的3层已足够，则吉林油田存在流程冗余。"));

children.push(h2("缺陷7：塔里木油田、大庆油田调查表质量差"));
children.push(p("塔里木油田421行、大庆油田100行数据中均未有效填写付款审批流程链，调查表填写质量差，导致无法评估实际控制强度。"));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 七、13家共性缺陷 =====
children.push(h1("七、13家共性缺陷识别"));
children.push(p("通过对12家油田（浙江油田因格式问题除外）的横向对比，识别出以下7项共性缺陷："));

children.push(h2("共性缺陷1：付款审批金额分档无集团统一标准"));
children.push(p("12家油田的金额档次从0档到8档不等，标准各自为政。青海油田完全不分档，玉门/冀东分7档，辽河标注「自行简化」。集团层面缺乏《付款审批金额分档指引》，导致同集团内不同单位对同类付款的审批强度天差地别。"));

children.push(h2("共性缺陷2：系统碎片化，数据孤岛严重"));
children.push(p("每家油田涉及4~8+个系统，跨系统数据传输依赖接口，但接口的异常处理、数据对账、回滚机制普遍缺失。冀东油田部分流程仍依赖线下操作和纸质文件。"));

children.push(h2("共性缺陷3：内部/关联交易结算管控普遍薄弱"));
children.push(p("12家中仅5家有内部/关联交易的专门审批流程。其余7家将关联交易与外部交易混同处理，或完全依赖系统自动化，缺乏价格公允性验证和独立审批。"));

children.push(h2("共性缺陷4：合同签订缺乏多部门并行审查机制"));
children.push(p("12家中仅4家（长庆、冀东、吉林、新疆）明确标注合同签订审批采用多部门「并行」审查，其余8家为串行审批。串行审查延长合同签订周期，且后序部门倾向于「附和式审批」，审查独立性不足。"));

children.push(h2("共性缺陷5：调查表填写质量参差不齐"));
children.push(p("华北油田将制度编号填入系统名称栏、大庆油田和塔里木油田关键审批链数据缺失、新疆油田仅提取1档付款标准。调查表本身是内控评估的基础工具，填写质量差反映出各单位对内控工作的重视程度和管理规范性存在差距。"));

children.push(h2("共性缺陷6：缺少合同余额与付款的联动校验"));
children.push(p("所有油田的调查表中均未见「付款前校验合同累计付款额」的控制节点。对于多期结算合同（如年度框架协议），存在超合同金额付款的风险。"));

children.push(h2("共性缺陷7：结算归档缺乏强制信息化手段"));
children.push(p("验收单、工程量签证单、付款审批单等关键结算资料普遍为线下纸质/Excel文件，未与合同管理系统强制关联。合同关闭前无归档完整性自动核查机制。"));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 八、各家特有缺陷 =====
children.push(h1("八、各家特有缺陷详析"));

children.push(h2("辽河油田（最高风险）"));
children.push(p("• 「二级单位结合自身规模，自行简化审批流程」——这是制度层面的显性授权，但无下限约束，可能导致单人审批；"));
children.push(p("• 内部结算「系统自动生成付款单，无审批流程」——覆盖全部12类合同，关联交易缺乏事前控制；"));
children.push(p("• 审批链平均仅3.0层，为12家中最短。"));

children.push(h2("青海油田"));
children.push(p("• 付款审批全部标注「不限」金额——507条数据无一例外，大额付款无升级审批机制；"));
children.push(p("• 使用A8~A11四个版本，系统版本碎片化；"));
children.push(p("• 缺少合资合作经营合同、其它合同2个类别。"));

children.push(h2("华北油田"));
children.push(p("• 调查表填写异常：操作平台栏填写制度编号而非系统名称，付款金额栏填写资料清单而非金额；"));
children.push(p("• 合同类别识别失败——一级类别显示为序号而非名称（1.0~13.0），表明数据填写不规范；"));
children.push(p("• 无法有效评估其系统化程度和付款审批金额标准。"));

children.push(h2("塔里木油田"));
children.push(p("• 421行数据中未提取到有效付款审批链信息，调查表关键字段缺失严重；"));
children.push(p("• 系统数量最多（8+个），但集成度存疑；"));
children.push(p("• 缺少合资合作经营合同类别。"));

children.push(h2("大庆油田"));
children.push(p("• 付款审批仅1档「各单位按照不同结算金额配置审批流程」——实际上未在调查表中体现分档标准；"));
children.push(p("• 100行数据中未提取到有效审批链信息；"));
children.push(p("• 无合同签订并行审查，无关联交易特殊流程。"));

children.push(h2("吉林油田"));
children.push(p("• 审批链最长（5.6层），可能存在冗余审批节点；"));
children.push(p("• 金额档次表述混乱——同一100万门槛出现6种不同写法；"));
children.push(p("• 部分流程仍在线下操作（纸质工作量确认单）。"));

children.push(h2("冀东油田"));
children.push(p("• 审批链5.0层偏高；"));
children.push(p("• 使用A5系统（版本老旧），且有造价系统独立运行；"));
children.push(p("• 存在线下操作和无系统操作流程。"));

children.push(h2("西南油气田"));
children.push(p("• 采用多Sheet分合同类型填写（6个Sheet），结构复杂，汇总分析困难；"));
children.push(p("• 主Sheet仅含2类合同数据，数据分散。"));

children.push(h2("玉门油田"));
children.push(p("• 相对规范，但金额档次达7档（过多），可能增加管理复杂度；"));
children.push(p("• 系统引用中出现「spa」（疑为SAP拼写错误），需核实。"));

children.push(h2("长庆油田"));
children.push(p("• 付款审批金额标准结合单位资金支出总量而非统一金额，复杂度高；"));
children.push(p("• A8~A10三个A系统版本并存。"));

children.push(h2("大港油田"));
children.push(p("• 操作平台栏填写的是制度编号而非系统名称（与华北油田类似）；"));
children.push(p("• 付款审批链仅1条有效样本，调查表数据质量差；"));
children.push(p("• 无合同签订并行审查，无关联交易特殊流程。"));

children.push(h2("新疆油田"));
children.push(p("• 523行数据量最大，但仅提取1档付款标准，且为叙述性描述；"));
children.push(p("• 无关联交易特殊流程。"));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 九、流程繁琐度评分 =====
children.push(h1("九、流程繁琐度评分排名"));
children.push(p("从系统数量、审批链长度、金额分档复杂度、线上线下混合度四个维度对12家油田的流程繁琐度进行评分（1-5分，5分最繁琐）："));

const scoreTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [1200, 1000, 1000, 1000, 1000, 1000, 3360],
  rows: [
    new TableRow({ children: [
      hCell("油田", 1200), hCell("系统碎片", 1000), hCell("审批深度", 1000),
      hCell("金额分档", 1000), hCell("线上线下", 1000), hCell("综合评分", 1000),
      hCell("评价", 3360)
    ]}),
    // 数据: name, sys_frag, appr_depth, tier, mix, total, eval
    ...[
      ["塔里木油田", "5", "2", "3", "3", "3.3", "系统最多(8+)但审批链数据缺失"],
      ["长庆油田", "5", "3", "4", "3", "3.8", "系统最多、A版本混乱、金额标准最复杂"],
      ["青海油田", "4", "3", "1", "2", "2.5", "系统多但审批无金额分档，控制最弱"],
      ["冀东油田", "4", "4", "4", "4", "4.0", "含线下操作、系统老旧(A5)"],
      ["吉林油田", "3", "5", "4", "3", "3.8", "审批链最长(5.6层)、金额表述混乱"],
      ["玉门油田", "4", "3", "4", "2", "3.3", "金额7档偏多但系统相对集中"],
      ["辽河油田", "3", "2", "1", "3", "2.3", "审批链最短但控制最弱"],
      ["西南油气田", "4", "2", "2", "2", "2.5", "多Sheet结构复杂"],
      ["大庆油田", "3", "1", "1", "2", "1.8", "调查表质量最差"],
      ["大港油田", "2", "2", "2", "2", "2.0", "数据质量差"],
      ["华北油田", "1", "1", "1", "2", "1.3", "调查表填写严重异常"],
      ["新疆油田", "3", "2", "1", "2", "2.0", "数据量大但关键信息缺失"],
    ].map(([name, sf, ad, t, m, total, eval_]) => {
      const tc = parseFloat(total) >= 3.5 ? "CC0000" : parseFloat(total) >= 2.5 ? "CC6600" : "008800";
      return new TableRow({ children: [
        dCell(name, 1200, { bold: true }),
        dCell(sf, 1000, { align: AlignmentType.CENTER }),
        dCell(ad, 1000, { align: AlignmentType.CENTER }),
        dCell(t, 1000, { align: AlignmentType.CENTER }),
        dCell(m, 1000, { align: AlignmentType.CENTER }),
        dCell(total, 1000, { align: AlignmentType.CENTER, bold: true, color: tc }),
        dCell(eval_, 3360),
      ]});
    })
  ]
});
children.push(scoreTable);
children.push(pBr());
children.push(p("结论：冀东油田(4.0)和长庆/吉林(3.8)流程最为繁琐，且繁琐不等于控制力强——青海油田审批虽不繁琐但控制最弱。最理想的区间是2.5-3.0分，兼具适当控制和合理效率。"));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 十、系统性优化建议 =====
children.push(h1("十、系统性优化建议"));

children.push(h2("建议1：制定集团统一的付款审批金额分档标准"));
children.push(p("当前12家油田的金额档次从0到8档不等，建议集团层面出台《付款审批金额分档指引》，统一设定6档标准："));
children.push(p("• 第1档：≤5万元 — 经办人+部门负责人双签"));
children.push(p("• 第2档：5万~50万元 — 经办人→部门负责人→财务负责人"));
children.push(p("• 第3档：50万~300万元 — 经办人→部门负责人→财务负责人→分管领导"));
children.push(p("• 第4档：300万~1000万元 — 经办人→部门负责人→财务负责人→分管领导→总会计师"));
children.push(p("• 第5档：1000万~5000万元 — 以上+总经理审批"));
children.push(p("• 第6档：≥5000万元 — 以上+总经理办公会集体决策"));
children.push(p("各油田可在集团标准基础上收紧但不得放宽。青海油田须立即从「无分档」切换到集团标准。"));

children.push(h2("建议2：推动系统整合，减少跨平台操作"));
children.push(p("目标是将合同签订→履约确认→结算审批→付款的全链条整合到3个以内系统。优先方案："));
children.push(p("• 合同管理系统（合同签订、履约确认）→ ERP/SAP（发票预制、结算审批）→ TR/资金系统（付款执行）"));
children.push(p("• 逐步淘汰A5~A11多版本A系统，统一迁移至集团统建平台；"));
children.push(p("• 冀东油田的线下操作流程须在6个月内全部上线。"));

children.push(h2("建议3：规范关联交易结算审批"));
children.push(p("要求所有油田对内部/关联交易设置专门审批流程，至少包含："));
children.push(p("• 财务部门对关联价格的公允性确认；"));
children.push(p("• 对≥100万的内部结算须上传服务验收证明；"));
children.push(p("• 每季度对主要内部结算价格与市场价格进行比对。"));

children.push(h2("建议4：建立合同余额联动校验机制"));
children.push(p("在TR付款系统中设置合同余额硬性校验：每次付款前自动核查累计付款额是否超出合同金额。超出合同金额的付款必须触发审批升级至分管领导。"));

children.push(h2("建议5：标准化调查表填写规范"));
children.push(p("当前调查表填写质量差异巨大（华北油田填写制度编号、大庆/塔里木关键字段缺失），建议："));
children.push(p("• 下发填写样例和必填项清单；"));
children.push(p("• 对已提交调查表进行质量复核，不合格的限期重填；"));
children.push(p("• 将调查表填写质量纳入各单位内控考核指标。"));

children.push(h2("建议6：精简过长审批链，补充过短审批链"));
children.push(p("• 吉林油田（5.6层）和冀东油田（5.0层）审批链偏长，建议审查是否存在「橡皮图章」式冗余节点，目标压缩至4层以内；"));
children.push(p("• 辽河油田（3.0层）审批链偏短，在取消「自行简化」条款后，至少恢复至4层标准审批链。"));

children.push(h2("建议7：补充缺失合同类别"));
children.push(p("青海油田缺少2个类别（合资合作经营合同、其它合同），塔里木油田缺少1个类别（合资合作经营合同），须补充相应合同类型的流程设计。"));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 十一、整改优先级 =====
children.push(h1("十一、整改优先级矩阵"));

const prioTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [600, 3000, 1400, 1560, 1200, 1600],
  rows: [
    new TableRow({ children: [
      hCell("序号", 600), hCell("整改事项", 3000), hCell("涉及油田", 1400),
      hCell("风险等级", 1560), hCell("优先级", 1200), hCell("建议时限", 1600)
    ]}),
    ...[
      ["1", "辽河油田取消「自行简化审批」条款，设定审批底线", "辽河油田", "极高", "立即整改", "30日内发文"],
      ["2", "青海油田建立付款审批金额分档标准", "青海油田", "极高", "立即整改", "60日内完成"],
      ["3", "华北油田/大港油田/大庆油田/塔里木油田重新规范填写调查表", "4家", "高", "立即整改", "15日内重交"],
      ["4", "出台集团统一付款审批金额分档指引", "全部", "高", "集团层面", "90日内发布"],
      ["5", "辽河油田/大港油田/大庆油田补充关联交易专门审批流程", "3家", "高", "高", "60日内完成"],
      ["6", "系统整合：淘汰A5~A11多版本，统一平台", "青海/冀东/长庆", "高", "中", "6个月"],
      ["7", "吉林/冀东审批链审查，精简冗余节点", "2家", "中", "中", "90日内完成"],
      ["8", "冀东油田线下操作全部上线", "冀东油田", "高", "高", "6个月"],
      ["9", "建立合同余额联动校验机制", "全部", "高", "集团层面", "纳入年度IT计划"],
      ["10", "补充缺失合同类别", "青海/塔里木", "中", "中", "60日内"],
      ["11", "标准化调查表填写规范", "全部", "中", "低", "下次填报前"],
    ].map(([num, item, fields, risk, prio, deadline]) => {
      const rc = risk === "极高" ? "CC0000" : risk === "高" ? "CC6600" : "000000";
      const pc = prio === "立即整改" ? "CC0000" : prio === "高" ? "CC6600" : "000000";
      return new TableRow({ children: [
        dCell(num, 600, { align: AlignmentType.CENTER }),
        dCell(item, 3000),
        dCell(fields, 1400, { align: AlignmentType.CENTER }),
        dCell(risk, 1560, { align: AlignmentType.CENTER, bold: risk === "极高", color: rc }),
        dCell(prio, 1200, { align: AlignmentType.CENTER, bold: prio === "立即整改", color: pc }),
        dCell(deadline, 1600, { align: AlignmentType.CENTER }),
      ]});
    })
  ]
});
children.push(prioTable);
children.push(pBr());

children.push(new Paragraph({ children: [new PageBreak()] }));

// ===== 结尾 =====
children.push(new Paragraph({ spacing: { before: 400 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "—— 报告完 ——", font: "Arial", size: 22, italics: true, color: "888888" })]
}));
children.push(pBr());
children.push(p("免责声明：本报告基于13家油田提交的《合同结算流程调查表》进行案头分析。部分油田调查表填写质量较差，可能导致分析偏差。建议结合现场走访和系统穿行测试进一步验证。浙江油田因提交格式（WPS .et）无法读取，未纳入本次分析。"));

// ===== 生成文档 =====
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: "1A3C6E" },
        paragraph: { spacing: { before: 360, after: 200 } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "2E75B6" },
        paragraph: { spacing: { before: 280, after: 160 } } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: "333333" },
        paragraph: { spacing: { before: 200, after: 120 } } },
    ]
  },
  sections: [{
    properties: {
      page: { size: { width: 11906, height: 16838 }, margin: { top: 1200, right: 1100, bottom: 1200, left: 1100 } }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: "13家油田合同结算业务流程横向对比分析报告 | 2026年5月", font: "Arial", size: 16, color: "AAAAAA", italics: true })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "第 ", font: "Arial", size: 16, color: "888888" }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16, color: "888888" }),
            new TextRun({ text: " 页", font: "Arial", size: 16, color: "888888" }),
          ]
        })]
      })
    },
    children
  }]
});

const outPath = "D:/Users/12844/Desktop/13家油田合同结算业务流程横向对比分析报告.docx";
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outPath, buffer);
  console.log("DOCX generated: " + outPath);
  console.log("Size: " + (buffer.length / 1024).toFixed(1) + " KB");
});
