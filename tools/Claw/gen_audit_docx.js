const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType,
  HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageNumber, PageBreak
} = require("docx");

const border = { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 60, bottom: 60, left: 100, right: 100 };
const headerBorder = { style: BorderStyle.SINGLE, size: 1, color: "2E75B6" };
const headerBorders = { top: headerBorder, bottom: headerBorder, left: headerBorder, right: headerBorder };

function makeHeaderCell(text, width) {
  return new TableCell({
    borders: headerBorders,
    width: { size: width, type: WidthType.DXA },
    margins: cellMargins,
    shading: { fill: "2E75B6", type: ShadingType.CLEAR },
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold: true, font: "Arial", size: 20, color: "FFFFFF" })]
    })]
  });
}

function makeDataCell(text, width, opts) {
  opts = opts || {};
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    margins: cellMargins,
    shading: opts.shading ? { fill: opts.shading, type: ShadingType.CLEAR } : undefined,
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: opts.alignment || AlignmentType.LEFT,
      children: [new TextRun({
        text: text || "",
        font: "Arial",
        size: 20,
        bold: !!opts.bold,
        color: opts.color || "000000"
      })]
    })]
  });
}

function h1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] });
}
function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] });
}
function h3(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] });
}
function p(text, opts) {
  opts = opts || {};
  return new Paragraph({
    spacing: { after: 100, line: 340 },
    alignment: opts.alignment || AlignmentType.LEFT,
    children: [new TextRun({ text: text || "", font: "Arial", size: 21, bold: !!opts.bold, color: opts.color || "000000" })]
  });
}
function pRuns(runs) {
  return new Paragraph({
    spacing: { after: 100, line: 340 },
    children: runs.map(function(r) {
      return new TextRun({ font: "Arial", size: 21, text: r.text || "", bold: !!r.bold, color: r.color || "000000" });
    })
  });
}
function blank() {
  return new Paragraph({ spacing: { after: 60 }, children: [] });
}
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

var children = [];

// ===== 封面 =====
children.push(new Paragraph({
  spacing: { after: 200, before: 800 },
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "辽河油田公司", bold: true, font: "Arial", size: 40, color: "1A3C6E" })]
}));
children.push(new Paragraph({
  spacing: { after: 600 },
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "业务生命周期审计分析报告", bold: true, font: "Arial", size: 40, color: "1A3C6E" })]
}));
children.push(p("审计对象：辽河油田公司（单位编号：0102001）"));
children.push(p("资料依据：内部控制管理手册（2026版）、合同结算流程调查表（0515）、合同管理办法（2023年56号文）、业务活动层面风险数据库（1255条）、权限手册、信息系统统计表"));
children.push(p("报告日期：2026年5月"));
children.push(p("审计视角：业务生命周期全链条——合同签订 → 履约确认 → 发票预制 → 结算审批 → 付款 → 档案归档"));
children.push(blank());

// ===== 一、总体评价 =====
children.push(h1("一、总体评价"));
children.push(p("辽河油田公司建立了以COSO框架为基础的内部控制体系，覆盖业务流程200余个，系统支撑涵盖大集中ERP、合同管理系统、A8、云梦泽、司库系统等9套统建系统，整体框架较为完整。但经深入审阅相关资料，发现在业务生命周期各关键节点存在16项实质性缺陷，分属流程设计缺陷、审批管控漏洞、系统集成风险、权限管理不足、归档合规五大类别，具体说明如下。"));
children.push(blank());

// ===== 二、缺陷识别与分析 =====
children.push(h1("二、业务生命周期缺陷识别与分析"));

// 第一环节
children.push(h2("【第一环节】合同签订前——立项申报与资质审核"));

children.push(h3("缺陷1：合同申报阶段供应商资质核验控制薄弱"));
children.push(pRuns([{ text: "发现依据：", bold: true }, { text: "风险数据库条目 MP09.01.01.03-F01合同双方没有资质或履行合同能力被列为一般控制（非关键控制点），表明该环节无关键控制点覆盖。" }]));
children.push(p("问题描述：", { bold: true }));
children.push(p("• 合同管理办法（56号文）规定由企管法规部统一归口管理，但合同相对方的资质审核职责未在制度层面明确区分形式审核与实质审核；"));
children.push(p("• 供应商准入资格在云梦泽平台管理，但合同系统与云梦泽之间采购订单双向传输，合同签订时是否强制联动校验供应商在库状态，制度上无明文约束；"));
children.push(p("• 合同申报环节的F01（资质问题）、F02（未纳入预算）、F03（资料不齐）三项风险全部仅设一般控制，缺乏关键控制兜底。"));
children.push(pRuns([{ text: "风险等级：", bold: true }, { text: "较高" }]));
children.push(p("改进建议：在合同管理系统中设置强制校验节点——合同创建时自动调取云梦泽供应商状态，未通过准入或资质过期者禁止推进；将F01升级为关键控制点（K级），纳入年度审计必查项。"));
children.push(blank());

children.push(h3("缺陷2：合同签订授权体系存在层级空白"));
children.push(pRuns([{ text: "发现依据：", bold: true }, { text: "权限手册19.1合同管理各类合同审批权限与实际合同管理办法（56号文）对比分析。" }]));
children.push(p("问题描述：", { bold: true }));
children.push(new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [2000, 2360, 2360, 2640],
  rows: [
    new TableRow({ children: [makeHeaderCell("合同类型", 2000), makeHeaderCell("权限手册审批门槛", 2360), makeHeaderCell("56号文重大合同标准", 2360), makeHeaderCell("差异分析", 2640)] }),
    new TableRow({ children: [makeDataCell("技术开发/服务/咨询合同", 2000), makeDataCell("100万以上总经理审批", 2360), makeDataCell("500万以上重大合同", 2360), makeDataCell("100万~500万之间管理层级不明", 2640)] }),
    new TableRow({ children: [makeDataCell("承揽/服务合同", 2000), makeDataCell("300万以上总经理审批", 2360), makeDataCell("同上", 2360), makeDataCell("300万以下由谁审批未明示", 2640)] }),
    new TableRow({ children: [makeDataCell("建设工程施工合同", 2000), makeDataCell("500万以上总经理审批", 2360), makeDataCell("同上", 2360), makeDataCell("500万以下二级单位自行审批，下限无规定", 2640)] }),
    new TableRow({ children: [makeDataCell("买卖合同（非贸易）", 2000), makeDataCell("1000万以上总经理审批", 2360), makeDataCell("500万以上", 2360), makeDataCell("500万~1000万存在重大合同但无总经理审批", 2640)] }),
  ]
}));
children.push(blank());
children.push(p("• 权限手册与56号文对重大合同界定逻辑不一致，500万~1000万区间的买卖合同属于56号文重大合同，但权限手册无明确对应审批要求；"));
children.push(p("• 二级单位自行签订合同的内部授权程序在资料中未见规范约束。"));
children.push(pRuns([{ text: "风险等级：", bold: true }, { text: "高", color: "CC6600" }]));
children.push(p("改进建议：修订权限手册19.1条款，与56号文统一重大合同门槛为500万，补充各层级完整授权矩阵；二级单位签约授权须上报企管法规部备案。"));
children.push(blank());

// 第二环节
children.push(h2("【第二环节】合同履约——验收与确认"));

children.push(h3("缺陷3：工程验收专业资质要求执行未被列为关键控制"));
children.push(pRuns([{ text: "发现依据：", bold: true }, { text: "风险数据库 KP01.02.01.01.03-F02、KP01.02.01.02.03-F02进行专业检验的人员没有必要的资质均为一般控制。" }]));
children.push(p("• 物探、探井、油田工程多类验收流程中，验收人员资质均仅设一般控制，无K级关键控制点约束；"));
children.push(p("• 合同结算流程调查表显示，工程验收确认通过A8/EPM/SAP系统执行，但系统中是否有验收人员资质校验字段，制度上无明确要求；"));
children.push(p("• 结算表单为Excel模板，签字栏无角色资质要求说明。"));
children.push(pRuns([{ text: "风险等级：", bold: true }, { text: "较高" }]));
children.push(p("改进建议：将验收人员资质要求升级为关键控制点；系统侧在A8/SAP验收节点增加验收人角色绑定。"));
children.push(blank());

children.push(h3("缺陷4：履约过程中合同变更、工程量签证控制不完整"));
children.push(pRuns([{ text: "发现依据：", bold: true }, { text: "结算表单目录中存在外包工程工作量签证单.docx，系线下Word文档；流程目录有KP02.01.02.04施工图设计变更管理但仅在KP建设工程领域，其他合同类型无对应变更流程。" }]));
children.push(p("• 工程量签证单为线下文档，与SAP/A8系统无集成，存在签证金额与系统结算金额不一致、人工录入差错的风险；"));
children.push(p("• 设计变更仅有KP02类建设工程流程覆盖，油田服务类工程（钻井、测井、试油等合同）的现场变更签证无对应流程文档；"));
children.push(p("• 合同履行情况报告单仅要求合同承办部门负责人签字，未要求财务、业务、法务三方会签。"));
children.push(pRuns([{ text: "风险等级：", bold: true }, { text: "较高" }]));
children.push(p("改进建议：工程量签证单纳入合同管理系统线上管理，与SAP结算单强制关联；补充油田服务合同变更签证控制流程；合同履行情况报告单增加财务核实意见栏。"));
children.push(blank());

// 第三环节
children.push(h2("【第三环节】发票预制与结算发起"));

children.push(h3("缺陷5：发票预制环节控制规则过于宽松"));
children.push(pRuns([{ text: "发现依据：", bold: true }, { text: "风险数据库 MP04.01.04.02.05-F01采购发票的信息不完整、不准确仅为一般控制；F02物资采购结算未经有效审批为K3关键控制。" }]));
children.push(p("• 发票预制（BPM审批）由业务部门在SAP中操作，F01仅设一般控制，系统侧无强制校验逻辑说明；"));
children.push(p("• 结算审批前是否需要与合同金额、验收工程量进行三单比对（合同-验收单-发票），在制度和系统配置层面均未明确；"));
children.push(p("• 调查表中内部结算明确记载系统自动生成付款单，无审批流程，内部结算发票的合规性审核完全依赖系统自动化。"));
children.push(pRuns([{ text: "风险等级：", bold: true }, { text: "高", color: "CC6600" }]));
children.push(p("改进建议：在SAP发票预制节点增加三单比对强制逻辑（合同→验收→发票金额/数量自动比对），差异超过阈值（如±5%）须经额外审批；内部结算系统自动化规则应形成书面文档，纳入信息系统控制审计范围。"));
children.push(blank());

children.push(h3("缺陷6：多合同类型结算审批起点不统一，存在管控盲区【立即整改】"));
children.push(pRuns([{ text: "发现依据：", bold: true }, { text: "合同结算流程调查表草稿Sheet对比分析——不同合同类型的结算流程节点数量差异显著。" }]));
children.push(new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [1800, 2520, 2520, 2520],
  rows: [
    new TableRow({ children: [makeHeaderCell("合同类型", 1800), makeHeaderCell("结算发起方式", 2520), makeHeaderCell("审批路径", 2520), makeHeaderCell("差异说明", 2520)] }),
    new TableRow({ children: [makeDataCell("油田钻井工程", 1800), makeDataCell("EPM系统确认→SAP发票预制BPM", 2520), makeDataCell("业务部门→财务结算审批", 2520), makeDataCell("相对规范", 2520)] }),
    new TableRow({ children: [makeDataCell("建设工程施工", 1800), makeDataCell("A8确认→SAP发票预制BPM", 2520), makeDataCell("含工程审计环节", 2520), makeDataCell("含专项控制", 2520)] }),
    new TableRow({ children: [makeDataCell("劳务合同", 1800), makeDataCell("业务部门发起→直接结算", 2520), makeDataCell("业务部门自行审批", 2520), makeDataCell("审批链最短", 2520)] }),
    new TableRow({ children: [makeDataCell("内部结算（关联）", 1800), makeDataCell("系统自动生成", 2520, { color: "CC0000", bold: true }), makeDataCell("无审批流程", 2520, { color: "CC0000", bold: true }), makeDataCell("管控最弱", 2520, { color: "CC0000", bold: true })] }),
    new TableRow({ children: [makeDataCell("供用水电气热", 1800), makeDataCell("业务部门审批", 2520), makeDataCell("二级单位自行简化", 2520), makeDataCell("审批弹性最大", 2520)] }),
  ]
}));
children.push(blank());
children.push(p("• 劳务合同和供用类合同的结算发起由业务部门主导，审批流程最简化，缺乏财务独立审核节点；"));
children.push(p("• 内部结算（关联交易）无审批流程是最突出的设计缺陷，关联交易的准确性和公允性缺乏事前控制。"));
children.push(pRuns([{ text: "风险等级：", bold: true }, { text: "极高", color: "CC0000" }]));
children.push(p("改进建议：内部结算不应完全依赖系统自动化，须设置财务复核确认节点；劳务类合同结算须补充财务部门独立审核签字；关联结算价格应设置系统预警机制。"));
children.push(blank());

// 第四环节
children.push(h2("【第四环节】结算审批——付款审批"));

children.push(h3("缺陷7：付款审批层级允许自行简化缺乏上限约束【立即整改】"));
children.push(pRuns([{ text: "发现依据：", bold: true }, { text: "合同结算流程调查表明确记载：二级单位结合自身规模，自行简化审批流程。" }]));
children.push(p("• 公司级付款审批按金额设四档（<1万、<30万、<5000万、≥5000万），但允许二级单位自行简化，同一金额在不同二级单位可能经历截然不同的审批层级；"));
children.push(p("• 自行简化无明确下限——是否允许简化至单人审批、是否允许绕过财务部门，制度上无约束；"));
children.push(p("• 风险数据库 MP02.01.05.01.09-B01银行UK权限设置不合规为K1关键控制，说明资金安全已有认识，但付款审批流程弹性与资金安全控制之间存在矛盾。"));
children.push(pRuns([{ text: "风险等级：", bold: true }, { text: "极高", color: "CC0000" }]));
children.push(p("改进建议：明确规定二级单位简化审批的底线要求：至少保留业务部门+财务负责人双签；禁止单人审批付款；30万以上付款必须在系统（TR/BPM）中执行，不得线下操作；对自行简化的实施方案，须上报公司财务部备案审批。"));
children.push(blank());

children.push(h3("缺陷8：5000万以上超大额付款审批流程文件化程度不足"));
children.push(p("权限手册对合同签订权限有详细规定，但付款审批的≥5000万档次对应的审批主体（总经理办公会、董事会等）在调查表中仅有档次描述，无具体流程节点。资金安全管理办法属于旧版文件（引用试行规范），与现行TR付款系统的衔接关系不清晰。"));
children.push(pRuns([{ text: "风险等级：", bold: true }, { text: "较高" }]));
children.push(p("改进建议：补充大额付款（≥3000万或按公司实际确定）集体决策规程，将超大额付款纳入总经理办公会审议事项；更新资金安全管理办法，明确与TR系统的对接关系。"));
children.push(blank());

children.push(h3("缺陷9：付款审批与合同剩余余额缺乏系统联动校验"));
children.push(p("付款审批节点未见合同累计付款/合同总额比对校验机制说明，存在超合同金额付款的风险。对于多期结算合同（如年度框架协议），每次付款是否触发合同余额预警，系统配置文件中无记录。"));
children.push(pRuns([{ text: "风险等级：", bold: true }, { text: "高", color: "CC6600" }]));
children.push(p("改进建议：在TR付款系统中设置合同余额硬性校验：每次付款前自动核查合同累计付款额，超过合同金额的付款必须触发审批升级；将线下Excel进度结算单纳入系统线上管理。"));
children.push(blank());

// 第五环节
children.push(h2("【第五环节】信息系统集成——数据完整性"));

children.push(h3("缺陷10：公共主数据编码平台（MDG）三项控制全部为否【立即整改】"));
children.push(pRuns([{ text: "发现依据：", bold: true }, { text: "信息系统统计表序号7公共数据编码平台3.0（MDG）：总体控制=否、应用控制=否、权限控制=否，系统等级为第二等级。" }]));
children.push(p("• MDG作为主数据中枢，向ERP、云梦泽、合同系统、共享平台输出组织单元、财务编码、物料编码、往来单位等核心主数据，但三项系统控制全部缺失；"));
children.push(p("• 主数据的错误或被篡改将影响所有下游系统的数据质量（合同签订的相对方、发票编码、结算金额归集等全部依赖主数据）；"));
children.push(p("• 等级为第二等级但控制均为否，与第二等级应具备基本控制要求不符。"));
children.push(pRuns([{ text: "风险等级：", bold: true }, { text: "极高", color: "CC0000" }]));
children.push(p("改进建议：立即对MDG系统补充设计并实施：(1) 权限控制——主数据的新增、变更须双人审批，财务编码变更须财务部审核；(2) 应用控制——主数据变更须留存变更日志；(3) 总体控制——定期对MDG与下游系统数据进行比对核查；将MDG升级为重点审计关注对象。"));
children.push(blank());

children.push(h3("缺陷11：油气水井生产数据系统（A2）三项控制同样全部为否"));
children.push(p("A2系统向A8系统提供油气水井生产数据，A8系统又向ERP（财务/投资/成本）传输业务数据，形成生产数据→A8→ERP→财务结算的核心数据链。A2三项控制均为否，意味着生产数据的准确性、完整性和接触安全均无系统控制。"));
children.push(pRuns([{ text: "风险等级：", bold: true }, { text: "高", color: "CC6600" }]));
children.push(p("改进建议：对A2系统补充权限管理（生产数据录入与审核岗位分离）；在A2→A8数据传输接口设置异常数据识别规则；将A2纳入年度IT审计范围。"));
children.push(blank());

children.push(h3("缺陷12：多系统并行导致结算数据在系统间存在断点风险"));
children.push(p("业务流程跨越5个以上系统（合同系统→云梦泽→A8/EPM→SAP ERP→TR），每个接口均存在数据传输失败或延迟的风险。接口数据传输方向标注为双向，但双向传输的触发条件、失败回滚机制、数据一致性校验规则均无文档记录。"));
children.push(pRuns([{ text: "风险等级：", bold: true }, { text: "较高" }]));
children.push(p("改进建议：建立系统间数据传输日志监控机制；针对关键业务接口（合同→ERP、A8→ERP）设计数据对账规程，定期（每月）进行跨系统数据比对；制定接口异常应急预案。"));
children.push(blank());

// 第六环节
children.push(h2("【第六环节】归档管理"));

children.push(h3("缺陷13：合同归档完整性控制缺乏强制信息化手段"));
children.push(p("56号文规定企管法规部负责合同管理信息系统推广应用，但实际结算表单（合同履行情况报告单、验收单、进度结算单等）均为线下纸质/Excel文件，未与合同信息系统关联。合同管理系统与ERP之间仅传输采购订单、结算单，合同的验收单据、履行记录、签证文件、付款凭证是否同步归入合同档案，无自动化机制保障。"));
children.push(pRuns([{ text: "风险等级：", bold: true }, { text: "较高" }]));
children.push(p("改进建议：设置合同关闭强制归档节点——付款完成后，系统自动触发归档核查，验收单、结算单、发票影像全部上传后方可关闭合同；将F03升级为关键控制点。"));
children.push(blank());

// 横向专题
children.push(h2("【横向专题】权限管理与职责分离"));

children.push(h3("缺陷14：权限手册格式问题影响可读性和执行性"));
children.push(p("权限手册共407行权限条目，但各条目的协办部门/单位栏位内容大量为空，协同审核机制不完整。合同管理权限（19.1）与采购管理权限（15.1~15.3）对同类业务的审批门槛存在差异，未统一口径。"));
children.push(pRuns([{ text: "风险等级：", bold: true }, { text: "一般" }]));
children.push(p("改进建议：更新权限手册电子版本，确保内容完整准确；补充协办部门栏位；统一合同签订与采购执行的审批金额门槛。"));
children.push(blank());

children.push(h3("缺陷15：合同管理与财务结算之间缺乏独立的第三方复核"));
children.push(p("合同管理办法规定企管法规部统一归口，但在结算付款流程中，企管法规部无明确参与节点。风险数据库中合同管理（MP09）仅有6条风险识别条目，远少于物资采购（MP04，400余条），与合同管理在业务链中的核心地位不匹配，说明风险识别本身存在遗漏。"));
children.push(pRuns([{ text: "风险等级：", bold: true }, { text: "高", color: "CC6600" }]));
children.push(p("改进建议：对500万以上重大合同，结算前须企管法规部出具合同履行符合性意见；扩充风险数据库中MP09合同管理的风险识别范围（当前6条明显偏少，建议补充至20条以上）。"));
children.push(blank());

children.push(h3("缺陷16：关联交易管控机制系统化程度低"));
children.push(p("辽河油田公司与集团内部单位大量内部结算依赖系统自动化，关联交易价格的公允性缺乏独立验证机制。关联价格风险仅设一般控制，未建立关联价格定期审查和与市场价格比较的机制。"));
children.push(pRuns([{ text: "风险等级：", bold: true }, { text: "高", color: "CC6600" }]));
children.push(p("改进建议：建立关联交易价格合规性审查机制，每季度对主要内部结算价格与市场价格进行比对；将关联交易结算纳入审批流程，至少须财务部门确认；F01、F02升级为关键控制点。"));

children.push(pageBreak());

// ===== 三、缺陷汇总矩阵 =====
children.push(h1("三、缺陷汇总与优先级矩阵"));
var matrixRows = [
  ["1", "合同申报资质核验控制薄弱（仅一般控制）", "合同签订前", "较高", "中"],
  ["2", "合同签订授权体系层级空白（500万~1000万买卖合同）", "合同签订", "高", "高"],
  ["3", "验收人员资质要求无关键控制点覆盖", "履约验收", "较高", "中"],
  ["4", "工程变更签证线下管理、无系统集成", "履约变更", "较高", "中"],
  ["5", "发票预制无三单比对强制逻辑", "发票预制", "高", "高"],
  ["6", "内部结算无审批流程、劳务类审批链过短", "结算发起", "极高", "立即整改"],
  ["7", "二级单位付款审批自行简化无下限约束", "付款审批", "极高", "立即整改"],
  ["8", "超大额付款（≥5000万）集体决策程序不明", "付款审批", "较高", "中"],
  ["9", "付款审批缺乏合同余额联动校验", "付款执行", "高", "高"],
  ["10", "MDG主数据系统三项控制全部为否", "系统集成", "极高", "立即整改"],
  ["11", "A2生产数据系统三项控制全部为否", "系统集成", "高", "高"],
  ["12", "多系统接口无异常处理文档", "系统集成", "较高", "中"],
  ["13", "合同归档无强制信息化手段", "归档", "较高", "中"],
  ["14", "权限手册格式异常、协办栏位大量空白", "权限管理", "一般", "低"],
  ["15", "结算缺乏独立第三方复核、MP09风险识别严重不足", "全链条", "高", "高"],
  ["16", "关联交易管控系统化程度低（价格公允性无独立验证）", "关联交易", "高", "高"],
];

var summaryTableRows = [
  new TableRow({ children: [
    makeHeaderCell("序号", 600), makeHeaderCell("缺陷描述", 3000), makeHeaderCell("环节", 1400), makeHeaderCell("风险等级", 1280), makeHeaderCell("整改优先级", 1280),
  ]})
];
matrixRows.forEach(function(row) {
  var riskColor = row[3] === "极高" ? "CC0000" : row[3] === "高" ? "CC6600" : "000000";
  var prioColor = row[4] === "立即整改" ? "CC0000" : row[4] === "高" ? "CC6600" : "000000";
  summaryTableRows.push(new TableRow({ children: [
    makeDataCell(row[0], 600, { alignment: AlignmentType.CENTER }),
    makeDataCell(row[1], 3000),
    makeDataCell(row[2], 1400, { alignment: AlignmentType.CENTER }),
    makeDataCell(row[3], 1280, { color: riskColor, bold: row[3] === "极高", alignment: AlignmentType.CENTER }),
    makeDataCell(row[4], 1280, { color: prioColor, bold: row[4] === "立即整改", alignment: AlignmentType.CENTER }),
  ]}));
});
children.push(new Table({ width: { size: 9560, type: WidthType.DXA }, columnWidths: [600, 3000, 1400, 1280, 1280], rows: summaryTableRows }));

children.push(pageBreak());

// ===== 四、立即整改 =====
children.push(h1("四、立即整改事项详述"));

children.push(h2("整改项1：内部结算须补充审批流程（缺陷6）"));
children.push(p("具体措施："));
children.push(p("1. 在ERP系统中为内部结算增设财务负责人确认节点，不允许系统完全自动化执行；"));
children.push(p("2. 对金额≥100万的内部结算，须上传对应服务/产品的验收证明；"));
children.push(p("3. 内部结算价格定期（每季度）与市场价进行比对，差异>10%须报企管法规部审核。"));
children.push(blank());

children.push(h2("整改项2：付款审批自行简化须设置底线（缺陷7）"));
children.push(p("具体措施："));
children.push(p("1. 立即发文明确：二级单位付款审批简化后，至少需要：业务经办人→业务部门负责人→财务负责人三级签批；"));
children.push(p("2. 金额>30万的付款必须通过TR系统执行，禁止线下审批；"));
children.push(p("3. 所有二级单位的简化方案须在60日内上报公司财务部备案，逾期按最严格标准执行。"));
children.push(blank());

children.push(h2("整改项3：MDG主数据系统控制补充（缺陷10）"));
children.push(p("具体措施："));
children.push(p("1. 立即组织IT部门（科技信息部）对MDG系统现状进行评估，确认当前实际控制执行情况；"));
children.push(p("2. 制定MDG访问控制矩阵，区分主数据的查询/申请/审批/发布权限；"));
children.push(p("3. 在年度内部审计计划中将MDG列为重点专项审计对象。"));

children.push(pageBreak());

// ===== 五、系统性建议 =====
children.push(h1("五、对内控体系设计的系统性建议"));

children.push(h2("建议1：提升MP09合同管理风险识别覆盖度"));
children.push(p("当前风险数据库MP09（合同管理）仅识别6条风险，而MP04（物资采购）达400余条，严重失衡。建议参照合同全生命周期（策划/谈判/签订/履行/结算/变更/归档/纠纷）补充风险识别，重点增加："));
children.push(p("• 合同条款不完整导致纠纷风险"));
children.push(p("• 结算超期付款违约风险"));
children.push(p("• 合同到期未续期或自动续期风险"));
children.push(p("• 合同权利义务不对等风险"));
children.push(blank());

children.push(h2("建议2：推动合同管理系统与ERP深度集成"));
children.push(p("当前合同系统与ERP之间仅传输采购订单、结算单，建议扩展集成范围至："));
children.push(p("• 验收单（含联合验收单）"));
children.push(p("• 工程量签证单"));
children.push(p("• 合同变更单"));
children.push(p("• 质保金台账"));
children.push(blank());

children.push(h2("建议3：建立全链条业务审计追踪机制"));
children.push(p("对重大合同（500万以上）建立从签订到付款到归档的完整审计轨迹，ERP系统中每笔付款均可追溯至对应合同编号、验收单号、发票号，实现三单一致可追溯审计。"));
children.push(blank());

children.push(h2("建议4：定期开展付款审批流程穿行测试"));
children.push(p("建议每年至少两次，抽取各金额档次的付款样本，执行穿行测试，重点关注："));
children.push(p("• 二级单位是否按规定执行审批流程"));
children.push(p("• 系统内付款单是否与纸质凭证一致"));
children.push(p("• U-Key管理是否符合资金安全管理办法要求"));

children.push(pageBreak());

// ===== 六、附：资料清单 =====
children.push(h1("六、附：主要审阅资料清单"));
var refData = [
  ["内部控制管理手册（2026版）", "01目录", "COSO框架六章，内控组织架构"],
  ["业务活动层面风险数据库", "01目录", "1255条风险识别，涵盖13个字段"],
  ["基本业务流程目录", "01目录", "200+个五级流程，KP/MP/SP三类"],
  ["信息系统及控制适用情况统计表", "01目录", "9个统建系统，控制设计情况"],
  ["辽河油田公司权限手册", "03目录", "407行权限条目，19类业务"],
  ["合同结算流程调查表0515", "根目录", "各合同类型结算流程，两个Sheet"],
  ["合同管理办法（2023年56号文）", "表单目录", "合同分类、审批权限、归口管理"],
  ["资金安全管理办法（TR付款文件）", "施工结算", "现金/银行/票据/印鉴管理要求"],
  ["合同履行情况报告单", "施工结算", "结算验收签字表单"],
];
var refRows = [
  new TableRow({ children: [makeHeaderCell("文件名称", 2900), makeHeaderCell("来源", 1500), makeHeaderCell("内容摘要", 5160)] })
];
refData.forEach(function(r) {
  refRows.push(new TableRow({ children: [
    makeDataCell(r[0], 2900, { bold: true }),
    makeDataCell(r[1], 1500, { alignment: AlignmentType.CENTER }),
    makeDataCell(r[2], 5160),
  ]}));
});
children.push(new Table({ width: { size: 9560, type: WidthType.DXA }, columnWidths: [2900, 1500, 5160], rows: refRows }));
children.push(blank());
children.push(new Paragraph({
  spacing: { after: 120, line: 340 },
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "—— 报告完 ——", font: "Arial", size: 21, italics: true })]
}));
children.push(blank());
children.push(p("本报告基于所提供资料进行案头分析，如需进行现场核查或系统功能验证，建议结合实地走访和系统演示进一步确认。"));

// ===== 生成文档 =====
var doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Arial", size: 21 } }
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: "1A3C6E" },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "2E75B6" },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 }
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 23, bold: true, font: "Arial", color: "333333" },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 }
      },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1260, bottom: 1440, left: 1260 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: "辽河油田公司业务生命周期审计分析报告", font: "Arial", size: 18, color: "888888", italics: true })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "第 ", font: "Arial", size: 18, color: "888888" }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: "888888" }),
            new TextRun({ text: " 页", font: "Arial", size: 18, color: "888888" }),
          ]
        })]
      })
    },
    children: children
  }]
});

var outPath = "D:/Users/12844/Desktop/10-梳理-----0528/辽河油田公司业务生命周期审计分析报告.docx";
Packer.toBuffer(doc).then(function(buffer) {
  fs.writeFileSync(outPath, buffer);
  console.log("DONE: " + outPath);
}).catch(function(err) {
  console.error("ERROR: " + err.message);
  process.exit(1);
});
