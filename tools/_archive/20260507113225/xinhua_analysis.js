const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageNumber, Header, Footer
} = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 6, color: "AAAAAA" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function cell(text, opts = {}) {
  const { bold = false, fill = "FFFFFF", colSpan, width = 2200, shade = ShadingType.CLEAR } = opts;
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill, type: shade },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    ...(colSpan ? { columnSpan: colSpan } : {}),
    children: [new Paragraph({
      alignment: AlignmentType.LEFT,
      children: [new TextRun({ text, bold, size: 20, font: "宋体" })]
    })]
  });
}

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 180 },
    children: [new TextRun({ text, bold: true, size: 28, font: "宋体", color: "1F4E79" })]
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: 24, font: "宋体", color: "2E74B5" })]
  });
}

function heading3(text) {
  return new Paragraph({
    spacing: { before: 160, after: 80 },
    children: [new TextRun({ text, bold: true, size: 22, font: "宋体", color: "404040" })]
  });
}

function para(text, opts = {}) {
  const { indent = true, bold = false, color = "000000", size = 22 } = opts;
  return new Paragraph({
    spacing: { before: 60, after: 60, line: 360 },
    indent: indent ? { firstLine: 440 } : {},
    children: [new TextRun({ text, bold, size, font: "宋体", color })]
  });
}

function riskRow(level, riskName, detail, levelColor) {
  const levelColors = { "高": "C00000", "中": "ED7D31", "低": "70AD47" };
  const levelFills = { "高": "FFE2E2", "中": "FFF2E2", "低": "E8F5E9" };
  return new TableRow({
    children: [
      new TableCell({
        borders,
        width: { size: 1400, type: WidthType.DXA },
        shading: { fill: levelFills[level] || "FFFFFF", type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: level, bold: true, size: 20, font: "宋体", color: levelColors[level] || "000000" })]
        })]
      }),
      new TableCell({
        borders,
        width: { size: 2200, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          children: [new TextRun({ text: riskName, bold: true, size: 20, font: "宋体" })]
        })]
      }),
      new TableCell({
        borders,
        width: { size: 5760, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          children: [new TextRun({ text: detail, size: 20, font: "宋体" })]
        })]
      }),
    ]
  });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "宋体", size: 22 } } },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1800 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1F4E79", space: 1 } },
          children: [new TextRun({ text: "河北信华能源科技集团有限公司 — 融资合规性及风险分析报告", size: 18, font: "宋体", color: "888888" })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "第 ", size: 18, font: "宋体", color: "888888" }),
            new TextRun({ children: [PageNumber.CURRENT], size: 18, font: "宋体", color: "888888" }),
            new TextRun({ text: " 页", size: 18, font: "宋体", color: "888888" }),
          ]
        })]
      })
    },
    children: [
      // ====== 封面标题 ======
      new Paragraph({ spacing: { before: 720, after: 120 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "河北信华能源科技集团有限公司", bold: true, size: 40, font: "宋体", color: "1F4E79" })] }),
      new Paragraph({ spacing: { before: 0, after: 120 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "资金规划融资合规性及主要风险分析报告", bold: true, size: 36, font: "宋体", color: "2E74B5" })] }),
      new Paragraph({ spacing: { before: 120, after: 480 }, alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "分析日期：2026年5月7日    分析依据：资金规划情况汇报（2026年4月）", size: 20, font: "宋体", color: "888888" })] }),

      // ====== 一、基本情况概述 ======
      heading1("一、企业基本资金状况概述"),
      para("据汇报文件披露，信华集团自2022年11月成立至2026年3月，现金流量净额由正转负，截至2026年3月，集团整体现金流量净额为-56,818万元，现金流严重枯竭。集团本部账面实际资金余额为-13,352万元（账面资金2,648万元，短期付息负债16,000万元）。"),
      para("累计对外分红45,136万元，其中2024年和2025年各分红19,754万元，两年分红合计占累计总分红的87.7%。本部资金已严重依赖向资金池借款维持分红义务，2024年借款2,000万元、2025年借款14,000万元，合计16,000万元。"),
      para("为应对资金枯竭，公司拟采取三类措施：（1）开通银行承兑汇票业务；（2）向商业银行多渠道融资；（3）处置部分股权回笼资金。"),

      // ====== 二、融资方式合规性分析 ======
      heading1("二、各类融资方式合规合法性分析"),

      heading2("（一）银行承兑汇票业务"),
      heading3("1. 汇报中的法律分析"),
      para('汇报文件援引《保障中小企业款项支付条例》第十一条，论证信华集团所属企业（中小型企业）不受该条款"大型企业不得强制使用非现金支付"的约束，并据此认定开通银行承兑汇票不存在法律风险。'),
      heading3("2. 合规性评价"),
      para("汇报文件对该条例的援引整体准确，逻辑推演基本成立。但需注意以下几点补充："),
      para('▶ 须确认合同相对方的知情同意：条例同时禁止"未经约定强制接受"，故须在招标文件或合同中明确约定汇票支付比例，不可在合同签署后单方面要求对方接受。', { indent: false }),
      para("▶ 承兑银行资质合规：所使用的商业汇票须通过具有相应资质的银行出具，并确保贴现渠道合规，防止涉及非法票据融资。", { indent: false }),
      para("▶ 管控流程设计须健全：文件提及需在一体化平台审批流程中增加结算中心审查权限，该机制若落实不到位，可能存在内控漏洞。", { indent: false }),

      heading2("（二）向商业银行融资"),
      heading3("1. 流动资金贷款（建设银行、工商银行）"),
      para("此类贷款为标准银行信贷产品，受《商业银行法》《贷款通则》规范，只要：①贷款用途真实，用于日常生产经营周转；②担保方式（信用担保）经过银行审核；③按时披露财务信息；则合规性较高，无重大法律障碍。"),
      para('注意：贷款用于支付分红属于"非经营性用途"，可能触发银行"贷款用途限制"条款，若将流贷资金用于支付股东分红，属于违规使用贷款，须明确资金用途边界。', { indent: false, color: "C00000" }),

      heading3("2. 应收账款融资（昆仑银行）"),
      para("以中石油应收账款为质押，属于《民法典》规定的应收账款质押业务，须在中国人民银行征信中心动产融资统一登记公示系统（PPRS）完成质押登记，否则质押不生效。须确认：①应收账款真实存在；②无重复质押；③经中石油核心企业确认。"),

      heading3("3. 供应链融资—E信通、职工惠（建设银行）"),
      para("此为银行供应链金融产品，基于真实贸易背景，合规性较高。E信通本质是银行承兑汇票的电子化替代，须保证背后交易的真实性，不得虚构贸易背景套取融资资金。职工惠用于支付劳务费等，须留存完整的劳务合同、发票等凭据。"),

      heading3("4. 电子凭证保理融资（昆仑银行）"),
      para("属于保理业务，受《民法典》第十三章和《商业银行保理业务管理暂行办法》规制。须注意：①不得以未来应收账款（尚未形成的债权）作为保理基础；②不得重复保理，即同一笔应收账款不得同时向多家机构融资；③须向中国人民银行完成保理登记。"),

      heading3("5. 融资租赁（昆仑金租）"),
      para("售后回租业务须确保租赁物权属清晰、价值评估合理（按账面净值110%操作），且须实际交付租赁物，不可仅进行合同安排而无实际资产转移，否则可能被认定为变相贷款而违反相关监管规定。《融资租赁公司监督管理暂行办法》要求租赁物须为真实存在的有形动产，须具备合法产权证明。"),

      heading3("6. 固定资产贷款（工商银行）"),
      para("须确保贷款与具体项目挂钩，贷款资金专项用于项目采购建设，不得挪作他用。项目前期贷款须有可信的后续还款来源（明确预期项目贷款或其他合法资金），抵押物（土地及在建工程）须完成正式抵押登记。"),

      // ====== 三、资金池借款合规性 ======
      heading2("（三）资金池内部借款合规性"),
      para("汇报显示，信华集团本部累计向资金池借款16,000万元，子公司（工程项目管理公司、华油蠡县新能源公司）亦存在资金池透支。内部资金池模式须关注："),
      para("▶ 须签订规范的内部借款合同，明确利率、期限、还款安排，借款利率不得偏离市场正常水平，否则可能被税务机关认定为转移利润，面临企业所得税调整风险。", { indent: false }),
      para('▶ 本部向资金池借款用于支付分红（而非生产经营），该资金流向具有特殊性，须在法律和税务层面进行充分论证，防范被认定为"抽逃出资"或违规分配。', { indent: false }),
      para("▶ 子公司透支额已超出其授信额度（工程项目管理公司透支3,941万元，超过5,000万元额度；华油蠡县新能源公司透支568万元），应确认是否有相应审批授权，否则存在内控缺陷。", { indent: false }),

      // ====== 三、主要风险点 ======
      heading1("三、主要风险点汇总"),
      para("以下风险矩阵综合评估文件所涉及的主要风险，按风险等级由高到低排序：", { indent: false }),

      new Paragraph({ spacing: { before: 120 } }),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [1400, 2200, 5760],
        rows: [
          // 表头
          new TableRow({
            tableHeader: true,
            children: [
              new TableCell({
                borders, width: { size: 1400, type: WidthType.DXA },
                shading: { fill: "1F4E79", type: ShadingType.CLEAR },
                margins: { top: 100, bottom: 100, left: 120, right: 120 },
                verticalAlign: VerticalAlign.CENTER,
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "风险等级", bold: true, size: 20, font: "宋体", color: "FFFFFF" })] })]
              }),
              new TableCell({
                borders, width: { size: 2200, type: WidthType.DXA },
                shading: { fill: "1F4E79", type: ShadingType.CLEAR },
                margins: { top: 100, bottom: 100, left: 120, right: 120 },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "风险名称", bold: true, size: 20, font: "宋体", color: "FFFFFF" })] })]
              }),
              new TableCell({
                borders, width: { size: 5760, type: WidthType.DXA },
                shading: { fill: "1F4E79", type: ShadingType.CLEAR },
                margins: { top: 100, bottom: 100, left: 120, right: 120 },
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "风险描述及潜在影响", bold: true, size: 20, font: "宋体", color: "FFFFFF" })] })]
              }),
            ]
          }),
          riskRow("高", "流动性危机风险", "截至2026年3月集团整体现金流净额为-56,818万元，本部实际可动用资金为负值。若无法及时落地融资方案，6月末可动用资金将耗尽（-4,582万元），存在资金链断裂、无法履行到期债务的极高风险。"),
          riskRow("高", "违规使用贷款风险", "流动资金贷款用于支付股东分红，属违规挪用，将触发贷款提前到期条款，并面临银行罚息、列入失信名单乃至刑事追责（违规贷款）的风险。须严格区分贷款用途，坚决不以流贷资金支付分红。"),
          riskRow("高", "大额分红可持续性风险", "公司本部无经营性收入，连续两年借款支付分红（合计16,000万元），已构成资不抵债隐患。若分红义务继续按现有约定履行，必然加速资金链断裂，并可能触发股东协议中的违约条款。"),
          riskRow("高", "股权处置估值与合规风险", "处置中蓝公司及华港公司股权涉及重大资产交易，须依法进行资产评估（须具备资质的评估机构出具评估报告）、股东会决议（须满足章程规定比例）、可能涉及竞争对手信息保护及国资监管要求，流程缺失将导致交易被撤销或承担赔偿责任。"),
          riskRow("中", "应收账款重复质押风险", "若同一笔中石油应收账款被用于多项融资（如既做保理又做质押），将构成欺诈，不仅融资无效，还将面临刑事责任追究。须建立严格的融资台账管理机制。"),
          riskRow("中", "融资租赁物权不清晰风险", "售后回租须确保租赁物合法、真实存在，若存在已抵押或权属纠纷的资产参与租赁，将导致融资租赁合同无效，并面临虚构交易的法律责任。"),
          riskRow("中", "内部资金池税务合规风险", "本部借款用于支付分红，若未按规定签订借款合同、未收取市场利率，税务机关可能认定为关联方不合理资金占用，并对免收利息部分按市场利率补征企业所得税，亦可能触发转让定价调查。"),
          riskRow("中", "保证金暂无法动用风险", "工程项目管理公司3,242万元保证金冻结在资金池外，在特定压力情况下将无法作为应急资金使用，存在进一步削减流动性的风险。"),
          riskRow("低", "银行承兑汇票操作合规风险", "若在招标文件或合同中未明确约定汇票支付比例，可能被供应商主张违反《保障中小企业款项支付条例》，引发合同争议。该风险可通过规范合同管理加以防控。"),
          riskRow("低", "信息披露与内控风险", "一体化平台审批流程若管控不到位（如结算中心审查权限未及时增设），存在内控薄弱点，可能被审计发现，影响内部合规评级。"),
        ]
      }),
      new Paragraph({ spacing: { before: 120 } }),

      // ====== 四、合规建议 ======
      heading1("四、合规建议"),

      heading2("（一）资金使用边界管理（高优先级）"),
      para("立即建立融资资金专项用途管理台账，明确各笔贷款对应的合规用途（生产经营、项目建设），严禁将任何商业银行贷款资金用于支付股东分红或偿还历史分红借款，以免触发贷款违约和刑事风险。"),

      heading2("（二）分红机制审查与重谈（高优先级）"),
      para("建议尽快组织法务和财务团队，对股东分红协议进行审查：①评估现有分红义务是否超出公司盈利能力；②依据《公司法》第二百一十条（原第一百六十六条），公司应以当年可分配利润为限进行分红，若无利润则不应强制分红；③启动与股东方的协商，暂缓或减少2026年分红，以保障日常运营资金需求。"),

      heading2("（三）股权处置合规流程（高优先级）"),
      para("在启动中蓝公司及华港公司股权处置前，须完成：①委托具有证券从业资格的资产评估机构出具估值报告；②召开股东会并形成合规决议（注意章程规定的表决比例）；③确认是否涉及国资监管要求（如有国有资产成分须履行国资委审批程序）；④在交易文件中明确知识产权、未决诉讼等声明与保证条款。"),

      heading2("（四）融资台账与质押登记（中优先级）"),
      para("建立统一的应收账款和资产融资台账，对已质押、已保理的应收账款进行标记管理，防止重复融资。应收账款质押须及时在人民银行PPRS系统完成登记，确保质押权利的对抗效力。"),

      heading2("（五）内部资金池规范化（中优先级）"),
      para("对已发生的16,000万元内部借款，补充签订规范的内部借贷协议，明确：借款期限、利率（参考LPR，建议不低于银行同期贷款利率）、还款计划；税务层面确保利息收入和支出在关联方之间对称确认，并留存定价说明文件备查。"),

      heading2("（六）票据业务合同管控（低优先级）"),
      para("在所有招标文件及采购合同中，明确注明商业汇票支付比例及期限，并保留对方签字确认记录，以规避潜在的合规争议。"),

      // ====== 五、综合结论 ======
      heading1("五、综合结论"),
      para("从合规合法性角度评估，信华集团所规划的各类融资方式在产品层面均有法律依据，属于商业金融领域的常规操作，整体合法性基础较为扎实。"),
      para("然而，企业当前面临的根本性问题不在于融资工具的合法性，而在于：一是超出盈利能力的持续高额分红义务已使公司陷入以债还债的恶性循环；二是本部完全依赖投资收益且持续借款支付分红，已呈现资不抵债特征；三是多元融资方案若不配套严格的用途管控，极有可能在执行层面触发违规风险。"),
      para('建议管理层将"规范分红机制"与"落实融资方案"并行推进，在解决短期流动性危机的同时，从根本上修复企业的资本结构和盈利能力，方能实现可持续经营。'),

      new Paragraph({ spacing: { before: 480 }, alignment: AlignmentType.RIGHT,
        children: [new TextRun({ text: "—— 分析报告生成于 2026年5月7日", size: 20, font: "宋体", color: "888888" })] }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('C:/Users/12844/WorkBuddy/20260507113225/信华集团融资合规性及风险分析报告.docx', buf);
  console.log('done');
});
