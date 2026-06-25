const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  HeadingLevel, PageNumber, Header, Footer, UnderlineType,
  LevelFormat, TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

// ─────────────── 通用样式 ───────────────
const FONT = "仿宋_GB2312";
const FONT_TITLE = "方正小标宋简体";
const FONT_HEADING = "黑体";
const FONT_SUBTITLE = "仿宋_GB2312";

function normalPara(text, options = {}) {
  return new Paragraph({
    children: [new TextRun({ text, font: FONT, size: 28, ...options })],
    spacing: { line: 480, lineRule: "auto" },
    alignment: AlignmentType.JUSTIFIED,
  });
}

function boldPara(text, options = {}) {
  return new Paragraph({
    children: [new TextRun({ text, font: FONT_HEADING, size: 28, bold: true, ...options })],
    spacing: { line: 480, lineRule: "auto" },
    alignment: AlignmentType.JUSTIFIED,
  });
}

function centerPara(text, options = {}) {
  return new Paragraph({
    children: [new TextRun({ text, font: FONT, size: 28, ...options })],
    spacing: { line: 480, lineRule: "auto" },
    alignment: AlignmentType.CENTER,
  });
}

function indentPara(text, options = {}) {
  return new Paragraph({
    children: [new TextRun({ text, font: FONT, size: 28, ...options })],
    indent: { firstLine: 560 },
    spacing: { line: 480, lineRule: "auto" },
    alignment: AlignmentType.JUSTIFIED,
  });
}

// ─────────────── 文档正文 ───────────────
const border = { style: BorderStyle.SINGLE, size: 1, color: "000000" };
const borders = { top: border, bottom: border, left: border, right: border };

function makeCell(text, w, bold = false, shade = false) {
  return new TableCell({
    borders,
    width: { size: w, type: WidthType.DXA },
    shading: shade ? { fill: "D9E1F2", type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      children: [new TextRun({ text, font: FONT, size: 26, bold })],
      alignment: AlignmentType.CENTER,
      spacing: { line: 360 },
    })]
  });
}

// ─── 资产明细表（示例行，可替换） ───
const assetTableRows = [
  new TableRow({
    tableHeader: true,
    children: [
      makeCell("序号", 600, true, true),
      makeCell("资产名称", 2200, true, true),
      makeCell("资产类别", 1600, true, true),
      makeCell("账面价值（元）", 2000, true, true),
      makeCell("是否设有抵押/质押", 2000, true, true),
      makeCell("是否提供对外担保", 1960, true, true),
    ]
  }),
  new TableRow({
    children: [
      makeCell("合计", 600, true),
      makeCell("—", 2200, true),
      makeCell("—", 1600, true),
      makeCell("以实际审计确认金额为准", 2000, true),
      makeCell("无", 2000, true),
      makeCell("无", 1960, true),
    ]
  }),
];

// ─────────────── 构建文档 ───────────────
const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: FONT, size: 28 }
      }
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT_HEADING, size: 32, bold: true, color: "000000" },
        paragraph: {
          spacing: { before: 240, after: 120, line: 480, lineRule: "auto" },
          outlineLevel: 0
        }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: FONT_HEADING, size: 28, bold: true, color: "000000" },
        paragraph: {
          spacing: { before: 180, after: 60, line: 480, lineRule: "auto" },
          outlineLevel: 1
        }
      }
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1701, right: 1418, bottom: 1701, left: 1984 }
      }
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            children: [new TextRun({ text: "铜峰会计师事务所（特殊普通合伙）", font: FONT, size: 22, color: "808080" })],
            alignment: AlignmentType.CENTER,
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: "第 ", font: FONT, size: 22, color: "808080" }),
              new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 22, color: "808080" }),
              new TextRun({ text: " 页，共 ", font: FONT, size: 22, color: "808080" }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 22, color: "808080" }),
              new TextRun({ text: " 页", font: FONT, size: 22, color: "808080" }),
            ],
            alignment: AlignmentType.CENTER,
          })
        ]
      })
    },
    children: [

      // ═══════ 封面区域 ═══════
      new Paragraph({ children: [], spacing: { line: 480 } }),
      new Paragraph({ children: [], spacing: { line: 480 } }),

      // 报告标题
      new Paragraph({
        children: [new TextRun({
          text: "关于大港工程职业技术学院划转",
          font: FONT_TITLE, size: 44, bold: true, color: "000000"
        })],
        alignment: AlignmentType.CENTER,
        spacing: { line: 480, lineRule: "auto" }
      }),
      new Paragraph({
        children: [new TextRun({
          text: "天津石油职业技术学院资产",
          font: FONT_TITLE, size: 44, bold: true, color: "000000"
        })],
        alignment: AlignmentType.CENTER,
        spacing: { line: 480, lineRule: "auto" }
      }),
      new Paragraph({
        children: [new TextRun({
          text: "无抵押担保专项审计报告",
          font: FONT_TITLE, size: 44, bold: true, color: "000000"
        })],
        alignment: AlignmentType.CENTER,
        spacing: { line: 480, lineRule: "auto" }
      }),

      new Paragraph({ children: [], spacing: { line: 480 } }),

      // 报告编号
      new Paragraph({
        children: [new TextRun({ text: "铜峰审专字〔2026〕第××号", font: FONT, size: 28, color: "595959" })],
        alignment: AlignmentType.CENTER,
        spacing: { line: 480, lineRule: "auto" }
      }),

      new Paragraph({ children: [], spacing: { line: 960 } }),

      // 封面信息块
      new Paragraph({
        children: [
          new TextRun({ text: "委  托  方：", font: FONT, size: 28, bold: false }),
          new TextRun({ text: "天津石油职业技术学院", font: FONT, size: 28 }),
        ],
        spacing: { line: 480, lineRule: "auto" }
      }),
      new Paragraph({
        children: [
          new TextRun({ text: "出 具 机 构：", font: FONT, size: 28 }),
          new TextRun({ text: "铜峰会计师事务所（特殊普通合伙）", font: FONT, size: 28 }),
        ],
        spacing: { line: 480, lineRule: "auto" }
      }),
      new Paragraph({
        children: [
          new TextRun({ text: "报 告 日 期：", font: FONT, size: 28 }),
          new TextRun({ text: "二○二六年  月  日", font: FONT, size: 28 }),
        ],
        spacing: { line: 480, lineRule: "auto" }
      }),

      new Paragraph({ children: [], spacing: { line: 480 } }),

      // 分隔线（底部边框段落）
      new Paragraph({
        children: [new TextRun({ text: "", font: FONT, size: 28 })],
        border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: "000000" } },
        spacing: { before: 480, after: 120 }
      }),

      new Paragraph({
        children: [
          new TextRun({ text: "地址：", font: FONT, size: 22, color: "595959" }),
          new TextRun({ text: "（事务所地址）", font: FONT, size: 22, color: "595959" }),
          new TextRun({ text: "          电话：（联系电话）", font: FONT, size: 22, color: "595959" }),
        ],
        alignment: AlignmentType.CENTER,
        spacing: { line: 360 }
      }),

      // ═══════ 第一部分：审计报告正文（新页） ═══════
      new Paragraph({ children: [], pageBreakBefore: true }),

      // 一、报告标题（正文）
      new Paragraph({
        children: [new TextRun({
          text: "关于大港工程职业技术学院划转天津石油职业技术学院",
          font: FONT_TITLE, size: 36, bold: true
        })],
        alignment: AlignmentType.CENTER,
        spacing: { line: 480 }
      }),
      new Paragraph({
        children: [new TextRun({
          text: "资产无抵押担保专项审计报告",
          font: FONT_TITLE, size: 36, bold: true
        })],
        alignment: AlignmentType.CENTER,
        spacing: { line: 480 }
      }),

      centerPara("铜峰审专字〔2026〕第××号", { color: "595959", size: 26 }),

      new Paragraph({ children: [], spacing: { line: 240 } }),

      // 二、收件方
      indentPara("天津石油职业技术学院："),

      new Paragraph({ children: [], spacing: { line: 240 } }),

      // ─── 一、业务背景 ───
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "一、基本情况", font: FONT_HEADING, size: 28, bold: true })]
      }),

      indentPara(
        "受天津石油职业技术学院（以下简称"委托方"或"贵校"）委托，铜峰会计师事务所（特殊普通合伙）（以下简称"本所"）对大港工程职业技术学院（以下简称"划转方"）无偿划转给天津石油职业技术学院的相关资产（以下简称"划转资产"）是否存在抵押、质押或对外提供担保情形进行专项审计。本次审计基准日为           （以实际确认日期为准）。"
      ),
      indentPara(
        "大港工程职业技术学院与天津石油职业技术学院同属同一控制主体下的高等学校，本次无偿划转系同一控制下的资产转移行为，不涉及商业对价。划转完成后，上述资产归属天津石油职业技术学院，相应权利义务亦随之转移。为保证委托方接收资产的无瑕疵性、合规性，委托方要求本所就划转资产是否设有抵押、质押及是否对外提供担保出具专项审计报告。"
      ),

      // ─── 二、管理层责任 ───
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "二、被审计单位管理层的责任", font: FONT_HEADING, size: 28, bold: true })]
      }),
      indentPara(
        "大港工程职业技术学院管理层（以下简称"被审计单位管理层"）负责按照国家法律法规和相关制度的规定，对划转资产进行确认、计量和记录，并保证所提供的会计凭证、账簿、报表及相关说明材料真实、合法、完整；负责确认划转资产上不存在任何抵押、质押、担保及其他权利负担，并就本次专项审计事项作出书面声明。"
      ),

      // ─── 三、注册会计师的责任 ───
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "三、注册会计师的责任", font: FONT_HEADING, size: 28, bold: true })]
      }),
      indentPara(
        "本所的责任是在执行审计工作的基础上，对划转资产是否存在抵押、质押或对外担保情形发表独立审计意见。本所按照中国注册会计师审计准则及相关职业道德规范，独立执行了专项审计工作，以充分适当的审计证据为基础，对出具本报告负责。"
      ),

      // ─── 四、审计范围与审计程序 ───
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "四、审计范围与主要审计程序", font: FONT_HEADING, size: 28, bold: true })]
      }),
      indentPara("（一）审计范围"),
      indentPara(
        "本次专项审计的范围为大港工程职业技术学院拟无偿划转给天津石油职业技术学院的全部资产，具体清单以双方签署的《资产划转协议》及附件为准。"
      ),
      indentPara("（二）主要审计程序"),
      indentPara("本所实施了以下主要审计程序："),
      new Paragraph({
        children: [new TextRun({
          text: "1．审阅被审计单位提供的资产清单、资产权属证明文件及相关账簿凭证，核实划转资产的账面记录情况；",
          font: FONT, size: 28
        })],
        indent: { left: 560, firstLine: 560 },
        spacing: { line: 480, lineRule: "auto" },
        alignment: AlignmentType.JUSTIFIED,
      }),
      new Paragraph({
        children: [new TextRun({
          text: "2．查阅被审计单位与银行等金融机构签订的借款合同、抵押合同及相关抵押登记记录，确认划转资产是否已被设定抵押或质押；",
          font: FONT, size: 28
        })],
        indent: { left: 560, firstLine: 560 },
        spacing: { line: 480, lineRule: "auto" },
        alignment: AlignmentType.JUSTIFIED,
      }),
      new Paragraph({
        children: [new TextRun({
          text: "3．通过不动产登记中心等权威机构查询不动产抵押登记情况，核实相关房产、土地使用权是否存在抵押登记；",
          font: FONT, size: 28
        })],
        indent: { left: 560, firstLine: 560 },
        spacing: { line: 480, lineRule: "auto" },
        alignment: AlignmentType.JUSTIFIED,
      }),
      new Paragraph({
        children: [new TextRun({
          text: "4．查阅被审计单位对外担保台账及相关协议，核实划转资产是否已被用于提供对外担保；",
          font: FONT, size: 28
        })],
        indent: { left: 560, firstLine: 560 },
        spacing: { line: 480, lineRule: "auto" },
        alignment: AlignmentType.JUSTIFIED,
      }),
      new Paragraph({
        children: [new TextRun({
          text: "5．获取被审计单位管理层关于划转资产上不存在任何抵押、质押、担保的书面声明；",
          font: FONT, size: 28
        })],
        indent: { left: 560, firstLine: 560 },
        spacing: { line: 480, lineRule: "auto" },
        alignment: AlignmentType.JUSTIFIED,
      }),
      new Paragraph({
        children: [new TextRun({
          text: "6．对以上程序所获取的审计证据进行综合评价，形成审计结论。",
          font: FONT, size: 28
        })],
        indent: { left: 560, firstLine: 560 },
        spacing: { line: 480, lineRule: "auto" },
        alignment: AlignmentType.JUSTIFIED,
      }),

      // ─── 五、审计结论 ───
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "五、审计结论", font: FONT_HEADING, size: 28, bold: true })]
      }),
      indentPara(
        "经本所专项审计，截至审计基准日，大港工程职业技术学院拟无偿划转给天津石油职业技术学院的划转资产，均不存在以下情形："
      ),
      new Paragraph({
        children: [new TextRun({
          text: "（一）划转资产不存在抵押或质押担保，均未被设定抵押权或质权；",
          font: FONT, size: 28
        })],
        indent: { left: 560, firstLine: 560 },
        spacing: { line: 480, lineRule: "auto" },
        alignment: AlignmentType.JUSTIFIED,
      }),
      new Paragraph({
        children: [new TextRun({
          text: "（二）划转资产未被用于对外提供担保（包括但不限于保证担保、抵押担保、质押担保）；",
          font: FONT, size: 28
        })],
        indent: { left: 560, firstLine: 560 },
        spacing: { line: 480, lineRule: "auto" },
        alignment: AlignmentType.JUSTIFIED,
      }),
      new Paragraph({
        children: [new TextRun({
          text: "（三）划转资产不存在其他权利负担或法律上的权利瑕疵，不影响划转的合法有效性。",
          font: FONT, size: 28
        })],
        indent: { left: 560, firstLine: 560 },
        spacing: { line: 480, lineRule: "auto" },
        alignment: AlignmentType.JUSTIFIED,
      }),
      indentPara(
        "综上，本所认为，上述划转资产权属明确、无抵押、无质押、无对外担保，可无风险接收。"
      ),

      // ─── 六、划转资产明细 ───
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "六、划转资产明细", font: FONT_HEADING, size: 28, bold: true })]
      }),
      indentPara(
        "划转资产明细情况如下表所示（具体金额以双方签认的资产清单及最终审计确认数为准）："
      ),

      new Paragraph({ children: [], spacing: { line: 240 } }),

      new Table({
        width: { size: 10360, type: WidthType.DXA },
        columnWidths: [600, 2200, 1600, 2000, 2000, 1960],
        rows: assetTableRows
      }),

      new Paragraph({ children: [], spacing: { line: 240 } }),
      centerPara("（注：资产明细清单以附件形式附于报告后）", { size: 24, color: "595959" }),

      // ─── 七、特别说明 ───
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: "七、特别说明", font: FONT_HEADING, size: 28, bold: true })]
      }),
      indentPara(
        "（一）本报告仅就划转资产是否存在抵押、质押及对外担保情形发表意见，不涉及资产价值评估及其他审计事项。"
      ),
      indentPara(
        "（二）本报告依赖于被审计单位管理层提供的资料及书面声明，如被审计单位提供的资料存在不真实、不完整情形，本所将不承担相应法律责任。"
      ),
      indentPara(
        "（三）本报告系专项出具，仅供天津石油职业技术学院在本次资产无偿划转事项中使用，不得用于其他目的。如将本报告用于其他目的，本所概不负责。"
      ),
      indentPara(
        "（四）本报告自出具之日起一年内有效。如划转资产情况发生变化，本报告自动失效。"
      ),

      // ─── 签章区 ───
      new Paragraph({ children: [], spacing: { line: 480 } }),

      new Paragraph({
        children: [new TextRun({ text: "铜峰会计师事务所（特殊普通合伙）", font: FONT, size: 28, bold: true })],
        alignment: AlignmentType.RIGHT,
        spacing: { line: 480 }
      }),
      new Paragraph({
        children: [new TextRun({ text: "（盖章）", font: FONT, size: 28 })],
        alignment: AlignmentType.RIGHT,
        spacing: { line: 480 }
      }),
      new Paragraph({
        children: [new TextRun({ text: "注册会计师（签字）：                ", font: FONT, size: 28 })],
        alignment: AlignmentType.RIGHT,
        spacing: { line: 480 }
      }),
      new Paragraph({
        children: [new TextRun({ text: "注册会计师（签字）：                ", font: FONT, size: 28 })],
        alignment: AlignmentType.RIGHT,
        spacing: { line: 480 }
      }),
      new Paragraph({
        children: [new TextRun({ text: "二○二六年   月   日", font: FONT, size: 28 })],
        alignment: AlignmentType.RIGHT,
        spacing: { line: 480 }
      }),

      // ═══════ 附件页 ═══════
      new Paragraph({ children: [], pageBreakBefore: true }),

      new Paragraph({
        children: [new TextRun({ text: "附件：被审计单位管理层声明书", font: FONT_HEADING, size: 32, bold: true })],
        alignment: AlignmentType.CENTER,
        spacing: { line: 480 }
      }),
      new Paragraph({ children: [], spacing: { line: 240 } }),
      indentPara("铜峰会计师事务所（特殊普通合伙）："),
      new Paragraph({ children: [], spacing: { line: 240 } }),
      indentPara(
        "本单位就大港工程职业技术学院无偿划转给天津石油职业技术学院资产事项，特向贵所郑重声明如下："
      ),
      indentPara(
        "一、本次划转资产清单所列各项资产，均系本单位合法拥有，权属清晰，无产权争议。"
      ),
      indentPara(
        "二、上述资产截至声明日，均未向任何单位或个人设定抵押权或质权，不存在任何形式的抵押或质押担保。"
      ),
      indentPara(
        "三、上述资产截至声明日，均未被用于对外提供任何形式的担保（包括但不限于保证担保、抵押担保、质押担保等）。"
      ),
      indentPara(
        "四、上述资产不存在任何查封、扣押、冻结等司法强制措施，不存在任何涉诉纠纷。"
      ),
      indentPara(
        "五、本单位已向贵所提供了与本次专项审计相关的全部真实、完整、合法的资料，所有书面资料及口头解释均不存在重大错误、遗漏或误导性陈述。"
      ),
      indentPara(
        "六、如上述声明内容存在不实，由此产生的一切法律责任由本单位承担，与贵所无关。"
      ),
      new Paragraph({ children: [], spacing: { line: 480 } }),
      new Paragraph({
        children: [new TextRun({ text: "大港工程职业技术学院（公章）", font: FONT, size: 28 })],
        alignment: AlignmentType.RIGHT,
        spacing: { line: 480 }
      }),
      new Paragraph({
        children: [new TextRun({ text: "法定代表人（签字）：                ", font: FONT, size: 28 })],
        alignment: AlignmentType.RIGHT,
        spacing: { line: 480 }
      }),
      new Paragraph({
        children: [new TextRun({ text: "二○二六年   月   日", font: FONT, size: 28 })],
        alignment: AlignmentType.RIGHT,
        spacing: { line: 480 }
      }),
    ]
  }]
});

// ─── 输出文件 ───
const outPath = "D:\\Users\\12844\\Desktop\\天津石油职业技术学院\\无抵押担保专项审计报告.docx";
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outPath, buffer);
  console.log("报告已生成：" + outPath);
}).catch(e => {
  console.error("生成失败：", e);
  process.exit(1);
});
