const fs = require('fs');
const docx = require('docx');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
        BorderStyle, WidthType, ShadingType, Table, TableRow, TableCell,
        PageBreak } = docx;

const children = [];

// 标题页
children.push(new Paragraph({
  children: [new TextRun({ text: '国家税务总局', size: 28 })],
  heading: HeadingLevel.HEADING_2,
  alignment: AlignmentType.CENTER,
  spacing: { after: 200 }
}));

children.push(new Paragraph({
  children: [new TextRun({ text: '关于印发《特别纳税调整实施办法（试行）》的通知', bold: true, size: 32 })],
  heading: HeadingLevel.HEADING_1,
  alignment: AlignmentType.CENTER,
  spacing: { after: 200 }
}));

children.push(new Paragraph({
  children: [new TextRun({ text: '国税发〔2009〕2号', bold: true, size: 28 })],
  alignment: AlignmentType.CENTER,
  spacing: { after: 400 }
}));

// 有效性说明表格
const noteBorder = { style: BorderStyle.SINGLE, size: 1, color: 'E74C3C' };
children.push(new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [9360],
  rows: [
    new TableRow({
      children: [
        new TableCell({
          borders: { top: noteBorder, bottom: noteBorder, left: noteBorder, right: noteBorder },
          shading: { fill: 'FEF9E7', type: ShadingType.CLEAR },
          margins: { top: 100, bottom: 100, left: 150, right: 150 },
          children: [
            new Paragraph({ children: [new TextRun({ text: '⚠️ 文件有效性说明', bold: true, size: 22, color: 'E74C3C' })], spacing: { after: 120 } }),
            new Paragraph({ children: [new TextRun({ text: '• 发布日期：2009年1月8日', size: 20 })], spacing: { after: 60 } }),
            new Paragraph({ children: [new TextRun({ text: '• 生效日期：2008年1月1日', size: 20 })], spacing: { after: 60 } }),
            new Paragraph({ children: [new TextRun({ text: '• 有效性：大部分条款有效，仅第七十八条被废止', size: 20 })], spacing: { after: 60 } }),
            new Paragraph({ children: [new TextRun({ text: '• 废止条款：第七十八条（自2023年5月26日起废止）', size: 20, strike: true, color: 'E74C3C' })], spacing: { after: 60 } }),
            new Paragraph({ children: [new TextRun({ text: '• 废止依据：《国家税务总局关于公布全文和部分条款失效废止的税务规范性文件目录的公告》（2023年）', size: 20 })] }),
          ]
        })
      ]
    })
  ]
}));

children.push(new Paragraph({ spacing: { after: 400 } }));

// 通知正文
children.push(new Paragraph({
  children: [new TextRun({ text: '各省、自治区、直辖市和计划单列市国家税务局、地方税务局：', size: 22 })],
  spacing: { before: 200, after: 200 }
}));

children.push(new Paragraph({
  children: [new TextRun({ text: '    为贯彻落实《中华人民共和国企业所得税法》及其实施条例，规范和加强特别纳税调整管理，国家税务总局制定了《特别纳税调整实施办法（试行）》，现印发给你们，请遵照执行。', size: 22 })],
  spacing: { after: 200 }
}));

children.push(new Paragraph({
  children: [new TextRun({ text: '税务总局', size: 22 })],
  alignment: AlignmentType.RIGHT,
  spacing: { after: 100 }
}));

children.push(new Paragraph({
  children: [new TextRun({ text: '二○○九年一月八日', size: 22 })],
  alignment: AlignmentType.RIGHT,
  spacing: { after: 400 }
}));

children.push(new Paragraph({ children: [new PageBreak()] }));

// 章节内容
const chapters = [
  {
    title: '第一章 总则',
    articles: [
      { num: '第一条', content: '为了规范特别纳税调整管理，根据《中华人民共和国企业所得税法》（以下简称所得税法）、《中华人民共和国企业所得税法实施条例》（以下简称所得税法实施条例）、《中华人民共和国税收征收管理法》（以下简称征管法）、《中华人民共和国税收征收管理法实施细则》（以下简称征管法实施细则）以及我国政府与有关国家（地区）政府签署的避免双重征税协定（安排）（以下简称税收协定）的有关规定，制定本办法。' },
      { num: '第二条', content: '本办法适用于税务机关对企业的转让定价、预约定价安排、成本分摊协议、受控外国企业、资本弱化以及一般反避税等特别纳税调整事项的管理。' },
      { num: '第三条', content: '转让定价管理是指税务机关按照所得税法第六章和征管法第三十六条的有关规定，对企业与其关联方之间的业务往来（以下简称关联交易）是否符合独立交易原则进行审核评估和调查调整等工作的总称。' },
      { num: '第四条', content: '预约定价安排管理是指税务机关按照所得税法第四十二条和征管法实施细则第五十三条的规定，对企业提出的未来年度关联交易的定价原则和计算方法进行审核评估，并与企业协商达成预约定价安排等工作的总称。' },
      { num: '第五条', content: '成本分摊协议管理是指税务机关按照所得税法第四十一条第二款的规定，对企业与其关联方签署的成本分摊协议是否符合独立交易原则进行审核评估和调查调整等工作的总称。' },
      { num: '第六条', content: '受控外国企业管理是指税务机关按照所得税法第四十五条的规定，对受控外国企业不作利润分配或减少分配进行审核评估和调查，并对归属于中国居民企业所得进行调整等工作的总称。' },
      { num: '第七条', content: '资本弱化管理是指税务机关按照所得税法第四十六条的规定，对企业接受关联方债权性投资与企业接受的权益性投资的比例是否符合规定比例或独立交易原则进行审核评估和调查调整等工作的总称。' },
      { num: '第八条', content: '一般反避税管理是指税务机关按照所得税法第四十七条的规定，对企业实施其他不具有合理商业目的的安排而减少其应纳税收入或所得额进行审核评估和调查调整等工作的总称。' },
    ]
  },
  {
    title: '第二章 关联申报',
    articles: [
      { num: '第九条', content: '所得税法实施条例第一百零九条及征管法实施细则第五十一条所称关联关系，主要是指企业与其他企业、组织或个人具有下列之一关系：（一）一方直接或间接持有另一方的股份总和达到25%以上，或者双方直接或间接同为第三方所持有的股份达到25%以上；（二）一方与另一方之间借贷资金占一方实收资本50%以上，或者一方借贷资金总额的10%以上是由另一方担保；（三）一方半数以上的高级管理人员或至少一名可以控制董事会的董事会高级成员是由另一方委派；（四）一方半数以上的高级管理人员同时担任另一方的高级管理人员，或者一方至少一名可以控制董事会的董事会高级成员同时担任另一方的董事会高级成员；（五）一方的生产经营活动必须由另一方提供的工业产权、专有技术等特许权才能正常进行；（六）一方的购买或销售活动主要由另一方控制；（七）一方接受或提供劳务主要由另一方控制；（八）一方对另一方的生产经营、交易具有实质控制，或者双方在利益上具有相关联的其他关系。' },
      { num: '第十条', content: '关联交易主要包括以下类型：（一）有形资产的购销、转让和使用；（二）无形资产的转让和使用；（三）融通资金；（四）提供劳务。' },
      { num: '第十一条', content: '实行查账征收的居民企业和在中国境内设立机构、场所并据实申报缴纳企业所得税的非居民企业向税务机关报送年度企业所得税纳税申报表时，应附送《中华人民共和国企业年度关联业务往来报告表》。' },
      { num: '第十二条', content: '企业按规定期限报送本办法第十一条规定的报告表确有困难，需要延期的，应按征管法及其实施细则的有关规定办理。' },
    ]
  },
  {
    title: '第三章 同期资料管理',
    articles: [
      { num: '第十三条', content: '企业应根据所得税法实施条例第一百一十四条的规定，按纳税年度准备、保存、并按税务机关要求提供其关联交易的同期资料。' },
      { num: '第十四条', content: '同期资料主要包括以下内容：（一）组织结构；（二）生产经营情况；（三）关联交易情况；（四）可比性分析；（五）转让定价方法的选择和使用。' },
      { num: '第十五条', content: '属于下列情形之一的企业，可免于准备同期资料：（一）年度发生的关联购销金额在2亿元人民币以下且其他关联交易金额在4000万元人民币以下；（二）关联交易属于执行预约定价安排所涉及的范围；（三）外资股份低于50%且仅与境内关联方发生关联交易。' },
      { num: '第十六条', content: '除本办法第七章另有规定外，企业应在关联交易发生年度的次年5月31日之前准备完毕该年度同期资料，并自税务机关要求之日起20日内提供。' },
      { num: '第十七条', content: '企业按照税务机关要求提供的同期资料，须加盖公章，并由法定代表人或法定代表人授权的代表签字或盖章。' },
      { num: '第十八条', content: '企业因合并、分立等原因变更或注销税务登记的，应由合并、分立后的企业保存同期资料。' },
      { num: '第十九条', content: '同期资料应使用中文。如原始资料为外文的，应附送中文副本。' },
      { num: '第二十条', content: '同期资料应自企业关联交易发生年度的次年6月1日起保存10年。' },
    ]
  },
  {
    title: '第四章 转让定价方法',
    articles: [
      { num: '第二十一条', content: '企业发生关联交易以及税务机关审核、评估关联交易均应遵循独立交易原则，选用合理的转让定价方法。转让定价方法包括可比非受控价格法、再销售价格法、成本加成法、交易净利润法、利润分割法和其他符合独立交易原则的方法。' },
      { num: '第二十二条', content: '选用合理的转让定价方法应进行可比性分析。可比性分析因素主要包括：交易资产或劳务特性、交易各方功能和风险、合同条款、经济环境、经营策略。' },
      { num: '第二十三条', content: '可比非受控价格法以非关联方之间进行的与关联交易相同或类似业务活动所收取的价格作为关联交易的公平成交价格。可比非受控价格法可以适用于所有类型的关联交易。' },
      { num: '第二十四条', content: '再销售价格法以关联方购进商品再销售给非关联方的价格减去可比非关联交易毛利后的金额作为关联方购进商品的公平成交价格。再销售价格法通常适用于再销售者未对商品进行实质性增值加工的简单加工或单纯购销业务。' },
      { num: '第二十五条', content: '成本加成法以关联交易发生的合理成本加上可比非关联交易毛利作为关联交易的公平成交价格。成本加成法通常适用于有形资产的购销、转让和使用，劳务提供或资金融通的关联交易。' },
      { num: '第二十六条', content: '交易净利润法以可比非关联交易的利润率指标确定关联交易的净利润。交易净利润法通常适用于有形资产的购销、转让和使用，无形资产的转让和使用以及劳务提供等关联交易。' },
      { num: '第二十七条', content: '利润分割法根据企业与其关联方对关联交易合并利润的贡献计算各自应该分配的利润额。利润分割法分为一般利润分割法和剩余利润分割法。' },
    ]
  },
  {
    title: '第五章 转让定价调查及调整',
    articles: [
      { num: '第二十八条', content: '税务机关有权依据税收征管法及其实施细则有关税务检查的规定，确定调查企业，进行转让定价调查、调整。被调查企业必须据实报告其关联交易情况，并提供相关资料，不得拒绝或隐瞒。' },
      { num: '第二十九条', content: '转让定价调查应重点选择以下企业：（一）关联交易数额较大或类型较多的企业；（二）长期亏损、微利或跳跃性盈利的企业；（三）低于同行业利润水平的企业；（四）利润水平与其所承担的功能风险明显不相匹配的企业；（五）与避税港关联方发生业务往来的企业；（六）未按规定进行关联申报或准备同期资料的企业；（七）其他明显违背独立交易原则的企业。' },
      { num: '第三十条', content: '实际税负相同的境内关联方之间的交易，只要该交易没有直接或间接导致国家总体税收收入的减少，原则上不做转让定价调查、调整。' },
      { num: '第四十一条', content: '税务机关采用四分位法分析、评估企业利润水平时，企业利润水平低于可比企业利润率区间中位值的，原则上应按照不低于中位值进行调整。' },
      { num: '第四十二条', content: '经调查，企业关联交易符合独立交易原则的，税务机关应做出转让定价调查结论，并向企业送达《特别纳税调查结论通知书》。' },
      { num: '第四十三条', content: '经调查，企业关联交易不符合独立交易原则而减少其应纳税收入或者所得额的，税务机关应按程序实施转让定价纳税调整，包括拟定调整方案、与企业协商谈判、确定最终调整方案并送达《特别纳税调查调整通知书》。' },
    ]
  },
  {
    title: '第六章 预约定价安排管理',
    articles: [
      { num: '第四十六条', content: '企业可以依据所得税法第四十二条、所得税法实施条例第一百一十三条及征管法实施细则第五十三条的规定，与税务机关就企业未来年度关联交易的定价原则和计算方法达成预约定价安排。预约定价安排的谈签与执行通常经过预备会谈、正式申请、审核评估、磋商、签订安排和监控执行6个阶段。' },
      { num: '第四十七条', content: '预约定价安排应由设区的市、自治州以上的税务机关受理。' },
      { num: '第四十八条', content: '预约定价安排一般适用于同时满足以下条件的企业：（一）年度发生的关联交易金额在4000万元人民币以上；（二）依法履行关联申报义务；（三）按规定准备、保存和提供同期资料。' },
      { num: '第四十九条', content: '预约定价安排适用于自企业提交正式书面申请年度的次年起3至5个连续年度的关联交易。' },
    ]
  },
  {
    title: '第七章 成本分摊协议管理',
    articles: [
      { num: '第六十四条', content: '根据所得税法第四十一条第二款及所得税法实施条例第一百一十二条的规定，企业与其关联方签署成本分摊协议，共同开发、受让无形资产，或者共同提供、接受劳务，应符合本章规定。' },
      { num: '第七十五条', content: '企业与其关联方签署成本分摊协议，有下列情形之一的，其自行分摊的成本不得税前扣除：（一）不具有合理商业目的和经济实质；（二）不符合独立交易原则；（三）没有遵循成本与收益配比原则；（四）未按规定备案或准备、保存和提供同期资料；（五）自签署成本分摊协议之日起经营期限少于20年。' },
    ]
  },
  {
    title: '第八章 受控外国企业管理',
    articles: [
      { num: '第七十六条', content: '受控外国企业是指根据所得税法第四十五条的规定，由居民企业，或者由居民企业和居民个人控制的设立在实际税负低于所得税法第四条第一款规定税率水平50%的国家（地区），并非出于合理经营需要对利润不作分配或减少分配的外国企业。' },
      { num: '第七十七条', content: '本办法第七十六条所称控制，是指在股份、资金、经营、购销等方面构成实质控制。其中，股份控制是指由中国居民股东在纳税年度任何一天单层直接或多层间接单一持有外国企业10%以上有表决权股份，且共同持有该外国企业50%以上股份。' },
      { num: '【已废止】第七十八条', content: '中国居民企业股东应在年度企业所得税纳税申报时提供对外投资信息，附送《对外投资情况表》。【本条自2023年5月26日起废止】', revoked: true },
      { num: '第七十九条', content: '税务机关应汇总、审核中国居民企业股东申报的对外投资信息，向受控外国企业的中国居民企业股东送达《受控外国企业中国居民股东确认通知书》。' },
      { num: '第八十四条', content: '中国居民企业股东能够提供资料证明其控制的外国企业满足以下条件之一的，可免于将外国企业不作分配或减少分配的利润视同股息分配额：（一）设立在国家税务总局指定的非低税率国家（地区）；（二）主要取得积极经营活动所得；（三）年度利润总额低于500万元人民币。' },
    ]
  },
  {
    title: '第九章 资本弱化管理',
    articles: [
      { num: '第八十五条', content: '所得税法第四十六条所称不得在计算应纳税所得额时扣除的利息支出应按以下公式计算：不得扣除利息支出＝年度实际支付的全部关联方利息×（1－标准比例/关联债资比例）。' },
      { num: '第八十六条', content: '关联债资比例＝年度各月平均关联债权投资之和/年度各月平均权益投资之和。' },
      { num: '第九十条', content: '企业未按规定准备、保存和提供同期资料证明关联债权投资金额、利率、期限、融资条件以及债资比例等符合独立交易原则的，其超过标准比例的关联方利息支出，不得在计算应纳税所得额时扣除。' },
    ]
  },
  {
    title: '第十章 一般反避税管理',
    articles: [
      { num: '第九十二条', content: '税务机关可依据所得税法第四十七条及所得税法实施条例第一百二十条的规定对存在以下避税安排的企业，启动一般反避税调查：（一）滥用税收优惠；（二）滥用税收协定；（三）滥用公司组织形式；（四）利用避税港避税；（五）其他不具有合理商业目的的安排。' },
      { num: '第九十三条', content: '税务机关应按照实质重于形式的原则审核企业是否存在避税安排，并综合考虑安排的以下内容：（一）安排的形式和实质；（二）安排订立的时间和执行期间；（三）安排实现的方式；（四）安排各个步骤或组成部分之间的联系；（五）安排涉及各方财务状况的变化；（六）安排的税收结果。' },
      { num: '第九十七条', content: '一般反避税调查及调整须层报国家税务总局批准。' },
    ]
  },
  {
    title: '第十一章 相应调整及国际磋商',
    articles: [
      { num: '第九十八条', content: '关联交易一方被实施转让定价调查调整的，应允许另一方做相应调整，以消除双重征税。' },
      { num: '第一百条', content: '企业应自企业或其关联方收到转让定价调整通知书之日起三年内提出相应调整的申请，超过三年的，税务机关不予受理。' },
    ]
  },
  {
    title: '第十二章 法律责任',
    articles: [
      { num: '第一百零五条', content: '企业未按照本办法的规定向税务机关报送企业年度关联业务往来报告表，或者未保存同期资料或其他相关资料的，依照征管法第六十条和第六十二条的规定处理。' },
      { num: '第一百零七条', content: '税务机关根据所得税法及其实施条例的规定，对企业做出特别纳税调整的，应对2008年1月1日以后发生交易补征的企业所得税税款，按日加收利息。（一）计息期间自税款所属纳税年度的次年6月1日起至补缴税款入库之日止；（二）利息率按照税款所属纳税年度12月31日实行的中国人民银行人民币贷款基准利率加5个百分点计算；（三）企业按照规定提供同期资料的可只按基准利率计算加收利息。' },
    ]
  },
  {
    title: '第十三章 附则',
    articles: [
      { num: '第一百一十条', content: '税务机关对转让定价管理和预约定价安排管理以外的其他特别纳税调整事项实施的调查调整程序可参照适用本办法第五章的有关规定。' },
      { num: '第一百一十一条', content: '各级国家税务局和地方税务局对企业实施特别纳税调查调整要加强联系，可根据需要组成联合调查组进行调查。' },
      { num: '第一百一十七条', content: '本办法由国家税务总局负责解释和修订。' },
      { num: '第一百一十八条', content: '本办法自2008年1月1日起施行。在本办法发布前实施的有关规定与本办法不一致的，以本办法为准。' },
    ]
  },
];

for (const ch of chapters) {
  children.push(new Paragraph({
    children: [new TextRun({ text: ch.title, bold: true, size: 28 })],
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 300, after: 200 }
  }));
  for (const art of ch.articles) {
    const isRevoked = art.revoked === true;
    children.push(new Paragraph({
      children: [
        new TextRun({ text: art.num + ' ', bold: true, size: 22, strike: isRevoked, color: isRevoked ? 'FF0000' : undefined }),
        new TextRun({ text: art.content, size: 22, strike: isRevoked, color: isRevoked ? 'FF0000' : undefined })
      ],
      spacing: { before: 80, after: 80 },
      indent: { left: 360 }
    }));
    if (isRevoked) {
      children.push(new Paragraph({
        children: [new TextRun({ text: '【本条已废止】根据2023年公告，本条自2023年5月26日起废止，不再作为执法依据。', bold: true, size: 20, color: 'FF0000' })],
        spacing: { before: 40, after: 120 },
        indent: { left: 720 }
      }));
    }
  }
}

// 生成文档
const doc = new Document({
  styles: {
    default: { document: { run: { font: '宋体', size: 24 } } },
    paragraphStyles: [
      {
        id: 'Heading1',
        name: 'Heading 1',
        basedOn: 'Normal',
        next: 'Normal',
        run: { size: 32, bold: true, font: '黑体' },
        paragraph: { spacing: { before: 240, after: 240 }, alignment: AlignmentType.CENTER }
      },
      {
        id: 'Heading2',
        name: 'Heading 2',
        basedOn: 'Normal',
        next: 'Normal',
        run: { size: 28, bold: true, font: '黑体' },
        paragraph: { spacing: { before: 200, after: 200 } }
      },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    children: children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('国税发〔2009〕2号_特别纳税调整实施办法（试行）.docx', buffer);
  console.log('Word文档生成成功！');
  console.log('文件路径：' + process.cwd() + '/国税发〔2009〕2号_特别纳税调整实施办法（试行）.docx');
}).catch(err => {
  console.error('生成失败:', err.message);
  process.exit(1);
});
