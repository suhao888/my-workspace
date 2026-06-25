const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
        BorderStyle, WidthType, ShadingType, PageNumber, PageBreak } = require('docx');
const fs = require('fs');

// 边框设置
const border = { style: BorderStyle.SINGLE, size: 1, color: "999999" };
const borders = { top: border, bottom: border, left: border, right: border };

// 表头样式
const headerCell = (text, width) => new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: "1F4E79", type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text, bold: true, color: "FFFFFF", font: "Arial", size: 22 })]
    })]
});

// 普通单元格
const dataCell = (text, width, align = AlignmentType.LEFT) => new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
        alignment: align,
        children: [new TextRun({ text: String(text), font: "Arial", size: 20 })]
    })]
});

// 高亮单元格
const highlightCell = (text, width, fillColor = "FFF2CC") => new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: fillColor, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: String(text), bold: true, font: "Arial", size: 20 })]
    })]
});

const doc = new Document({
    styles: {
        default: { document: { run: { font: "Arial", size: 21 } } },
        paragraphStyles: [
            { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 36, bold: true, font: "Arial", color: "1F4E79" },
                paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
            { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 28, bold: true, font: "Arial", color: "2E75B6" },
                paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 } },
            { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
                run: { size: 24, bold: true, font: "Arial", color: "404040" },
                paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 } },
        ]
    },
    numbering: {
        config: [
            { reference: "bullets",
              levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
                style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
            { reference: "numbers",
              levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
                style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
        ]
    },
    sections: [{
        properties: {
            page: {
                size: { width: 11906, height: 16838 },
                margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
            }
        },
        headers: {
            default: new Header({
                children: [new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [new TextRun({ text: "车辆服务中心2025年度司机差旅费补助审计报告", font: "Arial", size: 18, color: "666666" })]
                })]
            })
        },
        footers: {
            default: new Footer({
                children: [new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                        new TextRun({ text: "第 ", font: "Arial", size: 18, color: "666666" }),
                        new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: "666666" }),
                        new TextRun({ text: " 页", font: "Arial", size: 18, color: "666666" })
                    ]
                })]
            })
        },
        children: [
            // 标题
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 720, after: 480 },
                children: [new TextRun({ text: "车辆服务中心2025年度", font: "Arial", size: 52, bold: true, color: "1F4E79" })]
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 480 },
                children: [new TextRun({ text: "司机差旅费补助审计报告", font: "Arial", size: 52, bold: true, color: "1F4E79" })]
            }),
            new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 720 },
                children: [new TextRun({ text: "（五大风险点专项分析）", font: "Arial", size: 28, color: "666666" })]
            }),

            // 基本信息表
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [2000, 7026],
                rows: [
                    new TableRow({ children: [
                        new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, shading: { fill: "E7E6E6", type: ShadingType.CLEAR },
                            children: [new Paragraph({ children: [new TextRun({ text: "审计期间", bold: true, font: "Arial", size: 20 })] })] }),
                        new TableCell({ borders, width: { size: 7026, type: WidthType.DXA },
                            children: [new Paragraph({ children: [new TextRun({ text: "2025年度（含2026年2月追溯核查）", font: "Arial", size: 20 })] })] })
                    ]}),
                    new TableRow({ children: [
                        new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, shading: { fill: "E7E6E6", type: ShadingType.CLEAR },
                            children: [new Paragraph({ children: [new TextRun({ text: "审计范围", bold: true, font: "Arial", size: 20 })] })] }),
                        new TableCell({ borders, width: { size: 7026, type: WidthType.DXA },
                            children: [new Paragraph({ children: [new TextRun({ text: "车辆服务中心全部94名司机", font: "Arial", size: 20 })] })] })
                    ]}),
                    new TableRow({ children: [
                        new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, shading: { fill: "E7E6E6", type: ShadingType.CLEAR },
                            children: [new Paragraph({ children: [new TextRun({ text: "补助总额", bold: true, font: "Arial", size: 20 })] })] }),
                        new TableCell({ borders, width: { size: 7026, type: WidthType.DXA },
                            children: [new Paragraph({ children: [new TextRun({ text: "146.9万元（全年汇总）", font: "Arial", size: 20 })] })] })
                    ]}),
                    new TableRow({ children: [
                        new TableCell({ borders, width: { size: 2000, type: WidthType.DXA }, shading: { fill: "E7E6E6", type: ShadingType.CLEAR },
                            children: [new Paragraph({ children: [new TextRun({ text: "审计日期", bold: true, font: "Arial", size: 20 })] })] }),
                        new TableCell({ borders, width: { size: 7026, type: WidthType.DXA },
                            children: [new Paragraph({ children: [new TextRun({ text: "2026年5月15日", font: "Arial", size: 20 })] })] })
                    ]}),
                ]
            }),

            new Paragraph({ spacing: { before: 600, after: 200 }, children: [] }),

            // 风险总览
            new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("一、审计风险总览")] }),
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [1500, 5000, 1500, 1026],
                rows: [
                    new TableRow({ children: [
                        headerCell("序号", 1500),
                        headerCell("风险点", 5000),
                        headerCell("风险等级", 1500),
                        headerCell("涉及金额估算", 1026)
                    ]}),
                    new TableRow({ children: [
                        dataCell("1", 1500, AlignmentType.CENTER),
                        dataCell("制度沿革合规性", 5000),
                        dataCell("中风险", 1500, AlignmentType.CENTER),
                        dataCell("-", 1026, AlignmentType.CENTER)
                    ]}),
                    new TableRow({ children: [
                        highlightCell("2", 1500, "FFD966"),
                        highlightCell("保底公里制度套利", 5000, "FFD966"),
                        highlightCell("高风险", 1500, "FFD966"),
                        highlightCell("虚增87,336公里", 1026, "FFD966")
                    ]}),
                    new TableRow({ children: [
                        highlightCell("3", 1500, "FF6B6B"),
                        highlightCell("出车天数超月份天数", 5000, "FF6B6B"),
                        highlightCell("高风险", 1500, "FF6B6B"),
                        highlightCell("直接矛盾", 1026, "FF6B6B")
                    ]}),
                    new TableRow({ children: [
                        highlightCell("4", 1500, "FF6B6B"),
                        highlightCell("个别司机里程异常高", 5000, "FF6B6B"),
                        highlightCell("高风险", 1500, "FF6B6B"),
                        highlightCell("需交叉验证", 1026, "FF6B6B")
                    ]}),
                    new TableRow({ children: [
                        dataCell("5", 1500, AlignmentType.CENTER),
                        dataCell("周末补助超限", 5000),
                        dataCell("中高风险", 1500, AlignmentType.CENTER),
                        dataCell("14,690元/月", 1026, AlignmentType.CENTER)
                    ]}),
                ]
            }),

            new Paragraph({ children: [new PageBreak()] }),

            // 风险一
            new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("二、风险一：制度沿革合规性（中风险）")] }),

            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("2.1 审计发现")] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "现行细则附录C引用华北概预（2013）46号文，距今已超过10年", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "缺少上会审批记录证据，无法确认制度的有效性和时效性", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "制度依据可能已过时，需重新确认适用性", font: "Arial", size: 21 })] }),

            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("2.2 审计建议")] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "调取2013年原文件，确认其是否仍为现行有效制度依据", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "补充上会审批记录，完善制度修订的正式程序", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "如制度已更新，需更新细则引用条款", font: "Arial", size: 21 })] }),

            new Paragraph({ spacing: { before: 400, after: 200 }, children: [] }),

            // 风险二
            new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("三、风险二：保底公里制度套利（高风险）")] }),

            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("3.1 制度规定")] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "月包：3000公里/月（折合100公里/天）", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "日包：144公里/天（司机每日最低出车保障公里数）", font: "Arial", size: 21 })] }),

            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("3.2 审计发现")] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "保底制度导致低出车司机与高出车司机获得相同补助", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "8月全队计提公里：435,982公里 vs 实际行驶公里：348,646公里", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "虚增公里数：87,336公里（虚增比例：+25%）", font: "Arial", size: 21, bold: true, color: "C00000" })] }),

            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("3.3 典型案例")] }),
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [1500, 1500, 2000, 1500, 1500, 1026],
                rows: [
                    new TableRow({ children: [
                        headerCell("司机", 1500),
                        headerCell("出车天数", 1500),
                        headerCell("实际公里", 2000),
                        headerCell("计提公里", 1500),
                        headerCell("补助金额", 1500),
                        headerCell("等效补助/公里", 1026)
                    ]}),
                    new TableRow({ children: [
                        dataCell("虞岩", 1500),
                        dataCell("1天", 1500),
                        dataCell("2公里", 2000),
                        highlightCell("3,000公里", 1500, "FF6B6B"),
                        dataCell("600元", 1500),
                        dataCell("300元/km", 1026)
                    ]}),
                ]
            }),

            new Paragraph({ spacing: { before: 200, after: 100 }, children: [] }),
            new Paragraph({ children: [new TextRun({ text: "案例分析：虞岩仅出车1天、行驶2公里，却按月包3000公里计提，补助600元，等效每实际公里补助300元，显著高于正常标准，存在明显套利行为。", font: "Arial", size: 21, italics: true, color: "666666" })] }),

            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("3.4 审计建议")] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "重新评估月包制度的合理性，考虑按实际出车天数计算补助", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "对保底公里数设置合理上限，防止制度套利", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "对已发生的异常补助进行追溯核查", font: "Arial", size: 21 })] }),

            new Paragraph({ children: [new PageBreak()] }),

            // 风险三
            new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("四、风险三：出车天数超月份天数（高风险）")] }),

            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("4.1 审计发现（直接数据矛盾）")] }),
            new Paragraph({ children: [new TextRun({ text: "发现多名司机出车天数超出当月实际天数，直接构成数据矛盾，属于高风险事项：", font: "Arial", size: 21 })] }),

            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [2000, 1500, 2000, 1500, 2026],
                rows: [
                    new TableRow({ children: [
                        headerCell("司机", 2000),
                        headerCell("月份", 1500),
                        headerCell("出车天数", 2000),
                        headerCell("当月总天数", 1500),
                        headerCell("异常情况", 2026)
                    ]}),
                    new TableRow({ children: [
                        highlightCell("王小虎", 2000, "FF6B6B"),
                        dataCell("2月", 1500),
                        highlightCell("32天", 2000, "FF6B6B"),
                        dataCell("28天", 1500),
                        highlightCell("超出4天", 2026, "FF6B6B")
                    ]}),
                    new TableRow({ children: [
                        highlightCell("王红平", 2000, "FF6B6B"),
                        dataCell("2月", 1500),
                        highlightCell("31天", 2000, "FF6B6B"),
                        dataCell("28天", 1500),
                        highlightCell("超出3天", 2026, "FF6B6B")
                    ]}),
                    new TableRow({ children: [
                        highlightCell("王红平", 2000, "FFD966"),
                        dataCell("8月", 1500),
                        highlightCell("32天", 2000, "FFD966"),
                        dataCell("31天", 1500),
                        highlightCell("超出1天", 2026, "FFD966")
                    ]}),
                ]
            }),

            new Paragraph({ spacing: { before: 200, after: 100 }, children: [] }),
            new Paragraph({ children: [new TextRun({ text: "重点关注：王红平全年补助金额最高，达41,888元，为一级核查对象。", font: "Arial", size: 21, bold: true, color: "C00000" })] }),

            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("4.2 审计建议")] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "立即核实上述数据差异，调取原始考勤记录、出车日志等证据", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "对王红平进行重点约谈，了解数据异常原因", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "排查数据录入环节是否存在人为错误或故意虚报", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "追溯调整相关补助金额，追回多发款项", font: "Arial", size: 21 })] }),

            new Paragraph({ spacing: { before: 400, after: 200 }, children: [] }),

            // 风险四
            new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("五、风险四：个别司机里程异常高（高风险）")] }),

            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("5.1 审计发现")] }),
            new Paragraph({ children: [new TextRun({ text: "8月数据中发现日均公里数异常偏高的司机，建议通过油卡消费、GPS行车记录仪进行交叉验证：", font: "Arial", size: 21 })] }),

            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [2000, 2500, 2000, 2526],
                rows: [
                    new TableRow({ children: [
                        headerCell("司机", 2000),
                        headerCell("8月日均公里", 2500),
                        headerCell("全年补助", 2000),
                        headerCell("备注", 2526)
                    ]}),
                    new TableRow({ children: [
                        highlightCell("童庆龙", 2000, "FF6B6B"),
                        highlightCell("586公里/天", 2500, "FF6B6B"),
                        dataCell("30,673元", 2000),
                        dataCell("需交叉验证", 2526)
                    ]}),
                    new TableRow({ children: [
                        highlightCell("周华", 2000, "FF6B6B"),
                        highlightCell("477公里/天", 2500, "FF6B6B"),
                        dataCell("24,278元", 2000),
                        dataCell("需交叉验证", 2526)
                    ]}),
                    new TableRow({ children: [
                        dataCell("姚东旭", 2000),
                        dataCell("283公里/天", 2500),
                        highlightCell("37,098元", 2000, "FFD966"),
                        highlightCell("全年第二高", 2526, "FFD966")
                    ]}),
                ]
            }),

            new Paragraph({ spacing: { before: 200, after: 100 }, children: [] }),
            new Paragraph({ children: [new TextRun({ text: "注：按每天行驶8小时、平均时速60公里计算，日均理论上限约480公里。童庆龙586公里/天已超出合理范围。", font: "Arial", size: 20, italics: true, color: "666666" })] }),

            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("5.2 审计建议")] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "调取上述司机的油卡充值记录、消费记录，验证里程真实性", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "核查GPS行车记录仪数据，与计提公里进行比对", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "核实是否存在重复计算里程、虚报里程等情况", font: "Arial", size: 21 })] }),

            new Paragraph({ children: [new PageBreak()] }),

            // 风险五
            new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("六、风险五：周末补助超限（中高风险）")] }),

            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("6.1 制度规定")] }),
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [3000, 3000, 3026],
                rows: [
                    new TableRow({ children: [
                        headerCell("出车类型", 3000),
                        headerCell("周末补助上限", 3000),
                        headerCell("补助标准", 3026)
                    ]}),
                    new TableRow({ children: [
                        dataCell("月包车", 3000),
                        dataCell("每月≤4天", 3000),
                        dataCell("65元/天", 3026)
                    ]}),
                    new TableRow({ children: [
                        dataCell("非月包车", 3000),
                        dataCell("每月≤5天", 3000),
                        dataCell("65元/天", 3026)
                    ]}),
                    new TableRow({ children: [
                        dataCell("外埠", 3000),
                        dataCell("每月≤8天", 3000),
                        dataCell("65元/天", 3026)
                    ]}),
                ]
            }),

            new Paragraph({ spacing: { before: 300, after: 200 }, children: [] }),
            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("6.2 审计发现")] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "王胜路8月周末出车8天，超过月包车4天上限2倍", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "苏保国8月周末出车6天，超过月包车4天上限", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "26名司机周末出车达5天，触及非月包车上限", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "8月全队周末补助合计：14,690元", font: "Arial", size: 21, bold: true })] }),

            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("6.3 审计建议")] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "核查上述超限司机是否属于外埠出车，如非外埠则超出规定上限", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "核实出车类型的分类是否准确，防止类型错报", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "超限部分应不予支付或追回", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "建议优化制度，明确周末补助的审批流程", font: "Arial", size: 21 })] }),

            new Paragraph({ spacing: { before: 600, after: 200 }, children: [] }),

            // 总结
            new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("七、审计总结与后续建议")] }),

            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("7.1 重点关注事项")] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "王红平：全年补助最高（41,888元），2月、8月出车天数均超月份天数，一级核查对象", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "虞岩：保底公里套利典型案例，1天仅行驶2公里却按3000公里计提", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "numbers", level: 0 }, children: [new TextRun({ text: "童庆龙、周华：日均公里数异常高，需交叉验证", font: "Arial", size: 21 })] }),

            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("7.2 制度完善建议")] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "重新评估保底公里制度，防止制度套利", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "完善出车天数统计机制，确保数据准确性", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "建立里程异常预警机制，设置合理上限", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "定期通过油卡、GPS数据进行交叉验证", font: "Arial", size: 21 })] }),
            new Paragraph({ numbering: { reference: "bullets", level: 0 }, children: [new TextRun({ text: "更新制度依据文件，完善审批程序", font: "Arial", size: 21 })] }),

            new Paragraph({ spacing: { before: 600, after: 200 }, children: [] }),

            // 附件
            new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("八、附件清单")] }),
            new Table({
                width: { size: 9026, type: WidthType.DXA },
                columnWidths: [1500, 5000, 2526],
                rows: [
                    new TableRow({ children: [
                        headerCell("序号", 1500),
                        headerCell("文件名称", 5000),
                        headerCell("说明", 2526)
                    ]}),
                    new TableRow({ children: [
                        dataCell("1", 1500, AlignmentType.CENTER),
                        dataCell("2025年8月司机补助.xlsx", 5000),
                        dataCell("8月详细数据", 2526)
                    ]}),
                    new TableRow({ children: [
                        dataCell("2", 1500, AlignmentType.CENTER),
                        dataCell("车辆服务中心－2025年司机差费发放表.xlsx", 5000),
                        dataCell("全年汇总数据", 2526)
                    ]}),
                    new TableRow({ children: [
                        dataCell("3", 1500, AlignmentType.CENTER),
                        dataCell("2月份服务一队、司机计提公里统计.xls", 5000),
                        dataCell("2月数据（追溯核查）", 2526)
                    ]}),
                    new TableRow({ children: [
                        dataCell("4", 1500, AlignmentType.CENTER),
                        dataCell("差旅费管理细则.docx", 5000),
                        dataCell("制度依据文件", 2526)
                    ]}),
                    new TableRow({ children: [
                        dataCell("5", 1500, AlignmentType.CENTER),
                        dataCell("10月司机差费计算表--模板.xlsx", 5000),
                        dataCell("阶梯费率标准表", 2526)
                    ]}),
                ]
            }),

            new Paragraph({ spacing: { before: 800, after: 200 }, children: [] }),
            new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "审计人员：苏浩", font: "Arial", size: 21 })] }),
            new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "上会会计师事务所（特殊普通合伙）", font: "Arial", size: 21 })] }),
            new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "2026年5月15日", font: "Arial", size: 21 })] }),
        ]
    }]
});

Packer.toBuffer(doc).then(buffer => {
    fs.writeFileSync("D:/Users/12844/Desktop/车辆服务中心2025年度司机差旅费补助审计报告.docx", buffer);
    console.log("Word文档已成功生成！");
    console.log("文件路径：D:/Users/12844/Desktop/车辆服务中心2025年度司机差旅费补助审计报告.docx");
});
