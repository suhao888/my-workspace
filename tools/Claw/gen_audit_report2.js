const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  HeadingLevel, PageNumber, Header, Footer
} = require('docx');
const fs = require('fs');

const FONT = "\u4effSong_GB2312";
const FONT_SONG = "\u4eff\u5b8b_GB2312";
const FONT_HEI = "\u9ed1\u4f53";
const FONT_BIAOSONG = "\u65b9\u6b63\u5c0f\u6807\u5b8b\u7b80\u4f53";

function ip(text, opts) {
  return new Paragraph({
    children: [new TextRun(Object.assign({ text: text, font: "\u4effSong_GB2312", size: 28 }, opts || {}))],
    indent: { firstLine: 560 },
    spacing: { line: 480, lineRule: "auto" },
    alignment: AlignmentType.JUSTIFIED,
  });
}
function cp(text, opts) {
  return new Paragraph({
    children: [new TextRun(Object.assign({ text: text, font: "\u4effSong_GB2312", size: 28 }, opts || {}))],
    alignment: AlignmentType.CENTER,
    spacing: { line: 480, lineRule: "auto" },
  });
}
function rp(text, opts) {
  return new Paragraph({
    children: [new TextRun(Object.assign({ text: text, font: "\u4effSong_GB2312", size: 28 }, opts || {}))],
    alignment: AlignmentType.RIGHT,
    spacing: { line: 480, lineRule: "auto" },
  });
}
function np(text, opts) {
  return new Paragraph({
    children: [new TextRun(Object.assign({ text: text, font: "\u4effSong_GB2312", size: 28 }, opts || {}))],
    spacing: { line: 480, lineRule: "auto" },
    alignment: AlignmentType.JUSTIFIED,
  });
}
function lp(text, opts) {
  return new Paragraph({
    children: [new TextRun(Object.assign({ text: text, font: "\u4effSong_GB2312", size: 28 }, opts || {}))],
    indent: { left: 560, firstLine: 560 },
    spacing: { line: 480, lineRule: "auto" },
    alignment: AlignmentType.JUSTIFIED,
  });
}
function empty() { return new Paragraph({ children: [], spacing: { line: 480 } }); }

const br = { style: BorderStyle.SINGLE, size: 1, color: "000000" };
const brs = { top: br, bottom: br, left: br, right: br };

function mc(text, w, bold, shade) {
  return new TableCell({
    borders: brs,
    width: { size: w, type: WidthType.DXA },
    shading: shade ? { fill: "D9E1F2", type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      children: [new TextRun({ text: text, font: "\u4effSong_GB2312", size: 26, bold: !!bold })],
      alignment: AlignmentType.CENTER,
      spacing: { line: 360 },
    })]
  });
}

// ----- texts -----
var t = {
  mainTitle1: "\u5173\u4e8e\u5927\u6e2f\u5de5\u7a0b\u804c\u4e1a\u6280\u672f\u5b66\u9662\u5212\u8f6c",
  mainTitle2: "\u5929\u6d25\u77f3\u6cb9\u804c\u4e1a\u6280\u672f\u5b66\u9662\u8d44\u4ea7",
  mainTitle3: "\u65e0\u629a\u62bc\u62c5\u4fdd\u4e13\u9879\u5ba1\u8ba1\u62a5\u544a",
  reportNo: "\u4e0a\u4f1a\u5ba1\u4e13\u5b57\u30162026\u3017\u7b2c\u00d7\u00d7\u53f7",
  client: "\u5929\u6d25\u77f3\u6cb9\u804c\u4e1a\u6280\u672f\u5b66\u9662",
  firm: "\u4e0a\u4f1a\u4f1a\u8ba1\u5e08\u4e8b\u52a1\u6240\uff08\u7279\u6b8a\u666e\u901a\u5408\u4f19\uff09",
  reportDate: "\u4e8c\u25cb\u4e8c\u516d\u5e74  \u6708  \u65e5",

  // 正文
  salutation: "\u5929\u6d25\u77f3\u6cb9\u804c\u4e1a\u6280\u672f\u5b66\u9662\uff1a",
  h1: "\u4e00\u3001\u57fa\u672c\u60c5\u51b5",
  h2: "\u4e8c\u3001\u88ab\u5ba1\u8ba1\u5355\u4f4d\u7ba1\u7406\u5c42\u7684\u8d23\u4efb",
  h3: "\u4e09\u3001\u6ce8\u518c\u4f1a\u8ba1\u5e08\u7684\u8d23\u4efb",
  h4: "\u56db\u3001\u5ba1\u8ba1\u8303\u56f4\u4e0e\u4e3b\u8981\u5ba1\u8ba1\u7a0b\u5e8f",
  h5: "\u4e94\u3001\u5ba1\u8ba1\u7ed3\u8bba",
  h6: "\u516d\u3001\u5212\u8f6c\u8d44\u4ea7\u660e\u7ec6",
  h7: "\u4e03\u3001\u7279\u522b\u8bf4\u660e",

  p1a: "\u53d7\u5929\u6d25\u77f3\u6cb9\u804c\u4e1a\u6280\u672f\u5b66\u9662\uff08\u4ee5\u4e0b\u7b80\u79f0\u201c\u59d4\u6258\u65b9\u201d\u6216\u201c\u8d35\u6821\u201d\uff09\u59d4\u6258\uff0c\u4e0a\u4f1a\u4f1a\u8ba1\u5e08\u4e8b\u52a1\u6240\uff08\u7279\u6b8a\u666e\u901a\u5408\u4f19\uff09\uff08\u4ee5\u4e0b\u7b80\u79f0\u201c\u672c\u6240\u201d\uff09\u5bf9\u5927\u6e2f\u5de5\u7a0b\u804c\u4e1a\u6280\u672f\u5b66\u9662\uff08\u4ee5\u4e0b\u7b80\u79f0\u201c\u5212\u8f6c\u65b9\u201d\uff09\u65e0\u507f\u5212\u8f6c\u7ed9\u5929\u6d25\u77f3\u6cb9\u804c\u4e1a\u6280\u672f\u5b66\u9662\u7684\u76f8\u5173\u8d44\u4ea7\uff08\u4ee5\u4e0b\u7b80\u79f0\u201c\u5212\u8f6c\u8d44\u4ea7\u201d\uff09\u662f\u5426\u5b58\u5728\u629a\u62bc\u3001\u8d28\u62bc\u6216\u5bf9\u5916\u63d0\u4f9b\u62c5\u4fdd\u60c5\u5f62\u8fdb\u884c\u4e13\u9879\u5ba1\u8ba1\u3002\u672c\u6b21\u5ba1\u8ba1\u57fa\u51c6\u65e5\u4e3a       \uff08\u4ee5\u5b9e\u9645\u786e\u8ba4\u65e5\u671f\u4e3a\u51c6\uff09\u3002",
  p1b: "\u5927\u6e2f\u5de5\u7a0b\u804c\u4e1a\u6280\u672f\u5b66\u9662\u4e0e\u5929\u6d25\u77f3\u6cb9\u804c\u4e1a\u6280\u672f\u5b66\u9662\u540c\u5c5e\u540c\u4e00\u63a7\u5236\u4e3b\u4f53\u4e0b\u7684\u9ad8\u7b49\u5b66\u6821\uff0c\u672c\u6b21\u65e0\u507f\u5212\u8f6c\u7cfb\u540c\u4e00\u63a7\u5236\u4e0b\u7684\u8d44\u4ea7\u8f6c\u79fb\u884c\u4e3a\uff0c\u4e0d\u6d89\u53ca\u5546\u4e1a\u5bf9\u4ef7\u3002\u5212\u8f6c\u5b8c\u6210\u540e\uff0c\u4e0a\u8ff0\u8d44\u4ea7\u5f52\u5c5e\u5929\u6d25\u77f3\u6cb9\u804c\u4e1a\u6280\u672f\u5b66\u9662\uff0c\u76f8\u5e94\u6743\u5229\u4e49\u52a1\u4ea6\u968f\u4e4b\u8f6c\u79fb\u3002\u4e3a\u4fdd\u8bc1\u59d4\u6258\u65b9\u63a5\u6536\u8d44\u4ea7\u7684\u65e0\u7455\u75b5\u6027\u3001\u5408\u89c4\u6027\uff0c\u59d4\u6258\u65b9\u8981\u6c42\u672c\u6240\u5c31\u5212\u8f6c\u8d44\u4ea7\u662f\u5426\u8bbe\u6709\u629a\u62bc\u3001\u8d28\u62bc\u53ca\u662f\u5426\u5bf9\u5916\u63d0\u4f9b\u62c5\u4fdd\u51fa\u5177\u4e13\u9879\u5ba1\u8ba1\u62a5\u544a\u3002",

  p2a: "\u5927\u6e2f\u5de5\u7a0b\u804c\u4e1a\u6280\u672f\u5b66\u9662\u7ba1\u7406\u5c42\uff08\u4ee5\u4e0b\u7b80\u79f0\u201c\u88ab\u5ba1\u8ba1\u5355\u4f4d\u7ba1\u7406\u5c42\u201d\uff09\u8d1f\u8d23\u6309\u7167\u56fd\u5bb6\u6cd5\u5f8b\u6cd5\u89c4\u548c\u76f8\u5173\u5236\u5ea6\u7684\u89c4\u5b9a\uff0c\u5bf9\u5212\u8f6c\u8d44\u4ea7\u8fdb\u884c\u786e\u8ba4\u3001\u8ba1\u91cf\u548c\u8bb0\u5f55\uff0c\u5e76\u4fdd\u8bc1\u6240\u63d0\u4f9b\u7684\u4f1a\u8ba1\u51ed\u8bc1\u3001\u8d26\u7c3f\u3001\u62a5\u8868\u53ca\u76f8\u5173\u8bf4\u660e\u6750\u6599\u771f\u5b9e\u3001\u5408\u6cd5\u3001\u5b8c\u6574\uff1b\u8d1f\u8d23\u786e\u8ba4\u5212\u8f6c\u8d44\u4ea7\u4e0a\u4e0d\u5b58\u5728\u4efb\u4f55\u629a\u62bc\u3001\u8d28\u62bc\u3001\u62c5\u4fdd\u53ca\u5176\u4ed6\u6743\u5229\u8d1f\u62c5\uff0c\u5e76\u5c31\u672c\u6b21\u4e13\u9879\u5ba1\u8ba1\u4e8b\u9879\u4f5c\u51fa\u4e66\u9762\u58f0\u660e\u3002",
  p3a: "\u672c\u6240\u7684\u8d23\u4efb\u662f\u5728\u6267\u884c\u5ba1\u8ba1\u5de5\u4f5c\u7684\u57fa\u7840\u4e0a\uff0c\u5bf9\u5212\u8f6c\u8d44\u4ea7\u662f\u5426\u5b58\u5728\u629a\u62bc\u3001\u8d28\u62bc\u6216\u5bf9\u5916\u62c5\u4fdd\u60c5\u5f62\u53d1\u8868\u72ec\u7acb\u5ba1\u8ba1\u610f\u89c1\u3002\u672c\u6240\u6309\u7167\u4e2d\u56fd\u6ce8\u518c\u4f1a\u8ba1\u5e08\u5ba1\u8ba1\u51c6\u5219\u53ca\u76f8\u5173\u804c\u4e1a\u9053\u5fb7\u89c4\u8303\uff0c\u72ec\u7acb\u6267\u884c\u4e86\u4e13\u9879\u5ba1\u8ba1\u5de5\u4f5c\uff0c\u4ee5\u5145\u5206\u9002\u5f53\u7684\u5ba1\u8ba1\u8bc1\u636e\u4e3a\u57fa\u7840\uff0c\u5bf9\u51fa\u5177\u672c\u62a5\u544a\u8d1f\u8d23\u3002",
  h41: "\uff08\u4e00\uff09\u5ba1\u8ba1\u8303\u56f4",
  p4a: "\u672c\u6b21\u4e13\u9879\u5ba1\u8ba1\u7684\u8303\u56f4\u4e3a\u5927\u6e2f\u5de5\u7a0b\u804c\u4e1a\u6280\u672f\u5b66\u9662\u62df\u65e0\u507f\u5212\u8f6c\u7ed9\u5929\u6d25\u77f3\u6cb9\u804c\u4e1a\u6280\u672f\u5b66\u9662\u7684\u5168\u90e8\u8d44\u4ea7\uff0c\u5177\u4f53\u6e05\u5355\u4ee5\u53cc\u65b9\u7b7e\u7f72\u7684\u300a\u8d44\u4ea7\u5212\u8f6c\u534f\u8bae\u300b\u53ca\u9644\u4ef6\u4e3a\u51c6\u3002",
  h42: "\uff08\u4e8c\uff09\u4e3b\u8981\u5ba1\u8ba1\u7a0b\u5e8f",
  p4b: "\u672c\u6240\u5b9e\u65bd\u4e86\u4ee5\u4e0b\u4e3b\u8981\u5ba1\u8ba1\u7a0b\u5e8f\uff1a",
  pa1: "1.\u5ba1\u9605\u88ab\u5ba1\u8ba1\u5355\u4f4d\u63d0\u4f9b\u7684\u8d44\u4ea7\u6e05\u5355\u3001\u8d44\u4ea7\u6743\u5c5e\u8bc1\u660e\u6587\u4ef6\u53ca\u76f8\u5173\u8d26\u7c3f\u51ed\u8bc1\uff0c\u6838\u5b9e\u5212\u8f6c\u8d44\u4ea7\u7684\u8d26\u9762\u8bb0\u5f55\u60c5\u51b5\uff1b",
  pa2: "2.\u67e5\u9605\u88ab\u5ba1\u8ba1\u5355\u4f4d\u4e0e\u9280\u884c\u7b49\u91d1\u878d\u673a\u6784\u7b7e\u8ba2\u7684\u501f\u6b3e\u5408\u540c\u3001\u629a\u62bc\u5408\u540c\u53ca\u76f8\u5173\u629a\u62bc\u767b\u8bb0\u8bb0\u5f55\uff0c\u786e\u8ba4\u5212\u8f6c\u8d44\u4ea7\u662f\u5426\u5df2\u88ab\u8bbe\u5b9a\u629a\u62bc\u6216\u8d28\u62bc\uff1b",
  pa3: "3.\u901a\u8fc7\u4e0d\u52a8\u4ea7\u767b\u8bb0\u4e2d\u5fc3\u7b49\u6743\u5a01\u673a\u6784\u67e5\u8be2\u4e0d\u52a8\u4ea7\u629a\u62bc\u767b\u8bb0\u60c5\u51b5\uff0c\u6838\u5b9e\u76f8\u5173\u623f\u4ea7\u3001\u571f\u5730\u4f7f\u7528\u6743\u662f\u5426\u5b58\u5728\u629a\u62bc\u767b\u8bb0\uff1b",
  pa4: "4.\u67e5\u9605\u88ab\u5ba1\u8ba1\u5355\u4f4d\u5bf9\u5916\u62c5\u4fdd\u53f0\u8d26\u53ca\u76f8\u5173\u534f\u8bae\uff0c\u6838\u5b9e\u5212\u8f6c\u8d44\u4ea7\u662f\u5426\u5df2\u88ab\u7528\u4e8e\u5bf9\u5916\u63d0\u4f9b\u62c5\u4fdd\uff1b",
  pa5: "5.\u83b7\u53d6\u88ab\u5ba1\u8ba1\u5355\u4f4d\u7ba1\u7406\u5c42\u5173\u4e8e\u5212\u8f6c\u8d44\u4ea7\u4e0a\u4e0d\u5b58\u5728\u4efb\u4f55\u629a\u62bc\u3001\u8d28\u62bc\u3001\u62c5\u4fdd\u7684\u4e66\u9762\u58f0\u660e\uff1b",
  pa6: "6.\u5bf9\u4ee5\u4e0a\u7a0b\u5e8f\u6240\u83b7\u53d6\u7684\u5ba1\u8ba1\u8bc1\u636e\u8fdb\u884c\u7efc\u5408\u8bc4\u4ef7\uff0c\u5f62\u6210\u5ba1\u8ba1\u7ed3\u8bba\u3002",

  p5a: "\u7ecf\u672c\u6240\u4e13\u9879\u5ba1\u8ba1\uff0c\u622a\u81f3\u5ba1\u8ba1\u57fa\u51c6\u65e5\uff0c\u5927\u6e2f\u5de5\u7a0b\u804c\u4e1a\u6280\u672f\u5b66\u9662\u62df\u65e0\u507f\u5212\u8f6c\u7ed9\u5929\u6d25\u77f3\u6cb9\u804c\u4e1a\u6280\u672f\u5b66\u9662\u7684\u5212\u8f6c\u8d44\u4ea7\uff0c\u5747\u4e0d\u5b58\u5728\u4ee5\u4e0b\u60c5\u5f62\uff1a",
  p5b1: "\uff08\u4e00\uff09\u5212\u8f6c\u8d44\u4ea7\u4e0d\u5b58\u5728\u629a\u62bc\u6216\u8d28\u62bc\u62c5\u4fdd\uff0c\u5747\u672a\u88ab\u8bbe\u5b9a\u629a\u62bc\u6743\u6216\u8d28\u6743\uff1b",
  p5b2: "\uff08\u4e8c\uff09\u5212\u8f6c\u8d44\u4ea7\u672a\u88ab\u7528\u4e8e\u5bf9\u5916\u63d0\u4f9b\u62c5\u4fdd\uff08\u5305\u62ec\u4f46\u4e0d\u9650\u4e8e\u4fdd\u8bc1\u62c5\u4fdd\u3001\u629a\u62bc\u62c5\u4fdd\u3001\u8d28\u62bc\u62c5\u4fdd\uff09\uff1b",
  p5b3: "\uff08\u4e09\uff09\u5212\u8f6c\u8d44\u4ea7\u4e0d\u5b58\u5728\u5176\u4ed6\u6743\u5229\u8d1f\u62c5\u6216\u6cd5\u5f8b\u4e0a\u7684\u6743\u5229\u7455\u75b5\uff0c\u4e0d\u5f71\u54cd\u5212\u8f6c\u7684\u5408\u6cd5\u6709\u6548\u6027\u3002",
  p5c: "\u7efc\u4e0a\uff0c\u672c\u6240\u8ba4\u4e3a\uff0c\u4e0a\u8ff0\u5212\u8f6c\u8d44\u4ea7\u6743\u5c5e\u660e\u786e\u3001\u65e0\u629a\u62bc\u3001\u65e0\u8d28\u62bc\u3001\u65e0\u5bf9\u5916\u62c5\u4fdd\uff0c\u53ef\u65e0\u98ce\u9669\u63a5\u6536\u3002",

  p6a: "\u5212\u8f6c\u8d44\u4ea7\u660e\u7ec6\u60c5\u51b5\u5982\u4e0b\u8868\u6240\u793a\uff08\u5177\u4f53\u91d1\u989d\u4ee5\u53cc\u65b9\u7b7e\u8ba4\u7684\u8d44\u4ea7\u6e05\u5355\u53ca\u6700\u7ec8\u5ba1\u8ba1\u786e\u8ba4\u6570\u4e3a\u51c6\uff09\uff1a",
  p6b: "\uff08\u6ce8\uff1a\u8d44\u4ea7\u660e\u7ec6\u6e05\u5355\u4ee5\u9644\u4ef6\u5f62\u5f0f\u9644\u4e8e\u62a5\u544a\u540e\uff09",

  p7a: "\uff08\u4e00\uff09\u672c\u62a5\u544a\u4ec5\u5c31\u5212\u8f6c\u8d44\u4ea7\u662f\u5426\u5b58\u5728\u629a\u62bc\u3001\u8d28\u62bc\u53ca\u5bf9\u5916\u62c5\u4fdd\u60c5\u5f62\u53d1\u8868\u610f\u89c1\uff0c\u4e0d\u6d89\u53ca\u8d44\u4ea7\u4ef7\u5024\u8bc4\u4f30\u53ca\u5176\u4ed6\u5ba1\u8ba1\u4e8b\u9879\u3002",
  p7b: "\uff08\u4e8c\uff09\u672c\u62a5\u544a\u4f9d\u8d56\u4e8e\u88ab\u5ba1\u8ba1\u5355\u4f4d\u7ba1\u7406\u5c42\u63d0\u4f9b\u7684\u8d44\u6599\u53ca\u4e66\u9762\u58f0\u660e\uff0c\u5982\u88ab\u5ba1\u8ba1\u5355\u4f4d\u63d0\u4f9b\u7684\u8d44\u6599\u5b58\u5728\u4e0d\u771f\u5b9e\u3001\u4e0d\u5b8c\u6574\u60c5\u5f62\uff0c\u672c\u6240\u5c06\u4e0d\u627f\u62c5\u76f8\u5e94\u6cd5\u5f8b\u8d23\u4efb\u3002",
  p7c: "\uff08\u4e09\uff09\u672c\u62a5\u544a\u7cfb\u4e13\u9879\u51fa\u5177\uff0c\u4ec5\u4f9b\u5929\u6d25\u77f3\u6cb9\u804c\u4e1a\u6280\u672f\u5b66\u9662\u5728\u672c\u6b21\u8d44\u4ea7\u65e0\u507f\u5212\u8f6c\u4e8b\u9879\u4e2d\u4f7f\u7528\uff0c\u4e0d\u5f97\u7528\u4e8e\u5176\u4ed6\u76ee\u7684\u3002\u5982\u5c06\u672c\u62a5\u544a\u7528\u4e8e\u5176\u4ed6\u76ee\u7684\uff0c\u672c\u6240\u6982\u4e0d\u8d1f\u8d23\u3002",
  p7d: "\uff08\u56db\uff09\u672c\u62a5\u544a\u81ea\u51fa\u5177\u4e4b\u65e5\u8d77\u4e00\u5e74\u5185\u6709\u6548\u3002\u5982\u5212\u8f6c\u8d44\u4ea7\u60c5\u51b5\u53d1\u751f\u53d8\u5316\uff0c\u672c\u62a5\u544a\u81ea\u52a8\u5931\u6548\u3002",

  // 附件
  attTitle: "\u9644\u4ef6\uff1a\u88ab\u5ba1\u8ba1\u5355\u4f4d\u7ba1\u7406\u5c42\u58f0\u660e\u4e66",
  attSal: "\u4e0a\u4f1a\u4f1a\u8ba1\u5e08\u4e8b\u52a1\u6240\uff08\u7279\u6b8a\u666e\u901a\u5408\u4f19\uff09\uff1a",
  attIntro: "\u672c\u5355\u4f4d\u5c31\u5927\u6e2f\u5de5\u7a0b\u804c\u4e1a\u6280\u672f\u5b66\u9662\u65e0\u507f\u5212\u8f6c\u7ed9\u5929\u6d25\u77f3\u6cb9\u804c\u4e1a\u6280\u672f\u5b66\u9662\u8d44\u4ea7\u4e8b\u9879\uff0c\u7279\u5411\u8d35\u6240\u90d1\u91cd\u58f0\u660e\u5982\u4e0b\uff1a",
  att1: "\u4e00\u3001\u672c\u6b21\u5212\u8f6c\u8d44\u4ea7\u6e05\u5355\u6240\u5217\u5404\u9879\u8d44\u4ea7\uff0c\u5747\u7cfb\u672c\u5355\u4f4d\u5408\u6cd5\u62e5\u6709\uff0c\u6743\u5c5e\u6e05\u6670\uff0c\u65e0\u4ea7\u6743\u4e89\u8bae\u3002",
  att2: "\u4e8c\u3001\u4e0a\u8ff0\u8d44\u4ea7\u622a\u81f3\u58f0\u660e\u65e5\uff0c\u5747\u672a\u5411\u4efb\u4f55\u5355\u4f4d\u6216\u4e2a\u4eba\u8bbe\u5b9a\u629a\u62bc\u6743\u6216\u8d28\u6743\uff0c\u4e0d\u5b58\u5728\u4efb\u4f55\u5f62\u5f0f\u7684\u629a\u62bc\u6216\u8d28\u62bc\u62c5\u4fdd\u3002",
  att3: "\u4e09\u3001\u4e0a\u8ff0\u8d44\u4ea7\u622a\u81f3\u58f0\u660e\u65e5\uff0c\u5747\u672a\u88ab\u7528\u4e8e\u5bf9\u5916\u63d0\u4f9b\u4efb\u4f55\u5f62\u5f0f\u7684\u62c5\u4fdd\uff08\u5305\u62ec\u4f46\u4e0d\u9650\u4e8e\u4fdd\u8bc1\u62c5\u4fdd\u3001\u629a\u62bc\u62c5\u4fdd\u3001\u8d28\u62bc\u62c5\u4fdd\u7b49\uff09\u3002",
  att4: "\u56db\u3001\u4e0a\u8ff0\u8d44\u4ea7\u4e0d\u5b58\u5728\u4efb\u4f55\u67e5\u5c01\u3001\u6263\u62bc\u3001\u51bb\u7ed3\u7b49\u53f8\u6cd5\u5f3a\u5236\u63aa\u65bd\uff0c\u4e0d\u5b58\u5728\u4efb\u4f55\u6d89\u8bc9\u7ea0\u7eb7\u3002",
  att5: "\u4e94\u3001\u672c\u5355\u4f4d\u5df2\u5411\u8d35\u6240\u63d0\u4f9b\u4e86\u4e0e\u672c\u6b21\u4e13\u9879\u5ba1\u8ba1\u76f8\u5173\u7684\u5168\u90e8\u771f\u5b9e\u3001\u5b8c\u6574\u3001\u5408\u6cd5\u7684\u8d44\u6599\uff0c\u6240\u6709\u4e66\u9762\u8d44\u6599\u53ca\u53e3\u5934\u89e3\u91ca\u5747\u4e0d\u5b58\u5728\u91cd\u5927\u9519\u8bef\u3001\u9057\u6f0f\u6216\u8bef\u5bfc\u6027\u9648\u8ff0\u3002",
  att6: "\u516d\u3001\u5982\u4e0a\u8ff0\u58f0\u660e\u5185\u5bb9\u5b58\u5728\u4e0d\u5b9e\uff0c\u7531\u6b64\u4ea7\u751f\u7684\u4e00\u5207\u6cd5\u5f8b\u8d23\u4efb\u7531\u672c\u5355\u4f4d\u627f\u62c5\uff0c\u4e0e\u8d35\u6240\u65e0\u5173\u3002",
  attFirm: "\u5927\u6e2f\u5de5\u7a0b\u804c\u4e1a\u6280\u672f\u5b66\u9662\uff08\u516c\u7ae0\uff09",
  attSign: "\u6cd5\u5b9a\u4ee3\u8868\u4eba\uff08\u7b7e\u5b57\uff09\uff1a                ",
  attDate: "\u4e8c\u25cb\u4e8c\u516d\u5e74   \u6708   \u65e5"
};

const doc = new Document({
  styles: {
    default: { document: { run: { font: "\u4effSong_GB2312", size: 28 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: "\u9ed1\u4f53", size: 28, bold: true, color: "000000" },
        paragraph: { spacing: { before: 240, after: 120, line: 480, lineRule: "auto" }, outlineLevel: 0 }
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
            children: [new TextRun({ text: t.firm, font: "\u4effSong_GB2312", size: 22, color: "808080" })],
            alignment: AlignmentType.CENTER,
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" } }
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: "\u7b2c ", font: "\u4effSong_GB2312", size: 22, color: "808080" }),
              new TextRun({ children: [PageNumber.CURRENT], font: "\u4effSong_GB2312", size: 22, color: "808080" }),
              new TextRun({ text: " \u9875\uff0c\u5171 ", font: "\u4effSong_GB2312", size: 22, color: "808080" }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "\u4effSong_GB2312", size: 22, color: "808080" }),
              new TextRun({ text: " \u9875", font: "\u4effSong_GB2312", size: 22, color: "808080" }),
            ],
            alignment: AlignmentType.CENTER,
          })
        ]
      })
    },
    children: [
      // ===== 封面 =====
      empty(), empty(),
      new Paragraph({
        children: [new TextRun({ text: t.mainTitle1, font: "\u65b9\u6b63\u5c0f\u6807\u5b8b\u7b80\u4f53", size: 44, bold: true })],
        alignment: AlignmentType.CENTER, spacing: { line: 480 }
      }),
      new Paragraph({
        children: [new TextRun({ text: t.mainTitle2, font: "\u65b9\u6b63\u5c0f\u6807\u5b8b\u7b80\u4f53", size: 44, bold: true })],
        alignment: AlignmentType.CENTER, spacing: { line: 480 }
      }),
      new Paragraph({
        children: [new TextRun({ text: t.mainTitle3, font: "\u65b9\u6b63\u5c0f\u6807\u5b8b\u7b80\u4f53", size: 44, bold: true })],
        alignment: AlignmentType.CENTER, spacing: { line: 480 }
      }),
      empty(),
      cp(t.reportNo, { color: "595959", size: 26 }),
      empty(), empty(),
      np("\u59d4  \u6258  \u65b9\uff1a" + t.client),
      np("\u51fa \u5177 \u673a \u6784\uff1a" + t.firm),
      np("\u62a5 \u544a \u65e5 \u671f\uff1a" + t.reportDate),
      empty(),
      new Paragraph({
        children: [new TextRun({ text: "", font: "\u4effSong_GB2312", size: 28 })],
        border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: "000000" } },
        spacing: { before: 480, after: 120 }
      }),
      cp("\u5730\u5740\uff1a\uff08\u4e8b\u52a1\u6240\u5730\u5740\uff09          \u7535\u8bdd\uff1a\uff08\u8054\u7cfb\u7535\u8bdd\uff09", { size: 22, color: "595959" }),

      // ===== 正文（新页） =====
      new Paragraph({ children: [], pageBreakBefore: true }),
      new Paragraph({
        children: [new TextRun({ text: t.mainTitle1, font: "\u65b9\u6b63\u5c0f\u6807\u5b8b\u7b80\u4f53", size: 36, bold: true })],
        alignment: AlignmentType.CENTER, spacing: { line: 480 }
      }),
      new Paragraph({
        children: [new TextRun({ text: t.mainTitle2, font: "\u65b9\u6b63\u5c0f\u6807\u5b8b\u7b80\u4f53", size: 36, bold: true })],
        alignment: AlignmentType.CENTER, spacing: { line: 480 }
      }),
      new Paragraph({
        children: [new TextRun({ text: t.mainTitle3, font: "\u65b9\u6b63\u5c0f\u6807\u5b8b\u7b80\u4f53", size: 36, bold: true })],
        alignment: AlignmentType.CENTER, spacing: { line: 480 }
      }),
      cp(t.reportNo, { color: "595959", size: 26 }),
      empty(),
      ip(t.salutation),
      empty(),
      // 一
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t.h1, font: "\u9ed1\u4f53", size: 28, bold: true })] }),
      ip(t.p1a),
      ip(t.p1b),
      // 二
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t.h2, font: "\u9ed1\u4f53", size: 28, bold: true })] }),
      ip(t.p2a),
      // 三
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t.h3, font: "\u9ed1\u4f53", size: 28, bold: true })] }),
      ip(t.p3a),
      // 四
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t.h4, font: "\u9ed1\u4f53", size: 28, bold: true })] }),
      ip(t.h41, { bold: false }),
      ip(t.p4a),
      ip(t.h42, { bold: false }),
      ip(t.p4b),
      lp(t.pa1),
      lp(t.pa2),
      lp(t.pa3),
      lp(t.pa4),
      lp(t.pa5),
      lp(t.pa6),
      // 五
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t.h5, font: "\u9ed1\u4f53", size: 28, bold: true })] }),
      ip(t.p5a),
      lp(t.p5b1),
      lp(t.p5b2),
      lp(t.p5b3),
      ip(t.p5c),
      // 六
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t.h6, font: "\u9ed1\u4f53", size: 28, bold: true })] }),
      ip(t.p6a),
      empty(),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [500, 2100, 1500, 1900, 1880, 1480],
        rows: [
          new TableRow({
            tableHeader: true,
            children: [
              mc("\u5e8f\u53f7", 500, true, true),
              mc("\u8d44\u4ea7\u540d\u79f0", 2100, true, true),
              mc("\u8d44\u4ea7\u7c7b\u522b", 1500, true, true),
              mc("\u8d26\u9762\u4ef7\u5024\uff08\u5143\uff09", 1900, true, true),
              mc("\u662f\u5426\u8bbe\u6709\u629a\u62bc/\u8d28\u62bc", 1880, true, true),
              mc("\u662f\u5426\u5bf9\u5916\u62c5\u4fdd", 1480, true, true),
            ]
          }),
          new TableRow({
            children: [
              mc("\u5408\u8ba1", 500, true),
              mc("\u2014", 2100, true),
              mc("\u2014", 1500, true),
              mc("\u4ee5\u5b9e\u9645\u5ba1\u8ba1\u786e\u8ba4\u6570\u4e3a\u51c6", 1900, false),
              mc("\u65e0", 1880, true),
              mc("\u65e0", 1480, true),
            ]
          }),
        ]
      }),
      empty(),
      cp(t.p6b, { size: 24, color: "595959" }),
      // 七
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t.h7, font: "\u9ed1\u4f53", size: 28, bold: true })] }),
      ip(t.p7a),
      ip(t.p7b),
      ip(t.p7c),
      ip(t.p7d),
      empty(),
      // 签章
      rp(t.firm, { bold: true }),
      rp("\uff08\u76d6\u7ae0\uff09"),
      rp("\u6ce8\u518c\u4f1a\u8ba1\u5e08\uff08\u7b7e\u5b57\uff09\uff1a                "),
      rp("\u6ce8\u518c\u4f1a\u8ba1\u5e08\uff08\u7b7e\u5b57\uff09\uff1a                "),
      rp(t.reportDate),

      // ===== 附件页 =====
      new Paragraph({ children: [], pageBreakBefore: true }),
      new Paragraph({
        children: [new TextRun({ text: t.attTitle, font: "\u9ed1\u4f53", size: 32, bold: true })],
        alignment: AlignmentType.CENTER, spacing: { line: 480 }
      }),
      empty(),
      ip(t.attSal),
      empty(),
      ip(t.attIntro),
      ip(t.att1),
      ip(t.att2),
      ip(t.att3),
      ip(t.att4),
      ip(t.att5),
      ip(t.att6),
      empty(), empty(),
      rp(t.attFirm),
      rp(t.attSign),
      rp(t.attDate),
    ]
  }]
});

const outPath = "D:\\Users\\12844\\Desktop\\\\u5929\\u6d25\\u77f3\\u6cb9\\u804c\\u4e1a\\u6280\\u672f\\u5b66\\u9662\\\\u65e0\\u629a\\u62bc\\u62c5\\u4fdd\\u4e13\\u9879\\u5ba1\\u8ba1\\u62a5\\u544a.docx";
const outPath2 = "D:/Users/12844/Desktop/\u5929\u6d25\u77f3\u6cb9\u804c\u4e1a\u6280\u672f\u5b66\u9662/\u65e0\u629c\u62bc\u62c5\u4fdd\u4e13\u9879\u5ba1\u8ba1\u62a5\u544a.docx";
Packer.toBuffer(doc).then(function(buffer) {
  fs.writeFileSync(outPath2, buffer);
  console.log("\u62a5\u544a\u5df2\u751f\u6210\uff1a" + outPath2);
}).catch(function(e) {
  console.error("\u751f\u6210\u5931\u8d25\uff1a", e);
  process.exit(1);
});
