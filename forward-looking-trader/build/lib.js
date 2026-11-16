// Shared docx helpers for The Forward-Looking Trader documents.
if (process.env.FLT_HTML) { module.exports = require("./lib_html.js"); return; }
const fs = require("fs");
const d = require("docx");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, AlignmentType,
  HeadingLevel, BorderStyle, LevelFormat, Header, Footer, PageNumber, PageBreak, TableOfContents } = d;
const FONT = "Arial", NAVY = "1F3A5F", GREY = "555555", W = 9360;
const VERSION = "v1.1 · 16 Nov 2026 (UTC)";

function runs(text, base = {}) {
  // mini-markup: **bold**, *italic*
  const out = []; const re = /(\*\*[^*]+\*\*|\*[^*]+\*)/g; let last = 0, m;
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), font: FONT, ...base }));
    const t = m[0];
    if (t.startsWith("**")) out.push(new TextRun({ text: t.slice(2, -2), bold: true, font: FONT, ...base }));
    else out.push(new TextRun({ text: t.slice(1, -1), italics: true, font: FONT, ...base }));
    last = m.index + t.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), font: FONT, ...base }));
  return out;
}
const P = (t, o = {}) => new Paragraph({ children: runs(t, o.run), spacing: { after: 140, line: 288 }, ...o.para });
const H1 = t => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t, font: FONT })] });
const H2 = t => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: t, font: FONT })] });
const H3 = t => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun({ text: t, font: FONT })] });
const B = arr => arr.map(t => new Paragraph({ numbering: { reference: "bul", level: 0 }, children: runs(t), spacing: { after: 80, line: 276 } }));
let numInst = 0;
const N = arr => { numInst++; return arr.map(t => new Paragraph({ numbering: { reference: "num", level: 0, instance: numInst }, children: runs(t), spacing: { after: 80, line: 276 } })); };
const F = t => new Paragraph({ children: [new TextRun({ text: t, font: "Cambria Math", size: 22 })], indent: { left: 720 }, spacing: { before: 60, after: 140 } });
const NOTE = t => new Paragraph({ children: runs(t, { size: 18, color: GREY, italics: true }), spacing: { after: 140 } });
const PB = () => new Paragraph({ children: [new PageBreak()] });
const bd = { style: BorderStyle.SINGLE, size: 4, color: "BBBBBB" };
const borders = { top: bd, bottom: bd, left: bd, right: bd };
function cell(t, w, o = {}) {
  const paras = String(t).split("\n").map(line => new Paragraph({ children: runs(line, { size: 18, bold: o.bold, color: o.color }), spacing: { after: 40 }, alignment: o.align }));
  return new TableCell({ width: { size: w, type: WidthType.DXA }, borders, margins: { top: 60, bottom: 60, left: 100, right: 100 },
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, children: paras });
}
function T(headers, rows, widths) {
  const tot = widths.reduce((a, b) => a + b, 0); if (tot !== W) throw new Error("widths sum " + tot + " for " + headers.join("|"));
  const out = [];
  if (headers) out.push(new TableRow({ tableHeader: true, cantSplit: true, children: headers.map((h, i) => cell(h, widths[i], { bold: true, fill: NAVY, color: "FFFFFF" })) }));
  rows.forEach((r, k) => out.push(new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, widths[i], { fill: k % 2 ? "F4F6F9" : undefined })) })));
  return [new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: widths, rows: out }), new Paragraph({ children: [], spacing: { after: 120 } })];
}
function CALLOUT(title, lines, fill = "EEF3F9") {
  const paras = [new Paragraph({ children: [new TextRun({ text: title, bold: true, font: FONT, size: 20, color: NAVY })], spacing: { after: 60 } })]
    .concat([].concat(lines).map(l => new Paragraph({ children: runs(l, { size: 20 }), spacing: { after: 60, line: 276 } })));
  const side = { style: BorderStyle.SINGLE, size: 4, color: fill };
  return [new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: [W], rows: [new TableRow({ cantSplit: true, children: [new TableCell({
    width: { size: W, type: WidthType.DXA }, shading: { fill, type: ShadingType.CLEAR, color: "auto" }, margins: { top: 120, bottom: 120, left: 200, right: 200 },
    borders: { top: side, bottom: side, right: side, left: { style: BorderStyle.SINGLE, size: 24, color: NAVY } }, children: paras })] })] }),
    new Paragraph({ children: [], spacing: { after: 120 } })];
}
function TITLE(kicker, title, sub, meta) {
  return [
    new Paragraph({ children: [new TextRun({ text: kicker.toUpperCase(), font: FONT, size: 18, bold: true, color: GREY, characterSpacing: 40 })], spacing: { before: 1800, after: 200 } }),
    new Paragraph({ children: [new TextRun({ text: title, font: FONT, size: 56, bold: true, color: NAVY })], spacing: { after: 160 } }),
    new Paragraph({ children: [new TextRun({ text: sub, font: FONT, size: 26, color: "333333" })], spacing: { after: 400 }, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 12 } } }),
    ...meta.map(m => new Paragraph({ children: runs(m, { size: 20, color: GREY }), spacing: { after: 80 } })),
  ];
}
const TOC = () => [H1("Contents"), new TableOfContents("Contents", { hyperlink: true, headingStyleRange: "1-2" }), NOTE("If the table of contents is empty, right-click it in Word and choose Update Field."), PB()];
async function build(file, shortTitle, children) {
  const hs = (size, before, after, color) => ({ run: { font: FONT, size, bold: true, color }, paragraph: { spacing: { before, after }, keepNext: true } });
  const doc = new Document({
    creator: "The Forward-Looking Trader", title: shortTitle, features: { updateFields: true },
    styles: { default: { document: { run: { font: FONT, size: 21 } }, heading1: hs(32, 360, 160, NAVY), heading2: hs(25, 280, 120, NAVY), heading3: hs(22, 200, 100, "333333") } },
    numbering: { config: [
      { reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
      { reference: "num", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 360 } } } }] }] },
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, bottom: 1300, left: 1440, right: 1440 } } },
      headers: { default: new Header({ children: [new Paragraph({ children: [new TextRun({ text: "The Forward-Looking Trader  ·  " + shortTitle, font: FONT, size: 16, color: GREY })] })] }) },
      footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: VERSION + "  ·  Educational material, not investment advice  ·  Page ", font: FONT, size: 16, color: GREY }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: GREY })] })] }) },
      children: children.flat(Infinity),
    }],
  });
  fs.writeFileSync(file, await Packer.toBuffer(doc)); console.log("wrote", file);
}
module.exports = { P, H1, H2, H3, B, N, F, NOTE, PB, T, CALLOUT, TITLE, TOC, build, VERSION };
