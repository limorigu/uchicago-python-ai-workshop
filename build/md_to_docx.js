// ==========================================================================
// WARNING — BUILD OUTPUT IS NOT THE FINAL VERSION OF THE DOCUMENT
//
// The checked-in docs/workshop_design.docx received additional formatting
// tweaks (hyperlink styling, layout polish) made by hand, outside this
// script. Running this converter regenerates a document whose CONTENT is
// identical and fully reproducible from docs/workshop_design.md, but whose
// formatting will not match the final checked-in file. Do not overwrite
// the final .docx without re-applying those tweaks.
// ==========================================================================
// Minimal Markdown -> .docx converter tuned for docs/workshop_design.md (docx-js).
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType,
  AlignmentType, LevelFormat, ShadingType, BorderStyle, ExternalHyperlink, PageBreak,
} = require("docx");

const [, , inPath, outPath] = process.argv;
const md = fs.readFileSync(inPath, "utf8").split("\n");

const NAVY = "1F3A5F", GREY = "595959", CODE_BG = "F3F4F6", TABLE_HEAD = "E8EEF5";
const BODY_FONT = "Calibri", CODE_FONT = "Consolas";

// ---------- inline parsing ----------
function inline(text, base = {}) {
  const runs = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`|\[[^\]]+\]\([^)]+\)|https?:\/\/[^\s)]+)/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) runs.push(new TextRun({ text: text.slice(last, m.index), ...base }));
    const t = m[0];
    if (t.startsWith("**")) runs.push(new TextRun({ text: t.slice(2, -2), bold: true, ...base }));
    else if (t.startsWith("`")) runs.push(new TextRun({ text: t.slice(1, -1), font: CODE_FONT, size: (base.size || 21) - 1, shading: { type: ShadingType.CLEAR, fill: CODE_BG }, ...base }));
    else if (t.startsWith("*")) runs.push(new TextRun({ text: t.slice(1, -1), italics: true, ...base }));
    else if (t.startsWith("[")) {
      const mm = t.match(/\[([^\]]+)\]\(([^)]+)\)/);
      runs.push(new ExternalHyperlink({ link: mm[2], children: [new TextRun({ text: mm[1], style: "Hyperlink", ...base })] }));
    } else {
      const trail = t.match(/[.,;:]+$/) ? t.match(/[.,;:]+$/)[0] : "";
      const url = t.slice(0, t.length - trail.length);
      runs.push(new ExternalHyperlink({ link: url, children: [new TextRun({ text: url, style: "Hyperlink", ...base })] }));
      if (trail) runs.push(new TextRun({ text: trail, ...base }));
    }
    last = m.index + t.length;
  }
  if (last < text.length) runs.push(new TextRun({ text: text.slice(last), ...base }));
  return runs;
}

// ---------- block parsing ----------
const children = [];
let i = 0, firstH1 = true, numberedInstance = 0;
function para(text, opts = {}) { children.push(new Paragraph({ children: inline(text, opts.run || {}), spacing: { after: 120, line: 276 }, ...opts.p })); }

while (i < md.length) {
  const line = md[i];
  if (/^\s*$/.test(line)) { i++; continue; }
  if (/^---\s*$/.test(line)) { i++; continue; }
  if (line.startsWith("\\pagebreak")) { children.push(new Paragraph({ children: [new PageBreak()] })); i++; continue; }

  let h;
  if ((h = line.match(/^(#{1,3})\s+(.*)$/))) {
    const level = h[1].length, text = h[2];
    if (level === 1) {
      children.push(new Paragraph({ children: inline(text, { size: firstH1 ? 52 : 40, bold: true, color: NAVY }), spacing: { before: firstH1 ? 0 : 360, after: 120 } }));
      firstH1 = false;
    } else if (level === 2) {
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: inline(text, { size: 30, bold: true, color: NAVY }), spacing: { before: 360, after: 120 }, keepNext: true }));
    } else {
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: inline(text, { size: 24, bold: true, color: NAVY }), spacing: { before: 240, after: 80 }, keepNext: true }));
    }
    i++; continue;
  }

  if (line.startsWith("```")) {
    const buf = []; i++;
    while (i < md.length && !md[i].startsWith("```")) buf.push(md[i++]);
    i++;
    buf.forEach((l, k) => children.push(new Paragraph({
      children: [new TextRun({ text: l.replace(/ /g, " ") || " ", font: CODE_FONT, size: 17 })],
      shading: { type: ShadingType.CLEAR, fill: CODE_BG }, spacing: { after: 0, line: 240 },
      indent: { left: 200, right: 200 },
      border: k === 0 ? { top: { style: BorderStyle.SINGLE, size: 4, color: "D0D5DD", space: 4 } } : k === buf.length - 1 ? { bottom: { style: BorderStyle.SINGLE, size: 4, color: "D0D5DD", space: 4 } } : undefined,
    })));
    children.push(new Paragraph({ spacing: { after: 120 } }));
    continue;
  }

  if (line.startsWith("|")) {
    const rows = [];
    while (i < md.length && md[i].startsWith("|")) { rows.push(md[i]); i++; }
    const cells = rows.filter(r => !/^\|\s*-+/.test(r)).map(r => r.slice(1, r.lastIndexOf("|")).split("|").map(c => c.trim()));
    const ncol = cells[0].length;
    const total = 9360;
    const lens = Array.from({ length: ncol }, (_, c) => Math.sqrt(Math.max(...cells.map(r => (r[c] || "").length), 6)));
    const sum = lens.reduce((a, b) => a + b, 0);
    let widths = lens.map(l => Math.max(1500, Math.round((l / sum) * total)));
    const wsum = widths.reduce((a, b) => a + b, 0); widths = widths.map(w => Math.round(w * total / wsum));
    widths[ncol - 1] += total - widths.reduce((a, b) => a + b, 0);
    const table = new Table({
      width: { size: total, type: WidthType.DXA }, columnWidths: widths,
      rows: cells.map((r, ri) => new TableRow({
        tableHeader: ri === 0,
        children: r.map((c, ci) => new TableCell({
          width: { size: widths[ci], type: WidthType.DXA },
          shading: ri === 0 ? { type: ShadingType.CLEAR, fill: TABLE_HEAD } : undefined,
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          children: [new Paragraph({ children: inline(c, { size: 18, bold: ri === 0 }), spacing: { after: 0, line: 252 } })],
        })),
      })),
    });
    children.push(table);
    children.push(new Paragraph({ spacing: { after: 120 } }));
    continue;
  }

  let b;
  if ((b = line.match(/^(\s*)[-*]\s+(.*)$/))) {
    const level = Math.min(2, Math.floor(b[1].length / 2));
    children.push(new Paragraph({ children: inline(b[2]), numbering: { reference: "bullets", level }, spacing: { after: 60, line: 264 } }));
    i++; continue;
  }
  if ((b = line.match(/^(\s*)(\d+)\.\s+(.*)$/))) {
    // restart numbering when the item number is 1
    if (b[2] === "1") numberedInstance++;
    children.push(new Paragraph({ children: inline(b[3]), numbering: { reference: "numbers", level: 0, instance: numberedInstance }, spacing: { after: 60, line: 264 } }));
    i++; continue;
  }

  // paragraph: merge soft-wrapped lines
  const buf = [line]; i++;
  while (i < md.length && md[i].trim() && !/^(#|\||```|[-*]\s|\d+\.\s|---|\\pagebreak)/.test(md[i])) buf.push(md[i++]);
  para(buf.join(" "));
}

const doc = new Document({
  creator: "Workshop design", title: "From Excel to Python and AI — workshop design",
  styles: {
    default: { document: { run: { font: BODY_FONT, size: 21, color: "1F2937" } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 30, bold: true, color: NAVY, font: BODY_FONT }, paragraph: { outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 24, bold: true, color: NAVY, font: BODY_FONT }, paragraph: { outlineLevel: 1 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } },
        { level: 1, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1000, hanging: 270 } } } },
        { level: 2, format: LevelFormat.BULLET, text: "·", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1460, hanging: 270 } } } },
      ] },
      { reference: "numbers", levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 300 } } } },
      ] },
    ],
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1300, bottom: 1300, left: 1440, right: 1440 } } },
    headers: { default: { options: { children: [new Paragraph({ children: [new TextRun({ text: "From Excel to Python and AI · Workshop design · University of Chicago", size: 16, color: GREY })], alignment: AlignmentType.RIGHT })] } } },
    children,
  }],
});

Packer.toBuffer(doc).then(buf => { fs.writeFileSync(outPath, buf); console.log("wrote", outPath, buf.length, "bytes"); });
