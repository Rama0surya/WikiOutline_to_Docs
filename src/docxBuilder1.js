const {
  Document, Packer, Paragraph, TextRun, PageBreak,
  Table, TableRow, TableCell, ImageRun,
  AlignmentType, LevelFormat, BorderStyle, WidthType,
  ShadingType, VerticalAlign, HeadingLevel,
} = require('docx');
const axios = require('axios');

const FONT = 'Arial';
// docx size unit = half-points → pt * 2
const SZ = { h1: 32, h2: 28, h3: 24, body: 22, small: 18 };
// Page content width: A4 with 1 inch (1440 DXA) margins each side → 11906 - 2880 = 9026
const CONTENT_WIDTH_DXA = 9026;

// ─── Inline parser: **bold**, *italic*, `code` ─────────────────────────────
function parseInline(text, size = SZ.body, color = '333333') {
  const runs = [];
  const regex = /\*\*(.+?)\*\*|\*(.+?)\*|`([^`]+)`/g;
  let last = 0, m;
  while ((m = regex.exec(text)) !== null) {
    if (m.index > last)
      runs.push(new TextRun({ text: text.slice(last, m.index), font: FONT, size, color }));
    if (m[1] !== undefined)
      runs.push(new TextRun({ text: m[1], font: FONT, size, bold: true, color }));
    else if (m[2] !== undefined)
      runs.push(new TextRun({ text: m[2], font: FONT, size, italics: true, color }));
    else if (m[3] !== undefined)
      runs.push(new TextRun({ text: m[3], font: 'Courier New', size, color: '555555' }));
    last = m.index + m[0].length;
  }
  if (last < text.length)
    runs.push(new TextRun({ text: text.slice(last), font: FONT, size, color }));
  return runs.length ? runs : [new TextRun({ text, font: FONT, size, color })];
}

// ─── Image fetcher ──────────────────────────────────────────────────────────
async function fetchImage(src, outlineUrl, apiToken) {
  try {
    // Resolve relative URLs
    const url = src.startsWith('http') ? src : outlineUrl.replace(/\/$/, '') + src;
    const resp = await axios.get(url, {
      responseType: 'arraybuffer',
      timeout: 15000,
      headers: apiToken ? { Authorization: `Bearer ${apiToken}` } : {},
    });
    const contentType = resp.headers['content-type'] || 'image/png';
    const ext = contentType.includes('jpeg') || contentType.includes('jpg') ? 'jpg'
      : contentType.includes('gif') ? 'gif'
      : contentType.includes('bmp') ? 'bmp'
      : 'png';
    return { data: Buffer.from(resp.data), type: ext };
  } catch {
    return null;
  }
}

// ─── Markdown table parser ───────────────────────────────────────────────────
function parseTable(lines) {
  // lines[0] = header row, lines[1] = separator, lines[2+] = data rows
  const parseRow = (line) =>
    line.replace(/^\||\|$/g, '').split('|').map((c) => c.trim());

  const headers = parseRow(lines[0]);
  const colCount = headers.length;
  const colWidth = Math.floor(CONTENT_WIDTH_DXA / colCount);
  const columnWidths = Array(colCount).fill(colWidth);

  const border = { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' };
  const borders = { top: border, bottom: border, left: border, right: border };

  const makeRow = (cells, isHeader) =>
    new TableRow({
      children: cells.map((cellText, ci) =>
        new TableCell({
          borders,
          width: { size: columnWidths[ci], type: WidthType.DXA },
          shading: isHeader
            ? { fill: 'F0F0F0', type: ShadingType.CLEAR }
            : { fill: 'FFFFFF', type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          verticalAlign: VerticalAlign.CENTER,
          children: [
            new Paragraph({
              children: parseInline(cellText, SZ.body, '1a1a1a'),
              ...(isHeader ? { run: { bold: true } } : {}),
            }),
          ],
        })
      ),
    });

  const rows = [makeRow(headers, true)];
  for (let i = 2; i < lines.length; i++) {
    const cells = parseRow(lines[i]);
    // Pad or trim to colCount
    while (cells.length < colCount) cells.push('');
    rows.push(makeRow(cells.slice(0, colCount), false));
  }

  return new Table({
    width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
    columnWidths,
    rows,
  });
}

// ─── Main markdown → docx paragraph builder ─────────────────────────────────
async function buildParagraphs(md, title, opts, outlineUrl, apiToken) {
  const out = [];

  if (opts?.title && title) {
    out.push(new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ text: title, font: FONT, size: SZ.h1, bold: true, color: '000000' })],
    }));
  }
  if (opts?.date) {
    out.push(new Paragraph({
      spacing: { after: 80 },
      children: [new TextRun({
        text: 'Diekspor: ' + new Date().toLocaleDateString('id-ID'),
        font: FONT, size: SZ.small, italics: true, color: '999999',
      })],
    }));
  }

  const lines = md.split('\n');
  let i = 0;

  while (i < lines.length) {
    const line = lines[i].trimEnd();

    // ── Empty line ────────────────────────────────────────────────────────
    if (!line.trim()) {
      out.push(new Paragraph({ spacing: { after: 0 }, children: [new TextRun('')] }));
      i++; continue;
    }

    // ── Markdown table (detect header + separator pattern) ────────────────
    if (line.startsWith('|') && lines[i + 1]?.match(/^\|[\s|:-]+\|/)) {
      const tableLines = [line];
      let j = i + 1;
      while (j < lines.length && lines[j].startsWith('|')) {
        tableLines.push(lines[j]);
        j++;
      }
      out.push(parseTable(tableLines));
      // Add spacing after table
      out.push(new Paragraph({ spacing: { after: 80 }, children: [new TextRun('')] }));
      i = j; continue;
    }

    // ── Image: ![alt](url) ───────────────────────────────────────────────
    const imgMatch = line.match(/^!\[([^\]]*)\]\(([^)]+)\)/);
    if (imgMatch) {
      const imgData = await fetchImage(imgMatch[2], outlineUrl, apiToken);
      if (imgData) {
        // Cap image width to content width (~6.27 inches = 9026 DXA → EMU: 9026/1440*914400)
        const maxWidthEmu = Math.round((CONTENT_WIDTH_DXA / 1440) * 914400);
        out.push(new Paragraph({
          spacing: { after: 80 },
          children: [
            new ImageRun({
              data: imgData.data,
              type: imgData.type,
              transformation: { width: Math.round(maxWidthEmu / 9525), height: Math.round(maxWidthEmu / 9525 * 0.5625) },
            }),
          ],
        }));
      } else {
        // Fallback: show alt text
        out.push(new Paragraph({
          spacing: { after: 60 },
          children: [new TextRun({ text: `[Image: ${imgMatch[1] || imgMatch[2]}]`, font: FONT, size: SZ.body, color: '999999', italics: true })],
        }));
      }
      i++; continue;
    }

    // ── Headings ─────────────────────────────────────────────────────────
    if (line.startsWith('### ')) {
      out.push(new Paragraph({
        spacing: { before: 120, after: 60 },
        children: [new TextRun({ text: line.slice(4).trim(), font: FONT, size: SZ.h3, bold: true, color: '000000' })],
      }));
    } else if (line.startsWith('## ')) {
      out.push(new Paragraph({
        spacing: { before: 160, after: 80 },
        children: [new TextRun({ text: line.slice(3).trim(), font: FONT, size: SZ.h2, bold: true, color: '000000' })],
      }));
    } else if (line.startsWith('# ')) {
      out.push(new Paragraph({
        spacing: { before: 200, after: 100 },
        children: [new TextRun({ text: line.slice(2).trim(), font: FONT, size: SZ.h1, bold: true, color: '000000' })],
      }));

    // ── Bullet list ──────────────────────────────────────────────────────
    } else if (line.match(/^(\s*)[-*] /)) {
      const indent = line.match(/^(\s*)/)[1].length;
      const text = line.replace(/^\s*[-*] /, '');
      out.push(new Paragraph({
        numbering: { reference: 'bullets', level: Math.min(Math.floor(indent / 2), 8) },
        spacing: { after: 40 },
        children: parseInline(text, SZ.body, '333333'),
      }));

    // ── Numbered list ────────────────────────────────────────────────────
    } else if (line.match(/^(\s*)\d+\. /)) {
      const indent = line.match(/^(\s*)/)[1].length;
      const text = line.replace(/^\s*\d+\. /, '');
      out.push(new Paragraph({
        numbering: { reference: 'numbers', level: Math.min(Math.floor(indent / 2), 8) },
        spacing: { after: 40 },
        children: parseInline(text, SZ.body, '333333'),
      }));

    // ── Blockquote ────────────────────────────────────────────────────────
    } else if (line.startsWith('> ')) {
      out.push(new Paragraph({
        indent: { left: 720 },
        spacing: { after: 60 },
        children: [new TextRun({ text: line.slice(2).trim(), font: FONT, size: SZ.body, italics: true, color: '666666' })],
      }));

    // ── Horizontal rule ───────────────────────────────────────────────────
    } else if (line.startsWith('---') || line.startsWith('***')) {
      out.push(new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'CCCCCC', space: 1 } },
        children: [new TextRun('')],
      }));

    // ── Code block (``` ... ```) ──────────────────────────────────────────
    } else if (line.startsWith('```')) {
      const codeLines = [];
      i++;
      while (i < lines.length && !lines[i].startsWith('```')) {
        codeLines.push(lines[i]);
        i++;
      }
      for (const codeLine of codeLines) {
        out.push(new Paragraph({
          indent: { left: 720 },
          spacing: { after: 0, before: 0 },
          children: [new TextRun({ text: codeLine || ' ', font: 'Courier New', size: 18, color: '333333' })],
        }));
      }
      // spacing after code block
      out.push(new Paragraph({ spacing: { after: 80 }, children: [new TextRun('')] }));

    // ── Normal paragraph ──────────────────────────────────────────────────
    } else {
      out.push(new Paragraph({
        spacing: { after: 60 },
        children: parseInline(line, SZ.body, '333333'),
      }));
    }

    i++;
  }

  return out;
}

// ─── Main export function ────────────────────────────────────────────────────
async function buildDocx(docs, opts, outlineUrl, apiToken) {
  // Build bullet/number levels 0-8
  const bulletLevels = Array.from({ length: 9 }, (_, level) => ({
    level,
    format: LevelFormat.BULLET,
    text: ['•', '◦', '▪', '▸', '–', '•', '◦', '▪', '▸'][level],
    alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 720 + level * 360, hanging: 360 } } },
  }));
  const numberLevels = Array.from({ length: 9 }, (_, level) => ({
    level,
    format: LevelFormat.DECIMAL,
    text: `%${level + 1}.`,
    alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 720 + level * 360, hanging: 360 } } },
  }));

  const allChildren = [];

  for (let d = 0; d < docs.length; d++) {
    if (d > 0) {
      allChildren.push(new Paragraph({ children: [new PageBreak()] }));
    }
    const paragraphs = await buildParagraphs(
      docs[d].markdown, docs[d].title, opts, outlineUrl, apiToken
    );
    allChildren.push(...paragraphs);
  }

  const doc = new Document({
    numbering: {
      config: [
        { reference: 'bullets', levels: bulletLevels },
        { reference: 'numbers', levels: numberLevels },
      ],
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      children: allChildren,
    }],
  });

  return Packer.toBuffer(doc);
}

module.exports = { buildDocx };
