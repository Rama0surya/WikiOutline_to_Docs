'use strict';
/**
 * docxBuilder.js — Outline Markdown → Word (.docx)
 *
 * Design goal: output that visually mirrors Outline wiki layout
 *   - Calibri / Calibri Light  (closest to Outline's Inter in Word)
 *   - 1.5× line-height body text
 *   - Clean heading hierarchy with bottom-border separators
 *   - Tables with blue-tinted header + alternating row shading
 *   - Code blocks: full-width shaded box with left accent border
 *   - Callout blocks (:::info :::warning :::tip :::danger :::note)
 *   - Task-list checkboxes  - [ ] / - [x]
 *   - Inline: **bold**, *italic*, ***b+i***, `code`, ~~strike~~, ==highlight==, [link](url)
 *   - Backslash-escape clean  (\* \_ \\ etc. never leak into output)
 *   - Images: real aspect-ratio via image-size, centred, max content-width
 *   - Blockquotes: nested depth-aware
 */

const {
  Document, Packer, Paragraph, TextRun, PageBreak,
  Table, TableRow, TableCell, ImageRun, ExternalHyperlink,
  AlignmentType, LevelFormat, BorderStyle, LineRuleType,
  WidthType, ShadingType, VerticalAlign, UnderlineType,
  TableLayoutType,
} = require('docx');
const axios           = require('axios');
const { imageSize }   = require('image-size');

// ─── Design tokens ────────────────────────────────────────────────────────────
// Calibri is bundled with every copy of Microsoft Office and is the closest
// freely-available match to Outline's Inter sans-serif look.
const FONT      = 'Calibri';
const FONT_MONO = 'Courier New';   // code / pre

// Half-points (1 pt = 2 units).  Calibri reads slightly larger than Arial
// so we keep sizes 1–2pt smaller than comparable Arial settings.
const SZ = {
  h1:    36,   // 18 pt  — document / page title
  h2:    30,   // 15 pt
  h3:    26,   // 13 pt
  h4:    24,   // 12 pt
  body:  22,   // 11 pt
  small: 18,   // 9 pt   — dates, captions, muted
  code:  20,   // 10 pt
};

// Colour palette — mirrors Outline's clean white-background UI
const C = {
  black:       '111111',   // headings
  body:        '1F2937',   // body text  (Outline's near-black)
  muted:       '6B7280',   // secondary / captions
  link:        '155EEF',   // hyperlinks (Outline blue)

  code_fg:     '1F2937',
  code_bg:     'F8F9FA',   // very light grey — code / pre background
  code_border: 'E5E7EB',   // code block left-border / frame

  bq_border:   'CBD5E1',   // blockquote left bar
  bq_bg:       'F8FAFC',
  bq_text:     '475569',

  hr:          'E5E7EB',   // horizontal rule

  tbl_head_bg: 'EFF6FF',   // table header fill  (Outline blue-50)
  tbl_head_fg: '1E3A5F',   // table header text
  tbl_border:  'BFDBFE',   // table outer border (blue-200)
  tbl_row_alt: 'F8FBFF',   // alternating body row  (very faint blue)

  // Callout palettes: [fill, left-bar, text]
  info:    ['EFF6FF', '3B82F6', '1E40AF'],
  warning: ['FFFBEB', 'F59E0B', '92400E'],
  tip:     ['F0FDF4', '22C55E', '166534'],
  danger:  ['FFF1F2', 'F43F5E', '9F1239'],
  note:    ['F5F3FF', '8B5CF6', '4C1D95'],
};

// A4 geometry (DXA; 1 cm ≈ 567, 1 in = 1440)
const PAGE_W    = 11906;
const MARGIN_H  = 1134;         // 2 cm left/right
const MARGIN_V  = 1134;         // 2 cm top/bottom
const CONTENT_W = PAGE_W - MARGIN_H * 2;   // 9638 DXA ≈ 16.9 cm

// Image sizing (EMU; 1 in = 914 400, 96 dpi)
const EMU_PER_PX  = 914400 / 96;
const MAX_IMG_EMU = Math.round((CONTENT_W / 1440) * 914400);

// ─── Utility ──────────────────────────────────────────────────────────────────
const decodeEntities = (s) => s
  .replace(/&amp;/g,  '&').replace(/&lt;/g,  '<')
  .replace(/&gt;/g,   '>').replace(/&quot;/g, '"')
  .replace(/&#39;/g,  "'").replace(/&nbsp;/g, ' ');

const indentLevel = (prefix, step = 2) =>
  Math.min(Math.floor(prefix.replace(/\t/g, '  ').length / step), 8);

// Strip ALL markdown syntax from a string, returning plain text
const stripMd = (s) =>
  s.replace(/[*_`~=]/g, '').replace(/\[([^\]]+)\]\([^)]+\)/g, '$1').trim();

// ─── Inline Markdown → TextRun[] ────────────────────────────────────────────
function parseInline(rawText, size = SZ.body, baseColor = C.body) {
  if (rawText == null) return [new TextRun({ text: '', font: FONT, size })];
  const input = String(rawText);
  if (!input.trim()) return [new TextRun({ text: input, font: FONT, size, color: baseColor })];

  // Replace backslash-escaped chars with a private-use marker so they are
  // never confused with Markdown syntax during tokenisation.
  const ESC  = '\uE000';
  const safe = input.replace(/\\([*_`~=\[\]()\\!#])/g, (_, ch) => ESC + ch + ESC);

  // Regex token order: bold+italic  >  bold  >  italic  >  code  >  strike
  //                    >  highlight  >  link
  const TOKEN =
    /\*\*\*(.+?)\*\*\*|___(.+?)___|(?<!\*)\*\*(?!\*)(.+?)(?<!\*)\*\*(?!\*)|__(.+?)__|(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)|(?<!_)_(?!_)(.+?)(?<!_)_(?!_)|`([^`]+)`|~~(.+?)~~|==(.+?)==|\[([^\]]*)\]\(([^)\s]+(?:\s+"[^"]*")?)\)/gs;

  const restore = (s) =>
    decodeEntities(s.replace(new RegExp(ESC + '(.)', 'g'), '$1'));

  const run = ({ txt, bold, italic, strike, code, hl, href }) => {
    const text = restore(txt || '');
    if (!text) return null;
    if (href) {
      const link = href.replace(/\s+"[^"]*"$/, '').trim();
      return new ExternalHyperlink({
        link,
        children: [new TextRun({
          text, font: FONT, size,
          color: C.link,
          underline: { type: UnderlineType.SINGLE, color: C.link },
        })],
      });
    }
    return new TextRun({
      text,
      font:    code ? FONT_MONO : FONT,
      size:    code ? SZ.code : size,
      bold:    bold   || false,
      italics: italic || false,
      strike:  strike || false,
      color:   code ? C.code_fg : baseColor,
      ...(hl   ? { highlight: 'yellow' } : {}),
      ...(code ? { shading: { type: ShadingType.CLEAR, fill: C.code_bg } } : {}),
    });
  };

  const runs = [];
  let last = 0, m;
  while ((m = TOKEN.exec(safe)) !== null) {
    if (m.index > last) { const r = run({ txt: safe.slice(last, m.index) }); if (r) runs.push(r); }
    /* eslint-disable no-multi-spaces */
    if      (m[1])  { const r = run({ txt: m[1],  bold: true, italic: true }); if (r) runs.push(r); }
    else if (m[2])  { const r = run({ txt: m[2],  bold: true, italic: true }); if (r) runs.push(r); }
    else if (m[3])  { const r = run({ txt: m[3],  bold: true               }); if (r) runs.push(r); }
    else if (m[4])  { const r = run({ txt: m[4],  bold: true               }); if (r) runs.push(r); }
    else if (m[5])  { const r = run({ txt: m[5],  italic: true             }); if (r) runs.push(r); }
    else if (m[6])  { const r = run({ txt: m[6],  italic: true             }); if (r) runs.push(r); }
    else if (m[7])  { const r = run({ txt: m[7],  code: true               }); if (r) runs.push(r); }
    else if (m[8])  { const r = run({ txt: m[8],  strike: true             }); if (r) runs.push(r); }
    else if (m[9])  { const r = run({ txt: m[9],  hl: true                 }); if (r) runs.push(r); }
    else if (m[10]) { const r = run({ txt: m[10], href: m[11]              }); if (r) runs.push(r); }
    /* eslint-enable no-multi-spaces */
    last = m.index + m[0].length;
  }
  if (last < safe.length) { const r = run({ txt: safe.slice(last) }); if (r) runs.push(r); }

  const filtered = runs.filter(Boolean);
  return filtered.length
    ? filtered
    : [new TextRun({ text: restore(safe), font: FONT, size, color: baseColor })];
}

// ─── Image fetch + aspect-ratio sizing ───────────────────────────────────────
async function fetchImage(src, outlineUrl, apiToken) {
  try {
    const url  = src.startsWith('http') ? src : (outlineUrl || '').replace(/\/$/, '') + src;
    const resp = await axios.get(url, {
      responseType: 'arraybuffer',
      timeout:      25000,
      headers:      apiToken ? { Authorization: `Bearer ${apiToken}` } : {},
    });
    const ct   = resp.headers['content-type'] || 'image/png';
    const type = ct.includes('jpeg') || ct.includes('jpg') ? 'jpg'
               : ct.includes('gif')  ? 'gif'
               : ct.includes('bmp')  ? 'bmp'
               : 'png';                         // webp → treated as png
    const data = Buffer.from(resp.data);

    // Detect real pixel dimensions to preserve aspect ratio
    let w = 800, h = 450;
    try { const d = imageSize(data); w = d.width || 800; h = d.height || 450; } catch {}

    const natW  = w * EMU_PER_PX;
    const natH  = h * EMU_PER_PX;
    const scale = natW > MAX_IMG_EMU ? MAX_IMG_EMU / natW : 1;

    return {
      data,
      type,
      width:  Math.round((natW * scale) / EMU_PER_PX),
      height: Math.round((natH * scale) / EMU_PER_PX),
    };
  } catch {
    return null;
  }
}

// ─── Heading builder ─────────────────────────────────────────────────────────
// Returns one or two Paragraphs (H1 gets a separator line below)
function buildHeading(text, level) {
  const szMap    = [SZ.h1, SZ.h2, SZ.h3, SZ.h4];
  // Space above/below in DXA (twips). Larger headings need more air above.
  const before   = [360, 300, 240, 180];
  const after    = [120, 100,  80,  60];

  const headingRuns = parseInlineForHeading(text, szMap[level - 1] || SZ.h4);

  const para = new Paragraph({
    spacing: { before: before[level - 1], after: after[level - 1], line: 320, lineRule: LineRuleType.AUTO },
    children: headingRuns,
  });

  if (level === 1) {
    // H1: thick bottom border — mirrors Outline's page-title separator
    return [
      para,
      new Paragraph({
        spacing: { before: 0, after: 160 },
        border:  { bottom: { style: BorderStyle.SINGLE, size: 8, color: C.hr, space: 1 } },
        children: [new TextRun('')],
      }),
    ];
  }
  if (level === 2) {
    // H2: subtle bottom border for section breaks
    return [
      para,
      new Paragraph({
        spacing: { before: 0, after: 100 },
        border:  { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.hr, space: 1 } },
        children: [new TextRun('')],
      }),
    ];
  }
  return [para];
}

// Heading inline parse — forces bold while respecting italic/code/link inside
function parseInlineForHeading(rawText, size) {
  const ESC  = '\uE000';
  const safe = String(rawText).replace(/\\([*_`~=\[\]()\\!#])/g, (_, ch) => ESC + ch + ESC);
  const TOKEN = /\*\*\*(.+?)\*\*\*|___(.+?)___|(?<!\*)\*\*(?!\*)(.+?)(?<!\*)\*\*(?!\*)|__(.+?)__|(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)|(?<!_)_(?!_)(.+?)(?<!_)_(?!_)|`([^`]+)`|~~(.+?)~~|\[([^\]]*)\]\(([^)\s]+(?:\s+"[^"]*")?)\)/gs;
  const restore = (s) => decodeEntities(s.replace(new RegExp(ESC + '(.)', 'g'), '$1'));

  const runs = [];
  let last = 0, m;
  while ((m = TOKEN.exec(safe)) !== null) {
    if (m.index > last)
      runs.push(new TextRun({ text: restore(safe.slice(last, m.index)), font: FONT, size, bold: true, color: C.black }));
    if      (m[1] || m[2]) runs.push(new TextRun({ text: restore(m[1] || m[2]), font: FONT, size, bold: true, italics: true, color: C.black }));
    else if (m[3] || m[4]) runs.push(new TextRun({ text: restore(m[3] || m[4]), font: FONT, size, bold: true, color: C.black }));
    else if (m[5] || m[6]) runs.push(new TextRun({ text: restore(m[5] || m[6]), font: FONT, size, bold: true, italics: true, color: C.black }));
    else if (m[7])         runs.push(new TextRun({ text: restore(m[7]),         font: FONT_MONO, size: SZ.code, bold: false, color: C.code_fg, shading: { type: ShadingType.CLEAR, fill: C.code_bg } }));
    else if (m[8])         runs.push(new TextRun({ text: restore(m[8]),         font: FONT, size, bold: true, strike: true, color: C.black }));
    else if (m[9]) {
      const link = m[10].replace(/\s+"[^"]*"$/, '').trim();
      runs.push(new ExternalHyperlink({ link, children: [new TextRun({ text: restore(m[9]), font: FONT, size, bold: true, color: C.link, underline: { type: UnderlineType.SINGLE, color: C.link } })] }));
    }
    last = m.index + m[0].length;
  }
  if (last < safe.length)
    runs.push(new TextRun({ text: restore(safe.slice(last)), font: FONT, size, bold: true, color: C.black }));

  return runs.length ? runs.filter(Boolean)
    : [new TextRun({ text: restore(safe), font: FONT, size, bold: true, color: C.black })];
}

// ─── Table builder ────────────────────────────────────────────────────────────
function buildTable(lines) {
  const splitRow = (line) =>
    line.replace(/^\s*\||\|\s*$/g, '').split('|').map(c => c.trim());

  const headers  = splitRow(lines[0]);
  const sepCells = splitRow(lines[1]);
  const colCount = headers.length;

  const aligns = sepCells.map(c => {
    const t = c.trim();
    if (/^:-+:$/.test(t)) return AlignmentType.CENTER;
    if (/^-+:$/.test(t))  return AlignmentType.RIGHT;
    return AlignmentType.LEFT;
  });

  // Distribute column widths evenly, last column absorbs rounding remainder
  const baseW    = Math.floor(CONTENT_W / colCount);
  const colWidths = Array(colCount).fill(baseW);
  colWidths[colCount - 1] += CONTENT_W - baseW * colCount;

  const bdr = (color, sz = 4) => ({ style: BorderStyle.SINGLE, size: sz, color });
  const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };

  const makeCell = (rawTxt, ci, isHeader, rowIndex) => {
    const altRow = !isHeader && rowIndex % 2 === 0;
    const fill   = isHeader ? C.tbl_head_bg : (altRow ? C.tbl_row_alt : 'FFFFFF');

    // Header cells: strip markdown from raw text, then make bold TextRun
    // Body cells: full inline parse (preserves bold/italic/code/link)
    const cellChildren = isHeader
      ? [new TextRun({ text: stripMd(rawTxt), font: FONT, size: SZ.body, bold: true, color: C.tbl_head_fg })]
      : parseInline(rawTxt, SZ.body, C.body);

    return new TableCell({
      width:         { size: colWidths[ci], type: WidthType.DXA },
      shading:       { fill, type: ShadingType.CLEAR },
      borders: {
        top:    isHeader ? bdr(C.tbl_border, 6) : bdr(C.code_border, 2),
        bottom: isHeader ? bdr(C.tbl_border, 6) : bdr(C.code_border, 2),
        left:   ci === 0             ? bdr(C.tbl_border, 4) : bdr(C.code_border, 2),
        right:  ci === colCount - 1  ? bdr(C.tbl_border, 4) : bdr(C.code_border, 2),
      },
      margins:       { top: 120, bottom: 120, left: 180, right: 180 },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({
        alignment: aligns[ci] || AlignmentType.LEFT,
        spacing:   { before: 0, after: 0, line: 276, lineRule: LineRuleType.AUTO },
        children:  cellChildren,
      })],
    });
  };

  const rows = [
    new TableRow({
      tableHeader: true,
      children: headers.map((h, ci) => makeCell(h, ci, true, 0)),
    }),
  ];
  for (let i = 2; i < lines.length; i++) {
    const trimmed = lines[i].trim();
    if (!trimmed || trimmed === '|') continue;
    const cells = splitRow(lines[i]);
    while (cells.length < colCount) cells.push('');
    const rowIndex = i - 1; // 0-based data row index
    rows.push(new TableRow({
      children: cells.slice(0, colCount).map((c, ci) => makeCell(c, ci, false, rowIndex)),
    }));
  }

  return new Table({
    width:        { size: CONTENT_W, type: WidthType.DXA },
    layout:       TableLayoutType.FIXED,
    columnWidths: colWidths,
    rows,
  });
}

// ─── Code block ───────────────────────────────────────────────────────────────
function buildCodeBlock(codeLines) {
  // Renders as a shaded box with a coloured left accent — identical to
  // Outline's code-block appearance.
  const makeCodePara = (text, isFirst, isLast) => new Paragraph({
    spacing: {
      before:   isFirst ? 40  : 0,
      after:    isLast  ? 40  : 0,
      line:     240,
      lineRule: LineRuleType.AUTO,
    },
    shading: { type: ShadingType.CLEAR, fill: C.code_bg },
    indent:  { left: 400, right: 400 },
    border:  {
      left:  { style: BorderStyle.SINGLE, size: 18, color: '4F81BD', space: 8 },
      ...(isFirst ? { top:    { style: BorderStyle.SINGLE, size: 4, color: C.code_border } } : {}),
      ...(isLast  ? { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.code_border } } : {}),
      right: { style: BorderStyle.SINGLE, size: 4, color: C.code_border },
    },
    children: [new TextRun({
      text:  text === '' ? '\u00A0' : text,  // preserve blank lines
      font:  FONT_MONO,
      size:  SZ.code,
      color: C.code_fg,
    })],
  });

  if (!codeLines.length) return [];
  const paras = codeLines.map((l, idx) =>
    makeCodePara(l, idx === 0, idx === codeLines.length - 1)
  );
  // Add spacing after block
  paras.push(new Paragraph({ spacing: { before: 0, after: 140 }, children: [new TextRun('')] }));
  return paras;
}

// ─── Callout block (:::info / :::warning / :::tip / :::danger / :::note) ─────
function buildCallout(type, lines) {
  const palette   = C[type] || C.info;
  const [bg, bar, fg] = palette;
  const ICONS     = { info: 'ℹ', warning: '⚠', tip: '✓', danger: '✕', note: '✎' };
  const label     = (ICONS[type] || '') + '  ' + type.toUpperCase();

  const paras = [];
  const border = (isFirst, isLast) => ({
    left:   { style: BorderStyle.SINGLE, size: 18, color: bar,             space: 10 },
    top:    isFirst ? { style: BorderStyle.SINGLE, size: 4, color: bar }   : { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    bottom: isLast  ? { style: BorderStyle.SINGLE, size: 4, color: bar }   : { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    right:  { style: BorderStyle.SINGLE, size: 4, color: bar },
  });

  // Label row
  paras.push(new Paragraph({
    spacing: { before: 160, after: 0, line: 276, lineRule: LineRuleType.AUTO },
    shading: { type: ShadingType.CLEAR, fill: bg },
    indent:  { left: 360, right: 360 },
    border:  border(true, lines.length === 0),
    children: [new TextRun({ text: label, font: FONT, size: SZ.small, bold: true, color: fg })],
  }));

  lines.forEach((l, idx) => {
    const isLast = idx === lines.length - 1;
    paras.push(new Paragraph({
      spacing: { before: 0, after: isLast ? 0 : 0, line: 320, lineRule: LineRuleType.AUTO },
      shading: { type: ShadingType.CLEAR, fill: bg },
      indent:  { left: 360, right: 360 },
      border:  border(false, isLast),
      children: parseInline(l || '\u00A0', SZ.body, fg),
    }));
  });

  paras.push(new Paragraph({ spacing: { before: 0, after: 160 }, children: [new TextRun('')] }));
  return paras;
}

// ─── Blockquote ───────────────────────────────────────────────────────────────
function buildBlockquote(text, depth = 1) {
  return new Paragraph({
    indent:  { left: depth * 500 },
    spacing: { before: 40, after: 40, line: 320, lineRule: LineRuleType.AUTO },
    shading: { type: ShadingType.CLEAR, fill: C.bq_bg },
    border:  { left: { style: BorderStyle.SINGLE, size: 16, color: C.bq_border, space: 12 } },
    children: parseInline(text, SZ.body, C.bq_text),
  });
}

// ─── Main Markdown → Paragraph list ─────────────────────────────────────────
async function buildParagraphs(md, title, opts, outlineUrl, apiToken) {
  const out = [];

  // ── Optional document title block ──────────────────────────────────────
  if (opts?.title && title) {
    out.push(new Paragraph({
      spacing: { before: 0, after: 80, line: 320, lineRule: LineRuleType.AUTO },
      children: [new TextRun({ text: title, font: FONT, size: SZ.h1, bold: true, color: C.black })],
    }));
    out.push(new Paragraph({
      spacing: { before: 0, after: 200 },
      border:  { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.hr, space: 1 } },
      children: [new TextRun('')],
    }));
  }

  if (opts?.date) {
    out.push(new Paragraph({
      spacing: { after: 180 },
      children: [new TextRun({
        text: 'Diekspor: ' + new Date().toLocaleDateString('id-ID', { day: '2-digit', month: 'long', year: 'numeric' }),
        font: FONT, size: SZ.small, italics: true, color: C.muted,
      })],
    }));
  }

  // ── Line-by-line parse ──────────────────────────────────────────────────
  const lines     = md.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  let i           = 0;
  let prevBlank   = false;

  while (i < lines.length) {
    const raw  = lines[i];
    const line = raw.trimEnd();

    // ── Blank lines — collapse runs of blanks into one small spacer ──────
    if (!line.trim()) {
      if (!prevBlank) {
        out.push(new Paragraph({
          spacing: { before: 0, after: 60 },
          children: [new TextRun('')],
        }));
      }
      prevBlank = true;
      i++; continue;
    }
    prevBlank = false;

    // ── Setext H1  (=====) ───────────────────────────────────────────────
    if (lines[i + 1]?.trim().match(/^=+$/) && line.trim()) {
      out.push(...buildHeading(line.trim(), 1));
      i += 2; continue;
    }
    // ── Setext H2  (-----) ───────────────────────────────────────────────
    if (lines[i + 1]?.trim().match(/^-+$/) && line.trim() && !line.startsWith('-')) {
      out.push(...buildHeading(line.trim(), 2));
      i += 2; continue;
    }

    // ── Callout  (:::type … :::) ─────────────────────────────────────────
    const calloutOpen = line.match(/^:::\s*(\w+)\s*$/);
    if (calloutOpen) {
      const type  = calloutOpen[1].toLowerCase();
      const body  = [];
      i++;
      while (i < lines.length && !lines[i].trim().match(/^:::\s*$/)) {
        body.push(lines[i]); i++;
      }
      out.push(...buildCallout(type, body));
      i++; continue;
    }

    // ── Fenced code block  (``` … ```) ───────────────────────────────────
    if (line.match(/^```/)) {
      const codeLines = [];
      i++;
      while (i < lines.length && !lines[i].trimEnd().match(/^```\s*$/)) {
        codeLines.push(lines[i]); i++;
      }
      out.push(...buildCodeBlock(codeLines));
      i++; continue;
    }

    // ── Indented code block  (4-space / tab) ─────────────────────────────
    if (line.match(/^(    |\t)/)) {
      const codeLines = [];
      while (i < lines.length && (lines[i].match(/^(    |\t)/) || !lines[i].trim())) {
        codeLines.push(lines[i].replace(/^(    |\t)/, '')); i++;
      }
      out.push(...buildCodeBlock(codeLines));
      continue;
    }

    // ── Markdown table ────────────────────────────────────────────────────
    if (line.startsWith('|') && lines[i + 1]?.match(/^\s*\|[\s|:=-]+\|/)) {
      const tLines = [line];
      let j = i + 1;
      while (j < lines.length && lines[j].trimEnd().startsWith('|')) {
        tLines.push(lines[j]); j++;
      }
      out.push(buildTable(tLines));
      out.push(new Paragraph({ spacing: { before: 0, after: 140 }, children: [new TextRun('')] }));
      i = j; continue;
    }

    // ── ATX headings  (# / ## / ### / ####) ──────────────────────────────
    const hMatch = line.match(/^(#{1,4})\s+(.+?)(?:\s+#+\s*)?$/);
    if (hMatch) {
      out.push(...buildHeading(hMatch[2].trim(), hMatch[1].length));
      i++; continue;
    }

    // ── Horizontal rule  (--- / *** / ___) ───────────────────────────────
    if (line.match(/^(\*{3,}|-{3,}|_{3,})\s*$/)) {
      out.push(new Paragraph({
        spacing: { before: 120, after: 120 },
        border:  { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.hr, space: 1 } },
        children: [new TextRun('')],
      }));
      i++; continue;
    }

    // ── Blockquote  (> …) ────────────────────────────────────────────────
    if (line.startsWith('>')) {
      while (i < lines.length && lines[i].trimEnd().startsWith('>')) {
        const depth = (lines[i].match(/^(>+)/)?.[1]?.length) || 1;
        out.push(buildBlockquote(lines[i].replace(/^>+\s?/, ''), depth));
        i++;
      }
      continue;
    }

    // ── Task list  (- [ ] / - [x]) ────────────────────────────────────────
    const taskMatch = line.match(/^(\s*)[-*+]\s+\[([ xX])\]\s+(.*)$/);
    if (taskMatch) {
      const done   = taskMatch[2].toLowerCase() === 'x';
      const lvl    = indentLevel(taskMatch[1], 2);
      const symbol = done ? '☑' : '☐';
      const textColor = done ? C.muted : C.body;
      out.push(new Paragraph({
        numbering: { reference: 'bullets', level: lvl },
        spacing:   { before: 0, after: 50, line: 320, lineRule: LineRuleType.AUTO },
        children:  [
          new TextRun({ text: symbol + ' ', font: FONT, size: SZ.body, color: textColor }),
          ...parseInline(taskMatch[3], SZ.body, textColor).map(r =>
            r instanceof TextRun
              ? new TextRun({ text: (r.root?.text ?? ''), font: FONT, size: SZ.body, color: textColor, strike: done, italics: false, bold: false })
              : r
          ),
        ],
      }));
      i++; continue;
    }

    // ── Bullet list  (-, *, +) ────────────────────────────────────────────
    const bulletMatch = line.match(/^(\s*)([-*+])\s+(.*)$/);
    if (bulletMatch) {
      const lvl = indentLevel(bulletMatch[1], 2);
      out.push(new Paragraph({
        numbering: { reference: 'bullets', level: lvl },
        spacing:   { before: 0, after: 50, line: 320, lineRule: LineRuleType.AUTO },
        children:  parseInline(bulletMatch[3], SZ.body, C.body),
      }));
      i++; continue;
    }

    // ── Numbered list  (1. / 1)) ──────────────────────────────────────────
    const numMatch = line.match(/^(\s*)\d+[.)]\s+(.*)$/);
    if (numMatch) {
      const lvl = indentLevel(numMatch[1], 3);
      out.push(new Paragraph({
        numbering: { reference: 'numbers', level: lvl },
        spacing:   { before: 0, after: 50, line: 320, lineRule: LineRuleType.AUTO },
        children:  parseInline(numMatch[2], SZ.body, C.body),
      }));
      i++; continue;
    }

    // ── Standalone image  (![alt](url)) ──────────────────────────────────
    const imgMatch = line.match(/^!\[([^\]]*)\]\(([^)\s]+)(?:\s+"[^"]*")?\)(\s*\{[^}]*\})?$/);
    if (imgMatch) {
      const img = await fetchImage(imgMatch[2].trim(), outlineUrl, apiToken);
      if (img) {
        out.push(new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing:   { before: 120, after: 120 },
          children:  [new ImageRun({
            data:           img.data,
            type:           img.type,
            transformation: { width: img.width, height: img.height },
          })],
        }));
      } else {
        // Fallback placeholder when image can't be fetched
        out.push(new Paragraph({
          spacing: { after: 80 },
          children: [new TextRun({
            text: `[Gambar: ${imgMatch[1] || imgMatch[2]}]`,
            font: FONT, size: SZ.body, italics: true, color: C.muted,
          })],
        }));
      }
      i++; continue;
    }

    // ── Normal paragraph ──────────────────────────────────────────────────
    // Detect trailing 2-space hard line break
    const hardBreak = raw.endsWith('  ');
    out.push(new Paragraph({
      spacing: { before: 0, after: hardBreak ? 0 : 100, line: 360, lineRule: LineRuleType.AUTO },
      children: parseInline(line.trimEnd(), SZ.body, C.body),
    }));
    i++;
  }

  return out;
}

// ─── Document assembler ───────────────────────────────────────────────────────
async function buildDocx(docs, opts, outlineUrl, apiToken) {

  // 9-level bullet list — Outline-style symbols
  const BULLETS = ['•', '◦', '▪', '▸', '–', '•', '◦', '▪', '▸'];
  const bulletLevels = BULLETS.map((ch, level) => ({
    level,
    format:    LevelFormat.BULLET,
    text:      ch,
    alignment: AlignmentType.LEFT,
    style: {
      paragraph: {
        indent:  { left: 480 + level * 360, hanging: 360 },
        spacing: { before: 0, after: 0, line: 320, lineRule: LineRuleType.AUTO },
      },
      run: { font: FONT, size: SZ.body, color: C.body },
    },
  }));

  // 9-level numbered list
  const numberLevels = Array.from({ length: 9 }, (_, level) => ({
    level,
    format:    LevelFormat.DECIMAL,
    text:      `%${level + 1}.`,
    alignment: AlignmentType.LEFT,
    style: {
      paragraph: {
        indent:  { left: 480 + level * 360, hanging: 360 },
        spacing: { before: 0, after: 0, line: 320, lineRule: LineRuleType.AUTO },
      },
      run: { font: FONT, size: SZ.body, color: C.body },
    },
  }));

  const allChildren = [];

  for (let d = 0; d < docs.length; d++) {
    if (d > 0) {
      allChildren.push(new Paragraph({
        spacing: { before: 0, after: 0 },
        pageBreakBefore: true,
        children: [new TextRun('')],
      }));
    }
    const paras = await buildParagraphs(
      docs[d].markdown || '',
      docs[d].title    || 'Untitled',
      opts,
      outlineUrl,
      apiToken,
    );
    allChildren.push(...paras);
  }

  const doc = new Document({
    creator:     'Outline Exporter',
    description: 'Exported from Outline Wiki',
    styles: {
      default: {
        document: {
          run: {
            font:  FONT,
            size:  SZ.body,
            color: C.body,
          },
          paragraph: {
            spacing: { line: 360, lineRule: LineRuleType.AUTO, after: 100 },
          },
        },
      },
      paragraphStyles: [
        {
          id: 'Normal', name: 'Normal', basedOn: 'Normal', next: 'Normal',
          run:       { font: FONT, size: SZ.body, color: C.body },
          paragraph: { spacing: { line: 360, lineRule: LineRuleType.AUTO, after: 100 } },
        },
      ],
    },
    numbering: {
      config: [
        { reference: 'bullets', levels: bulletLevels },
        { reference: 'numbers', levels: numberLevels },
      ],
    },
    sections: [{
      properties: {
        page: {
          size:   { width: PAGE_W, height: 16838 },
          margin: { top: MARGIN_V, right: MARGIN_H, bottom: MARGIN_V, left: MARGIN_H },
        },
      },
      children: allChildren,
    }],
  });

  return Packer.toBuffer(doc);
}

module.exports = { buildDocx };