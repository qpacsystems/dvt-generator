#!/usr/bin/env node
// DVT generator — compact table format.
// Usage:
//   node build_dvt.js [data.json]
//   default: data.json = ./dvt_data.json
//
// Two modes:
//   - Default (procedure): cover, TOC, known-issues, external-ref pages, per-section
//     test tables, appendices A/B/C. Unchanged from upstream.
//   - Report (meta.mode === 'report'): the body is an ordered `reportSections` array
//     of free-form blocks (narrative / subcomponents / scope / units / tests /
//     anomalies). Additive and backward compatible.

const fs = require('fs');
const path = require('path');
const JSZip = require('jszip'); // transitive dep of docx
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, PageOrientation,
  Bookmark, InternalHyperlink,
  HeadingLevel, BorderStyle, WidthType, ShadingType, PageBreak, PageNumber,
  SimpleField, ImageRun,
} = require('docx');

// ============ CLI ============
const INPUT = process.argv[2] || 'dvt_data.json';
if (!fs.existsSync(INPUT)) {
  console.error(`Input not found: ${INPUT}`);
  console.error('Usage: node build_dvt.js [data.json]');
  process.exit(1);
}
let data = JSON.parse(fs.readFileSync(INPUT, 'utf-8'));
data.meta = data.meta || {};

// ---------- Today's date default ----------
if (!data.meta.date) {
  const t = new Date();
  data.meta.date = `${t.getMonth() + 1}/${t.getDate()}/${t.getFullYear()}`;
}

// ---------- Section letter abstraction ----------
let letterCode = 'A'.charCodeAt(0);
const codeToLetter = {};
for (const ext of data.meta.externalRefs || []) {
  if (ext.sectionLetter) {
    letterCode = Math.max(letterCode, ext.sectionLetter.charCodeAt(0) + 1);
  } else {
    ext.sectionLetter = String.fromCharCode(letterCode++);
  }
  if (ext.code) codeToLetter[ext.code] = ext.sectionLetter;
  codeToLetter[ext.sectionLetter] = ext.sectionLetter;
}
for (const sec of data.sections || []) {
  if (sec.id && /^[A-Z]$/.test(sec.id)) {
    sec.letter = sec.id;
    letterCode = Math.max(letterCode, sec.id.charCodeAt(0) + 1);
  } else if (!sec.letter) {
    sec.letter = String.fromCharCode(letterCode++);
  }
  sec.id = sec.letter;
  if (sec.code) codeToLetter[sec.code] = sec.letter;
  codeToLetter[sec.letter] = sec.letter;
  if (!sec.title) {
    sec.title = sec.name ? `Section ${sec.letter} — ${sec.name}` : `Section ${sec.letter}`;
  }
}
for (const step of data.steps || []) {
  if (step.ref) continue;
  if (step.section && step.subref != null) {
    const letter = codeToLetter[step.section];
    if (!letter) {
      console.error(`Step references unknown section code "${step.section}":`, step);
      process.exit(1);
    }
    step.ref = `${letter}.${step.subref}`;
  } else {
    console.error('Step is missing `ref` or `section`+`subref`:', step);
    process.exit(1);
  }
}

// ---------- Cross-reference expansion ----------
function expandCrossRefs(s, map) {
  return s.replace(/\[([A-Za-z_]\w*)((?:\.[\w.-]+)?)\]/g, (m, code, suffix) => {
    if (!(code in map)) return m;
    return map[code] + (suffix || '');
  });
}
function deepCrossRef(obj, map) {
  if (typeof obj === 'string') return expandCrossRefs(obj, map);
  if (Array.isArray(obj)) return obj.map((x) => deepCrossRef(x, map));
  if (obj && typeof obj === 'object') {
    const out = {};
    for (const k of Object.keys(obj)) out[k] = deepCrossRef(obj[k], map);
    return out;
  }
  return obj;
}
data = deepCrossRef(data, codeToLetter);

// ---------- Meta interpolation ({key} → meta.key) ----------
function interpolateStr(s, vars) {
  return s.replace(/\{(\w+)\}/g, (m, k) => (k in vars ? String(vars[k]) : m));
}
function deepInterpolate(obj, vars) {
  if (typeof obj === 'string') return interpolateStr(obj, vars);
  if (Array.isArray(obj)) return obj.map((x) => deepInterpolate(x, vars));
  if (obj && typeof obj === 'object') {
    const out = {};
    for (const k of Object.keys(obj)) out[k] = deepInterpolate(obj[k], vars);
    return out;
  }
  return obj;
}
const vars = {};
for (const [k, v] of Object.entries(data.meta)) {
  if (typeof v === 'string' || typeof v === 'number') vars[k] = v;
}
data = deepInterpolate(data, vars);
const meta = data.meta || {};

// ============ Report mode flag ============
const REPORT = meta.mode === 'report';

// ---------- Lint ----------
const allRefs = new Set((data.steps || []).map((s) => s.ref));
const validRefsAndPrefixes = new Set(allRefs);
for (const ref of allRefs) {
  const parts = ref.split('.');
  for (let i = parts.length - 1; i >= 1; i--) {
    validRefsAndPrefixes.add(parts.slice(0, i).join('.'));
  }
}
const lintWarnings = [];
const lintErrors = [];
function lintString(s, where) {
  for (const m of s.matchAll(/\[([A-Za-z_]\w*)(?:\.[\w.-]+)?\]/g)) {
    lintWarnings.push(`Unresolved cross-ref ${m[0]} in ${where}`);
  }
  for (const m of s.matchAll(/\b([A-Z])\.(\d+(?:\.\d+)*)\b/g)) {
    const ref = `${m[1]}.${m[2]}`;
    if (!validRefsAndPrefixes.has(ref)) {
      lintWarnings.push(`Unknown ref "${ref}" in ${where} — stale cross-ref?`);
    }
  }
  for (const m of s.matchAll(/\{(\w+)\}/g)) {
    lintWarnings.push(`Unresolved placeholder ${m[0]} in ${where}`);
  }
}
function walkLint(obj, path) {
  if (typeof obj === 'string') lintString(obj, path);
  else if (Array.isArray(obj)) obj.forEach((x, i) => walkLint(x, `${path}[${i}]`));
  else if (obj && typeof obj === 'object') for (const k of Object.keys(obj)) walkLint(obj[k], `${path}.${k}`);
}
walkLint(data, 'data');
{
  const seen = new Set();
  for (const step of data.steps || []) {
    if (seen.has(step.ref)) lintErrors.push(`Duplicate step ref: ${step.ref}`);
    seen.add(step.ref);
  }
}
for (const k of ['productName', 'headerLine']) {
  if (!meta[k]) lintWarnings.push(`meta.${k} is missing or empty`);
}
if (lintWarnings.length) {
  console.warn(`\nLint warnings (${lintWarnings.length}):`);
  for (const w of lintWarnings) console.warn('  ' + w);
  console.warn('');
}
if (lintErrors.length) {
  console.error(`Lint errors (${lintErrors.length}):`);
  for (const e of lintErrors) console.error('  ' + e);
  process.exit(1);
}

// ---------- Output name / dir ----------
function sanitizeForFilename(s) {
  return String(s)
    .replace(/[—–]/g, '-')
    .replace(/[\/\\:*?"<>|]/g, '_')
    .replace(/\s+/g, ' ')
    .trim();
}
function resolveOutName() {
  if (meta.outputFilename) return meta.outputFilename;
  if (meta.headerLine) return sanitizeForFilename(meta.headerLine) + '.docx';
  return path.basename(INPUT, '.json') + '.docx';
}
function resolveOutDir() {
  const inputDir = path.dirname(path.resolve(INPUT));
  if (!meta.outputDir) return inputDir;
  return path.isAbsolute(meta.outputDir) ? meta.outputDir : path.resolve(inputDir, meta.outputDir);
}
const OUT_NAME = resolveOutName();
const OUT_DIR = resolveOutDir();
const OUT = path.join(OUT_DIR, OUT_NAME);
if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });

// ============ Page / layout constants ============
const PAGE = { width: 12240, height: 15840, margin: 1008 }; // US Letter, 0.7"
const CONTENT_W = PAGE.width - 2 * PAGE.margin; // 10224 DXA

const COL_REF = 720, COL_HEADER = 1700, COL_BRIEF = 3200, COL_CRIT = 2400, COL_MEAS = 1500, COL_PF = 704;
const ROW_COLS = [COL_REF, COL_HEADER, COL_BRIEF, COL_CRIT, COL_MEAS, COL_PF];

const border = { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 60, bottom: 60, left: 100, right: 100 };

const HEADER_FILL = 'D9E2F3';
const ALT_FILL = 'F5F7FB';
const SECTION_FILL = 'B4C7E7';

function getImageSize(buf) {
  if (buf.length > 24 && buf[0] === 0x89 && buf[1] === 0x50) return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
  if (buf.length > 4 && buf[0] === 0xFF && buf[1] === 0xD8) {
    let o = 2;
    while (o < buf.length - 8) {
      if (buf[o] !== 0xFF) { o++; continue; }
      const marker = buf[o + 1];
      if (marker >= 0xC0 && marker <= 0xCF && marker !== 0xC4 && marker !== 0xC8 && marker !== 0xCC) return { h: buf.readUInt16BE(o + 5), w: buf.readUInt16BE(o + 7) };
      o += 2 + buf.readUInt16BE(o + 2);
    }
  }
  return { w: 480, h: 360 };
}

// ============ Bookmark IDs ============
const refToBkm = (ref) => 'test_' + String(ref).replace(/[^A-Za-z0-9]/g, '_');
const refToPfBkm = (ref) => 'pf_' + String(ref).replace(/[^A-Za-z0-9]/g, '_');

// ============ Helpers ============
const cell = (text, widthDxa, opts = {}) => {
  const children = Array.isArray(text)
    ? text
    : [new Paragraph({
        spacing: { before: 20, after: 20 },
        alignment: opts.align || AlignmentType.LEFT,
        children: [new TextRun({ text: String(text == null ? '' : text), bold: !!opts.bold, size: opts.size || 18 })],
      })];
  return new TableCell({
    borders,
    width: { size: widthDxa, type: WidthType.DXA },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: cellMargins,
    children,
  });
};

const headerRow = (labels) => new TableRow({
  tableHeader: true,
  children: labels.map((label, i) => cell(label, ROW_COLS[i], { bold: true, fill: HEADER_FILL, align: AlignmentType.CENTER })),
});

const stepRow = (s, alt) => {
  const fill = alt ? ALT_FILL : undefined;
  const refPara = new Paragraph({
    spacing: { before: 20, after: 20 },
    alignment: AlignmentType.LEFT,
    children: [new Bookmark({ id: refToBkm(s.ref), children: [new TextRun({ text: s.ref, bold: true, size: 18 })] })],
  });
  const pfPara = new Paragraph({
    spacing: { before: 20, after: 20 },
    alignment: AlignmentType.CENTER,
    children: [new Bookmark({ id: refToPfBkm(s.ref), children: [new TextRun({ text: '☐', size: 22 })] })],
  });
  return new TableRow({
    cantSplit: true,
    children: [
      new TableCell({ borders, width: { size: COL_REF, type: WidthType.DXA }, shading: fill ? { fill, type: ShadingType.CLEAR } : undefined, margins: cellMargins, children: [refPara] }),
      cell(s.header, COL_HEADER, { fill, bold: true }),
      cell(s.brief, COL_BRIEF, { fill }),
      cell(s.criteria, COL_CRIT, { fill }),
      cell(s.measure || '', COL_MEAS, { fill }),
      new TableCell({ borders, width: { size: COL_PF, type: WidthType.DXA }, shading: fill ? { fill, type: ShadingType.CLEAR } : undefined, margins: cellMargins, children: [pfPara] }),
    ],
  });
};

const sectionTable = (steps) => new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: ROW_COLS,
  rows: [
    headerRow(['Ref', 'Header', 'Brief', 'Criteria', 'Measure / Result', 'P/F']),
    ...steps.map((s, i) => stepRow(s, i % 2 === 1)),
  ],
});

// ============ Cover sheet ============
const COVER_L = 3600;
const COVER_R = CONTENT_W - COVER_L;
const coverCell = (label, value, isLabel) => new TableCell({
  borders,
  width: { size: isLabel ? COVER_L : COVER_R, type: WidthType.DXA },
  shading: isLabel ? { fill: 'F2F2F2', type: ShadingType.CLEAR } : undefined,
  margins: { top: 100, bottom: 100, left: 140, right: 140 },
  children: [new Paragraph({ children: [new TextRun({ text: isLabel ? label : value, bold: isLabel, size: 20 })] })],
});
const coverRow = (label, value) => new TableRow({ children: [coverCell(label, '', true), coverCell('', value || '', false)] });
const coverTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [COVER_L, COVER_R],
  rows: (meta.coverRows || []).map((r) => coverRow(r.label, r.value)),
});

const coverPara = [];
if (meta.title) coverPara.push(new Paragraph({ spacing: { before: 100, after: 100 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: meta.title, bold: true, size: 32 })] }));
if (meta.productName) coverPara.push(new Paragraph({ spacing: { before: 100, after: 60 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: meta.productName, bold: true, size: 28 })] }));
if (meta.subtitle) coverPara.push(new Paragraph({ spacing: { before: 60, after: 60 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: meta.subtitle, size: 22 })] }));
if (meta.descriptor) coverPara.push(new Paragraph({ spacing: { before: 60, after: 240 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: meta.descriptor, italics: true, size: 22 })] }));
coverPara.push(coverTable);

// ============ Default-mode tables (built lazily; only used when !REPORT) ============
const KI_W = [4400, 2800, 3024];
const knownIssuesTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: KI_W,
  rows: [
    new TableRow({ tableHeader: true, children: ['Defect', 'Suite test', 'DVT step'].map((label, i) => cell(label, KI_W[i], { bold: true, fill: HEADER_FILL })) }),
    ...(data.knownIssues || []).map((ki, i) => new TableRow({
      cantSplit: true,
      children: [
        cell(ki.defect, KI_W[0], { fill: i % 2 ? ALT_FILL : undefined }),
        cell(ki.suite, KI_W[1], { fill: i % 2 ? ALT_FILL : undefined }),
        cell(ki.dvt, KI_W[2], { fill: i % 2 ? ALT_FILL : undefined }),
      ],
    })),
  ],
});

const EQ_W = [3000, 4200, 3024];
const equipmentTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: EQ_W,
  rows: [
    new TableRow({ tableHeader: true, children: ['Equipment', 'Action / use', 'Spec'].map((label, i) => cell(label, EQ_W[i], { bold: true, fill: HEADER_FILL })) }),
    ...(data.equipment || []).map((eq, i) => new TableRow({
      cantSplit: true,
      children: [
        cell(eq.name, EQ_W[0], { fill: i % 2 ? ALT_FILL : undefined, bold: true }),
        cell(eq.action, EQ_W[1], { fill: i % 2 ? ALT_FILL : undefined }),
        cell(eq.spec, EQ_W[2], { fill: i % 2 ? ALT_FILL : undefined }),
      ],
    })),
  ],
});

const RG_CAT_W = [800, 4200, 5224];
const releaseCatTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: RG_CAT_W,
  rows: [
    new TableRow({ tableHeader: true, children: ['Cat', 'Change type', 'Required sections'].map((label, i) => cell(label, RG_CAT_W[i], { bold: true, fill: HEADER_FILL })) }),
    ...((data.releaseGate && data.releaseGate.categories) || []).map((c, i) => new TableRow({
      cantSplit: true,
      children: [
        cell(c.id, RG_CAT_W[0], { fill: i % 2 ? ALT_FILL : undefined, bold: true, align: AlignmentType.CENTER }),
        cell(c.name, RG_CAT_W[1], { fill: i % 2 ? ALT_FILL : undefined }),
        cell(c.required, RG_CAT_W[2], { fill: i % 2 ? ALT_FILL : undefined }),
      ],
    })),
  ],
});

const matrixHeaders = (meta.appendixB && meta.appendixB.matrixHeaders) || [];
const defaultMatrixWidths = () => {
  if (matrixHeaders.length === 0) return [];
  const w = Math.floor(CONTENT_W / matrixHeaders.length);
  const arr = Array(matrixHeaders.length).fill(w);
  arr[arr.length - 1] = CONTENT_W - w * (matrixHeaders.length - 1);
  return arr;
};
const RG_MATRIX_W = (meta.appendixB && meta.appendixB.matrixColumnWidths) || defaultMatrixWidths();
const matrixRows = ((data.releaseGate && data.releaseGate.matrix) || []);
const releaseMatrixTable = matrixHeaders.length > 0 ? new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: RG_MATRIX_W,
  rows: [
    new TableRow({ tableHeader: true, children: matrixHeaders.map((label, i) => cell(label, RG_MATRIX_W[i], { bold: true, fill: HEADER_FILL, align: AlignmentType.CENTER })) }),
    ...matrixRows.map((r, i) => new TableRow({
      cantSplit: true,
      children: (r.cells || []).map((c, j) => cell(c, RG_MATRIX_W[j] || RG_MATRIX_W[RG_MATRIX_W.length - 1], { fill: i % 2 ? ALT_FILL : undefined, bold: j === 0, align: (j > 0 && j < matrixHeaders.length - 1) ? AlignmentType.CENTER : AlignmentType.LEFT })),
    })),
  ],
}) : null;

// ============ Default-mode TOC ============
const TOC_W = [1000, 7700, 800, 724];
function tocSectionDividerRow(text, anchorId) {
  return new TableRow({
    cantSplit: true,
    children: [new TableCell({
      borders,
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnSpan: 4,
      shading: { fill: SECTION_FILL, type: ShadingType.CLEAR },
      margins: { top: 40, bottom: 40, left: 100, right: 100 },
      children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [new InternalHyperlink({ anchor: anchorId, children: [new TextRun({ text, bold: true, size: 18, color: '1F3864' })] })] })],
    })],
  });
}
function tocTestRow(step) {
  const linkRun = (text) => new InternalHyperlink({ anchor: refToBkm(step.ref), children: [new TextRun({ text, style: 'Hyperlink', color: '0066CC', size: 16 })] });
  return new TableRow({
    cantSplit: true,
    children: [
      new TableCell({ borders, width: { size: TOC_W[0], type: WidthType.DXA }, margins: { top: 20, bottom: 20, left: 80, right: 80 }, children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [linkRun(step.ref)] })] }),
      new TableCell({ borders, width: { size: TOC_W[1], type: WidthType.DXA }, margins: { top: 20, bottom: 20, left: 80, right: 80 }, children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [linkRun(step.header)] })] }),
      new TableCell({ borders, width: { size: TOC_W[2], type: WidthType.DXA }, margins: { top: 20, bottom: 20, left: 80, right: 80 }, children: [new Paragraph({ spacing: { before: 0, after: 0 }, alignment: AlignmentType.CENTER, children: [new SimpleField('PAGEREF ' + refToBkm(step.ref) + ' \\h')] })] }),
      new TableCell({ borders, width: { size: TOC_W[3], type: WidthType.DXA }, margins: { top: 20, bottom: 20, left: 80, right: 80 }, children: [new Paragraph({ spacing: { before: 0, after: 0 }, alignment: AlignmentType.CENTER, children: [new SimpleField('REF ' + refToPfBkm(step.ref) + ' \\h')] })] }),
    ],
  });
}
function tocRefPageRow(sectionLetter, sectionTitle) {
  return new TableRow({
    cantSplit: true,
    children: [
      new TableCell({ borders, width: { size: TOC_W[0], type: WidthType.DXA }, margins: { top: 20, bottom: 20, left: 80, right: 80 }, children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: 'Sec ' + sectionLetter, size: 16 })] })] }),
      new TableCell({ borders, width: { size: TOC_W[1], type: WidthType.DXA }, margins: { top: 20, bottom: 20, left: 80, right: 80 }, children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [new InternalHyperlink({ anchor: 'sec_' + sectionLetter, children: [new TextRun({ text: sectionTitle + ' (external reference)', style: 'Hyperlink', italics: true, color: '0066CC', size: 16 })] })] })] }),
      new TableCell({ borders, width: { size: TOC_W[2], type: WidthType.DXA }, margins: { top: 20, bottom: 20, left: 80, right: 80 }, children: [new Paragraph({ spacing: { before: 0, after: 0 }, alignment: AlignmentType.CENTER, children: [new SimpleField('PAGEREF sec_' + sectionLetter + ' \\h')] })] }),
      new TableCell({ borders, width: { size: TOC_W[3], type: WidthType.DXA }, margins: { top: 20, bottom: 20, left: 80, right: 80 }, children: [new Paragraph({ spacing: { before: 0, after: 0 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: '—', size: 16 })] })] }),
    ],
  });
}
const tocHeaderRow = new TableRow({
  tableHeader: true,
  children: ['Ref', 'Test header', 'Page', 'P/F'].map((label, i) => cell(label, TOC_W[i], { bold: true, fill: HEADER_FILL, align: AlignmentType.CENTER })),
});
const tocRows = [tocHeaderRow];
for (const ref of (meta.externalRefs || [])) {
  tocRows.push(tocRefPageRow(ref.sectionLetter, `Section ${ref.sectionLetter} — ${ref.title}`));
}
for (const section of (data.sections || [])) {
  tocRows.push(tocSectionDividerRow(section.title, 'sec_' + section.id));
  const sectionSteps = (data.steps || []).filter((s) => s.ref.startsWith(section.id + '.'));
  for (const step of sectionSteps) tocRows.push(tocTestRow(step));
}
const tocTable = new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: TOC_W, rows: tocRows });

// ============ Compose document ============
const children = [];
children.push(...coverPara);

if (!REPORT) {
  // ---- Default (procedure) body ----
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 0, after: 80 }, children: [new TextRun({ text: 'Contents', bold: true, size: 28 })] }));
  children.push(new Paragraph({ spacing: { before: 40, after: 120 }, children: [new TextRun({ text: 'Click any row to jump to the test. Page numbers and P/F mirror update on field refresh — Word: Ctrl+A then F9; LibreOffice: Tools → Update → Fields.', italics: true, size: 16 })] }));
  children.push(tocTable);

  const ki = meta.knownIssues || {};
  if ((data.knownIssues || []).length > 0 || ki.heading) {
    children.push(new Paragraph({ children: [new PageBreak()] }));
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 0, after: 120 }, children: [new TextRun({ text: ki.heading || 'Known Issues', bold: true, size: 28 })] }));
    if (ki.intro) children.push(new Paragraph({ spacing: { before: 60, after: 120 }, children: [new TextRun({ text: ki.intro, italics: true, size: 18 })] }));
    children.push(knownIssuesTable);
  }

  const ATT_W = [Math.round(CONTENT_W * 0.55), 0];
  ATT_W[1] = CONTENT_W - ATT_W[0];
  const externalRefPage = (opts) => {
    const letter = opts.sectionLetter;
    children.push(new Paragraph({ children: [new PageBreak()] }));
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 0, after: 120 }, children: [new Bookmark({ id: 'sec_' + letter, children: [new TextRun({ text: `Section ${letter} — ${opts.title} (External Reference)`, bold: true, size: 28 })] })] }));
    if (opts.intro) children.push(new Paragraph({ spacing: { before: 60, after: 160 }, children: [new TextRun({ text: opts.intro, italics: true, size: 20 })] }));
    children.push(new Paragraph({ spacing: { before: 200, after: 80 }, children: [new TextRun({ text: 'Reference', bold: true, size: 22 })] }));
    children.push(new Paragraph({ spacing: { before: 40, after: 40 }, children: [new TextRun({ text: `${opts.title} — latest released revision: `, size: 20 }), new TextRun({ text: opts.linkPlaceholder || '<Link>', bold: true, color: '0066CC', size: 20 })] }));
    if (opts.passCriteria) {
      children.push(new Paragraph({ spacing: { before: 200, after: 80 }, children: [new TextRun({ text: `Pass criteria for Section ${letter}`, bold: true, size: 22 })] }));
      children.push(new Paragraph({ spacing: { before: 40, after: 40 }, children: [new TextRun({ text: opts.passCriteria, size: 20 })] }));
    }
    if (opts.attestHeaders && opts.attestHeaders.length === 2) {
      children.push(new Paragraph({ spacing: { before: 200, after: 80 }, children: [new TextRun({ text: 'Tester attestation', bold: true, size: 22 })] }));
      children.push(new Table({
        width: { size: CONTENT_W, type: WidthType.DXA },
        columnWidths: ATT_W,
        rows: [
          new TableRow({ tableHeader: true, children: opts.attestHeaders.map((label, i) => cell(label, ATT_W[i], { bold: true, fill: HEADER_FILL })) }),
          new TableRow({ cantSplit: true, children: [cell('', ATT_W[0]), cell('', ATT_W[1])] }),
        ],
      }));
    }
  };
  for (const ref of (meta.externalRefs || [])) externalRefPage(ref);

  for (const section of (data.sections || [])) {
    const sectionSteps = (data.steps || []).filter((s) => s.ref.startsWith(section.id + '.'));
    children.push(new Paragraph({ children: [new PageBreak()] }));
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 0, after: 120 }, children: [new Bookmark({ id: 'sec_' + section.id, children: [new TextRun({ text: section.title, bold: true, size: 28 })] })] }));
    if (section.intro) children.push(new Paragraph({ spacing: { before: 60, after: 160 }, children: [new TextRun({ text: section.intro, italics: true, size: 18 })] }));
    children.push(sectionTable(sectionSteps));
  }

  const appA = meta.appendixA || {};
  if ((data.equipment || []).length > 0) {
    children.push(new Paragraph({ children: [new PageBreak()] }));
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 0, after: 120 }, children: [new TextRun({ text: appA.heading || 'Appendix A — Equipment', bold: true, size: 28 })] }));
    if (appA.intro) children.push(new Paragraph({ spacing: { before: 60, after: 160 }, children: [new TextRun({ text: appA.intro, italics: true, size: 18 })] }));
    children.push(equipmentTable);
  }

  const appB = meta.appendixB || {};
  if (((data.releaseGate && data.releaseGate.categories) || []).length > 0 || matrixRows.length > 0) {
    children.push(new Paragraph({ children: [new PageBreak()] }));
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 0, after: 120 }, children: [new TextRun({ text: appB.heading || 'Appendix B — Release-Gate Matrix', bold: true, size: 28 })] }));
    if (appB.intro) children.push(new Paragraph({ spacing: { before: 60, after: 120 }, children: [new TextRun({ text: appB.intro, italics: true, size: 18 })] }));
    if (((data.releaseGate && data.releaseGate.categories) || []).length > 0) {
      children.push(new Paragraph({ spacing: { before: 120, after: 80 }, children: [new TextRun({ text: appB.categoriesTitle || 'Change-severity categories', bold: true, size: 22 })] }));
      children.push(releaseCatTable);
    }
    if (releaseMatrixTable && matrixRows.length > 0) {
      children.push(new Paragraph({ spacing: { before: 200, after: 80 }, children: [new TextRun({ text: appB.matrixTitle || 'Compatibility matrix', bold: true, size: 22 })] }));
      children.push(releaseMatrixTable);
    }
  }

  const appC = meta.appendixC || {};
  if ((data.elaborations || []).length > 0) {
    children.push(new Paragraph({ children: [new PageBreak()] }));
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 0, after: 120 }, children: [new TextRun({ text: appC.heading || 'Appendix C — Elaborations and Rationale', bold: true, size: 28 })] }));
    if (appC.intro) children.push(new Paragraph({ spacing: { before: 60, after: 160 }, children: [new TextRun({ text: appC.intro, italics: true, size: 18 })] }));
    for (const e of data.elaborations) {
      children.push(new Paragraph({ spacing: { before: 140, after: 40 }, children: [new TextRun({ text: e.ref, bold: true, size: 20 })] }));
      children.push(new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: e.text, size: 18 })] }));
    }
  }
}

// ============ Report-mode body (ordered free-form sections) ============
if (REPORT) {
  const RSEC = data.reportSections || [];
  const para = (text, o = {}) => new Paragraph({ spacing: { before: o.before != null ? o.before : 40, after: o.after != null ? o.after : 80 }, children: [new TextRun({ text: String(text == null ? '' : text), size: o.size || 20, bold: !!o.bold, italics: !!o.italics })] });
  const subhead = (text, before) => new Paragraph({ spacing: { before: before != null ? before : 200, after: 60 }, children: [new TextRun({ text, bold: true, size: 22 })] });
  const genTable = (headers, rows, widths, opts = {}) => {
    const n = headers.length;
    let w = (widths && widths.length === n) ? widths.slice() : Array(n).fill(Math.floor(CONTENT_W / n));
    if (!(widths && widths.length === n)) w[n - 1] = CONTENT_W - w.slice(0, -1).reduce((a, b) => a + b, 0);
    const hrow = new TableRow({ tableHeader: true, children: headers.map((l, j) => cell(l, w[j], { bold: true, fill: HEADER_FILL, align: opts.centerHeader ? AlignmentType.CENTER : AlignmentType.LEFT })) });
    const body = (rows || []).map((r, ri) => new TableRow({ cantSplit: true, children: r.map((c, j) => cell(c, w[j], { fill: ri % 2 ? ALT_FILL : undefined, bold: j === 0 && opts.boldFirst })) }));
    return new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: w, rows: [hrow, ...body] });
  };

  // Contents
  children.push(new Paragraph({ children: [new PageBreak()] }));
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 0, after: 80 }, children: [new TextRun({ text: 'Contents', bold: true, size: 28 })] }));
  children.push(new Paragraph({ spacing: { before: 40, after: 120 }, children: [new TextRun({ text: 'Click a section to jump to it. Page numbers update on field refresh — Word: Ctrl+A then F9; LibreOffice: Tools → Update → Fields.', italics: true, size: 16 })] }));
  {
    const TW = [9200, 1024];
    const rows = [new TableRow({ tableHeader: true, children: [cell('Section', TW[0], { bold: true, fill: HEADER_FILL }), cell('Page', TW[1], { bold: true, fill: HEADER_FILL, align: AlignmentType.CENTER })] })];
    RSEC.forEach((s, i) => {
      rows.push(new TableRow({ cantSplit: true, children: [
        new TableCell({ borders, width: { size: TW[0], type: WidthType.DXA }, margins: cellMargins, children: [new Paragraph({ children: [new InternalHyperlink({ anchor: 'rsec_' + i, children: [new TextRun({ text: s.name, style: 'Hyperlink', color: '0066CC', size: 18 })] })] })] }),
        new TableCell({ borders, width: { size: TW[1], type: WidthType.DXA }, margins: cellMargins, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new SimpleField('PAGEREF rsec_' + i + ' \\h')] })] }),
      ] }));
    });
    children.push(new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: TW, rows }));
  }

  RSEC.forEach((s, i) => {
    children.push(new Paragraph({ children: [new PageBreak()] }));
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 0, after: 120 }, children: [new Bookmark({ id: 'rsec_' + i, children: [new TextRun({ text: s.name, bold: true, size: 28 })] })] }));
    if (s.intro) children.push(para(s.intro, { italics: true, size: 18, after: 140 }));
    const kind = s.kind || 'narrative';
    if (kind === 'narrative') {
      for (const p of (s.body || [])) children.push(para(p));
    } else if (kind === 'subcomponents') {
      for (const r of (s.refs || [])) {
        children.push(new Paragraph({ spacing: { before: 180, after: 40 }, children: [new TextRun({ text: r.title, bold: true, size: 22 })] }));
        if (r.link) children.push(new Paragraph({ spacing: { before: 20, after: 40 }, children: [new TextRun({ text: 'Reference: ', size: 20 }), new TextRun({ text: r.link, color: '0066CC', size: 20 })] }));
        if (r.passCriteria) children.push(para(r.passCriteria, { size: 20, before: 20 }));
        if (r.attestHeaders && r.attestHeaders.length === 2) {
          const AW = [Math.round(CONTENT_W * 0.5), CONTENT_W - Math.round(CONTENT_W * 0.5)];
          children.push(new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: AW, rows: [
            new TableRow({ tableHeader: true, children: r.attestHeaders.map((l, j) => cell(l, AW[j], { bold: true, fill: HEADER_FILL })) }),
            new TableRow({ cantSplit: true, children: [cell('', AW[0]), cell('', AW[1])] }),
          ] }));
        }
      }
    } else if (kind === 'units' || kind === 'table') {
      children.push(genTable(s.columns || [], s.rows || [], s.columnWidths, { centerHeader: true, boldFirst: true }));
    } else if (kind === 'scope') {
      if (s.legend && s.legend.length) {
        if (s.legendTitle) children.push(subhead(s.legendTitle, 60));
        children.push(genTable(['Trigger', 'Description', 'This release'], s.legend.map((c) => [c.id, c.name, c.fired || '']), [1300, 5900, 3024], { boldFirst: true }));
      }
      if (s.headers && s.rows) {
        if (s.matrixTitle) children.push(subhead(s.matrixTitle));
        children.push(genTable(s.headers, s.rows, s.columnWidths, { centerHeader: true, boldFirst: true }));
      }
    } else if (kind === 'tests') {
      // Report test table — no P/F column (this is a report, not a checklist).
      const TW2 = [820, 1900, 3100, 2400, 2004];
      const hr = new TableRow({ tableHeader: true, children: ['Ref', 'Header', 'Brief', 'Criteria', 'Result'].map((l, j) => cell(l, TW2[j], { bold: true, fill: HEADER_FILL, align: AlignmentType.CENTER })) });
      const br = (s.steps || []).map((st, ri) => new TableRow({ cantSplit: true, children: [
        cell(st.ref, TW2[0], { fill: ri % 2 ? ALT_FILL : undefined, bold: true }),
        cell(st.header, TW2[1], { fill: ri % 2 ? ALT_FILL : undefined, bold: true }),
        cell(st.brief, TW2[2], { fill: ri % 2 ? ALT_FILL : undefined }),
        cell(st.criteria, TW2[3], { fill: ri % 2 ? ALT_FILL : undefined }),
        cell(st.measure || '', TW2[4], { fill: ri % 2 ? ALT_FILL : undefined }),
      ] }));
      children.push(new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: TW2, rows: [hr, ...br] }));
    } else if (kind === 'figures') {
      const MAXW = 520;
      for (const g of (s.figures || [])) {
        if (g.heading) children.push(subhead(g.heading));
        try {
          const imgPath = path.isAbsolute(g.file) ? g.file : path.resolve(path.dirname(path.resolve(INPUT)), g.file);
          const imgData = fs.readFileSync(imgPath);
          const sz = getImageSize(imgData);
          const scale = Math.min(1, MAXW / sz.w);
          const W = Math.round(sz.w * scale), H = Math.round(sz.h * scale);
          const ext = (g.file.split('.').pop() || '').toLowerCase();
          const type = (ext === 'jpg' || ext === 'jpeg') ? 'jpg' : (ext === 'gif' ? 'gif' : (ext === 'bmp' ? 'bmp' : 'png'));
          children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 20 }, children: [new ImageRun({ type, data: imgData, transformation: { width: W, height: H } })] }));
        } catch (e) {
          children.push(para('[image not found: ' + g.file + ']', { italics: true, size: 16 }));
        }
        if (g.caption) children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 160 }, children: [new TextRun({ text: g.caption, italics: true, size: 16 })] }));
      }
    } else if (kind === 'anomalies') {
      if (s.anomalies && s.anomalies.length) {
        children.push(genTable(['ID', 'Activity', 'Finding', 'Disposition'], s.anomalies.map((a) => [a.id, a.activity, a.finding, a.disposition]), [900, 1500, 4000, 3824], { boldFirst: true }));
      }
      if (s.deviations) children.push(para(s.deviations, { before: 160 }));
      if (s.recommendationsTitle) children.push(subhead(s.recommendationsTitle));
      for (const r of (s.recommendations || [])) children.push(para(r, { size: 18, before: 20, after: 20 }));
      if (s.conclusionTitle) children.push(subhead(s.conclusionTitle));
      for (const p of (s.conclusion || [])) children.push(para(p, { size: 20 }));
    }
  });
}

// ============ Document config ============
const doc = new Document({
  creator: meta.creator || 'DVT Generator',
  styles: {
    default: { document: { run: { font: 'Arial', size: 20 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 28, bold: true, font: 'Arial' }, paragraph: { spacing: { before: 120, after: 120 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 22, bold: true, font: 'Arial' }, paragraph: { spacing: { before: 100, after: 100 }, outlineLevel: 1 } },
    ],
  },
  sections: [{
    properties: { page: { size: { width: PAGE.width, height: PAGE.height, orientation: PageOrientation.PORTRAIT }, margin: { top: PAGE.margin, right: PAGE.margin, bottom: PAGE.margin, left: PAGE.margin } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: meta.headerLine || (meta.productName ? `${meta.productName} DVT` : 'DVT'), color: '666666', size: 16 })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Page ', size: 16, color: '666666' }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: '666666' }), new TextRun({ text: ' of ', size: 16, color: '666666' }), new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: '666666' })] })] }) },
    children,
  }],
});

async function injectCheckboxControls(buf) {
  const zip = await JSZip.loadAsync(buf);
  let docXml = await zip.file('word/document.xml').async('string');
  const w14Ns = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
  if (!docXml.includes(w14Ns)) docXml = docXml.replace(/<w:document\b/, `<w:document ${w14Ns}`);
  const bkmRangeRe = /(<w:bookmarkStart\b[^>]*\bw:name="pf_[^"]+"[^>]*\/>)([\s\S]*?)(<w:bookmarkEnd\b[^>]*\/>)/g;
  const runWithBoxRe = /<w:r\b[\s\S]*?<w:t[^>]*>(?:☐|&#9744;|&#x2610;|&#X2610;)<\/w:t>[\s\S]*?<\/w:r>/;
  let sdtId = 50000, replaceCount = 0;
  docXml = docXml.replace(bkmRangeRe, (full, bkmStart, inner, bkmEnd) => {
    const m = inner.match(runWithBoxRe);
    if (!m) return full;
    const runXml = m[0];
    const before = inner.slice(0, m.index);
    const after = inner.slice(m.index + runXml.length);
    replaceCount++;
    const id = sdtId++;
    const sdt = '<w:sdt><w:sdtPr><w:rPr/><w:id w:val="' + id + '"/><w14:checkbox><w14:checked w:val="0"/><w14:checkedState w:val="2612" w:font="MS Gothic"/><w14:uncheckedState w:val="2610" w:font="MS Gothic"/></w14:checkbox></w:sdtPr><w:sdtContent>' + runXml + '</w:sdtContent></w:sdt>';
    return `${bkmStart}${before}${sdt}${after}${bkmEnd}`;
  });
  zip.file('word/document.xml', docXml);
  const out = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  return { buf: out, count: replaceCount };
}

Packer.toBuffer(doc).then(async (buf) => {
  const { buf: finalBuf, count } = await injectCheckboxControls(buf);
  fs.writeFileSync(OUT, finalBuf);
  console.log(`Wrote ${OUT} (${finalBuf.length} bytes); injected ${count} content-control checkboxes.`);
}).catch((err) => { console.error(err); process.exit(1); });
