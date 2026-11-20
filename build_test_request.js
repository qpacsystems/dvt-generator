#!/usr/bin/env node
// R&D Test Request generator — same visual language as build_dvt.js.
// Usage:
//   node build_test_request.js [request.json]
//   default: request.json = ./test_request_data.json
//
// The request covers the "why" and "what", not the "how": the device under
// test, why we're testing, what we hope to learn, the broad test approach,
// envisioned tests, and what pass / fail looks like. The lab plans and
// records the detailed testing in its own format.
//
// Any string can use {key} to pull a string/number from `meta`.
// Any text field may be a string (split on \n into paragraphs) or an array of
// strings (one paragraph each).

const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, PageOrientation,
  HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber,
} = require('docx');

// ============ CLI / load ============
const INPUT = process.argv[2] || 'test_request_data.json';
if (!fs.existsSync(INPUT)) {
  console.error(`Input not found: ${INPUT}`);
  console.error('Usage: node build_test_request.js [request.json]');
  process.exit(1);
}
let data = JSON.parse(fs.readFileSync(INPUT, 'utf-8'));
data.meta = data.meta || {};
if (!data.meta.date) {
  const t = new Date();
  data.meta.date = `${t.getMonth() + 1}/${t.getDate()}/${t.getFullYear()}`;
}

// ---------- {key} interpolation from meta ----------
function deepInterpolate(obj, vars) {
  if (typeof obj === 'string') return obj.replace(/\{(\w+)\}/g, (m, k) => (k in vars ? String(vars[k]) : m));
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
const meta = data.meta;

// ---------- Lint ----------
const warnings = [];
(function walk(obj, where) {
  if (typeof obj === 'string') {
    for (const m of obj.matchAll(/\{(\w+)\}/g)) warnings.push(`Unresolved placeholder ${m[0]} in ${where}`);
    if (/<[^<>]*>/.test(obj) && !/^https?:/.test(obj)) warnings.push(`Template text left in ${where}: "${obj.slice(0, 60)}"`);
  } else if (Array.isArray(obj)) obj.forEach((x, i) => walk(x, `${where}[${i}]`));
  else if (obj && typeof obj === 'object') for (const k of Object.keys(obj)) walk(obj[k], `${where}.${k}`);
})(data, 'data');
const tests = data.envisionedTests || [];
const objectives = (data.purpose && data.purpose.objectives) || [];
{
  const seen = new Set();
  tests.forEach((t, i) => {
    if (!t.id) t.id = `R${i + 1}`;
    if (seen.has(t.id)) warnings.push(`Duplicate envisioned-test id: ${t.id}`);
    seen.add(t.id);
    if (!t.pass || !t.fail) warnings.push(`Envisioned test ${t.id} is missing a pass or fail description`);
  });
}
if (!objectives.length) warnings.push('purpose.objectives is empty — what do you hope to learn?');
if (!data.successCriteria || !data.successCriteria.pass || !data.successCriteria.fail) {
  warnings.push('successCriteria.pass / successCriteria.fail should both be filled in');
}
if (warnings.length) {
  console.warn(`\nLint warnings (${warnings.length}):`);
  for (const w of warnings) console.warn('  ' + w);
  console.warn('');
}

// ---------- Output path ----------
const sanitize = (s) => String(s).replace(/[—–]/g, '-').replace(/[\/\\:*?"<>|]/g, '_').replace(/\s+/g, ' ').trim();
const OUT_NAME = meta.outputFilename || (meta.headerLine ? sanitize(meta.headerLine) + '.docx' : path.basename(INPUT, '.json') + '.docx');
const inputDir = path.dirname(path.resolve(INPUT));
const OUT_DIR = !meta.outputDir ? inputDir : (path.isAbsolute(meta.outputDir) ? meta.outputDir : path.resolve(inputDir, meta.outputDir));
const OUT = path.join(OUT_DIR, OUT_NAME);
if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });

// ============ Layout constants (match build_dvt.js) ============
const PAGE = { width: 12240, height: 15840, margin: 1008 }; // US Letter, 0.7"
const CONTENT_W = PAGE.width - 2 * PAGE.margin; // 10224 DXA

const border = { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 60, bottom: 60, left: 100, right: 100 };
const HEADER_FILL = 'D9E2F3';
const ALT_FILL = 'F5F7FB';
const LABEL_FILL = 'F2F2F2';
const NOTE_COLOR = '808080';

// ============ Helpers ============
const toParas = (text) => (Array.isArray(text) ? text : String(text == null ? '' : text).split('\n'));

function textParas(text, opts = {}) {
  return toParas(text).map((line) => new Paragraph({
    spacing: { before: opts.before ?? 20, after: opts.after ?? 20 },
    alignment: opts.align || AlignmentType.LEFT,
    children: [new TextRun({ text: line, bold: !!opts.bold, italics: !!opts.italics, size: opts.size || 18, color: opts.color })],
  }));
}

function cell(content, widthDxa, opts = {}) {
  const children = Array.isArray(content) && content[0] instanceof Paragraph ? content : textParas(content, opts);
  return new TableCell({
    borders,
    width: { size: widthDxa, type: WidthType.DXA },
    shading: opts.fill ? { fill: opts.fill, type: ShadingType.CLEAR } : undefined,
    margins: cellMargins,
    columnSpan: opts.span,
    children,
  });
}

const row = (cells, opts = {}) => new TableRow({
  cantSplit: true,
  tableHeader: !!opts.header,
  children: cells,
});

const headerRow = (labels, widths) => row(
  labels.map((l, i) => cell(l, widths[i], { bold: true, fill: HEADER_FILL, align: AlignmentType.CENTER })),
  { header: true },
);

const table = (widths, rows) => new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: widths, rows });

// Label / value two-column table.
const KV_W = [3000, CONTENT_W - 3000];
function kvTable(rows) {
  return table(KV_W, rows.map((r) => row([
    cell(r.label, KV_W[0], { bold: true, fill: LABEL_FILL }),
    cell(r.value == null ? '' : r.value, KV_W[1]),
  ])));
}

const h1 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 240, after: 100 }, children: [new TextRun({ text, bold: true, size: 26 })] });
const h2 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 80 }, children: [new TextRun({ text, bold: true, size: 22 })] });
const note = (text) => new Paragraph({ keepNext: true, keepLines: true, spacing: { before: 20, after: 100 }, children: [new TextRun({ text, italics: true, size: 16, color: NOTE_COLOR })] });

// ============ Compose ============
const children = [];

// ---- Title block + cover rows ----
children.push(new Paragraph({ spacing: { before: 0, after: 60 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: meta.title || 'R&D TEST REQUEST', bold: true, size: 32 })] }));
if (meta.subtitle) children.push(new Paragraph({ spacing: { before: 0, after: 160 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: meta.subtitle, size: 22 })] }));
if ((meta.coverRows || []).length) children.push(kvTable(meta.coverRows));


// 1. Device under test
children.push(h1('1  Device Under Test'));
children.push(kvTable(data.deviceUnderTest || []));

// 2. Why — purpose & what we hope to learn
const purpose = data.purpose || {};
children.push(h1('2  Purpose — Why Are We Testing?'));
if (purpose.background) children.push(...textParas(purpose.background, { size: 20, after: 80 }));
if (objectives.length) {
  children.push(h2('What I hope to learn'));
  const OBJ_W = [700, CONTENT_W - 700];
  children.push(table(OBJ_W, [
    headerRow(['#', 'Question / objective'], OBJ_W),
    ...objectives.map((o, i) => row([
      cell(`Q${i + 1}`, OBJ_W[0], { bold: true, align: AlignmentType.CENTER, fill: i % 2 ? ALT_FILL : undefined }),
      cell(o, OBJ_W[1], { fill: i % 2 ? ALT_FILL : undefined }),
    ])),
  ]));
}

// 3. Approach — broad strokes
if (data.approach) {
  children.push(h1('3  Test Approach — Broad Strokes'));
  children.push(...textParas(data.approach, { size: 20, after: 80 }));
}

// 4. Envisioned tests
if (tests.length) {
  children.push(h1('4  Envisioned Tests'));
  children.push(note('High-level only — the lab owns the detailed method, setup, and equipment. Pass / fail describes the outcome, not the procedure.'));
  const T_W = [600, 1800, 2800, 2512, 2512];
  children.push(table(T_W, [
    headerRow(['ID', 'Test area', 'What we want to learn', 'Pass looks like', 'Fail looks like'], T_W),
    ...tests.map((t, i) => {
      const fill = i % 2 ? ALT_FILL : undefined;
      return row([
        cell(t.id, T_W[0], { bold: true, fill }),
        cell(t.name || '', T_W[1], { bold: true, fill }),
        cell(t.learn || '', T_W[2], { fill }),
        cell(t.pass || '', T_W[3], { fill }),
        cell(t.fail || '', T_W[4], { fill }),
      ]);
    }),
  ]));
}

// 5. Overall success criteria
const sc = data.successCriteria || {};
if (sc.pass || sc.fail || sc.inconclusive) {
  children.push(h1('5  Overall Pass / Fail'));
  children.push(kvTable([
    sc.pass && { label: 'PASS looks like', value: sc.pass },
    sc.fail && { label: 'FAIL looks like', value: sc.fail },
    sc.inconclusive && { label: 'INCONCLUSIVE if', value: sc.inconclusive },
  ].filter(Boolean)));
}

// 6. Constraints / notes for the lab
if ((data.constraints || []).length) {
  children.push(h1('6  Constraints & Notes for the Lab'));
  for (const c of data.constraints) {
    children.push(new Paragraph({ bullet: { level: 0 }, spacing: { before: 20, after: 20 }, children: [new TextRun({ text: c, size: 20 })] }));
  }
}

// ============ Document ============
const doc = new Document({
  creator: meta.creator || 'Test Request Generator',
  styles: {
    default: { document: { run: { font: 'Arial', size: 20 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 26, bold: true, font: 'Arial', color: '1F3864' }, paragraph: { spacing: { before: 240, after: 100 }, outlineLevel: 0, keepNext: true } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 22, bold: true, font: 'Arial' }, paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 1, keepNext: true } },
    ],
  },
  sections: [{
    properties: { page: { size: { width: PAGE.width, height: PAGE.height, orientation: PageOrientation.PORTRAIT }, margin: { top: PAGE.margin, right: PAGE.margin, bottom: PAGE.margin, left: PAGE.margin } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: meta.headerLine || 'R&D Test Request', color: '666666', size: 16 })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Page ', size: 16, color: '666666' }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: '666666' }), new TextRun({ text: ' of ', size: 16, color: '666666' }), new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: '666666' })] })] }) },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(OUT, buf);
  console.log(`Wrote ${OUT} (${buf.length} bytes); ${tests.length} envisioned test(s).`);
}).catch((err) => {
  console.error(err);
  process.exit(1);
});
