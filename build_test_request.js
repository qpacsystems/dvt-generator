#!/usr/bin/env node
// R&D Test Request generator — same visual language as build_dvt.js.
// Usage:
//   node build_test_request.js [request.json]
//   default: request.json = ./test_request_data.json
//
// The document has two parts:
//   Part 1 — Request (filled by the requester, from the JSON): the device under
//            test, why we're testing, what we hope to learn, the broad test
//            approach, envisioned tests, and what pass / fail looks like.
//   Part 2 — Lab Execution (left blank for the lab to fill in by hand / in
//            Word): intake, specific tests performed, dates, notes, P/F,
//            findings against each objective, overall result, sign-off.
//
// Any string can use {key} to pull a string/number from `meta`.
// Any text field may be a string (split on \n into paragraphs) or an array of
// strings (one paragraph each).

const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, PageOrientation, HeightRule,
  HeadingLevel, BorderStyle, WidthType, ShadingType, PageBreak, PageNumber,
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
const PART_FILL = '1F3864';
const LAB_HINT = '808080';
const BOX = '☐'; // every ☐ run becomes a clickable Word checkbox in post-processing

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
  height: opts.minHeight ? { value: opts.minHeight, rule: HeightRule.ATLEAST } : undefined,
  children: cells,
});

const headerRow = (labels, widths) => row(
  labels.map((l, i) => cell(l, widths[i], { bold: true, fill: HEADER_FILL, align: AlignmentType.CENTER })),
  { header: true },
);

const table = (widths, rows) => new Table({ width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: widths, rows });

// Label / value two-column table. Empty value → blank write-in cell.
const KV_W = [3000, CONTENT_W - 3000];
function kvTable(rows, opts = {}) {
  return table(KV_W, rows.map((r) => row([
    cell(r.label, KV_W[0], { bold: true, fill: LABEL_FILL }),
    cell(r.value == null ? '' : r.value, KV_W[1]),
  ], { minHeight: (r.value ? 0 : opts.blankHeight) || 0 })));
}

// A bordered write-in box of a given minimum height (DXA; 1440 = 1").
function writeInBox(minHeight, hint) {
  return table([CONTENT_W], [row([
    cell(hint ? textParas(hint, { italics: true, color: LAB_HINT, size: 16 }) : '', CONTENT_W),
  ], { minHeight })]);
}

// Paragraph containing "☐ label   ☐ label ..." checkbox choices.
function checkChoices(labels, opts = {}) {
  const runs = [];
  labels.forEach((l, i) => {
    if (i) runs.push(new TextRun({ text: opts.vertical ? '' : '     ', size: opts.size || 18 }));
    runs.push(new TextRun({ text: BOX, size: (opts.size || 18) + 4 }));
    runs.push(new TextRun({ text: ' ' + l, size: opts.size || 18, bold: !!opts.bold }));
  });
  return new Paragraph({ spacing: { before: 20, after: 20 }, alignment: opts.align, children: runs });
}

const h1 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 240, after: 100 }, children: [new TextRun({ text, bold: true, size: 26 })] });
const h2 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 80 }, children: [new TextRun({ text, bold: true, size: 22 })] });
const note = (text) => new Paragraph({ keepNext: true, keepLines: true, spacing: { before: 20, after: 100 }, children: [new TextRun({ text, italics: true, size: 16, color: LAB_HINT })] });
const spacer = () => new Paragraph({ spacing: { before: 0, after: 0 }, children: [] });

// Full-width dark banner marking "who fills this part in".
function partBanner(title, who) {
  return table([CONTENT_W], [row([new TableCell({
    borders,
    width: { size: CONTENT_W, type: WidthType.DXA },
    shading: { fill: PART_FILL, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 140, right: 140 },
    children: [new Paragraph({ children: [
      new TextRun({ text: title, bold: true, size: 26, color: 'FFFFFF' }),
      new TextRun({ text: `   —   ${who}`, italics: true, size: 18, color: 'FFFFFF' }),
    ] })],
  })])]);
}

// ============ Compose ============
const children = [];

// ---- Title block + cover rows ----
children.push(new Paragraph({ spacing: { before: 0, after: 60 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: meta.title || 'R&D TEST REQUEST', bold: true, size: 32 })] }));
if (meta.subtitle) children.push(new Paragraph({ spacing: { before: 0, after: 160 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: meta.subtitle, size: 22 })] }));
if ((meta.coverRows || []).length) children.push(kvTable(meta.coverRows));

// ---- PART 1 — REQUEST ----
children.push(spacer());
children.push(partBanner('PART 1 — REQUEST', meta.part1Who || 'completed by requester'));

// 1.1 Device under test
children.push(h1('1.1  Device Under Test'));
children.push(kvTable(data.deviceUnderTest || []));

// 1.2 Why — purpose & what we hope to learn
const purpose = data.purpose || {};
children.push(h1('1.2  Purpose — Why Are We Testing?'));
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

// 1.3 Approach — broad strokes
if (data.approach) {
  children.push(h1('1.3  Test Approach — Broad Strokes'));
  children.push(...textParas(data.approach, { size: 20, after: 80 }));
}

// 1.4 Envisioned tests
if (tests.length) {
  children.push(h1('1.4  Envisioned Tests'));
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

// 1.5 Overall success criteria
const sc = data.successCriteria || {};
if (sc.pass || sc.fail || sc.inconclusive) {
  children.push(h1('1.5  Overall Pass / Fail'));
  children.push(kvTable([
    sc.pass && { label: 'PASS looks like', value: sc.pass },
    sc.fail && { label: 'FAIL looks like', value: sc.fail },
    sc.inconclusive && { label: 'INCONCLUSIVE if', value: sc.inconclusive },
  ].filter(Boolean)));
}

// 1.6 Constraints / notes for the lab
if ((data.constraints || []).length) {
  children.push(h1('1.6  Constraints & Notes for the Lab'));
  for (const c of data.constraints) {
    children.push(new Paragraph({ bullet: { level: 0 }, spacing: { before: 20, after: 20 }, children: [new TextRun({ text: c, size: 20 })] }));
  }
}

// ---- PART 2 — LAB EXECUTION ----
const lab = data.lab || {};
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(partBanner('PART 2 — LAB EXECUTION', meta.part2Who || 'completed by lab'));

// 2.1 Intake
children.push(h1('2.1  Intake'));
const intakeRows = (lab.intakeRows || [
  'Received by', 'Date received', 'Unit(s) received — SN', 'Condition on arrival',
  'Assigned technician', 'Planned start', 'Planned completion',
]).map((label) => ({ label, value: '' }));
children.push(kvTable(intakeRows, { blankHeight: 400 }));

// 2.2 Test log
children.push(h1('2.2  Test Plan & Log'));
children.push(note('One row per specific test. Link each to the request item it addresses (R#), or "—" for tests the lab adds. Describe method / setup / equipment briefly; attach data files and photos and reference them in the notes.'));
const L_W = [600, 700, 2900, 1100, 3400, 1524];
const blankRows = lab.blankRows ?? 6;
const logRow = (i, reqId) => {
  const fill = i % 2 ? ALT_FILL : undefined;
  return row([
    cell(`L${i + 1}`, L_W[0], { bold: true, fill }),
    cell(reqId || '', L_W[1], { fill, align: AlignmentType.CENTER }),
    cell('', L_W[2], { fill }),
    cell('', L_W[3], { fill }),
    cell('', L_W[4], { fill }),
    cell([checkChoices(['Pass']), checkChoices(['Fail'])], L_W[5], { fill }),
  ], { minHeight: lab.logRowHeight || 900 });
};
const logRows = [];
tests.forEach((t, i) => logRows.push(logRow(i, t.id)));
for (let j = 0; j < blankRows; j++) logRows.push(logRow(tests.length + j, ''));
children.push(table(L_W, [
  headerRow(['Ref', 'Req.', 'Test performed (method / setup)', 'Date(s)', 'Results & notes', 'Result'], L_W),
  ...logRows,
]));

// 2.3 Findings vs. objectives
if (objectives.length) {
  children.push(h1('2.3  Findings vs. Objectives'));
  children.push(note('Answer each question from section 1.2 in a sentence or two, citing the supporting L# rows.'));
  const F_W = [700, 3800, CONTENT_W - 4500];
  children.push(table(F_W, [
    headerRow(['#', 'Question / objective', 'What we learned'], F_W),
    ...objectives.map((o, i) => row([
      cell(`Q${i + 1}`, F_W[0], { bold: true, align: AlignmentType.CENTER }),
      cell(o, F_W[1], { size: 16, color: '404040' }),
      cell('', F_W[2]),
    ], { minHeight: 800 })),
  ]));
}

// 2.4 Observations / deviations
children.push(h1('2.4  Observations, Anomalies & Deviations'));
children.push(writeInBox(lab.notesHeight || 2400, 'Anything unexpected, deviations from the request, equipment issues, damage, or ideas for follow-up testing.'));

// 2.5 Overall result
children.push(h1('2.5  Overall Result'));
children.push(checkChoices(['PASS', 'FAIL', 'INCONCLUSIVE'], { bold: true, size: 22 }));
children.push(note('Judged against section 1.5.'));
children.push(writeInBox(1400, 'Summary / recommendation'));

// 2.6 Sign-off
children.push(h1('2.6  Sign-off'));
const S_W = [2900, 2800, 2800, 1724];
const signers = lab.signers || ['Lab technician', 'Lab lead / reviewer', 'Requester — results accepted'];
children.push(table(S_W, [
  headerRow(['Role', 'Name', 'Signature', 'Date'], S_W),
  ...signers.map((r) => row([cell(r, S_W[0], { bold: true, fill: LABEL_FILL }), cell('', S_W[1]), cell('', S_W[2]), cell('', S_W[3])], { minHeight: 500 })),
]));

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

// ============ Post-process: every ☐ run → Word content-control checkbox ============
async function injectCheckboxControls(buf) {
  const zip = await JSZip.loadAsync(buf);
  let xml = await zip.file('word/document.xml').async('string');
  const w14Ns = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';
  if (!xml.includes(w14Ns)) xml = xml.replace(/<w:document\b/, `<w:document ${w14Ns}`);
  // Match a single <w:r> (not <w:rPr>) whose only text is ☐, without spanning into other runs.
  const runRe = /<w:r(?:\s[^>]*)?>(?:(?!<\/w:r>)[\s\S])*?<w:t[^>]*>(?:☐|&#9744;|&#x2610;)<\/w:t>(?:(?!<\/w:r>)[\s\S])*?<\/w:r>/gi;
  let id = 50000, count = 0;
  xml = xml.replace(runRe, (runXml) => {
    count++;
    return '<w:sdt><w:sdtPr><w:rPr/>' + `<w:id w:val="${id++}"/>` +
      '<w14:checkbox><w14:checked w:val="0"/>' +
      '<w14:checkedState w:val="2612" w:font="MS Gothic"/>' +
      '<w14:uncheckedState w:val="2610" w:font="MS Gothic"/>' +
      `</w14:checkbox></w:sdtPr><w:sdtContent>${runXml}</w:sdtContent></w:sdt>`;
  });
  zip.file('word/document.xml', xml);
  return { buf: await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }), count };
}

Packer.toBuffer(doc).then(async (buf) => {
  const { buf: finalBuf, count } = await injectCheckboxControls(buf);
  fs.writeFileSync(OUT, finalBuf);
  console.log(`Wrote ${OUT} (${finalBuf.length} bytes); ${tests.length} envisioned test(s), ${count} checkboxes.`);
}).catch((err) => {
  console.error(err);
  process.exit(1);
});
