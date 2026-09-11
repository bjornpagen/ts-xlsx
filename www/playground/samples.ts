/**
 * The workbooks the playground offers, as builder functions.
 *
 * Each one exists to make a claim checkable rather than to look busy: a reader who does not
 * believe the type round-trip, or that a four-section number format survives, can pick the
 * sample that shows it and read the bytes that came out. The builders are ordinary calls
 * against the public API, and the page shows each one's source, so what a reader sees is
 * what produced the file.
 *
 * Nothing here touches a DOM, and nothing here reads a file.
 */

import {Workbook} from '../../src/index.ts';

export interface Sample {
  readonly id: string;
  readonly title: string;
  /** One line, rendered under the title. */
  readonly description: string;
  /** Where the claim this sample makes is written down, if it is written down. */
  readonly source?: {readonly label: string; readonly path: string};
  readonly build: () => Workbook;
}

// A four-section accounting format: quoted literals, alignment placeholders, and group
// separators. `test/corpus/cases/custom-numfmt-string-roundtrips-verbatim.case.ts` locks it,
// and the round-trip lane here is the same assertion with the bytes on screen.
const ACCOUNTING = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)';

function buildValues(): Workbook {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('Values');
  sheet.getCell('A1').value = 'kind';
  sheet.getCell('B1').value = 'value';
  sheet.getCell('A2').value = 'number';
  sheet.getCell('B2').value = 1234.5;
  sheet.getCell('A3').value = 'string';
  sheet.getCell('B3').value = 'hello';
  sheet.getCell('A4').value = 'boolean';
  sheet.getCell('B4').value = true;
  sheet.getCell('A5').value = 'date';
  sheet.getCell('B5').value = new Date(Date.UTC(2026, 7, 27));
  sheet.getCell('B5').numFmt = 'yyyy-mm-dd';
  sheet.getCell('A6').value = 'formula';
  sheet.getCell('B6').value = {formula: 'B2*2', result: 2469};
  sheet.getCell('A7').value = 'error';
  sheet.getCell('B7').value = {error: '#DIV/0!'};
  sheet.getCell('A8').value = 'hyperlink';
  sheet.getCell('B8').value = {hyperlink: 'https://github.com/shbernal/ts-xlsx', text: 'ts-xlsx'};
  sheet.getCell('A9').value = 'rich text';
  sheet.getCell('B9').value = {
    richText: [
      {text: 'half ', font: {bold: true}},
      {text: 'and half', font: {italic: true}},
    ],
  };
  sheet.getColumn(1).width = 14;
  sheet.getColumn(2).width = 28;
  return workbook;
}

function buildStyles(): Workbook {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('Styles');
  sheet.getCell('A1').value = 'Quarter';
  sheet.getCell('B1').value = 'Booked';
  sheet.getCell('C1').value = 'Closed';
  const header = sheet.getRange('A1:C1');
  header.font = {bold: true, color: {argb: 'FFFFFFFF'}};
  header.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF146C43'}};
  header.alignment = {horizontal: 'center', vertical: 'center'};

  const quarters = ['Q1', 'Q2', 'Q3', 'Q4'];
  const booked = [18400.5, 22750, 19980.25, 30500];
  const closed = [0.62, 0.71, 0.58, 0.83];
  for (const [index, quarter] of quarters.entries()) {
    const row = index + 2;
    sheet.getCell(`A${row}`).value = quarter;
    sheet.getCell(`B${row}`).value = booked[index] ?? 0;
    sheet.getCell(`B${row}`).numFmt = ACCOUNTING;
    sheet.getCell(`C${row}`).value = closed[index] ?? 0;
    sheet.getCell(`C${row}`).numFmt = '0.0%';
  }
  sheet.getRange('A1:C5').border = {
    top: {style: 'thin', color: {argb: 'FFBBBBBB'}},
    bottom: {style: 'thin', color: {argb: 'FFBBBBBB'}},
    left: {style: 'thin', color: {argb: 'FFBBBBBB'}},
    right: {style: 'thin', color: {argb: 'FFBBBBBB'}},
  };
  sheet.getColumn(1).width = 10;
  sheet.getColumn(2).width = 18;
  sheet.getColumn(3).width = 12;
  return workbook;
}

function buildStructure(): Workbook {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('Structure');
  sheet.getCell('A1').value = 'Regional summary';
  sheet.mergeCells('A1:C1');
  sheet.getCell('A1').alignment = {horizontal: 'center'};
  sheet.getCell('A1').font = {bold: true, size: 14};

  sheet.getCell('A2').value = 'Region';
  sheet.getCell('B2').value = 'Units';
  sheet.getCell('C2').value = 'Revenue';
  const regions: readonly (readonly [string, number, number])[] = [
    ['North', 1240, 84200],
    ['South', 980, 61050],
    ['East', 1655, 99310],
    ['West', 1102, 70480],
  ];
  for (const [index, [name, units, revenue]] of regions.entries()) {
    const row = index + 3;
    sheet.getCell(`A${row}`).value = name;
    sheet.getCell(`B${row}`).value = units;
    sheet.getCell(`C${row}`).value = revenue;
  }
  sheet.addTable({
    name: 'Regions',
    ref: 'A2',
    columns: [{name: 'Region'}, {name: 'Units'}, {name: 'Revenue'}],
    rowCount: regions.length,
    style: {name: 'TableStyleMedium2', showRowStripes: true},
  });
  sheet.freeze(2, 1);
  sheet.getColumn(1).width = 14;
  sheet.getColumn(2).width = 10;
  sheet.getColumn(3).width = 14;
  return workbook;
}

function buildScale(): Workbook {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('Scale');
  sheet.getCell('A1').value = 'id';
  sheet.getCell('B1').value = 'label';
  sheet.getCell('C1').value = 'amount';
  sheet.getCell('D1').value = 'running';
  for (let row = 2; row <= 4001; row++) {
    const n = row - 1;
    sheet.getCell(`A${row}`).value = n;
    sheet.getCell(`B${row}`).value = `row ${n}`;
    sheet.getCell(`C${row}`).value = (n * 37) % 1000;
    sheet.getCell(`D${row}`).value = {formula: `SUM(C$2:C${row})`};
  }
  sheet.getColumn(2).width = 12;
  sheet.getColumn(4).width = 16;
  return workbook;
}

function buildCorpusNumberFormat(): Workbook {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('NumFmt');
  sheet.getCell('A1').value = 'the format code';
  sheet.getCell('B1').value = ACCOUNTING;
  sheet.getCell('A2').value = 'a positive amount';
  sheet.getCell('B2').value = 1234.5;
  sheet.getCell('B2').numFmt = ACCOUNTING;
  sheet.getCell('A3').value = 'a negative amount';
  sheet.getCell('B3').value = -87.25;
  sheet.getCell('B3').numFmt = ACCOUNTING;
  sheet.getCell('A4').value = 'zero';
  sheet.getCell('B4').value = 0;
  sheet.getCell('B4').numFmt = ACCOUNTING;
  sheet.getColumn(1).width = 20;
  sheet.getColumn(2).width = 46;
  return workbook;
}

export const SAMPLES: readonly Sample[] = [
  {
    id: 'values',
    title: 'Every value kind',
    description:
      'One row per kind a cell can hold, so what the reader gets back can be compared with what went in.',
    build: buildValues,
  },
  {
    id: 'styles',
    title: 'Fonts, fills, borders and formats',
    description:
      'A styled table with a percentage format and a four-section accounting format on the money column.',
    build: buildStyles,
  },
  {
    id: 'structure',
    title: 'Merges, widths, a frozen pane and a table',
    description:
      'The structure around the values: a merged title, column widths, frozen headers, and a real table object.',
    build: buildStructure,
  },
  {
    id: 'scale',
    title: 'Four thousand rows',
    description:
      'Enough rows for the timings to mean something, each carrying a running-total formula.',
    build: buildScale,
  },
  {
    id: 'corpus-numfmt',
    title: 'A number format the corpus locks',
    description:
      'The four-section accounting code from a regression case, which must read back character for character.',
    source: {
      label: 'custom-numfmt-string-roundtrips-verbatim',
      path: 'test/corpus/cases/custom-numfmt-string-roundtrips-verbatim.case.ts',
    },
    build: buildCorpusNumberFormat,
  },
];

export function findSample(id: string): Sample | undefined {
  return SAMPLES.find((sample) => sample.id === id);
}

/**
 * A builder's own source, cut out of this module's text.
 *
 * The page shows the code that produced the file it is showing, and the only way for that to
 * stay true is to take it from this file rather than from a second copy of it. The caller
 * supplies the text, because the two callers get it differently: the page imports it with
 * Vite's `?raw`, and the test reads it off disk. Brace matching rather than a regular
 * expression, because a builder body contains braces and a lazy match would stop at the
 * first one.
 */
export function builderSource(moduleText: string, id: string): string {
  const sample = findSample(id);
  if (sample === undefined) throw new Error(`no sample with id "${id}"`);
  const name = sample.build.name;
  const start = moduleText.indexOf(`function ${name}(`);
  if (start === -1) {
    throw new Error(
      `builder "${name}" is not a top-level function declaration in this module's text, ` +
        'so the page cannot show the code that produced the file.',
    );
  }
  let depth = 0;
  for (let index = moduleText.indexOf('{', start); index < moduleText.length; index++) {
    const char = moduleText[index];
    if (char === '{') depth += 1;
    else if (char === '}') {
      depth -= 1;
      if (depth === 0) return moduleText.slice(start, index + 1);
    }
  }
  throw new Error(`builder "${name}" has no closing brace; the module text is truncated.`);
}
