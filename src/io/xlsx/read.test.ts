import assert from 'node:assert/strict';
import {test} from 'node:test';

import {strToU8, zipSync} from 'fflate';

import {isDateFormat} from '../../core/date.ts';
import type {Fill} from '../../core/style.ts';
import {isFormulaValue} from '../../core/value.ts';
import {DEFAULT_WORKBOOK_VIEW, Workbook, type WorkbookView} from '../../core/workbook.ts';
import {Worksheet} from '../../core/worksheet.ts';
import {parseXmlPasses} from '../../xml/xml-read.ts';
import {UnsupportedFormatError} from '../opc/errors.ts';
import {conditionalFormattingPass} from './conditional-formatting.ts';
import {dataValidationPass, extendedDataValidationPass} from './data-validation.ts';
import {sheetHyperlinkPass} from './hyperlinks.ts';
import {
  captureIn,
  elementIn,
  foreignPackage,
  foreignSheet,
  optionalPartText,
  partBytes,
  partText,
  patchParts,
  roundtrip,
  SHEET1,
  sheetXml,
} from './package.test-support.ts';
import {workbookViewPass} from './read-workbook-xml.ts';
import {worksheetPass} from './read-worksheet.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

// Drive the one pass under test over a fragment. The reader is a pass rather than a function over the
// part's text because the whole workbook part is read in one scan; a test of one element still wants
// to name only that element.
function readView(view: WorkbookView, xml: string): void {
  parseXmlPasses(xml, [workbookViewPass(view)]);
}

/** The foreground ARGB of a pattern fill, narrowing the Fill union (a gradient has no fgColor). */
function fillFgArgb(fill: Fill | undefined): string | undefined {
  return fill?.type === 'pattern' ? fill.fgColor?.argb : undefined;
}

test('a saved window view is read off <bookViews>, flags and all', () => {
  const view = {...DEFAULT_WORKBOOK_VIEW};
  readView(
    view,
    '<workbook><bookViews><workbookView visibility="hidden" minimized="1" xWindow="0" ' +
      'yWindow="480" windowWidth="9000" windowHeight="6000" activeTab="2"/></bookViews></workbook>',
  );
  assert.deepEqual(view, {
    x: 0,
    y: 480,
    width: 9000,
    height: 6000,
    activeTab: 2,
    visibility: 'hidden',
    minimized: true,
  });
});

test('a workbook with no <bookViews> keeps the default window view', () => {
  const view = {...DEFAULT_WORKBOOK_VIEW};
  readView(view, '<workbook><sheets/></workbook>');
  assert.deepEqual(view, {...DEFAULT_WORKBOOK_VIEW});
});

test('an unusable window attribute falls back to the default rather than a NaN geometry', () => {
  const view = {...DEFAULT_WORKBOOK_VIEW};
  readView(view, '<bookViews><workbookView xWindow="wide" windowWidth="8000"/></bookViews>');
  assert.equal(view.x, DEFAULT_WORKBOOK_VIEW.x);
  assert.equal(view.width, 8000);
});

test('only the first <workbookView> is read: the model carries one window', () => {
  const view = {...DEFAULT_WORKBOOK_VIEW};
  readView(
    view,
    '<bookViews><workbookView windowWidth="100"/><workbookView windowWidth="200"/></bookViews>',
  );
  assert.equal(view.width, 100);
});

test('scalar cell values survive the round-trip with their types', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 42;
  sheet.getCell('A2').value = 'hello';
  sheet.getCell('A3').value = true;
  sheet.getCell('A4').value = false;

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getCell('A1').value, 42);
  assert.equal(back?.getCell('A2').value, 'hello');
  assert.equal(back?.getCell('A3').value, true);
  assert.equal(back?.getCell('A4').value, false);
});

test('a string with markup-significant and leading/trailing space round-trips exactly', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = '  <a> & "b" \t end  ';
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getCell('A1').value, '  <a> & "b" \t end  ');
});

test('a formula with a numeric result round-trips as {formula, result}', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: 'SUM(B1:B2)', result: 7};
  const value = roundtrip(wb).getWorksheet('S')?.getCell('A1').value;
  assert.ok(value && isFormulaValue(value));
  assert.equal(value.formula, 'SUM(B1:B2)');
  assert.equal(value.result, 7);
});

test('a formula with a string result carries its t="str" result back', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: 'CONCAT("a","b")', result: 'ab'};
  const value = roundtrip(wb).getWorksheet('S')?.getCell('A1').value;
  assert.ok(value && isFormulaValue(value));
  assert.equal(value.result, 'ab');
});

test('a formula with no cached result round-trips without inventing one', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: 'NOW()'};
  const value = roundtrip(wb).getWorksheet('S')?.getCell('A1').value;
  assert.ok(value && isFormulaValue(value));
  assert.equal(value.result, undefined);
});

test('a modern function is stored with _xlfn. on disk but round-trips as its plain name', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: 'FILTER(B1:D1,B2:D2=1)', result: 0};
  const xml = sheetXml(writeXlsx(wb));
  assert.match(xml, /<f>_xlfn\._xlws\.FILTER\(B1:D1,B2:D2=1\)<\/f>/);

  const value = roundtrip(wb).getWorksheet('S')?.getCell('A1').value;
  assert.ok(value && isFormulaValue(value));
  assert.equal(value.formula, 'FILTER(B1:D1,B2:D2=1)');
});

test('a LET formula stores its parameters with _xlpm. on disk but round-trips as plain names', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: 'LET(x,B1,x+1)', result: 0};
  const xml = sheetXml(writeXlsx(wb));
  assert.match(xml, /<f>_xlfn\.LET\(_xlpm\.x,B1,_xlpm\.x\+1\)<\/f>/);

  const value = roundtrip(wb).getWorksheet('S')?.getCell('A1').value;
  assert.ok(value && isFormulaValue(value));
  assert.equal(value.formula, 'LET(x,B1,x+1)');
});

test('a formula that already carries _xlfn. is not double-prefixed on write', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: '_xlfn.XLOOKUP(1,B:B,C:C)', result: 0};
  const xml = sheetXml(writeXlsx(wb));
  assert.match(xml, /_xlfn\.XLOOKUP/);
  assert.doesNotMatch(xml, /_xlfn\._xlfn/);
});

test('a dotted statistical function is stored whole with _xlfn. and round-trips as its plain name', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: 'NORM.DIST(A2,0,1,TRUE)', result: 0.5};
  const xml = sheetXml(writeXlsx(wb));
  assert.match(xml, /<f>_xlfn\.NORM\.DIST\(A2,0,1,TRUE\)<\/f>/);
  assert.doesNotMatch(xml, /_xlfn\.DIST/);

  const value = roundtrip(wb).getWorksheet('S')?.getCell('A1').value;
  assert.ok(value && isFormulaValue(value));
  assert.equal(value.formula, 'NORM.DIST(A2,0,1,TRUE)');
});

test('column width and visibility round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getColumn(2).width = 24;
  sheet.getColumn(4).hidden = true;
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getColumn(2).width, 24);
  assert.equal(back?.getColumn(4).hidden, true);
});

test('column outline grouping round-trips', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getColumn(2).outlineLevel = 2;
  sheet.getColumn(3).collapsed = true;
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getColumn(2).outlineLevel, 2);
  assert.equal(back?.getColumn(3).collapsed, true);
});

test('row height and visibility round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.getRow(1).height = 33;
  sheet.getRow(2).hidden = true;
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getRow(1).height, 33);
  assert.equal(back?.getRow(2).hidden, true);
});

test('merged ranges round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.mergeCells('A1:B2');
  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual([...(back?.merges ?? [])], ['A1:B2']);
});

test('page margins round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.pageMargins.left = 0.5;
  sheet.pageMargins.top = 1.25;
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.pageMargins.left, 0.5);
  assert.equal(back?.pageMargins.top, 1.25);
});

test('workbook core properties round-trip', () => {
  const wb = new Workbook();
  wb.properties.creator = 'Ada Lovelace';
  wb.properties.lastModifiedBy = 'Grace Hopper';
  wb.properties.created = new Date('2020-01-02T03:04:05.000Z');
  wb.properties.modified = new Date('2021-06-07T08:09:10.000Z');
  wb.addWorksheet('S').getCell('A1').value = 1;

  const back = roundtrip(wb);
  assert.equal(back.properties.creator, 'Ada Lovelace');
  assert.equal(back.properties.lastModifiedBy, 'Grace Hopper');
  assert.equal(back.properties.created?.toISOString(), '2020-01-02T03:04:05.000Z');
  assert.equal(back.properties.modified?.toISOString(), '2021-06-07T08:09:10.000Z');
});

test('multiple sheets round-trip in order and are addressable by name', () => {
  const wb = new Workbook();
  for (const name of ['First', 'Second', 'Third']) wb.addWorksheet(name).getCell('A1').value = name;
  const back = roundtrip(wb);
  assert.deepEqual(
    back.worksheets.map((s) => s.name),
    ['First', 'Second', 'Third'],
  );
  assert.equal(back.getWorksheet('Second')?.getCell('A1').value, 'Second');
});

test('rowCount spans a gap; actualRowCount counts only populated rows', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'a';
  sheet.getCell('A3').value = 'c';
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.rowCount, 3);
  assert.equal(back?.actualRowCount, 2);
});

test('a solid pattern fill round-trips with its foreground colour on a single cell', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'painted';
  sheet.getCell('A1').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF0000'}};
  sheet.getCell('B2').value = 'plain';

  const back = roundtrip(wb).getWorksheet('S');
  const fill = back?.getCell('A1').fill;
  assert.equal(fill?.type, 'pattern');
  assert.equal(fill?.pattern, 'solid');
  assert.equal(fill?.fgColor?.argb, 'FFFF0000');
  assert.equal(fill?.bgColor?.indexed, 64, 'solid fill keeps the automatic indexed background');
  assert.equal(back?.getCell('B2').fill, undefined, 'an unfilled cell reads back with no fill');
});

test('a formatted-but-empty cell round-trips with its fill and a null value', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'anchor';
  // B2 is styled but never given a value: the fill must survive without a value being invented.
  sheet.getCell('B2').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF00FF00'}};

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(
    back?.getCell('B2').value,
    null,
    'the empty cell stays empty, not fabricated a value',
  );
  assert.equal(
    fillFgArgb(back?.getCell('B2').fill),
    'FF00FF00',
    'the fill survives on the empty cell',
  );
});

test('two cells with different fills stay distinct across the round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 1;
  sheet.getCell('A1').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF0000'}};
  sheet.getCell('A2').value = 2;
  sheet.getCell('A2').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF0000FF'}};

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(fillFgArgb(back?.getCell('A1').fill), 'FFFF0000');
  assert.equal(fillFgArgb(back?.getCell('A2').fill), 'FF0000FF');
});

test('a row-level fill is inherited by the row cells that carry no fill of their own', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  for (let r = 1; r <= 4; r++) sheet.getCell(`A${r}`).value = `r${r}`;
  sheet.getRow(3).fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF4500'}};

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(
    fillFgArgb(back?.getCell('A3').fill),
    'FFFF4500',
    'the formatted row paints its cell',
  );
  for (const ref of ['A1', 'A2', 'A4']) {
    assert.equal(back?.getCell(ref).fill, undefined, `${ref} does not inherit the row-3 fill`);
  }
});

test('many identically-filled cells collapse to one shared style entry in the package', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  for (let r = 1; r <= 40; r++) {
    sheet.getCell(`A${r}`).value = r;
    sheet.getCell(`A${r}`).fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFDDEEFF'}};
  }
  const stylesXml = partText(writeXlsx(wb), 'xl/styles.xml');
  // Default xf + the single shared fill = two entries, never ~40.
  assert.match(stylesXml, /<cellXfs count="2">/);
});

test('a custom cell number format round-trips byte-for-byte', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  const fmt = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)';
  sheet.getCell('A1').value = 1234.5;
  sheet.getCell('A1').numFmt = fmt;
  sheet.getCell('B1').value = 'plain';

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(
    back?.getCell('A1').numFmt,
    fmt,
    'the custom format survives character-for-character',
  );
  assert.equal(
    back?.getCell('B1').numFmt,
    undefined,
    'an unformatted cell reads back with no numFmt',
  );
});

test('a column number format round-trips and is inherited by the column cells that carry no format', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getColumn(1).numFmt = '0.00%';
  sheet.getColumn(2).numFmt = '"$"#,##0.00';
  sheet.getCell('A1').value = 0.1;
  sheet.getCell('B1').value = 3;

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getColumn(1).numFmt, '0.00%', 'column 1 keeps its own format');
  assert.equal(
    back?.getColumn(2).numFmt,
    '"$"#,##0.00',
    'column 2 keeps its own format, not the last-assigned one',
  );
  assert.equal(back?.getCell('A1').numFmt, '0.00%', 'a bare cell inherits its column format');
  assert.equal(
    back?.getCell('B1').numFmt,
    '"$"#,##0.00',
    'a bare cell inherits its own column, not a sibling’s',
  );
});

test('a cell fill and its column number format both survive: overriding one facet keeps the other', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getColumn(1).numFmt = '0.00';
  sheet.getCell('A1').value = 1;
  sheet.getCell('A2').value = 2;
  sheet.getCell('A2').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFFFF00'}};
  sheet.getCell('A3').value = 3;

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(fillFgArgb(back?.getCell('A2').fill), 'FFFFFF00', 'the filled cell keeps its fill');
  for (const ref of ['A1', 'A2', 'A3']) {
    assert.equal(back?.getCell(ref).numFmt, '0.00', `${ref} keeps the column number format`);
  }
  assert.equal(back?.getCell('A1').fill, undefined, 'a sibling does not pick up the fill');
  assert.equal(back?.getCell('A3').fill, undefined, 'a sibling does not pick up the fill');
});

test('a column alignment round-trips and is inherited by its own cells only', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getColumn(1).alignment = {textRotation: 45};
  sheet.getCell('A1').value = 'a';
  sheet.getCell('A2').value = 'c';
  sheet.getCell('B1').value = 'b';

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(
    back?.getColumn(1).alignment?.textRotation,
    45,
    'the column keeps its own alignment',
  );
  assert.equal(
    back?.getCell('A1').alignment?.textRotation,
    45,
    'a bare cell inherits its column alignment',
  );
  assert.equal(
    back?.getCell('A2').alignment?.textRotation,
    45,
    'a second cell in the column inherits it too',
  );
  assert.equal(
    back?.getCell('B1').alignment,
    undefined,
    'a cell in another column does not inherit it',
  );
});

test('a column border applies only to its declaring column, not to later width-only columns', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getColumn(1).border = {right: {style: 'thin', color: {argb: 'FF000000'}}};
  sheet.getColumn(2).width = 10;
  sheet.getCell('A1').value = 'a';
  sheet.getCell('B1').value = 'b';
  sheet.getCell('C1').value = 'c';

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(
    back?.getCell('A1').border?.right?.style,
    'thin',
    'the declaring column’s cell has the border',
  );
  assert.equal(back?.getCell('B1').border, undefined, 'a width-only column’s cell gets no border');
  assert.equal(back?.getCell('C1').border, undefined, 'an undeclared column’s cell gets no border');
});

test('a cell overriding one facet keeps the column’s other facet default', () => {
  // The column defaults both an alignment and a number format; a cell that overrides only its
  // alignment must still carry the column's number format: the generalized column composition
  // must not drop a non-overridden facet.
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getColumn(1).alignment = {horizontal: 'center'};
  sheet.getColumn(1).numFmt = '0.00';
  sheet.getCell('A1').value = 1;
  sheet.getCell('A2').value = 2;
  sheet.getCell('A2').alignment = {horizontal: 'right'};

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(
    back?.getCell('A1').alignment?.horizontal,
    'center',
    'a bare cell inherits the column alignment',
  );
  assert.equal(
    back?.getCell('A2').alignment?.horizontal,
    'right',
    'the overriding cell keeps its own alignment',
  );
  assert.equal(
    back?.getCell('A2').numFmt,
    '0.00',
    'the overriding cell still carries the column number format',
  );
});

test('a built-in numFmt id on a foreign cell resolves to its standard format code', () => {
  // A foreign generator names a built-in format by id with no <numFmt> entry; the reader
  // resolves it from the standard table (id 10 = "0.00%").
  const pkg = foreignPackage({
    'xl/styles.xml':
      '<?xml version="1.0"?><styleSheet><fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
      '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>',
    [SHEET1]: foreignSheet('<row r="1"><c r="A1" s="1"><v>0.5</v></c></row>'),
  });
  const back = readXlsx(pkg).getWorksheet('S');
  assert.equal(back?.getCell('A1').numFmt, '0.00%');
});

test('a custom indexed-color palette survives a read → write round-trip verbatim', () => {
  const palette =
    '<colors><indexedColors><rgbColor rgb="ff000000"/><rgbColor rgb="ffffffff"/>' +
    '<rgbColor rgb="ff3f6797"/><rgbColor rgb="ffaaaaaa"/></indexedColors></colors>';
  const pkg = foreignPackage({
    'xl/styles.xml': `<?xml version="1.0"?><styleSheet><dxfs count="0"/>${palette}</styleSheet>`,
    [SHEET1]: foreignSheet('<row r="1"><c r="A1"><v>1</v></c></row>'),
  });
  const loaded = readXlsx(pkg);
  assert.deepEqual(loaded.indexedColors, [
    '<rgbColor rgb="ff000000"/>',
    '<rgbColor rgb="ffffffff"/>',
    '<rgbColor rgb="ff3f6797"/>',
    '<rgbColor rgb="ffaaaaaa"/>',
  ]);
  // The re-written styles part carries the same palette, entry for entry.
  const rewritten = partText(writeXlsx(loaded), 'xl/styles.xml');
  assert.match(
    rewritten,
    /<colors><indexedColors><rgbColor rgb="ff000000"\/>.*<rgbColor rgb="ffaaaaaa"\/><\/indexedColors><\/colors>/,
  );
});

test('a gradient fill in the <fills> list keeps its id slot so later fills still resolve', () => {
  // fills: 0 none, 1 gray125, 2 gradient, 3 solid red. A cell naming fill 3 must read back red, so the
  // gradient at index 2 must not swallow the slot and shift red down to where nothing references it.
  const styles =
    '<?xml version="1.0"?><styleSheet><fills count="4">' +
    '<fill><patternFill patternType="none"/></fill>' +
    '<fill><patternFill patternType="gray125"/></fill>' +
    '<fill><gradientFill degree="90"><stop position="0"><color rgb="FFFFFFFF"/></stop>' +
    '<stop position="1"><color rgb="FF4472C4"/></stop></gradientFill></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/><bgColor indexed="64"/></patternFill></fill>' +
    '</fills><cellXfs count="3">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0"/>' +
    '</cellXfs></styleSheet>';
  const pkg = foreignPackage({
    'xl/styles.xml': styles,
    [SHEET1]:
      '<?xml version="1.0"?><worksheet><sheetData><row r="1">' +
      '<c r="A1" s="1"><v>1</v></c><c r="A2" s="2"><v>2</v></c></row></sheetData></worksheet>',
  });
  const back = readXlsx(pkg).getWorksheet('S');
  assert.equal(
    fillFgArgb(back?.getCell('A1').fill),
    'FFFF0000',
    'the solid fill after the gradient still resolves',
  );

  const gradient = back?.getCell('A2').fill;
  assert.ok(gradient?.type === 'gradient', 'the gradient-filled cell reads back a gradient fill');
  assert.equal(gradient.gradient, 'linear');
  assert.equal(gradient.degree, 90);
  assert.deepEqual(gradient.stops, [
    {position: 0, color: {argb: 'FFFFFFFF'}},
    {position: 1, color: {argb: 'FF4472C4'}},
  ]);
});

test('a linear gradient fill on a cell survives a model → write → read round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'g';
  sheet.getCell('A1').fill = {
    type: 'gradient',
    gradient: 'linear',
    degree: 45,
    stops: [
      {position: 0, color: {argb: 'FFFF0000'}},
      {position: 0.5, color: {argb: 'FF00FF00'}},
      {position: 1, color: {argb: 'FF0000FF'}},
    ],
  };

  const back = roundtrip(wb).getWorksheet('S');
  const fill = back?.getCell('A1').fill;
  assert.ok(fill?.type === 'gradient', 'the cell keeps a gradient fill');
  assert.equal(fill.gradient, 'linear');
  assert.equal(fill.degree, 45);
  assert.deepEqual(fill.stops, [
    {position: 0, color: {argb: 'FFFF0000'}},
    {position: 0.5, color: {argb: 'FF00FF00'}},
    {position: 1, color: {argb: 'FF0000FF'}},
  ]);
});

test('a path gradient fill round-trips its inner-rectangle insets', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'p';
  sheet.getCell('A1').fill = {
    type: 'gradient',
    gradient: 'path',
    left: 0.5,
    right: 0.5,
    top: 0.5,
    bottom: 0.5,
    stops: [
      {position: 0, color: {argb: 'FFFFFFFF'}},
      {position: 1, color: {argb: 'FF000000'}},
    ],
  };

  const back = roundtrip(wb).getWorksheet('S');
  const fill = back?.getCell('A1').fill;
  assert.ok(fill?.type === 'gradient', 'the cell keeps a gradient fill');
  assert.equal(fill.gradient, 'path');
  assert.equal(fill.degree, undefined, 'a path gradient carries no degree');
  assert.deepEqual(
    {left: fill.left, right: fill.right, top: fill.top, bottom: fill.bottom},
    {
      left: 0.5,
      right: 0.5,
      top: 0.5,
      bottom: 0.5,
    },
  );
});

test('a Strict-mode t="d" cell parses to the ISO date it states, not a 1900 serial', () => {
  const pkg = foreignPackage({
    [SHEET1]: foreignSheet('<row r="1"><c r="A1" t="d"><v>2024-02-09</v></c></row>'),
  });
  const cell = readXlsx(pkg).getWorksheet('S')?.getCell('A1');
  assert.equal(cell?.type, 'date');
  assert.equal((cell.value as Date).toISOString(), '2024-02-09T00:00:00.000Z');
});

test('a serial under a built-in locale date id (57) reads as a date, not a bare number', () => {
  const pkg = foreignPackage({
    // Built-in id 57 carries no <numFmt> entry: it is resolved from the built-in table.
    'xl/styles.xml':
      '<?xml version="1.0"?><styleSheet><cellXfs count="2"><xf numFmtId="0"/>' +
      '<xf numFmtId="57" applyNumberFormat="1"/></cellXfs></styleSheet>',
    [SHEET1]: foreignSheet('<row r="1"><c r="A1" s="1"><v>45809</v></c></row>'),
  });
  const cell = readXlsx(pkg).getWorksheet('S')?.getCell('A1');
  assert.equal(cell?.type, 'date', 'a serial under a built-in date id reads as a date');
  assert.ok(cell?.numFmt, 'the built-in id resolves to a non-empty format code');
});

test('a cell font round-trips through the <fonts> table, and only the styled cell carries it', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'heading';
  sheet.getCell('A1').font = {
    bold: true,
    italic: true,
    size: 14,
    name: 'Arial',
    color: {argb: 'FF3A80D5'},
  };
  sheet.getCell('B2').value = 'plain';

  const back = roundtrip(wb).getWorksheet('S');
  const font = back?.getCell('A1').font;
  assert.deepEqual(font, {
    bold: true,
    italic: true,
    size: 14,
    color: {argb: 'FF3A80D5'},
    name: 'Arial',
  });
  // The unstyled sibling renders in the workbook default font, not A1's, so it must not pick up
  // any of A1's overrides (bold/italic/Arial/the blue), though it does resolve the default face.
  const b2 = back?.getCell('B2').font;
  assert.equal(b2?.bold, undefined, 'an unstyled sibling does not inherit the bold');
  assert.equal(b2?.italic, undefined, 'an unstyled sibling does not inherit the italic');
  assert.notEqual(b2?.name, 'Arial', 'an unstyled sibling keeps the default face, not A1 Arial');
});

test('an unstyled cell resolves to the workbook default font (a concrete face), not nothing', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'plain';

  const font = roundtrip(wb).getWorksheet('S')?.getCell('A1').font;
  assert.ok(font, 'an unstyled cell must resolve the default font, not undefined');
  assert.equal(font?.name, 'Calibri', 'the default face is Calibri');
  assert.equal(font?.size, 11, 'the default size is 11');
});

test('a cell carrying exactly the default font interns back to font id 0: no redundant entry', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  // Assign the default face explicitly; on write it must collapse to id 0, so the emitted <fonts>
  // table stays a single default entry rather than gaining a duplicate.
  sheet.getCell('A1').font = {
    size: 11,
    color: {theme: 1},
    name: 'Calibri',
    family: 2,
    scheme: 'minor',
  };

  const styles = partText(writeXlsx(wb), 'xl/styles.xml');
  const fontsBlock = elementIn(styles, /<fonts\b[^>]*>[\s\S]*?<\/fonts>/);
  assert.equal(
    (fontsBlock.match(/<font\b/g) ?? []).length,
    1,
    'only the single default font entry',
  );
});

test('an underline font round-trips: single stays single, a named variant keeps its value', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.getCell('A1').font = {underline: true};
  sheet.getCell('A2').value = 'y';
  sheet.getCell('A2').font = {underline: 'double'};

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getCell('A1').font?.underline, true, 'a bare underline reads back as true');
  assert.equal(
    back?.getCell('A2').font?.underline,
    'double',
    'a named underline keeps its variant',
  );
});

test('a foreign font’s <u val="none"/> reads back as not underlined, not the truthy string "none"', () => {
  // Real producers write <u val="none"/> for the explicit ABSENCE of an underline. Surfacing the
  // literal "none" would be truthy: a consumer’s `if (font.underline)` would mistake it for an
  // underline, so the reader must read it back falsy.
  const pkg = foreignPackage({
    'xl/styles.xml':
      '<?xml version="1.0"?><styleSheet>' +
      '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><u val="none"/></font></fonts>' +
      '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
      '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>',
    [SHEET1]: foreignSheet('<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>x</t></is></c></row>'),
  });
  const underline = readXlsx(pkg).getWorksheet('S')?.getCell('A1').font?.underline;
  assert.ok(
    !underline,
    `<u val="none"/> must read back falsy, not the truthy string "none"; got ${JSON.stringify(underline)}`,
  );
});

test('a foreign font honours an explicit-false boolean flag rather than tag presence', () => {
  // A foreign generator writes <b/> (bold on) but <i val="0"/> (italic explicitly off). The
  // reader must honour the val: a present tag is not truthy on its own.
  const pkg = foreignPackage({
    'xl/styles.xml':
      '<?xml version="1.0"?><styleSheet>' +
      '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><i val="0"/></font></fonts>' +
      '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
      '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>',
    [SHEET1]: foreignSheet('<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>x</t></is></c></row>'),
  });
  const font = readXlsx(pkg).getWorksheet('S')?.getCell('A1').font;
  assert.equal(font?.bold, true, 'a bare <b/> is bold');
  assert.equal(font?.italic, false, '<i val="0"/> is explicitly not italic: the val is honoured');
});

test('a cell border round-trips through the <borders> table, and only the styled cell carries it', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'boxed';
  sheet.getCell('A1').border = {
    top: {style: 'thin'},
    bottom: {style: 'medium', color: {argb: 'FF3A80D5'}},
  };
  sheet.getCell('B2').value = 'plain';

  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual(back?.getCell('A1').border, {
    top: {style: 'thin'},
    bottom: {style: 'medium', color: {argb: 'FF3A80D5'}},
  });
  assert.equal(back?.getCell('B2').border, undefined, 'an unbordered sibling stays borderless');
});

test('a cell bordered on one side does not fabricate the other three across a round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.getCell('A1').border = {top: {style: 'thin'}};

  const border = roundtrip(wb).getWorksheet('S')?.getCell('A1').border;
  assert.equal(border?.top?.style, 'thin', 'the declared top edge survives');
  for (const side of ['left', 'right', 'bottom', 'diagonal'] as const) {
    assert.equal(border?.[side], undefined, `${side} is not fabricated`);
  }
});

test('a foreign diagonal border reads its edge and diagonal direction', () => {
  // A foreign generator declares a diagonal border with diagonalUp; the reader must carry the
  // edge and the direction flag rather than dropping either.
  const pkg = foreignPackage({
    'xl/styles.xml':
      '<?xml version="1.0"?><styleSheet>' +
      '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
      '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
      '<border diagonalUp="1"><left/><right/><top/><bottom/><diagonal style="thin"/></border></borders>' +
      '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/></cellXfs></styleSheet>',
    [SHEET1]: foreignSheet('<row r="1"><c r="A1" s="1" t="inlineStr"><is><t>x</t></is></c></row>'),
  });
  const border = readXlsx(pkg).getWorksheet('S')?.getCell('A1').border;
  assert.equal(border?.diagonal?.style, 'thin', 'the diagonal edge survives');
  assert.equal(border?.diagonalUp, true, 'the diagonalUp direction is honoured');
  assert.equal(border?.top, undefined, 'a styleless edge is not fabricated');
});

test('cell alignment round-trips through the xf, and only the aligned cell carries it', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.getCell('A1').alignment = {
    horizontal: 'center',
    vertical: 'top',
    wrapText: true,
    indent: 2,
    textRotation: 45,
  };
  sheet.getCell('B2').value = 'plain';

  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual(back?.getCell('A1').alignment, {
    horizontal: 'center',
    vertical: 'top',
    wrapText: true,
    indent: 2,
    textRotation: 45,
  });
  assert.equal(
    back?.getCell('B2').alignment,
    undefined,
    'an unaligned sibling stays alignment-free',
  );
});

test('alignment boolean flags left off do not read back spuriously enabled', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.getCell('A1').alignment = {wrapText: false, shrinkToFit: false};

  // An all-default alignment serialises to nothing, so it resolves to xf 0 and reads back absent.
  assert.equal(roundtrip(wb).getWorksheet('S')?.getCell('A1').alignment, undefined);
});

test('a foreign alignment carrying only wrapText="0" reads back with no alignment', () => {
  // Excel writes an all-false alignment as wrapText="0"; the raw "0" is a truthy JS string, so a
  // reader that guards on presence rather than the parsed boolean would mistake it for present.
  const pkg = foreignPackage({
    'xl/styles.xml':
      '<?xml version="1.0"?><styleSheet>' +
      '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
      '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">' +
      '<alignment wrapText="0"/></xf></cellXfs></styleSheet>',
    [SHEET1]: foreignSheet('<row r="1"><c r="A1" s="0" t="inlineStr"><is><t>x</t></is></c></row>'),
  });
  assert.equal(readXlsx(pkg).getWorksheet('S')?.getCell('A1').alignment, undefined);
});

test('cell protection round-trips the meaningful flags, and only the protected cell carries them', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.getCell('A1').protection = {locked: false, hidden: true};
  sheet.getCell('B2').value = 'plain';

  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual(back?.getCell('A1').protection, {locked: false, hidden: true});
  assert.equal(
    back?.getCell('B2').protection,
    undefined,
    'a default-protection sibling stays protection-free',
  );
});

test('a default-locked cell does not read back as explicitly protected', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  // locked defaults to on, hidden to off, so this restates the default and serialises to nothing.
  sheet.getCell('A1').protection = {locked: true, hidden: false};

  assert.equal(roundtrip(wb).getWorksheet('S')?.getCell('A1').protection, undefined);
});

test('a foreign <protection locked="1">, an explicit default, reads back with no protection', () => {
  // A foreign generator states the default explicitly (locked on). Since locked defaults true,
  // that carries no information; the reader must not fabricate a { locked: true } protection.
  const pkg = foreignPackage({
    'xl/styles.xml':
      '<?xml version="1.0"?><styleSheet>' +
      '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
      '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1">' +
      '<protection locked="1"/></xf></cellXfs></styleSheet>',
    [SHEET1]: foreignSheet('<row r="1"><c r="A1" s="0" t="inlineStr"><is><t>x</t></is></c></row>'),
  });
  assert.equal(readXlsx(pkg).getWorksheet('S')?.getCell('A1').protection, undefined);
});

test('the inflate bound rejects a part whose declared size is over the cap', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const buffer = writeXlsx(wb);
  assert.throws(() => readXlsx(buffer, {maxUncompressedBytes: 16}), /possible zip bomb/);
});

test('a zip that is not an xlsx (no workbook part) is rejected, not misread', () => {
  const bogus = zipSync({'hello.txt': strToU8('not a spreadsheet')});
  assert.throws(
    () => readXlsx(bogus),
    (err: unknown) =>
      err instanceof UnsupportedFormatError &&
      err.format === 'unknown' &&
      /not a valid \.xlsx package/.test(err.message),
  );
});

test('a t="s" shared-string cell resolves against the shared table', () => {
  // Our writer emits inlineStr, but the reader must also resolve shared strings that
  // foreign generators use. Assemble a minimal package by hand to exercise that path.
  const pkg = foreignPackage({
    'xl/sharedStrings.xml': '<?xml version="1.0"?><sst><si><t>shared</t></si></sst>',
    [SHEET1]: foreignSheet('<row r="1"><c r="A1" t="s"><v>0</v></c></row>'),
  });
  const back = readXlsx(pkg).getWorksheet('S');
  assert.equal(back?.getCell('A1').value, 'shared');
});

test('a password-protected sheet round-trips its credential and permissive flags', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 'x';
  ws.protect('secret', {sort: true, autoFilter: true, selectLockedCells: false});

  const back = roundtrip(wb).getWorksheet('S');
  const protection = back?.protection;
  assert.ok(protection, 'the reloaded sheet reports protection');

  // The permissive flags the author set survive; a default-valued flag stays absent.
  assert.equal(protection.flags.sort, true);
  assert.equal(protection.flags.autoFilter, true);
  assert.equal(protection.flags.selectLockedCells, false);
  assert.equal(protection.flags.deleteRows, undefined);

  // The agile credential is preserved verbatim: the reader cannot (and must not) re-hash it.
  const original = ws.protection?.credential;
  assert.deepEqual(protection.credential, original);
  assert.equal(protection.credential?.algorithmName, 'SHA-512');
});

test('a passwordless protected sheet round-trips as protected with no credential', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 'x';
  ws.protect();

  const protection = roundtrip(wb).getWorksheet('S')?.protection;
  assert.ok(protection, 'a soft (passwordless) lock still round-trips as protection');
  assert.equal(protection.credential, undefined);
});

// The 16-bit `password` hash is what pre-2010 Excel, XlsxWriter, openpyxl and LibreOffice write, and it
// had no field: a sheet protected that way was rewritten still marked protected and without a password,
// so anyone could unprotect it. And `sheet` defaults to false, so an element without it was locking a
// sheet the file had left open.
function withSheetProtection(element: string): Uint8Array {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  return patchParts(writeXlsx(wb), {
    [SHEET1]: (xml) => xml.replace('</sheetData>', `</sheetData>${element}`),
  });
}

test('a legacy password hash a file carries survives a round-trip verbatim', () => {
  const read = readXlsx(
    withSheetProtection('<sheetProtection password="CC3D" sheet="1" objects="1" scenarios="1"/>'),
  );
  const protection = read.getWorksheet('S')?.protection;
  assert.ok(protection, 'the sheet reads as protected');
  assert.equal(protection.legacyPasswordHash, 'CC3D');
  assert.equal(protection.credential, undefined);
  const rewritten = elementIn(sheetXml(writeXlsx(read)), /<sheetProtection\b[^>]*\/>/);
  assert.match(rewritten, /password="CC3D"/);
  assert.match(rewritten, /sheet="1"/);
});

test('a malformed legacy password hash is dropped, and the protection it came with is kept', () => {
  const read = readXlsx(withSheetProtection('<sheetProtection password="not-a-hash" sheet="1"/>'));
  const protection = read.getWorksheet('S')?.protection;
  assert.ok(protection);
  assert.equal(protection.legacyPasswordHash, undefined);
  assert.doesNotMatch(sheetXml(writeXlsx(read)), /password=/);
});

test('a sheetProtection element without sheet="1" records an unprotected sheet', () => {
  const read = readXlsx(withSheetProtection('<sheetProtection objects="1" scenarios="1"/>'));
  assert.equal(read.getWorksheet('S')?.protection, undefined);
  assert.doesNotMatch(sheetXml(writeXlsx(read)), /<sheetProtection/, 'and a save does not lock it');
});

test('an unprotected sheet reports no protection after a round-trip', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  assert.equal(roundtrip(wb).getWorksheet('S')?.protection, undefined);
});

test('a re-written loaded protection preserves the credential byte-for-byte', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 'x';
  ws.protect('pw', {formatCells: true});

  const firstXml = sheetXml(writeXlsx(wb));
  const secondXml = sheetXml(writeXlsx(roundtrip(wb)));
  const prot = (xml: string): string => (xml.match(/<sheetProtection\b[^>]*\/>/) ?? [''])[0];

  assert.ok(prot(firstXml), 'the first write emits a sheetProtection element');
  assert.equal(
    prot(secondXml),
    prot(firstXml),
    'a passthrough save re-emits the identical protection',
  );
});

test('a Date value round-trips as a Date on the calendar day it was written', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = new Date('2020-03-04T00:00:00.000Z');

  const back = roundtrip(wb).getWorksheet('S')?.getCell('A1');
  assert.equal(back?.type, 'date');
  assert.ok(back?.value instanceof Date);
  assert.equal(back.value.toISOString(), '2020-03-04T00:00:00.000Z');
});

test('a bare Date is written under a date number format so it reads back as a date', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = new Date('2020-03-04T00:00:00.000Z');

  const pkg = writeXlsx(wb);
  const cell = elementIn(sheetXml(pkg), /<c r="A1"[^>]*>/);
  assert.doesNotMatch(cell, /t="/, 'a date serial is a plain number cell, no t=');

  const styles = partText(pkg, 'xl/styles.xml');
  const xfIndex = Number(captureIn(cell, /\bs="(\d+)"/, 'the cell style index'));
  const cellXfs = elementIn(styles, /<cellXfs\b[\s\S]*?<\/cellXfs>/);
  const xf = [...cellXfs.matchAll(/<xf\b[^>]*>/g)][xfIndex]?.[0];
  assert.ok(xf, `the cell names xf ${xfIndex}, which cellXfs must hold`);
  const numFmtId = captureIn(xf, /\bnumFmtId="(\d+)"/, 'the xf number format');
  const formatCode = captureIn(
    styles,
    new RegExp(`<numFmt numFmtId="${numFmtId}" formatCode="([^"]*)"`),
    `the format code for numFmtId ${numFmtId}`,
  );
  assert.ok(isDateFormat(formatCode), `${formatCode} is a date format`);
});

test('an explicit date numFmt on a Date cell survives verbatim, not swapped for the default', () => {
  const wb = new Workbook();
  const cell = wb.addWorksheet('S').getCell('A1');
  cell.value = new Date('2020-03-04T00:00:00.000Z');
  cell.numFmt = 'DD/MM/YYYY';

  const back = roundtrip(wb).getWorksheet('S')?.getCell('A1');
  assert.equal(back?.numFmt, 'DD/MM/YYYY');
  assert.equal(back?.type, 'date');
});

test('a serial under a non-date format reads back as a plain number, not a date', () => {
  const wb = new Workbook();
  const cell = wb.addWorksheet('S').getCell('A1');
  cell.value = 43_894; // the serial for 2020-03-04
  cell.numFmt = '0.00';

  const back = roundtrip(wb).getWorksheet('S')?.getCell('A1');
  assert.equal(back?.type, 'number');
  assert.equal(back?.value, 43_894);
});

test('an Invalid Date does not throw on write and does not drop sibling cells', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = new Date(Number.NaN);
  sheet.getCell('B1').value = 'still here';
  sheet.getCell('C1').value = 42;

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.getCell('B1').value, 'still here');
  assert.equal(back?.getCell('C1').value, 42);
  assert.equal(back?.getCell('A1').value, null, 'the invalid date carries no serial');
});

test('a worksheet tab colour round-trips as the exact 8-digit ARGB', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.tabColor = {argb: 'FFFF0000'};
  sheet.getCell('A1').value = 'x';

  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual(back?.tabColor, {argb: 'FFFF0000'});
});

test('the written package carries the tab colour under <sheetPr>', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').tabColor = {argb: 'FF00FF00'};

  const xml = sheetXml(writeXlsx(wb));
  assert.match(xml, /<sheetPr><tabColor rgb="FF00FF00"\/><\/sheetPr>/);
  // <sheetPr> must lead the worksheet, before <dimension>.
  assert.ok(xml.indexOf('<sheetPr>') < xml.indexOf('<dimension'), 'sheetPr precedes dimension');
});

test('a sheet with no tab colour acquires none and emits no <sheetPr>', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'y';

  const xml = sheetXml(writeXlsx(wb));
  assert.doesNotMatch(xml, /<sheetPr>/);
  assert.equal(roundtrip(wb).getWorksheet('S')?.tabColor, undefined);
});

test('a theme-based tab colour with a tint round-trips', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').tabColor = {theme: 4, tint: -0.25};

  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual(back?.tabColor, {theme: 4, tint: -0.25});
});

test('a tab colour survives a worksheet model export/import', () => {
  const wb = new Workbook();
  const src = wb.addWorksheet('Src');
  src.tabColor = {argb: 'FF0000FF'};
  const dst = wb.addWorksheet('Dst');
  dst.model = src.model;
  assert.deepEqual(dst.tabColor, {argb: 'FF0000FF'});
});

test('inverted outline summary positions round-trip as false', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.outline.summaryBelow = false;
  sheet.outline.summaryRight = false;
  sheet.getCell('A1').value = 'x';

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.outline.summaryBelow, false);
  assert.equal(back?.outline.summaryRight, false);
});

test('the written package carries the outline flags under <sheetPr>', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.outline.summaryBelow = false;
  sheet.outline.summaryRight = false;

  const xml = sheetXml(writeXlsx(wb));
  assert.match(xml, /<sheetPr><outlinePr summaryBelow="0" summaryRight="0"\/><\/sheetPr>/);
});

test('the tab colour and outline flags share one <sheetPr> in CT_SheetPr order', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.tabColor = {argb: 'FFFF0000'};
  sheet.outline.summaryBelow = false;

  const xml = sheetXml(writeXlsx(wb));
  assert.match(
    xml,
    /<sheetPr><tabColor rgb="FFFF0000"\/><outlinePr summaryBelow="0"\/><\/sheetPr>/,
  );
});

test('only the set outline flag serializes; the untouched one stays absent', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').outline.summaryRight = false;

  const xml = sheetXml(writeXlsx(wb));
  assert.match(xml, /<outlinePr summaryRight="0"\/>/);
  assert.doesNotMatch(xml, /summaryBelow/);
});

test('a sheet with default outline positions emits no <outlinePr>', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'y';

  const xml = sheetXml(writeXlsx(wb));
  assert.doesNotMatch(xml, /<outlinePr/);
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.outline.summaryBelow, undefined);
  assert.equal(back?.outline.summaryRight, undefined);
});

// These optional booleans were read with `boolPresent`, which takes `"yes"` as present-and-true, so a
// token the schema does not allow came back out of the writer as `"1"`.
test('an unrecognised boolean token on outlinePr or pageSetUpPr is dropped, not written as "1"', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 1;
  sheet.outline.summaryBelow = false;
  sheet.pageSetup.fitToPage = true;
  const pkg = patchParts(writeXlsx(wb), {
    [SHEET1]: (xml) =>
      xml
        .replace('summaryBelow="0"', 'summaryBelow="yes"')
        .replace('fitToPage="1"', 'fitToPage="yes"'),
  });
  const back = readXlsx(pkg);
  assert.equal(back.getWorksheet('S')?.outline.summaryBelow, undefined);
  assert.equal(back.getWorksheet('S')?.pageSetup.fitToPage, undefined);
  const rewritten = sheetXml(writeXlsx(back));
  assert.doesNotMatch(rewritten, /summaryBelow=|fitToPage=/, 'neither is written back at all');
});

test('outline flags survive a worksheet model export/import', () => {
  const wb = new Workbook();
  const src = wb.addWorksheet('Src');
  src.outline.summaryBelow = false;
  const dst = wb.addWorksheet('Dst');
  dst.model = src.model;
  assert.equal(dst.outline.summaryBelow, false);
});

test('a fit-to-page setup round-trips its flag, counts, and scale', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.pageSetup.fitToPage = true;
  sheet.pageSetup.fitToWidth = 1;
  sheet.pageSetup.fitToHeight = 0;
  sheet.pageSetup.scale = 80;
  sheet.getCell('A1').value = 'x';

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.pageSetup.fitToPage, true);
  assert.equal(back?.pageSetup.fitToWidth, 1);
  assert.equal(back?.pageSetup.fitToHeight, 0);
  assert.equal(back?.pageSetup.scale, 80);
});

test('the fit-to-page flag rides <pageSetUpPr> under <sheetPr>, the counts ride <pageSetup>', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.pageSetup.fitToPage = true;
  sheet.pageSetup.fitToWidth = 1;

  const xml = sheetXml(writeXlsx(wb));
  assert.match(xml, /<sheetPr><pageSetUpPr fitToPage="1"\/><\/sheetPr>/);
  assert.match(xml, /<pageSetup fitToWidth="1"\/>/);
});

test('<pageSetUpPr> follows <outlinePr> under <sheetPr> in CT_SheetPr order', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.outline.summaryBelow = false;
  sheet.pageSetup.fitToPage = true;

  const xml = sheetXml(writeXlsx(wb));
  assert.match(
    xml,
    /<sheetPr><outlinePr summaryBelow="0"\/><pageSetUpPr fitToPage="1"\/><\/sheetPr>/,
  );
});

test('<pageSetup> sits between <pageMargins> and <headerFooter>', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.pageMargins.top = 1;
  sheet.pageSetup.scale = 90;
  sheet.headerFooter.oddHeader = 'H';

  const xml = sheetXml(writeXlsx(wb));
  assert.ok(
    xml.indexOf('<pageMargins') < xml.indexOf('<pageSetup'),
    'pageSetup follows pageMargins',
  );
  assert.ok(
    xml.indexOf('<pageSetup') < xml.indexOf('<headerFooter'),
    'pageSetup precedes headerFooter',
  );
});

test('orientation and pageOrder round-trip and emit only when set', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.pageSetup.orientation = 'landscape';
  sheet.pageSetup.pageOrder = 'overThenDown';

  const xml = sheetXml(writeXlsx(wb));
  assert.match(xml, /<pageSetup pageOrder="overThenDown" orientation="landscape"\/>/);
  assert.doesNotMatch(xml, /scale=|fitToWidth=|fitToHeight=/);

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.pageSetup.orientation, 'landscape');
  assert.equal(back?.pageSetup.pageOrder, 'overThenDown');
});

test('paperSize round-trips and leads the <pageSetup> attributes', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.pageSetup.paperSize = 9;
  sheet.pageSetup.scale = 96;

  const xml = sheetXml(writeXlsx(wb));
  assert.match(xml, /<pageSetup paperSize="9" scale="96"\/>/);

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.pageSetup.paperSize, 9);
});

test('an opaque printer-settings blob round-trips its exact bytes', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  const blob = new Uint8Array([0x00, 0x01, 0xff, 0x7f, 0x80, 0x42]);
  sheet.pageSetup.printerSettings = blob;

  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual(back?.pageSetup.printerSettings, blob, 'the DEVMODE bytes survive verbatim');
});

test('a printer-settings blob wires up the r:id, the .bin part, its rel, and a content type', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').pageSetup.printerSettings = new Uint8Array([1, 2, 3]);

  const pkg = writeXlsx(wb);
  // The blob is the only reason the element exists, so <pageSetup> emits carrying just the r:id.
  assert.match(sheetXml(pkg), /<pageSetup r:id="rId1"\/>/);

  assert.deepEqual(
    partBytes(pkg, 'xl/printerSettings/printerSettings1.bin'),
    new Uint8Array([1, 2, 3]),
  );

  const rels = partText(pkg, 'xl/worksheets/_rels/sheet1.xml.rels');
  assert.match(rels, /Id="rId1"[^>]*Target="\.\.\/printerSettings\/printerSettings1\.bin"/);
  assert.match(rels, /Type="[^"]*\/printerSettings"/);

  const contentTypes = partText(pkg, '[Content_Types].xml');
  assert.match(contentTypes, /<Default Extension="bin" ContentType="[^"]*printerSettings"\/>/);
});

test('a printer-settings blob rides alongside a table without stealing its rel id', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'h';
  sheet.getCell('A2').value = 'v';
  sheet.addTable({name: 'T', ref: 'A1', columns: [{name: 'h'}], rowCount: 1});
  sheet.pageSetup.printerSettings = new Uint8Array([9]);

  const pkg = writeXlsx(wb);
  const rels = partText(pkg, 'xl/worksheets/_rels/sheet1.xml.rels');
  // The table keeps rId1; the printer-settings blob follows it at rId2, so neither reference collides.
  assert.match(rels, /Id="rId1"[^>]*Target="\.\.\/tables\/table1\.xml"/);
  assert.match(rels, /Id="rId2"[^>]*Target="\.\.\/printerSettings\/printerSettings1\.bin"/);
  assert.match(sheetXml(pkg), /<pageSetup r:id="rId2"\/>/);
});

test('a sheet with no printer settings writes no .bin part and no r:id', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').pageSetup.scale = 90;

  const pkg = writeXlsx(wb);
  assert.doesNotMatch(sheetXml(pkg), /r:id=/);
  assert.equal(optionalPartText(pkg, 'xl/printerSettings/printerSettings1.bin'), undefined);
  assert.doesNotMatch(partText(pkg, '[Content_Types].xml'), /Extension="bin"/);
});

test('a non-numeric paperSize is dropped on read, not stored as NaN', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').pageSetup.scale = 96;
  const patched = patchParts(writeXlsx(wb), {
    [SHEET1]: (xml) => xml.replace('<pageSetup ', '<pageSetup paperSize="A4" '),
  });
  const back = readXlsx(patched).getWorksheet('S');
  assert.equal(back?.pageSetup.paperSize, undefined);
  assert.equal(back?.pageSetup.scale, 96);
});

test('a sheet with no page setup emits neither <pageSetUpPr> nor <pageSetup>', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'y';

  const xml = sheetXml(writeXlsx(wb));
  assert.doesNotMatch(xml, /<pageSetUpPr/);
  assert.doesNotMatch(xml, /<pageSetup/);
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.pageSetup.fitToPage, undefined);
  assert.equal(back?.pageSetup.scale, undefined);
});

test('a <pageSetUpPr> present only for other reasons leaves fitToPage unset', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'y';
  const patched = patchParts(writeXlsx(wb), {
    [SHEET1]: (xml) =>
      xml.replace('<dimension', '<sheetPr><pageSetUpPr autoPageBreaks="0"/></sheetPr><dimension'),
  });
  const back = readXlsx(patched).getWorksheet('S');
  assert.equal(back?.pageSetup.fitToPage, undefined);
});

test('page setup survives a worksheet model export/import', () => {
  const wb = new Workbook();
  const src = wb.addWorksheet('Src');
  src.pageSetup.fitToPage = true;
  src.pageSetup.scale = 75;
  const dst = wb.addWorksheet('Dst');
  dst.model = src.model;
  assert.equal(dst.pageSetup.fitToPage, true);
  assert.equal(dst.pageSetup.scale, 75);
});

test("a sheet-list entry's state attribute is read back as the worksheet visibility state", () => {
  const wb = new Workbook();
  wb.addWorksheet('Visible').getCell('A1').value = 'v';
  wb.addWorksheet('Hid', {state: 'hidden'}).getCell('A1').value = 'h';
  wb.addWorksheet('VeryHid', {state: 'veryHidden'}).getCell('A1').value = 'y';
  const back = roundtrip(wb);
  assert.equal(back.getWorksheet('Visible')?.state, 'visible');
  assert.equal(back.getWorksheet('Hid')?.state, 'hidden');
  assert.equal(back.getWorksheet('VeryHid')?.state, 'veryHidden');
});

test('an unknown state attribute on a sheet-list entry falls back to visible, not the raw string', () => {
  const pkg = foreignPackage({
    'xl/workbook.xml':
      '<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="S" sheetId="1" state="bogus" r:id="rId1"/></sheets></workbook>',
  });
  assert.equal(readXlsx(pkg).getWorksheet('S')?.state, 'visible');
});

test('workbook structure protection survives a read→write round-trip', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  wb.protection = {lockStructure: true, lockWindows: false};
  const wbXml = partText(writeXlsx(wb), 'xl/workbook.xml');
  assert.match(wbXml, /<workbookProtection lockStructure="1"\/>/);
  // The element precedes <sheets> in CT_Workbook order.
  assert.ok(wbXml.indexOf('<workbookProtection') < wbXml.indexOf('<sheets>'));

  const back = roundtrip(wb);
  assert.equal(back.protection?.lockStructure, true);
  assert.equal(back.protection?.lockWindows, undefined);
});

test('a workbook with no protection emits no <workbookProtection> and reads back undefined', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const wbXml = partText(writeXlsx(wb), 'xl/workbook.xml');
  assert.doesNotMatch(wbXml, /workbookProtection/);
  assert.equal(roundtrip(wb).protection, undefined);
});

test('a workbook protection password/hash credential is preserved verbatim across a round-trip', () => {
  const pkg = foreignPackage({
    'xl/workbook.xml':
      '<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<workbookProtection workbookAlgorithmName="SHA-512" workbookHashValue="aGFzaA==" ' +
      'workbookSaltValue="c2FsdA==" workbookSpinCount="100000" lockStructure="1"/>' +
      '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>',
  });
  const back = readXlsx(pkg);
  assert.equal(back.protection?.lockStructure, true);
  assert.equal(back.protection?.credentials?.workbookAlgorithmName, 'SHA-512');
  assert.equal(back.protection?.credentials?.workbookSpinCount, '100000');

  const wbXml = partText(writeXlsx(back), 'xl/workbook.xml');
  assert.match(wbXml, /workbookAlgorithmName="SHA-512"/);
  assert.match(wbXml, /workbookHashValue="aGFzaA=="/);
  assert.match(wbXml, /workbookSpinCount="100000"/);
});

test('an unknown attribute on <workbookProtection> is dropped, not echoed back on write', () => {
  const pkg = foreignPackage({
    'xl/workbook.xml':
      '<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<workbookProtection lockStructure="1" bogusAttr="evil"/>' +
      '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>',
  });
  const wbXml = partText(writeXlsx(readXlsx(pkg)), 'xl/workbook.xml');
  assert.match(wbXml, /lockStructure="1"/);
  assert.doesNotMatch(wbXml, /bogusAttr/);
});

// --- the workbook default font, end to end ------------------------------------------------------

// The `<fonts>` table of a written package, as one string.
function fontsBlockOf(workbook: Workbook): string {
  const styles = partText(writeXlsx(workbook), 'xl/styles.xml');
  return elementIn(styles, /<fonts\b[^>]*>[\s\S]*?<\/fonts>/);
}

test("an authored theme's body face reaches font 0, so it reaches every unstyled cell", () => {
  // The reported symptom: setTheme wrote the theme part correctly while styles.xml went on claiming
  // `scheme="minor"` *and* naming Calibri outright. Excel resolves the explicit name, so every
  // unstyled cell rendered Calibri and callers had to set `font` on every column to work around it.
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'plain';
  wb.setTheme({fonts: {major: 'Aptos Display', minor: 'Aptos'}});

  assert.equal(
    fontsBlockOf(wb),
    '<fonts count="1"><font><sz val="11"/><color theme="1"/>' +
      '<name val="Aptos"/><family val="2"/><scheme val="minor"/></font></fonts>',
    'font 0 names the theme body face and its scheme claim is true',
  );
});

test('a foreign non-Calibri font 0 survives a round-trip without gaining a duplicate', () => {
  // Read→write used to replace the declared default with Calibri and re-add the real face as a
  // redundant custom entry: populated cells kept it via that entry, while empty cells (and any
  // consumer reading font 0 as "the workbook default") silently got Calibri.
  const seed = new Workbook();
  seed.addWorksheet('S').getCell('A1').value = 'plain';
  const patched = patchParts(writeXlsx(seed), {
    'xl/styles.xml': (xml) => xml.replace('<name val="Calibri"/>', '<name val="Aptos"/>'),
  });

  const back = readXlsx(patched);
  assert.equal(
    back.declaredDefaultFont?.name,
    'Aptos',
    'the declared face is surfaced on the model',
  );
  assert.equal(
    fontsBlockOf(back),
    '<fonts count="1"><font><sz val="11"/><color theme="1"/>' +
      '<name val="Aptos"/><family val="2"/><scheme val="minor"/></font></fonts>',
    'one entry, still Aptos: no Calibri, no duplicate',
  );
});

test('an authored default font moves the cells that only inherited the file’s, not those that named a face', () => {
  const seed = new Workbook();
  const seedSheet = seed.addWorksheet('S');
  seedSheet.getCell('A1').value = 'inherits the default';
  // A fill forces a real xf that still names font 0, so the reader flattens the default face onto
  // this cell even though the source file never said anything about its font.
  seedSheet.getCell('A2').value = 'filled, but says nothing about its font';
  seedSheet.getCell('A2').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFEEEEEE'}};
  seedSheet.getCell('A3').value = 'names a face outright';
  seedSheet.getCell('A3').font = {name: 'Courier New', size: 9};

  const back = roundtrip(seed);
  back.setDefaultFont({name: 'Georgia', size: 12});

  const fonts = fontsBlockOf(back);
  assert.match(fonts, /<font><sz val="12"\/><color theme="1"\/><name val="Georgia"\/><\/font>/);
  // Georgia is not the theme body face, so the `scheme="minor"` claim is dropped rather than left to
  // contradict the name it sits beside.
  assert.doesNotMatch(fonts, /scheme/);
  // The filled cell inherited; it must follow the new default rather than pin itself to the old face.
  assert.doesNotMatch(fonts, /Calibri/);
  // The cell that named a face keeps it: that was an authored intent, not an inherited default.
  assert.match(fonts, /<name val="Courier New"\/>/);
});

// Five readers want the worksheet part, and it is the largest in the package. Reading it once each
// spent 45% of a large file's read on four scans that matched no element, so they share one parse.
//
// The guard counts the scanning rather than timing it. It used to be a wall-clock ratio between the
// two arrangements, taken as the fastest of several runs to duck the collector; what that could not
// duck was another process, and one `verify` run failed it with a neighbouring gate on the CPU. A
// gate that fails on scheduling teaches an agent to re-run rather than to read.
//
// What is counted is every read the scanner makes of the source, by handing it a source that reports
// them. The unit is arbitrary - it is neither characters nor tags - but it is *exactly* proportional
// to how many times the document was walked, which is the only thing this test is about: five
// separate parses read five times what one shared parse reads, and any other ratio means the passes
// are not sharing a scan.
function countingSource(xml: string): {source: string; reads: () => number} {
  let reads = 0;
  const target = new String(xml);
  const source = new Proxy(target, {
    get(boxed, property, receiver) {
      reads++;
      const value: unknown = Reflect.get(boxed, property, receiver);
      // Bound to the boxed string, so a method the scanner calls does its own internal reads against
      // the target rather than back through this trap: what is counted is the scanner's reads, not
      // the standard library's.
      return typeof value === 'function'
        ? (value as (...args: never[]) => unknown).bind(boxed)
        : value;
    },
  });
  return {source: source as unknown as string, reads: () => reads};
}

test('the worksheet part is read in one pass, not once per reader', () => {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('S');
  for (let row = 1; row <= 500; row++) sheet.addRow([`a${row}`, row, row * 2, 'text', row / 3]);
  const xml = sheetXml(writeXlsx(workbook));

  const apart = countingSource(xml);
  parseXmlPasses(apart.source, [worksheetPass(new Worksheet('S', 1), [], [], 1900)]);
  parseXmlPasses(apart.source, [sheetHyperlinkPass()]);
  parseXmlPasses(apart.source, [dataValidationPass()]);
  parseXmlPasses(apart.source, [extendedDataValidationPass()]);
  parseXmlPasses(apart.source, [conditionalFormattingPass()]);

  const together = countingSource(xml);
  parseXmlPasses(together.source, [
    worksheetPass(new Worksheet('S', 1), [], [], 1900),
    sheetHyperlinkPass(),
    dataValidationPass(),
    extendedDataValidationPass(),
    conditionalFormattingPass(),
  ]);

  assert.equal(
    apart.reads(),
    together.reads() * 5,
    `five separate parses read the part ${apart.reads()} times against ${together.reads()} for one shared parse: ` +
      'that is not five times one, so the passes are not sharing a scan',
  );
});
