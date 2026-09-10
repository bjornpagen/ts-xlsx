import assert from 'node:assert/strict';
import {test} from 'node:test';

import {MAX_ROW} from '../../core/address.ts';
import {isFormulaValue} from '../../core/value.ts';
import {Workbook} from '../../core/workbook.ts';
import {foreignPackage, foreignSheet, SHEET1} from './package.test-support.ts';
import {readSheetRows, readWorkbookStream, type StreamedRow} from './read-rows.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

function rows(data: Uint8Array, options?: Parameters<typeof readSheetRows>[1]): StreamedRow[] {
  return [...readSheetRows(data, options)];
}

// xf 0 is the format of every cell nothing else formats, and a workbook can make it a date format. The
// streaming reader stopped short of it, so the same bare cell was a Date through readXlsx and a serial
// number through readSheetRows.
test('a cell with no format of its own decodes under xf 0 in both readers, and reports no style', () => {
  const data = foreignPackage({
    'xl/styles.xml':
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<cellXfs count="1"><xf numFmtId="14"/></cellXfs></styleSheet>',
    [SHEET1]: foreignSheet('<row r="1"><c r="A1"><v>45000</v></c></row>'),
  });
  const buffered = readXlsx(data).worksheets[0]?.getCell('A1').value;
  assert.ok(buffered instanceof Date, 'the buffered read is a date');
  const [cell] = rows(data)[0]?.cells ?? [];
  assert.ok(cell);
  assert.ok(cell.value instanceof Date, 'and so is the streamed one');
  assert.equal(cell.value.getTime(), buffered.getTime());
  assert.equal(cell.style, undefined, 'xf 0 is a floor, not a format the cell declared');
});

test('yields rows in order, non-empty cells only, with decoded values', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'name';
  sheet.getCell('B1').value = 'qty';
  sheet.getCell('A2').value = 'widget';
  sheet.getCell('B2').value = 42;
  sheet.getCell('B3').value = true;

  const streamed = rows(writeXlsx(wb));
  assert.deepEqual(
    streamed.map((row) => row.number),
    [1, 2, 3],
  );
  assert.deepEqual(
    streamed[0]?.cells.map((cell) => [cell.address, cell.value]),
    [
      ['A1', 'name'],
      ['B1', 'qty'],
    ],
  );
  assert.deepEqual(
    streamed[1]?.cells.map((cell) => [cell.col, cell.value]),
    [
      [1, 'widget'],
      [2, 42],
    ],
  );
  // Row 3 has only B3: A3 was never written, so it is absent, not a null cell.
  assert.deepEqual(
    streamed[2]?.cells.map((cell) => [cell.address, cell.value]),
    [['B3', true]],
  );
});

test('falsy-but-present values survive; only a truly empty cell drops', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 0;
  sheet.getCell('B1').value = false;
  sheet.getCell('C1').value = '';
  // D1 given a style but no value: a blank the data read should omit.
  sheet.getCell('D1').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF0000'}};

  const [row] = rows(writeXlsx(wb));
  assert.deepEqual(
    row?.cells.map((cell) => [cell.address, cell.value]),
    [
      ['A1', 0],
      ['B1', false],
      ['C1', ''],
    ],
  );
});

test('row numbers are preserved across a gap', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'top';
  sheet.getCell('A5').value = 'bottom';

  assert.deepEqual(
    rows(writeXlsx(wb)).map((row) => [row.number, row.cells[0]?.value]),
    [
      [1, 'top'],
      [5, 'bottom'],
    ],
  );
});

test('a formula cell yields a formula value with its cached result', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: 'SUM(B1:B2)', result: 7};

  const value = rows(writeXlsx(wb))[0]?.cells[0]?.value;
  assert.ok(value !== undefined);
  assert.ok(isFormulaValue(value));
  assert.equal(value.formula, 'SUM(B1:B2)');
  assert.equal(value.result, 7);
});

test('a modern function is unmangled back to its plain name', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: 'XLOOKUP(B1,C:C,D:D)', result: 0};

  const value = rows(writeXlsx(wb))[0]?.cells[0]?.value;
  assert.ok(value !== undefined);
  assert.ok(isFormulaValue(value));
  assert.equal(value.formula, 'XLOOKUP(B1,C:C,D:D)');
});

test('a date value round-trips as a Date, not a bare serial', () => {
  const when = new Date(Date.UTC(2021, 5, 15));
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = when;

  const value = rows(writeXlsx(wb))[0]?.cells[0]?.value;
  assert.ok(value instanceof Date);
  assert.equal(value.getTime(), when.getTime());
});

test('selects a worksheet by name', () => {
  const wb = new Workbook();
  wb.addWorksheet('first').getCell('A1').value = 'one';
  wb.addWorksheet('second').getCell('A1').value = 'two';

  assert.equal(rows(writeXlsx(wb), {sheet: 'second'})[0]?.cells[0]?.value, 'two');
});

test('selects a worksheet by 1-based position; default is the first', () => {
  const wb = new Workbook();
  wb.addWorksheet('first').getCell('A1').value = 'one';
  wb.addWorksheet('second').getCell('A1').value = 'two';
  const data = writeXlsx(wb);

  assert.equal(rows(data)[0]?.cells[0]?.value, 'one');
  assert.equal(rows(data, {sheet: 1})[0]?.cells[0]?.value, 'one');
  assert.equal(rows(data, {sheet: 2})[0]?.cells[0]?.value, 'two');
});

test('a missing sheet selector is an error, not silent emptiness', () => {
  const wb = new Workbook();
  wb.addWorksheet('only').getCell('A1').value = 'x';
  const data = writeXlsx(wb);

  assert.throws(() => rows(data, {sheet: 'nope'}), /no worksheet named/);
  assert.throws(() => rows(data, {sheet: 3}), /no worksheet at position 3/);
});

test('an inline string cell decodes through the streaming SAX path', () => {
  // writeXlsx pools strings into sharedStrings, so hand-build an inlineStr sheet to prove the
  // `<is><t>` path the streaming reader must also handle.
  const archive = foreignPackage({
    [SHEET1]: foreignSheet('<row r="1"><c r="A1" t="inlineStr"><is><t>hi there</t></is></c></row>'),
  });

  assert.equal(rows(archive)[0]?.cells[0]?.value, 'hi there');
});

test('a row past the last row is dropped, as the buffered reader drops it', () => {
  // The two readers have to agree on what an out-of-grid `<r>` means. The buffered one drops the
  // row (clamping would move its formatting onto 1048576); this one must not hand the consumer a
  // `number` that names no cell, and must retain nothing for the row it refuses.
  const archive = foreignPackage({
    [SHEET1]: foreignSheet(
      '<row r="1"><c r="A1" t="inlineStr"><is><t>in grid</t></is></c></row>' +
        `<row r="${MAX_ROW + 1}"><c r="A${MAX_ROW + 1}" t="inlineStr"><is><t>past it</t></is></c></row>` +
        '<row r="2"><c r="A2" t="inlineStr"><is><t>after</t></is></c></row>',
    ),
  });

  const streamed = rows(archive);
  assert.deepEqual(
    streamed.map((row) => row.number),
    [1, 2],
    'the out-of-grid row is not yielded, and the rows around it still are',
  );
  assert.deepEqual(
    streamed.map((row) => row.cells.map((cell) => cell.value)),
    [['in grid'], ['after']],
  );
});

test('a stylesheet at an unconventional path is still found, so a date stays a date', () => {
  // The resolution both readers run: through the relationship that names the part, with the
  // conventional path only as the fallback. A package is free to name any part anything, and a
  // stylesheet resolved conventional-path-first reads as no styles at all -- which changes a cell's
  // *type*, because the date test reads `numFmt` off the resolved style to tell 45000 from a date.
  // Nothing about the resulting workbook is malformed, so only a check like this one catches it.
  const archive = foreignPackage({
    'xl/_rels/workbook.xml.rels':
      '<Relationships>' +
      '<Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="x/styles" Target="theStyles.xml"/>' +
      '</Relationships>',
    // Number format 14 is the built-in short date, so `s="1"` is what makes the serial a date.
    'xl/theStyles.xml':
      '<styleSheet><cellXfs count="2"><xf numFmtId="0"/>' +
      '<xf numFmtId="14" applyNumberFormat="1"/></cellXfs></styleSheet>',
    // A stylesheet also sits at the conventional path, and it is not this workbook's: the two
    // resolution orders give different answers only when both parts exist, which is the case a
    // package that renames one of them and leaves the other behind actually presents.
    'xl/styles.xml': '<styleSheet><cellXfs count="0"/></styleSheet>',
    [SHEET1]: foreignSheet(
      '<row r="1"><c r="A1" s="1"><v>45000</v></c><c r="B1"><v>45000</v></c></row>',
    ),
  });

  const [row] = rows(archive);
  assert.ok(row?.cells[0]?.value instanceof Date, 'the styled cell reads as a date');
  assert.equal(row?.cells[1]?.value, 45_000, 'and the unstyled one beside it stays a number');
  // And the buffered reader answers the same package the same way: a streamed read must not decode
  // a cell differently from a buffered one.
  const model = readXlsx(archive).getWorksheet('S');
  assert.deepEqual(model?.getCell('A1').value, row?.cells[0]?.value);
});

test('streamed values agree with readXlsx cell-for-cell', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'text';
  sheet.getCell('B1').value = 3.14;
  sheet.getCell('A2').value = false;
  sheet.getCell('B2').value = {formula: 'A1&"!"', result: 'text!'};
  sheet.getCell('A3').value = new Date(Date.UTC(2000, 0, 1));
  const data = writeXlsx(wb);

  const model = readXlsx(data).getWorksheet('S');
  for (const row of rows(data)) {
    for (const cell of row.cells) {
      assert.deepEqual(cell.value, model?.getCell(cell.address).value, cell.address);
    }
  }
});

test('readWorkbookStream yields every worksheet in order, by declared name', () => {
  const wb = new Workbook();
  wb.addWorksheet('Alpha').getCell('A1').value = 1;
  wb.addWorksheet('Beta').getCell('A1').value = 2;
  wb.addWorksheet('Gamma').getCell('A1').value = 3;

  const sheets = [...readWorkbookStream(writeXlsx(wb))];
  assert.deepEqual(
    sheets.map((sheet) => sheet.name),
    ['Alpha', 'Beta', 'Gamma'],
  );
  // Each sheet streams its own rows independently.
  assert.equal([...sheets[1]!.rows()][0]?.cells[0]?.value, 2);
});

test('a streamed sheet surfaces its hidden columns after the rows are consumed', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getColumn(2).hidden = true;
  sheet.getColumn(4).hidden = true;
  sheet.getCell('A1').value = 'a';

  const [streamed] = [...readWorkbookStream(writeXlsx(wb))];
  assert.ok(streamed !== undefined);
  for (const _row of streamed.rows()) void _row;
  assert.deepEqual(streamed.hiddenColumns, [2, 4]);
});

test('hidden columns and merges resolve lazily even without iterating rows', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getColumn(3).hidden = true;
  sheet.getCell('A1').value = 'm';
  sheet.mergeCells('A1:B2');

  const [streamed] = [...readWorkbookStream(writeXlsx(wb))];
  assert.ok(streamed !== undefined);
  // Reading a summary before touching rows() drives a scan of its own.
  assert.deepEqual(streamed.hiddenColumns, [3]);
  assert.deepEqual(streamed.merges, ['A1:B2']);
});

test('a streamed row carries its hidden flag; visible rows report false', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'top';
  sheet.getCell('A2').value = 'mid';
  sheet.getRow(2).hidden = true;

  const [streamed] = [...readWorkbookStream(writeXlsx(wb))];
  assert.deepEqual(
    [...streamed!.rows()].map((row) => [row.number, row.hidden]),
    [
      [1, false],
      [2, true],
    ],
  );
});

test('readWorkbookStream agrees with readXlsx across every sheet and cell', () => {
  const wb = new Workbook();
  const a = wb.addWorksheet('A');
  a.getCell('A1').value = 'text';
  a.getCell('B2').value = {formula: 'A1&"!"', result: 'text!'};
  const b = wb.addWorksheet('B');
  b.getCell('A1').value = 42;
  b.getCell('C3').value = new Date(Date.UTC(2005, 2, 4));
  const data = writeXlsx(wb);

  const model = readXlsx(data);
  for (const sheet of readWorkbookStream(data)) {
    const worksheet = model.getWorksheet(sheet.name);
    for (const row of sheet.rows()) {
      for (const cell of row.cells) {
        assert.deepEqual(
          cell.value,
          worksheet?.getCell(cell.address).value,
          `${sheet.name}!${cell.address}`,
        );
      }
    }
  }
});

test('rows() is re-iterable: a second pass yields the same rows and summaries', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getColumn(2).hidden = true;
  sheet.getCell('A1').value = 'x';
  sheet.mergeCells('A1:B1');

  const [streamed] = [...readWorkbookStream(writeXlsx(wb))];
  assert.ok(streamed !== undefined);
  const first = [...streamed.rows()].map((row) => row.number);
  const second = [...streamed.rows()].map((row) => row.number);
  assert.deepEqual(first, second);
  assert.deepEqual(streamed.hiddenColumns, [2]);
  assert.deepEqual(streamed.merges, ['A1:B1']);
});

test('a streamed cell surfaces its resolved style facets for a read→write copy', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  const cell = sheet.getCell('A1');
  cell.value = 'hello';
  cell.font = {bold: true, color: {argb: 'FFFF0000'}};
  cell.numFmt = '0.00%';
  cell.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF00FF00'}};

  const [streamed] = [...readWorkbookStream(writeXlsx(wb))];
  const style = [...streamed!.rows()][0]?.cells[0]?.style;
  assert.ok(style, 'the styled cell carries a resolved style');
  assert.equal(style.numFmt, '0.00%');
  assert.equal(style.font?.bold, true);
  assert.equal(style.font?.color?.argb, 'FFFF0000');
  assert.equal(style.fill?.type, 'pattern');
});

test('an unstyled streamed cell carries no style facet', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'plain';

  const [streamed] = [...readWorkbookStream(writeXlsx(wb))];
  assert.equal([...streamed!.rows()][0]?.cells[0]?.style, undefined);
});

test('the generator is lazy: the first row is available before the rest are pulled', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  for (let r = 1; r <= 1000; r += 1) sheet.getCell(`A${r}`).value = r;

  const iterator = readSheetRows(writeXlsx(wb));
  const first = iterator.next();
  assert.ok(!first.done);
  assert.equal(first.value.number, 1);
  assert.equal(first.value.cells[0]?.value, 1);
  // Pulling one row must not have required draining the sheet; the next pull continues in order.
  const second = iterator.next();
  assert.ok(!second.done);
  assert.equal(second.value.number, 2);
});
