import assert from 'node:assert/strict';
import {test} from 'node:test';

import type {TotalsRowFunction} from '../../core/table.ts';
import {Workbook} from '../../core/workbook.ts';
import {partIn, partsWritten, roundtrip} from './package.test-support.ts';
import {parseTable} from './tables.ts';
import {writeXlsx} from './write.ts';

// Author a workbook whose single sheet carries one table, round-trip it, and hand back the
// reconstructed table for assertions.
function roundtripTable(options: {
  name: string;
  displayName?: string;
  ref: string;
  columns: {name: string; totalsRowLabel?: string; totalsRowFunction?: TotalsRowFunction}[];
  rowCount: number;
  headerRow?: boolean;
  totalsRow?: boolean;
  totalsRowShown?: boolean;
  autoFilter?: boolean;
  style?: {
    name?: string;
    showFirstColumn?: boolean;
    showLastColumn?: boolean;
    showRowStripes?: boolean;
    showColumnStripes?: boolean;
  };
}) {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable(options);
  const back = roundtrip(wb);
  const sheet = back.getWorksheet('S');
  assert.ok(sheet !== undefined);
  return sheet.tables;
}

test('a table read back from a written package exposes its name, columns, and ref', () => {
  const [table, ...rest] = roundtripTable({
    name: 'Inventory',
    ref: 'A1',
    columns: [{name: 'Item'}, {name: 'Qty'}],
    rowCount: 3,
  });
  assert.equal(rest.length, 0, 'exactly one table is reconstructed');
  assert.ok(table !== undefined);
  assert.equal(table.name, 'Inventory');
  assert.deepEqual(
    table.columns.map((c) => c.name),
    ['Item', 'Qty'],
  );
  assert.equal(table.range, 'A1:B4', 'header + 3 data rows spans four rows');
  assert.equal(table.options.ref, 'A1', 'the anchor reconstructs to the top-left cell');
});

const tablePart = (ref: string, columns: string, style = ''): string =>
  '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" ' +
  `displayName="T" ref="${ref}" totalsRowCount="1"><tableColumns>${columns}</tableColumns>${style}</table>`;

// A `<tableColumn>` with no name was skipped with nothing checking the count, so the columns after it
// shifted one place left: removing the third column's name from an A1:C3 table read back columns a
// and b, the range A1:B3, and b owning the third column's `SUM(T[c])`.
test('a nameless column drops the table rather than shifting its columns and totals formula', () => {
  const custom = (name: string) =>
    `<tableColumn id="3"${name} totalsRowFunction="custom"><totalsRowFormula>SUM(T[c])</totalsRowFormula></tableColumn>`;
  const columns = '<tableColumn id="1" name="a"/><tableColumn id="2" name="b"/>';

  const whole = parseTable(tablePart('A1:C3', columns + custom(' name="c"')));
  assert.deepEqual(
    whole?.columns.map((column) => [column.name, column.totalsRowFormula]),
    [
      ['a', undefined],
      ['b', undefined],
      ['c', 'SUM(T[c])'],
    ],
    'control: with the name present the formula sits on its own column',
  );
  assert.equal(parseTable(tablePart('A1:C3', columns + custom(''))), undefined);
});

test('a table whose column count differs from its ref width is dropped whole', () => {
  const two = '<tableColumn id="1" name="a"/><tableColumn id="2" name="b"/>';
  assert.ok(parseTable(tablePart('A1:B3', two)), 'control: two columns across a two-wide ref');
  assert.equal(
    parseTable(tablePart('A1:C3', two)),
    undefined,
    'fewer columns than the ref is wide',
  );
  assert.equal(parseTable(tablePart('A1:A3', two)), undefined, 'more columns than the ref is wide');
});

// `boolPresent` read `"yes"` as present-and-true, so the writer turned it into `"1"`.
test('an unrecognised boolean token on a table or its style info is dropped, not read as true', () => {
  const two = '<tableColumn id="1" name="a"/><tableColumn id="2" name="b"/>';
  const style =
    '<tableStyleInfo name="TableStyleLight1" showRowStripes="yes" showFirstColumn="0"/>';
  const part = tablePart('A1:B3', two, style).replace('totalsRowCount="1"', 'totalsRowShown="yes"');
  const options = parseTable(part);
  assert.deepEqual(options?.style, {name: 'TableStyleLight1', showFirstColumn: false});
  assert.equal(options?.totalsRowShown, undefined);
});

test('a loaded table exposes its data-row count, not an empty rows array', () => {
  const [table] = roundtripTable({
    name: 'T',
    ref: 'A1',
    columns: [{name: 'Name'}],
    rowCount: 5,
  });
  assert.ok(table !== undefined);
  assert.equal(table.options.rowCount, 5, 'the data-row count survives the round-trip');
});

test('an empty-body table (header only) round-trips with a zero data-row count', () => {
  const [table] = roundtripTable({
    name: 'Empty',
    ref: 'A1',
    columns: [{name: 'C1'}, {name: 'C2'}],
    rowCount: 0,
  });
  assert.ok(table !== undefined);
  assert.equal(table.options.rowCount, 0);
  assert.equal(table.range, 'A1:B1', 'a header-only table occupies a single row');
});

test('a headerless table round-trips with headerRow false and no autofilter', () => {
  const [table] = roundtripTable({
    name: 'Bare',
    ref: 'A1',
    columns: [{name: 'C1'}],
    rowCount: 2,
    headerRow: false,
  });
  assert.ok(table !== undefined);
  assert.equal(table.options.headerRow, false);
  assert.equal(table.options.rowCount, 2, 'both rows are data rows when there is no header');
  assert.equal(table.autoFilterRef, undefined, 'a headerless table anchors no autofilter');
});

test('a totals-row table round-trips its totals flag and per-column totals behaviour', () => {
  const [table] = roundtripTable({
    name: 'Totalled',
    ref: 'A1',
    columns: [
      {name: 'Label', totalsRowLabel: 'Total'},
      {name: 'Amount', totalsRowFunction: 'sum'},
    ],
    rowCount: 2,
    totalsRow: true,
  });
  assert.ok(table !== undefined);
  assert.equal(table.options.totalsRow, true);
  assert.equal(table.options.rowCount, 2, 'the totals row is not counted as a data row');
  assert.equal(table.range, 'A1:B4', 'header + 2 data + totals spans four rows');
  assert.equal(table.columns[0]?.totalsRowLabel, 'Total');
  assert.equal(table.columns[1]?.totalsRowFunction, 'sum');
});

test('an unrecognised totalsRowFunction is dropped rather than trusted in verbatim', () => {
  const xml =
    '<?xml version="1.0"?>' +
    '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" ' +
    'name="T" displayName="T" ref="A1:A3" totalsRowCount="1">' +
    '<tableColumns count="1"><tableColumn id="1" name="A" totalsRowFunction="avg"/></tableColumns>' +
    '</table>';
  const table = parseTable(xml);
  assert.ok(table !== undefined);
  assert.equal(table.columns[0]?.totalsRowFunction, undefined);
});

test('a table with no totalsRowShown flag stays without one across a round-trip', () => {
  const [table] = roundtripTable({name: 'T', ref: 'A1', columns: [{name: 'A'}], rowCount: 2});
  assert.ok(table !== undefined);
  assert.equal(
    table.options.totalsRowShown,
    undefined,
    'an absent flag must not be fabricated on read-back',
  );
  assert.equal(
    'totalsRowShown' in table.options,
    false,
    'the key is omitted, not set to undefined',
  );
});

test('an explicit totalsRowShown flag survives a round-trip in both states', () => {
  const [off] = roundtripTable({
    name: 'T',
    ref: 'A1',
    columns: [{name: 'A'}],
    rowCount: 2,
    totalsRowShown: false,
  });
  assert.equal(off?.options.totalsRowShown, false);

  const [on] = roundtripTable({
    name: 'T',
    ref: 'A1',
    columns: [{name: 'A'}],
    rowCount: 2,
    totalsRowShown: true,
  });
  assert.equal(on?.options.totalsRowShown, true);
});

test('a custom table style name and banding flags survive a round-trip verbatim', () => {
  const [table] = roundtripTable({
    name: 'Styled',
    ref: 'A1',
    columns: [{name: 'A'}],
    rowCount: 2,
    style: {
      name: 'TableStyleLight9',
      showFirstColumn: true,
      showLastColumn: false,
      showRowStripes: false,
      showColumnStripes: true,
    },
  });
  assert.ok(table !== undefined);
  assert.deepEqual(table.options.style, {
    name: 'TableStyleLight9',
    showFirstColumn: true,
    showLastColumn: false,
    showRowStripes: false,
    showColumnStripes: true,
  });
});

test('a style that omits its name round-trips still nameless, not defaulted', () => {
  const [table] = roundtripTable({
    name: 'NoName',
    ref: 'A1',
    columns: [{name: 'A'}],
    rowCount: 1,
    style: {showRowStripes: true},
  });
  assert.ok(table !== undefined);
  assert.equal(
    'name' in (table.options.style ?? {}),
    false,
    'an absent style name is not fabricated',
  );
  assert.equal(table.options.style?.showRowStripes, true);
});

test("a table authored without a style reads back with Excel's default style", () => {
  const [table] = roundtripTable({name: 'Plain', ref: 'A1', columns: [{name: 'A'}], rowCount: 1});
  assert.ok(table !== undefined);
  assert.equal(table.options.style?.name, 'TableStyleMedium2');
  assert.equal(table.options.style?.showRowStripes, true);
});

test('a distinct display name round-trips independently of the internal name', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({
    name: 'MyTable',
    displayName: 'My Display Name',
    ref: 'A1',
    columns: [{name: 'C'}],
    rowCount: 1,
  });
  const table = roundtrip(wb).getWorksheet('S')?.tables[0];
  assert.ok(table !== undefined);
  assert.equal(table.name, 'MyTable', 'the internal identifier is unaffected');
  assert.equal(table.displayName, 'My Display Name', 'the display label survives the round-trip');
});

test('a display name defaults to the table name when unset', () => {
  const [table] = roundtripTable({name: 'Plain', ref: 'A1', columns: [{name: 'C'}], rowCount: 1});
  assert.ok(table !== undefined);
  assert.equal(table.displayName, 'Plain');
});

test('a header table read without an autoFilter does not gain one on round-trip', () => {
  const [table] = roundtripTable({
    name: 'Unfiltered',
    ref: 'A1',
    columns: [{name: 'C'}],
    rowCount: 2,
    autoFilter: false,
  });
  assert.ok(table !== undefined);
  assert.equal(table.headerRow, true, 'it keeps its header row');
  assert.equal(table.autoFilterRef, undefined, 'but exposes no autoFilter range');
  assert.equal(table.options.autoFilter, false, 'and the flag survives the round-trip');
});

test('a header table gains an autoFilter by default', () => {
  const [table] = roundtripTable({
    name: 'Filtered',
    ref: 'A1',
    columns: [{name: 'C'}],
    rowCount: 2,
  });
  assert.ok(table !== undefined);
  assert.equal(table.autoFilterRef, 'A1:A3', 'the default autoFilter spans header + data rows');
  assert.equal(table.options.autoFilter, true);
});

test('several tables on one sheet all read back in definition order', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.addTable({name: 'First', ref: 'A1', columns: [{name: 'A'}], rowCount: 1});
  sheet.addTable({name: 'Second', ref: 'D1', columns: [{name: 'B'}, {name: 'C'}], rowCount: 2});

  const back = roundtrip(wb);
  const tables = back.getWorksheet('S')?.tables ?? [];
  assert.deepEqual(
    tables.map((t) => t.name),
    ['First', 'Second'],
  );
  assert.equal(tables[1]?.range, 'D1:E3');
});

test('tables on distinct sheets each reconstruct on their own sheet', () => {
  const wb = new Workbook();
  wb.addWorksheet('One').addTable({name: 'TA', ref: 'A1', columns: [{name: 'X'}], rowCount: 1});
  wb.addWorksheet('Two').addTable({name: 'TB', ref: 'A1', columns: [{name: 'Y'}], rowCount: 1});

  const back = roundtrip(wb);
  assert.deepEqual(
    back.getWorksheet('One')?.tables.map((t) => t.name),
    ['TA'],
  );
  assert.deepEqual(
    back.getWorksheet('Two')?.tables.map((t) => t.name),
    ['TB'],
  );
});

test('declaring a table fills its header row cells with the column names', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.addTable({name: 'T', ref: 'C3', columns: [{name: 'Alpha'}, {name: 'Beta'}], rowCount: 1});
  assert.equal(sheet.getCell('C3').value, 'Alpha', 'headers land at the table anchor, not row 1');
  assert.equal(sheet.getCell('D3').value, 'Beta');
});

test('a headerless table claims no header row', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.addTable({
    name: 'T',
    ref: 'A1',
    columns: [{name: 'Alpha'}],
    rowCount: 1,
    headerRow: false,
  });
  assert.equal(sheet.getCell('A1').value, null, 'row 1 belongs to the data, not a header');
});

test('a header cell that already holds a value is not overwritten by a table declaration', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'Existing';
  sheet.addTable({name: 'T', ref: 'A1', columns: [{name: 'Alpha'}, {name: 'Beta'}], rowCount: 1});
  assert.equal(sheet.getCell('A1').value, 'Existing', 'authored content wins over the column name');
  assert.equal(sheet.getCell('B1').value, 'Beta', 'the empty header cell is still filled');
});

// Reading re-registers every table through addTable *after* the sheet's cells are loaded, so the
// header-filling step runs against a populated grid. It must never write over what the file said:
// these two cases are the ones where the loaded cell carries more than the column name does.
test('reading does not flatten a rich-text header cell to its plain column name', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.addTable({name: 'T', ref: 'A1', columns: [{name: 'Alpha'}], rowCount: 1});
  sheet.getCell('A1').value = {richText: [{text: 'Al', font: {bold: true}}, {text: 'pha'}]};

  const value = roundtrip(wb).getWorksheet('S')?.getCell('A1').value;
  assert.ok(value !== null && typeof value === 'object' && 'richText' in value);
  assert.equal(value.richText.length, 2, 'the runs survive the read');
  assert.equal(value.richText[0]?.font?.bold, true);
});

test('reading preserves a header cell whose text drifted from the declared column name', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.addTable({name: 'T', ref: 'A1', columns: [{name: 'Alpha'}], rowCount: 1});
  sheet.getCell('A1').value = 'Drifted';

  const back = roundtrip(wb);
  assert.equal(
    back.getWorksheet('S')?.getCell('A1').value,
    'Drifted',
    'the file is authoritative on read: a re-declaration must not repair it',
  );
});

test('a table survives a second read → write → read round-trip unchanged', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({
    name: 'Persist',
    ref: 'B2',
    columns: [{name: 'One'}, {name: 'Two'}, {name: 'Three'}],
    rowCount: 4,
  });
  const once = roundtrip(wb);
  const twice = roundtrip(once);
  const table = twice.getWorksheet('S')?.tables[0];
  assert.ok(table !== undefined);
  assert.equal(table.name, 'Persist');
  assert.equal(table.range, 'B2:D6', 'the range is stable across two round-trips');
  assert.deepEqual(
    table.columns.map((c) => c.name),
    ['One', 'Two', 'Three'],
  );
});

test('a table part spelling its booleans "false" reads them off, as "0" does', () => {
  const part = (value: string): string =>
    '<?xml version="1.0"?>' +
    '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" ' +
    `name="T" displayName="T" ref="A1:A3" totalsRowCount="1" totalsRowShown="${value}">` +
    '<tableColumns count="1"><tableColumn id="1" name="A"/></tableColumns>' +
    `<tableStyleInfo name="TableStyleMedium2" showFirstColumn="${value}" ` +
    `showLastColumn="${value}" showRowStripes="${value}" showColumnStripes="${value}"/>` +
    '</table>';
  for (const value of ['0', 'false']) {
    const table = parseTable(part(value));
    assert.ok(table !== undefined);
    assert.equal(table.totalsRowShown, false, `totalsRowShown="${value}"`);
    assert.deepEqual(
      {...table.style, name: undefined},
      {
        name: undefined,
        showFirstColumn: false,
        showLastColumn: false,
        showRowStripes: false,
        showColumnStripes: false,
      },
      `every tableStyleInfo flag spelled "${value}" reads off`,
    );
  }
});

test('an empty-body table refs the full header row and writes a table part', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({
    name: 'Empty',
    ref: 'A1',
    columns: [{name: 'Alpha'}, {name: 'Beta'}],
    rowCount: 0,
  });
  const parts = partsWritten(wb);
  const table = partIn(parts, 'xl/tables/table1.xml');
  assert.match(table, /ref="A1:B1"/);
  assert.match(table, /<tableColumns count="2">/);
  assert.match(table, /<autoFilter ref="A1:B1"\/>/);
  assert.match(partIn(parts, '[Content_Types].xml'), /\/xl\/tables\/table1\.xml/);
  assert.match(
    partIn(parts, 'xl/worksheets/_rels/sheet1.xml.rels'),
    /Target="\.\.\/tables\/table1\.xml"/,
  );
  assert.match(
    partIn(parts, 'xl/worksheets/sheet1.xml'),
    /<tableParts count="1"><tablePart r:id="rId1"\/><\/tableParts>/,
  );
});

test('a data row extends the table ref by one row', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({
    name: 'T',
    ref: 'A1',
    columns: [{name: 'A'}, {name: 'B'}],
    rowCount: 1,
  });
  const table = partIn(partsWritten(wb), 'xl/tables/table1.xml');
  assert.match(table, /ref="A1:B2"/);
});

test('a headerless table sets headerRowCount="0" and emits no autoFilter', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({
    name: 'H',
    ref: 'A1',
    columns: [{name: 'A'}, {name: 'B'}],
    rowCount: 2,
    headerRow: false,
  });
  const table = partIn(partsWritten(wb), 'xl/tables/table1.xml');
  assert.match(table, /headerRowCount="0"/);
  assert.doesNotMatch(table, /<autoFilter/);
});

test('a totals-row column serialises its function and keeps every column', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({
    name: 'T',
    ref: 'A1',
    columns: [
      {name: 'Item', totalsRowLabel: 'Total'},
      {name: 'Amount', totalsRowFunction: 'sum'},
    ],
    rowCount: 2,
    totalsRow: true,
  });
  const table = partIn(partsWritten(wb), 'xl/tables/table1.xml');
  assert.match(table, /ref="A1:B4"/);
  assert.match(table, /totalsRowCount="1"/);
  assert.match(table, /<tableColumn id="1" name="Item" totalsRowLabel="Total"\/>/);
  assert.match(table, /<tableColumn id="2" name="Amount" totalsRowFunction="sum"\/>/);
});

test('a no-totals table omits totalsRowShown unless the flag is set explicitly', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({name: 'T', ref: 'A1', columns: [{name: 'A'}], rowCount: 1});
  const table = partIn(partsWritten(wb), 'xl/tables/table1.xml');
  assert.doesNotMatch(
    table,
    /totalsRowShown/,
    'an unset flag emits no attribute: Excel must not see a spurious one',
  );
});

test('an explicit totalsRowShown flag is written as "0" or "1"', () => {
  const off = new Workbook();
  off
    .addWorksheet('S')
    .addTable({name: 'T', ref: 'A1', columns: [{name: 'A'}], rowCount: 1, totalsRowShown: false});
  assert.match(partIn(partsWritten(off), 'xl/tables/table1.xml'), /totalsRowShown="0"/);

  const on = new Workbook();
  on.addWorksheet('S').addTable({
    name: 'T',
    ref: 'A1',
    columns: [{name: 'A'}],
    rowCount: 1,
    totalsRowShown: true,
  });
  assert.match(partIn(partsWritten(on), 'xl/tables/table1.xml'), /totalsRowShown="1"/);
});

test("a table with no explicit style is written with Excel's default TableStyleMedium2", () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({name: 'T', ref: 'A1', columns: [{name: 'A'}], rowCount: 1});
  const table = partIn(partsWritten(wb), 'xl/tables/table1.xml');
  assert.match(table, /<tableStyleInfo name="TableStyleMedium2"[^>]*showRowStripes="1"[^>]*\/>/);
});

test('an explicit table style is emitted verbatim, omitting the attributes it leaves unset', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({
    name: 'T',
    ref: 'A1',
    columns: [{name: 'A'}],
    rowCount: 1,
    style: {name: 'Assignment schedule', showRowStripes: false},
  });
  const table = partIn(partsWritten(wb), 'xl/tables/table1.xml');
  assert.match(
    table,
    /name="Assignment schedule"/,
    'a custom style name survives instead of being rewritten',
  );
  assert.match(
    table,
    /showRowStripes="0"/,
    'the source\'s stripe choice is preserved, not forced to "1"',
  );
  assert.doesNotMatch(table, /showFirstColumn/, 'an unset banding flag emits no attribute');
});

test('an illegal table name is rejected at definition time', () => {
  const s = new Workbook().addWorksheet('S');
  assert.throws(
    () => s.addTable({name: "Bob's Accounts", ref: 'A1', columns: [{name: 'A'}], rowCount: 1}),
    /identifier/,
  );
  assert.throws(
    () => s.addTable({name: '1Digit', ref: 'A1', columns: [{name: 'A'}], rowCount: 1}),
    /identifier/,
  );
  assert.throws(
    () => s.addTable({name: 'test-name', ref: 'A1', columns: [{name: 'A'}], rowCount: 1}),
    /identifier/,
  );
});

test('a valid identifier table name is written verbatim', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({
    name: 'Valid_Name',
    ref: 'A1',
    columns: [{name: 'A'}],
    rowCount: 1,
  });
  const table = partIn(partsWritten(wb), 'xl/tables/table1.xml');
  assert.match(table, /name="Valid_Name"/);
  assert.match(table, /displayName="Valid_Name"/);
});

test('tables are numbered globally across sheets with sheet-local rel ids', () => {
  const wb = new Workbook();
  wb.addWorksheet('One').addTable({name: 'Ta', ref: 'A1', columns: [{name: 'A'}], rowCount: 1});
  wb.addWorksheet('Two').addTable({name: 'Tb', ref: 'A1', columns: [{name: 'A'}], rowCount: 1});
  const parts = partsWritten(wb);
  assert.ok(parts['xl/tables/table1.xml'], 'first table part');
  assert.ok(parts['xl/tables/table2.xml'], 'second table part (globally numbered)');
  assert.match(
    partIn(parts, 'xl/worksheets/_rels/sheet2.xml.rels'),
    /Target="\.\.\/tables\/table2\.xml"/,
  );
});

test('a merge overlapping a table is rejected; a disjoint merge is written', () => {
  const overlap = new Workbook();
  const s1 = overlap.addWorksheet('S');
  s1.addTable({name: 'T', ref: 'A1', columns: [{name: 'A'}, {name: 'B'}], rowCount: 2});
  s1.mergeCells('A2:B2');
  assert.throws(() => writeXlsx(overlap), /overlaps table/);

  const disjoint = new Workbook();
  const s2 = disjoint.addWorksheet('S');
  s2.addTable({name: 'T', ref: 'A1', columns: [{name: 'A'}, {name: 'B'}], rowCount: 2});
  s2.mergeCells('D5:E5');
  const xml = partIn(partsWritten(disjoint), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<mergeCells count="1"><mergeCell ref="D5:E5"\/><\/mergeCells>/);
});
