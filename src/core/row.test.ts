import assert from 'node:assert/strict';
import test from 'node:test';

import {columnToNumber, MAX_COLUMN, MAX_ROW, numberToColumn} from './address.ts';
import {Worksheet} from './worksheet.ts';

function thrownMessage(run: () => unknown): string {
  try {
    run();
  } catch (error) {
    return error instanceof Error ? `${error.name}: ${error.message}` : String(error);
  }
  return 'no throw';
}

test("a row's cell refuses a column off the grid in the words the column decoders use", () => {
  const sheet = new Worksheet('S', 1);
  const row = sheet.getRow(1);
  assert.equal(
    thrownMessage(() => row.getCell(MAX_COLUMN + 1)),
    thrownMessage(() => numberToColumn(MAX_COLUMN + 1)),
  );
  assert.equal(
    thrownMessage(() => row.getCell('XFE')),
    thrownMessage(() => columnToNumber('XFE')),
  );
  assert.deepEqual([...sheet.rows()], [], 'and nothing was materialised');
});

test('a row handle is a live view, not a copy: two handles on one number agree', () => {
  const sheet = new Worksheet('S', 1);
  const first = sheet.getRow(3);
  const second = sheet.getRow(3);

  first.height = 22;
  assert.equal(second.height, 22, 'the second handle sees a write made through the first');

  second.getCell(1).value = 'a';
  assert.equal(first.cells.length, 1, 'and sees a cell materialised through the other');
  assert.equal(first.getCell('A').value, 'a');
});

test('reading a row creates nothing: no record, no cells, no used range', () => {
  const sheet = new Worksheet('S', 1);
  const row = sheet.getRow(500);

  assert.equal(row.properties, undefined, 'no format record fabricated');
  assert.deepEqual(row.cells, [], 'no cells fabricated');
  assert.deepEqual(row.values, [], 'no values fabricated');
  assert.equal(sheet.rowCount, 0, 'merely asking about row 500 does not extend the used range');
  assert.deepEqual([...sheet.rows()], [], 'and the row does not appear in the sheet iteration');
});

test('writing formatting through a row creates the record, and clearing it removes the key', () => {
  const sheet = new Worksheet('S', 1);
  const row = sheet.getRow(2);

  row.hidden = true;
  row.height = 30;
  assert.deepEqual(
    row.properties,
    {hidden: true, height: 30},
    'each write created exactly the key it named',
  );
  assert.equal(sheet.rowCount, 2, 'a formatted row is in the used range');

  row.hidden = undefined;
  assert.deepEqual(
    row.properties,
    {height: 30},
    'clearing deletes the key rather than storing an explicit undefined',
  );
  assert.equal(sheet.rowCount, 2, 'and the row is still formatted, so still in the used range');
});

test("clearing a row's last property takes the record with it, and the used range with that", () => {
  const sheet = new Worksheet('S', 1);
  sheet.getRow(500).height = 20;
  assert.equal(sheet.rowCount, 500);

  sheet.getRow(500).height = undefined;
  assert.equal(
    sheet.getRow(500).properties,
    undefined,
    'an emptied record is not a record: `properties` promises a read that never fabricates',
  );
  assert.equal(sheet.rowCount, 0, 'and nothing is left declaring row 500, so the extent recedes');
});

test('clearing a property on a row that has no record creates no record', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getRow(4).height = undefined;
  assert.equal(sheet.getRow(4).properties, undefined);
  assert.equal(sheet.rowCount, 0);
});

test('every RowProperties facet round-trips through the handle', () => {
  const sheet = new Worksheet('S', 1);
  const row = sheet.getRow(1);
  const fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF0000'}} as const;

  row.height = 30;
  row.hidden = true;
  row.outlineLevel = 2;
  row.collapsed = true;
  row.fill = fill;

  assert.equal(row.height, 30);
  assert.equal(row.hidden, true);
  assert.equal(row.outlineLevel, 2);
  assert.equal(row.collapsed, true);
  assert.deepEqual(row.fill, fill);
  assert.deepEqual(sheet.getRow(1).properties, {
    height: 30,
    hidden: true,
    outlineLevel: 2,
    collapsed: true,
    fill,
  });
});

test('row.getCell takes an index or column letters, and resolves through merges', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('A1').value = 'master';
  sheet.mergeCells('A1:B2');

  const row = sheet.getRow(1);
  assert.equal(row.getCell(1).address, 'A1');
  assert.equal(row.getCell('B').address, 'A1', 'a covered column resolves to the region master');
  assert.equal(sheet.getRow(2).getCell('B').address, 'A1', 'from any row of the region');
});

test('row.cells is the materialised cells in ascending column order, sparsely', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('C1').value = 'c';
  sheet.getCell('A1').value = 'a';

  assert.deepEqual(
    sheet.getRow(1).cells.map((cell) => cell.address),
    ['A1', 'C1'],
    'ascending by column regardless of the order they were written',
  );
});

test('row.values reads by position with column A at index 0, leaving gaps as holes', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('A1').value = 'a';
  sheet.getCell('C1').value = 'c';

  const values = sheet.getRow(1).values;
  assert.equal(values.length, 3);
  assert.equal(values[0], 'a');
  assert.equal(1 in values, false, 'an unwritten column is a hole, not an explicit undefined');
  assert.equal(values[2], 'c');
});

test('a cell holding null reads back as null, distinct from a column that was never written', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('A1').value = null;
  const values = sheet.getRow(1).values;
  assert.equal(0 in values, true, 'the materialised cell is present');
  assert.equal(values[0], null);
});

test('assigning row.values places what it names and leaves every other column untouched', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getRow(1).values = ['a', 'b', 'c'];
  assert.deepEqual(sheet.getRow(1).values, ['a', 'b', 'c']);

  sheet.getRow(1).values = ['A'];
  assert.deepEqual(
    sheet.getRow(1).values,
    ['A', 'b', 'c'],
    'a shorter array does not clear the tail: addRow rules, not replacement',
  );

  sheet.getRow(1).values = [undefined, 'B'];
  assert.deepEqual(sheet.getRow(1).values, ['A', 'B', 'c'], 'an explicit undefined skips a column');
});

test('a row handle is positional: a splice moves content, not the handle', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('A1').value = 'first';
  sheet.getRow(1).height = 40;

  const row = sheet.getRow(1);
  sheet.insertRow(1, ['inserted']);

  assert.equal(row.number, 1, 'the handle still names row 1');
  assert.equal(
    row.getCell(1).value,
    'inserted',
    "and now sees row 1's new content: a handle fixes its position, exactly as Cell does",
  );
  assert.equal(sheet.getRow(2).height, 40, 'the original row and its formatting shifted down');
});

test('a row is out of bounds below 1, and rejects a non-integer', () => {
  const sheet = new Worksheet('S', 1);
  assert.throws(() => sheet.getRow(0), RangeError);
  assert.throws(() => sheet.getRow(-1), RangeError);
  assert.throws(() => sheet.getRow(1.5), RangeError);
});

test('a row past 1048576 is out of bounds', () => {
  const sheet = new Worksheet('S', 1);
  assert.equal(sheet.getRow(MAX_ROW).number, MAX_ROW, 'the last row in the grid stands');
  assert.throws(() => sheet.getRow(MAX_ROW + 1), {
    name: 'RangeError',
    message: `row ${MAX_ROW + 1} is out of bounds: Excel supports 1..${MAX_ROW}`,
  });
});

test('sheet.rows() yields handles over the union of celled and formatted rows, ascending', () => {
  const sheet = new Worksheet('S', 1);
  sheet.getCell('A3').value = 'three';
  sheet.getRow(1).hidden = true;

  const rows = [...sheet.rows()];
  assert.deepEqual(
    rows.map((row) => row.number),
    [1, 3],
  );
  assert.equal(rows[0]?.properties?.hidden, true, 'the formatting-only row carries its record');
  assert.deepEqual(rows[0]?.cells, [], 'and no cells');
  assert.equal(rows[1]?.properties, undefined, 'the cells-only row carries no record');
  assert.equal(rows[1]?.cells[0]?.value, 'three');
});
