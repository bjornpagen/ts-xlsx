import assert from 'node:assert/strict';
import {test} from 'node:test';

import {XlsxError} from '../errors.ts';
import {Table, type TableOptions} from './table.ts';

function table(overrides: Partial<TableOptions> = {}): Table {
  return new Table({
    name: 'T',
    ref: 'A3',
    columns: [{name: 'H1'}, {name: 'H2'}],
    rowCount: 2,
    ...overrides,
  });
}

test('duplicate column names are disambiguated with a numeric suffix', () => {
  const t = new Table({
    name: 'T',
    ref: 'A1',
    columns: [{name: 'Dup'}, {name: 'Dup'}, {name: 'Dup'}],
    rowCount: 1,
  });
  assert.deepStrictEqual(
    t.columns.map((c) => c.name),
    ['Dup', 'Dup2', 'Dup3'],
    'the first name is kept; later clashes gain the smallest resolving suffix',
  );
});

test('duplicate column names are disambiguated case-insensitively', () => {
  const t = new Table({
    name: 'T',
    ref: 'A1',
    columns: [{name: 'Name'}, {name: 'name'}],
    rowCount: 1,
  });
  const [first, second] = t.columns.map((c) => c.name);
  assert.strictEqual(first, 'Name');
  assert.notStrictEqual(
    second?.toLowerCase(),
    'name',
    'a case-insensitive clash is still resolved',
  );
});

test('distinct column names are accepted', () => {
  assert.doesNotThrow(() => table());
});

test('inserting a row above a table shifts its whole range down', () => {
  const t = table(); // A3:B5 (header + 2 data rows)
  assert.strictEqual(t.range, 'A3:B5');
  const alive = t.shiftRows({axis: 'row', start: 1, count: 0, delta: 1}); // insert one row at the top
  assert.strictEqual(alive, true);
  assert.strictEqual(t.range, 'A4:B6');
});

test('inserting rows inside a table grows its data rows', () => {
  const t = table(); // A3:B5
  t.shiftRows({axis: 'row', start: 4, count: 0, delta: 2}); // two rows inserted within the data body (row 4)
  assert.strictEqual(t.range, 'A3:B7', 'the table absorbs the inserted rows');
});

test('a table left entirely above the splice is untouched', () => {
  const t = table(); // A3:B5
  t.shiftRows({axis: 'row', start: 10, count: 0, delta: 5}); // insert well below the table
  assert.strictEqual(t.range, 'A3:B5');
});

test('deleting every row of a table reports it as removed', () => {
  const t = table(); // rows 3..5
  const alive = t.shiftRows({axis: 'row', start: 3, count: 3, delta: -3}); // delete the whole span
  assert.strictEqual(alive, false);
});

test('a column splice to the left shifts the table anchor', () => {
  const t = table(); // anchored at column A (A3:B5)
  t.shiftColumns({axis: 'col', start: 1, count: 0, delta: 2}); // insert two columns before it
  assert.strictEqual(t.range, 'C3:D5');
});

test('the sentinel style name "None" is normalised to an absent name', () => {
  const t = table({style: {name: 'None', showRowStripes: true}});
  assert.strictEqual(
    t.style?.name,
    undefined,
    '"None" means unstyled, not a literal style reference',
  );
  assert.strictEqual(t.style?.showRowStripes, true, 'flags set alongside the theme survive');
});

test('a real style name is preserved', () => {
  const t = table({style: {name: 'TableStyleMedium2'}});
  assert.strictEqual(t.style?.name, 'TableStyleMedium2');
});

test('rowCount reports the data-row count', () => {
  assert.strictEqual(table().rowCount, 2);
});

test('addRow grows the data-row count and the range', () => {
  const t = table(); // A3:B5, 2 data rows
  t.addRow();
  assert.strictEqual(t.rowCount, 3);
  assert.strictEqual(t.range, 'A3:B6');
});

test('addRow with values on a table not attached to a worksheet throws', () => {
  assert.throws(() => table().addRow(['a', 1]), /not attached to a worksheet/);
});

test('addRow rejects more values than the table has columns', () => {
  assert.throws(() => table().addRow(['a', 1, 'x']), /has 2 columns/);
});

test('addRow on a detached table with a totals row throws: relocation needs the grid', () => {
  assert.throws(
    () => table({totalsRow: true}).addRow(),
    /not attached to a worksheet.*relocate its totals row/,
  );
});

test('a table name that is not an Excel identifier is refused, natively', () => {
  // A name is a single scalar: out of range is a `RangeError`, unparseable a `SyntaxError`, the same
  // way a comment id that is not a GUID is. The composite claims about a table (its columns spanning
  // its range, a duplicate column name) are what stay in the library's own taxonomy. Asserted as
  // *not* an XlsxError so a later re-wrap reddens the suite instead of changing what callers catch.
  assert.throws(() => table({name: ''}), {name: 'RangeError', message: /1 and 255/});
  assert.throws(() => table({name: 'x'.repeat(256)}), {name: 'RangeError'});
  assert.throws(() => table({name: '1st quarter'}), {
    name: 'SyntaxError',
    message: /valid Excel identifier/,
  });
  for (const name of ['', '1st quarter']) {
    assert.throws(
      () => table({name}),
      (error: unknown) => !(error instanceof XlsxError),
    );
  }
});

test('a table name that reads as a reference is refused, natively', () => {
  // Each of these made Excel 16.0 offer to repair a package carrying the table; the three beside them
  // opened clean.
  for (const name of ['T1', 'c1', 'R', 'RC', 'R1C1', 'R1X', 'TRUE']) {
    assert.throws(() => table({name}), {name: 'SyntaxError', message: /reads as a cell reference/});
  }
  for (const name of ['CC', 'XFE1', 'R1.5']) assert.equal(table({name}).name, name);
});

test('a table whose derived corner leaves the grid is refused at construction', () => {
  // The anchor was validated and the far corner, derived from it, was not. `XFC1` plus three columns
  // reached column XFE, and five million data rows produced a `<table ref>` naming rows that cannot
  // exist -- which the writer emitted, and which `range` then threw on when anyone read it back.
  assert.throws(() => table({ref: 'XFC1', columns: [{name: 'a'}, {name: 'b'}, {name: 'c'}]}), {
    name: 'RangeError',
    message: /table "T" spans 3 columns from XFC, past the last column \(XFD\)/,
  });
  assert.throws(() => table({ref: 'A1', columns: [{name: 'a'}], rowCount: 5_000_000}), {
    name: 'RangeError',
    message: /table "T" spans 5000001 rows from 1, past the last row \(1048576\)/,
  });
});

test('a table reaching the last column or row exactly is accepted', () => {
  assert.strictEqual(table({ref: 'XFC1', columns: [{name: 'a'}, {name: 'b'}]}).range, 'XFC1:XFD3');
  assert.strictEqual(
    table({ref: 'A1048574', columns: [{name: 'a'}], rowCount: 2}).range,
    'A1048574:A1048576',
  );
});

test('a column splice that would push a table past the last column drops it', () => {
  // Clamping the anchor is not the same as bounding the table: an anchor clamped onto XFD still puts
  // a two-column table's right edge at XFE, where `range`, `autoFilterRef` and `region` all throw.
  const t = table({ref: 'XFC1', columns: [{name: 'a'}, {name: 'b'}]}); // XFC..XFD
  assert.strictEqual(
    t.shiftColumns({axis: 'col', start: 1, count: 0, delta: 1}),
    false,
    'no room left for its columns',
  );
});
