import assert from 'node:assert/strict';
import {test} from 'node:test';

import {
  columnToNumber,
  decodeAddress,
  decodeCellRef,
  decodeRange,
  encodeAddress,
  MAX_COLUMN,
  MAX_ROW,
  nameReadsAsReference,
  numberToColumn,
  tryColumnToNumber,
  tryDecodeCellRef,
  tryDecodeRange,
} from './address.ts';
import {Workbook} from './workbook.ts';

test('numberToColumn covers the Excel range boundaries', () => {
  assert.equal(numberToColumn(1), 'A');
  assert.equal(numberToColumn(26), 'Z');
  assert.equal(numberToColumn(27), 'AA');
  assert.equal(numberToColumn(702), 'ZZ');
  assert.equal(numberToColumn(703), 'AAA');
  assert.equal(numberToColumn(MAX_COLUMN), 'XFD');
});

test('numberToColumn rejects out-of-bounds and non-integers', () => {
  assert.throws(() => numberToColumn(0), RangeError);
  assert.throws(() => numberToColumn(MAX_COLUMN + 1), RangeError);
  assert.throws(() => numberToColumn(1.5), RangeError);
});

test('columnToNumber is the inverse of numberToColumn across the whole range', () => {
  for (const n of [1, 26, 27, 52, 702, 703, 16383, MAX_COLUMN]) {
    assert.equal(columnToNumber(numberToColumn(n)), n);
  }
});

test('columnToNumber rejects invalid letters and overflow', () => {
  assert.throws(() => columnToNumber(''), RangeError);
  assert.throws(() => columnToNumber('a'), RangeError);
  assert.throws(() => columnToNumber('A1'), RangeError);
  assert.throws(() => columnToNumber('XFE'), RangeError); // 16385, just past XFD
  // The two failures keep their own wording: one is a typo, the other a column too far right, and a
  // caller fixes them differently.
  assert.throws(() => columnToNumber('A1'), /invalid column letters/);
  assert.throws(() => columnToNumber('XFE'), /out of bounds/);
});

test('tryColumnToNumber answers undefined where columnToNumber throws', () => {
  assert.equal(tryColumnToNumber('A'), 1);
  assert.equal(tryColumnToNumber('XFD'), MAX_COLUMN);
  for (const letters of ['', 'a', 'A1', 'XFE', 'ZZZ', 'AAAA', ' A']) {
    assert.equal(tryColumnToNumber(letters), undefined, letters);
    assert.throws(() => columnToNumber(letters), RangeError, letters);
  }
});

test('decodeAddress reads a plain and an absolute cell identically', () => {
  assert.deepEqual(decodeAddress('B2'), {address: 'B2', col: 2, row: 2});
  assert.deepEqual(decodeAddress('$B$2'), {address: 'B2', col: 2, row: 2});
});

test('decodeAddress leaves the omitted axis undefined, not a sentinel', () => {
  assert.deepEqual(decodeAddress('$1'), {address: '1', col: undefined, row: 1});
  assert.deepEqual(decodeAddress('$A'), {address: 'A', col: 1, row: undefined});
});

test('decodeAddress rejects an empty reference', () => {
  assert.throws(() => decodeAddress('$'), SyntaxError);
  assert.throws(() => decodeAddress(''), SyntaxError);
});

// The narrowing every caller that needs an actual cell used to re-derive: a bare row, a bare column
// and a range each parse fine but name no cell, and that is the case worth a single spelling.
const NOT_ONE_CELL = ['1', '$1', 'A', '$A', 'A1:B2', '$', '', 'a1', 'Sheet1!A1'];

test('decodeCellRef narrows a full reference to its two axes', () => {
  assert.deepEqual(decodeCellRef('B2'), {col: 2, row: 2});
  assert.deepEqual(decodeCellRef('$B$2'), {col: 2, row: 2}, 'anchors are dropped, not rejected');
  assert.deepEqual(decodeCellRef('XFD1048576'), {col: MAX_COLUMN, row: 1048576});
});

test('decodeCellRef refuses everything that does not name one cell', () => {
  for (const reference of NOT_ONE_CELL) {
    assert.throws(() => decodeCellRef(reference), SyntaxError, `"${reference}" names no one cell`);
  }
});

test('a column past XFD is a bounds failure, not a syntax one', () => {
  assert.throws(() => decodeCellRef('XFE1'), RangeError);
});

test('tryDecodeCellRef answers undefined where decodeCellRef throws, on either error', () => {
  assert.deepEqual(tryDecodeCellRef('B2'), {col: 2, row: 2});
  for (const reference of [...NOT_ONE_CELL, 'XFE1']) {
    assert.equal(tryDecodeCellRef(reference), undefined, `"${reference}" names no one cell`);
  }
});

test('decodeCellRef bounds the row where it already bounded the column', () => {
  // A decoder that returned row 0 handed it to every caller, and `getCell('A0')` stored it before
  // the cell constructor refused it, leaving a sheet that threw on every later read or write.
  assert.throws(() => decodeCellRef('A0'), {
    name: 'RangeError',
    message: `row 0 is out of bounds: Excel supports 1..${MAX_ROW}`,
  });
  assert.throws(() => decodeCellRef(`A${MAX_ROW + 1}`), RangeError);
  assert.throws(() => decodeAddress('0'), RangeError, 'a bare row is bounded too');
  assert.throws(() => decodeRange(`A1:B${MAX_ROW + 1}`), RangeError, 'and a range endpoint');
});

// The read side's question is not "does this parse" but "can this cell exist", and the bound that
// answers it now lives in the throwing decoder, so the tolerant one only has to catch.
test('tryDecodeCellRef refuses a reference that parses but names no possible cell', () => {
  assert.deepEqual(tryDecodeCellRef(`XFD${MAX_ROW}`), {col: MAX_COLUMN, row: MAX_ROW});
  for (const reference of ['A0', `A${MAX_ROW + 1}`, 'ZZZZ1']) {
    assert.equal(tryDecodeCellRef(reference), undefined, `"${reference}" is off the grid`);
  }
});

test('tryDecodeRange answers undefined where decodeRange throws', () => {
  assert.equal(tryDecodeRange('B2:D6')?.dimensions, 'B2:D6');
  for (const reference of ['junk!!', 'ZZZZ0:!!', '', ':']) {
    assert.equal(tryDecodeRange(reference), undefined, `"${reference}" names no region`);
  }
});

// An axis neither endpoint mentions is unbounded, not unreadable: a whole-column range is a
// legitimate reference, and a caller that needs a bounded rectangle checks the corners itself.
test('tryDecodeRange keeps an unbounded range and refuses one off the grid', () => {
  assert.equal(tryDecodeRange('A:A')?.dimensions, 'A:A');
  assert.equal(tryDecodeRange('1:1')?.dimensions, '1:1');
  assert.equal(tryDecodeRange(`A1:XFD${MAX_ROW}`)?.bottom, MAX_ROW);
  for (const reference of ['A0:B2', `A1:B${MAX_ROW + 1}`, 'A1:ZZZZ2']) {
    assert.equal(tryDecodeRange(reference), undefined, `"${reference}" leaves the grid`);
  }
});

test('decodeRange resolves an ordinary rectangle and normalizes corner order', () => {
  const range = decodeRange('B2:D6');
  assert.equal(range.top, 2);
  assert.equal(range.left, 2);
  assert.equal(range.bottom, 6);
  assert.equal(range.right, 4);
  assert.equal(range.dimensions, 'B2:D6');
  // reversed input yields the same normalized corners
  assert.deepEqual(decodeRange('D6:B2'), range);
});

test('decodeRange on a whole-row range leaks no undefined/NaN and keeps row bounds', () => {
  const range = decodeRange('$1:$1');
  assert.equal(range.top, 1);
  assert.equal(range.bottom, 1);
  assert.equal(range.left, undefined);
  assert.equal(range.right, undefined);
  assert.equal(range.dimensions, '1:1');
  const serialized = JSON.stringify(range);
  assert.ok(!serialized.includes('undefined'), serialized);
  assert.ok(!serialized.includes('NaN'), serialized);
});

test('decodeRange on a whole-column range is symmetric', () => {
  const range = decodeRange('$A:$C');
  assert.equal(range.left, 1);
  assert.equal(range.right, 3);
  assert.equal(range.top, undefined);
  assert.equal(range.bottom, undefined);
  assert.equal(range.dimensions, 'A:C');
});

test('decodeRange carries a quoted sheet name, unescaping doubled apostrophes', () => {
  assert.equal(decodeRange('Sheet1!A1:B2').sheetName, 'Sheet1');
  assert.equal(decodeRange("'Bob''s data'!A1:B2").sheetName, "Bob's data");
  assert.equal(decodeRange('A1:B2').sheetName, undefined);
});

test('encodeAddress round-trips with decodeAddress', () => {
  assert.equal(encodeAddress(2, 6), 'B6');
  assert.equal(encodeAddress(MAX_COLUMN, 1048576), 'XFD1048576');
  const decoded = decodeAddress(encodeAddress(30, 42));
  assert.equal(decoded.col, 30);
  assert.equal(decoded.row, 42);
});

test('one coordinate mistake gets one message, whichever door the caller came through', () => {
  // Measured before the guards were shared, on a row 0 offered three ways:
  //   getRange(0,1,1,1)  ->  "row 0 is out of bounds: rows start at 1"
  //   getRow(0)          ->  "row 0 is out of bounds: Excel supports 1..1048576"
  //   spliceRows(0, 1)   ->  "splice start 0 is out of bounds: rows start at 1"
  // Three answers to one question, from a library whose addressing module states in a comment that
  // these "refuse the same mistake with the same words".
  const sheet = new Workbook().addWorksheet('S');
  const expected = 'row 0 is out of bounds: Excel supports 1..1048576';
  const message = (run: () => unknown): string => {
    try {
      run();
    } catch (error) {
      return (error as Error).message;
    }
    return '<no throw>';
  };
  assert.equal(
    message(() => sheet.getRange(0, 1, 1, 1)),
    expected,
  );
  assert.equal(
    message(() => sheet.getRow(0)),
    expected,
  );
  assert.equal(
    message(() => sheet.spliceRows(0, 1)),
    expected,
  );
  assert.equal(
    message(() => sheet.getRange(1, 0, 1, 1)),
    'column 0 is out of bounds: Excel supports 1..16384',
    'and the column axis says the same thing about itself',
  );
});

test('a coordinate past the end of the grid is refused on both axes, not just the column', () => {
  // `numberToColumn` has always bounded the column; the row half of `encodeAddress` checked only
  // that it was at least 1, so one axis produced an address for a position Excel has no reference
  // for while the other refused.
  assert.throws(() => encodeAddress(1, MAX_ROW + 1), RangeError);
  assert.throws(() => encodeAddress(MAX_COLUMN + 1, 1), RangeError);
  assert.equal(encodeAddress(MAX_COLUMN, MAX_ROW), 'XFD1048576');
});

// Observed in Excel 16.0 build 20326, twice over: the names it quotes in a formula's sheet prefix, and
// the table names that make it offer to repair a package. The two sets agreed on every name tried.
const NAMES_READ_AS_REFERENCES = [
  // A1 cells on the grid, in any case and with a leading zero.
  ['A1', 'T1', 'Q1', 'Z1', 'c1', 'A01', 'R01', 'RR1', 'CR1', 'Rx1', 'XFD1048576', 'C16385'],
  // R1C1 references, whole: a row, a column, both, each with or without its number.
  ['R', 'C', 'r', 'RC', 'Rc', 'rC1', 'R1C', 'RC1', 'R1C1', 'R1048576', 'C16384'],
  // An R1C1 row or column number followed by a name character.
  ['R1X', 'R1A', 'C1X', 'R1C0', 'R1R1', 'C1C1', 'R1C1A', 'R1C1X', 'R1C1_', 'R1C1.x', 'R2C3D4'],
  ['TRUE', 'FALSE', 'True', 'fAlse'],
].flat();

const NAMES_NOT_READ_AS_REFERENCES = [
  // Past the grid, or naming row 0.
  ['XFE1', 'XFD1048577', 'A1048577', 'R1048577', 'R99999999', 'AAAA1', 'A0', 'R0', 'RC0', 'R0C1'],
  // No R1C1 number for a name character to follow, or no name character after it.
  ['RCX', 'RX', 'CX', 'RC_', 'R_1', 'C_', 'R1.5', 'R1.A', 'RC.1'],
  ['TRUE1', 'TRUEX', 'A1B', 'AB1C', 'Q1X', 'CC', 'RR', 'T', 'Table1', 'Sales.1'],
].flat();

test('nameReadsAsReference answers what Excel reads as a reference where it expects a name', () => {
  for (const name of NAMES_READ_AS_REFERENCES) assert.ok(nameReadsAsReference(name), name);
  for (const name of NAMES_NOT_READ_AS_REFERENCES) assert.ok(!nameReadsAsReference(name), name);
});
