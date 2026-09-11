import assert from 'node:assert/strict';
import test from 'node:test';

import {MAX_COLUMN, MAX_ROW} from '../../core/address.ts';
import {XlsxError} from '../../errors.ts';
import {CsvParseError} from './errors.ts';
import {readCsv} from './read.ts';

// The single worksheet's rows as a plain 2-D array of cell values, for terse assertions.
function rowsOf(csv: string, options?: Parameters<typeof readCsv>[1]): unknown[][] {
  const sheet = readCsv(csv, options).worksheets[0]!;
  const rows: unknown[][] = [];
  for (const {cells} of sheet.rows()) {
    let width = 0;
    for (const cell of cells) if (cell.col > width) width = cell.col;
    const fields: unknown[] = new Array(width).fill(null);
    for (const cell of cells) fields[cell.col - 1] = cell.value;
    rows.push(fields);
  }
  return rows;
}

test('a configured delimiter splits fields and numeric fields coerce', () => {
  assert.deepEqual(rowsOf('a;b;c\n1;2;3', {delimiter: ';'}), [
    ['a', 'b', 'c'],
    [1, 2, 3],
  ]);
});

test('a non-single-character delimiter is rejected rather than silently collapsing rows', () => {
  // Native `RangeError`, not the library's taxonomy: one argument out of range. Asserted as *not*
  // an XlsxError too, so a later re-wrap reddens the suite rather than quietly changing the catch.
  assert.throws(() => readCsv('a,b,c', {delimiter: ''}), {name: 'RangeError'});
  assert.throws(() => readCsv('a::b::c', {delimiter: '::'}), {name: 'RangeError'});
  assert.throws(
    () => readCsv('a::b::c', {delimiter: '::'}),
    (error: unknown) => !(error instanceof XlsxError),
  );
});

test('a quote, CR or LF is rejected as a delimiter rather than read as two things at once', () => {
  // `a"b"c` with `"` as the delimiter read as the single field `abc`.
  for (const delimiter of ['"', '\r', '\n']) {
    assert.throws(() => readCsv('a"b"c', {delimiter}), {name: 'RangeError'});
  }
});

test('an over-precision numeric string is preserved verbatim; in-range numbers coerce', () => {
  const big = '56343416020533614003';
  assert.deepEqual(rowsOf(`${big},42\n1.5,7`), [
    [big, 42],
    [1.5, 7],
  ]);
});

test('a leading-zero id coerces to a number by default but survives the identity map', () => {
  assert.deepEqual(rowsOf('007,32.5'), [[7, 32.5]]);
  assert.deepEqual(rowsOf('007,32.5', {map: (v) => v}), [['007', '32.5']]);
});

test('padded ids and dash-codes stay strings; a strict ISO date becomes a Date', () => {
  assert.deepEqual(rowsOf('2020-00001,1-3,3-4'), [['2020-00001', '1-3', '3-4']]);
  const cell = rowsOf('2018-01-05')[0]![0];
  assert.ok(cell instanceof Date);
  assert.equal(cell.toISOString(), '2018-01-05T00:00:00.000Z');
});

test('a whitespace-only field is a string, an empty field is null: neither is 0', () => {
  const [row] = rowsOf('firstValue,   ,secondValue\n');
  assert.equal(typeof row![1], 'string');
  assert.notEqual(row![1], 0);
  assert.equal(rowsOf('firstValue,,secondValue')[0]![1], null);
});

test('header mode consumes the first line, leaving data rows', () => {
  assert.deepEqual(rowsOf('name,age\nalice,30', {headers: true}), [['alice', 30]]);
  assert.deepEqual(rowsOf('name,age\nalice,30'), [
    ['name', 'age'],
    ['alice', 30],
  ]);
});

test('quoted fields carry embedded delimiters, quotes, and newlines', () => {
  assert.deepEqual(rowsOf('"a,b","he said ""hi""","line\nbreak"'), [
    ['a,b', 'he said "hi"', 'line\nbreak'],
  ]);
});

test('a quote inside an unquoted field is a literal character, not a field opener', () => {
  assert.deepEqual(rowsOf(`John,5'10",tall\nJane,5'2",short`), [
    ['John', `5'10"`, 'tall'],
    ['Jane', `5'2"`, 'short'],
  ]);
  assert.deepEqual(rowsOf('a,b"\nc,d'), [
    ['a', 'b"'],
    ['c', 'd'],
  ]);
});

test('text after a closing quote is kept literally, including a later quote', () => {
  assert.deepEqual(rowsOf('"ab"cd,x'), [['abcd', 'x']]);
  assert.deepEqual(rowsOf('"ab"c"d,x'), [['abc"d', 'x']]);
});

test('an ISO date naming no calendar day stays a string', () => {
  assert.deepEqual(rowsOf('2024-02-30,2023-04-31,2023-02-29'), [
    ['2024-02-30', '2023-04-31', '2023-02-29'],
  ]);
  const leap = rowsOf('2024-02-29')[0]![0];
  assert.ok(leap instanceof Date);
  assert.equal(leap.toISOString(), '2024-02-29T00:00:00.000Z');
});

test('a two-digit-era year is that year, not 1900 plus it', () => {
  const cell = rowsOf('0099-01-01 12:30')[0]![0];
  assert.ok(cell instanceof Date);
  assert.equal(cell.getUTCFullYear(), 99);
  assert.equal(cell.toISOString(), '0099-01-01T12:30:00.000Z');
});

test('a decimal with more significant digits than a double holds stays text', () => {
  const pi = '3.14159265358979323846';
  const card = '4111111111111111';
  assert.deepEqual(rowsOf(`${pi},${card}`), [[pi, card]]);
  // Fifteen significant digits fit; leading and trailing zeros only place the decimal point.
  assert.deepEqual(rowsOf('3.14159265358979,0.000123456789012345000,123456789012345'), [
    [3.14159265358979, 0.000123456789012345, 123456789012345],
  ]);
});

test('a record wider than the grid throws CsvParseError naming the record', () => {
  const wide = new Array(MAX_COLUMN + 1).fill('x').join(',');
  assert.throws(
    () => readCsv(`a,b\n${wide}`),
    (error: unknown) =>
      error instanceof CsvParseError &&
      error.code === 'malformed-input' &&
      /record 2 has more than 16384 fields/.test(error.message),
  );
  const full = new Array(MAX_COLUMN).fill('x').join(',');
  assert.equal(rowsOf(full)[0]!.length, MAX_COLUMN);
});

test('more data records than the grid has rows throws CsvParseError', () => {
  assert.throws(
    () => readCsv('x\n'.repeat(MAX_ROW + 1)),
    (error: unknown) => error instanceof CsvParseError && /more than 1048576/.test(error.message),
  );
});

test('a leading UTF-8 BOM is stripped and bytes decode', () => {
  const withBom = new Uint8Array([0xef, 0xbb, 0xbf, ...Buffer.from('x,y', 'utf8')]);
  assert.deepEqual(rowsOfBytes(withBom), [['x', 'y']]);
});

function rowsOfBytes(bytes: Uint8Array): unknown[][] {
  const sheet = readCsv(bytes).worksheets[0]!;
  const rows: unknown[][] = [];
  for (const {cells} of sheet.rows()) rows.push(cells.map((c) => c.value));
  return rows;
}
