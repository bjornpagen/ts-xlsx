import assert from 'node:assert/strict';
import test from 'node:test';

import {Workbook} from '../../core/workbook.ts';
import {AuthoringError} from '../../errors.ts';
import {writeCsv, writeCsvText} from './write.ts';

test('each row is sized to its own extent, not clamped to a narrower earlier row', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.addRow(['a']);
  sheet.addRow(['b', 'c', 'd']);
  assert.equal(writeCsvText(wb), 'a\nb,c,d');
});

test('a number renders bare and fields join on the configured delimiter', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addRow(['a', 'b', 'c']);
  assert.equal(writeCsvText(wb, {delimiter: ';'}), 'a;b;c');
});

test('sheetName selects the worksheet; the default is the first', () => {
  const wb = new Workbook();
  wb.addWorksheet('First').addRow(['a', 1]);
  const second = wb.addWorksheet('Second');
  second.addRow(['b', 2]);
  second.addRow(['c', 3]);
  assert.equal(writeCsvText(wb), 'a,1');
  assert.equal(writeCsvText(wb, {sheetName: 'Second'}), 'b,2\nc,3');
});

test('a name matching no worksheet throws rather than emitting empty output', () => {
  const wb = new Workbook();
  wb.addWorksheet('First').addRow(['a']);
  assert.throws(
    () => writeCsvText(wb, {sheetName: 'Nope'}),
    /no worksheet "Nope"; this workbook has "First"/,
  );
});

test('a Date renders in a token format in UTC, or as a full ISO timestamp by default', () => {
  const wb = new Workbook();
  const noon = new Date('2018-01-05T12:00:00.000Z');
  wb.addWorksheet('S').addRow([noon]);
  assert.equal(writeCsvText(wb, {dateFormat: 'MM/DD/YYYY', dateUTC: true}), '01/05/2018');
  assert.match(writeCsvText(wb, {dateUTC: true}), /^2018-01-05T12:00:00/);
});

test('a formula cell renders its cached result', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addRow([{formula: '1+1', result: 2}]);
  assert.equal(writeCsvText(wb), '2');
});

test('a field carrying the delimiter, a quote, or a newline is quoted and escaped', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addRow(['a,b', 'he said "hi"', 'line\nbreak']);
  assert.equal(writeCsvText(wb), '"a,b","he said ""hi""","line\nbreak"');
});

test('writeCsv prepends a UTF-8 BOM for the default UTF-8 encoding', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addRow(['café']);
  const bytes = writeCsv(wb);
  assert.deepEqual([bytes[0], bytes[1], bytes[2]], [0xef, 0xbb, 0xbf]);
  assert.equal(Buffer.from(bytes.slice(3)).toString('utf8'), 'café');
});

test('writeCsv honours a requested non-UTF-8 encoding and adds no BOM there', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addRow(['café']);
  const bytes = writeCsv(wb, {encoding: 'utf16le'});
  assert.equal(Buffer.from(bytes).toString('utf16le'), 'café');
  assert.notEqual(Buffer.from(bytes).toString('utf8'), 'café');
});

test('map overrides the default rendering per field, receiving the value and column index', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 7;
  sheet.getCell('C1').value = 'x';
  assert.equal(
    // The cast is this test's claim about its own fixture: the sheet holds one number and one
    // string, and `CellValue` at large includes shapes a template would render as `[object Object]`.
    writeCsvText(wb, {
      map: (value, index) => `${index}:${value === null ? '_' : (value as number | string)}`,
    }),
    '0:7,1:_,2:x',
  );
});

test("map's returned text is still quoted like any other field", () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addRow(['a']);
  assert.equal(writeCsvText(wb, {map: () => 'x,y'}), '"x,y"');
});

test('emoji and CJK survive the default UTF-8 byte path', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addRow(['😀🎉', '日本語']);
  const body = Buffer.from(writeCsv(wb).slice(3)).toString('utf8');
  assert.equal(body, '😀🎉,日本語');
});

// A lone surrogate is the one character the CSV encode step can lose in silence, because
// `Buffer.from` substitutes U+FFFD for it rather than failing. Refusing is what the XLSX writer's
// escapers do with the characters they cannot represent; CSV has no `_xHHHH_` to fall back on.

test('a lone surrogate is refused on the UTF-8 path rather than encoded as U+FFFD', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addRow(['a\uD800b']);

  assert.equal(
    writeCsvText(wb),
    'a\uD800b',
    'the logical text loses nothing: a JS string holds it',
  );
  assert.throws(
    () => writeCsv(wb),
    (error: unknown) => {
      assert.ok(error instanceof AuthoringError);
      assert.match(error.message, /U\+D800 at offset 1/);
      return true;
    },
  );
});

test('a lone surrogate passes through the UTF-16 path, which encodes it losslessly', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addRow(['a\uD800b']);
  const bytes = writeCsv(wb, {encoding: 'utf16le'});
  assert.equal(Buffer.from(bytes).toString('utf16le'), 'a\uD800b');
});

test('a well-formed astral pair is not mistaken for a lone surrogate', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addRow(['\u{1F600}']);
  assert.equal(Buffer.from(writeCsv(wb).slice(3)).toString('utf8'), '\u{1F600}');
});

test('the two spellings of UTF-8 are the same path, BOM and refusal both', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addRow(['café']);
  const bytes = writeCsv(wb, {encoding: 'utf-8'});
  assert.deepEqual([bytes[0], bytes[1], bytes[2]], [0xef, 0xbb, 0xbf]);

  const broken = new Workbook();
  broken.addWorksheet('S').addRow(['\uDC00']);
  assert.throws(() => writeCsv(broken, {encoding: 'utf-8'}), AuthoringError);
});

// The encoder used to be `Buffer.from(text, encoding)`, and dropping that Node global was the
// point of writing one (`scripts/check-browser-safe.ts`). Byte-for-byte agreement with Buffer is
// therefore the contract, not merely a plausible encoding: every spelling, and the astral and
// high-Latin characters where the byte-narrow encodings do something lossy on purpose.
test('every encoding produces exactly the bytes Buffer would have', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addRow(['héllo ÿ € 😀 漢字']);
  const text = writeCsvText(wb);
  for (const encoding of ['ascii', 'latin1', 'ucs-2', 'ucs2', 'utf-16le', 'utf16le'] as const) {
    assert.deepEqual(
      writeCsv(wb, {encoding}),
      new Uint8Array(Buffer.from(text, encoding)),
      encoding,
    );
  }
  for (const encoding of ['utf-8', 'utf8'] as const) {
    assert.deepEqual(
      writeCsv(wb, {bom: false, encoding}),
      new Uint8Array(Buffer.from(text, encoding)),
      encoding,
    );
  }
});

test('a delimiter the reader cannot honour is refused by the writer too', () => {
  // `readCsv(writeCsv(wb, {delimiter: "||"}))` used to throw on text this same codec produced, and
  // `{delimiter: ""}` was worse: `field.includes("")` holds for every field, so every field was
  // quoted and the output carried no separators at all.
  const wb = new Workbook();
  wb.addWorksheet('S').addRow(['a', 'b']);
  assert.throws(() => writeCsvText(wb, {delimiter: '||'}), RangeError);
  assert.throws(() => writeCsvText(wb, {delimiter: ''}), RangeError);
  assert.equal(writeCsvText(wb, {delimiter: ';'}), 'a;b');
});

test('a quote, CR or LF is refused as a delimiter: the reader gives each its own meaning', () => {
  // With `"` as the delimiter, `a"b"c` read back as the single field `abc`.
  const wb = new Workbook();
  wb.addWorksheet('S').addRow(['a', 'b']);
  for (const delimiter of ['"', '\r', '\n']) {
    assert.throws(() => writeCsvText(wb, {delimiter}), /reserved/, JSON.stringify(delimiter));
  }
});

test('a row delimiter that is empty, holds the field delimiter, or holds a quote is refused', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addRow(['a', 'b']);
  for (const rowDelimiter of ['', ',', ';,', '"']) {
    assert.throws(() => writeCsvText(wb, {rowDelimiter}), RangeError, JSON.stringify(rowDelimiter));
  }
  assert.equal(writeCsvText(wb, {rowDelimiter: '\r\n'}), 'a,b');
});

test('a field containing a custom row delimiter is quoted', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.addRow(['a|b', 'c']);
  sheet.addRow(['d']);
  assert.equal(writeCsvText(wb, {rowDelimiter: '|'}), '"a|b",c|d');
});

test('an empty row between populated ones is an empty line, so row positions survive', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'a';
  sheet.getCell('A3').value = 'c';
  assert.equal(writeCsvText(wb), 'a\n\nc');
});

test('empty rows before the first populated one are empty lines too', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('B3').value = 'x';
  assert.equal(writeCsvText(wb), '\n\n,x');
});

test('a row holding only formatting writes the same text as a row holding nothing', () => {
  const plain = new Workbook();
  const plainSheet = plain.addWorksheet('S');
  plainSheet.getCell('A1').value = 'a';
  plainSheet.getCell('A3').value = 'c';

  const formatted = new Workbook();
  const formattedSheet = formatted.addWorksheet('S');
  formattedSheet.getCell('A1').value = 'a';
  formattedSheet.getCell('A3').value = 'c';
  formattedSheet.getRow(2).hidden = true;
  formattedSheet.getRow(5).height = 30;

  assert.equal(writeCsvText(formatted), writeCsvText(plain));
});
