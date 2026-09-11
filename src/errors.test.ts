import assert from 'node:assert/strict';
import {test} from 'node:test';

import type {CellValue} from './core/value.ts';
import {Workbook} from './core/workbook.ts';
import type {Worksheet} from './core/worksheet.ts';
import {CustomUiParseError} from './customui/errors.ts';
import {parseCustomUi} from './customui/ribbon.ts';
import {AuthoringError, InternalError, quoted, XlsxError, type XlsxErrorCode} from './errors.ts';
import {CsvParseError} from './io/csv/errors.ts';
import {PackageReadError, UnsupportedFormatError} from './io/opc/errors.ts';
import {XlsbParseError} from './io/xlsb/errors.ts';
import {XlsxParseError} from './io/xlsx/errors.ts';
import {writeXlsx} from './io/xlsx/write.ts';
import {VbaAuthorError, VbaParseError} from './vba/errors.ts';
import {XmlParseError} from './xml/errors.ts';
import {xmlEvents} from './xml/xml-scan.ts';

// Every class in the taxonomy, with the code it is contracted to carry. A class added without a
// row here is one whose category nobody chose. The `every class` tests below are only as complete
// as this table, so it is the thing to extend first.
const TAXONOMY: ReadonlyArray<
  // Two constructor parameters, not one: every subclass inherits Error's `(message, options)` and
  // none may shadow it away, which the cause test below calls directly. Declaring one parameter and
  // casting at the call site said the same thing in a place the type could not enforce it.
  readonly [new (message?: string, options?: ErrorOptions) => XlsxError, XlsxErrorCode]
> = [
  [AuthoringError, 'authoring'],
  [VbaAuthorError, 'authoring'],
  [InternalError, 'internal'],
  [XmlParseError, 'malformed-input'],
  [XlsxParseError, 'malformed-input'],
  [XlsbParseError, 'malformed-input'],
  [CsvParseError, 'malformed-input'],
  [VbaParseError, 'malformed-input'],
  [CustomUiParseError, 'malformed-input'],
  [PackageReadError, 'malformed-input'],
];

test('every taxonomy error is catchable as one XlsxError and as a plain Error', () => {
  for (const [Class] of TAXONOMY) {
    const error = new Class('boom');
    assert.ok(error instanceof XlsxError, `${Class.name} is not an XlsxError`);
    assert.ok(error instanceof Error, `${Class.name} is not an Error`);
    // `InternalError` appends a report pointer below the message and is asserted exactly in its own
    // test; every other class is contracted to hand back what it was given, unchanged.
    if (Class === InternalError) {
      assert.match(error.message, /^boom\n/, 'InternalError did not lead with its own message');
    } else {
      assert.equal(error.message, 'boom');
    }
  }
});

test('every taxonomy error reports its documented code and its own name', () => {
  const names = new Set<string>();
  for (const [Class, code] of TAXONOMY) {
    const error = new Class();
    assert.equal(error.code, code, `${Class.name} carries the wrong code`);
    assert.equal(error.name, Class.name, `${Class.name} does not report its own name`);
    names.add(error.name);
  }
  // `name` is the 1:1 discriminant, so a copy-pasted class that kept its neighbour's name would
  // make two distinct failures indistinguishable to a caller who cannot use `instanceof`.
  assert.equal(names.size, TAXONOMY.length);
});

test('a taxonomy error carries a cause through to the standard Error field', () => {
  const cause = new Error('underlying');
  for (const [Class] of TAXONOMY) {
    const error = new Class('wrapped', {cause});
    assert.equal(error.cause, cause, `${Class.name} dropped its cause`);
  }
});

test('UnsupportedFormatError keeps its format branch and takes a cause after the message', () => {
  const cause = new Error('underlying');
  const error = new UnsupportedFormatError('xls', 'custom', {cause});
  assert.ok(error instanceof XlsxError);
  assert.equal(error.code, 'unsupported-format');
  assert.equal(error.format, 'xls');
  assert.equal(error.message, 'custom');
  assert.equal(error.cause, cause);
});

test('UnsupportedFormatError still defaults its message per format', () => {
  assert.match(new UnsupportedFormatError('xlsb').message, /binary \.xlsb/);
});

// The report pointer is the only channel that reaches a caller who never reads our docs: it rides
// the stack trace they are already staring at. These pin the two halves of that bargain: the
// invariant stays first so the message is still diagnosable, and the pointer is actually there.
test('InternalError states the broken invariant first, then where to report it', () => {
  const error = new InternalError('pivot record references an uncatalogued item');
  assert.match(error.message, /^pivot record references an uncatalogued item\n\n/);
  assert.match(error.message, /bug in ts-xlsx/);
  assert.match(error.message, /https:\/\/github\.com\/shbernal\/ts-xlsx\/issues\/new\?template=/);
});

test('InternalError without a message reports the pointer alone, not a stringified undefined', () => {
  const message = new InternalError().message;
  assert.doesNotMatch(message, /undefined/);
  assert.match(message, /^This is a bug in ts-xlsx/);
});

test('malformed XML markup fails as a typed parse error, not a bare SyntaxError', () => {
  assert.throws(
    () => [...xmlEvents('<a')],
    (error: unknown) =>
      error instanceof XmlParseError &&
      error.code === 'malformed-input' &&
      /unterminated tag/.test(error.message),
  );
});

test('a wrapping parser keeps the failure it wrapped on cause', () => {
  assert.throws(
    () => parseCustomUi('<customUI'),
    (error: unknown) => error instanceof CustomUiParseError && error.cause instanceof XmlParseError,
  );
});

// The taxonomy is only worth its weight if the throw sites actually use it. These sample the two
// layers that used to throw a bare `Error`, the model and the writer, so a regression to
// `new Error(...)` in either fails here rather than only in the docs.
test('the model rejects an unauthorable document with an AuthoringError', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('Sheet1');
  assert.throws(() => workbook.addWorksheet('sheet1'), AuthoringError);
  assert.throws(() => workbook.addWorksheet(''), AuthoringError);

  const sheet = workbook.worksheets[0] as Worksheet;
  sheet.mergeCells('A1:B2');
  assert.throws(() => sheet.mergeCells('B2:C3'), AuthoringError);
});

test('the writer rejects an unwritable workbook with an AuthoringError', () => {
  assert.throws(() => writeXlsx(new Workbook()), AuthoringError);
});

// The writer's unhandled-value arms are `InternalError` rather than a "not implemented" report
// because the model refuses a foreign value long before serialisation. This is what proves it, and
// what would fail if the model's guard were ever loosened into the writer's lap.
test('a value the model does not admit never reaches the writer', () => {
  const sheet = new Workbook().addWorksheet('Sheet1');
  assert.throws(() => {
    sheet.getCell('A1').value = {nonsense: true} as unknown as CellValue;
  }, TypeError);
});

// One spelling for a name inside a message, and the one that survives a name a spreadsheet permits.
// The double-quoted and single-quoted literals this replaced both rendered the first case as a
// message whose reader cannot tell where the name ends.
test('quoted sets a name off from the prose without becoming ambiguous', () => {
  assert.equal(quoted('Sheet1'), '"Sheet1"');
  assert.equal(quoted('Q1 "draft"'), '"Q1 \\"draft\\""');
  assert.equal(quoted('two\nlines'), '"two\\nlines"');
  assert.equal(quoted('C:\\book.xlsx'), '"C:\\\\book.xlsx"');
});
