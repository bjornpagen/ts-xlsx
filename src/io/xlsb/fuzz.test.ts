import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import path from 'node:path';
import {test} from 'node:test';
import {fileURLToPath} from 'node:url';

import {unzipSync, zipSync} from 'fflate';

import type {Workbook} from '../../core/workbook.ts';
import {xorshift32} from '../../fuzz.test-support.ts';
import {PackageReadError, UnsupportedFormatError} from '../opc/errors.ts';
import {XlsbParseError} from './errors.ts';
import {readXlsb} from './read.ts';
import {readRecords} from './record-stream.ts';
import {BRT} from './record-types.ts';
import {concat, frame, nameCitations, wide, word} from './records.test-support.ts';

// An adversarial pass over the BIFF12 reader.
//
// The contract under test is not "these mutations produce these workbooks": a mutated file has no
// correct reading. It is the *hostile-input* contract: whatever the bytes say, the reader either
// produces a model or fails closed with a typed error, in bounded time and bounded memory. Two
// failure shapes are called out specifically because each is a bug wearing a crash's clothes:
//
// - a `TypeError` or `RangeError` escaping means an index or a length reached the model unchecked
//   (`undefined` folded into arithmetic, an address outside the grid handed to the encoder);
// - a run that does not finish means a count taken from the file became a loop bound.
//
// The corpus fixture is the seed rather than random noise, so every mutation lands inside a structure
// real enough to reach deep into the parsers instead of being rejected at the first record.

const FIXTURE = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  '../../../test/corpus/fixtures/xlsb-binary-workbook-reads-like-its-xlsx-twin/source.xlsb',
);

// The binary parts: the only ones this reader parses, and so the only ones worth mutating.
const BINARY_PARTS = [
  'xl/workbook.bin',
  'xl/worksheets/sheet1.bin',
  'xl/worksheets/sheet2.bin',
  'xl/worksheets/sheet3.bin',
  'xl/styles.bin',
  'xl/sharedStrings.bin',
];

// A formula literal JavaScript spells but no spreadsheet grammar does. Eight mutated bytes of a
// `PtgNum` decode to one of these as readily as to any other double, and the decoder used to render
// whatever `String(value)` gave it: the text `INFINITY` then reached the model, was escaped as
// ordinary formula text by the writer, and produced a package Excel reports as damaged. A read that
// succeeds is only a success if what it produced can be written back out.
const UNSPELLABLE_LITERAL = /\b(?:INFINITY|NAN)\b/;

/** Read a mutated package, asserting only that it fails the way a reader is allowed to fail. */
function readOrFailClosed(archive: Uint8Array, label: string): void {
  let workbook: Workbook;
  try {
    workbook = readXlsb(archive);
  } catch (error) {
    // A mutation inside a compressed part can break the deflate stream itself, so the archive failing
    // to unpack (`PackageReadError`) is as legitimate a closed failure as a part failing to parse.
    if (
      error instanceof XlsbParseError ||
      error instanceof UnsupportedFormatError ||
      error instanceof PackageReadError
    ) {
      return;
    }
    // An `authoring` failure is never a closed one on this path. That code blames the caller, and the
    // caller here did nothing but open a file: a sheet name a mutation duplicated, a merge it made
    // overlap, reaches the model through the read repair or it is a missing repair.
    assert.ok(
      error instanceof Error &&
        !(error instanceof TypeError) &&
        !(error instanceof RangeError) &&
        (error as {code?: unknown}).code !== 'authoring',
      `${label}: expected a typed, closed failure but got ${String(error)}`,
    );
    return;
  }
  // Outside the `try`, deliberately: this assertion is the *test* failing, and the catch above would
  // read an `AssertionError` as one more closed failure and pass.
  assertWritableFormulas(workbook, label);
}

// Every formula the read produced, checked for a literal the format cannot carry. Throws an
// `AssertionError`, which `readOrFailClosed` deliberately does not catch: this is the test failing,
// not the reader failing closed.
function assertWritableFormulas(workbook: Workbook, label: string): void {
  for (const sheet of workbook.worksheets) {
    for (const {cells} of sheet.rows()) {
      for (const cell of cells) {
        const value: unknown = cell.value;
        if (value === null || typeof value !== 'object' || !('formula' in value)) continue;
        const formula = String(value.formula);
        assert.ok(
          !UNSPELLABLE_LITERAL.test(formula),
          `${label}: decoded a formula no spreadsheet can spell (${formula})`,
        );
      }
    }
  }
}

test('single-byte mutations anywhere in the binary parts never escape the typed failure modes', () => {
  const parts = unzipSync(readFileSync(FIXTURE));
  const next = xorshift32(0x5eed1);
  for (let round = 0; round < 300; round++) {
    const partName = BINARY_PARTS[next() % BINARY_PARTS.length] ?? '';
    const original = parts[partName];
    assert.ok(original, `fixture is missing ${partName}`);
    const mutated = Uint8Array.from(original);
    const offset = next() % mutated.length;
    mutated[offset] = next() & 0xff;
    readOrFailClosed(
      zipSync({...parts, [partName]: mutated}),
      `${partName}@${offset} round ${round}`,
    );
  }
});

test('a part truncated at any point fails closed rather than reading half a workbook into a crash', () => {
  const parts = unzipSync(readFileSync(FIXTURE));
  const next = xorshift32(0x7ac6);
  for (let round = 0; round < 200; round++) {
    const partName = BINARY_PARTS[next() % BINARY_PARTS.length] ?? '';
    const original = parts[partName];
    assert.ok(original);
    const cut = next() % original.length;
    readOrFailClosed(
      zipSync({...parts, [partName]: original.subarray(0, cut)}),
      `${partName} cut at ${cut}`,
    );
  }
});

test('a record length inflated to the maximum is rejected, not allocated', () => {
  // The direct expression of the attack the framing guard exists for: every plausible size prefix
  // rewritten to the largest the encoding can express, one record at a time.
  const parts = unzipSync(readFileSync(FIXTURE));
  for (const partName of BINARY_PARTS) {
    const original = parts[partName];
    assert.ok(original);
    // Every other byte: a size prefix is at least two bytes wide once inflated, so a stride of two
    // still lands on each record's prefix while halving a pass that is otherwise the slowest here.
    for (let offset = 0; offset < Math.min(original.length, 300); offset += 2) {
      const mutated = Uint8Array.from(original);
      mutated[offset] = 0xff;
      mutated[offset + 1] = 0xff;
      mutated[offset + 2] = 0xff;
      mutated[offset + 3] = 0x7f;
      readOrFailClosed(zipSync({...parts, [partName]: mutated}), `${partName}@${offset}`);
    }
  }
});

test('a cell, row, or column addressed outside the grid is dropped, not encoded', () => {
  // Positional fields are the one place a lying number cannot simply be believed: an address beyond
  // Excel's grid has no representation, and a column *run* beyond it is a loop bound. Both are driven
  // to their extreme here, with every 32-bit positional field set to its maximum.
  const parts = unzipSync(readFileSync(FIXTURE));
  const sheet = parts['xl/worksheets/sheet2.bin'];
  assert.ok(sheet);
  // Stride 3 against 4-byte fields, so every positional field is struck at more than one alignment.
  for (let offset = 0; offset + 4 <= sheet.length; offset += 3) {
    const mutated = Uint8Array.from(sheet);
    mutated[offset] = 0xff;
    mutated[offset + 1] = 0xff;
    mutated[offset + 2] = 0xff;
    mutated[offset + 3] = 0x0f;
    readOrFailClosed(zipSync({...parts, 'xl/worksheets/sheet2.bin': mutated}), `grid@${offset}`);
  }
});

test('a cell formula citing a megabyte defined name hundreds of times fails no worse than closed', () => {
  // Random mutation cannot find this one: it needs a long name *and* a stream that cites it, and each
  // is harmless without the other. The decoded text would be the name's length times the citations.
  const parts = unzipSync(readFileSync(FIXTURE));
  const workbookPart = parts['xl/workbook.bin'];
  const sheetPart = parts['xl/worksheets/sheet1.bin'];
  assert.ok(workbookPart && sheetPart);
  // A `PtgName` cites by 1-based position among every `BrtName` in the file.
  const index = [...readRecords(workbookPart)].filter(({type}) => type === BRT.Name).length + 1;
  const longName = frame(
    BRT.Name,
    concat(
      word(0),
      Uint8Array.of(0),
      word(0xffffffff),
      wide('x'.repeat(1 << 20)),
      word(3),
      Uint8Array.of(0x1e, 1, 0),
      word(0),
    ),
  );
  const rgce = nameCitations(600, index);
  const citingCell = concat(
    frame(BRT.RowHdr, concat(word(199), word(0), Uint8Array.of(0, 0, 0, 0))),
    frame(
      BRT.FmlaNum,
      concat(
        word(0),
        word(0),
        new Uint8Array(8),
        Uint8Array.of(0, 0),
        word(rgce.length),
        rgce,
        word(0),
      ),
    ),
  );
  readOrFailClosed(
    zipSync({
      ...parts,
      'xl/workbook.bin': concat(workbookPart, longName),
      'xl/worksheets/sheet1.bin': concat(sheetPart, citingCell),
    }),
    'a megabyte name cited 600 times',
  );
});

test('a deeply repeated collection marker does not accumulate unbounded state', () => {
  // The style sheet's Begin/End collection markers are the reader's only nesting-shaped state. A file
  // that repeats one a hundred thousand times must cost a hundred thousand cheap assignments, not a
  // hundred thousand stack frames or a growing stack of contexts.
  const parts = unzipSync(readFileSync(FIXTURE));
  const beginFonts = Uint8Array.of(0xeb, 0x04, 0x04, 0x01, 0x00, 0x00, 0x00);
  const flood = new Uint8Array(beginFonts.length * 100_000);
  for (let index = 0; index < 100_000; index++) flood.set(beginFonts, index * beginFonts.length);
  readOrFailClosed(zipSync({...parts, 'xl/styles.bin': flood}), 'repeated BrtBeginFonts');
});
