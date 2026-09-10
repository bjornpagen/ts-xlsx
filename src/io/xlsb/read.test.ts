import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import path from 'node:path';
import {test} from 'node:test';
import {fileURLToPath} from 'node:url';

import {strToU8, unzipSync, zipSync} from 'fflate';

import type {Workbook} from '../../core/workbook.ts';
import {UnsupportedFormatError} from '../opc/errors.ts';
import {foreignPackage, foreignSheet} from '../xlsx/package.test-support.ts';
import {readXlsx} from '../xlsx/read.ts';
import {XlsbParseError} from './errors.ts';
import {RecordReader} from './primitives.ts';
import {readXlsb} from './read.ts';
import {BRT} from './record-types.ts';
import {concat, frame, reframe, wide, word} from './records.test-support.ts';

// The corpus owns the implementation-blind "reads like its XML twin" property; these are the
// white-box checks that go with the modules: the wiring, and the failure modes a fixture cannot show.
const FIXTURE = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  '../../../test/corpus/fixtures/xlsb-binary-workbook-reads-like-its-xlsx-twin/source.xlsb',
);

const fixture = (): Uint8Array => readFileSync(FIXTURE);

test('readXlsb reads a binary package into the workbook model', () => {
  const workbook = readXlsb(fixture());
  assert.deepEqual(
    workbook.worksheets.map((sheet) => sheet.name),
    ['Quiet', 'Grid', 'Values'],
  );
  assert.equal(workbook.getWorksheet('Values')?.getCell('B2').value, 10);
});

test('readXlsx auto-detects a binary package and produces the same workbook readXlsb does', () => {
  // The auto-detect path must not be a second, subtly different reader: it is the same codec, handed
  // an already-inflated package.
  const auto = readXlsx(fixture());
  const explicit = readXlsb(fixture());
  const project = (workbook: ReturnType<typeof readXlsb>): string =>
    JSON.stringify(workbook.worksheets.map((sheet) => sheet.model));
  assert.equal(project(auto), project(explicit));
});

test('readXlsb rejects an XML .xlsx package rather than reading it as empty', () => {
  const archive = zipSync({
    '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types/>'),
    'xl/workbook.xml': strToU8('<workbook/>'),
  });
  assert.throws(
    () => readXlsb(archive),
    (error: unknown) => {
      assert.ok(error instanceof UnsupportedFormatError);
      assert.match(error.message, /workbook\.bin/);
      return true;
    },
  );
});

test('a binary workbook part with a lying record length fails closed', () => {
  const archive = zipSync({
    'xl/workbook.bin': Uint8Array.of(0x81, 0x00, 0x02, 0xff, 0xff, 0xff, 0x7f, 0x01),
  });
  assert.throws(() => readXlsb(archive), XlsbParseError);
});

test('a sheet whose relationship target is missing yields an empty sheet, not a crash', () => {
  // A dangling reference is a damaged file, not a hostile one: the workbook still names the sheet, so
  // the model keeps it rather than dropping the sheet the caller can see in Excel's tab bar.
  const workbook = readXlsb(withoutPart('xl/worksheets/sheet3.bin'));
  assert.equal(workbook.worksheets.length, 3);
  assert.equal(workbook.getWorksheet('Values')?.model.cells.length, 0);
});

test('a package with no style sheet reads its values with every cell unstyled', () => {
  const workbook = readXlsb(withoutPart('xl/styles.bin'));
  const cell = workbook.getWorksheet('Values')?.getCell('A1');
  assert.equal(cell?.value, 'kind');
  assert.equal(cell?.font, undefined);
  assert.equal(cell?.fill, undefined);
});

test('a package with no shared-string table reads its pooled cells as empty strings', () => {
  const workbook = readXlsb(withoutPart('xl/sharedStrings.bin'));
  // The value is gone with the pool, but the cell, and the rest of the sheet, still reads.
  assert.equal(workbook.getWorksheet('Values')?.getCell('A1').value, '');
  assert.equal(workbook.getWorksheet('Values')?.getCell('B2').value, 10);
});

// Rebuild the fixture archive without one part, to exercise the reader's tolerance of a damaged
// package without committing a second fixture for each variant.
function withoutPart(dropped: string): Uint8Array {
  const parts = unzipSync(fixture());
  delete parts[dropped];
  return zipSync(parts);
}

// ── A file never raises an authoring failure ────────────────────────────────────────────────────────
// Each shape below made `readXlsb` throw `AuthoringError`, blaming a caller who had only opened a file,
// while the XML reader repaired or dropped the same shape. Each is built by re-framing one record of
// the Excel-authored fixture, and each is held to what `readXlsx` makes of the same shape in XML.

// The fixture with one part rewritten, failing the test if the rewrite changed nothing.
function withPart(name: string, edit: (part: Uint8Array) => Uint8Array): Uint8Array {
  const parts = unzipSync(fixture());
  const part = parts[name];
  assert.ok(part, `expected part ${name}`);
  const edited = edit(part);
  assert.notDeepEqual(edited, part, `edit to ${name} changed nothing`);
  return zipSync({...parts, [name]: edited});
}

// `BrtBundleSh` ends with the sheet name, so renaming a sheet replaces the payload's tail.
function renameSheets(...names: string[]): (part: Uint8Array) => Uint8Array {
  let index = 0;
  return (part) =>
    reframe(part, (type, data) => {
      const name = type === BRT.BundleSh ? names[index++] : undefined;
      if (name === undefined) return frame(type, data);
      const reader = new RecordReader(data);
      reader.skip(8);
      reader.nullableWideString();
      const declared = reader.wideString();
      return frame(
        type,
        concat(data.subarray(0, data.length - 4 - declared.length * 2), wide(name)),
      );
    });
}

// An XML workbook with these sheets, the first carrying `firstSheet`, and `definedNames` after them.
function xmlWorkbook(
  names: readonly string[],
  {firstSheet = foreignSheet(''), definedNames = ''} = {},
): Workbook {
  return readXlsx(
    foreignPackage({
      'xl/workbook.xml':
        '<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
        names
          .map((name, i) => `<sheet name="${name}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`)
          .join('') +
        `</sheets>${definedNames}</workbook>`,
      'xl/_rels/workbook.xml.rels':
        '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        names
          .map(
            (_, i) =>
              `<Relationship Id="rId${i + 1}" Type="x" Target="worksheets/sheet${i + 1}.xml"/>`,
          )
          .join('') +
        '</Relationships>',
      ...Object.fromEntries(
        names.map((_, i) => [
          `xl/worksheets/sheet${i + 1}.xml`,
          i === 0 ? firstSheet : foreignSheet(''),
        ]),
      ),
    }),
  );
}

const sheetNames = (workbook: Workbook): string[] => workbook.worksheets.map((sheet) => sheet.name);

test('two sheets whose names differ only in case both read, the second renamed as XML renames it', () => {
  const binary = readXlsb(withPart('xl/workbook.bin', renameSheets('S', 's')));
  assert.deepEqual(sheetNames(binary), ['S', 's (2)', 'Values']);
  assert.deepEqual(sheetNames(binary), sheetNames(xmlWorkbook(['S', 's', 'Values'])));
});

test('a sheet name past the length limit is truncated, as the XML reader truncates it', () => {
  const long = 'x'.repeat(40);
  const binary = readXlsb(withPart('xl/workbook.bin', renameSheets(long)));
  assert.equal(binary.worksheets[0]?.name, 'x'.repeat(31));
  assert.deepEqual(sheetNames(binary), sheetNames(xmlWorkbook([long, 'Grid', 'Values'])));
});

test('a merge overlapping one already read is dropped and the sheet still reads', () => {
  // Far from anything the fixture merges, so the only overlap is the one planted here.
  const merge = (rowFirst: number, colFirst: number): Uint8Array =>
    frame(
      BRT.MergeCell,
      concat(word(rowFirst), word(rowFirst + 1), word(colFirst), word(colFirst + 1)),
    );
  const binary = readXlsb(
    withPart('xl/worksheets/sheet1.bin', (part) => concat(part, merge(99, 26), merge(100, 27))),
  );
  const merges = binary.worksheets.flatMap((sheet) => [...sheet.merges]);
  assert.ok(merges.includes('AA100:AB101'), `the first merge reads; got ${merges.join(', ')}`);
  assert.ok(!merges.includes('AB101:AC102'), 'the overlapping one is dropped');

  const xml = xmlWorkbook(['S'], {
    firstSheet:
      '<?xml version="1.0"?><worksheet><sheetData/><mergeCells>' +
      '<mergeCell ref="AA100:AB101"/><mergeCell ref="AB101:AC102"/></mergeCells></worksheet>',
  });
  assert.deepEqual([...(xml.worksheets[0]?.merges ?? [])], ['AA100:AB101']);
});

test('a defined name the model refuses is dropped and the names after it still read', () => {
  // `BrtName`: flags, the Alt-key byte, a workbook-global scope, the name, then a one-token formula
  // (`PtgInt` 1) and an empty extra-data block.
  const name = (text: string): Uint8Array =>
    frame(
      BRT.Name,
      concat(
        word(0),
        Uint8Array.of(0),
        word(0xffffffff),
        wide(text),
        word(3),
        Uint8Array.of(0x1e, 1, 0),
        word(0),
      ),
    );
  const binary = readXlsb(
    withPart('xl/workbook.bin', (part) => concat(part, name(''), name('Kept'))),
  );
  const planted = (workbook: Workbook) =>
    workbook.definedNames.filter((defined) => defined.name === '' || defined.name === 'Kept');
  assert.deepEqual(planted(binary), [{name: 'Kept', refersTo: '1'}]);

  const xml = xmlWorkbook(['S'], {
    definedNames:
      '<definedNames><definedName name="">1</definedName><definedName name="Kept">1</definedName></definedNames>',
  });
  assert.deepEqual(planted(binary), planted(xml));
});
