import assert from 'node:assert/strict';
import {test} from 'node:test';

import type {ArrayFormulaValue, CellValue} from '../../core/value.ts';
import {Workbook} from '../../core/workbook.ts';
import {parseDynamicArrayCellMetadata} from './cell-metadata.ts';
import {
  assertRelationshipsWired,
  elementIn,
  foreignPackage,
  foreignSheet,
  optionalPartIn,
  partIn,
  partsWritten,
  refuses,
  relationship,
  relationshipsPart,
  roundtrip,
  SHEET1,
  sheetXml,
} from './package.test-support.ts';
import {readSheetRows} from './read-rows.ts';
import {readXlsx} from './read.ts';
import {WorkbookStreamWriter} from './write-stream.ts';
import {writeXlsx} from './write.ts';

const SPILLING: ArrayFormulaValue = {
  shareType: 'array',
  formula: 'SEQUENCE(3)',
  ref: 'B1:B3',
  dynamic: true,
  result: 1,
};
const LEGACY: ArrayFormulaValue = {shareType: 'array', formula: 'A1:A3*2', ref: 'D1:D3', result: 2};

function workbookWith(cells: Record<string, CellValue>): Workbook {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('S');
  for (const [address, value] of Object.entries(cells)) sheet.getCell(address).value = value;
  return workbook;
}

// A package whose first sheet holds `rows` and whose workbook reaches `metadata` as its cell metadata.
function withCellMetadata(rows: string, metadata: string): Uint8Array {
  return foreignPackage({
    'xl/_rels/workbook.xml.rels': relationshipsPart(
      relationship('rId1', 'worksheet', 'worksheets/sheet1.xml') +
        relationship('rId2', 'sheetMetadata', 'metadata.xml'),
    ),
    'xl/metadata.xml': metadata,
    [SHEET1]: foreignSheet(rows),
  });
}

test('a legacy array formula is written over its range, with no cell metadata', () => {
  const parts = partsWritten(workbookWith({D1: LEGACY}));
  assert.equal(
    elementIn(partIn(parts, SHEET1), /<c r="D1"[\s\S]*?<\/c>/),
    '<c r="D1"><f t="array" ref="D1:D3">A1:A3*2</f><v>2</v></c>',
  );
  assert.equal(optionalPartIn(parts, 'xl/metadata.xml'), undefined, 'no part is written for it');
  assert.doesNotMatch(partIn(parts, '[Content_Types].xml'), /sheetMetadata/);
});

test('a dynamic array points its cell at the cell metadata that marks it', () => {
  const parts = partsWritten(workbookWith({B1: SPILLING, D1: LEGACY}));
  const sheet = partIn(parts, SHEET1);
  assert.equal(
    elementIn(sheet, /<c r="B1"[\s\S]*?<\/c>/),
    '<c r="B1" cm="1"><f t="array" ref="B1:B3">_xlfn.SEQUENCE(3)</f><v>1</v></c>',
    'the cell and formula Excel saves for the same formula',
  );
  assert.doesNotMatch(elementIn(sheet, /<c r="D1"[^>]*>/), /cm=/, 'the legacy one points nowhere');
  const cm = Number(elementIn(sheet, /<c r="B1" cm="(\d+)"/).match(/cm="(\d+)"/)?.[1]);
  assert.ok(parseDynamicArrayCellMetadata(partIn(parts, 'xl/metadata.xml')).has(cm));
  assert.match(
    partIn(parts, '[Content_Types].xml'),
    /<Override PartName="\/xl\/metadata.xml" ContentType="application\/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata\+xml"\/>/,
  );
  assert.match(
    partIn(parts, 'xl/_rels/workbook.xml.rels'),
    /relationships\/sheetMetadata" Target="metadata.xml"/,
  );
});

test('both kinds read back as written: range, mark and cached result', () => {
  const sheet = roundtrip(workbookWith({B1: SPILLING, D1: LEGACY})).getWorksheet('S');
  assert.deepEqual(sheet?.getCell('B1').value, SPILLING);
  assert.deepEqual(sheet?.getCell('D1').value, LEGACY);
});

test('the streaming reader reads the array formulas the buffered reader does', () => {
  const bytes = writeXlsx(workbookWith({B1: SPILLING, D1: LEGACY}));
  const streamed = [...readSheetRows(bytes)].flatMap((row) => row.cells);
  assert.deepEqual(
    streamed.map((cell) => [cell.address, cell.value]),
    [
      ['B1', SPILLING],
      ['D1', LEGACY],
    ],
  );
});

test('a one-cell range is spelled as the cell alone, as Excel spells it', () => {
  const workbook = workbookWith({F1: {shareType: 'array', formula: 'SUM(A1:A3*2)', ref: 'F1:F1'}});
  assert.match(sheetXml(writeXlsx(workbook)), /<f t="array" ref="F1">/);
  assert.deepEqual(roundtrip(workbook).getWorksheet('S')?.getCell('F1').value, {
    shareType: 'array',
    formula: 'SUM(A1:A3*2)',
    ref: 'F1',
  });
});

test('the writer refuses a range that does not start at the formula cell', () => {
  for (const ref of ['B1:C2', 'C:C', 'S!C1:C2', 'nonsense']) {
    refuses((workbook) => {
      const sheet = workbook.getWorksheet('S');
      assert.ok(sheet);
      sheet.getCell('C1').value = {shareType: 'array', formula: '1', ref};
    });
  }
});

test('an array formula whose range is missing or starts elsewhere reads as its plain formula', () => {
  const read = (f: string) =>
    readXlsx(
      foreignPackage({[SHEET1]: foreignSheet(`<row r="1"><c r="B1">${f}<v>6</v></c></row>`)}),
    )
      .getWorksheet('S')
      ?.getCell('B1').value;
  const plain = {formula: 'SUM(A1:A3)', result: 6};
  assert.deepEqual(read('<f t="array">SUM(A1:A3)</f>'), plain, 'no ref');
  assert.deepEqual(
    read('<f t="array" ref="A1:B2">SUM(A1:A3)</f>'),
    plain,
    'a range starting at A1',
  );
  assert.deepEqual(read('<f t="array" ref="junk">SUM(A1:A3)</f>'), plain, 'an unreadable range');
  assert.deepEqual(read('<f t="array" ref="$B$1:$B$3">SUM(A1:A3)</f>'), {
    shareType: 'array',
    formula: 'SUM(A1:A3)',
    ref: 'B1:B3',
    result: 6,
  });
});

test('a cm is a dynamic array only where the workbook cell metadata marks it', () => {
  const metadata = writeXlsxMetadata();
  const cell = (cm: string) =>
    readXlsx(
      withCellMetadata(
        `<row r="1"><c r="B1"${cm}><f t="array" ref="B1:B3">SEQUENCE(3)</f><v>1</v></c></row>`,
        metadata,
      ),
    )
      .getWorksheet('S')
      ?.getCell('B1').value;
  assert.deepEqual(cell(' cm="1"'), SPILLING);
  const legacy = {shareType: 'array', formula: 'SEQUENCE(3)', ref: 'B1:B3', result: 1};
  assert.deepEqual(cell(' cm="2"'), legacy, 'a block the part does not hold');
  assert.deepEqual(cell(''), legacy, 'no cm at all');
  const noPart = readXlsx(
    foreignPackage({
      [SHEET1]: foreignSheet(
        '<row r="1"><c r="B1" cm="1"><f t="array" ref="B1:B3">SEQUENCE(3)</f><v>1</v></c></row>',
      ),
    }),
  );
  assert.deepEqual(noPart.getWorksheet('S')?.getCell('B1').value, legacy, 'no cell metadata part');
});

// The cell metadata this library writes, which the reader must take as Excel's own.
function writeXlsxMetadata(): string {
  return partIn(partsWritten(workbookWith({B1: SPILLING})), 'xl/metadata.xml');
}

test('a dynamic array in a committed streamed row still reaches its cell metadata', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.addRow([null, SPILLING]).commit();
  sheet.commit();
  const bytes = await writer.commit();
  assert.ok(bytes);
  assertRelationshipsWired(bytes);
  assert.deepEqual(readXlsx(bytes).getWorksheet('S')?.getCell('B1').value, SPILLING);
});

test('a splice moves an array formula, its range and its references together', () => {
  const workbook = workbookWith({B2: {...SPILLING, formula: 'A2:A4*2', ref: 'B2:B4'}});
  const sheet = workbook.getWorksheet('S');
  assert.ok(sheet);
  sheet.spliceRows(1, 0, []);
  assert.deepEqual(sheet.getCell('B3').value, {...SPILLING, formula: 'A3:A5*2', ref: 'B3:B5'});
  sheet.spliceColumns(1, 0, []);
  assert.deepEqual(sheet.getCell('C3').value, {...SPILLING, formula: 'B3:B5*2', ref: 'C3:C5'});
  sheet.spliceRows(4, 1);
  assert.deepEqual(
    sheet.getCell('C3').value,
    {...SPILLING, formula: 'B3:B4*2', ref: 'C3:C4'},
    'a delete inside the range shrinks it',
  );
});

test('a duplicated row copies an array formula over a range of the same shape', () => {
  const workbook = workbookWith({A1: 1, B1: {...SPILLING, formula: 'A1*{1,2}', ref: 'B1:C1'}});
  const sheet = workbook.getWorksheet('S');
  assert.ok(sheet);
  sheet.duplicateRow(1);
  assert.deepEqual(sheet.getCell('B2').value, {
    shareType: 'array',
    formula: 'A2*{1,2}',
    ref: 'B2:C2',
    dynamic: true,
  });
});
