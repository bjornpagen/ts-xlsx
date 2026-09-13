import assert from 'node:assert/strict';
import {test} from 'node:test';

import type {CellValue} from '../../core/value.ts';
import {Workbook} from '../../core/workbook.ts';
import {patchParts, SHEET1} from './package.test-support.ts';
import {readXlsx} from './read.ts';
import {WorkbookStreamWriter} from './write-stream.ts';
import {writeXlsx} from './write.ts';

// Excel 16.0 offered to repair every package below the writer refuses, and opened the ones it writes
// clean (`test/corpus/fixtures/excel-oracle/array-formula-ranges.json`).
const legacy = (ref: string): CellValue => ({shareType: 'array', formula: 'A1:A3*2', ref});

function workbookWith(cells: Record<string, CellValue>): Workbook {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('S');
  for (const [address, value] of Object.entries(cells)) sheet.getCell(address).value = value;
  return workbook;
}

test('an array formula whose range holds another formula is refused, naming both cells', () => {
  for (const other of [{formula: 'A2*10'}, legacy('B2'), {sharedFormula: 'B1'}] as CellValue[]) {
    const workbook = workbookWith({B1: legacy('B1:B3'), B2: other});
    assert.throws(
      () => writeXlsx(workbook),
      /array formula in B1 fills "B1:B3", which holds the formula in B2/,
    );
  }
  const dynamic = workbookWith({
    B1: {shareType: 'array', formula: 'A1:A3*2', ref: 'B1:B3', dynamic: true},
    B3: {formula: 'A3'},
  });
  assert.throws(() => writeXlsx(dynamic), /which holds the formula in B3/, 'a dynamic array too');
});

test('array formula ranges sharing a cell are refused, whichever holds whose formula', () => {
  const corner = workbookWith({B1: legacy('B1:C2'), A2: legacy('A2:B3')});
  assert.throws(
    () => writeXlsx(corner),
    /array formula in A2 fills "A2:B3", which shares cells with the range of the array formula in B1/,
  );
});

test('an array formula whose other cells hold values, or nothing, is written', () => {
  assert.doesNotThrow(() => writeXlsx(workbookWith({B1: legacy('B1:B3'), B2: 4, B3: 6})));
  assert.doesNotThrow(() => writeXlsx(workbookWith({B1: legacy('B1:B3'), C1: legacy('C1:C3')})));
});

test('the streaming writer asks the same of rows it already flushed', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.addRow([1, legacy('B1:B3')]).commit();
  sheet.addRow([2, {formula: 'A2*10'}]).commit();
  await assert.rejects(writer.commit(), /which holds the formula in B2/);
});

// Excel 16.0 repaired each file below by keeping the array formula and taking the other formula, and the
// reader reads it the same way.
const kept = {shareType: 'array', formula: 'A1:A3*2', ref: 'B1:B3', result: 2} satisfies CellValue;

test('a formula inside a file array formula range loses its value, keeping its style', () => {
  const workbook = workbookWith({A1: 1, B1: kept, B2: 4});
  const authored = workbook.getWorksheet('S');
  assert.ok(authored);
  authored.getCell('B2').font = {bold: true};
  const foreign = patchParts(writeXlsx(workbook), {
    [SHEET1]: (xml) =>
      xml.replace(/<c r="B2"( s="\d+")><v>4<\/v><\/c>/, '<c r="B2"$1><f>A2*10</f><v>20</v></c>'),
  });
  const sheet = readXlsx(foreign).getWorksheet('S');
  assert.deepEqual(sheet?.getCell('B1').value, kept, 'the array formula keeps its range');
  assert.equal(sheet?.getCell('B2').value, null);
  assert.equal(sheet?.getCell('B2').font?.bold, true);
  assert.doesNotThrow(() => writeXlsx(readXlsx(foreign)), 'and the reading writes back');
});

test('a shared formula whose master a file array formula range holds keeps only its value', () => {
  const foreign = patchParts(writeXlsx(workbookWith({A1: 1, B1: kept, B2: 4, C2: 5})), {
    [SHEET1]: (xml) =>
      xml
        .replace(
          '<c r="B2"><v>4</v></c>',
          '<c r="B2"><f t="shared" ref="B2:C2" si="0">A2*10</f><v>20</v></c>',
        )
        .replace('<c r="C2"><v>5</v></c>', '<c r="C2"><f t="shared" si="0"/><v>200</v></c>'),
  });
  const sheet = readXlsx(foreign).getWorksheet('S');
  assert.equal(sheet?.getCell('B2').value, null);
  assert.equal(sheet?.getCell('C2').value, 200);
  assert.doesNotThrow(() => writeXlsx(readXlsx(foreign)));
});

test('of two array formula ranges sharing a cell in a file, the later keeps only its value', () => {
  const written = writeXlsx(workbookWith({B1: legacy('B1:C2'), D5: legacy('D5')}));
  const foreign = patchParts(written, {
    [SHEET1]: (xml) =>
      xml.replace(
        '<c r="D5">',
        '<c r="A2"><f t="array" ref="A2:B3">A2:A3*2</f><v>3</v></c><c r="D5">',
      ),
  });
  const sheet = readXlsx(foreign).getWorksheet('S');
  const first = sheet?.getCell('B1').value as {ref?: string} | undefined;
  assert.equal(first?.ref, 'B1:C2', 'the first keeps its range');
  assert.equal(sheet?.getCell('A2').value, 3);
  assert.doesNotThrow(() => writeXlsx(readXlsx(foreign)));
});
