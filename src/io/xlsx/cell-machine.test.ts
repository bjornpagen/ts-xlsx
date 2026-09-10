// The two worksheet readers drive one `<c>` machine (`CellAccumulator`), each contributing only what
// committing a cell means to it. This is the claim that machine exists to make, and the one the
// comment it replaced could only ask for: read one package both ways and the cells agree.
//
// Two differences are deliberate and are the reason this asserts on decoded values rather than
// deep-equality of the models: the row stream flattens a rich inline string to its text, and it does
// not resolve a shared formula against its master. Everything else must match.

import assert from 'node:assert/strict';
import {test} from 'node:test';

import {isFormulaValue} from '../../core/value.ts';
import {Workbook} from '../../core/workbook.ts';
import {patchParts, SHEET1, sheetXml} from './package.test-support.ts';
import {readSheetRows} from './read-rows.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

// Every payload shape the machine has a branch for: a shared string, an inline one, a number, a
// boolean, an error, a date, a formula with a cached result, and a formatted-but-empty `<c/>`.
function sample(): Uint8Array {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'shared text';
  sheet.getCell('B1').value = 42.5;
  sheet.getCell('C1').value = true;
  sheet.getCell('D1').value = {error: '#DIV/0!'};
  sheet.getCell('E1').value = new Date(Date.UTC(2020, 4, 17));
  sheet.getCell('F1').value = {formula: 'B1*2', result: 85};
  sheet.getCell('G1').value = '';
  sheet.getCell('H1').numFmt = '0.00';
  sheet.getCell('A2').value = 'text with an & and a <bracket>';
  sheet.getCell('B2').value = 'escaped\u0001control';
  return writeXlsx(wb);
}

test('a package reads the same cell values buffered and streamed', () => {
  const data = sample();
  const sheet = readXlsx(data).getWorksheet('S')!;

  const seen: string[] = [];
  for (const row of readSheetRows(data)) {
    for (const streamed of row.cells) {
      seen.push(streamed.address);
      const buffered = sheet.getCell(streamed.address).value;
      const expected = isFormulaValue(buffered) ? buffered.result : buffered;
      const actual = isFormulaValue(streamed.value) ? streamed.value.result : streamed.value;
      assert.deepEqual(actual, expected, `${streamed.address} disagrees between the two readers`);
    }
  }
  // H1 carries a number format and no value, and a data read yields only cells that carry something.
  assert.deepEqual(seen, ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'A2', 'B2']);
});

test('an inline string reads the same both ways, flattened on the streaming side', () => {
  // Authored as raw XML: the writer pools strings, and `<is>` is the branch this is about.
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'seed';
  const data = writeXlsx(wb);

  const inline =
    '<c r="A1" t="inlineStr"><is><t>plain &amp; simple</t></is></c>' +
    '<c r="B1" t="inlineStr"><is><r><t>bold</t></r><r><t> and not</t></r></is></c>';
  const patched = patchSheetBody(data, inline);

  const sheet = readXlsx(patched).getWorksheet('S')!;
  assert.equal(sheet.getCell('A1').value, 'plain & simple');
  assert.deepEqual(sheet.getCell('B1').value, {richText: [{text: 'bold'}, {text: ' and not'}]});

  const streamed = [...readSheetRows(patched)][0]?.cells ?? [];
  assert.equal(streamed[0]?.value, 'plain & simple', 'a plain inline string is identical');
  assert.equal(
    streamed[1]?.value,
    'bold and not',
    'and a rich one flattens, which is the contract',
  );
});

test('a <v> that is not a number reads as no value, identically both ways', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'seed';
  const patched = patchSheetBody(
    writeXlsx(wb),
    '<c r="A1" s="0"><v>abc</v></c>' +
      '<c r="B1"><v></v></c>' +
      '<c r="C1"><v>Infinity</v></c>' +
      '<c r="D1"><v>12.5</v></c>',
  );

  const sheet = readXlsx(patched).getWorksheet('S')!;
  assert.equal(sheet.getCell('A1').value, null, 'unparseable text is not NaN');
  assert.equal(sheet.getCell('B1').value, null, 'and a blank <v> is not zero');
  assert.equal(sheet.getCell('C1').value, null, 'nor is a non-finite spelling a number');
  assert.equal(sheet.getCell('D1').value, 12.5, 'a real number is untouched');

  // The two readers agree on "no value" and spell it differently by contract: a data read yields
  // only cells that carry something, so the three that decode to null drop out of the stream.
  const streamed = [...readSheetRows(patched)][0]?.cells ?? [];
  assert.deepEqual(
    streamed.map((cell) => [cell.address, cell.value]),
    [['D1', 12.5]],
  );
});

test('a malformed numeric cell survives a re-write as a cell, carrying no value', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').numFmt = '0.00';
  const patched = patchSheetBody(writeXlsx(wb), '<c r="A1" s="1"><v>abc</v></c>');

  const rewritten = writeXlsx(readXlsx(patched));
  const xml = sheetXml(rewritten);
  assert.match(xml, /<c r="A1" s="\d+"\/>/, 'the cell and its style are kept');
  assert.doesNotMatch(xml, /NaN/);
});

test('a formula whose cached <v> is unparseable keeps its formula and caches nothing', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: 'B1*2', result: 85};
  const patched = patchSheetBody(writeXlsx(wb), '<c r="A1"><f>B1*2</f><v>abc</v></c>');

  const value = readXlsx(patched).getWorksheet('S')!.getCell('A1').value;
  assert.deepEqual(value, {formula: 'B1*2'}, 'no `result` key, not a NaN one');

  const streamed = [...readSheetRows(patched)][0]?.cells ?? [];
  assert.deepEqual(streamed[0]?.value, {formula: 'B1*2'}, 'and the two readers agree');
});

// Replace the whole `<sheetData>` body of the first worksheet part with one authored row.
function patchSheetBody(data: Uint8Array, cells: string): Uint8Array {
  // Round-tripping through the reader would re-pool the strings, so the bytes are edited directly.
  return patchParts(data, {
    [SHEET1]: (xml) =>
      xml.replace(
        /<sheetData>[\s\S]*?<\/sheetData>/,
        `<sheetData><row r="1">${cells}</row></sheetData>`,
      ),
  });
}
