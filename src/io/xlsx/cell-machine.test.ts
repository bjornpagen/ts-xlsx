// The two worksheet readers drive one `<c>` machine (`CellAccumulator`), each contributing only what
// committing a cell means to it. This is the claim that machine exists to make, and the one the
// comment it replaced could only ask for: read one package both ways and the cells agree.
//
// One difference is deliberate and is the reason this asserts on decoded values rather than
// deep-equality of the models: the row stream does not resolve a shared formula against its master.
// Everything else must match, a rich string's runs included.

import assert from 'node:assert/strict';
import {test} from 'node:test';

import {MAX_ROW} from '../../core/address.ts';
import {isFormulaValue} from '../../core/value.ts';
import {Workbook} from '../../core/workbook.ts';
import {partText, patchParts, SHEET1, sheetXml} from './package.test-support.ts';
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

// The streamer flattened an inline rich string to its text but kept a pooled one's runs, because the
// shared-string reader reads runs for both readers. Excel pools rich text, so the same runs streamed as
// `"bold and not"` or as `{richText: [...]}` depending on how the producer stored them.
test('a rich string reads with its runs both ways, whether inline or pooled', () => {
  // With `useSharedStrings` this writer pools plain text but still inlines rich text, so the pooled rich
  // string is made by replacing a pooled plain one's `<si>`, and the inline cells are authored beside it.
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'seed';
  const data = writeXlsx(wb, {useSharedStrings: true});
  const seed = '<si><t>seed</t></si>';
  assert.ok(
    partText(data, 'xl/sharedStrings.xml').includes(seed),
    'precondition: the seed is pooled',
  );

  const patched = patchParts(
    patchSheetBody(
      data,
      '<c r="A1" t="inlineStr"><is><t>plain &amp; simple</t></is></c>' +
        '<c r="B1" t="inlineStr"><is><r><rPr><b/></rPr><t>bold</t></r><r><t> and not</t></r></is></c>' +
        '<c r="C1" t="s"><v>0</v></c>',
    ),
    {
      'xl/sharedStrings.xml': (xml) =>
        xml.replace(seed, '<si><r><rPr><b/></rPr><t>pooled</t></r><r><t> run</t></r></si>'),
    },
  );

  const sheet = readXlsx(patched).getWorksheet('S')!;
  assert.deepEqual(sheet.getCell('B1').value, {
    richText: [{text: 'bold', font: {bold: true}}, {text: ' and not'}],
  });
  const streamed = [...readSheetRows(patched)][0]?.cells ?? [];
  assert.deepEqual(
    streamed.map((cell) => cell.address),
    ['A1', 'B1', 'C1'],
  );
  for (const cell of streamed) {
    assert.deepEqual(cell.value, sheet.getCell(cell.address).value, `${cell.address} agrees`);
  }
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

// Excel shows each of these as blank (Excel 16.0 build 20326), and a reader decoding by `t` alone read
// the boolean as FALSE and the rest as "", which a save then wrote back as data.
const TYPED_WITHOUT_VALUE: readonly (readonly [string, string])[] = [
  ['b', '<c r="A1" t="b"/>'],
  ['s', '<c r="A1" t="s"/>'],
  ['e', '<c r="A1" t="e"/>'],
  ['str', '<c r="A1" t="str"/>'],
  ['d', '<c r="A1" t="d"/>'],
  ['n', '<c r="A1" t="n"/>'],
  ['inlineStr', '<c r="A1" t="inlineStr"/>'],
  ['b, empty <v>', '<c r="A1" t="b"><v></v></c>'],
  ['b, <v/>', '<c r="A1" t="b"><v/></c>'],
  ['s, empty <v>', '<c r="A1" t="s"><v></v></c>'],
  ['s, blank <v>', '<c r="A1" t="s"><v> </v></c>'],
  ['e, empty <v>', '<c r="A1" t="e"><v></v></c>'],
  ['inlineStr, <v> and no <is>', '<c r="A1" t="inlineStr"><v>ignored</v></c>'],
];

test('a typed cell with no value reads as none, identically both ways', () => {
  for (const [label, cell] of TYPED_WITHOUT_VALUE) {
    const patched = patchSheetBody(seeded(), `${cell}<c r="B1"><v>1</v></c>`);
    assert.equal(readXlsx(patched).getWorksheet('S')!.getCell('A1').value, null, `${label}`);
    const streamed = [...readSheetRows(patched)][0]?.cells ?? [];
    assert.deepEqual(
      streamed.map((c) => c.address),
      ['B1'],
      `${label} is not a data cell`,
    );
  }
});

test('a string cell whose <v> or <is> is present but empty reads as the empty string both ways', () => {
  for (const cell of [
    '<c r="A1" t="str"><v></v></c>',
    // The self-closing spelling fires no close, and is still the same element as `<v></v>`.
    '<c r="A1" t="str"><v/></c>',
    '<c r="A1" t="inlineStr"><is/></c>',
    '<c r="A1" t="inlineStr"><is></is></c>',
  ]) {
    const patched = patchSheetBody(seeded(), cell);
    assert.equal(readXlsx(patched).getWorksheet('S')!.getCell('A1').value, '', cell);
    assert.deepEqual(
      ([...readSheetRows(patched)][0]?.cells ?? []).map((c) => [c.address, c.value]),
      [['A1', '']],
      cell,
    );
  }
});

// Each of these makes Excel refuse the package, so there is no reading to match, and none is invented.
test('a typed token with no reading for its type is no value, but an unlisted error code is kept', () => {
  const patched = patchSheetBody(
    seeded(),
    '<c r="A1" t="b"><v>2</v></c><c r="B1" t="s"><v>abc</v></c><c r="C1" t="s"><v>99</v></c>' +
      '<c r="D1" t="e"><v>#BUSY!</v></c>',
  );
  const sheet = readXlsx(patched).getWorksheet('S')!;
  assert.equal(sheet.getCell('A1').value, null, 'not FALSE');
  assert.equal(sheet.getCell('B1').value, null, 'not ""');
  assert.equal(sheet.getCell('C1').value, null, 'nor "" for an index past the pool');
  assert.equal(sheet.getCell('D1').value, '#BUSY!', 'a newer error code is data');
});

test('a typed cell with no value is not written back as one', () => {
  const body = TYPED_WITHOUT_VALUE.map(([, cell], i) =>
    cell.replace('A1', `${String.fromCharCode(65 + i)}1`),
  ).join('');
  const xml = sheetXml(writeXlsx(readXlsx(patchSheetBody(seeded(), body))));
  // The cell elements themselves may stay, as a bare `<c r="A1"/>` does; the invented values may not.
  assert.doesNotMatch(xml, /<c [^>]*\bt="|<v>|<is>/);
});

// A row past the grid is dropped by both readers, cells included: the streamer dropped them, and the
// buffered reader placed one whose own `r` was in the grid.
test('buffered and streamed reads agree that a row past the grid places no cell', () => {
  const book = new Workbook();
  book.addWorksheet('S').getCell('A1').value = 1;
  const data = patchParts(writeXlsx(book), {
    [SHEET1]: (xml) =>
      xml.replace(
        /<sheetData>[\s\S]*?<\/sheetData>/,
        `<sheetData><row r="${MAX_ROW + 1}"><c r="A5"><v>9</v></c></row></sheetData>`,
      ),
  });
  assert.equal(readXlsx(data).getWorksheet('S')?.getCell('A5').value, null, 'buffered');
  assert.deepEqual([...readSheetRows(data)], [], 'streamed');
});

// A one-sheet package whose pool holds one string, for a patched body to index into.
function seeded(): Uint8Array {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'seed';
  return writeXlsx(wb, {useSharedStrings: true});
}

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
