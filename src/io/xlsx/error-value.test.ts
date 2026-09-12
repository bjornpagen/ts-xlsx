import assert from 'node:assert/strict';
import {test} from 'node:test';

import {encodeAddress} from '../../core/address.ts';
import {ERROR_CODES, isErrorValue} from '../../core/value.ts';
import {Workbook} from '../../core/workbook.ts';
import {
  assertRelationshipsWired,
  optionalPartIn,
  partIn,
  partsOf,
  partsWritten,
  patchParts,
  roundtrip,
  SHEET1,
  sheetXml,
} from './package.test-support.ts';
import {readSheetRows} from './read-rows.ts';
import {readXlsx} from './read.ts';
import {WorkbookStreamWriter} from './write-stream.ts';
import {writeXlsx} from './write.ts';

function reReadA1(wb: Workbook): unknown {
  return roundtrip(wb).getWorksheet('S')?.getCell('A1').value;
}

test('an error cell round-trips as its error value', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {error: '#REF!'};

  const back = reReadA1(wb);
  assert.ok(isErrorValue(back as never), 'reads back as an error value');
  assert.deepEqual(back, {error: '#REF!'});
});

test('every canonical error code round-trips', () => {
  for (const code of ERROR_CODES) {
    const wb = new Workbook();
    wb.addWorksheet('S').getCell('A1').value = {error: code};
    assert.deepEqual(reReadA1(wb), {error: code}, `${code} survives`);
  }
});

test('an error cell serialises under t="e" with the code as its value', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {error: '#DIV/0!'};

  assert.match(sheetXml(writeXlsx(wb)), /<c r="A1" t="e"><v>#DIV\/0!<\/v><\/c>/);
});

test('a formula whose cached result is an error round-trips with both', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: 'A2/A3', result: {error: '#DIV/0!'}};

  const xml = sheetXml(writeXlsx(wb));
  assert.match(xml, /<c r="A1" t="e"><f>A2\/A3<\/f><v>#DIV\/0!<\/v><\/c>/);

  const back = reReadA1(wb) as {formula: string; result: unknown};
  assert.equal(back.formula, 'A2/A3', 'the formula survives');
  assert.deepEqual(back.result, {error: '#DIV/0!'}, 'the cached error result survives');
});

test('a styled error cell keeps its style across the round-trip', () => {
  const wb = new Workbook();
  const cell = wb.addWorksheet('S').getCell('A1');
  cell.value = {error: '#N/A'};
  cell.font = {bold: true};

  const back = roundtrip(wb).getWorksheet('S')?.getCell('A1');
  assert.deepEqual(back?.value, {error: '#N/A'}, 'the error value survives');
  assert.equal(back?.font?.bold, true, 'the font survives alongside it');
});

test('a foreign t="e" cell carrying a code the model does not list reads back as a plain string', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {error: '#REF!'};

  const patched = patchParts(writeXlsx(wb), {
    [SHEET1]: (xml) => xml.replace('<v>#REF!</v>', '<v>#PYTHON!</v>'),
  });

  const back = readXlsx(patched).getWorksheet('S')?.getCell('A1').value;
  assert.equal(back, '#PYTHON!', 'an unrecognised error literal falls back to its raw text');
});

test('a code the model does not list is refused rather than written', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {error: '#PYTHON!'} as never;
  assert.throws(() => writeXlsx(wb), {name: 'AuthoringError'});
});

test('an error Excel has no literal for is written as #VALUE! beside the rich value naming it', () => {
  // Excel 16.0 (build 20326) opened a cell `<c t="e"><v>…</v></c>` clean for `#N/A`, `#GETTING_DATA` and
  // `#BUSY!`, and with its repair prompt for `#SPILL!`, `#CALC!`, `#FIELD!`, `#BLOCKED!`, `#CONNECT!` and
  // `#UNKNOWN!`, each of which it stores as `#VALUE!` with a `vm`.
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = {error: '#SPILL!'};
  sheet.getCell('B1').value = {error: '#N/A'};
  sheet.getCell('C1').value = {formula: 'NA()', result: {error: '#CALC!'}};
  sheet.getCell('D1').value = {error: '#SPILL!'};

  const bytes = writeXlsx(wb);
  const parts = partsOf(bytes);
  const xml = partIn(parts, SHEET1);
  assert.match(xml, /<c r="A1" t="e" vm="1"><v>#VALUE!<\/v><\/c>/);
  assert.match(xml, /<c r="B1" t="e"><v>#N\/A<\/v><\/c>/, 'a literal points nowhere');
  assert.match(xml, /<c r="C1" t="e" vm="2"><f>NA\(\)<\/f><v>#VALUE!<\/v><\/c>/);
  assert.match(xml, /<c r="D1" t="e" vm="1">/, 'the same error shares its rich value');
  assert.match(partIn(parts, 'xl/richData/rdrichvalue.xml'), /<rv s="0"><v>8<\/v><v>0<\/v><\/rv>/);
  assert.match(partIn(parts, '[Content_Types].xml'), /rdrichvaluestructure\+xml/);
  assertRelationshipsWired(bytes);
});

test('every error round-trips, the ones stored as rich values included', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  ERROR_CODES.forEach((code, i) => {
    sheet.getCell(encodeAddress(i + 1, 1)).value = {error: code};
    sheet.getCell(encodeAddress(i + 1, 2)).value = {formula: 'NA()', result: {error: code}};
  });
  const bytes = writeXlsx(wb);
  const back = readXlsx(bytes).getWorksheet('S');
  const streamed = [...readSheetRows(bytes)];
  ERROR_CODES.forEach((code, i) => {
    assert.deepEqual(
      back?.getCell(encodeAddress(i + 1, 1)).value,
      {error: code},
      `${code} as a value`,
    );
    assert.deepEqual(
      back?.getCell(encodeAddress(i + 1, 2)).value,
      {formula: 'NA()', result: {error: code}},
      `${code} as a cached result`,
    );
    assert.deepEqual(streamed[0]?.cells[i]?.value, {error: code}, `${code} streamed`);
  });
});

test('a workbook with no rich-value error writes no rich-value part', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {error: '#DIV/0!'};
  const parts = partsWritten(wb);
  assert.equal(optionalPartIn(parts, 'xl/metadata.xml'), undefined);
  assert.equal(optionalPartIn(parts, 'xl/richData/rdrichvalue.xml'), undefined);
});

test('the streaming writer shares the rich values across the rows it flushes', async () => {
  const writer = new WorkbookStreamWriter();
  const sheet = writer.addWorksheet('S');
  sheet.addRow([{error: '#CALC!'}]).commit();
  sheet.addRow([{error: '#FIELD!'}, {error: '#CALC!'}]).commit();
  const bytes = await writer.commit();
  assert.ok(bytes !== undefined);
  const back = readXlsx(bytes).getWorksheet('S');
  assert.deepEqual(
    [back?.getCell('A1').value, back?.getCell('A2').value, back?.getCell('B2').value],
    [{error: '#CALC!'}, {error: '#FIELD!'}, {error: '#CALC!'}],
  );
  assert.match(sheetXml(bytes), /<c r="B2" t="e" vm="1">/);
});

test('a vm naming a rich-value error wins over the #VALUE! beside it', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {error: '#SPILL!'};
  const cell = '<c r="A1" t="e" vm="1"><v>#VALUE!</v></c>';
  const patched = patchParts(writeXlsx(wb), {
    [SHEET1]: (xml) =>
      xml.replace(
        cell,
        `${cell}<c r="B1" t="e"><v>#VALUE!</v></c><c r="C1" t="e" vm="7"><v>#VALUE!</v></c>`,
      ),
  });
  const back = readXlsx(patched).getWorksheet('S');
  assert.deepEqual(back?.getCell('A1').value, {error: '#SPILL!'});
  assert.deepEqual(back?.getCell('B1').value, {error: '#VALUE!'}, 'no vm');
  assert.deepEqual(back?.getCell('C1').value, {error: '#VALUE!'}, 'a vm naming no block');
});
