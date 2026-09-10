import assert from 'node:assert/strict';
import {test} from 'node:test';

import {ERROR_CODES, isErrorValue} from '../../core/value.ts';
import {Workbook} from '../../core/workbook.ts';
import {patchParts, roundtrip, SHEET1, sheetXml} from './package.test-support.ts';
import {readXlsx} from './read.ts';
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

test('a foreign t="e" cell carrying a non-canonical code reads back as a plain string', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {error: '#REF!'};

  const patched = patchParts(writeXlsx(wb), {
    [SHEET1]: (xml) => xml.replace('<v>#REF!</v>', '<v>#UNKNOWN!</v>'),
  });

  const back = readXlsx(patched).getWorksheet('S')?.getCell('A1').value;
  assert.equal(back, '#UNKNOWN!', 'an unrecognised error literal falls back to its raw text');
});
