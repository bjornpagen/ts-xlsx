import assert from 'node:assert/strict';
import {test} from 'node:test';

import type {Fill} from '../../core/style.ts';
import {isSharedFormulaValue, type SharedFormulaValue} from '../../core/value.ts';
import {Workbook} from '../../core/workbook.ts';
import {roundtrip, sheetXml} from './package.test-support.ts';
import {writeXlsx} from './write.ts';

// A master formula filled down a column: B1 is the master, B2/B3 are clones referencing it.
function filledColumn(): Workbook {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 1;
  sheet.getCell('A2').value = 2;
  sheet.getCell('A3').value = 3;
  sheet.getCell('B1').value = {formula: 'A1*2', result: 2};
  sheet.getCell('B2').value = {sharedFormula: 'B1', result: 4};
  sheet.getCell('B3').value = {sharedFormula: 'B1', result: 6};
  return wb;
}

function sharedOf(wb: Workbook, ref: string): SharedFormulaValue {
  const value = wb.getWorksheet('S')?.getCell(ref).value;
  assert.ok(
    value !== undefined && value !== null && isSharedFormulaValue(value),
    `${ref} is a shared formula`,
  );
  return value;
}

test('a shared-formula clone reads back its master formula translated to its own address', () => {
  const back = roundtrip(filledColumn());

  const master = back.getWorksheet('S')?.getCell('B1').value as {formula: string; result: number};
  assert.equal(master.formula, 'A1*2', 'the master keeps its own formula');
  assert.equal(master.result, 2);

  const b2 = sharedOf(back, 'B2');
  assert.equal(b2.formula, 'A2*2', 'a clone one row down resolves to A2*2, not empty');
  assert.equal(b2.sharedFormula, 'B1', 'the clone records its master reference');
  assert.equal(b2.result, 4, 'the clone keeps its cached result');
  assert.equal(sharedOf(back, 'B3').formula, 'A3*2', 'two rows down is A3*2');
});

test('the master seeds the group with t="shared" ref/si and clones reference it by si', () => {
  const xml = sheetXml(writeXlsx(filledColumn()));
  assert.match(xml, /<c r="B1"><f t="shared" ref="B1:B3" si="0">A1\*2<\/f><v>2<\/v><\/c>/);
  assert.match(xml, /<c r="B2"><f t="shared" si="0"\/><v>4<\/v><\/c>/);
  assert.match(xml, /<c r="B3"><f t="shared" si="0"\/><v>6<\/v><\/c>/);
});

test('a shared formula filled across a row translates the column, not the row', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = {formula: 'A2+1', result: 0};
  sheet.getCell('B1').value = {sharedFormula: 'A1', result: 0};
  sheet.getCell('C1').value = {sharedFormula: 'A1', result: 0};

  const back = roundtrip(wb);
  assert.equal(sharedOf(back, 'B1').formula, 'B2+1', 'one column across');
  assert.equal(sharedOf(back, 'C1').formula, 'C2+1', 'two columns across');
});

test('an absolute reference in the master is not shifted in a clone', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('B1').value = {formula: '$A$1+A1', result: 0};
  sheet.getCell('B2').value = {sharedFormula: 'B1', result: 0};

  assert.equal(sharedOf(roundtrip(wb), 'B2').formula, '$A$1+A2', 'the anchored term stays put');
});

test('a shared formula survives a read → write → read round-trip', () => {
  const once = roundtrip(filledColumn());
  const twice = roundtrip(once);

  const b2 = sharedOf(twice, 'B2');
  assert.equal(b2.formula, 'A2*2', 'the clone still resolves after a second round-trip');
  assert.equal(b2.sharedFormula, 'B1');
  assert.equal(b2.result, 4);
  // The re-write reconstructed the shared grouping rather than expanding to concrete formulas.
  assert.match(sheetXml(writeXlsx(once)), /<f t="shared" si="0"\/>/);
});

test('a clone whose master has no formula is refused, naming the offending cell', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('B2').value = {sharedFormula: 'A1', result: 0};

  assert.throws(
    () => writeXlsx(wb),
    (error: Error) => {
      assert.match(error.message, /master/i, 'the error explains the missing master');
      assert.match(error.message, /B2/, 'the error names the offending clone');
      return true;
    },
  );
});

test('a clone above or left of its master is rejected', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('B2').value = {formula: 'A1', result: 0};
  sheet.getCell('B1').value = {sharedFormula: 'B2', result: 0}; // above the master

  assert.throws(() => writeXlsx(wb), /above and\/or left/);
});

test('inserting a column into a shared-formula sheet re-anchors the master so the write succeeds', () => {
  const wb = roundtrip(filledColumn());
  const sheet = wb.getWorksheet('S');
  assert.ok(sheet !== undefined);

  // The whole group shifts one column right; without re-anchoring the clones would still point at the
  // now-empty B1 and the writer would reject them as orphaned masters.
  sheet.spliceColumns(1, 0, []);
  assert.equal(sharedOf(wb, 'C2').sharedFormula, 'C1', 'the clone tracks its master to column C');
  assert.doesNotThrow(() => writeXlsx(wb));

  // The grouping survives the write and re-reads intact at its new position.
  const back = roundtrip(wb);
  const c2 = sharedOf(back, 'C2');
  assert.equal(c2.sharedFormula, 'C1');
  assert.match(sheetXml(writeXlsx(wb)), /<f t="shared" ref="C1:C3" si="0">/);
});

test('a shared formula a splice splits writes the part still shared as a group, and reads back as it was left', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  // B1 reads A10, B2 A11, B3 A12. A row inserted at 12 moves only B3's reference, so B2 is still B1
  // translated and B3 is not.
  sheet.getCell('B1').value = {formula: 'A10', result: 0};
  sheet.getCell('B2').value = {sharedFormula: 'B1', result: 0};
  sheet.getCell('B3').value = {sharedFormula: 'B1', result: 0};
  sheet.spliceRows(12, 0, []);

  const xml = sheetXml(writeXlsx(wb));
  assert.match(xml, /<c r="B1"><f t="shared" ref="B1:B2" si="0">A10<\/f>/);
  assert.match(xml, /<c r="B3"><f>A13<\/f>/);
  const back = roundtrip(wb);
  assert.equal(sharedOf(back, 'B2').formula, 'A11');
  assert.deepEqual(back.getWorksheet('S')?.getCell('B3').value, {formula: 'A13', result: 0});
});

test('a styled shared-formula clone keeps its fill and font on read, not just its value', () => {
  const red: Fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF0000'}};
  const wb = filledColumn();
  const sheet = wb.getWorksheet('S');
  assert.ok(sheet !== undefined);
  // The clone B2 carries style facets the classic reader dropped: it committed the clone's value
  // directly and never re-applied the resolved xf, so the fill/font vanished on round-trip.
  sheet.getCell('B2').fill = red;
  sheet.getCell('B2').font = {bold: true};

  const back = roundtrip(wb);
  const b2 = back.getWorksheet('S')?.getCell('B2');
  assert.ok(b2 !== undefined);
  assert.equal(
    sharedOf(back, 'B2').sharedFormula,
    'B1',
    'the clone is still a shared-formula clone',
  );
  // A solid fill reads back with the default indexed bgColor the writer emits alongside it.
  assert.deepEqual(b2.fill, {...red, bgColor: {indexed: 64}}, 'the clone keeps its fill');
  assert.deepEqual(b2.font, {bold: true}, 'the clone keeps its font');
});
