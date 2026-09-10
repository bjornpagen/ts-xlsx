import assert from 'node:assert/strict';
import {test} from 'node:test';

import {isDataTableFormulaValue} from '../../core/value.ts';
import {Workbook} from '../../core/workbook.ts';
import {
  elementIn,
  foreignPackage,
  foreignSheet,
  roundtrip,
  SHEET1,
  sheetXml,
} from './package.test-support.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

test('a data-table formula writes its t="dataTable" declaration with input cells', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('B2').value = {
    shareType: 'dataTable',
    ref: 'B2:B5',
    dataTableRow: true,
    r1: 'A1',
    result: 99,
  };

  const cell = elementIn(sheetXml(writeXlsx(wb)), /<c r="B2"[\s\S]*?<\/c>/);
  assert.match(cell, /<f t="dataTable"/, 'the formula is emitted as the data-table kind');
  assert.match(cell, /ref="B2:B5"/, 'the data-table range is emitted');
  assert.match(cell, /dtr="1"/, 'the row-input flag is emitted');
  assert.match(cell, /r1="A1"/, 'the input cell is emitted');
  assert.match(cell, /<v>99<\/v>/, 'the cached result travels with the cell');
});

test('a data-table formula round-trips its kind, range, inputs, and result', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('B2').value = {
    shareType: 'dataTable',
    ref: 'B2:B5',
    dataTableRow: true,
    r1: 'A1',
    result: 99,
  };

  const value = roundtrip(wb).getWorksheet('S')?.getCell('B2').value ?? null;
  assert.ok(isDataTableFormulaValue(value), 'the cell reads back as a data-table formula');
  assert.strictEqual(value.ref, 'B2:B5');
  assert.strictEqual(value.dataTableRow, true);
  assert.strictEqual(value.r1, 'A1');
  assert.strictEqual(value.result, 99);
});

test('a two-variable data table round-trips its 2-D flag and both input cells', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  // A 2-D What-If table drives a row input (r1) and a column input (r2) at once; dtr does not apply.
  sheet.getCell('B2').value = {
    shareType: 'dataTable',
    ref: 'B2:E5',
    dataTable2D: true,
    r1: 'A1',
    r2: 'A2',
    result: 99,
  };

  const value = roundtrip(wb).getWorksheet('S')?.getCell('B2').value ?? null;
  assert.ok(isDataTableFormulaValue(value), 'the cell reads back as a data-table formula');
  assert.strictEqual(value.dataTable2D, true, 'the 2-D kind survives');
  assert.strictEqual(value.r1, 'A1', 'the row input cell survives');
  assert.strictEqual(value.r2, 'A2', 'the column input cell survives');
  assert.strictEqual(value.dataTableRow, undefined, 'a 2-D table carries no row-orientation flag');
});

test('a column-input data table round-trips with no row-orientation flag', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  // A one-variable table down a column omits dataTableRow entirely (the writer emits dtr="0").
  sheet.getCell('B2').value = {
    shareType: 'dataTable',
    ref: 'B2:B5',
    r1: 'A1',
    result: 5,
  };

  const value = roundtrip(wb).getWorksheet('S')?.getCell('B2').value ?? null;
  assert.ok(isDataTableFormulaValue(value), 'the cell reads back as a data-table formula');
  assert.strictEqual(value.r1, 'A1', 'the input cell survives');
  assert.strictEqual(value.dataTableRow, undefined, 'no row-orientation flag is invented');
  assert.strictEqual(value.dataTable2D, undefined, 'a one-variable table is not marked 2-D');
});

test('a ref-less t="dataTable" declaration is tolerated, not read as a data table', () => {
  // The declaration is meaningless without its range; a hostile or corrupt sheet can still emit one.
  // The reader must not surface a data-table value from it: the cell decodes as its plain payload.
  const archive = foreignPackage({
    [SHEET1]: foreignSheet(
      '<row r="2"><c r="B2"><f t="dataTable" dt2D="0" dtr="1" r1="A1"/><v>7</v></c></row>',
    ),
  });

  const value = readXlsx(archive).getWorksheet('S')?.getCell('B2').value ?? null;
  assert.ok(
    !isDataTableFormulaValue(value),
    'the ref-less declaration does not become a data table',
  );
  assert.strictEqual(value, 7, 'the cell falls back to its plain cached payload');
});

test('a re-written data-table formula still declares t="dataTable" after a read-modify-write', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('B2').value = {
    shareType: 'dataTable',
    ref: 'B2:B5',
    dataTableRow: true,
    r1: 'A1',
    result: 99,
  };

  const reloaded = roundtrip(wb);
  const sheet2 = reloaded.getWorksheet('S');
  assert.ok(sheet2, 'the sheet reloads');
  sheet2.getCell('A1').value = 'edited elsewhere';
  assert.match(
    sheetXml(writeXlsx(reloaded)),
    /t="dataTable"/,
    'the kind survives an unrelated edit',
  );
});
