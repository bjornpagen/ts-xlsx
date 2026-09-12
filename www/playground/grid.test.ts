import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Workbook} from '../../src/index.ts';
import {DEFAULT_BOUNDS, toGrid} from './grid.ts';
import {findSample} from './samples.ts';

function sheetOf(id: string, name: string) {
  const sample = findSample(id);
  assert.ok(sample !== undefined, id);
  return sample.build().requireWorksheet(name);
}

test('an empty sheet is a grid with nothing in it, not a crash', () => {
  const sheet = new Workbook().addWorksheet('Empty');
  const grid = toGrid(sheet);
  assert.deepEqual(grid.rows, []);
  assert.deepEqual(grid.columns, []);
  assert.equal(grid.totalRows, 0);
  assert.equal(grid.hiddenRows, 0);
});

test('the grid bounds its own output and says how much it left out', () => {
  // The bound lives in the model on purpose: one applied while painting is a bound that has
  // already walked four thousand rows.
  const sheet = sheetOf('scale', 'Scale');
  const grid = toGrid(sheet, {maxRows: 10, maxColumns: 2});
  assert.equal(grid.rows.length, 10);
  assert.equal(grid.columns.length, 2);
  assert.equal(grid.totalRows, 4001);
  assert.equal(grid.hiddenRows, 3991);
  assert.equal(grid.totalColumns, 4);
  assert.equal(grid.hiddenColumns, 2);
});

test('a sheet smaller than the bound hides nothing', () => {
  const grid = toGrid(sheetOf('values', 'Values'), DEFAULT_BOUNDS);
  assert.equal(grid.hiddenRows, 0);
  assert.equal(grid.hiddenColumns, 0);
  assert.equal(grid.rows.length, grid.totalRows);
});

test('column headers are the letters of the columns actually shown', () => {
  const grid = toGrid(sheetOf('structure', 'Structure'), {maxRows: 3, maxColumns: 3});
  assert.deepEqual(grid.columns, ['A', 'B', 'C']);
});

test('a merge is reported once, on its anchor, and the cells it covers are absent', () => {
  const grid = toGrid(sheetOf('structure', 'Structure'), DEFAULT_BOUNDS);
  const [firstRow] = grid.rows;
  assert.ok(firstRow !== undefined);
  assert.equal(firstRow.cells.length, 1, 'the merged title occupies the whole first row');
  const [anchor] = firstRow.cells;
  assert.ok(anchor !== undefined);
  assert.equal(anchor.address, 'A1');
  assert.deepEqual(anchor.merge, {rowSpan: 1, colSpan: 3});
});

test('a cell carries its number format as data rather than as applied formatting', () => {
  // The library models formats and does not implement a formatting engine, so the grid
  // shows the value's text and hands the format to the page as a fact about the cell.
  const grid = toGrid(sheetOf('corpus-numfmt', 'NumFmt'), DEFAULT_BOUNDS);
  const cell = grid.rows
    .flatMap((row) => row.cells)
    .find((candidate) => candidate.address === 'B2');
  assert.ok(cell !== undefined);
  assert.equal(cell.text, '1234.5');
  assert.equal(cell.numFmt, '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)');
});

test('a cell that was never written reads as empty rather than materialising one', () => {
  const sheet = new Workbook().addWorksheet('Sparse');
  sheet.getCell('A1').value = 'here';
  sheet.getCell('C3').value = 'there';
  const before = sheet.actualRowCount;
  const grid = toGrid(sheet);
  const middle = grid.rows
    .flatMap((row) => row.cells)
    .find((candidate) => candidate.address === 'B2');
  assert.ok(middle !== undefined);
  assert.equal(middle.text, '');
  assert.equal(middle.type, 'null');
  assert.equal(sheet.actualRowCount, before, 'walking the grid did not grow the sheet');
});

test('the type of each cell is the kind the library reports, not one the grid guesses', () => {
  const grid = toGrid(sheetOf('values', 'Values'), DEFAULT_BOUNDS);
  const byAddress = new Map(
    grid.rows.flatMap((row) => row.cells).map((cell) => [cell.address, cell]),
  );
  assert.equal(byAddress.get('B2')?.type, 'number');
  assert.equal(byAddress.get('B3')?.type, 'string');
  assert.equal(byAddress.get('B4')?.type, 'boolean');
  assert.equal(byAddress.get('B5')?.type, 'date');
  assert.equal(byAddress.get('B6')?.type, 'formula');
  assert.equal(byAddress.get('B7')?.type, 'error');
  // The sample's link sits on the sheet over B8; the cell itself holds its label.
  assert.equal(byAddress.get('B8')?.type, 'string');
  assert.equal(byAddress.get('B9')?.type, 'richText');
});
