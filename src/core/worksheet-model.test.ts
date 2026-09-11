// Field-agnostic cover for the model round-trip. The per-field cases live in worksheet.test.ts;
// these assert the properties the facet registry is what guarantees, so they keep their meaning,
// and keep catching a half-wired field, as WorksheetModel grows.

import assert from 'node:assert/strict';
import {test} from 'node:test';

import {AuthoringError} from '../errors.ts';
import {WORKSHEET_MODEL_FACETS} from './worksheet-model.ts';
import {Worksheet, type WorksheetModel} from './worksheet.ts';

// Every field of a WorksheetModel, populated, so a round-trip has something to lose in each of them.
function populatedSheet(): Worksheet {
  const sheet = new Worksheet('Src', 1);
  sheet.state = 'hidden';
  sheet.tabColor = {argb: 'FFFF0000'};
  sheet.properties.defaultRowHeight = 18;
  sheet.outline.summaryBelow = false;
  sheet.freeze(1, 1);
  sheet.pageSetup.fitToPage = true;
  sheet.printOptions.gridLines = true;
  sheet.pageMargins.left = 0.25;
  sheet.headerFooter.oddHeader = '&Ctitle';
  sheet.rowBreaks.push({id: 3});
  sheet.columnBreaks.push({id: 2});
  sheet.getColumn(2).width = 20;
  sheet.getRow(7).height = 33; // a row that carries formatting and no cells
  sheet.getCell('A1').value = 'title';
  sheet.getCell('B2').value = 42;
  sheet.getCell('B2').font = {bold: true};
  sheet.mergeCells('D1:E1');
  sheet.addDataValidation('F1', {type: 'whole', operator: 'between', formulae: [0, 9]});
  sheet.addConditionalFormatting({
    ref: 'G1:G3',
    rules: [{type: 'cellIs', operator: 'greaterThan', formulae: [5], priority: 1}],
  });
  sheet.addTable({name: 'Totals', ref: 'A10', columns: [{name: 'Col'}], rowCount: 2});
  sheet.autoFilter = 'A1:B2';
  sheet.protect('secret');
  return sheet;
}

test('the exported model carries exactly the fields the facet registry declares', () => {
  // The registry is proved exhaustive over `keyof WorksheetModel` at compile time; this is the
  // runtime half: that the getter actually emits a key for each facet and invents none besides.
  assert.deepEqual(
    Object.keys(populatedSheet().model).sort(),
    WORKSHEET_MODEL_FACETS.map((facet) => facet.key).sort(),
  );
});

test('a fully populated sheet round-trips to an identical model', () => {
  const src = populatedSheet();
  const dst = new Worksheet('Dst', 2);
  dst.model = src.model;

  // A facet whose two directions disagree (a field read but not written, or written into the wrong
  // place) shows up here as a difference, whichever field it is.
  assert.deepEqual(dst.model, src.model);
});

test('a round-trip through a sheet holding unrelated content leaves no residue', () => {
  const dst = new Worksheet('Dst', 2);
  dst.getCell('Z9').value = 'stale';
  dst.getColumn(26).width = 99;
  dst.getRow(9).height = 99;
  dst.mergeCells('X1:Y1');
  dst.addDataValidation('W1', {type: 'whole', formulae: [1]});
  dst.protect('other');

  const src = populatedSheet();
  dst.model = src.model;

  assert.deepEqual(dst.model, src.model);
});

// Assigning a model resets the sheet and replays each field, and a field that threw part-way used to
// leave the destination emptied and half loaded: its own content gone, the model's earlier fields
// applied, the rest not.
test('a model that cannot be applied throws and leaves the destination exactly as it was', () => {
  const extent = (sheet: Worksheet): number[] | undefined => {
    const range = sheet.usedRange;
    return range && [range.top, range.left, range.bottom, range.right];
  };
  const valid = populatedSheet().model;
  const broken: [string, WorksheetModel][] = [
    ['overlapping merges', {...valid, merges: ['A1:B2', 'B2:C3']}],
    [
      'a data validation naming no cells',
      {...valid, dataValidations: [{sqref: '', rule: {type: 'whole', formulae: [1]}}]},
    ],
  ];

  for (const [what, model] of broken) {
    const dst = new Worksheet('Dst', 2);
    dst.getCell('A1').value = 'precious';
    dst.getCell('Z9').value = 'far';
    dst.mergeCells('X1:Y1');
    const before = dst.model;
    const extentBefore = extent(dst);

    assert.throws(() => {
      dst.model = model;
    }, AuthoringError);

    assert.deepEqual(dst.model, before, `${what}: the destination's content is untouched`);
    assert.deepEqual(extent(dst), extentBefore, `${what}: and so is its extent`);
  }
});

test('the exported model orders cells, rows, and columns ascending whatever order they were written', () => {
  const sheet = new Worksheet('S', 1);
  for (const reference of ['C3', 'A1', 'B2', 'A3']) sheet.getCell(reference).value = reference;
  for (const index of [5, 1, 3]) sheet.getColumn(index).width = index;
  for (const number of [6, 2, 4]) sheet.getRow(number).height = number;

  const model = sheet.model;
  assert.deepEqual(
    model.cells.map((cell) => [cell.row, cell.col]),
    [
      [1, 1],
      [2, 2],
      [3, 1],
      [3, 3],
    ],
    'cells come out row-major, ascending within each row',
  );
  assert.deepEqual(
    model.columns.map((column) => column.index),
    [1, 3, 5],
  );
  assert.deepEqual(
    model.rows.map((row) => row.number),
    [2, 4, 6],
  );
});

test('a model exported from an empty sheet clears every field of the sheet it is assigned to', () => {
  const dst = populatedSheet();
  const empty = new Worksheet('Empty', 3).model;

  dst.model = empty;

  assert.deepEqual(dst.model, empty);
});

test('a frozen pane survives a model transplant', () => {
  // The named regression behind the facet: `view` was omitted from the model for no stated reason,
  // so a transplant reproduced every other sheet-level field and silently unfroze the header row.
  // It is workbook-independent state, the test the model's boundary is drawn on, not an attached
  // part, so it belongs in the snapshot exactly as the autofilter does.
  const src = new Worksheet('Src', 1);
  src.freeze(2, 1);

  const dst = new Worksheet('Dst', 2);
  dst.model = src.model;

  assert.deepEqual(dst.view, {state: 'frozen', xSplit: 1, ySplit: 2, topLeftCell: 'B3'});
});

test('a model transplant clears a frozen pane the destination held', () => {
  const dst = new Worksheet('Dst', 2);
  dst.freeze(1);

  dst.model = new Worksheet('Src', 1).model;

  assert.deepEqual(dst.view, {}, 'a normal-view source leaves no pane behind');
});
