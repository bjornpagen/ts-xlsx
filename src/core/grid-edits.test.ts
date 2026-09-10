// A splice moves everything anchored to the grid, not only the cells. The four participants that live
// outside the cell grid (data validations, conditional formats, comment threads, the autofilter) are
// the ones a splice used to leave behind, pointing a dropdown or a highlight rule at whatever moved
// into their place. These lock the re-anchoring on both axes.

import assert from 'node:assert/strict';
import {test} from 'node:test';

import {threadAt} from './comment-thread.test-support.ts';
import type {CellValue} from './value.ts';
import {Workbook} from './workbook.ts';
import type {Worksheet} from './worksheet.ts';

// A sheet carrying one of each range-bound overlay over the same block, so a single splice exercises
// all four and any that fails to move stands out beside the three that did.
function anchoredSheet(): Worksheet {
  const sheet = new Workbook().addWorksheet('S');
  sheet.getCell('B5').value = 'x';
  sheet.addDataValidation('B5:B6', {type: 'list', formulae: ['"a,b"']});
  sheet.addConditionalFormatting({ref: 'B5:B6', rules: [{type: 'dataBar', priority: 1}]});
  sheet.autoFilter = 'B5:C6';
  sheet.addCommentThread(threadAt('B5'));
  return sheet;
}

const anchors = (sheet: Worksheet) => ({
  validation: sheet.dataValidations.map((entry) => entry.sqref),
  formatting: sheet.conditionalFormattings.map((entry) => entry.ref),
  filter: sheet.autoFilter?.ref,
  threads: sheet.commentThreads.map((thread) => thread.ref),
});

test('an insert above carries every range-bound overlay down with the cells it covers', () => {
  const sheet = anchoredSheet();
  sheet.insertRow(1, ['hdr']);

  assert.equal(sheet.getCell('B6').value, 'x');
  assert.deepEqual(anchors(sheet), {
    validation: ['B6:B7'],
    formatting: ['B6:B7'],
    filter: 'B6:C7',
    threads: ['B6'],
  });
});

test('a delete above carries every range-bound overlay up with the cells it covers', () => {
  const sheet = anchoredSheet();
  sheet.spliceRows(1, 2);

  assert.equal(sheet.getCell('B3').value, 'x');
  assert.deepEqual(anchors(sheet), {
    validation: ['B3:B4'],
    formatting: ['B3:B4'],
    filter: 'B3:C4',
    threads: ['B3'],
  });
});

test('an overlay whose every anchor row is deleted goes with the rows', () => {
  const sheet = anchoredSheet();
  sheet.spliceRows(5, 2);

  assert.deepEqual(anchors(sheet), {
    validation: [],
    formatting: [],
    filter: undefined,
    threads: [],
  });
});

test('an insert left carries every range-bound overlay right with the cells it covers', () => {
  const sheet = anchoredSheet();
  sheet.insertColumn(1, ['hdr']);

  assert.equal(sheet.getCell('C5').value, 'x');
  assert.deepEqual(anchors(sheet), {
    validation: ['C5:C6'],
    formatting: ['C5:C6'],
    filter: 'C5:D6',
    threads: ['C5'],
  });
});

test('a delete left carries every range-bound overlay left with the cells it covers', () => {
  const sheet = anchoredSheet();
  sheet.spliceColumns(1, 1);

  assert.equal(sheet.getCell('A5').value, 'x');
  assert.deepEqual(anchors(sheet), {
    validation: ['A5:A6'],
    formatting: ['A5:A6'],
    filter: 'A5:B6',
    threads: ['A5'],
  });
});

test('an overlay whose every anchor column is deleted goes with the columns', () => {
  const sheet = anchoredSheet();
  sheet.spliceColumns(2, 2);

  assert.deepEqual(anchors(sheet), {
    validation: [],
    formatting: [],
    filter: undefined,
    threads: [],
  });
});

test('a splice below an overlay leaves it alone', () => {
  const sheet = anchoredSheet();
  sheet.insertRow(9, ['later']);

  assert.deepEqual(anchors(sheet), {
    validation: ['B5:B6'],
    formatting: ['B5:B6'],
    filter: 'B5:C6',
    threads: ['B5'],
  });
});

test('a whole-column sqref survives a row splice with its spelling intact', () => {
  const sheet = new Workbook().addWorksheet('S');
  sheet.addDataValidation('B:B', {type: 'list', formulae: ['"a,b"']});
  sheet.addConditionalFormatting({ref: 'D:D', rules: [{type: 'dataBar', priority: 1}]});
  sheet.insertRow(1, ['hdr']);

  // Not `B2:B1048577`: a column covers every row, so a row splice cannot move it, and re-spelling it
  // would rewrite a foreign file's own wording for no gain.
  assert.equal(sheet.dataValidations[0]?.sqref, 'B:B');
  assert.equal(sheet.conditionalFormattings[0]?.ref, 'D:D');
});

test('a whole-row sqref shifts through a row splice and stays row-shaped', () => {
  const sheet = new Workbook().addWorksheet('S');
  sheet.addDataValidation('5:5', {type: 'list', formulae: ['"a,b"']});
  sheet.insertRow(1, ['hdr']);

  assert.equal(sheet.dataValidations[0]?.sqref, '6:6');
});

test('each area of a multi-area sqref shifts on its own', () => {
  const sheet = new Workbook().addWorksheet('S');
  sheet.addDataValidation('A1:C1 A3:C3 A9:C9', {type: 'list', formulae: ['"a,b"']});
  // Deletes row 3 outright and pulls row 9 up by one; row 1 is above the cut and stays put.
  sheet.spliceRows(3, 1);

  assert.equal(sheet.dataValidations[0]?.sqref, 'A1:C1 A8:C8');
});

test('a single-cell sqref shifts without growing into a range', () => {
  const sheet = new Workbook().addWorksheet('S');
  sheet.addDataValidation('B5', {type: 'list', formulae: ['"a,b"']});
  sheet.insertRow(1, ['hdr']);

  assert.equal(sheet.dataValidations[0]?.sqref, 'B6');
});

test('a shifted validation answers a point lookup from its new geometry', () => {
  const sheet = anchoredSheet();
  sheet.insertRow(1, ['hdr']);

  assert.equal(sheet.dataValidationAt('B5'), undefined);
  assert.equal(sheet.dataValidationAt('B6')?.type, 'list');
});

test("a thread and its cell's note stay on the same cell through a splice", () => {
  const sheet = new Workbook().addWorksheet('S');
  sheet.getCell('B5').note = 'legacy note';
  sheet.addCommentThread(threadAt('B5'));
  sheet.insertRow(1, ['hdr']);

  // A note is cell state and travels with the cell; the thread anchored to the same cell has to make
  // the same move, or the writer emits a conversation on one cell and its fallback note on another.
  assert.equal(sheet.getCell('B6').note, 'legacy note');
  assert.equal(sheet.commentThreads[0]?.ref, 'B6');
});

test('a column splice inside a filter re-measures its criteria against the new left edge', () => {
  const sheet = new Workbook().addWorksheet('S');
  sheet.autoFilter = {
    ref: 'B1:E10',
    columns: [
      {colId: 0, criteria: {kind: 'values', values: ['keep'], blank: false}},
      {colId: 2, criteria: {kind: 'values', values: ['shift'], blank: false}},
    ],
  };
  // Column C is the filter's colId 1; deleting it narrows the range and pulls colId 2 down to 1.
  sheet.spliceColumns(3, 1);

  assert.equal(sheet.autoFilter?.ref, 'B1:D10');
  assert.deepEqual(
    sheet.autoFilter?.columns.map((column) => column.colId),
    [0, 1],
  );
});

test('a criterion on a deleted column goes with the column', () => {
  const sheet = new Workbook().addWorksheet('S');
  sheet.autoFilter = {
    ref: 'B1:E10',
    columns: [
      {colId: 1, criteria: {kind: 'values', values: ['doomed'], blank: false}},
      {colId: 3, criteria: {kind: 'values', values: ['survivor'], blank: false}},
    ],
  };
  sheet.spliceColumns(3, 1);

  assert.deepEqual(sheet.autoFilter?.columns, [
    {colId: 2, criteria: {kind: 'values', values: ['survivor'], blank: false}},
  ]);
});

test('a row splice leaves a filter criterion addressed as it was', () => {
  const sheet = new Workbook().addWorksheet('S');
  sheet.autoFilter = {
    ref: 'B5:E10',
    columns: [{colId: 2, criteria: {kind: 'values', values: ['keep'], blank: false}}],
  };
  sheet.insertRow(1, ['hdr']);

  assert.equal(sheet.autoFilter?.ref, 'B6:E11');
  assert.deepEqual(
    sheet.autoFilter?.columns.map((column) => column.colId),
    [2],
  );
});

// ── The grid's edges ────────────────────────────────────────────────────────────────────────────

// A whole column is written as a bounded range to the last row, which is what Excel itself writes,
// so every one of these regions is already sitting on the grid's edge before the splice touches it.
// Pushing an edge past that produced a package Excel met with its repair prompt; the same workbook
// with the edges inside the grid opened clean. See
// docs/knowledge/specs/a-splice-must-not-push-geometry-off-the-grid.md.
const LAST_ROW = 1_048_576;
const LAST_COLUMN = 16_384;

test('an insert above a whole-column validation leaves its bottom edge on the last row', () => {
  const sheet = new Workbook().addWorksheet('S');
  sheet.addDataValidation(`B1:B${LAST_ROW}`, {type: 'list', formulae: ['"a,b,c"']});
  sheet.spliceRows(900, 0, ['inserted'], ['also']);

  assert.equal(
    sheet.dataValidations[0]?.sqref,
    `B1:B${LAST_ROW}`,
    "the sqref Excel's own row insert produces for the same rule",
  );
});

test('an insert above a full-height autofilter and conditional format clamps both', () => {
  const sheet = new Workbook().addWorksheet('S');
  sheet.autoFilter = `A1:A${LAST_ROW}`;
  sheet.addConditionalFormatting({
    ref: `C1:C${LAST_ROW}`,
    rules: [{type: 'dataBar', priority: 1}],
  });
  sheet.insertRow(1, ['header']);

  // Each region's top edge moves with the inserted row, as it should; what may not move is the
  // bottom edge, which has nowhere to go.
  assert.equal(sheet.autoFilter?.ref, `A2:A${LAST_ROW}`);
  assert.equal(sheet.conditionalFormattings[0]?.ref, `C2:C${LAST_ROW}`);
});

test('a merge on the bottom edge shrinks rather than naming a row past the grid', () => {
  // The one place the clamp is visibly lossy: the merge loses the row it had no room to move into.
  // Excel refuses the insert outright here, since a merge is content occupying the last rows, but a
  // refusal from a library that has already accepted the splice for everything else is worse than a
  // region one row shorter, and the alternative measured was a file that will not open.
  const sheet = new Workbook().addWorksheet('S');
  sheet.mergeCells(`C${LAST_ROW - 1}:D${LAST_ROW}`);
  sheet.insertRow(1, ['header']);

  assert.deepEqual(sheet.merges, [`C${LAST_ROW}:D${LAST_ROW}`]);
});

test('a column splice holds the right edge at the last column, the same rule on the other axis', () => {
  const sheet = new Workbook().addWorksheet('S');
  sheet.addDataValidation('A1:XFD1', {type: 'list', formulae: ['"a,b"']});
  sheet.autoFilter = 'A2:XFD2';
  sheet.spliceColumns(2, 0, ['inserted']);

  assert.equal(sheet.dataValidations[0]?.sqref, 'A1:XFD1');
  assert.equal(
    sheet.autoFilter?.ref,
    'A2:XFD2',
    'the insert is inside it: only its right edge could move, and it cannot',
  );
});

test('a splice with content on the last line completes instead of dying inside the grid', () => {
  // The cell grid was the one splice participant doing raw arithmetic where every other one went
  // through `shiftIndex`. `new Cell` asserts its coordinates, so an insert on a sheet holding a cell
  // in the last row or column threw a `RangeError` from inside the splice, and the column axis wrote
  // each row back as it went, so the throw left half the sheet shifted and the other half not.
  const rows = new Workbook().addWorksheet('R');
  rows.getCell('A1').value = 'first';
  rows.getCell(`A${LAST_ROW}`).value = 'last';
  rows.insertRow(1, ['inserted']);
  assert.equal(rows.getCell('A1').value, 'inserted');
  assert.equal(rows.getCell('A2').value, 'first');

  const columns = new Workbook().addWorksheet('C');
  columns.getCell('A1').value = 'a1';
  columns.getCell('XFD2').value = 'edge';
  columns.getCell('A3').value = 'a3';
  columns.spliceColumns(1, 0, ['inserted']);

  assert.equal(columns.getCell('A1').value, 'inserted', 'the insert landed');
  assert.equal(columns.getCell('B1').value, 'a1', 'row 1 shifted');
  assert.equal(columns.getCell('B3').value, 'a3', 'and so did row 3, which the throw used to skip');
});

test('a splice with line properties on the last line keeps them inside the grid too', () => {
  // The cell grid was fixed to clamp; its row and column *metadata* was the last participant still
  // doing the arithmetic by hand, so an insert pushed a height off the bottom to row 1048577. That
  // made `rowCount` name a row `new Row` refuses to construct, so iterating the sheet threw and the
  // sheet could no longer be written or inspected -- after a legal public call.
  const rows = new Workbook().addWorksheet('R');
  rows.getRow(LAST_ROW).height = 20;
  rows.insertRow(1, ['inserted']);
  assert.equal(rows.rowCount, LAST_ROW, 'the height clamped onto the last row, as its cells do');
  assert.doesNotThrow(() => [...rows.rows()]);
  assert.equal(rows.getRow(LAST_ROW).height, 20);

  const columns = new Workbook().addWorksheet('C');
  columns.getColumn(LAST_COLUMN).width = 12;
  columns.spliceColumns(1, 0, ['inserted']);
  assert.equal(columns.columnCount, LAST_COLUMN);
  assert.equal(columns.getColumn(LAST_COLUMN).width, 12);
});

test('a splice drops the line properties of a line it deleted', () => {
  const sheet = new Workbook().addWorksheet('S');
  sheet.getRow(2).height = 30;
  sheet.getRow(3).height = 40;
  sheet.spliceRows(2, 1);
  assert.equal(sheet.getRow(2).height, 40, 'row 3 and its height shifted up into row 2');
  assert.equal(sheet.getRow(3).height, undefined, 'nothing was left behind at the old index');
});

test('a column splice that deletes a table whole drops the table rather than re-pointing it', () => {
  // `shiftRows` asks `isDeletedSpan` and returns false so the caller prunes; `shiftColumns` did not,
  // so a table survived a splice that deleted its every column, still carrying the names of columns
  // that no longer exist, declared over whatever slid left. The writer then emits that.
  const sheet = new Workbook().addWorksheet('S');
  sheet.getCell('B1').value = 'h1';
  sheet.getCell('C1').value = 'h2';
  sheet.getCell('B2').value = 1;
  sheet.getCell('C2').value = 2;
  sheet.addTable({name: 'T', ref: 'B1', columns: [{name: 'h1'}, {name: 'h2'}], rowCount: 1});

  sheet.spliceColumns(2, 2);
  assert.deepEqual(sheet.tables, [], 'the table had no column left to occupy');
});

test('an insert with nowhere to put a column is refused whole, not clamped', () => {
  // The other half of the rule the regions above follow, and the reason it is the other half rather
  // than an inconsistency. A *region* pushed off the edge clamps, because a validation one column
  // narrower is a legible loss in a file that opens. *Content* pushed off the edge is what Excel
  // refuses outright ("can't insert new cells because it would push non-empty cells off the end of
  // the worksheet"), and so does this: clamping two inserted columns onto XFD would stack the second
  // on top of the first and call that success. See
  // docs/knowledge/specs/a-splice-must-not-push-geometry-off-the-grid.md.
  const sheet = new Workbook().addWorksheet('S');
  sheet.getCell('A1').value = 'a1';

  assert.throws(
    () => sheet.spliceColumns(LAST_COLUMN, 0, ['first'], ['second']),
    (error: unknown) =>
      error instanceof RangeError && /column 16385 is out of bounds/.test(error.message),
  );

  // Refused, not half-applied: the grid is swapped in after the inserts are built, so a throw leaves
  // the caller a sheet they can still act on.
  assert.equal(sheet.getCell('A1').value, 'a1');
  assert.equal(sheet.columnCount, 1);
});

test('an inserted column reaching past the last row is refused, the way addColumn refuses it', () => {
  // Same rule on the other axis of the same argument: an inserted column's values are indexed by row,
  // so a long enough array names a row that cannot exist. `addColumn` has always answered this with a
  // `RangeError` from `new Cell`; the splice path answers identically, which is what keeps the two
  // column-append doors from disagreeing about the same argument.
  const sheet = new Workbook().addWorksheet('S');
  sheet.getCell('A1').value = 'a1';
  // Sparse on purpose: only the last element exists, so this names row 1048577 without allocating a
  // million values to get there.
  const values: CellValue[] = [];
  values[LAST_ROW] = 'past the bottom';

  assert.throws(() => sheet.insertColumn(1, values), RangeError);
  assert.throws(() => sheet.addColumn(values), RangeError);

  assert.equal(sheet.getCell('A1').value, 'a1');
  assert.equal(sheet.columnCount, 1);
});

test('an insert of empty lines past the last one is refused before anything moves', () => {
  // An empty inserted row builds no cell, so the `Cell` constructor that refuses content past the
  // edge never saw it: the splice left a row key past the grid and the sheet threw from then on.
  const rows = new Workbook().addWorksheet('R');
  rows.getCell('A1').value = 'a1';
  assert.throws(() => rows.spliceRows(LAST_ROW, 0, [], []), {
    name: 'RangeError',
    message: `row ${LAST_ROW + 1} is out of bounds: Excel supports 1..${LAST_ROW}`,
  });
  assert.equal(rows.rowCount, 1);
  assert.doesNotThrow(() => [...rows.rows()]);

  // The column axis used to report success here, for columns that cannot exist.
  const columns = new Workbook().addWorksheet('C');
  columns.getCell('A1').value = 'a1';
  assert.throws(() => columns.spliceColumns(LAST_COLUMN, 0, [], []), RangeError);
  assert.equal(columns.columnCount, 1);
});

test('duplicating the last row has nowhere to put the copy, in either mode', () => {
  for (const insert of [true, false]) {
    const sheet = new Workbook().addWorksheet('S');
    sheet.getRow(LAST_ROW).height = 20;
    assert.throws(() => sheet.duplicateRow(LAST_ROW, {insert}), RangeError, `insert: ${insert}`);
    assert.equal(sheet.rowCount, LAST_ROW, `insert: ${insert}`);
    assert.doesNotThrow(() => [...sheet.rows()], `insert: ${insert}`);
  }
});

test('a replacing duplicate that would run past the last row writes none of its copies', () => {
  const sheet = new Workbook().addWorksheet('S');
  sheet.getCell(`A${LAST_ROW - 1}`).value = 'source';
  assert.throws(() => sheet.duplicateRow(LAST_ROW - 1, {count: 3, insert: false}), RangeError);
  assert.equal(sheet.hasCell(LAST_ROW, 1), false, 'not even the copy that would have fitted');
  assert.doesNotThrow(() => [...sheet.rows()]);
});

test('a table on the right edge keeps its anchor inside the grid', () => {
  // The unclamped increment could put the anchor past the last column, where `range`, `autoFilterRef`
  // and `region` all throw on read: a sheet that cannot be serialised or even inspected.
  const sheet = new Workbook().addWorksheet('S');
  sheet.getCell('XFD1').value = 'h';
  sheet.getCell('XFD2').value = 1;
  sheet.addTable({name: 'T', ref: 'XFD1', columns: [{name: 'h'}], rowCount: 1});

  sheet.spliceColumns(1, 0, ['inserted']);
  assert.doesNotThrow(() => sheet.tables[0]?.range);
});
