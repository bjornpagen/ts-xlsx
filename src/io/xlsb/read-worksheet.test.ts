import assert from 'node:assert/strict';
import {test} from 'node:test';

import {MAX_COLUMN} from '../../core/address.ts';
import {Workbook} from '../../core/workbook.ts';
import {parseWorksheet} from './read-worksheet.ts';
import {BRT} from './record-types.ts';
import {concat, frame, nameCitations, word} from './records.test-support.ts';

// `BrtColInfo` runs spanning the whole grid, hidden so the reader applies them: first and last column,
// a width, no style, and the hidden flag.
function fullWidthHiddenRuns(count: number): Uint8Array {
  const run = frame(
    BRT.ColInfo,
    concat(word(0), word(MAX_COLUMN - 1), word(2_560), word(0), Uint8Array.of(0b1, 0)),
  );
  return concat(...Array.from({length: count}, () => run));
}

test('a formula whose decoded text would outgrow any real formula keeps its cached value', () => {
  const cached = new Uint8Array(8);
  new DataView(cached.buffer).setFloat64(0, 7, true);
  const rgce = nameCitations(600);
  const part = concat(
    // `BrtRowHdr` for row 1, then `BrtFmlaNum` at A1: the cell, the cached number, two flag bytes,
    // the token stream and an empty extra-data block.
    frame(BRT.RowHdr, concat(word(0), word(0), Uint8Array.of(0, 0, 0, 0))),
    frame(
      BRT.FmlaNum,
      concat(word(0), word(0), cached, Uint8Array.of(0, 0), word(rgce.length), rgce, word(0)),
    ),
  );
  const sheet = new Workbook().addWorksheet('S');
  parseWorksheet(part, {
    sheet,
    sharedStrings: [],
    xfStyles: [],
    scope: {
      sheetNames: ['S'],
      externSheets: [],
      selfSupBook: undefined,
      names: ['x'.repeat(1 << 20)],
    },
    dateEpoch: 1900,
    definedNames: new Set(),
  });
  assert.equal(sheet.getCell('A1').value, 7);
});

// Clamping one run to the grid bounds that run and nothing else, so a part of full-width runs used to
// cost a column touch per run per column: 800 of them, 16 KB of part, took most of a second. Counted
// rather than timed, for the reason `core/merge-index.test.ts` gives.
test('full-width column runs are charged to one per-sheet budget, however many a part declares', () => {
  for (const count of [10, 800]) {
    const sheet = new Workbook().addWorksheet('S');
    let touched = 0;
    const getColumn = sheet.getColumn.bind(sheet);
    sheet.getColumn = (index) => {
      touched++;
      return getColumn(index);
    };
    parseWorksheet(fullWidthHiddenRuns(count), {
      sheet,
      sharedStrings: [],
      xfStyles: [],
      scope: {sheetNames: ['S'], externSheets: [], selfSupBook: undefined, names: []},
      dateEpoch: 1900,
      definedNames: new Set(),
    });
    assert.equal(touched, 4 * MAX_COLUMN, `${count} full-width runs touched ${touched} columns`);
    assert.equal(getColumn(MAX_COLUMN).hidden, true, 'and the runs the budget affords still apply');
  }
});
