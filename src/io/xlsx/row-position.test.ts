import assert from 'node:assert/strict';
import {test} from 'node:test';

import {RowPositionTracker} from './row-position.ts';

test('a row with no r takes the position after the last one', () => {
  const rows = new RowPositionTracker();
  assert.deepEqual(rows.open({}), {number: 1, inGrid: true});
  assert.deepEqual(rows.open({}), {number: 2, inGrid: true});
});

// The two spellings are mixable, and a declared number is the file speaking: the rows after it
// continue from where it said, not from where counting had got to.
test('a declared r re-anchors the count for the rows after it', () => {
  const rows = new RowPositionTracker();
  rows.open({});
  assert.deepEqual(rows.open({r: '10'}), {number: 10, inGrid: true});
  assert.deepEqual(rows.open({}), {number: 11, inGrid: true});
});

test('an unreadable r is an absent one: the row still gets its position', () => {
  const rows = new RowPositionTracker();
  rows.open({r: '4'});
  assert.deepEqual(rows.open({r: 'junk'}), {number: 5, inGrid: true});
});

// Reported rather than clamped: an `r` names one row, so there is nothing to fold it onto, and both
// readers drop such a row whole.
test('a row past the grid is reported out of grid, not clamped onto the last one', () => {
  const rows = new RowPositionTracker();
  assert.deepEqual(rows.open({r: '1048576'}), {number: 1_048_576, inGrid: true});
  assert.deepEqual(rows.open({r: '1048577'}), {number: 1_048_577, inGrid: false});
});
