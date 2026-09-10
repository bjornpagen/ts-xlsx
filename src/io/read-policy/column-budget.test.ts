import assert from 'node:assert/strict';
import {test} from 'node:test';

import {clampColumnSpan, ColumnRecordBudget} from './column-budget.ts';

test('a column run is clamped to the grid, and one starting past it reaches nothing', () => {
  const budget = new ColumnRecordBudget();
  assert.deepEqual(clampColumnSpan(3, 99_999_999, budget), {first: 3, last: 16_384});
  assert.equal(clampColumnSpan(16_385, 16_400, budget), undefined);
  assert.equal(clampColumnSpan(5, 4, budget), undefined, 'and an empty run reaches nothing');
});

// The budget is per sheet and spent by every run, so a file made of full-grid runs stops being
// applied rather than costing a column touch per record for as long as it cares to write them.
test('a column run past the sheet budget reaches nothing', () => {
  const budget = new ColumnRecordBudget();
  for (let run = 0; run < 4; run++) {
    assert.notEqual(clampColumnSpan(1, 16_384, budget), undefined);
  }
  assert.equal(clampColumnSpan(1, 16_384, budget), undefined);
});

test('a run the budget can only partly afford is truncated, not dropped', () => {
  const budget = new ColumnRecordBudget();
  for (let run = 0; run < 3; run++) clampColumnSpan(1, 16_384, budget);
  clampColumnSpan(1, 16_000, budget);
  assert.deepEqual(clampColumnSpan(1, 16_384, budget), {first: 1, last: 384});
});
