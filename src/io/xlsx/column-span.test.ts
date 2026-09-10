import assert from 'node:assert/strict';
import {test} from 'node:test';

import {ColumnRecordBudget} from '../read-policy/column-budget.ts';
import {takeColumnSpan} from './column-span.ts';

test('a col span is clamped to the grid, and one starting past it reaches nothing', () => {
  const budget = new ColumnRecordBudget();
  assert.deepEqual(takeColumnSpan({min: '3', max: '99999999'}, budget), {first: 3, last: 16_384});
  assert.equal(takeColumnSpan({min: '16385', max: '16400'}, budget), undefined);
  assert.equal(
    takeColumnSpan({max: '4'}, budget),
    undefined,
    'and unreadable bounds reach nothing',
  );
});
