import assert from 'node:assert/strict';
import {test} from 'node:test';

import {AuthoringError} from '../../errors.ts';
import {admitting, repairedSheetNames} from './read-repair.ts';

test('a workbook of sheet names is repaired in order, each against the names before it', () => {
  const repaired = repairedSheetNames([
    {name: 'S', part: 1},
    {name: 's', part: 2},
    {name: 'x'.repeat(40), part: 3},
    {name: '', part: 4},
  ]);
  assert.deepEqual(repaired, [
    {name: 'S', part: 1},
    {name: 's (2)', part: 2},
    {name: 'x'.repeat(31), part: 3},
    {name: 'Sheet1', part: 4},
  ]);
});

test('admitting swallows a refusal about the input and nothing else', () => {
  assert.equal(
    admitting((): number => {
      throw new AuthoringError('refused');
    }),
    undefined,
  );
  assert.equal(
    admitting(() => 7),
    7,
  );
  assert.throws(
    () =>
      admitting(() => {
        throw new TypeError('a bug');
      }),
    TypeError,
  );
});
