import assert from 'node:assert/strict';
import {test} from 'node:test';

import {FUNCTION_VALUE_NAMES, NOT_VALUES, OLDER_FUNCTION_VALUES} from './function-values.ts';
import {FUTURE_FUNCTION_PREFIXES} from './future-functions.ts';

// The list is written as the functions older than the future-function registry plus that registry
// minus a few, so each half has to stay what it claims to be.
test('every exception is a future function, so none outlives the name it excepts', () => {
  for (const name of NOT_VALUES) assert.ok(FUTURE_FUNCTION_PREFIXES.has(name), name);
});

test('the older functions repeat no future function and no name of their own', () => {
  for (const name of OLDER_FUNCTION_VALUES) assert.ok(!FUTURE_FUNCTION_PREFIXES.has(name), name);
  assert.equal(new Set(OLDER_FUNCTION_VALUES).size, OLDER_FUNCTION_VALUES.length);
});

// Excel 16.0 (build 20326) wrote each of these bare in value position.
test('what Excel leaves bare is not a function value', () => {
  for (const name of [
    'LOG10',
    'TRUE',
    'FALSE',
    'GOTO',
    'GET.CELL',
    'NETWORKDAYS.INTL',
    'WORKDAY.INTL',
    'ISO.CEILING',
    'ECMA.CEILING',
    'LAMBDA',
    'LET',
    'QUERYSTRING',
  ]) {
    assert.ok(!FUNCTION_VALUE_NAMES.has(name), name);
  }
});

test('a function of either age is a value', () => {
  for (const name of ['SUM', 'LEN', 'T', 'N', 'IFERROR', 'EDATE', 'CONCAT', 'FILTER', 'T.TEST']) {
    assert.ok(FUNCTION_VALUE_NAMES.has(name), name);
  }
});
