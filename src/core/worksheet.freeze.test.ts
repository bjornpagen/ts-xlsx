import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Worksheet} from './worksheet.ts';

test('a new sheet has a normal (empty) view', () => {
  const sheet = new Worksheet('S', 1);
  assert.deepEqual(sheet.view, {});
});

test('freeze(1) pins one header row and anchors the scrolling pane below it', () => {
  const sheet = new Worksheet('S', 1);
  sheet.freeze(1);
  assert.strictEqual(sheet.view.state, 'frozen');
  assert.strictEqual(sheet.view.ySplit, 1);
  assert.strictEqual(sheet.view.xSplit, 0);
  assert.strictEqual(sheet.view.topLeftCell, 'A2');
});

test('freeze(0, 2) pins two columns and no rows', () => {
  const sheet = new Worksheet('S', 1);
  sheet.freeze(0, 2);
  assert.strictEqual(sheet.view.state, 'frozen');
  assert.strictEqual(sheet.view.xSplit, 2);
  assert.strictEqual(sheet.view.topLeftCell, 'C1');
});

test('a freeze leaving nothing to scroll is refused and the previous view stands', () => {
  const sheet = new Worksheet('S', 1);
  sheet.freeze(2, 1);
  const before = {...sheet.view};
  assert.throws(() => sheet.freeze(1_048_576), RangeError);
  assert.throws(() => sheet.freeze(0, 16_384), RangeError);
  assert.deepEqual(sheet.view, before);
});

test('freeze(0, 0) is the same as unfreezing', () => {
  const sheet = new Worksheet('S', 1);
  sheet.freeze(2, 2);
  sheet.freeze(0, 0);
  assert.strictEqual(sheet.view.state, 'normal');
  assert.strictEqual(sheet.view.ySplit, undefined);
});

test('unfreeze clears the split and leaves a normal view', () => {
  const sheet = new Worksheet('S', 1);
  sheet.freeze(1, 1);
  sheet.unfreeze();
  assert.strictEqual(sheet.view.state, 'normal');
  assert.strictEqual(sheet.view.xSplit, undefined);
  assert.strictEqual(sheet.view.topLeftCell, undefined);
});

test('freeze rejects negative or fractional splits', () => {
  const sheet = new Worksheet('S', 1);
  assert.throws(() => sheet.freeze(-1), /non-negative integers/);
  assert.throws(() => sheet.freeze(1.5), /non-negative integers/);
});
