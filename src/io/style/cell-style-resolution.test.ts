// The format a cell resolves to decides the type its value decodes to, so the order and the xf 0
// floor are tested here once, rather than only through the three readers that drive them.

import {strict as assert} from 'node:assert';
import {test} from 'node:test';

import {CellStyleResolver} from './cell-style-resolution.ts';
import type {XfStyle} from './xf-style.ts';

const XF: readonly XfStyle[] = [
  {numFmt: 'mm-dd-yy'},
  {numFmt: '0.00'},
  {numFmt: '0%'},
  {numFmt: '@'},
];

test("a cell's own format wins over its row's and its column's", () => {
  const resolver = new CellStyleResolver();
  resolver.noteColumnSpan(1, 1, 3);
  resolver.openRow(2, true);
  assert.equal(resolver.styleFor(1, 1, XF), XF[1]);
  assert.equal(resolver.declaresFormat(1, 1), true);
});

test("a row's format reaches its bare cells only when the row is marked customFormat", () => {
  const resolver = new CellStyleResolver();
  resolver.openRow(2, false);
  assert.equal(resolver.styleFor(1, -1, XF), XF[0], 'an unflagged row format describes the row');
  resolver.openRow(2, true);
  assert.equal(resolver.styleFor(1, -1, XF), XF[2]);
});

test("a column's format is the last one declared, reached after the row's", () => {
  const resolver = new CellStyleResolver();
  resolver.noteColumnSpan(2, 4, 3);
  resolver.openRow(-1, false);
  assert.equal(resolver.styleFor(3, -1, XF), XF[3]);
  assert.equal(resolver.styleFor(5, -1, XF), XF[0], 'a column outside the span declares nothing');
  resolver.openRow(2, true);
  assert.equal(resolver.styleFor(3, -1, XF), XF[2], 'a flagged row comes before the column');
});

test('with nothing declared a cell resolves to xf 0, and says nothing was declared', () => {
  const resolver = new CellStyleResolver();
  assert.equal(resolver.styleFor(1, -1, XF), XF[0]);
  assert.equal(resolver.declaresFormat(1, -1), false);
});

test('closing a row stops its format reaching a cell outside it', () => {
  const resolver = new CellStyleResolver();
  resolver.openRow(2, true);
  resolver.closeRow();
  assert.equal(resolver.styleFor(1, -1, XF), XF[0]);
});

test('an index past the style table resolves to no style rather than to xf 0', () => {
  const resolver = new CellStyleResolver();
  assert.equal(resolver.styleFor(1, 9, XF), undefined);
  assert.equal(
    resolver.styleFor(1, -1, []),
    undefined,
    'and an empty table has no xf 0 to fall to',
  );
});
