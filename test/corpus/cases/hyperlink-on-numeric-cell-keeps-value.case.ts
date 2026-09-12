// Cluster: xlsx-io
//
// Real-world scenario: OOXML stores a hyperlink beside the cell, in a sheet-level `<hyperlinks>` list,
// so Excel lets a link sit on any cell: a total, a date, a formula. A reader that folds the link into
// the cell's value as a text label has nothing to label with when the value is not text: it either
// destroys the number, its type and any formula, or drops the link. Neither is necessary, because a
// link is not part of the value at all.

import type {Assert, Case, CorpusApi} from '../case.ts';

const REFS = ['A1', 'B1', 'C1', 'D1', 'E1', 'F1'];

export default {
  id: 'hyperlink-on-numeric-cell-keeps-value',
  provenance: {source: 'audit'},
  cluster: 'xlsx-io',
  description:
    'A hyperlink over a number, a boolean, a formula, a date, a text label or an empty cell is read, ' +
    'and the value beneath it keeps its type, before and after a save.',

  behavior: [
    {
      name: 'a number, a boolean, a formula, a date and a label keep their values under a link',
      expect(api: CorpusApi, assert: Assert) {
        const {cells} = api.hyperlinkOverNonTextCellsReport().read;
        assert.deepEqual(cells.A1, {kind: 'number', value: 42});
        assert.deepEqual(cells.B1, {kind: 'boolean', value: true});
        assert.deepEqual(cells.C1, {kind: 'formula', value: {formula: '1+1', result: 2}});
        assert.deepEqual(cells.D1, {kind: 'date', value: '2024-01-15T00:00:00.000Z'});
        assert.deepEqual(cells.E1, {kind: 'string', value: 'label'});
        assert.deepEqual(cells.F1, {kind: 'null', value: null});
      },
    },
    {
      name: 'every one of the six links is read, whatever the cell beneath it holds',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(
          api.hyperlinkOverNonTextCellsReport().read.links,
          REFS.map((ref) => ({ref, target: '#S!H1'})),
        );
      },
    },
    {
      name: 'a save keeps every value and every link it read',
      expect(api: CorpusApi, assert: Assert) {
        const {read, rewritten} = api.hyperlinkOverNonTextCellsReport();
        assert.deepEqual(rewritten, read);
        assert.deepEqual(
          rewritten.cells.A1,
          {kind: 'number', value: 42},
          'pinned, not only self-consistent',
        );
      },
    },
  ],
} satisfies Case;
