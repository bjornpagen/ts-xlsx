// Cluster: xlsx-io
//
// Real-world scenario: OOXML stores a hyperlink beside the cell, in a sheet-level `<hyperlinks>` list,
// so Excel lets a link sit on any cell: a total, a date, a formula. A reader that folds the link into
// the cell's value as a text label has nothing to label with when the value is not text, and replacing
// `42` with an empty label destroys the number, its type and any formula, after which a save writes an
// empty string where the value was. Losing the link is recoverable; losing the value is not.

import type {Assert, Case, CorpusApi} from '../case.ts';

const LINK = {hyperlink: '#S!H1'};

export default {
  id: 'hyperlink-on-numeric-cell-keeps-value',
  provenance: {source: 'audit'},
  cluster: 'xlsx-io',
  description:
    'A hyperlink over a number, a boolean, a formula or a date leaves that value and its type intact, ' +
    'before and after a save, while a hyperlink over a text label or an empty cell still reads as a ' +
    'hyperlink carrying that label.',

  behavior: [
    {
      name: 'a number, a boolean, a formula and a date keep their values under a link',
      expect(api: CorpusApi, assert: Assert) {
        const {read} = api.hyperlinkOverNonTextCellsReport();
        assert.deepEqual(read.A1, {kind: 'number', value: 42});
        assert.deepEqual(read.B1, {kind: 'boolean', value: true});
        assert.deepEqual(read.C1, {kind: 'formula', value: {formula: '1+1', result: 2}});
        assert.deepEqual(read.D1, {kind: 'date', value: '2024-01-15T00:00:00.000Z'});
      },
    },
    {
      name: 'a text label and an empty cell still read as links',
      expect(api: CorpusApi, assert: Assert) {
        const {read} = api.hyperlinkOverNonTextCellsReport();
        assert.deepEqual(read.E1, {kind: 'hyperlink', value: {...LINK, text: 'label'}});
        assert.deepEqual(read.F1, {kind: 'hyperlink', value: {...LINK, text: ''}});
      },
    },
    {
      name: 'a save keeps every value it read',
      expect(api: CorpusApi, assert: Assert) {
        const {read, rewritten} = api.hyperlinkOverNonTextCellsReport();
        assert.deepEqual(rewritten, read);
        assert.deepEqual(
          rewritten.A1,
          {kind: 'number', value: 42},
          'pinned, not only self-consistent',
        );
      },
    },
  ],
} satisfies Case;
