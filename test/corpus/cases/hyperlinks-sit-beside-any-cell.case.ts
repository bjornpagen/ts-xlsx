// Cluster: hyperlinks
//
// Real-world scenario: a report links its totals, its dates and its computed figures to the detail
// behind them, and a fill job then inserts and deletes rows through the linked ranges. Excel keeps a
// hyperlink beside the grid, covering a range of cells, and leaves each cell's value alone. A library
// that treats a link as a text label on one cell loses the link from every number, date and formula,
// and moves a ranged link as a single cell: a row deleted from the top of a linked range takes the whole
// link with it, where Excel shrinks the range and keeps the link.
//
// The fixture was saved by Excel Desktop 16.0 (build 20326) with a link on a number, a date formula, a
// formula, a boolean and a text cell. The edits below are the ones recorded, with what Excel did to the
// links, in `test/corpus/fixtures/excel-oracle/hyperlinks-through-edits.json`.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'hyperlinks-beside-any-cell/excel-saved.xlsx';
const REFS = ['A1', 'A2', 'A3', 'A4', 'A5'];

export default {
  id: 'hyperlinks-sit-beside-any-cell',
  cluster: 'hyperlinks',
  provenance: {source: 'excel-desktop', excel: '16.0 build 20326'},
  description:
    'A hyperlink covers a range of cells beside the grid, over any value: a workbook Excel saved with ' +
    'links on a number, a date, a formula and a boolean reads and writes back every link and every ' +
    'value, and links move through row inserts, deletes, copies and column inserts as Excel moves them.',
  behavior: [
    {
      name: 'every link Excel saved is read, and the value under each keeps its type',
      expect(api: CorpusApi, assert: Assert) {
        const {read} = api.fixtureHyperlinksReport(FIXTURE, REFS);
        assert.deepEqual(
          read.links,
          REFS.map((ref) => ({ref, target: 'https://example.com/', tooltip: 'tip'})),
        );
        assert.deepEqual(read.cells.A1, {kind: 'number', value: 42});
        assert.deepEqual(read.cells.A2, {
          kind: 'formula',
          value: {formula: 'DATE(2024,1,2)', result: '2024-01-02T00:00:00.000Z'},
        });
        assert.deepEqual(read.cells.A3, {kind: 'formula', value: {formula: '1+1', result: 2}});
        assert.deepEqual(read.cells.A4, {kind: 'boolean', value: true});
        assert.deepEqual(read.cells.A5, {kind: 'string', value: 'https://example.com/'});
      },
    },
    {
      name: 'a save writes back every link and keeps every value',
      expect(api: CorpusApi, assert: Assert) {
        const {read, rewritten} = api.fixtureHyperlinksReport(FIXTURE, REFS);
        assert.deepEqual(rewritten, read);
      },
    },
    {
      name: 'a link over a cell inside a wider link is kept beside it, and still wins there after a save',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(api.hyperlinksThroughEdits().overlap, {
          refs: ['A1:C1', 'B1'],
          b1: 'https://two.example/',
          c1: 'https://one.example/',
        });
      },
    },
    {
      name: 'a link grows with a row inserted inside it, shrinks with its first row deleted, and goes with all of them',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(api.hyperlinksThroughEdits().splice, {
          afterInsertInside: ['A2:A5'],
          afterTopRowDelete: ['A2:A4'],
          afterWholeDelete: [],
        });
      },
    },
    {
      name: 'a copied row carries its link, and a column inserted before the links moves them all',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(api.hyperlinksThroughEdits().copyAndColumn, {
          afterRowCopy: ['A1', 'A2'],
          afterColumnInsert: ['B1', 'B2'],
        });
      },
    },
  ],
} satisfies Case;
