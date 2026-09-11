// Cluster: tables
//
// Real-world scenario: a report template keeps its line items in a table whose calculated column
// reaches a rate cell below it (`$E$20`) and a lookup on another sheet, with a custom total beside it,
// and a pivot summarising a data sheet. The code filling the template inserts rows for its records.
// The cells' own formulas move with the rows, but a table stores its calculated column formula and
// its totals formula in the table part, and a pivot stores its source range in the cache: left as
// written, Excel fills the next row added to the table with a formula reaching the wrong rate, and
// refreshes the pivot over rows that are no longer the data.
//
// Excel moves all three on the same edit. The workbook in the fixture is Excel's, and every expected
// spelling below is what Excel stored after the same edits, recorded in
// `test/corpus/fixtures/excel-oracle/references-outside-cells-through-a-splice.json`.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'references-outside-cells-through-a-splice/before.xlsx';

export default {
  id: 'table-formulas-and-pivot-sources-move-with-a-splice',
  provenance: {
    source: 'excel-desktop-verification',
    ref: 'test/corpus/fixtures/excel-oracle/references-outside-cells-through-a-splice.json',
  },
  cluster: 'tables',
  description:
    "A row or column splice moves the references in a table column's calculated column formula and " +
    'totals row formula, and the source range of an authored pivot, as Excel moves them. A table ' +
    "formula's unqualified references mean the table's sheet; a pivot source grows and shrinks with " +
    'the lines it spans, and one whose every row or column is deleted stays as written.',

  behavior: [
    {
      name: "a table's calculated column and custom total move with inserts on its sheet and another, and a delete makes #REF!",
      expect(api: CorpusApi, assert: Assert) {
        const [aboveTable, onOtherSheet, deleteRate] = api.tableFormulasThroughSplices(FIXTURE, [
          {sheet: 'S1', op: 'insert', axis: 'row', start: 5, count: 2},
          {sheet: 'S2', op: 'insert', axis: 'row', start: 2, count: 2},
          {sheet: 'S1', op: 'delete', axis: 'row', start: 22, count: 1},
        ]);
        assert.deepEqual(aboveTable, {
          'a totals': "SUM($E$22:$E$24)+'S2'!A5",
          'c calculated': "T[[#This Row],[a]]+$E$22+'S2'!$A$5+E22",
        });
        assert.deepEqual(onOtherSheet, {
          'a totals': "SUM($E$22:$E$24)+'S2'!A7",
          'c calculated': "T[[#This Row],[a]]+$E$22+'S2'!$A$7+E22",
        });
        assert.deepEqual(deleteRate, {
          'a totals': "SUM($E$22:$E$23)+'S2'!A7",
          'c calculated': "T[[#This Row],[a]]+#REF!+'S2'!$A$7+#REF!",
        });
      },
    },
    {
      name: "an authored pivot's source grows with an insert inside it and shrinks with a delete",
      expect(api: CorpusApi, assert: Assert) {
        const source = (
          op: 'insert' | 'delete',
          axis: 'row' | 'column',
          start: number,
          count = 1,
        ) => api.authoredPivotSourceAfterSplice({op, axis, start, count});
        assert.deepEqual(source('insert', 'row', 3), {ref: 'A1:B6', sheet: 'S3'});
        assert.deepEqual(source('insert', 'column', 2), {ref: 'A1:C5', sheet: 'S3'});
        assert.deepEqual(source('delete', 'row', 1), {ref: 'A1:B4', sheet: 'S3'});
        assert.deepEqual(source('delete', 'row', 2, 4), {ref: 'A1:B1', sheet: 'S3'});
      },
    },
    {
      name: "a delete that takes every row or every column of a pivot's source leaves the source as written",
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(
          api.authoredPivotSourceAfterSplice({op: 'delete', axis: 'row', start: 1, count: 5}),
          {ref: 'A1:B5', sheet: 'S3'},
        );
        assert.deepEqual(
          api.authoredPivotSourceAfterSplice({op: 'delete', axis: 'column', start: 1, count: 2}),
          {ref: 'A1:B5', sheet: 'S3'},
        );
      },
    },
  ],
} satisfies Case;
