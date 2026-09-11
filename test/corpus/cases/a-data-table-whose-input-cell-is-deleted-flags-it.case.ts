// Cluster: formulas
//
// Real-world scenario: a sensitivity table built with What-If Analysis takes its input from a cell in a
// header block, and someone deletes the header rows above the table. The table moves up with the grid,
// and its input cell is gone. What Excel does was asked of Excel Desktop, recorded in
// `test/corpus/fixtures/excel-oracle/data-table-input-deleted.json`: whether the row, the column or the
// cell alone is deleted, it keeps the input reference exactly as written and flags it with `del1`, and
// the table shows `#REF!`. The two fixtures are Excel's own saves of the table before and after.
//
// A splice that left the reference unflagged would write a table reading its input from whatever cell
// now sits at the old address, which is a different calculation, silently. One that re-aimed the
// reference at a surviving neighbour would do the same with more confidence.

import type {Assert, Case, CorpusApi} from '../case.ts';

const DIR = 'data-table-input-cell-deleted';

export default {
  id: 'a-data-table-whose-input-cell-is-deleted-flags-it',
  provenance: {
    source: 'excel-desktop-verification',
    ref: 'test/corpus/fixtures/excel-oracle/data-table-input-deleted.json',
  },
  cluster: 'formulas',
  description:
    'A data table whose input cell a row or column delete takes keeps the input reference as written ' +
    'and flags it deleted (`del1`), as Excel does, while the table itself moves with the grid; a file ' +
    'Excel saved in that state reads with the flag and writes it back.',

  behavior: [
    {
      name: 'the table Excel saved after deleting the input row reads with its input flagged deleted',
      expect(api: CorpusApi, assert: Assert) {
        const {excel} = api.dataTableInputDeletionReport(DIR);
        assert.deepStrictEqual([excel.ref, excel.r1, excel.r1Deleted], ['B3:B5', 'D1', true]);
        assert.strictEqual(excel.written.del1, '1', 'and writes the flag back');
        assert.strictEqual(excel.written.r1, 'D1');
      },
    },
    {
      name: 'deleting the input row here does what Excel did: the table moves, the reference stays and is flagged',
      expect(api: CorpusApi, assert: Assert) {
        const {excel, spliced} = api.dataTableInputDeletionReport(DIR);
        assert.strictEqual(spliced.r1Deleted, true, 'the input the delete took is flagged');
        assert.deepStrictEqual(
          [spliced.ref, spliced.r1, spliced.r1Deleted],
          [excel.ref, excel.r1, excel.r1Deleted],
        );
        assert.strictEqual(spliced.written.ref, excel.written.ref);
        assert.strictEqual(spliced.written.del1, excel.written.del1);
        assert.strictEqual(spliced.written.r1, excel.written.r1);
      },
    },
  ],
} satisfies Case;
