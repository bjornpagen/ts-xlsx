// Cluster: formulas
//
// Real-world scenario: a column of `COUNTIF(A:A,A1)` filled right, or a row of `SUM(1:1)` filled down,
// is exactly what Excel stores as a shared formula: the master carries the text and every clone
// carries only a pointer to it. A reader recovers each clone's own text by shifting the master's
// references by the clone's offset. When only `A1`-shaped references move, a whole-column or whole-row
// range stays aimed at the master's lines, so every clone reads back counting the wrong column, and a
// save stores that wrong formula over the one the author wrote.

import type {Assert, Case, CorpusApi} from '../case.ts';

// A `COUNTIF` over a whole column with its clone one column to the right, and a `SUM` over a whole
// row with its clone one row below.
const SPEC = {
  sheets: [
    {
      name: 'S',
      cells: [
        {ref: 'A1', value: 1},
        {ref: 'B1', value: 2},
        {ref: 'C1', formula: 'COUNTIF(A:A,A1)', result: 1},
        {ref: 'D1', sharedFormula: 'C1', result: 0},
        {ref: 'A3', formula: 'SUM(1:1)', result: 3},
        {ref: 'A4', sharedFormula: 'A3', result: 0},
      ],
    },
  ],
};

export default {
  id: 'shared-formula-clones-shift-whole-column-and-row-ranges',
  provenance: {source: 'audit'},
  cluster: 'formulas',
  description:
    'A shared-formula clone reads back with its whole-column and whole-row ranges shifted by its ' +
    'offset from the master, exactly as its cell references are: a `COUNTIF(A:A,A1)` master filled ' +
    'one column right reads `COUNTIF(B:B,B1)`, and a `SUM(1:1)` master filled one row down reads ' +
    '`SUM(2:2)`.',

  behavior: [
    {
      name: 'a clone filled right shifts a whole-column range along with its cell reference',
      expect(api: CorpusApi, assert: Assert) {
        assert.equal(api.roundtripFormulas(SPEC).D1.formula, 'COUNTIF(B:B,B1)');
      },
    },
    {
      name: 'a clone filled down shifts a whole-row range',
      expect(api: CorpusApi, assert: Assert) {
        assert.equal(api.roundtripFormulas(SPEC).A4.formula, 'SUM(2:2)');
      },
    },
    {
      name: 'the masters keep their own text and the clones still name them',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.roundtripFormulas(SPEC);
        assert.equal(report.C1.formula, 'COUNTIF(A:A,A1)');
        assert.equal(report.A3.formula, 'SUM(1:1)');
        assert.equal(report.D1.sharedFormula, 'C1');
        assert.equal(report.A4.sharedFormula, 'A3');
      },
    },
  ],
} satisfies Case;
