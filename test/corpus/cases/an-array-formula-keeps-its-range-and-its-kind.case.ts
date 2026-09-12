// Cluster: formulas
//
// Real-world scenario: a workbook saved by current Excel holds formulas that fill more than their own
// cell. A dynamic-array formula (`=SEQUENCE(3)`, `=TRANSPOSE(A1:A3)`) spills its results down or
// across, and a legacy array formula entered with Ctrl+Shift+Enter fills the range it was entered over.
// Excel stores both as `<f t="array" ref>` on the cell the range starts at, and marks the dynamic kind
// only through the cell's `cm` link into the workbook's cell metadata. A user reads such a workbook,
// changes nothing, and saves it. Both kinds used to come back as plain formulas: every range collapsed
// to its formula's own cell, and a spilling formula opened as one that no longer spilled.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'an-array-formula-keeps-its-range-and-its-kind/excel-saved.xlsx';

export default {
  id: 'an-array-formula-keeps-its-range-and-its-kind',
  provenance: {
    source: 'excel-desktop-verification',
    ref: 'test/corpus/fixtures/excel-oracle/array-formulas.json',
  },
  cluster: 'formulas',
  description:
    'An array formula reads as its own kind over the range it fills, a dynamic-array formula is told ' +
    'from a Ctrl+Shift+Enter one by the cell metadata Excel marks it with, and a save stores each ' +
    'formula, range and metadata link as Excel stored them.',

  behavior: [
    {
      name: 'reading tells a dynamic-array formula from a Ctrl+Shift+Enter one, each over its range',
      expect(api: CorpusApi, assert: Assert) {
        const {read} = api.arrayFormulaReport(FIXTURE);
        const array = (ref: string, dynamic: boolean, formula: string, result: number) => ({
          kind: 'array',
          ref,
          dynamic,
          formula,
          result,
        });
        assert.deepStrictEqual(
          {B1: read.B1, D1: read.D1, F1: read.F1, H1: read.H1, J1: read.J1, L1: read.L1},
          {
            B1: array('B1:B3', true, 'SEQUENCE(3)', 1),
            D1: array('D1:D3', false, 'A1:A3*2', 2),
            // One cell each: entered with Ctrl+Shift+Enter, and entered in current Excel over arrays.
            F1: array('F1', false, 'SUM(A1:A3*2)', 12),
            H1: array('H1', true, 'SUM(A1:A3*2)', 12),
            J1: {kind: 'formula', ref: null, dynamic: false, formula: 'A1*2', result: 2},
            L1: array('L1:N1', true, 'TRANSPOSE(A1:A3)', 1),
          },
        );
      },
    },
    {
      name: 'the other cells of a range read as the values the formula produced',
      expect(api: CorpusApi, assert: Assert) {
        const {read} = api.arrayFormulaReport(FIXTURE);
        assert.deepStrictEqual(
          {B2: read.B2, B3: read.B3, D2: read.D2, D3: read.D3, M1: read.M1, N1: read.N1},
          {B2: 2, B3: 3, D2: 4, D3: 6, M1: 2, N1: 3},
        );
      },
    },
    {
      name: 'a save stores every formula as Excel stored it, range and spill link included',
      expect(api: CorpusApi, assert: Assert) {
        const {excel, written} = api.arrayFormulaReport(FIXTURE);
        assert.deepStrictEqual(written, excel);
      },
    },
    {
      name: 'an array formula authored here is stored the way Excel stores one',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepStrictEqual(api.authoredArrayFormulas(), {
          B1: {t: 'array', ref: 'B1:B3', dynamic: true},
          D1: {t: 'array', ref: 'D1:D3', dynamic: false},
        });
      },
    },
  ],
} satisfies Case;
