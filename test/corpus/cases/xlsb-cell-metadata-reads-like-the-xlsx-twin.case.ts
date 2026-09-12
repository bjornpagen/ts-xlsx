// Cluster: xlsx-io
//
// Real-world scenario: a workbook saved as `.xlsb` holds formulas current Excel writes. `=SEQUENCE(3)`
// spills, `=SUM(D1:D3*2)` entered in current Excel is a one-cell dynamic array, a formula entered with
// Ctrl+Shift+Enter fills its range, and a spill with no room shows `#SPILL!`. The binary form marks a
// dynamic array, and names an error it has no literal for, the way the XML form does: through metadata
// indices a cell carries, here as `BrtCellMeta` and `BrtValueMeta` records ahead of the cell record, into
// `xl/metadata.bin`. The binary reader looked at neither, so every array formula read as a legacy one and
// every such error as `#VALUE!`, and the `.xlsb` disagreed with its `.xlsx` twin.
//
// The pair is one workbook Excel saved both ways (`author.ps1` beside them). What each record means was
// checked against Excel itself, recorded in `test/corpus/fixtures/excel-oracle/xlsb-cell-metadata.json`.

import type {Assert, Case, CorpusApi} from '../case.ts';

const XLSB = 'an-error-value-is-stored-as-excel-stores-it/excel-saved.xlsb';
const XLSX = 'an-error-value-is-stored-as-excel-stores-it/excel-saved.xlsx';

export default {
  id: 'xlsb-cell-metadata-reads-like-the-xlsx-twin',
  provenance: {
    source: 'excel-desktop-verification',
    ref: 'test/corpus/fixtures/excel-oracle/xlsb-cell-metadata.json',
  },
  cluster: 'xlsx-io',
  description:
    'A binary .xlsb workbook reads its dynamic-array formulas as dynamic and its rich-value errors as ' +
    'the errors they name, from the metadata its cells point into, as its .xlsx twin does.',

  behavior: [
    {
      name: 'the binary and XML readings of the workbook produce the same model',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepStrictEqual(api.xlsbModelMatchesXlsxTwin(XLSB, XLSX), {
          identical: true,
          firstDifference: null,
        });
      },
    },
    {
      name: 'a dynamic-array formula reads as dynamic, and a Ctrl+Shift+Enter one as legacy',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepStrictEqual(api.xlsbCellValues(XLSB, ['M1', 'O1', 'Q1', 'S1']), {
          M1: {formula: 'SEQUENCE(3)', result: 1, array: {ref: 'M1:M3', dynamic: true}},
          O1: {formula: 'SUM(D1:D3*2)', result: 12, array: {ref: 'O1', dynamic: true}},
          Q1: {formula: 'D1:D3*2', result: 2, array: {ref: 'Q1:Q3', dynamic: false}},
          S1: {formula: 'D1*2', result: 2},
        });
      },
    },
    {
      name: 'an error the value metadata names reads as that error, computed or pasted',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepStrictEqual(api.xlsbCellValues(XLSB, ['A1', 'C1', 'E1', 'G1', 'G2', 'I1']), {
          A1: {
            formula: 'SEQUENCE(3)',
            result: {error: '#SPILL!'},
            array: {ref: 'A1', dynamic: true},
          },
          C1: {
            formula: 'FILTER(D1:D3,D1:D3>100)',
            result: {error: '#CALC!'},
            array: {ref: 'C1', dynamic: true},
          },
          E1: {
            formula: 'FIELDVALUE(1,"x")',
            result: {error: '#FIELD!'},
            array: {ref: 'E1', dynamic: true},
          },
          G1: {error: '#SPILL!'},
          G2: {error: '#CALC!'},
          I1: {formula: 'LAMBDA(x,x)', result: {error: '#CALC!'}},
        });
      },
    },
  ],
} satisfies Case;
