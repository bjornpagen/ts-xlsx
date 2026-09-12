// Cluster: types
//
// Real-world scenario: a workbook holds errors current Excel shows. `#SPILL!` from a dynamic array with no
// room, `#CALC!` from a FILTER that matched nothing, `#FIELD!` from a field a value does not have, beside
// the classic `#N/A` and `#DIV/0!`. A generator copies such results into a new workbook, or a user reads
// one and saves it. Excel stores the classic errors, `#GETTING_DATA` and `#BUSY!` literally, as a typed
// cell's `<v>`. It has no literal for `#SPILL!`, `#CONNECT!`, `#BLOCKED!`, `#UNKNOWN!`, `#FIELD!` or
// `#CALC!`: a cell spelling one of them makes Excel offer to repair the package, and Excel itself stores
// each as `<v>#VALUE!</v>` with a `vm` naming a rich value whose `errorType` is the real error. Such a cell
// used to read as `#VALUE!`, and a save dropped the rich value, so every one of them became `#VALUE!`.
//
// Which `errorType` is which error came from Excel Desktop (build 20326), recorded in
// `test/corpus/fixtures/excel-oracle/rich-value-errors.json`.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'an-error-value-is-stored-as-excel-stores-it/excel-saved.xlsx';
const LITERALS = ['#N/A', '#DIV/0!', '#GETTING_DATA', '#BUSY!'];
const ERROR_TYPES: Record<string, number> = {
  '#SPILL!': 8,
  '#CONNECT!': 9,
  '#BLOCKED!': 10,
  '#UNKNOWN!': 11,
  '#FIELD!': 12,
  '#CALC!': 13,
};

export default {
  id: 'an-error-value-is-stored-as-excel-stores-it',
  provenance: {
    source: 'excel-desktop-verification',
    ref: 'test/corpus/fixtures/excel-oracle/rich-value-errors.json',
  },
  cluster: 'types',
  description:
    'An error Excel reads from a typed cell is written as its literal; one Excel has no literal for, such ' +
    'as #SPILL! or #CALC!, is written as #VALUE! beside the rich value naming it, and a workbook Excel ' +
    'saved reads and saves every such error as the error Excel showed.',

  behavior: [
    {
      name: 'every error Excel reads literally is written as its literal and reads back as itself',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.errorLiteralReport(LITERALS);
        for (const code of LITERALS) {
          assert.deepStrictEqual(
            report[code],
            {written: code, errorType: null, readBack: {error: code}, refused: null},
            code,
          );
        }
      },
    },
    {
      name: 'an error Excel has no literal for is written as #VALUE! beside the rich value naming it',
      expect(api: CorpusApi, assert: Assert) {
        const codes = Object.keys(ERROR_TYPES);
        const report = api.errorLiteralReport(codes);
        for (const code of codes) {
          assert.deepStrictEqual(
            report[code],
            {
              written: '#VALUE!',
              errorType: ERROR_TYPES[code],
              readBack: {error: code},
              refused: null,
            },
            code,
          );
        }
      },
    },
    {
      name: 'a workbook Excel saved reads each error as the error Excel showed',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepStrictEqual(api.richValueErrorReport(FIXTURE).read, {
          // Computed: a spill blocked by a value, a FILTER matching nothing, a FIELDVALUE of a number, a
          // bare LAMBDA, a spill blocked by a merge and one off the sheet's edge.
          A1: '#SPILL!',
          C1: '#CALC!',
          E1: '#FIELD!',
          I1: '#CALC!',
          J1: '#SPILL!',
          A1048576: '#SPILL!',
          // Pasted as values from A1 and C1.
          G1: '#SPILL!',
          G2: '#CALC!',
        });
      },
    },
    {
      name: 'a save stores every error cell as Excel stored it, the rich value included',
      expect(api: CorpusApi, assert: Assert) {
        const {excel, written} = api.richValueErrorReport(FIXTURE);
        assert.deepStrictEqual(written, excel);
      },
    },
  ],
} satisfies Case;
