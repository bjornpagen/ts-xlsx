// Cluster: formulas
//
// Real-world scenario: a template has one formatted line item (`=B5*C5` for the amount) and the code
// filling it duplicates that row once per record. Every copy has to compute its own line, as it does
// when someone copies the row in Excel; a copy that still says `=B5*C5` shows the first line's amount on
// every line, in a file that opens without a complaint. Put to Excel Desktop over COM and recorded in
// `test/corpus/fixtures/excel-oracle/duplicate-row-formulas.json`: each copy is the source with its
// relative references moved down by its distance from the source, and after an insert the copy is
// taken from the source as the insert left it, so a source reading the inserted rows (`=A2`, now
// `=A4`) is copied as `=A5` and `=A6`.

import type {Assert, Case, CorpusApi} from '../case.ts';

const SOURCE = [
  {ref: 'B1', value: {formula: 'A1*2'}},
  {ref: 'C1', value: {formula: '$A$1+A$1+$A1'}},
  {ref: 'D1', value: {formula: 'A2'}},
];
const READ = ['B1', 'C1', 'D1', 'B2', 'C2', 'D2', 'B3', 'C3', 'D3'];

export default {
  id: 'duplicated-row-formulas-follow-each-copy',
  provenance: {
    source: 'excel-desktop-verification',
    ref: 'test/corpus/fixtures/excel-oracle/duplicate-row-formulas.json',
  },
  cluster: 'formulas',
  description:
    'A duplicated row copies its formulas as Excel copies a row: relative references move down with ' +
    'each copy and absolute ones stay, and an inserted copy is taken from the source as the insert ' +
    'left it.',

  behavior: [
    {
      name: 'copies inserted below the source read their own rows, as the insert left the source',
      expect(api: CorpusApi, assert: Assert) {
        const {cells, error} = api.mutateWorksheet({
          cells: SOURCE,
          ops: [{op: 'duplicateRow', start: 1, count: 2, insert: true}],
          read: READ,
        });
        assert.equal(error, null);
        assert.deepEqual(cells, {
          B1: {formula: 'A1*2'},
          C1: {formula: '$A$1+A$1+$A1'},
          D1: {formula: 'A4'},
          B2: {formula: 'A2*2'},
          C2: {formula: '$A$1+A$1+$A2'},
          D2: {formula: 'A5'},
          B3: {formula: 'A3*2'},
          C3: {formula: '$A$1+A$1+$A3'},
          D3: {formula: 'A6'},
        });
      },
    },
    {
      name: 'copies over the rows below read their own rows',
      expect(api: CorpusApi, assert: Assert) {
        const {cells, error} = api.mutateWorksheet({
          cells: SOURCE,
          ops: [{op: 'duplicateRow', start: 1, count: 2, insert: false}],
          read: READ,
        });
        assert.equal(error, null);
        assert.deepEqual(cells, {
          B1: {formula: 'A1*2'},
          C1: {formula: '$A$1+A$1+$A1'},
          D1: {formula: 'A2'},
          B2: {formula: 'A2*2'},
          C2: {formula: '$A$1+A$1+$A2'},
          D2: {formula: 'A3'},
          B3: {formula: 'A3*2'},
          C3: {formula: '$A$1+A$1+$A3'},
          D3: {formula: 'A4'},
        });
      },
    },
  ],
} satisfies Case;
