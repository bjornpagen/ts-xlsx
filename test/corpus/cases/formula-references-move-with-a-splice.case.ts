// Cluster: formulas
//
// Real-world scenario: a report template sums a block of rows (`SUM(B2:B20)`), another sheet pulls a
// total from it (`Data!B21`), a defined name points at the block, and a dropdown draws its choices from
// a range. The code filling the template inserts a row per record, or deletes the rows it does not
// need. Every coordinate the sheet stores as data moves with the rows, but formula text spells its
// coordinates, and a formula left as written sums the wrong cells in a file that opens without a
// complaint. Excel moves that text on the same edit, by rules put to Excel Desktop over COM and
// recorded in `test/corpus/fixtures/excel-oracle/formula-references-through-a-splice.json`; every
// expectation below is what Excel read back after the same edit.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'formula-references-move-with-a-splice',
  provenance: {
    source: 'excel-desktop-verification',
    ref: 'test/corpus/fixtures/excel-oracle/formula-references-through-a-splice.json',
  },
  cluster: 'formulas',
  description:
    'Inserting or deleting rows or columns moves the references formula text makes to the edited ' +
    "sheet, as Excel does: on that sheet, on other sheets, in defined names and in a validation's " +
    'source. An insert moves absolute and relative references alike and grows a range it lands in; a ' +
    'delete turns what it takes into #REF! and shrinks a range it cuts into; a 3-D span and a string ' +
    'do not move.',

  behavior: [
    {
      name: 'two rows inserted move every formula, name and validation source that refers to them',
      expect(api: CorpusApi, assert: Assert) {
        const after = api.formulasAfterSplice({
          formulas: {
            'S1!D1': '=SUM(A1:A10)',
            'S1!D2': '=A5+$A$5+A$5+$A5',
            'S1!D3': '=SUM(5:6)+SUM(A:A)+SUM(4:4)',
            'S1!E1': '=SUM(A11:A12)',
            'S2!A1': "='S1'!A5+SUM('S1'!A1:A10)+A5",
            'S2!A2': "=SUM(S1:'S3'!A5)",
            'S2!A3': '=INDIRECT("S1!A5")',
            'name:Five': "='S1'!$A$5",
            'name:Span': "='S1'!$A$1:$A$10",
            'dv:S1!F20': '=$A$5:$A$9',
          },
          edit: {op: 'insert', axis: 'row', start: 5, count: 2},
          read: [
            'S1!D1',
            'S1!D2',
            'S1!D3',
            'S1!E1',
            'S2!A1',
            'S2!A2',
            'S2!A3',
            'name:Five',
            'name:Span',
            'dv:S1!F22',
          ],
        });
        assert.deepEqual(after, {
          'S1!D1': '=SUM(A1:A12)',
          'S1!D2': '=A7+$A$7+A$7+$A7',
          'S1!D3': '=SUM(7:8)+SUM(A:A)+SUM(4:4)',
          'S1!E1': '=SUM(A13:A14)',
          'S2!A1': "='S1'!A7+SUM('S1'!A1:A12)+A5",
          'S2!A2': "=SUM(S1:'S3'!A5)",
          'S2!A3': '=INDIRECT("S1!A5")',
          'name:Five': "='S1'!$A$7",
          'name:Span': "='S1'!$A$1:$A$12",
          'dv:S1!F22': '=$A$7:$A$11',
        });
      },
    },
    {
      name: 'a row inserted at the top of a range moves it whole, and one below a range leaves it',
      expect(api: CorpusApi, assert: Assert) {
        const after = api.formulasAfterSplice({
          formulas: {'S1!D1': '=SUM(A5:A10)', 'S1!D2': '=SUM(A1:A4)'},
          edit: {op: 'insert', axis: 'row', start: 5, count: 1},
          read: ['S1!D1', 'S1!D2'],
        });
        assert.deepEqual(after, {'S1!D1': '=SUM(A6:A11)', 'S1!D2': '=SUM(A1:A4)'});
      },
    },
    {
      name: 'deleted rows turn what they held into #REF!, keeping its sheet, and shrink what they cut',
      expect(api: CorpusApi, assert: Assert) {
        const after = api.formulasAfterSplice({
          formulas: {
            'S1!D1': '=SUM(A3:A10)',
            'S1!D2': '=A5+A7+A12',
            'S1!D3': '=SUM(A5:A6)',
            'S1!D4': '=SUM(A6:A12)',
            'S1!E1': '=SUM(5:6)+SUM(4:7)',
            'S2!A1': "='S1'!A5+SUM('S1'!A5:A6)+'S1'!A8",
            'name:Five': "='S1'!$A$5",
          },
          edit: {op: 'delete', axis: 'row', start: 5, count: 2},
          read: ['S1!D1', 'S1!D2', 'S1!D3', 'S1!D4', 'S1!E1', 'S2!A1', 'name:Five'],
        });
        assert.deepEqual(after, {
          'S1!D1': '=SUM(A3:A8)',
          'S1!D2': '=#REF!+A5+A10',
          'S1!D3': '=SUM(#REF!)',
          'S1!D4': '=SUM(A5:A10)',
          'S1!E1': '=SUM(#REF!)+SUM(4:5)',
          'S2!A1': "='S1'!#REF!+SUM('S1'!#REF!)+'S1'!A6",
          'name:Five': "='S1'!#REF!",
        });
      },
    },
    {
      name: 'deleted columns follow the same rules across',
      expect(api: CorpusApi, assert: Assert) {
        const after = api.formulasAfterSplice({
          formulas: {'S1!A10': '=SUM(B1:E1)+C1+F1+SUM(C:D)+SUM(C1:D1)'},
          edit: {op: 'delete', axis: 'column', start: 3, count: 2},
          read: ['S1!A10'],
        });
        assert.deepEqual(after, {'S1!A10': '=SUM(B1:C1)+#REF!+D1+SUM(#REF!)+SUM(#REF!)'});
      },
    },
    {
      name: 'a reference pushed past the last row is #REF!, and a range bottom already on it stays',
      expect(api: CorpusApi, assert: Assert) {
        const after = api.formulasAfterSplice({
          formulas: {'S1!D1': '=A1048576', 'S1!D2': '=SUM(A1048570:A1048576)'},
          edit: {op: 'insert', axis: 'row', start: 3, count: 1},
          read: ['S1!D1', 'S1!D2'],
        });
        assert.deepEqual(after, {'S1!D1': '=#REF!', 'S1!D2': '=SUM(A1048571:A1048576)'});
      },
    },
  ],
} satisfies Case;
