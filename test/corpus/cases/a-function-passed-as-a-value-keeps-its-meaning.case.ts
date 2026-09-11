// Cluster: formulas
//
// Real-world scenario: a formula hands a built-in function to a LAMBDA helper by name, as in
// `=BYROW(A1:A3,SUM)` or `=LET(f,MAX,f(A1:A3))`. Excel stores such a name as `_xleta.SUM`. Stored bare,
// `SUM` is a reference to a defined name, so a workbook that defines none shows `#VALUE!` or `#NAME?`
// where Excel computed a value. A defined name spelled like a function (`LEN`) is legal, and wherever
// it is visible the bare name means that name, so the prefix is decided by the names in the formula's
// scope, not by the function alone.
//
// The fixture was saved by Excel Desktop 16.0 (build 20326), as `.xlsx` and as `.xlsb`, from a
// workbook of two sheets holding 1, 2 and 3 in A1:A3, a `LEN` scoped to S2 defined as `LAMBDA(x,x*7)`, a
// workbook-level `Totals` defined as `SUM(BYROW(S1!$A$1:$A$3,SUM))`, and the formulas below typed into
// column E. `test/corpus/fixtures/excel-oracle/function-values.json` records the values Excel computed
// for them, the same values for this library's own write of them, and a control without the prefix.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'function-passed-as-a-value/excel-saved';

// `read` is the formula text Excel's `Range.Formula` reports, without its `=`; `stored` is the text
// Excel stored in the package.
const EXCEL: Record<string, {read: string; stored: string}> = {
  'S1!E1': {read: 'SUM(BYROW(A1:A3,SUM))', stored: 'SUM(_xlfn.BYROW(A1:A3,_xleta.SUM))'},
  'S1!E2': {
    read: 'CONCAT(BYROW(A1:A3,CONCAT))',
    stored: '_xlfn.CONCAT(_xlfn.BYROW(A1:A3,_xleta.CONCAT))',
  },
  // `LEN` is defined on S2 only, so on S1 it means the function, and the prefix is still reported.
  'S1!E3': {read: 'SUM(MAP(A1:A3,_xleta.LEN))', stored: 'SUM(_xlfn.MAP(A1:A3,_xleta.LEN))'},
  'S1!E4': {read: 'LET(f,MAX,f(A1:A3))', stored: '_xlfn.LET(_xlpm.f,_xleta.MAX,_xlpm.f(A1:A3))'},
  'S1!E5': {read: 'ISERROR(SUM+1)', stored: 'ISERROR(_xleta.SUM+1)'},
  'S1!E6': {read: 'Totals', stored: 'Totals'},
  // On S2 the bare `LEN` is the defined name, and the prefixed one is the function.
  'S2!E1': {read: 'SUM(MAP(A1:A3,LEN))', stored: 'SUM(_xlfn.MAP(A1:A3,LEN))'},
  'S2!E2': {read: 'SUM(MAP(A1:A3,_xleta.LEN))', stored: 'SUM(_xlfn.MAP(A1:A3,_xleta.LEN))'},
  'name:S2!LEN': {read: 'LAMBDA(x,x*7)', stored: '_xlfn.LAMBDA(_xlpm.x,_xlpm.x*7)'},
  'name:Totals': {
    read: "SUM(BYROW('S1'!$A$1:$A$3,SUM))",
    stored: "SUM(_xlfn.BYROW('S1'!$A$1:$A$3,_xleta.SUM))",
  },
};

const project = (field: 'read' | 'stored') =>
  Object.fromEntries(Object.entries(EXCEL).map(([key, spelling]) => [key, spelling[field]]));

export default {
  id: 'a-function-passed-as-a-value-keeps-its-meaning',
  cluster: 'formulas',
  provenance: {source: 'excel-desktop', excel: '16.0 build 20326'},
  description:
    'A built-in function a formula passes as a value is stored under the _xleta. prefix Excel requires, ' +
    'except where a defined name spelled like it is in scope, and a workbook Excel saved reads back as ' +
    'the formula text Excel reports and writes back in the spelling Excel stored.',
  behavior: [
    {
      name: 'reading a workbook Excel saved gives every formula and name the text Excel reports for it',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(api.fixtureFormulaSpellings(`${FIXTURE}.xlsx`).read, project('read'));
      },
    },
    {
      name: 'writing that workbook back stores every formula and name in the spelling Excel stored',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(api.fixtureFormulaSpellings(`${FIXTURE}.xlsx`).written, project('stored'));
      },
    },
    {
      // The binary form cites each prefixed function and parameter through a name record of its own,
      // which is Excel's bookkeeping and not one of the workbook's names.
      name: 'the binary save of the same workbook reads the same formulas and no names beyond its own',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(api.fixtureFormulaSpellings(`${FIXTURE}.xlsb`), {
          read: project('read'),
          written: project('stored'),
        });
      },
    },
    {
      name: 'an authored function passed as a value is stored with the prefix, in any position',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(
          api.storedFormulas({
            'S1!E1': 'SUM(BYROW(A1:A3,sum))',
            'S1!E2': 'CONCAT(BYROW(A1:A3,CONCAT))',
            'S1!E4': 'LET(f,MAX,f(A1:A3))',
            'S1!E5': 'ISERROR(-SUM+1)',
          }),
          {
            'S1!E1': 'SUM(_xlfn.BYROW(A1:A3,_xleta.SUM))',
            'S1!E2': '_xlfn.CONCAT(_xlfn.BYROW(A1:A3,_xleta.CONCAT))',
            'S1!E4': '_xlfn.LET(_xlpm.f,_xleta.MAX,_xlpm.f(A1:A3))',
            'S1!E5': 'ISERROR(-_xleta.SUM+1)',
          },
        );
      },
    },
    {
      name: 'a defined name spelled like the function keeps it bare only where the name is visible',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(
          api.storedFormulas({
            'name:S2!LEN': 'LAMBDA(x,x*7)',
            'S1!E3': 'SUM(MAP(A1:A3,LEN))',
            'S2!E1': 'SUM(MAP(A1:A3,LEN))',
            'name:Totals': 'SUM(BYROW(S1!$A$1:$A$3,SUM))',
          }),
          {
            'name:S2!LEN': '_xlfn.LAMBDA(_xlpm.x,_xlpm.x*7)',
            'S1!E3': 'SUM(_xlfn.MAP(A1:A3,_xleta.LEN))',
            'S2!E1': 'SUM(_xlfn.MAP(A1:A3,LEN))',
            'name:Totals': 'SUM(_xlfn.BYROW(S1!$A$1:$A$3,_xleta.SUM))',
          },
        );
      },
    },
    {
      name: 'a function name that is called, names a sheet or sits inside a literal is stored as written',
      expect(api: CorpusApi, assert: Assert) {
        const formulas = {
          'S1!E1': 'SUM(A1:A3)&"SUM"',
          'S1!E2': 'IF(ISNA(#N/A),S2!A1,LOG10)',
          'S1!E3': 'LET(SUM,A1:A3,SUM)',
        };
        assert.deepEqual(api.storedFormulas(formulas), {
          ...formulas,
          'S1!E3': '_xlfn.LET(_xlpm.SUM,A1:A3,_xlpm.SUM)',
        });
      },
    },
  ],
} satisfies Case;
