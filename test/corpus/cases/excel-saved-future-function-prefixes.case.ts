// Cluster: formulas
//
// Real-world scenario: a workbook saved by current Excel Desktop stores every function added since
// OOXML froze its grammar under a prefix, and the prefix differs by function. FILTER and SORT are
// "worksheet-only" and take `_xlfn._xlws.`, the rest take `_xlfn.`, a handful of post-2007 names such
// as NETWORKDAYS.INTL and ISO.CEILING take none, a spill reference `E1#` is stored as ANCHORARRAY, and a
// LET parameter is stored in the spelling it was declared with. A library that knows only `_xlfn.`
// leaves `_xlws.FILTER` in the formula text it reads, and one whose list is wrong writes formulas
// Excel reads as `#NAME?`.
//
// The fixture was saved by Excel Desktop 16.0 (build 20326) from these formulas typed into E1, E5, and
// on every fourth row to E29: FILTER(A1:A3,A1:A3), SORT(A1:A3), NETWORKDAYS.INTL(1,30), ISO.CEILING(4.3),
// CEILING.MATH(4.3), TRIMRANGE(A1:A3), LET(x,1,X+1) and SUM(E1#). The `excel` strings below are the
// formula text that Excel wrote into the sheet.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'excel-saved-future-function-prefixes/excel-saved.xlsx';

const FORMULAS = [
  {ref: 'E1', model: 'FILTER(A1:A3,A1:A3)', excel: '_xlfn._xlws.FILTER(A1:A3,A1:A3)'},
  {ref: 'E5', model: 'SORT(A1:A3)', excel: '_xlfn._xlws.SORT(A1:A3)'},
  {ref: 'E9', model: 'NETWORKDAYS.INTL(1,30)', excel: 'NETWORKDAYS.INTL(1,30)'},
  {ref: 'E13', model: 'ISO.CEILING(4.3)', excel: 'ISO.CEILING(4.3)'},
  {ref: 'E17', model: 'CEILING.MATH(4.3)', excel: '_xlfn.CEILING.MATH(4.3)'},
  {ref: 'E21', model: 'TRIMRANGE(A1:A3)', excel: '_xlfn.TRIMRANGE(A1:A3)'},
  {ref: 'E25', model: 'LET(x,1,x+1)', excel: '_xlfn.LET(_xlpm.x,1,_xlpm.x+1)'},
  {ref: 'E29', model: 'SUM(ANCHORARRAY(E1))', excel: 'SUM(_xlfn.ANCHORARRAY(E1))'},
];

export default {
  id: 'excel-saved-future-function-prefixes',
  cluster: 'formulas',
  provenance: {source: 'excel-desktop', excel: '16.0 build 20326'},
  description:
    'A workbook Excel saved with worksheet-only, prefixed and unprefixed post-2007 functions, a spill ' +
    'reference and a LET reads back as plain formula text and is written back in the exact spelling ' +
    'Excel gave each formula.',
  behavior: [
    {
      name: 'reading strips every prefix Excel wrote, _xlfn._xlws. included, back to the plain formula',
      async expect(api: CorpusApi, assert: Assert) {
        const cells = await api.readFixtureCells(
          FIXTURE,
          FORMULAS.map(({ref}) => ref),
        );
        for (const {ref, model} of FORMULAS) {
          assert.strictEqual(
            cells[ref]?.value?.formula,
            model,
            `${ref}: ${JSON.stringify(cells[ref])}`,
          );
        }
      },
    },
    {
      name: 'writing the workbook back stores each formula in the spelling Excel wrote',
      async expect(api: CorpusApi, assert: Assert) {
        const {cells} = await api.roundtripFixtureCellXml(
          FIXTURE,
          FORMULAS.map(({ref}) => ref),
        );
        for (const {ref, excel} of FORMULAS) {
          assert.strictEqual(cells[ref]?.formula, excel, ref);
        }
      },
    },
  ],
} satisfies Case;
