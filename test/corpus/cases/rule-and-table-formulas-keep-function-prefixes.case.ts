// Cluster: formulas
//
// Real-world scenario: a conditional format highlights the rows an `XLOOKUP` finds, a validation checks
// an entry against a `MAXIFS`, a colour scale is anchored at `MINIFS` and `MAXIFS`, a table column is
// calculated from `XLOOKUP` and `LET`. Excel stores every one of those formulas the way it stores a cell
// formula, with `_xlfn.` on a post-2007 function, `_xlpm.` on a LET parameter and, in a table, `_xleta.`
// on a function passed as a value. Stored bare, the rule reads the function as an unknown name: the
// format never applies, the validation rejects every entry, and a row added to the table is calculated
// from a formula Excel cannot evaluate. Read without shedding the prefixes, the model holds the on-disk
// spelling rather than the formula.
//
// The fixture was saved by Excel Desktop 16.0 (build 20326); `author.ps1` beside it builds it. The rule
// Excel keeps only in the x14 extension, a format reaching another sheet, reads onto its range like the
// rest and is written back to the extension, where its formula takes the same prefixes.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'rule-and-table-formulas-keep-function-prefixes/excel-saved.xlsx';

const MIN = 'MINIFS($A$1:$A$3,$A$1:$A$3,">0")';
const MAX = 'MAXIFS($A$1:$A$3,$A$1:$A$3,">0")';

// The formula text of each place, as a formula reads (`read`) and as Excel stored it (`stored`).
const EXCEL: Record<string, {read: string; stored: string}> = {
  'cf:A1:A3:0': {
    read: 'XLOOKUP(A1,$B$1:$B$3,$B$1:$B$3)=1',
    stored: '_xlfn.XLOOKUP(A1,$B$1:$B$3,$B$1:$B$3)=1',
  },
  'cf:B1:B3:0': {read: 'LET(x,B1,x>1)', stored: '_xlfn.LET(_xlpm.x,B1,_xlpm.x>1)'},
  'cf:C1:C3:0': {read: MAX, stored: `_xlfn.${MAX}`},
  'cf:E1:E3:cfvo0': {read: MIN, stored: `_xlfn.${MIN}`},
  'cf:E1:E3:cfvo1': {read: MAX, stored: `_xlfn.${MAX}`},
  'cf:F1:F3:cfvo0': {read: MIN, stored: `_xlfn.${MIN}`},
  'cf:F1:F3:cfvo1': {read: MAX, stored: `_xlfn.${MAX}`},
  'dv:D1:0': {
    read: 'ISNUMBER(XLOOKUP(D1,$A$1:$A$3,$A$1:$A$3))',
    stored: 'ISNUMBER(_xlfn.XLOOKUP(D1,$A$1:$A$3,$A$1:$A$3))',
  },
  'dv:D2:0': {read: MIN, stored: `_xlfn.${MIN}`},
  'dv:D2:1': {read: MAX, stored: `_xlfn.${MAX}`},
  // Kept in the extension by Excel, because it reaches another sheet.
  'dv:D3:0': {
    read: "ISNUMBER(XLOOKUP(D3,'S2'!$A$1:$A$3,'S2'!$A$1:$A$3))",
    stored: "ISNUMBER(_xlfn.XLOOKUP(D3,'S2'!$A$1:$A$3,'S2'!$A$1:$A$3))",
  },
  'table:T:a:totals': {read: 'SUM(BYROW(T[a],SUM))', stored: 'SUM(_xlfn.BYROW(T[a],_xleta.SUM))'},
  'table:T:b:calculated': {
    read: 'XLOOKUP(T[[#This Row],[a]],$A$1:$A$3,$B$1:$B$3)+LET(x,1,x)',
    stored: '_xlfn.XLOOKUP(T[[#This Row],[a]],$A$1:$A$3,$B$1:$B$3)+_xlfn.LET(_xlpm.x,1,_xlpm.x)',
  },
  'table:T:b:totals': {
    read: 'SUBTOTAL(109,T[b])+XLOOKUP(1,$A$1:$A$3,$B$1:$B$3)',
    stored: 'SUBTOTAL(109,T[b])+_xlfn.XLOOKUP(1,$A$1:$A$3,$B$1:$B$3)',
  },
};

// The format Excel keeps in the extension, because it reaches another sheet. Its formula is read like
// the rest, but written back to the extension rather than to the classic blocks `stored` reports.
const CROSS_SHEET_FORMAT = {
  read: "XLOOKUP(1,'S2'!$A$1:$A$3,'S2'!$B$1:$B$3)=1",
  stored: "_xlfn.XLOOKUP(1,'S2'!$A$1:$A$3,'S2'!$B$1:$B$3)=1",
};

const project = (field: 'read' | 'stored') =>
  Object.fromEntries(Object.entries(EXCEL).map(([key, spelling]) => [key, spelling[field]]));

export default {
  id: 'rule-and-table-formulas-keep-function-prefixes',
  cluster: 'formulas',
  provenance: {source: 'excel-desktop', excel: '16.0 build 20326'},
  description:
    'A conditional format, data validation or table column formula calling a post-2007 function reads ' +
    'back as plain formula text and is stored with the prefixes Excel gives it, as a cell formula is.',
  behavior: [
    {
      name: 'reading a workbook Excel saved gives every rule and table formula its plain text',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(api.ruleFormulaSpellings(FIXTURE).read, {
          ...project('read'),
          'cf:C4:C6:0': CROSS_SHEET_FORMAT.read,
        });
      },
    },
    {
      name: 'writing that workbook back stores every rule and table formula as Excel stored it',
      expect(api: CorpusApi, assert: Assert) {
        const {written, writtenExtension} = api.ruleFormulaSpellings(FIXTURE);
        assert.deepEqual(written, project('stored'));
        // The data bar's anchors are repeated in its x14 extension, beside the cross-sheet format and
        // the extended validation.
        assert.deepEqual(writtenExtension, [
          `_xlfn.${MIN}`,
          `_xlfn.${MAX}`,
          CROSS_SHEET_FORMAT.stored,
          "ISNUMBER(_xlfn.XLOOKUP(D3,'S2'!$A$1:$A$3,'S2'!$A$1:$A$3))",
        ]);
      },
    },
    {
      name: 'authored rule and table formulas are stored with the prefixes and read back without them',
      expect(api: CorpusApi, assert: Assert) {
        const {read, written, writtenExtension} = api.authoredRuleFormulaSpellings();
        const expected = {
          'cf:A1:A3:0': EXCEL['cf:A1:A3:0'],
          'cf:E1:E3:cfvo0': EXCEL['cf:E1:E3:cfvo0'],
          'cf:E1:E3:cfvo1': EXCEL['cf:E1:E3:cfvo1'],
          'cf:F1:F3:cfvo0': EXCEL['cf:F1:F3:cfvo0'],
          'cf:F1:F3:cfvo1': EXCEL['cf:F1:F3:cfvo1'],
          'dv:D1:0': EXCEL['dv:D1:0'],
          'dv:D3:0': {
            read: 'ISNUMBER(XLOOKUP(D3,S2!$A$1:$A$3,S2!$A$1:$A$3))',
            stored: 'ISNUMBER(_xlfn.XLOOKUP(D3,S2!$A$1:$A$3,S2!$A$1:$A$3))',
          },
          'table:T:b:calculated': EXCEL['table:T:b:calculated'],
          'table:T:b:totals': EXCEL['table:T:b:totals'],
        };
        assert.deepEqual(
          read,
          Object.fromEntries(Object.entries(expected).map(([key, value]) => [key, value?.read])),
        );
        assert.deepEqual(
          written,
          Object.fromEntries(Object.entries(expected).map(([key, value]) => [key, value?.stored])),
        );
        assert.deepEqual(writtenExtension, [
          `_xlfn.${MIN}`,
          `_xlfn.${MAX}`,
          'ISNUMBER(_xlfn.XLOOKUP(D3,S2!$A$1:$A$3,S2!$A$1:$A$3))',
        ]);
      },
    },
  ],
} satisfies Case;
