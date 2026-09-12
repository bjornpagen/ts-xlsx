// Cluster: formulas
//
// Real-world scenario: a sheet holds a Ctrl+Shift+Enter formula filling B1:B3, and a report generator
// duplicates the header row above it, inserts a row of totals through it, or writes a formula into one of
// its cells. Excel refuses each of those edits, as a change to part of an array, and offers to repair any
// package whose array formula's range holds another formula or shares a cell with another array formula's
// range, dynamic arrays included. The library used to make every one of those edits and write every one
// of those shapes, so a duplicated row became a package Excel repaired.
//
// What Excel refuses and allows came from Excel Desktop (build 20326), recorded in
// `test/corpus/fixtures/excel-oracle/array-formula-ranges.json`.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'an-array-formula-range-holds-no-other-formula',
  provenance: {
    source: 'excel-desktop-verification',
    ref: 'test/corpus/fixtures/excel-oracle/array-formula-ranges.json',
  },
  cluster: 'formulas',
  description:
    'An array formula range holding another formula, or sharing a cell with another array formula range, ' +
    'is refused at write and read as a plain formula from a file; an edit cutting through a ' +
    'Ctrl+Shift+Enter range is refused, and one moving or removing it whole goes ahead, as in Excel.',

  behavior: [
    {
      name: 'a range holding another formula or sharing a cell with another range is refused at write',
      expect(api: CorpusApi, assert: Assert) {
        const {authored} = api.arrayRangeReport();
        assert.match(authored.formulaInsideLegacy ?? '', /B1.*B2/);
        assert.match(authored.formulaInsideDynamic ?? '', /B1.*B2/);
        assert.match(authored.cornerOverlap ?? '', /A2.*B1/);
        assert.strictEqual(
          authored.valuesInside,
          null,
          'the values an array formula produced are written',
        );
      },
    },
    {
      name: 'an edit cutting through a Ctrl+Shift+Enter range is refused and leaves the range as it was',
      expect(api: CorpusApi, assert: Assert) {
        const {edits} = api.arrayRangeReport();
        assert.deepStrictEqual(
          {
            insertRowInside: edits.insertRowInside,
            deleteRowInside: edits.deleteRowInside,
            deleteRowsCrossing: edits.deleteRowsCrossing,
            duplicateAnchorRow: edits.duplicateAnchorRow,
            insertColumnInside: edits.insertColumnInside,
            deleteColumnPart: edits.deleteColumnPart,
          },
          {
            insertRowInside: {refused: true, ref: 'B1:B3'},
            deleteRowInside: {refused: true, ref: 'B1:B3'},
            deleteRowsCrossing: {refused: true, ref: 'B1:B3'},
            duplicateAnchorRow: {refused: true, ref: 'B1:B3'},
            insertColumnInside: {refused: true, ref: 'B1:C3'},
            deleteColumnPart: {refused: true, ref: 'B1:C3'},
          },
        );
      },
    },
    {
      name: 'an edit moving or removing a range whole goes ahead, and so does one through a dynamic array',
      expect(api: CorpusApi, assert: Assert) {
        const {edits} = api.arrayRangeReport();
        assert.deepStrictEqual(
          {
            insertRowAbove: edits.insertRowAbove,
            insertRowBelow: edits.insertRowBelow,
            deleteWholeRange: edits.deleteWholeRange,
            insertColumnLeft: edits.insertColumnLeft,
            dynamicInsertInside: edits.dynamicInsertInside,
          },
          {
            insertRowAbove: {refused: false, ref: 'B2:B4'},
            insertRowBelow: {refused: false, ref: 'B1:B3'},
            deleteWholeRange: {refused: false, ref: null},
            insertColumnLeft: {refused: false, ref: 'C1:D3'},
            dynamicInsertInside: {refused: false, ref: 'B1:B4'},
          },
        );
      },
    },
    {
      name: 'a file whose range holds another formula reads that array formula as a plain one',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepStrictEqual(api.arrayRangeReport().foreign, {
          B1: {kind: 'formula', ref: null, dynamic: false, formula: 'A1:A3*2', result: null},
          rewrites: true,
        });
      },
    },
  ],
} satisfies Case;
