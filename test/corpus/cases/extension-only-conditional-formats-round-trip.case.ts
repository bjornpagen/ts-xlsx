// Cluster: conditional-formatting
//
// Real-world scenario: Excel stores a conditional format in the worksheet's 2009 extension
// (`<x14:conditionalFormatting>` inside `<extLst>`) rather than the classic `<conditionalFormatting>`
// element whenever the classic element cannot say it or Excel 2007 must not see it: a rule whose
// formula reaches another sheet, an icon set from the families the extension added (stars, triangles,
// boxes), an icon set with its own icon per threshold, a colour scale anchored on another sheet. Such a
// rule keeps its formula in `<xm:f>` and its look inline in `<x14:dxf>`. Reading and saving the workbook
// must keep every one of them, in the form Excel chose, or a template loses its cross-sheet highlights
// on the first save.
//
// The fixture is a workbook Excel 16.0 saved (see `author.ps1` beside it) with five such rules next to
// three classic ones: a same-sheet expression, an icon set in reverse order with its values hidden and a
// strict `>` threshold, and a data bar whose extra facets sit in the extension.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'extension-only-conditional-formats/excel-saved.xlsx';

export default {
  id: 'extension-only-conditional-formats-round-trip',
  cluster: 'conditional-formatting',
  provenance: {source: 'excel-desktop', excel: '16.0 build 20326'},
  description:
    'Conditional formats Excel stores only in the worksheet extension (cross-sheet formulas, 2009 icon ' +
    'families, custom icons) read onto their ranges and are written back to the extension as Excel ' +
    'saved them, beside classic rules that stay classic.',
  behavior: [
    {
      name: 'every rule reads onto its own range, the ones only the extension holds included',
      expect(api: CorpusApi, assert: Assert) {
        const {read} = api.conditionalFormatRulesAsStored(FIXTURE);
        assert.deepEqual(read, [
          {ref: 'A1:A6', rules: [{type: 'expression', formulae: ['A1>Data!$A$1']}]},
          {ref: 'B1:B6', rules: [{type: 'dataBar', formulae: []}]},
          {ref: 'C1:C6', rules: [{type: 'cellIs', formulae: ['Data!$A$1']}]},
          {ref: 'D1:D6', rules: [{type: 'expression', formulae: ['D1>4']}]},
          {ref: 'E1:E6', rules: [{type: 'iconSet', formulae: []}]},
          {ref: 'F1:F6', rules: [{type: 'iconSet', formulae: []}]},
          {ref: 'G1:G6', rules: [{type: 'iconSet', formulae: []}]},
          {ref: 'H1:H6', rules: [{type: 'colorScale', formulae: []}]},
        ]);
      },
    },
    {
      name: 'a save writes each extension-only rule back to the extension as Excel saved it, style and icons included',
      expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = api.conditionalFormatRulesAsStored(FIXTURE);
        const standalone = (rules: typeof source.extension) =>
          rules.filter((entry) => !entry.linksAClassicRule);
        assert.strictEqual(standalone(source.extension).length, 5, 'the fixture holds five');
        assert.deepEqual(standalone(rewritten.extension), standalone(source.extension));
      },
    },
    {
      name: 'a save leaves the classic rules classic, with their icon order, hidden values and strict threshold',
      expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = api.conditionalFormatRulesAsStored(FIXTURE);
        assert.strictEqual(source.classic.length, 3, 'the fixture holds three');
        assert.deepEqual(rewritten.classic, source.classic);
      },
    },
  ],
} satisfies Case;
