// Cluster: formulas
//
// Real-world scenario: a producer that does not calculate writes each formula with an empty `<v>`, or
// with none, and keeps whatever `t` its template had: `<c t="b"><f>TRUE</f><v></v></c>`, an error type
// over an empty value, a number with no digits. Excel shows the cached result until it recalculates,
// so what it shows for an empty one is the reading to match, and a reader that decoded by `t` alone
// invented a result: `FALSE` for the boolean, which a save then wrote back as a real FALSE.
//
// Excel Desktop was asked with calculation set to manual, so it shows the cache rather than recomputing
// it, and a control cell whose stale cached 5 stayed 5 proves that it did. Every empty or absent cached
// result showed as the same empty string, whatever the type: not FALSE, not an error, not 0. The
// observation is `test/corpus/fixtures/excel-oracle/formula-empty-cached-result.json`. Excel draws no
// line between an empty cached result and none, so neither does the reader, with one exception kept
// from plain cells: under `t="str"` a present `<v>` is text, the empty string.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'formula-empty-cached-result/cells.xlsx';

export default {
  id: 'a-formula-with-an-empty-cached-result-has-none',
  provenance: {
    source: 'excel-desktop-verification',
    ref: 'test/corpus/fixtures/excel-oracle/formula-empty-cached-result.json',
  },
  cluster: 'formulas',
  description:
    'A formula whose cached `<v>` is empty reads as a formula with no cached result under a boolean, ' +
    'error or numeric type, as one with no `<v>` at all does, because Excel shows both the same; only ' +
    'a string formula keeps the empty string. A formula with a cached value still carries it.',

  behavior: [
    {
      name: 'an empty cached result under a boolean, error or number type is no cached result',
      async expect(api: CorpusApi, assert: Assert) {
        const cells = await api.readFixtureCells(FIXTURE, ['B1', 'C1', 'F1']);
        assert.deepStrictEqual(cells.B1.value, {formula: 'TRUE'}, 'B1 t="b" is not FALSE');
        assert.deepStrictEqual(cells.C1.value, {formula: '1/0'}, 'C1 t="e"');
        assert.deepStrictEqual(cells.F1.value, {formula: '1+1'}, 'F1 is not 0');
      },
    },
    {
      name: 'a formula with no cached value reads the same way',
      async expect(api: CorpusApi, assert: Assert) {
        const {E1} = await api.readFixtureCells(FIXTURE, ['E1']);
        assert.deepStrictEqual(E1.value, {formula: 'TRUE'});
      },
    },
    {
      name: 'a string formula keeps the empty string, and a cached value is still read',
      async expect(api: CorpusApi, assert: Assert) {
        const {A1, D1} = await api.readFixtureCells(FIXTURE, ['A1', 'D1']);
        assert.deepStrictEqual(D1.value, {formula: '"x"', result: ''});
        assert.deepStrictEqual(A1.value, {formula: '1+1', result: 5});
      },
    },
  ],
} satisfies Case;
