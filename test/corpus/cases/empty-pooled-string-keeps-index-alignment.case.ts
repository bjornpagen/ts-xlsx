// Cluster: xlsx-io
//
// Real-world scenario: `xl/sharedStrings.xml` is an index-aligned table. A `t="s"` cell carries an
// ordinal into it and nothing else, and producers spell an empty pooled string two legal ways:
// `<si><t/></si>` and the self-closing `<si/>`. A reader whose parser fires no close for a
// self-closing element commits no entry for the second form, so the pool it builds is one short and
// every cell indexing past that point resolves to its *neighbour's* string. The same table is where
// an empty `<v/>` on a `t="s"` cell lands: read through a bare `Number()`, the empty string is 0, an
// integer, so the cell resolves to the first pooled string. Excel shows that cell as blank, not as
// an empty string (`test/corpus/fixtures/excel-oracle/typed-cell-without-value.json`).
//
// Both are the worst failure this library has: not a missing value, a different one. A downstream
// consumer sees a plausible string in every cell and has no way to tell it is the wrong one.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'empty-pooled-string-keeps-index-alignment',
  cluster: 'xlsx-io',
  description:
    'An empty shared-string entry written as `<si/>` occupies its slot exactly as `<si><t/></si>` ' +
    'does, so a `t="s"` cell indexing past it resolves to the string actually stored at that ' +
    'ordinal; and a `t="s"` cell whose `<v>` is present but empty holds no value, as Excel shows ' +
    'it, rather than shared string 0.',
  provenance: {source: 'audit'},

  behavior: [
    {
      name: 'both spellings of an empty pooled string decode to the empty string',
      expect(api: CorpusApi, assert: Assert) {
        const {pooled} = api.pooledStringIndexReport();
        assert.strictEqual(pooled[1], '', '`<si/>` is an empty pooled string');
        assert.strictEqual(pooled[2], '', '`<si><t/></si>` is an empty pooled string');
      },
    },
    {
      name: 'a self-closing `<si/>` still occupies its slot, so later indices stay aligned',
      expect(api: CorpusApi, assert: Assert) {
        const {pooled} = api.pooledStringIndexReport();
        assert.strictEqual(pooled[0], 'first', 'the entry before the empty ones is unaffected');
        assert.strictEqual(
          pooled[3],
          'fourth',
          'the cell indexing slot 3 must read slot 3, not the entry an uncommitted `<si/>` shifted into it',
        );
      },
    },
    {
      name: 'a `t="s"` cell with a present-but-empty `<v/>` is blank, not shared string 0',
      expect(api: CorpusApi, assert: Assert) {
        const {emptyValueCell} = api.pooledStringIndexReport();
        assert.strictEqual(
          emptyValueCell,
          null,
          'an empty `<v/>` carries no ordinal, so the cell holds no pooled string',
        );
      },
    },
  ],
} satisfies Case;
