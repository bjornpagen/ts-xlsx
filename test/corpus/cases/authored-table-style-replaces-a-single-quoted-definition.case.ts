// Cluster: styles
//
// Real-world scenario: a template from a producer that quotes attributes with apostrophes carries a
// custom table style, `<tableStyle name='Harbour' ...>`, and a report generator re-registers a style of
// the same name to restyle it. An authored style is meant to replace the definition it shares a name
// with, but the preserved definition was matched by a pattern that only recognised double quotes, so
// the saved stylesheet declared `Harbour` twice and which one a table used was up to the reader.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'authored-table-style-replaces-a-single-quoted-definition',
  cluster: 'styles',
  provenance: {source: 'audit'},
  description:
    'A custom table style a file declares with a single-quoted name is replaced, not duplicated, by an ' +
    'authored style of the same name: the rewritten stylesheet defines the name once and counts one.',

  behavior: [
    {
      name: 'the rewritten stylesheet defines the name once',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(api.singleQuotedTableStyleOverrideReport().definitions, ['Harbour']);
      },
    },
    {
      name: 'the declared count matches the definitions it holds',
      expect(api: CorpusApi, assert: Assert) {
        assert.equal(api.singleQuotedTableStyleOverrideReport().declaredCount, 1);
      },
    },
    {
      name: 'the definition written is the authored one',
      expect(api: CorpusApi, assert: Assert) {
        assert.equal(api.singleQuotedTableStyleOverrideReport().preservedDefinitionWritten, false);
      },
    },
  ],
} satisfies Case;
