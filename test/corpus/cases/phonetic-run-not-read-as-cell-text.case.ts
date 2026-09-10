// Cluster: xlsx-io
//
// Real-world scenario: Japanese Excel stores the furigana reading of a string beside its base text, as
// a phonetic run (`<rPh>`) inside the same string container: a shared-strings `<si>`, an inline
// `<is>`, or a note's `<text>`. The phonetic run has a `<t>` of its own, and a reader that gathers every
// `<t>` in the container appends the reading to the text it annotates, so a cell holding `漢字` reads
// as `漢字かんじ`, and saving the workbook writes that corruption back as the cell's value. Every
// furigana-bearing cell in a Japanese workbook is affected.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'phonetic-run-not-read-as-cell-text',
  provenance: {source: 'audit'},
  cluster: 'xlsx-io',
  description:
    'A phonetic run stored beside a string is not part of that string: a pooled string, a rich ' +
    "string's runs, an inline string and a note body each read back as their base text alone, and " +
    'a save writes none of the phonetic text into the value.',

  behavior: [
    {
      name: 'a pooled string reads its base text without the phonetic run',
      expect(api: CorpusApi, assert: Assert) {
        assert.equal(api.phoneticRunReport().read.pooledPlain, '漢字');
      },
    },
    {
      name: 'an inline string reads its base text without the phonetic run',
      expect(api: CorpusApi, assert: Assert) {
        assert.equal(api.phoneticRunReport().read.inline, '漢字');
      },
    },
    {
      name: "a rich string's runs carry only the base text",
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(api.phoneticRunReport().read.pooledRichRuns, ['漢字']);
      },
    },
    {
      name: 'a note reads its own text without the phonetic run',
      expect(api: CorpusApi, assert: Assert) {
        assert.equal(api.phoneticRunReport().read.note, 'note');
      },
    },
    {
      name: 'a save writes none of the phonetic text into the values',
      expect(api: CorpusApi, assert: Assert) {
        // Pinned to the base text rather than to the first reading: a reader that corrupts both the
        // same way would otherwise pass by agreeing with itself.
        assert.deepEqual(api.phoneticRunReport().rewritten, {
          pooledPlain: '漢字',
          pooledRichRuns: ['漢字'],
          inline: '漢字',
          note: 'note',
        });
      },
    },
  ],
} satisfies Case;
