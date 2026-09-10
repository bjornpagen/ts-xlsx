// Cluster: hyperlinks
//
// Real-world scenario: a report template carries a link over a merged banner, which Excel stores as one
// `<hyperlink ref="D1:H1">` whose destination and label sit on the merge's top-left cell, and a fill job
// inserts a header row above it. The merge moved down and so did the cell holding the link, but the
// link's clickable range stayed on the old row: the saved file put the hyperlink over cells the banner
// had left, and a re-read found an empty link above the banner and plain text inside it.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'ranged-hyperlink-moves-with-an-inserted-row',
  cluster: 'hyperlinks',
  provenance: {source: 'audit'},
  description:
    'A hyperlink whose clickable range spans a merged banner keeps that range on the banner after a ' +
    'row is inserted above it: the written hyperlink covers the moved merge, and a re-read finds the ' +
    "link, range included, on the merge's new top-left cell.",

  behavior: [
    {
      name: 'the written hyperlink covers the merge at its new position',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.rangedHyperlinkAfterInsertRow();
        assert.deepEqual(report.merges, ['D2:H2']);
        assert.equal(report.linkRef, 'D2:H2');
      },
    },
    {
      name: 'a re-read finds the link and its range on the banner, not above it',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(api.rangedHyperlinkAfterInsertRow().link, {
          hyperlink: 'https://example.com/',
          text: 'go',
          range: 'D2:H2',
        });
      },
    },
  ],
} satisfies Case;
