// Cluster: security
//
// Real-world scenario: what sits between a tag's name and its `>` is the file's choice, and the
// attribute scan runs on every open tag of every part. A scan that backtracks over a run of name
// characters no `=` follows pays for that run once per character in it, so a single junk token costs the
// square of its length. A megabyte of one repeated letter compresses to about a kilobyte, well inside
// the inflate ceiling, and cost minutes of CPU inside either reader: a CPU bomb that the zip-bomb guard
// cannot see, because after inflation the payload really is only a megabyte.

import type {Assert, Case, CorpusApi} from '../case.ts';

const TOKEN_LENGTH = 1_000_000;

// Both readers take this sheet in a few tens of milliseconds once the scan is linear, and a 64K-character
// token alone took seconds before, so the budget sits more than an order of magnitude clear of each side.
const TIME_BUDGET_MS = 1_500;

export default {
  id: 'a-megabyte-attribute-token-is-scanned-in-linear-time',
  provenance: {source: 'hostile-input-probe'},
  cluster: 'security',
  description:
    'A cell whose open tag carries a megabyte-long junk token with no "=" is read in bounded time by ' +
    'both the buffered and the streaming reader, and the attributes around the token still read.',

  behavior: [
    {
      name: 'the buffered reader finishes the sheet within its budget',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.junkAttributeTokenReport(TOKEN_LENGTH);
        assert.ok(
          report.bufferedMs < TIME_BUDGET_MS,
          `a ${report.tokenLength}-character token in ${report.zippedBytes} zipped bytes took ` +
            `${report.bufferedMs.toFixed(0)}ms to read`,
        );
      },
    },
    {
      name: 'the streaming reader is bounded on the same terms',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.junkAttributeTokenReport(TOKEN_LENGTH);
        assert.ok(
          report.streamingMs < TIME_BUDGET_MS,
          `streaming the same sheet took ${report.streamingMs.toFixed(0)}ms`,
        );
      },
    },
    {
      name: 'and the token costs nothing else: the cell it sits on reads in both',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.junkAttributeTokenReport(TOKEN_LENGTH);
        assert.deepEqual(
          {buffered: report.buffered, streamed: report.streamed},
          {buffered: 42, streamed: 42},
        );
      },
    },
  ],
} satisfies Case;
