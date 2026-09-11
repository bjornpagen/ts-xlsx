// Cluster: types
//
// Real-world scenario: a Strict-mode (ISO/IEC 29500 Strict) xlsx encodes a date cell with type
// `t="d"` and an ISO 8601 value (<v>2024-02-09</v>) rather than the transitional serial number.
// A reader that assumes the transitional encoding treats the ISO text as a numeric serial and
// produces a wildly wrong 1900-epoch date (e.g. 1905-07-16). A Strict-mode date cell must parse
// to the date it literally states, and sibling string cells in the same workbook must still read
// as their text.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'strict-mode-iso8601-date/sample.xlsx';

export default {
  id: 'strict-mode-iso8601-date-parses-correctly',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 2695},
  cluster: 'types',
  description:
    'A Strict-mode date cell (t="d" with an ISO 8601 value) parses to the stated date rather than ' +
    'a spurious 1900-epoch serial, while a string cell in the same workbook still reads as text.',

  behavior: [
    {
      name: 'the ISO 8601 date cell parses to the date it states',
      async expect(api: CorpusApi, assert: Assert) {
        const {A2} = await api.readFixtureCells(FIXTURE, ['A2']);
        assert.strictEqual(A2.type, 'date', 'the cell is typed as a date');
        assert.strictEqual(
          A2.value.date,
          '2024-02-09T00:00:00.000Z',
          `the ISO date must parse to 2024-02-09, not a 1900-epoch serial; got ${JSON.stringify(A2.value)}`,
        );
      },
    },
    {
      name: 'a string cell in the same Strict-mode workbook still reads as text',
      async expect(api: CorpusApi, assert: Assert) {
        const {A1} = await api.readFixtureCells(FIXTURE, ['A1']);
        assert.strictEqual(A1.type, 'string', 'the header cell reads as a string');
        assert.strictEqual(A1.value, 'Date', 'the shared-string value is intact');
      },
    },
    {
      // A Strict package spells the app-properties relationship `extendedProperties`, where
      // Transitional spells it `extended-properties`. A reader that recognises only the second keeps
      // the Strict one as an unmodelled part, and the writer then emits it beside the properties part
      // it generates, at the same path.
      name: 'a Strict-mode workbook writes back out and reads again with the same cells',
      async expect(api: CorpusApi, assert: Assert) {
        const report = api.dateSystemReport(FIXTURE, ['A1', 'A2']);
        assert.deepStrictEqual(
          report.roundtrip,
          report.eager,
          'the re-read cells match the first read',
        );
      },
    },
  ],
} satisfies Case;
