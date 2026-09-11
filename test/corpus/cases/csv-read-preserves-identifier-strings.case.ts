// Cluster: csv
//
// Real-world scenario: reading a CSV, fields that merely *resemble* a date must not be
// silently turned into Date values. Identifiers and codes like "2020-00001" (a padded ID),
// "1-3" (an inventory code), or "3-4" (a range label) are text and must stay text:
// coercing them corrupts the data and can crash downstream code. Genuinely numeric fields
// should still become numbers, and genuine ISO dates should still become dates. The reader
// must coerce conservatively: only clear numbers and strictly-formatted dates convert. A clear
// number is one a double holds exactly, and a strictly-formatted date names a real calendar day.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'csv-read-preserves-identifier-strings',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 711},
  cluster: 'csv',
  description:
    'CSV read coerces conservatively: an identifier or code that resembles a date ' +
    '("2020-00001", "1-3") is preserved as a string, while a clearly numeric field becomes ' +
    'a number and a strictly-formatted ISO date becomes a Date.',

  behavior: [
    {
      name: 'identifier-like strings that resemble dates are preserved as text',
      async expect(api: CorpusApi, assert: Assert) {
        const {ok, rows} = await api.csvRead({csv: '2020-00001,1-3,3-4', options: {}});
        assert.ok(ok, 'the read must succeed');
        assert.deepStrictEqual(
          rows[0],
          ['2020-00001', '1-3', '3-4'],
          'padded IDs and dash-codes stay strings, not coerced to dates',
        );
      },
    },
    {
      name: 'a clearly numeric field is read as a number',
      async expect(api: CorpusApi, assert: Assert) {
        const {rows} = await api.csvRead({csv: '123,45.6', options: {}});
        assert.deepStrictEqual(rows[0], [123, 45.6], 'numeric fields coerce to numbers');
      },
    },
    {
      name: 'a strictly-formatted ISO date field is read as a Date',
      async expect(api: CorpusApi, assert: Assert) {
        const {rows} = await api.csvRead({csv: '2018-01-05', options: {}});
        const cell = rows[0]![0];
        assert.ok(
          cell && typeof cell === 'object' && cell.date,
          `a real date should coerce; got ${JSON.stringify(cell)}`,
        );
        assert.match(
          cell.date,
          /^2018-01-0[45]T/,
          'the parsed date is Jan 5 2018 (modulo timezone)',
        );
      },
    },
    {
      name: 'an ISO-shaped field naming no calendar day stays text instead of rolling over',
      async expect(api: CorpusApi, assert: Assert) {
        const {rows} = await api.csvRead({csv: '2024-02-30,2023-04-31', options: {}});
        assert.deepStrictEqual(
          rows[0],
          ['2024-02-30', '2023-04-31'],
          'Feb 30 and Apr 31 are not dates, so they must not become March 1 and May 1',
        );
      },
    },
    {
      name: 'a year below 100 is that year, not 1900 plus it',
      async expect(api: CorpusApi, assert: Assert) {
        const {rows} = await api.csvRead({csv: '0099-01-01', options: {}});
        const cell = rows[0]![0];
        assert.ok(
          cell && typeof cell === 'object' && typeof cell.date === 'string',
          `a real date should coerce; got ${JSON.stringify(cell)}`,
        );
        assert.match(cell.date, /^0099-01-01T/, 'year 99 must not read as 1999');
      },
    },
    {
      name: 'a number with more significant digits than a double holds stays text',
      async expect(api: CorpusApi, assert: Assert) {
        const {rows} = await api.csvRead({
          csv: '3.14159265358979323846,4111111111111111',
          options: {},
        });
        assert.deepStrictEqual(
          rows[0],
          ['3.14159265358979323846', '4111111111111111'],
          'coercing either to a number would drop digits',
        );
      },
    },
  ],
} satisfies Case;
