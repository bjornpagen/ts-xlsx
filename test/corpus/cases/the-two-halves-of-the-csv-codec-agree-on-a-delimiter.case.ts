// Cluster: csv
//
// Real-world scenario: a caller picks a delimiter -- a semicolon for a European locale, a tab for a
// paste into a spreadsheet, and occasionally something the format cannot carry, like `||` from a
// pipe-delimited convention or an empty string from an unset config value. The reader refused a
// multi-character delimiter, because its character-scan parser cannot honour one. The writer checked
// nothing at all, in a module whose header is entirely about lossless round-tripping.
//
// So `readCsv(writeCsv(wb, {delimiter: '||'}))` threw on text this same codec had just produced. And
// the empty string was worse than a throw: `quoteField` asks `field.includes(delimiter)`, which is
// true of every field, so every field was quoted and the file went out with no separators in it --
// a file that parses cleanly as a single column and loses the shape of the data with nothing
// reported anywhere.
//
// The same audit found the single characters the length check let through. A quote, CR or LF is a
// single character the parser already gives a meaning to, so as a delimiter it is two things at once.
// And the row delimiter was not checked at all, nor was a field containing a custom one quoted: under
// `rowDelimiter: '|'` the fields `a|b` and `c` went out as `a|b,c`.
//
// The rule this locks: one validator, called from both entry points. Whatever the writer accepts,
// the reader reads back.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'the-two-halves-of-the-csv-codec-agree-on-a-delimiter',
  provenance: {source: 'codec-symmetry-audit'},
  cluster: 'csv',
  description:
    'A delimiter the CSV reader cannot honour is refused by the writer as well, with the same ' +
    'RangeError, so every delimiter the writer accepts round-trips through the reader.',

  behavior: [
    {
      name: 'every delimiter the writer accepts reads back as the rows that were written',
      async expect(api: CorpusApi, assert: Assert) {
        const written = (await api.csvDelimiterAgreement()).filter((row) => row.wroteOk);
        assert.equal(written.length, 3, 'comma, semicolon and tab');
        for (const row of written) {
          assert.equal(row.readOk, true, `${row.delimiter}: reads back (${row.readError})`);
          assert.deepEqual(
            row.rows,
            [
              ['a', 'b'],
              ['c', 'd'],
            ],
            `${row.delimiter}: the grid survives the round trip`,
          );
        }
      },
    },
    {
      name: 'a delimiter the reader cannot honour is refused by the writer, not silently emitted',
      async expect(api: CorpusApi, assert: Assert) {
        const rows = new Map(
          (await api.csvDelimiterAgreement()).map((row) => [row.delimiter, row]),
        );
        for (const delimiter of ['||', '']) {
          const row = rows.get(delimiter);
          assert.equal(row?.wroteOk, false, `${JSON.stringify(delimiter)}: the write is refused`);
          assert.match(
            String(row?.writeError),
            /single character/,
            'and says what a delimiter may be',
          );
        }
      },
    },
    {
      name: 'a quote, CR or LF is refused as a delimiter, because the reader gives each its own meaning',
      async expect(api: CorpusApi, assert: Assert) {
        // With `"` as the delimiter, `a"b"c` read back as the single field `abc`: the character is a
        // separator and a quote at once, and the parser can only take it as one of them.
        const rows = new Map(
          (await api.csvDelimiterAgreement()).map((row) => [row.delimiter, row]),
        );
        for (const delimiter of ['"', '\r', '\n']) {
          const row = rows.get(delimiter);
          assert.equal(row?.wroteOk, false, `${JSON.stringify(delimiter)}: the write is refused`);
          assert.match(String(row?.writeError), /reserved/, 'and says why');
        }
      },
    },
    {
      name: 'a row delimiter no consumer could split on is refused rather than written',
      expect(api: CorpusApi, assert: Assert) {
        const spec = {rows: [['a', 'b']]};
        for (const rowDelimiter of ['', ',', '"']) {
          const written = api.csvWrite({spec, options: {formatterOptions: {rowDelimiter}}});
          assert.equal(written.ok, false, `${JSON.stringify(rowDelimiter)}: the write is refused`);
        }
      },
    },
    {
      name: 'a field containing a custom row delimiter is quoted, so it is not split into two rows',
      expect(api: CorpusApi, assert: Assert) {
        const written = api.csvWrite({
          spec: {rows: [['a|b', 'c'], ['d']]},
          options: {formatterOptions: {rowDelimiter: '|'}},
        });
        assert.equal(written.text, '"a|b",c|d');
      },
    },
    {
      name: 'the delimiter actually separates: it is not swallowed by quoting',
      async expect(api: CorpusApi, assert: Assert) {
        const rows = new Map(
          (await api.csvDelimiterAgreement()).map((row) => [row.delimiter, row]),
        );
        assert.equal(rows.get(';')?.text, 'a;b\nc;d');
        assert.equal(rows.get('\t')?.text, 'a\tb\nc\td');
      },
    },
  ],
} satisfies Case;
