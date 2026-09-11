// Cluster: csv
//
// Real-world scenario: a CSV of people's heights, product sizes or survey free text, written by a
// tool that quotes only the fields that need it. `5'10"` and `12" pizza` hold a double quote that
// does not start the field, so the writer left them bare. Excel reads that quote as the character it
// is. A reader that opens a quoted field at any quote reads everything from there to the next quote
// in the file as one field: delimiters and row breaks included, so two rows collapse into one with no
// error, and the row count and every field after the quote are wrong.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'csv-read-quote-inside-unquoted-field-is-literal',
  cluster: 'csv',
  description:
    'A double quote that is not the first character of a CSV field is a literal character, so ' +
    '`5\'10"` reads as that text and the delimiters and row breaks after it still separate fields ' +
    'and rows; a quote at the start of a field still opens a quoted field.',
  provenance: {source: 'audit'},

  behavior: [
    {
      name: 'a mid-field quote does not swallow the fields and rows after it',
      expect(api: CorpusApi, assert: Assert) {
        const {ok, rows} = api.csvRead({csv: `John,5'10",tall\nJane,5'2",short`});
        assert.strictEqual(ok, true);
        assert.deepStrictEqual(rows, [
          ['John', `5'10"`, 'tall'],
          ['Jane', `5'2"`, 'short'],
        ]);
      },
    },
    {
      name: 'a trailing quote on a field is kept and the row still ends',
      expect(api: CorpusApi, assert: Assert) {
        const {ok, rows} = api.csvRead({csv: 'a,12"\nb,14"'});
        assert.strictEqual(ok, true);
        assert.deepStrictEqual(rows, [
          ['a', '12"'],
          ['b', '14"'],
        ]);
      },
    },
    {
      name: 'a quote opening a field still quotes it',
      expect(api: CorpusApi, assert: Assert) {
        const {ok, rows} = api.csvRead({csv: '"x,y",z'});
        assert.strictEqual(ok, true);
        assert.deepStrictEqual(rows, [['x,y', 'z']]);
      },
    },
  ],
} satisfies Case;
