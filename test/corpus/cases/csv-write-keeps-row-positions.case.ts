// Cluster: csv
//
// Real-world scenario: a sheet with a blank spacer row between a header block and the data, or a
// report whose first rows are left empty for a title that was never typed. CSV has no row numbers, so
// the only way to say "this is row 3" is to be the third line.
//
// The writer emitted one line per row the model yields, and the model yields only rows holding cells
// or formatting. So A1 and A3 went out as `a\nc`, and `c` read back in row 2. A sheet whose only cell
// is B3 went out as `,x`: the column offset kept, the row offset gone. A hidden empty row did produce
// a blank line, so whether the gap survived depended on formatting the CSV cannot show.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'csv-write-keeps-row-positions',
  provenance: {source: 'codec-symmetry-audit'},
  cluster: 'csv',
  description:
    'The CSV writer places row N on line N, so empty rows before or between populated ones survive ' +
    'a write and read back, and formatting-only rows change nothing about the text.',

  behavior: [
    {
      name: 'an empty row between populated rows reads back as the same gap',
      expect(api: CorpusApi, assert: Assert) {
        const result = api.csvRowPositionsRoundTrip({cells: {A1: 'a', A3: 'c'}});
        assert.deepEqual(result.cells, {A1: 'a', A3: 'c'});
      },
    },
    {
      name: 'empty rows above the first populated row keep its row number',
      expect(api: CorpusApi, assert: Assert) {
        const result = api.csvRowPositionsRoundTrip({cells: {B3: 'x'}});
        assert.deepEqual(result.cells, {B3: 'x'});
      },
    },
    {
      name: 'a hidden empty row writes the same text as an unformatted one',
      expect(api: CorpusApi, assert: Assert) {
        const plain = api.csvRowPositionsRoundTrip({cells: {A1: 'a', A3: 'c'}});
        const hidden = api.csvRowPositionsRoundTrip({
          cells: {A1: 'a', A3: 'c'},
          hiddenRows: [2, 6],
        });
        assert.equal(hidden.text, plain.text);
      },
    },
  ],
} satisfies Case;
