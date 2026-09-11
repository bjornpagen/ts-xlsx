// Cluster: tables
//
// Real-world scenario: a sales report pivots by year. The Year column is numbers, so the pivot cache's
// catalogue of years is a list of numbers, and Excel reads that catalogue against the description the
// cache field gives of it. A field that does not say it holds numbers is, by the schema's defaults, a
// field of strings, and a catalogue of `<n>` items under that description made Excel open the package
// with its repair prompt. Pivoting by Name, Region and Quarter opened clean, which is why it went
// unnoticed: it only took one numeric column on an axis.
//
// What each mix looks like was put to Excel Desktop, which built a pivot over columns of numbers, numbers
// with a blank, strings mixed with numbers, strings with a blank, decimals and plain strings, and saved
// it. Its attributes are recorded in `test/corpus/fixtures/excel-oracle/pivot-cache-field-kinds.json`,
// with the open verdicts before and after the writer described fields the same way.

import type {Assert, Case, CorpusApi} from '../case.ts';

const SOURCE = {
  Num: [1, 2, 3, 4, 2],
  NumBlank: [1, null, 3, 4, 1],
  Mixed: ['a', 2, 'b', 4, 'a'],
  StrBlank: ['a', null, 'b', 'c', 'a'],
  Float: [1.5, 2, 3.25, 4, 2],
  Str: ['x', 'y', 'z', 'x', 'y'],
};

// Copied from the sidecar: what Excel wrote for each column.
const EXCEL_WROTE: Record<string, string> = {
  Num: 'containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="1" maxValue="4"',
  NumBlank:
    'containsString="0" containsBlank="1" containsNumber="1" containsInteger="1" minValue="1" maxValue="4"',
  Mixed: 'containsMixedTypes="1" containsNumber="1" containsInteger="1" minValue="2" maxValue="4"',
  StrBlank: 'containsBlank="1"',
  Float:
    'containsSemiMixedTypes="0" containsString="0" containsNumber="1" minValue="1.5" maxValue="4"',
  Str: '',
};

const DISTINCT: Record<string, number> = {Num: 4, NumBlank: 4, Mixed: 4, StrBlank: 4, Float: 4};

export default {
  id: 'a-pivot-cache-field-says-what-kinds-of-value-it-holds',
  provenance: {
    source: 'excel-desktop-verification',
    ref: 'test/corpus/fixtures/excel-oracle/pivot-cache-field-kinds.json',
  },
  cluster: 'tables',
  description:
    'Every pivot cache field describes the kinds of value it holds (strings, numbers, blanks, their ' +
    'range and integrality) with the attributes Excel writes for that mix, whether the field carries a ' +
    'catalogue on an axis or stores its values inline, so a pivot over a numeric axis opens clean.',

  behavior: [
    {
      name: 'a field on an axis describes its catalogue as Excel does, then counts it',
      expect(api: CorpusApi, assert: Assert) {
        const kinds = api.pivotCacheFieldKinds({
          source: SOURCE,
          rows: ['Num', 'NumBlank', 'Mixed', 'StrBlank', 'Float'],
          columns: ['Str'],
        });
        for (const [name, count] of Object.entries(DISTINCT)) {
          const excel = EXCEL_WROTE[name] ?? '';
          assert.strictEqual(kinds[name], `${excel} count="${count}"`, name);
        }
      },
    },
    {
      name: 'a field no axis uses describes its inline values the same way, with no count',
      expect(api: CorpusApi, assert: Assert) {
        const kinds = api.pivotCacheFieldKinds({source: SOURCE, rows: ['Str'], columns: ['Mixed']});
        for (const name of ['Num', 'NumBlank', 'StrBlank', 'Float']) {
          assert.strictEqual(kinds[name], EXCEL_WROTE[name], name);
        }
      },
    },
    {
      name: 'a field of plain strings says nothing, as the defaults already describe it',
      expect(api: CorpusApi, assert: Assert) {
        const kinds = api.pivotCacheFieldKinds({source: SOURCE, rows: ['Num'], columns: ['Mixed']});
        assert.strictEqual(kinds.Str, EXCEL_WROTE.Str);
      },
    },
  ],
} satisfies Case;
