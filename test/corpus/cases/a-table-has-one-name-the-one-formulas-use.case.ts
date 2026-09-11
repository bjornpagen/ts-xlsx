// Cluster: tables
//
// Real-world scenario: a table part carries two name attributes, `name` and `displayName`, and a model
// that treated them as two things let a caller give a table an identifier for formulas and a friendlier
// label with spaces for people. Excel Desktop disagrees on both counts, recorded in
// `test/corpus/fixtures/excel-oracle/table-display-name.json`. A `displayName` with a space makes Excel
// offer to repair the package, and over a part whose `name` is `Internal` and whose `displayName` is
// `Shown`, `SUM(Shown[h])` computes while `SUM(Internal[h])` is `#REF!`. The `displayName` is the table's
// name, the one structured references use and the one Excel shows; `name` has no meaning of its own.
//
// So a table has one name. It is written as both attributes, and read from `displayName`, with `name`
// standing in only for a part that leaves `displayName` out.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'a-table-has-one-name-the-one-formulas-use',
  provenance: {
    source: 'excel-desktop-verification',
    ref: 'test/corpus/fixtures/excel-oracle/table-display-name.json',
  },
  cluster: 'tables',
  description:
    "A table's single name is written as both the part's `name` and its `displayName`, and a part " +
    'whose two attributes differ reads under its `displayName`, the name Excel resolves structured ' +
    'references against.',

  behavior: [
    {
      name: "a table's name is written as both of the part's name attributes",
      expect(api: CorpusApi, assert: Assert) {
        assert.deepStrictEqual(api.tableNamesReport().written, {
          name: 'Sales',
          displayName: 'Sales',
        });
      },
    },
    {
      name: 'a part whose two names differ reads under its displayName',
      expect(api: CorpusApi, assert: Assert) {
        assert.strictEqual(api.tableNamesReport().readDiffering, 'Shown');
      },
    },
    {
      name: 'a part with no displayName reads under its name',
      expect(api: CorpusApi, assert: Assert) {
        assert.strictEqual(api.tableNamesReport().readNameOnly, 'OnlyName');
      },
    },
  ],
} satisfies Case;
