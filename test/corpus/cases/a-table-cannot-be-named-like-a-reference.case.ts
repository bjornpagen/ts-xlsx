// Cluster: tables
//
// Real-world scenario: someone names their tables `T1`, `T2`, `T3`, or `R` and `C` for rows and
// columns, or `TRUE`. Each is a valid identifier by the character rules, and each is a name a
// formula reads as something else: a cell, a relative R1C1 reference, a boolean. Excel refuses all of
// them as table names, and a package carrying one opens with a prompt to repair it. The rule's edges
// are Excel's rather than the specification's, so they were put to Excel Desktop and recorded in
// `test/corpus/fixtures/excel-oracle/names-that-read-as-references.json`: a cell only counts while it
// is on the grid (`XFE1` is accepted), and an R1C1 number only counts when a letter or underscore
// follows it (`R1X` is refused, `R1.5` is not).
//
// The table is refused where it is authored, so a caller hears about it at the call rather than from
// a user whose Excel offered to repair the file.

import type {Assert, Case, CorpusApi} from '../case.ts';

const REFUSED_BY_EXCEL = ['T1', 'c1', 'A01', 'XFD1048576', 'R', 'r', 'RC', 'R1C1', 'R1X', 'TRUE'];
const ACCEPTED_BY_EXCEL = ['Table1', 'CC', 'XFE1', 'A0', 'R0', 'R1.5', 'RCX', 'TRUE1'];

export default {
  id: 'a-table-cannot-be-named-like-a-reference',
  provenance: {
    source: 'excel-desktop-verification',
    ref: 'test/corpus/fixtures/excel-oracle/names-that-read-as-references.json',
  },
  cluster: 'tables',
  description:
    'A table name that reads as an A1 cell on the grid, an R1C1 reference, or a boolean is refused ' +
    'when the table is authored, because Excel does not accept it and offers to repair a package ' +
    'carrying one; the names around those edges that Excel accepts are accepted.',

  behavior: [
    {
      name: 'every name Excel refused as a table name is refused at authoring',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.tableNameRefusals(REFUSED_BY_EXCEL);
        for (const name of REFUSED_BY_EXCEL) {
          assert.equal(report[name]?.refused, true, `${name} is refused`);
        }
      },
    },
    {
      name: 'every name Excel accepted beside those edges is accepted',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.tableNameRefusals(ACCEPTED_BY_EXCEL);
        for (const name of ACCEPTED_BY_EXCEL) {
          assert.equal(report[name]?.error, null, `${name} is accepted`);
        }
      },
    },
  ],
} satisfies Case;
