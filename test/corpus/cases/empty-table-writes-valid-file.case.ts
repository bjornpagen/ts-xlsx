// Cluster: tables
//
// Real-world scenario: a report defines a table with column headers but, on an empty
// result set, writes it with zero data rows. The produced xl/tables/tableN.xml must
// still be well-formed: its ref must span the full header row across all declared
// columns (not collapse to a single cell), and any autoFilter must stay consistent
// with that ref; otherwise Excel strips or repairs the table. Adding one data row
// must keep working (guarding the known-good path).

import type {Assert, Case, CorpusApi} from '../case.ts';

const EMPTY = {
  sheets: [{name: 'S', tables: [{name: 'Empty', ref: 'A1', headers: ['Alpha', 'Beta'], rows: []}]}],
};
const ONE_ROW = {
  sheets: [
    {
      name: 'S',
      tables: [{name: 'OneRow', ref: 'A1', headers: ['Alpha', 'Beta'], rows: [['x', 'y']]}],
    },
  ],
};

export default {
  id: 'empty-table-writes-valid-file',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1122},
  cluster: 'tables',
  description:
    'A table written with header columns and zero data rows must produce a ' +
    'well-formed table part whose ref spans the full header and whose autoFilter ' +
    '(if present) is consistent with that ref.',

  behavior: [
    {
      name: 'an empty-body table refs the full header row across all columns',
      async expect(api: CorpusApi, assert: Assert) {
        const [table] = (await api.inspectPackage(EMPTY)).tables;
        assert.ok(table, 'no table part written');
        assert.strictEqual(table.ref, 'A1:B1', `expected ref A1:B1, got ${table.ref}`);
        assert.strictEqual(table.columnCount, 2);
      },
    },
    {
      name: 'an empty-body table keeps its autoFilter consistent with its ref',
      async expect(api: CorpusApi, assert: Assert) {
        const [table] = (await api.inspectPackage(EMPTY)).tables;
        assert.ok(table!.xmlWellFormed, 'table XML not well-formed');
        if (table!.autoFilterRef !== null) {
          assert.strictEqual(
            table!.autoFilterRef,
            table!.ref,
            'autoFilter ref inconsistent with table ref',
          );
        }
      },
    },
    {
      name: 'the same table with one data row remains valid',
      async expect(api: CorpusApi, assert: Assert) {
        const [table] = (await api.inspectPackage(ONE_ROW)).tables;
        assert.strictEqual(table!.ref, 'A1:B2', `expected ref A1:B2, got ${table!.ref}`);
        assert.strictEqual(table!.columnCount, 2);
      },
    },
  ],
} satisfies Case;
