// Cluster: tables
//
// Real-world scenario: an Excel table can hold a CALCULATED column, one where every body cell shares
// a single formula. Excel records that formula once, on the column definition, as a
// `<calculatedColumnFormula>` child of the `<tableColumn>` element (not as a formula on each cell).
// A reader must consume that nested element and keep enumerating the remaining columns. The legacy
// table-column reader instead loses its place at the nested element: it stops collecting columns after
// the first calculated one, so the table's column list is truncated, and reconciling the sheet's
// autoFilter against the short list then dereferences a column that is not there
// ("Cannot set properties of undefined (setting 'filterButton')"), the whole workbook fails to load.
//
// The fixture is a three-column table (Qty, Double, Label) whose middle column, Double, is calculated
// (`Sales[Qty]*2`). A correct reader loads all three columns without error, and keeps the formula: it is
// what Excel fills into a row added to the table, so a save that dropped it would leave the column's
// cells as they are while every new row came in blank.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'table-with-calculated-column-loads-all-columns/calculated-column.xlsx';

export default {
  id: 'table-with-calculated-column-loads-all-columns',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'A table that declares a calculated column (a <calculatedColumnFormula> on the column definition) ' +
    'loads without crashing and preserves every column, rather than truncating the column list at the ' +
    'first calculated column and then dereferencing a missing column during autoFilter reconciliation.',

  behavior: [
    {
      name: 'a workbook whose table has a calculated column loads without throwing',
      async expect(api: CorpusApi, assert: Assert) {
        const {loaded, error} = await api.loadFixtureTableColumns(FIXTURE, 'Sales');
        assert.strictEqual(error, null, 'loading must not crash on the calculated-column table');
        assert.strictEqual(loaded, true, 'the workbook loads');
      },
    },
    {
      name: 'the calculated column does not truncate the table: all three columns survive',
      async expect(api: CorpusApi, assert: Assert) {
        const {columnCount, columnNames} = await api.loadFixtureTableColumns(FIXTURE, 'Sales');
        assert.strictEqual(
          columnCount,
          3,
          'every column, including those after the calculated one, is read',
        );
        assert.deepStrictEqual(
          columnNames,
          ['Qty', 'Double', 'Label'],
          'columns keep their names and order',
        );
      },
    },
    {
      name: 'the calculated column keeps its formula through a load and save',
      expect(api: CorpusApi, assert: Assert) {
        const [table] = api.roundtripFixtureTableXml(FIXTURE).tables;
        assert.deepStrictEqual(table?.source.calculatedColumnFormulas, {Double: 'Sales[Qty]*2'});
        assert.deepStrictEqual(table?.rewritten?.calculatedColumnFormulas, {
          Double: 'Sales[Qty]*2',
        });
      },
    },
  ],
} satisfies Case;
