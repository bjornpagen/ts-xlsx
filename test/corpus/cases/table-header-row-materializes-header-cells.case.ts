// Cluster: tables
//
// Real-world scenario: code adds a table with a header row and data rows, then writes the
// workbook. Excel refuses to open the result and offers to repair it.
//
// A table part declares its columns as metadata (`<tableColumn name="Alpha"/>`) and a range
// (`ref="A1:B3"`) that *includes* the header row. Excel treats the two as one fact: the
// worksheet cell sitting at each header position must exist and must carry exactly that
// column's name. A table whose declared header row is empty in `sheetData` is structurally
// inconsistent, and Excel takes the repair-on-open path rather than rendering it, verified
// against Excel Desktop, which refuses the package outright and, when forced through
// `xlRepairFile`, rewrites the headers to generic `Column1`/`Column2`.
//
// So writing the column names into the grid is the *library's* job, not the caller's: the
// caller already stated the names once, in the table definition. Requiring them to also set
// A1/B1 by hand makes silent corruption the default outcome of the obvious API call.

import type {Assert, Case, CorpusApi} from '../case.ts';
import type {Untyped} from '../untyped.ts';

const HEADER_TABLE = {
  sheets: [
    {
      name: 'S',
      tables: [{name: 'Header', ref: 'A1', headers: ['Alpha', 'Beta'], rows: [['x', 'y']]}],
    },
  ],
};

// Anchored away from A1 so a fix that hard-codes row 1 instead of using the table's anchor
// still fails this.
const OFFSET_TABLE = {
  sheets: [
    {
      name: 'S',
      tables: [{name: 'Anchored', ref: 'C3', headers: ['Gamma', 'Delta'], rows: [['x', 'y']]}],
    },
  ],
};

const HEADERLESS_TABLE = {
  sheets: [
    {
      name: 'S',
      tables: [
        {
          name: 'Headerless',
          ref: 'A1',
          headers: ['Alpha', 'Beta'],
          rows: [['x', 'y']],
          headerRow: false,
        },
      ],
    },
  ],
};

export default {
  id: 'table-header-row-materializes-header-cells',
  provenance: {source: 'excel-desktop-verification'},
  cluster: 'tables',
  description:
    'A table declaring a header row must write its column names into the worksheet cells of ' +
    'that row. Excel treats a declared-but-empty header row as corruption and repairs the file ' +
    'on open, discarding the column names.',

  behavior: [
    {
      name: 'a header-row table writes its column names into the header row cells',
      async expect(api: CorpusApi, assert: Assert) {
        const facts = await api.inspectPackage(HEADER_TABLE);
        const {cellText} = facts.sheets.S!;
        assert.strictEqual(cellText.A1, 'Alpha', `A1 should hold "Alpha", got ${cellText.A1}`);
        assert.strictEqual(cellText.B1, 'Beta', `B1 should hold "Beta", got ${cellText.B1}`);
      },
    },
    {
      name: 'every declared tableColumn name has a matching header cell in the grid',
      async expect(api: CorpusApi, assert: Assert) {
        const facts = await api.inspectPackage(HEADER_TABLE);
        const [table] = facts.tables;
        const {cellText} = facts.sheets.S!;
        assert.ok(table, 'no table part written');
        // The header row is the table ref's first row; the columns run left to right from its
        // top-left cell. Compare against the names the part itself declares, so the two can
        // never drift apart without this failing.
        table.columnNames.forEach((name: Untyped, index: number) => {
          const address = `${String.fromCharCode(65 + index)}1`;
          assert.strictEqual(
            cellText[address],
            name,
            `${address} should hold the declared column name "${name}", got ${cellText[address]}`,
          );
        });
      },
    },
    {
      name: 'a table anchored away from A1 writes its headers at its own anchor row',
      async expect(api: CorpusApi, assert: Assert) {
        const {cellText} = (await api.inspectPackage(OFFSET_TABLE)).sheets.S!;
        assert.strictEqual(cellText.C3, 'Gamma', `C3 should hold "Gamma", got ${cellText.C3}`);
        assert.strictEqual(cellText.D3, 'Delta', `D3 should hold "Delta", got ${cellText.D3}`);
      },
    },
    {
      name: 'a headerless table writes no header cells: its first row is data',
      async expect(api: CorpusApi, assert: Assert) {
        const {cellText} = (await api.inspectPackage(HEADERLESS_TABLE)).sheets.S!;
        // The first row is the data row ['x', 'y'], not the column names: a headerless table must
        // not materialize 'Alpha'/'Beta' into the grid the way a header-bearing one does.
        assert.strictEqual(
          cellText.A1,
          'x',
          `A1 should hold the first data value, got ${cellText.A1}`,
        );
        assert.strictEqual(
          cellText.B1,
          'y',
          `B1 should hold the first data value, got ${cellText.B1}`,
        );
      },
    },
  ],
} satisfies Case;
