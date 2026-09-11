// Cluster: tables
//
// Real-world scenario: a table declares a totals row (totalsRowCount="1") with per-column directives:
// a literal label on one column ("Total") and a built-in aggregate on another (sum, count, …). Excel
// renders that row on open: the labelled column shows its text and the aggregate column shows a
// SUBTOTAL formula over the table's column. A library that declares the totals row's geometry but
// writes nothing into those grid cells produces a file whose totals row is blank until the user
// interacts with the table, a UX-parity gap against Excel's on-open rendering.
//
// So the library materializes the row it declares: the label cell holds the label string, and the
// aggregate cell holds `SUBTOTAL(code, Table[Column])`: the same formula Excel would compute, with the
// correct SUBTOTAL function code per aggregate (the count/countNums → COUNTA/COUNT inversion is the
// trap). The formula carries no cached result: the library is not a calc engine, and Excel computes an
// uncached formula cell on open. A column with no built-in aggregate stays blank, exactly as the whole
// row did before, so nothing regresses for it: a `custom` column's stored formula is materialized
// separately (see the table-column-custom-totals-formula-roundtrip case).
//
// The materialization mirrors the header row's "fill only empty cells" guard, which keeps a round-trip
// idempotent: reloading a file re-registers the table after its cells are loaded, so the already-present
// totals cells are authoritative and must survive a write→read→write cycle untouched.
//
// This was confirmed against Excel Desktop through the Excel-oracle harness (ADR 0013) on Excel 16.0
// build 20131: sidecar `test/corpus/fixtures/excel-oracle/table-totals-row-materialized.json`. Excel
// opened the materialized totals row without repair and computed both uncached SUBTOTAL cells on open
// (COUNTA→2, SUM→30), so the no-cached-`<v>` emission is safe and `fullCalcOnLoad` is not required. It
// accepts our fully-qualified `T[Column]` structured reference and renders it in its own canonical
// intra-table form `[Column]` (the table name is elided for a reference inside its own table); both
// spellings compute identically, so the qualified spelling is a display-only difference. The behaviors
// below are the Tier-2 seam facts that LOCK that Tier-3 finding in CI (ADR 0012 seed+lock split): they
// assert our emitted label and formula text structurally, without re-opening Excel.

import type {Assert, Case, CorpusApi} from '../case.ts';

// Header row 1, data rows 2–3, totals row 4. Anchored at A1 so the totals row is row 4, columns A/B.
const LABEL_AND_SUM = {
  sheets: [
    {
      name: 'S',
      tables: [
        {
          name: 'T',
          ref: 'A1',
          totalsRow: true,
          columnDefs: [
            {name: 'Item', totalsRowLabel: 'Total'},
            {name: 'Amount', totalsRowFunction: 'sum'},
          ],
          rows: [
            ['a', 1],
            ['b', 2],
          ],
        },
      ],
    },
  ],
};

// Anchored away from A1 so a fix hard-coding the totals row (or column A) instead of using the table's
// own region still fails this. Totals row is row 5 (header C3, data C4, totals C5), columns C/D.
const OFFSET_TABLE = {
  sheets: [
    {
      name: 'S',
      tables: [
        {
          name: 'Anchored',
          ref: 'C3',
          totalsRow: true,
          columnDefs: [
            {name: 'Item', totalsRowLabel: 'Sum'},
            {name: 'Amount', totalsRowFunction: 'sum'},
          ],
          rows: [['a', 1]],
        },
      ],
    },
  ],
};

// count vs countNums map to different SUBTOTAL codes, COUNTA (103) and COUNT (102), the one easy
// inversion to get wrong. Totals row is row 3 (header 1, data 2, totals 3), columns A/B.
const COUNT_VARIANTS = {
  sheets: [
    {
      name: 'S',
      tables: [
        {
          name: 'T',
          ref: 'A1',
          totalsRow: true,
          columnDefs: [
            {name: 'Names', totalsRowFunction: 'count'},
            {name: 'Nums', totalsRowFunction: 'countNums'},
          ],
          rows: [['a', 1]],
        },
      ],
    },
  ],
};

// A column with neither directive (and one Excel cannot express as a built-in) has no aggregate to
// materialize; its totals cell stays blank. Totals row is row 3, columns A/B.
const NO_AGGREGATE = {
  sheets: [
    {
      name: 'S',
      tables: [
        {
          name: 'T',
          ref: 'A1',
          totalsRow: true,
          columnDefs: [{name: 'Item', totalsRowLabel: 'Total'}, {name: 'Note'}],
          rows: [['a', 'x']],
        },
      ],
    },
  ],
};

// A table with no totals row must not gain totals cells: nothing below its single data row.
const NO_TOTALS_ROW = {
  sheets: [
    {
      name: 'S',
      tables: [
        {
          name: 'T',
          ref: 'A1',
          columnDefs: [{name: 'Item'}, {name: 'Amount'}],
          rows: [['a', 1]],
        },
      ],
    },
  ],
};

export default {
  id: 'table-totals-row-materializes-label-and-subtotal',
  provenance: {
    source: 'excel-desktop-verification',
    ref: 'test/corpus/fixtures/excel-oracle/table-totals-row-materialized.json',
  },
  cluster: 'tables',
  description:
    'A table declaring a totals row writes its per-column directives into the totals-row grid cells ' +
    "to match Excel's on-open rendering: a label column holds its label string, an aggregate column " +
    'holds a SUBTOTAL(code, Table[Column]) formula with the correct function code, and a column with ' +
    'no built-in aggregate stays blank. The materialized row survives a round-trip unchanged.',

  behavior: [
    {
      name: 'a labelled totals column writes its label into the totals-row cell',
      async expect(api: CorpusApi, assert: Assert) {
        const {cellText} = (await api.inspectPackage(LABEL_AND_SUM)).sheets.S!;
        assert.strictEqual(cellText.A4, 'Total', `A4 should hold "Total", got ${cellText.A4}`);
      },
    },
    {
      name: 'a sum totals column writes SUBTOTAL(109, Table[Column]) into the totals-row cell',
      async expect(api: CorpusApi, assert: Assert) {
        const {formulas} = (await api.inspectPackage(LABEL_AND_SUM)).sheets.S!;
        assert.strictEqual(
          formulas.B4,
          'SUBTOTAL(109,T[Amount])',
          `B4 should hold the sum SUBTOTAL formula, got ${formulas.B4}`,
        );
      },
    },
    {
      name: 'the materialized totals formula caches no result (Excel recomputes on open)',
      async expect(api: CorpusApi, assert: Assert) {
        // A formula cell with a cached value would appear in cellText; the totals formula must not, so
        // Excel treats it as needing computation rather than trusting a value the library cannot supply.
        const {cellText} = (await api.inspectPackage(LABEL_AND_SUM)).sheets.S!;
        assert.strictEqual(
          cellText.B4,
          undefined,
          `B4 should carry a formula with no cached <v>, got cached ${cellText.B4}`,
        );
      },
    },
    {
      name: 'a totals row is materialized at the table anchor, not a hard-coded row or column',
      async expect(api: CorpusApi, assert: Assert) {
        const facts = (await api.inspectPackage(OFFSET_TABLE)).sheets.S;
        assert.strictEqual(
          facts!.cellText.C5,
          'Sum',
          `C5 should hold "Sum", got ${facts!.cellText.C5}`,
        );
        assert.strictEqual(
          facts!.formulas.D5,
          'SUBTOTAL(109,Anchored[Amount])',
          `D5 should hold the sum SUBTOTAL formula, got ${facts!.formulas.D5}`,
        );
      },
    },
    {
      name: 'count maps to SUBTOTAL 103 (COUNTA) and countNums to 102 (COUNT)',
      async expect(api: CorpusApi, assert: Assert) {
        const {formulas} = (await api.inspectPackage(COUNT_VARIANTS)).sheets.S!;
        assert.strictEqual(
          formulas.A3,
          'SUBTOTAL(103,T[Names])',
          `count should be COUNTA=103, got ${formulas.A3}`,
        );
        assert.strictEqual(
          formulas.B3,
          'SUBTOTAL(102,T[Nums])',
          `countNums should be COUNT=102, got ${formulas.B3}`,
        );
      },
    },
    {
      name: 'a totals column with no built-in aggregate leaves its totals cell blank',
      async expect(api: CorpusApi, assert: Assert) {
        const facts = (await api.inspectPackage(NO_AGGREGATE)).sheets.S;
        assert.strictEqual(
          facts!.cellText.A3,
          'Total',
          `A3 should hold "Total", got ${facts!.cellText.A3}`,
        );
        assert.strictEqual(
          facts!.cellText.B3,
          undefined,
          `B3 (no aggregate) should be blank, got ${facts!.cellText.B3}`,
        );
        assert.strictEqual(
          facts!.formulas.B3,
          undefined,
          `B3 (no aggregate) should carry no formula, got ${facts!.formulas.B3}`,
        );
      },
    },
    {
      name: 'a table without a totals row gains no totals cells',
      async expect(api: CorpusApi, assert: Assert) {
        const facts = (await api.inspectPackage(NO_TOTALS_ROW)).sheets.S;
        assert.strictEqual(
          facts!.cellText.A3,
          undefined,
          `A3 should be blank, got ${facts!.cellText.A3}`,
        );
        assert.strictEqual(
          facts!.formulas.B3,
          undefined,
          `B3 should carry no formula, got ${facts!.formulas.B3}`,
        );
      },
    },
    {
      name: 'the materialized totals row survives a write→read→write round-trip unchanged',
      async expect(api: CorpusApi, assert: Assert) {
        const {cellText, formulas} = (await api.roundtripInspectPackage(LABEL_AND_SUM)).sheets.S!;
        assert.strictEqual(
          cellText.A4,
          'Total',
          `A4 label should survive the round-trip, got ${cellText.A4}`,
        );
        assert.strictEqual(
          formulas.B4,
          'SUBTOTAL(109,T[Amount])',
          `the sum formula should survive the round-trip, got ${formulas.B4}`,
        );
        assert.strictEqual(
          cellText.B4,
          undefined,
          `the round-tripped formula must not gain a cached value, got ${cellText.B4}`,
        );
      },
    },
  ],
} satisfies Case;
