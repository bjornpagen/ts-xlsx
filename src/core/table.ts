// A worksheet table (OOXML `<table>`): a named, structured range with typed columns,
// an optional header row, and an optional totals row.
//
// The model stores the anchor and the column/row counts, not a pre-baked range string:
// the occupied geometry is derived, so it stays correct as an empty table (header row
// only), a headerless table (data rows only), or a totals-bearing table. The writer is
// the OOXML gatekeeper for serialization; this model owns the invariants Excel enforces
// on the *shape* itself: a legal name, at least one column, and at least one row.

import {AuthoringError, quoted} from '../errors.ts';
import {tokenSet} from '../token-set.ts';
import {
  type CellPosition,
  decodeCellRef,
  encodeAddress,
  encodeRect,
  type GridRect,
  MAX_COLUMN,
  MAX_ROW,
  numberToColumn,
} from './address.ts';
import {type ClonePlan, cloneWith} from './clone.ts';
import {type AxisSplice, isDeletedSpan, shiftIndex} from './grid-shift.ts';
import type {AssertNever} from './internal.ts';
import {MAX_TABLE_NAME_LENGTH, TABLE_NAME_PATTERN} from './limits.ts';
import {structuredColumnReference} from './structured-reference.ts';
import type {CellStyle} from './style.ts';
import type {CellValue} from './value.ts';

/** A per-column cell format applied to a table's body cells: the facets Excel's table-column style
 * bakes into the cells rather than storing as table metadata. Every facet ({@link CellStyle}) is
 * optional; only the ones set are applied, leaving the rest of each cell's style untouched. */
export type TableColumnStyle = Readonly<CellStyle>;

/**
 * The channel a registered table holds into its owning worksheet's grid. A worksheet supplies it
 * when it registers the table; a table built standalone (a unit test, a bare model) has none, so it
 * can be inspected but cannot materialise or append cells, and appending throws rather than
 * silently dropping the values.
 *
 * All three coordinates are 1-based.
 */
export interface TableGrid {
  /**
   * Whether the cell at this position already holds a value. The materialiser's round-trip guard
   * asks this and nothing else: it must not create the cell, because asking whether a table's frame
   * is already filled would otherwise fill the grid with the empty cells it was asking about.
   */
  holdsValue(row: number, col: number): boolean;
  /** Write a value, applying the column's style (if any) to the cell. */
  writeCell(row: number, col: number, value: CellValue, style?: TableColumnStyle): void;
  /**
   * Insert one empty row at `row`, shifting that row and everything below it down by one: how a
   * table with a totals row opens a slot for an appended data row above the totals. Relocating the
   * totals row lives in the grid, which is why this is the grid's job and not the table's.
   */
  insertRow(row: number): void;
}

/**
 * A table's visual style (`<tableStyleInfo>`): the named style to apply plus the banding/highlight
 * toggles. Every field is a tri-state so a round-trip stays faithful: a value present in the source
 * re-emits, one the source omitted stays omitted rather than being defaulted to `"0"`. A workbook
 * whose part carries no `<tableStyleInfo>` at all leaves {@link TableOptions.style} undefined.
 */
export interface TableStyleInfo {
  /**
   * Named table style to apply: one of the built-in gallery (`"TableStyleMedium2"`, …) or a custom
   * one the workbook defines with {@link Workbook.addTableStyle}.
   *
   * **Not validated.** A name that matches nothing renders the table unstyled, silently. Even so,
   * this library must not be the thing that rejects it. A reader has to accept a name from a newer
   * Excel than the gallery list it was built with, and a writer that threw would make round-tripping
   * such a file impossible; there is also no diagnostics channel to warn through, so the only
   * options were "throw" and "accept". Accepting is the one that never makes a readable file
   * unreadable. If a warning channel is ever added, this is the first thing that should use it.
   */
  readonly name?: string;
  /** Emphasise the first column. */
  readonly showFirstColumn?: boolean;
  /** Emphasise the last column. */
  readonly showLastColumn?: boolean;
  /** Band the rows (alternating fill). */
  readonly showRowStripes?: boolean;
  /** Band the columns (alternating fill). */
  readonly showColumnStripes?: boolean;
}

/** Copy a style, keeping only its defined fields off the literal so `exactOptionalPropertyTypes`
 * never sees a fabricated `key: undefined`: an absent attribute must stay absent across a copy.
 *
 * The sentinel name `"None"` (Excel's table-style gallery entry for *no* style) is normalised to an
 * absent name: OOXML expresses "unstyled" as a `<tableStyleInfo>` with no `name` attribute, so a
 * literal `name="None"` would reference a style that does not exist and make the file suspect. The
 * banding flags set alongside it are untouched. */
const STYLE_INFO_CLONE: ClonePlan<TableStyleInfo> = {
  // The one field with a rule of its own, applied here rather than after the copy so the sentinel
  // never lands in the clone at all.
  name: (name) => (name === 'None' ? undefined : name) as string,
  showFirstColumn: 'value',
  showLastColumn: 'value',
  showRowStripes: 'value',
  showColumnStripes: 'value',
};

/** The proof that {@link STYLE_INFO_CLONE} names every field of the style. */
export type EveryTableStyleInfoFieldIsCloned = AssertNever<
  Exclude<keyof Required<TableStyleInfo>, keyof typeof STYLE_INFO_CLONE>
>;

/** The `<tableStyleInfo>` banding flags, in CT_TableStyleInfo attribute order. Each is an optional
 * OOXML boolean, so the list is the whole of what the reader and the writer need to know about them. */
export const TABLE_STYLE_FLAGS = [
  'showFirstColumn',
  'showLastColumn',
  'showRowStripes',
  'showColumnStripes',
] as const satisfies readonly (keyof TableStyleInfo)[];

/** Compile-time proof that `name` and {@link TABLE_STYLE_FLAGS} cover every {@link TableStyleInfo}
 * field, so a flag added to the type cannot be read or written by one direction only. */
export type EveryTableStyleFlagIsDeclared = AssertNever<
  Exclude<keyof TableStyleInfo, 'name' | (typeof TABLE_STYLE_FLAGS)[number]>
>;

function cloneStyleInfo(style: TableStyleInfo): TableStyleInfo {
  const clone = cloneWith(style, STYLE_INFO_CLONE);
  // `cloneWith` skips a field the source omits; the `None` rule turns a present one into an omission,
  // which it cannot express, so the key is removed here instead.
  if (clone.name === undefined) delete (clone as {name?: string}).name;
  return clone;
}

/**
 * OOXML's totals-row function names (`ST_TotalsRowFunction`) to the `SUBTOTAL` first-argument code
 * Excel writes into a materialised totals cell. The `10x` band ignores manually hidden rows, the
 * behaviour Excel's totals row uses. The one inversion trap: `count` is COUNTA (103, non-empty) while
 * `countNums` is COUNT (102, numbers only). `none` (no aggregate) has no built-in code, so a column
 * carrying it is absent here and its totals cell is left unmaterialised: Excel accepts the blank
 * cell. `custom` is likewise absent: its aggregate is not a `SUBTOTAL` but the arbitrary formula
 * stored in {@link TableColumn.totalsRowFormula}, which the reader/writer round-trip and the
 * materialiser writes into the cell verbatim.
 */
export const TOTALS_ROW_SUBTOTAL_CODE: Readonly<Partial<Record<TotalsRowFunction, number>>> = {
  average: 101,
  countNums: 102,
  count: 103,
  max: 104,
  min: 105,
  stdDev: 107,
  sum: 109,
  var: 110,
};

/**
 * The values `ST_TotalsRowFunction` (ECMA-376 §18.18.86) can take: a closed OOXML enumeration Excel
 * does not extend over time (unlike, say, a conditional-formatting rule type), so an author-side typo
 * such as `"avg"` is a compile error here rather than a silently no-op attribute at write time.
 */
export type TotalsRowFunction =
  | 'average'
  | 'countNums'
  | 'count'
  | 'max'
  | 'min'
  | 'stdDev'
  | 'sum'
  | 'var'
  | 'custom'
  | 'none';

/** Narrow a raw `totalsRowFunction` attribute to a known {@link TotalsRowFunction}. */
export const isTotalsRowFunction = tokenSet<TotalsRowFunction>({
  average: true,
  countNums: true,
  count: true,
  max: true,
  min: true,
  stdDev: true,
  sum: true,
  var: true,
  custom: true,
  none: true,
});

/** One column of a table: a header name and its optional totals-row behaviour. */
export interface TableColumn {
  /** The column's header/display name. Must be unique within the table (case-insensitively):
   * Excel writes a table with colliding column names as corrupt. A collision supplied at construction
   * is disambiguated deterministically (the first keeps its name, later clashes gain a numeric
   * suffix), the same repair the reader applies to a loaded file, rather than being rejected. */
  readonly name: string;
  /** Literal label shown in the totals row (e.g. `"Total"`), mutually exclusive with a function. */
  readonly totalsRowLabel?: string;
  /** Built-in totals-row aggregate (`"sum"`, `"average"`, `"count"`, …), or `"custom"` when the
   * column's total is the arbitrary formula in {@link totalsRowFormula} rather than a `SUBTOTAL`. */
  readonly totalsRowFunction?: TotalsRowFunction;
  /** The formula (no leading `=`) backing a `totalsRowFunction: "custom"` column. This is OOXML's
   * `<totalsRowFormula>` child. Round-tripped verbatim and written into the totals cell as the
   * cell's formula. Meaningful only alongside `totalsRowFunction: "custom"`; ignored otherwise. */
  readonly totalsRowFormula?: string;
  /** A format applied to this column's body cells as they are written (see {@link TableColumnStyle}).
   * Excel bakes a table-column style into the cells rather than storing it as table metadata, so this
   * is an authoring convenience: it round-trips as the affected cells' own styles, not as the table. */
  readonly style?: TableColumnStyle;
}

export interface TableOptions {
  /** Table name: a valid Excel identifier, unique across the workbook. This is the name used in
   * structured formula references (`Table1[Column]`). */
  name: string;
  /** Human-facing display name shown in the UI. A free-form label (spaces allowed) that need not
   * be a valid identifier. Defaults to {@link name} when omitted. */
  displayName?: string;
  /** A1 reference of the table's top-left cell (an anchor, e.g. `"A1"`, not the full range). */
  ref: string;
  /** The table's columns, left to right. At least one is required. */
  columns: readonly TableColumn[];
  /** Number of data rows (excludes the header and totals rows). May be zero. */
  rowCount: number;
  /** Whether the table has a header row. Defaults to `true`. */
  headerRow?: boolean;
  /** Whether the table has a totals row. Defaults to `false`. */
  totalsRow?: boolean;
  /** The `totalsRowShown` flag on a table *without* a totals row: Excel's record of whether a
   * totals row has ever been toggled on. Tri-state so a round-trip is faithful: `false` re-emits
   * `totalsRowShown="0"`, `true` re-emits `totalsRowShown="1"`, and `undefined` (the authoring
   * default) emits nothing: a file read without the attribute must not have one injected. Ignored
   * when {@link totalsRow} is set, since a present totals row already implies it is shown. */
  totalsRowShown?: boolean;
  /** Whether the header row carries an autoFilter. Defaults to {@link headerRow}: a header table
   * gains an autoFilter, a headerless one never can. Set `false` to keep a header table's rows
   * unfiltered: a file read without an autoFilter must round-trip without one being injected. */
  autoFilter?: boolean;
  /** The table's visual style. Preserved verbatim across a round-trip; when omitted, a freshly
   * authored table is written with Excel's default (`TableStyleMedium2`, banded rows). A part read
   * with no `<tableStyleInfo>` sets this to `undefined`. See {@link TableStyleInfo}. */
  style?: TableStyleInfo;
}

/**
 * Return copies of `columns` with every name made unique (case-insensitively): the first occurrence
 * keeps its name; a later clash gains the smallest numeric suffix that resolves it (`foo`, `foo2`,
 * `foo3`, …). OOXML requires unique column names within a table, and Excel treats a collision as
 * corruption, so this is applied both when a table is authored and when one is read from a file,
 * keeping the two paths identical rather than rejecting a name list the reader would accept.
 */
function disambiguateColumnNames(columns: readonly TableColumn[]): TableColumn[] {
  const seen = new Set<string>();
  return columns.map((column) => {
    let candidate = column.name;
    for (let n = 2; seen.has(candidate.toLowerCase()); n++) candidate = `${column.name}${n}`;
    seen.add(candidate.toLowerCase());
    return candidate === column.name ? {...column} : {...column, name: candidate};
  });
}

// Both throws are native, which is the call this one needed: a table name is a single scalar, and
// "does not parse as an identifier" is the same kind of failure as a comment id that is not a GUID,
// which `comment-thread.ts` already raises as a `SyntaxError`. The composite claim about a table --
// that its columns span its range, that it does not name a column twice -- is elsewhere, and that
// one stays an `AuthoringError`.
function validateTableName(name: string): void {
  if (name.length === 0 || name.length > MAX_TABLE_NAME_LENGTH) {
    throw new RangeError(
      `table name ${quoted(name)} must be between 1 and ${MAX_TABLE_NAME_LENGTH} characters`,
    );
  }
  if (!TABLE_NAME_PATTERN.test(name)) {
    throw new SyntaxError(
      `table name ${quoted(name)} is not a valid Excel identifier: it must start with a letter, ` +
        'underscore, or backslash and contain only letters, digits, periods, and underscores',
    );
  }
}

/** The rectangle a table occupies, as the {@link GridRect} every range-shaped thing in the library is. */
export type TableRegion = GridRect;

export class Table {
  readonly name: string;
  readonly displayName: string;
  readonly columns: readonly TableColumn[];
  readonly headerRow: boolean;
  readonly totalsRow: boolean;
  readonly totalsRowShown: boolean | undefined;
  readonly autoFilter: boolean;
  readonly style: TableStyleInfo | undefined;

  // The anchor and data-row count move when a row/column splice shifts or resizes the table, so
  // they are mutable behind the class's controlled `shiftRows`/`shiftColumns` methods.
  #anchorCol: number;
  #anchorRow: number;
  #dataRowCount: number;

  readonly #grid: TableGrid | undefined;

  constructor(options: TableOptions, grid?: TableGrid) {
    validateTableName(options.name);
    if (options.columns.length === 0) {
      throw new AuthoringError(`table ${quoted(options.name)} must declare at least one column`);
    }
    if (!Number.isInteger(options.rowCount) || options.rowCount < 0) {
      throw new RangeError(
        `table ${quoted(options.name)} has an invalid data-row count (${options.rowCount})`,
      );
    }
    let anchor: CellPosition;
    try {
      anchor = decodeCellRef(options.ref);
    } catch (cause) {
      // A position off the grid is already the precise answer, in the words every other door gives.
      if (cause instanceof RangeError) throw cause;
      // The generic "not a single-cell reference" says less than naming the table, so it becomes
      // the cause of a message that does.
      throw new SyntaxError(
        `table ref ${quoted(options.ref)} must anchor at a single cell (e.g. "A1")`,
        {
          cause,
        },
      );
    }
    const {col, row} = anchor;

    this.name = options.name;
    this.displayName = options.displayName ?? options.name;
    this.columns = disambiguateColumnNames(options.columns);
    this.headerRow = options.headerRow ?? true;
    this.totalsRow = options.totalsRow ?? false;
    this.totalsRowShown = options.totalsRowShown;
    this.style = options.style === undefined ? undefined : cloneStyleInfo(options.style);
    // A header table gains an autoFilter by default (Excel's behaviour when a table is inserted);
    // a headerless table can never carry one: an autoFilter has no header row to anchor to.
    this.autoFilter = this.headerRow && (options.autoFilter ?? true);
    this.#anchorCol = col;
    this.#anchorRow = row;
    this.#dataRowCount = options.rowCount;
    this.#grid = grid;

    if (this.#rowSpan < 1) {
      throw new AuthoringError(
        `table ${quoted(this.name)} has no rows: it needs a header row or at least one data row`,
      );
    }
    // The anchor is inside the grid by construction; the far corner is derived from it and was never
    // checked. A table anchored at `XFC1` with three columns reached column XFE, and one with five
    // million data rows produced `A1:A5000001`, a `<table ref>` naming rows that cannot exist that
    // the writer then emitted. Neither is repairable after the fact: `range`, `autoFilterRef` and
    // `region` all throw on read once the corner is out of bounds, so the shape is refused here.
    if (this.#right > MAX_COLUMN) {
      throw new RangeError(
        `table ${quoted(this.name)} spans ${this.columns.length} columns from ` +
          `${numberToColumn(this.#anchorCol)}, past the last column (${numberToColumn(MAX_COLUMN)})`,
      );
    }
    if (this.#bottom > MAX_ROW) {
      throw new RangeError(
        `table ${quoted(this.name)} spans ${this.#rowSpan} rows from ${this.#anchorRow}, ` +
          `past the last row (${MAX_ROW})`,
      );
    }

    if (grid !== undefined) this.#materializeFrame(grid);
  }

  get columnCount(): number {
    return this.columns.length;
  }

  /** The number of data rows (excludes the header and totals rows). Always defined: a table loaded
   * from a file derives it from the stored range, so reading the height never throws. */
  get rowCount(): number {
    return this.#dataRowCount;
  }

  /**
   * Append a data row to the bottom of the table, growing its range by one row and writing `values`
   * left-to-right across its columns. A loaded table exposes its rows the same as a freshly-authored
   * one, so this works identically whether the table was built in memory or read from a file.
   *
   * A table carrying a totals row appends above it: the new data row lands where the totals row sat,
   * and the totals row (with any sheet content below) shifts down by one, exactly what inserting a
   * worksheet row does. That relocation lives in the grid, so a totals-row table not attached to a
   * worksheet throws, as does passing `values` on any detached table: there is nowhere to put them.
   */
  addRow(values: readonly CellValue[] = []): void {
    if (values.length > this.columnCount) {
      throw new RangeError(
        `row has ${values.length} values but table ${quoted(this.name)} has ${this.columnCount} columns`,
      );
    }

    // The append point is the row directly below the last data row: the totals row when one exists,
    // otherwise the first free row under the table.
    const target = this.#anchorRow + (this.headerRow ? 1 : 0) + this.#dataRowCount;

    const grid = this.#grid;
    if (this.totalsRow) {
      if (grid === undefined) {
        throw new AuthoringError(
          `table ${quoted(this.name)} is not attached to a worksheet: cannot relocate its totals row to append a data row`,
        );
      }
      // Opening a grid slot at the totals row shifts the totals down and grows this table by one
      // through the sheet's own table re-pinning, so #dataRowCount is not bumped again here.
      grid.insertRow(target);
      this.#writeRow(grid, target, values);
      return;
    }

    if (values.length > 0) {
      if (grid === undefined) {
        throw new AuthoringError(
          `table ${quoted(this.name)} is not attached to a worksheet: cannot write appended row values`,
        );
      }
      this.#writeRow(grid, target, values);
    }
    this.#dataRowCount += 1;
  }

  #writeRow(grid: TableGrid, row: number, values: readonly CellValue[]): void {
    values.forEach((value, index) => {
      grid.writeCell(row, this.#anchorCol + index, value, this.columns[index]?.style);
    });
  }

  // Fill the header and totals cells this table's own declaration implies, without clobbering
  // anything already there. It runs once, at the end of construction, because a table's frame is
  // part of what declaring the table means: a worksheet that registers one gets the cells with it,
  // and a table built standalone has no grid and so materialises nothing.
  //
  // Two independent findings below, kept apart because the reasons differ. The header fill is a
  // validity fix; the totals fill is a rendering-parity nicety. Both share one round-trip guard:
  // only an *empty* cell is filled.
  #materializeFrame(grid: TableGrid): void {
    // A table's declared range includes its header row, and Excel treats the column metadata and
    // the cells under it as one fact: a header row that is empty in the grid is corruption, and
    // Excel repairs the file on open, discarding the column names entirely. The caller already
    // named the columns once in the table definition, so materialising them here is what makes the
    // obvious API call produce a file that opens.
    //
    // Only *empty* header cells are filled. Reading a workbook re-registers each table after the
    // sheet's cells are loaded, and those cells are authoritative: they may carry rich text, a
    // style, or text that drifted from the column name, none of which a re-declaration may clobber.
    // An empty cell has no such content to lose.
    if (this.headerRow) {
      const {top, left} = this.region;
      this.columns.forEach((column, index) => {
        const col = left + index;
        if (grid.holdsValue(top, col)) return;
        grid.writeCell(top, col, column.name);
      });
    }

    // Materialize the totals row Excel renders on open, so our files show it immediately rather than a
    // blank strip until the user interacts. A labelled column writes its label string; an aggregate
    // column writes the `SUBTOTAL(code, Table[Column])` formula Excel would compute. Unlike the header
    // row, this is a UX-parity nicety, not a validity fix. Excel opens a declared-but-empty totals row
    // without repair; matching its on-open rendering is still the point.
    //
    // Same round-trip guard as the header row: only *empty* cells are filled. Reading a file
    // re-registers the table after its cells are loaded, so a materialized totals cell (ours, Excel's,
    // or a hand-set override) is authoritative and must survive untouched, keeping the round-trip
    // idempotent. The formula carries no cached result; Excel computes an uncached formula cell on open,
    // so the row shows real values without the library pretending to be a calc engine. A `custom` column
    // writes its stored `totalsRowFormula` verbatim; a `none` column (or a `custom` with no stored
    // formula) has nothing to write (see {@link TOTALS_ROW_SUBTOTAL_CODE}) and stays blank.
    if (this.totalsRow) {
      const {left, bottom} = this.region;
      this.columns.forEach((column, index) => {
        const col = left + index;
        if (grid.holdsValue(bottom, col)) return;
        if (column.totalsRowLabel !== undefined) {
          grid.writeCell(bottom, col, column.totalsRowLabel);
          return;
        }
        if (column.totalsRowFunction === undefined) return;
        const code = TOTALS_ROW_SUBTOTAL_CODE[column.totalsRowFunction];
        if (code !== undefined) {
          grid.writeCell(bottom, col, {
            formula: `SUBTOTAL(${code},${structuredColumnReference(this.name, column.name)})`,
          });
        } else if (column.totalsRowFunction === 'custom' && column.totalsRowFormula !== undefined) {
          // A `custom` total is the column's own stored formula, not a SUBTOTAL. Excel stores it
          // without a leading `=`, which is the formula string a cell value expects.
          grid.writeCell(bottom, col, {formula: column.totalsRowFormula});
        }
      });
    }
  }

  /**
   * Re-pin the table through a row splice. A splice entirely above the table moves its whole range
   * by the splice's `delta`; one landing inside grows or shrinks the data rows to absorb the change;
   * one that deletes the table's every row removes it. Returns `false` when the table no longer has
   * a row to occupy (the caller drops it), `true` when it survives.
   */
  shiftRows(splice: AxisSplice): boolean {
    // A table whose every row lies within the deleted span has nothing left to occupy.
    if (isDeletedSpan(this.#anchorRow, this.#bottom, splice)) return false;
    const top = shiftIndex(this.#anchorRow, splice);
    const bottom = shiftIndex(this.#bottom, splice);
    const span = bottom - top + 1;
    const fixedRows = (this.headerRow ? 1 : 0) + (this.totalsRow ? 1 : 0);
    const dataRows = span - fixedRows;
    if (span < 1 || dataRows < 0) return false;
    this.#anchorRow = top;
    this.#dataRowCount = dataRows;
    return true;
  }

  /**
   * Re-pin the table through a column splice, the mirror of {@link shiftRows}. A splice entirely to
   * the table's left moves its anchor by the splice's `delta`; one to its right leaves it untouched;
   * one that deletes the table's every column removes it. Returns `false` when the table no longer
   * has a column to occupy (the caller drops it), `true` when it survives.
   *
   * A splice landing *inside* the table's columns is structural surgery on named columns with no
   * unambiguous answer, so those columns are left as-is rather than fabricated or dropped. Whole-table
   * deletion is not that case: a table left declared over whatever slid into its place, carrying the
   * names of columns that no longer exist, is content the writer then emits.
   */
  shiftColumns(splice: AxisSplice): boolean {
    // A table whose every column lies within the deleted span has nothing left to occupy.
    if (isDeletedSpan(this.#anchorCol, this.#right, splice)) return false;
    // Clamped, like every other coordinate a splice moves: an unbounded increment could put the anchor
    // past the last column, where `range`, `autoFilterRef` and `region` all throw on read.
    const anchor = shiftIndex(this.#anchorCol, splice);
    // Clamping the anchor is not the same as bounding the table: `#right` is derived from the anchor
    // and the column count, so an anchor clamped onto XFD still puts a two-column table's right edge
    // at XFE, which is the unreadable table the clamp above was meant to prevent. A table with no
    // room left for its columns is dropped, the answer `shiftRows` gives one with no room for its rows.
    if (anchor + this.columns.length - 1 > MAX_COLUMN) return false;
    this.#anchorCol = anchor;
    return true;
  }

  /**
   * The options that reconstruct this table: the anchor as a single-cell ref (not the derived
   * full range), the columns, and the data-row count with the header/totals flags. Feeding this
   * back to the constructor yields an equivalent table, so a worksheet model can carry a table
   * losslessly across an export/import round-trip.
   */
  get options(): TableOptions {
    const options: TableOptions = {
      name: this.name,
      displayName: this.displayName,
      ref: encodeAddress(this.#anchorCol, this.#anchorRow),
      columns: this.columns.map((column) => ({...column})),
      rowCount: this.#dataRowCount,
      headerRow: this.headerRow,
      totalsRow: this.totalsRow,
      autoFilter: this.autoFilter,
    };
    // Kept off the literal so `undefined` (attribute absent) stays absent, not an explicit
    // `totalsRowShown: undefined`: the round-trip must not fabricate the flag.
    if (this.totalsRowShown !== undefined) options.totalsRowShown = this.totalsRowShown;
    if (this.style !== undefined) options.style = cloneStyleInfo(this.style);
    return options;
  }

  /** The full A1 range the table occupies: header (if any) + data rows + totals (if any). Distinct
   * from {@link TableOptions.ref} (and {@link options}'s own `ref`), which is only the single-cell
   * anchor a table is constructed from; this is the anchor plus the columns/rows it has grown to
   * cover. */
  get range(): string {
    return encodeRect({
      top: this.#anchorRow,
      left: this.#anchorCol,
      bottom: this.#bottom,
      right: this.#right,
    });
  }

  /**
   * The autoFilter range (the header row plus the data rows, never the totals row), or
   * `undefined` when the table has no autoFilter: either it is headerless (an autoFilter has
   * nothing to anchor to and Excel treats its presence as corruption) or its {@link autoFilter}
   * flag is off (a table read without one must not gain one on round-trip).
   */
  get autoFilterRef(): string | undefined {
    if (!this.autoFilter) return undefined;
    const bottom = this.#anchorRow + this.#dataRowCount;
    return encodeRect({top: this.#anchorRow, left: this.#anchorCol, bottom, right: this.#right});
  }

  /** The occupied rectangle, for conflict checks such as overlapping merges. */
  get region(): TableRegion {
    return {top: this.#anchorRow, left: this.#anchorCol, bottom: this.#bottom, right: this.#right};
  }

  get #right(): number {
    return this.#anchorCol + this.columns.length - 1;
  }

  get #rowSpan(): number {
    return (this.headerRow ? 1 : 0) + this.#dataRowCount + (this.totalsRow ? 1 : 0);
  }

  get #bottom(): number {
    return this.#anchorRow + this.#rowSpan - 1;
  }
}
