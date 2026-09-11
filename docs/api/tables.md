# Tables

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `isTotalsRowFunction`

<sub>const</sub>

Narrow a raw `totalsRowFunction` attribute to a known [`TotalsRowFunction`](./tables.md#totalsrowfunction).

```ts
const isTotalsRowFunction: (value: string) => value is TotalsRowFunction
```

---

### `Table`

<sub>class</sub>

```ts
class Table {
  readonly name: string;
  readonly columns: readonly TableColumn[];
  readonly headerRow: boolean;
  readonly totalsRow: boolean;
  readonly totalsRowShown: boolean | undefined;
  readonly autoFilter: boolean;
  readonly style: TableStyleInfo | undefined;
  get columnCount(): number;
  get rowCount(): number;
  addRow(values: readonly CellValue[] = []): void;
  shiftRows(splice: AxisSplice): boolean;
  shiftColumns(splice: AxisSplice): boolean;
  get options(): TableOptions;
  get range(): string;
  get autoFilterRef(): string | undefined;
  get region(): TableRegion;
}
```

**Members**

#### `Table.rowCount`

```ts
get rowCount(): number;
```

The number of data rows (excludes the header and totals rows). Always defined: a table loaded
from a file derives it from the stored range, so reading the height never throws.

#### `Table.addRow`

```ts
addRow(values: readonly CellValue[] = []): void;
```

Append a data row to the bottom of the table, growing its range by one row and writing `values`
left-to-right across its columns. A loaded table exposes its rows the same as a freshly-authored
one, so this works identically whether the table was built in memory or read from a file.

A table carrying a totals row appends above it: the new data row lands where the totals row sat,
and the totals row (with any sheet content below) shifts down by one, exactly what inserting a
worksheet row does. That relocation lives in the grid, so a totals-row table not attached to a
worksheet throws, as does passing `values` on any detached table: there is nowhere to put them.

#### `Table.shiftRows`

```ts
shiftRows(splice: AxisSplice): boolean;
```

Re-pin the table through a row splice. A splice entirely above the table moves its whole range
by the splice's `delta`; one landing inside grows or shrinks the data rows to absorb the change;
one that deletes the table's every row removes it. Returns `false` when the table no longer has
a row to occupy (the caller drops it), `true` when it survives.

#### `Table.shiftColumns`

```ts
shiftColumns(splice: AxisSplice): boolean;
```

Re-pin the table through a column splice, the mirror of [`shiftRows`](./tables.md#tableshiftrows). A splice entirely to
the table's left moves its anchor by the splice's `delta`; one to its right leaves it untouched;
one that deletes the table's every column removes it. Returns `false` when the table no longer
has a column to occupy (the caller drops it), `true` when it survives.

A splice landing *inside* the table's columns is structural surgery on named columns with no
unambiguous answer, so those columns are left as-is rather than fabricated or dropped. Whole-table
deletion is not that case: a table left declared over whatever slid into its place, carrying the
names of columns that no longer exist, is content the writer then emits.

#### `Table.options`

```ts
get options(): TableOptions;
```

The options that reconstruct this table: the anchor as a single-cell ref (not the derived
full range), the columns, and the data-row count with the header/totals flags. Feeding this
back to the constructor yields an equivalent table, so a worksheet model can carry a table
losslessly across an export/import round-trip.

#### `Table.range`

```ts
get range(): string;
```

The full A1 range the table occupies: header (if any) + data rows + totals (if any). Distinct
from [`TableOptions.ref`](./tables.md#tableoptions) (and [`options`](./tables.md#tableoptions)'s own `ref`), which is only the single-cell
anchor a table is constructed from; this is the anchor plus the columns/rows it has grown to
cover.

#### `Table.autoFilterRef`

```ts
get autoFilterRef(): string | undefined;
```

The autoFilter range (the header row plus the data rows, never the totals row), or
`undefined` when the table has no autoFilter: either it is headerless (an autoFilter has
nothing to anchor to and Excel treats its presence as corruption) or its `autoFilter`
flag is off (a table read without one must not gain one on round-trip).

#### `Table.region`

```ts
get region(): TableRegion;
```

The occupied rectangle, for conflict checks such as overlapping merges.

---

### `TableColumn`

<sub>interface</sub>

One column of a table: a header name and its optional totals-row behaviour.

```ts
interface TableColumn {
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
```

---

### `TableColumnStyle`

<sub>type</sub>

A per-column cell format applied to a table's body cells: the facets Excel's table-column style
bakes into the cells rather than storing as table metadata. Every facet ([`CellStyle`](./styles.md#cellstyle)) is
optional; only the ones set are applied, leaving the rest of each cell's style untouched.

```ts
type TableColumnStyle = Readonly<CellStyle>;
```

---

### `TableGrid`

<sub>interface</sub>

The channel a registered table holds into its owning worksheet's grid. A worksheet supplies it
when it registers the table; a table built standalone (a unit test, a bare model) has none, so it
can be inspected but cannot materialise or append cells, and appending throws rather than
silently dropping the values.

All three coordinates are 1-based.

```ts
interface TableGrid {
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
```

---

### `TableOptions`

<sub>interface</sub>

```ts
interface TableOptions {
  /**
   * Table name: a valid Excel identifier, unique across the workbook. It is the one name a table has:
   * the name structured references use (`Table1[Column]`) and the one Excel shows, written as both the
   * table part's `name` and its `displayName`.
   *
   * There is no separate display label. OOXML's `displayName` is the identifier Excel resolves a
   * structured reference against and shows in the UI, and a package whose `displayName` carries a space
   * opens with Excel's repair prompt.
   */
  name: string;
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
```

---

### `TableRegion`

<sub>type</sub>

The rectangle a table occupies, as the [`GridRect`](./addresses-ranges.md#gridrect) every range-shaped thing in the library is.

```ts
type TableRegion = GridRect;
```

---

### `TableStyleInfo`

<sub>interface</sub>

A table's visual style (`<tableStyleInfo>`): the named style to apply plus the banding/highlight
toggles. Every field is a tri-state so a round-trip stays faithful: a value present in the source
re-emits, one the source omitted stays omitted rather than being defaulted to `"0"`. A workbook
whose part carries no `<tableStyleInfo>` at all leaves [`TableOptions.style`](./tables.md#tableoptions) undefined.

```ts
interface TableStyleInfo {
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
```

---

### `TotalsRowFunction`

<sub>type</sub>

The values `ST_TotalsRowFunction` (ECMA-376 §18.18.86) can take: a closed OOXML enumeration Excel
does not extend over time (unlike, say, a conditional-formatting rule type), so an author-side typo
such as `"avg"` is a compile error here rather than a silently no-op attribute at write time.

```ts
type TotalsRowFunction =
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
```
