# Worksheet

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `CellModel`

<sub>interface</sub>

One materialised cell in a [`WorksheetModel`](./worksheet.md#worksheetmodel): its position, value, note, and every facet of its
formatting. Extends [`CellContent`](./styles.md#cellcontent) rather than [`CellStyle`](./styles.md#cellstyle) so the quote-prefix flag and
the named-style link travel with a model round-trip: they are written and read back like any other
facet, and leaving them off the tuple is what made a `dst.model = src.model` drop them.

```ts
interface CellModel extends CellContent {
  readonly row: number;
  readonly col: number;
  value: CellValue;
  note?: string | undefined;
}
```

---

### `ColumnProperties`

<sub>interface</sub>

Per-column formatting. A column may exist purely to carry these, with no cells. The style
facets are *defaults* for the column's cells: a cell that sets a facet of its own wins, but
one that leaves a facet unset inherits the column's, the same precedence Excel applies, and
symmetric with how a [`RowProperties`](./worksheet.md#rowproperties) fill defaults a row's cells.

```ts
interface ColumnProperties extends CellStyle {
  /** Stable key naming the column so a keyed-object row (see {@link Worksheet.addRow}) can place a
   * value under it by name rather than position. In-memory only: it is not serialized to OOXML. */
  key?: string;
  /** Column width in character units. */
  width?: number;
  /** Whether the column is hidden. */
  hidden?: boolean;
  /** Outline (grouping) depth; 0 or absent means ungrouped. */
  outlineLevel?: number;
  /** Whether this column is the collapsed summary of an outline group. */
  collapsed?: boolean;
}
```

---

### `isVisibility`

<sub>const</sub>

Narrow a raw `<sheet state>` or `<workbookView visibility>` token to a known [`Visibility`](./worksheet.md#visibility).

```ts
const isVisibility: (value: string) => value is Visibility
```

---

### `OutlineProperties`

<sub>interface</sub>

Placement of an outline's summary rows/columns. Excel's defaults are summary *below* the detail
rows and *right* of the detail columns; setting either to `false` inverts that placement so an
author who groups upward gets a file that honours it. An unset flag is omitted from the written
`<outlinePr>`, and an empty object emits no `<outlinePr>` at all.

```ts
interface OutlineProperties {
  summaryBelow?: boolean;
  summaryRight?: boolean;
}
```

---

### `RowInput`

<sub>type</sub>

A row handed to [`Worksheet.addRow`](./worksheet.md#worksheetaddrow): a positional array of cell values (a hole or `undefined`
leaves that column untouched), or an object keyed by column [`ColumnProperties.key`](./worksheet.md#columnproperties) whose
values land under the matching columns.

```ts
type RowInput = (CellValue | undefined)[] | Record<string, CellValue>;
```

---

### `RowProperties`

<sub>interface</sub>

Per-row formatting. A row may exist purely to carry these, with no cells.

```ts
interface RowProperties {
  /** Row height in points. */
  height?: number;
  /** Whether the row is hidden. */
  hidden?: boolean;
  /** Outline (grouping) depth; 0 or absent means ungrouped. */
  outlineLevel?: number;
  /** Whether this row is the collapsed summary of an outline group. */
  collapsed?: boolean;
  /** Background fill applied to the row's cells that carry no fill of their own. */
  fill?: Fill;
}
```

---

### `SheetView`

<sub>interface</sub>

A worksheet's frozen-pane view. `state` `'frozen'` locks the top `ySplit` rows and left `xSplit`
columns in place while the rest scrolls; `'normal'` (the default) has no split and emits no
`<pane>`: writing a normal view leaves no leftover pane markup that would trip Excel's repair
prompt. An empty object is a normal view.

```ts
interface SheetView {
  /** Freeze state. Absent or `'normal'` means no split. */
  state?: 'normal' | 'frozen';
  /** Number of columns frozen at the left; `0`/absent freezes no columns. */
  xSplit?: number;
  /** Number of rows frozen at the top; `0`/absent freezes no rows. */
  ySplit?: number;
  /** The cell anchoring the bottom-right scrolling pane; defaults to the first unfrozen cell. */
  topLeftCell?: string;
  /**
   * Whether the grid's cell lines are drawn on screen. Absent means Excel's default, which is on,
   * so only `false` is worth setting: a sheet that paints its own fills and borders reads as a
   * document rather than a spreadsheet once the grid behind it is off.
   *
   * On screen only, and unrelated to {@link PrintOptions.gridLines}, which decides whether the grid
   * is *printed*. Excel exposes them as two separate checkboxes because the answers differ.
   */
  showGridLines?: boolean;
}
```

---

### `Visibility`

<sub>type</sub>

Whether a thing Excel can hide is showing: a sheet's tab, or the document window itself.

One type for two schema enumerations. `ST_SheetState` and `ST_Visibility` are declared separately
in ECMA-376 and carry the same three tokens with the same meanings, and `veryHidden` means the same
thing in both: hidden, and not offered in the unhide list.

```ts
type Visibility = 'visible' | 'hidden' | 'veryHidden';
```

---

### `Worksheet`

<sub>class</sub>

```ts
class Worksheet {
  readonly name: string;
  readonly id: number;
  state: WorksheetState['state'];
  tabColor: Color | undefined;
  codeName?: string;
  readonly properties: WorksheetProperties = {};
  readonly outline: OutlineProperties = {};
  readonly view: SheetView = {};
  readonly pageSetup: PageSetup = {};
  readonly printOptions: PrintOptions = {};
  readonly pageMargins: PageMargins = {};
  readonly headerFooter: HeaderFooter = {};
  readonly rowBreaks: PageBreak[] = [];
  readonly columnBreaks: PageBreak[] = [];
  getCell(reference: string): Cell;
  hasCell(row: number, col: number): boolean;
  getColumn(index: number): Column;
  getRow(number: number): Row;
  getRange(reference: string): Range;
  getRange(top: number, left: number, bottom: number, right: number): Range;
  getRange(referenceOrTop: string | number, left?: number, bottom?: number, right?: number): Range;
  get rowCount(): number;
  get actualRowCount(): number;
  get columnCount(): number;
  get usedRange(): Range | undefined;
  *columns(): IterableIterator<Column>;
  *rows(): IterableIterator<Row>;
  addTable(options: TableOptions): Table;
  get tables(): readonly Table[];
  getTable(name: string): Table | undefined;
  addPivotTable(options: PivotTableOptions): PivotTable;
  get pivotTables(): readonly PivotTable[];
  get loadedPivotTables(): readonly ParsedPivotTable[];
  addCommentThread(thread: CommentThread): void;
  get commentThreads(): readonly CommentThread[];
  commentThreadAt(reference: string): CommentThread | undefined;
  addImage(
    imageId: number,
    anchor: {readonly tl: AnchorPoint; readonly br: AnchorPoint; readonly editAs?: ImageEditAs},
  ): void;
  addImage(
    imageId: number,
    anchor: {
      readonly tl: AnchorPoint;
      readonly ext: {readonly width: number; readonly height: number};
    },
  ): void;
  addImage(imageId: number, anchor: PixelAnchor): void;
  addImageAnchor(imageId: number, anchor: ImageAnchor): void;
  removeImage(imageId: number): void;
  get images(): readonly AnchoredImage[];
  addBackgroundImage(imageId: number): void;
  removeBackgroundImage(): void;
  get backgroundImageId(): number | undefined;
  get preservedReferences(): readonly PreservedWorksheetReference[];
  mergeCells(range: string): void;
  get merges(): readonly string[];
  get autoFilter(): AutoFilter | undefined;
  set autoFilter(filter: string | AutoFilter | undefined);
  unmergeCells(range: string): boolean;
  addDataValidation(sqref: string, rule: DataValidation, options: {extended?: boolean} = {}): void;
  get dataValidations(): readonly DataValidationEntry[];
  addConditionalFormatting(formatting: ConditionalFormatting): void;
  get conditionalFormattings(): readonly ConditionalFormatting[];
  dataValidationAt(reference: string): DataValidation | undefined;
  spliceRows(start: number, count: number, ...inserts: RowInput[]): void;
  insertRow(pos: number, values: RowInput): void;
  addRow(values: RowInput): Cell[];
  addRows(rows: RowInput[]): Cell[][];
  freeze(ySplit = 1, xSplit = 0): void;
  unfreeze(): void;
  duplicateRow(start: number, options: {count?: number; insert?: boolean} = {}): void;
  spliceColumns(start: number, count: number, ...inserts: CellValue[][]): void;
  insertColumn(pos: number, values: CellValue[]): void;
  addColumn(values: CellValue[]): Cell[];
  addColumns(columns: CellValue[][]): Cell[][];
  get model(): WorksheetModel;
  set model(model: WorksheetModel);
  protect(password?: string, options: SheetProtectionOptions = {}): void;
  unprotect(): void;
  get protection(): SheetProtection | undefined;
}
```

**Members**

#### `Worksheet.id`

```ts
readonly id: number;
```

1-based workbook-assigned id, stable for the sheet's lifetime.

#### `Worksheet.tabColor`

```ts
tabColor: Color | undefined;
```

Colour of the sheet's tab, as an ARGB/theme [`Color`](./styles.md#color). `undefined` leaves the tab its
default colour; the writer emits no `<tabColor>` for an uncoloured sheet, so a round-trip
never fabricates one.

#### `Worksheet.codeName`

```ts
codeName?: string;
```

The sheet's VBA identity (`<sheetPr codeName>`), the name a macro means by `Sheet1`. Undefined
for a sheet in a workbook with no VBA project, which is what Excel writes for one.

Preserved rather than modeled: nothing here reads it, but a `.xlsm` whose sheet code names are
dropped on a round trip has had the binding between its macros and its sheets cut. Deliberately
absent from [`WorksheetModel`](./worksheet.md#worksheetmodel), which is the copy shape: a code name identifies *this* sheet
to the workbook's VBA project, so copying it onto a second sheet would give the project two
sheets answering to one name.

#### `Worksheet.properties`

```ts
readonly properties: WorksheetProperties = {};
```

Sheet-level format defaults. Mutate in place: `sheet.properties.defaultRowHeight = 20`.

#### `Worksheet.outline`

```ts
readonly outline: OutlineProperties = {};
```

Outline summary-position flags. Mutate in place: `sheet.outline.summaryBelow = false`. Empty
means unset: the writer emits no `<outlinePr>` and a round-trip never fabricates one.

#### `Worksheet.view`

```ts
readonly view: SheetView = {};
```

The sheet's frozen-pane view. Empty (a normal view) emits no `<pane>`. Use [`freeze`](./worksheet.md#worksheetfreeze) and
[`unfreeze`](./worksheet.md#worksheetunfreeze) for the common cases, or mutate in place for finer control.

#### `Worksheet.pageSetup`

```ts
readonly pageSetup: PageSetup = {};
```

Print-scaling and orientation. Mutate in place: `sheet.pageSetup.fitToPage = true`. Empty means
unset: the writer emits neither `<pageSetUpPr>` nor `<pageSetup>` and a round-trip never
fabricates them.

#### `Worksheet.printOptions`

```ts
readonly printOptions: PrintOptions = {};
```

Print-toggle flags (`<printOptions>`): centring, and whether headings/gridlines print. Mutate in
place: `sheet.printOptions.gridLines = true`. Empty means unset. The writer emits no element and
a round-trip never fabricates one.

#### `Worksheet.pageMargins`

```ts
readonly pageMargins: PageMargins = {};
```

Print margins. Mutate in place: `sheet.pageMargins.left = 0.5`. Empty means unset.

#### `Worksheet.headerFooter`

```ts
readonly headerFooter: HeaderFooter = {};
```

Page header/footer text. Mutate in place: `sheet.headerFooter.oddHeader = '&C&"..."'`.

#### `Worksheet.rowBreaks`

```ts
readonly rowBreaks: PageBreak[] = [];
```

Horizontal page breaks (`<rowBreaks>`): each break's `id` is the last row before it, so
`sheet.rowBreaks.push({id: 3})` starts a new printed page at row 4. Mutate in place. A row splice
moves a break with the row after it. Empty means no row breaks and the writer emits no
`<rowBreaks>` element.

#### `Worksheet.columnBreaks`

```ts
readonly columnBreaks: PageBreak[] = [];
```

Vertical page breaks (`<colBreaks>`): each break's `id` is the last column before it, so
`sheet.columnBreaks.push({id: 3})` starts a new printed page at column D. Mutate in place. A column
splice moves a break with the column after it. Empty means no column breaks and the writer emits
no `<colBreaks>` element.

#### `Worksheet.getCell`

```ts
getCell(reference: string): Cell;
```

Get the cell at an A1 reference, creating it on first access. The reference must
name both a column and a row (`"B3"`); a whole-row or whole-column reference is
not a cell and is rejected.

Addressing a cell covered by a merged region resolves to that region's master
(top-left) cell, mirroring how a spreadsheet treats the merge as one cell: a value
or style written through a covered address lands on the master, and reading a
covered address returns the master's. Only the master ever holds an independent
value, so the serialized sheet stays well-formed (no stray value on a covered cell).

**Throws:** `SyntaxError` if the reference does not resolve to a single cell.

#### `Worksheet.hasCell`

```ts
hasCell(row: number, col: number): boolean;
```

Whether a cell has been materialised at the given 1-based position.

#### `Worksheet.getColumn`

```ts
getColumn(index: number): Column;
```

A handle on a 1-based column: its formatting, its cells, and its values. Cheap and stateless.
It creates neither cells nor a format record, so asking about a column costs nothing and does
not extend the used range. Writing through it (`getColumn(2).width = 12`) is what materialises
the record.

**Throws:** `RangeError` if the index is not a positive integer.

#### `Worksheet.getRow`

```ts
getRow(number: number): Row;
```

A handle on a 1-based row: its formatting, its cells, and its values. Cheap and stateless. It
creates neither cells nor a format record, so asking about a row costs nothing and does not
extend the used range. Writing through it (`getRow(3).height = 20`) is what materialises the
record.

**Throws:** `RangeError` if the number is not a positive integer.

#### `Worksheet.getRange`

```ts
getRange(reference: string): Range;
getRange(top: number, left: number, bottom: number, right: number): Range;
```

A handle on a rectangular block of cells: `getRange('B2:D5')`, or the same block by its
inclusive corners as `getRange(2, 2, 5, 4)`. Cheap and stateless like [`getRow`](./worksheet.md#worksheetgetrow) and
[`getColumn`](./worksheet.md#worksheetgetcolumn): it creates no cells and does not extend the used range.

Corners are stated **first and last, inclusive**, in either order, never as a start and a count.
That is the convention for every range-shaped accessor here, so the three axes cannot disagree
about what a pair of numbers means.

A whole-row (`'1:1'`) or whole-column (`'A:A'`) reference is refused rather than accepted as a
million-cell block: OOXML states a whole-axis default in one attribute, and [`getRow`](./worksheet.md#worksheetgetrow) /
[`getColumn`](./worksheet.md#worksheetgetcolumn) are how you write it.

**Throws:** `SyntaxError` if the reference is unparseable, names another worksheet, or leaves an
axis unbounded.
**Throws:** [`AuthoringError`](./errors.md#authoringerror) if the numeric form is called with fewer than four corners.
**Throws:** `RangeError` if a numeric corner is not a positive integer within the sheet's bounds.

#### `Worksheet.rowCount`

```ts
get rowCount(): number;
```

The 1-based index of the last row carrying anything (data or its own formatting),
or 0 for an empty sheet. Spans gaps: a value in row 5 makes this 5 even if rows 2–4
are empty. This is the used-range extent, not a populated-row tally (see
[`actualRowCount`](./worksheet.md#worksheetactualrowcount)).

*Carrying anything* is deliberately wide: a value, a cell's own style facet, quote-prefix flag
or named-style link, a note, a row height or outline level, or a merge reaching down. Styling a
band of empty rows is how a template is laid out, so those rows are used and [`addRow`](./worksheet.md#worksheetaddrow)
appends past them. A cell merely materialised by [`getCell`](./worksheet.md#worksheetgetcell) and left untouched carries
nothing, so reading a far address never grows the sheet.

#### `Worksheet.actualRowCount`

```ts
get actualRowCount(): number;
```

The number of rows that hold at least one non-empty cell, ignoring gaps and formatting-only rows.

#### `Worksheet.columnCount`

```ts
get columnCount(): number;
```

The 1-based index of the last column carrying anything (a used cell or the column's own format
properties), or 0 for an empty sheet. The used-range width, mirroring [`rowCount`](./worksheet.md#worksheetrowcount) for the
other axis, down to what *carrying anything* means: a value in column E makes this 5 even if
columns B–D are empty, and a column holding nothing but a styled empty cell is still used.

#### `Worksheet.usedRange`

```ts
get usedRange(): Range | undefined;
```

The sheet's used range as one handle: `A1` through the last row and column that carry
anything, or `undefined` when there is no rectangle to name.

This is [`rowCount`](./worksheet.md#worksheetrowcount) and [`columnCount`](./worksheet.md#worksheetcolumncount) said once, so a caller stops reassembling
`A1:${numberToColumn(sheet.columnCount)}${sheet.rowCount}` by hand. That is what an
[`autoFilter`](./worksheet.md#worksheetautofilter) covering the whole sheet wants (`sheet.autoFilter = sheet.usedRange.address`),
and Excel writes exactly that ref for a filter it applies itself. A header-only ref filters
nothing, which is the bug this exists to make hard to write.

It inherits both counts' definition of *used*, so it spans gaps (a value in `E5` and nothing
else still gives `A1:E5`) and includes a line carrying only its own formatting: a set column
width, an outline level, a merge reaching past the last value. `undefined` therefore means
strictly "no rectangle": an empty sheet, or one carrying only row formatting and no columns at
all (or the reverse), where an axis has no extent to bound the other against.

Not the same thing as the `<dimension>` a written package records. That is the *tight* box,
top-left at the first used cell and formatting-only rows excluded, because Excel writes it to
describe where the data is, not what the grid spans. This handle is anchored at `A1`, because
a caller asking for the used range means the block to read, style or filter.

#### `Worksheet.columns`

```ts
*columns(): IterableIterator<Column>;
```

The columns carrying format properties, as handles, in ascending index order.

#### `Worksheet.rows`

```ts
*rows(): IterableIterator<Row>;
```

The rows to serialise, as handles, in ascending row order: the union of rows holding cells and
rows holding only metadata (a hidden or grouped row need carry no data). Mirrors how OOXML
serialises (`<row>` wrapping `<c>`) and is the writer's row surface.

A handle yields its cells only when asked, so a pass that reads nothing but row attributes
never assembles a cell array it will not look at.

#### `Worksheet.addTable`

```ts
addTable(options: TableOptions): Table;
```

Define a table over a range of this sheet. The table's shape invariants (a legal
name, at least one column, at least one row) are enforced here; conflicts with the
rest of the sheet (e.g. an overlapping merge) are the writer's concern.

**Throws:** [`AuthoringError`](./errors.md#authoringerror) if the name, columns, or geometry are invalid.

#### `Worksheet.tables`

```ts
get tables(): readonly Table[];
```

The tables defined on this sheet, in definition order.

#### `Worksheet.getTable`

```ts
getTable(name: string): Table | undefined;
```

The table with the given name (case-sensitive, the identifier Excel uses), or `undefined`.
A table read back from a file is fully hydrated: its rows can be read and appended to.

#### `Worksheet.addPivotTable`

```ts
addPivotTable(options: PivotTableOptions): PivotTable;
```

Add a pivot table to this (destination) sheet, summarising a source sheet's data. The source is
read once, now, so the pivot is a snapshot: later edits to the source do not change it. The
supported shape (one summed value field, at least one row and column field) is enforced here.

**Throws:** [`AuthoringError`](./errors.md#authoringerror) if the metric, fields, or source shape are unsupported.

#### `Worksheet.pivotTables`

```ts
get pivotTables(): readonly PivotTable[];
```

The pivot tables hosted on this sheet, in definition order.

#### `Worksheet.loadedPivotTables`

```ts
get loadedPivotTables(): readonly ParsedPivotTable[];
```

Pivot tables reconstructed from a loaded package, in the order the reader found them: a
read-only inspection view (source range, field roles, value field, aggregation). A pivot
authored on this sheet via [`addPivotTable`](./worksheet.md#worksheetaddpivottable) does not appear here; a pivot loaded from a
file does not appear in [`pivotTables`](./worksheet.md#worksheetpivottables). The loaded pivots re-emit verbatim through
byte-preservation, so this collection is never itself serialised.

#### `Worksheet.addCommentThread`

```ts
addCommentThread(thread: CommentThread): void;
```

Anchor a threaded conversation to a cell. This is Excel's modern review comment: an opening
message, its replies, and whether the discussion was marked resolved. Distinct from a cell's
legacy note ([`Cell.note`](./cell.md#cellnote)), and mutually exclusive with one: Excel refuses to put both on
one cell, and a cell carrying both is written back as the conversation alone.

Every message supplies its own [`Comment.id`](./comment-thread.md#comment) and [`Comment.date`](./comment-thread.md#comment), and names its author by
[`Comment.personId`](./comment-thread.md#comment) into the workbook registry ([`Workbook.addPerson`](./workbook.md#workbookaddperson)): the writer has no
clock and no id generator, so nothing here is invented and the same workbook always serialises to the
same bytes. An author id nobody registered is refused when the workbook is written, where the
registry is complete; a message with no `personId` is written as Excel writes an unknown author. Every id is normalised to the brace-wrapped upper-case GUID form the format requires, so a
`crypto.randomUUID()` is accepted as-is.

Message ids must be unique **within this sheet**, because that is the scope in which they mean
anything: a reply names its thread by the head's id inside the sheet's own part, and the legacy
fallback comment binds its cell by the same id inside the sheet's own comments part. Two sheets reusing
one id is therefore harmless and is not rejected: Excel's ids happen to be globally unique, but
nothing resolves across a part boundary.

**Throws:** `SyntaxError` if the anchor does not resolve to a single cell, if any id is not a GUID, if a
message id is already used on this sheet, or if a mention's span is not a whole number the wire can
express.

#### `Worksheet.commentThreads`

```ts
get commentThreads(): readonly CommentThread[];
```

The threaded conversations on this sheet: Excel's modern review comments (author, timestamp,
replies, resolved state, `@mentions`). Empty for a sheet with none. Distinct from a cell's legacy note
([`Cell.note`](./cell.md#cellnote)).

#### `Worksheet.commentThreadAt`

```ts
commentThreadAt(reference: string): CommentThread | undefined;
```

The conversation anchored to a cell, or `undefined` when that cell carries none. The reference is
canonicalized, so an absolute `"$B$2"` finds the same thread as `"B2"`; it names the *anchor* cell,
so a cell merely covered by the anchor's merged region is not a match.

**Throws:** `SyntaxError` if the reference does not resolve to a single cell.

#### `Worksheet.addImage`

```ts
addImage(
    imageId: number,
    anchor: {readonly tl: AnchorPoint; readonly br: AnchorPoint; readonly editAs?: ImageEditAs},
  ): void;
addImage(
    imageId: number,
    anchor: {
      readonly tl: AnchorPoint;
      readonly ext: {readonly width: number; readonly height: number};
    },
  ): void;
```

Anchor a workbook image (the id returned by [`Workbook.addImage`](./workbook.md#workbookaddimage)) to this sheet. Two shapes:

- **Two-cell**: `{tl, br}` spans the rectangle from the top-left grid point to the bottom-right,
  reflowing as the spanned cells resize. `editAs` (`oneCell` by default) tunes how it follows.
- **One-cell**: `{tl, ext}` pins the image at `tl` at a fixed pixel size that the grid never
  resizes. `ext` is in pixels and converts to EMUs internally.

Grid points are 0-based (`{col: 0, row: 0}` is cell A1). A later row/column splice re-pins the
anchor to the same logical position.

#### `Worksheet.addImageAnchor`

```ts
addImageAnchor(imageId: number, anchor: ImageAnchor): void;
```

Anchor an image with a pre-built model anchor in the model's own units (EMUs). This is the
low-level primitive [`addImage`](./worksheet.md#worksheetaddimage) builds on and the reader uses to re-pin an image parsed from
a drawing part without a lossy pixel round-trip.

#### `Worksheet.removeImage`

```ts
removeImage(imageId: number): void;
```

Drop every anchor of the given workbook image from this sheet. The image stays registered on the
workbook (another sheet may still show it), so only this sheet's anchors are removed; the writer
then omits any media no sheet anchors any longer.

#### `Worksheet.images`

```ts
get images(): readonly AnchoredImage[];
```

The images anchored to this sheet, in the order they were added.

#### `Worksheet.addBackgroundImage`

```ts
addBackgroundImage(imageId: number): void;
```

Set this sheet's background image to a workbook image (the id [`Workbook.addImage`](./workbook.md#workbookaddimage) returned).
The picture tiles behind the whole grid; it is not anchored to any cell. Passing a new id replaces
the previous background.

#### `Worksheet.removeBackgroundImage`

```ts
removeBackgroundImage(): void;
```

Remove this sheet's background image, if any. The image stays registered on the workbook.

#### `Worksheet.backgroundImageId`

```ts
get backgroundImageId(): number | undefined;
```

The workbook image id set as this sheet's background, or `undefined` when it has none.

#### `Worksheet.preservedReferences`

```ts
get preservedReferences(): readonly PreservedWorksheetReference[];
```

The worksheet-level references to unmodeled package content preserved for round-tripping.

#### `Worksheet.mergeCells`

```ts
mergeCells(range: string): void;
```

Merge a range of cells (`"A1:B2"`). A range that overlaps an already-merged region is
rejected: Excel forbids overlapping merges and writes such geometry as a corrupt file.
Whole-row/column ranges (`"A:A"`) are unbounded, carry no rectangle, and are not overlap-checked.

Any value already sitting in a covered non-anchor cell is discarded, keeping only the top-left
anchor's, exactly how Excel collapses a range on merge. Leaving it would emit a populated `<c>`
under the `<mergeCell>` ref, the geometry Excel opens with a repair prompt. Covered-cell styles
survive (a border spanning the merge is legal), so only the conflicting value is cleared.

The range is stored in canonical form, `$` anchors dropped and the top-left corner first, so
merging `$B$2:$A$1` reports `A1:B2` through [`merges`](./worksheet.md#worksheetmerges).

**Throws:** `SyntaxError` if the range is unparseable or names a worksheet: a merge belongs to the
sheet it is made on, and a prefix naming another one used to be ignored.
**Throws:** [`AuthoringError`](./errors.md#authoringerror) if the range overlaps an already-merged region.

#### `Worksheet.merges`

```ts
get merges(): readonly string[];
```

The merged ranges on this sheet, in canonical form (`A1:B2`), in the order they were added.

#### `Worksheet.autoFilter`

```ts
get autoFilter(): AutoFilter | undefined;
set autoFilter(filter: string | AutoFilter | undefined);
```

The sheet's autofilter (its range plus any per-column criteria), or `undefined` when the sheet
carries none. Setting one turns on the header-row filter dropdowns Excel draws over the range;
the writer emits both the sheet's `<autoFilter>` element and the hidden `_FilterDatabase` defined
name Excel derives from it. Setting `undefined` clears the filter.

A bare range string is the ergonomic common case: `sheet.autoFilter = 'A1:C10'` for dropdowns
with no active criteria; pass an [`AutoFilter`](./autofilter.md#autofilter) object to narrow columns. Either way the
value is normalised on assignment (range to canonical `A1:C10` form) and the getter returns the
structured object. The range must be a bounded rectangle: a whole-row/column reference is not a
filterable region and is rejected.

#### `Worksheet.unmergeCells`

```ts
unmergeCells(range: string): boolean;
```

Remove a merged range previously added with [`mergeCells`](./worksheet.md#worksheetmergecells), returning whether such a merge
existed. The range matches however it is spelled (`$` anchors, corner order), since merges are
stored and compared in canonical form. The covering rectangle is dropped alongside it, so a cell
the merge had masked addresses independently again. The inverse of [`mergeCells`](./worksheet.md#worksheetmergecells).

**Throws:** `SyntaxError` if the range is unparseable or names a worksheet.

#### `Worksheet.addDataValidation`

```ts
addDataValidation(sqref: string, rule: DataValidation, options: {extended?: boolean} = {}): void;
```

Attach a data validation to a target range (`"B2:B20"`, a whole column `"B2:B1048576"`, or a
space-separated `sqref` of several ranges). The rule is stored once against the range, not copied
per covered cell, so a whole-column dropdown stays a single entry. A cell inside the range reports
the rule through [`dataValidationAt`](./worksheet.md#worksheetdatavalidationat).

Pass `{extended: true}` to mark a rule that belongs in the 2009 extension form
(`<x14:dataValidation>`), the carrier Excel uses for a list source on another sheet and other
shapes the standard element cannot express. The reader sets it for a rule found in that form so a
round-trip writes it back there instead of silently corrupting the cross-sheet reference.

#### `Worksheet.dataValidations`

```ts
get dataValidations(): readonly DataValidationEntry[];
```

The data validations on this sheet, each bound to its target range, in insertion order.

#### `Worksheet.addConditionalFormatting`

```ts
addConditionalFormatting(formatting: ConditionalFormatting): void;
```

Attach a conditional formatting to a target range. `formatting.ref` is an OOXML `sqref`: one
range (`"A1:A10"`), a whole column, or several space-separated areas (`"A1:C1 A3:C3"`) sharing one
rule set. The block is stored once against the range, defensively copied so the getter never hands
back a reference into the caller's object.

**Throws:** [`AuthoringError`](./errors.md#authoringerror) when `formatting.ref` names no cells: empty, or no area of it decodes.

#### `Worksheet.conditionalFormattings`

```ts
get conditionalFormattings(): readonly ConditionalFormatting[];
```

The conditional formattings on this sheet, each bound to its target range, in insertion order.

#### `Worksheet.dataValidationAt`

```ts
dataValidationAt(reference: string): DataValidation | undefined;
```

The validation covering a cell, or `undefined` when none does. The first added rule whose range
contains the cell wins, mirroring how a spreadsheet resolves overlapping validations.

#### `Worksheet.spliceRows`

```ts
spliceRows(start: number, count: number, ...inserts: RowInput[]): void;
```

Remove `count` rows starting at the 1-based `start`, then insert the given rows in their place.
Rows below the edit shift by `inserts.length - count`: a delete pulls the tail up, an insert
pushes it down, and doing both at once is a replace. Each inserted row takes either
[`RowInput`](./worksheet.md#rowinput) shape, a positional array from column A or a key-addressed object, exactly
like [`addRow`](./worksheet.md#worksheetaddrow). A `count` larger than the rows present simply clears the tail; it never
silently becomes a no-op. Cells carry their full style to the shifted position, and merged ranges
shift with the rows they cover.

**Throws:** `RangeError` if `start` is not a positive integer or `count` is negative.
**Throws:** `RangeError` if an inserted row would land past the last row of the grid. The sheet is
left untouched, so this is a refused edit rather than half of one: a region pushed off the edge
clamps and absorbs the loss, but content pushed off it is what Excel refuses outright.

#### `Worksheet.insertRow`

```ts
insertRow(pos: number, values: RowInput): void;
```

Insert one row of `values` at the 1-based `pos`, shifting the rows at and below it down by one.
`values` takes either [`RowInput`](./worksheet.md#rowinput) shape (positional array or keyed object), like
[`addRow`](./worksheet.md#worksheetaddrow). Shorthand for [`spliceRows`](./worksheet.md#worksheetsplicerows)`(pos, 0, values)`.

**Throws:** `RangeError` if `pos` is not a positive integer.

#### `Worksheet.addRow`

```ts
addRow(values: RowInput): Cell[];
```

Append a row of `values` after the last used row, returning the cells it materialised.
The append point is [`rowCount`](./worksheet.md#worksheetrowcount)` + 1`, so the row lands below every row that holds
data or its own formatting, never overwriting existing content, unlike [`insertRow`](./worksheet.md#worksheetinsertrow),
which shifts and needs a position. Unlike [`spliceRows`](./worksheet.md#worksheetsplicerows), appending shifts nothing, so
it never disturbs merges or the rows above.

A row takes either shape. A positional array maps its values to columns from A, and a hole in a
sparse array (`['a', , 'c']`) leaves that column untouched. A keyed object lands its values
under the columns carrying the matching [`ColumnProperties.key`](./worksheet.md#columnproperties).

#### `Worksheet.addRows`

```ts
addRows(rows: RowInput[]): Cell[][];
```

Append several rows after the last used row in one call, returning the cells materialised
for each. The rows stack in order: the first lands at [`rowCount`](./worksheet.md#worksheetrowcount)` + 1`, the next
directly below it, so a later row never collides with an earlier one even when both are
value-less. Each row is an array or a keyed object independently, so a mixed batch is fine.
The bulk form of [`addRow`](./worksheet.md#worksheetaddrow).

#### `Worksheet.freeze`

```ts
freeze(ySplit = 1, xSplit = 0): void;
```

Freeze the top `ySplit` rows and left `xSplit` columns in place; the rest of the sheet scrolls
beneath them. `freeze(1)` pins a header row; `freeze(0, 1)` pins the first column. Passing both
zero clears the freeze (equivalent to [`unfreeze`](./worksheet.md#worksheetunfreeze)).

**Throws:** `RangeError` if either split is a negative or non-integer count, or leaves no row or
column of the grid to scroll. The view is left as it was.

#### `Worksheet.unfreeze`

```ts
unfreeze(): void;
```

Clear any frozen split, returning the sheet to a normal (fully scrolling) view.

#### `Worksheet.duplicateRow`

```ts
duplicateRow(start: number, options: {count?: number; insert?: boolean} = {}): void;
```

Copy the row at the 1-based `start`, `options.count` times (default 1). With `options.insert`
(the default) the copies are inserted directly after the source, shifting the rows below, and
any merged range there, down by `count`; otherwise the copies *replace* the rows immediately
below without shifting. A destination row is not overlaid but wholly re-made, so a cell it held
in a column the source leaves empty is dropped, exactly as it would be with a shifting insert.

Each copy is a faithful duplicate of the source: its cell values, its per-cell styles, and its
row properties (height, hidden, outline level, row fill). It carries no merge of its own, so a
range can be merged onto a duplicated row afterwards.

**Throws:** `RangeError` if `start` is not a positive integer or `count` is negative, or if a copy
would land past the last row. The sheet is left untouched.

#### `Worksheet.spliceColumns`

```ts
spliceColumns(start: number, count: number, ...inserts: CellValue[][]): void;
```

Remove `count` columns starting at the 1-based `start`, then insert the given columns in their
place: the column analog of [`spliceRows`](./worksheet.md#worksheetsplicerows). Columns to the right shift by
`inserts.length - count`, keeping their values and styles, and a merged range lying wholly to
the right of the edit re-anchors to its new columns. Each inserted column is an array of values
indexed by row (index 0 → row 1); an empty array inserts a blank column.

**Throws:** `RangeError` if `start` is not a positive integer or `count` is negative.
**Throws:** `RangeError` if an inserted column would land past the last column, or one of its values
past the last row. The sheet is left untouched, so this is a refused edit rather than half of
one: a region pushed off the edge clamps and absorbs the loss, but content pushed off it is
what Excel refuses outright, and [`addColumn`](./worksheet.md#worksheetaddcolumn) refuses the same argument identically.

#### `Worksheet.insertColumn`

```ts
insertColumn(pos: number, values: CellValue[]): void;
```

Insert one column of `values` at the 1-based `pos`, shifting the columns at and right of it over
by one. `values` is an array of values indexed by row (index 0 → row 1), like
[`addColumn`](./worksheet.md#worksheetaddcolumn). Shorthand for [`spliceColumns`](./worksheet.md#worksheetsplicecolumns)`(pos, 0, values)`.

**Throws:** `RangeError` if `pos` is not a positive integer, or if the column would land past the
last column or one of its values past the last row; see [`spliceColumns`](./worksheet.md#worksheetsplicecolumns).

#### `Worksheet.addColumn`

```ts
addColumn(values: CellValue[]): Cell[];
```

Append a column of `values` after the last used column, returning the cells it materialised.
The append point is [`columnCount`](./worksheet.md#worksheetcolumncount)` + 1`, so the column lands right of every column that
holds data or its own formatting, never overwriting existing content, unlike [`insertColumn`](./worksheet.md#worksheetinsertcolumn),
which shifts and needs a position. Unlike [`spliceColumns`](./worksheet.md#worksheetsplicecolumns), appending shifts nothing, so it
never disturbs merges or the columns to its left.

`values` is an array indexed by row (index 0 → row 1); a hole or an explicit `undefined` leaves
that row untouched, mirroring [`addRow`](./worksheet.md#worksheetaddrow)'s positional-array shape. That is the only shape a
column takes: the other [`RowInput`](./worksheet.md#rowinput) form addresses columns by their
[`ColumnProperties.key`](./worksheet.md#columnproperties), and a column's values are indexed by *row*, which carries no key,
so there is nothing on this axis for a keyed object to name.

#### `Worksheet.addColumns`

```ts
addColumns(columns: CellValue[][]): Cell[][];
```

Append several columns after the last used column in one call, returning the cells materialised
for each. The columns stack in order: the first lands at [`columnCount`](./worksheet.md#worksheetcolumncount)` + 1`, the next
directly right of it, so a later column never collides with an earlier one even when both are
value-less. The bulk form of [`addColumn`](./worksheet.md#worksheetaddcolumn).

#### `Worksheet.model`

```ts
get model(): WorksheetModel;
set model(model: WorksheetModel);
```

A snapshot of this sheet's value and overlay content (see [`WorksheetModel`](./worksheet.md#worksheetmodel)). Reading it and
assigning it onto another sheet (`dst.model = src.model`) reproduces the source: merges, cells
and their styles, column/row metadata, tables, the autofilter, protection, the frozen-pane view,
and the page setup all survive, because the getter emits and the setter consumes exactly the same
fields. Identity (`name`, `id`) is not part of the model and is never touched by assignment; nor
are attached parts that carry workbook-level identity (images, pivots, threaded comments,
byte-preserved charts/drawings). See [`WorksheetModel`](./worksheet.md#worksheetmodel) for that boundary.

#### `Worksheet.protect`

```ts
protect(password?: string, options: SheetProtectionOptions = {}): void;
```

Protect the sheet, making the per-cell `locked`/`hidden` flags enforceable. Without a
password the protection is a soft lock any consumer can lift; with one, the password is
salted and hashed on the spot (the plaintext is never retained) so lifting the protection
requires re-supplying it. `options` names which operations stay available to a user while
the sheet is protected; anything unspecified falls to Excel's default for that operation.

Re-protecting replaces any prior protection; [`unprotect`](./worksheet.md#worksheetunprotect) clears it.

#### `Worksheet.unprotect`

```ts
unprotect(): void;
```

Remove any protection previously set by [`protect`](./worksheet.md#worksheetprotect).

#### `Worksheet.protection`

```ts
get protection(): SheetProtection | undefined;
```

The sheet's protection, or `undefined` if the sheet is unprotected.

---

### `WorksheetModel`

<sub>interface</sub>

A serialisable snapshot of a worksheet's value and overlay content: its cells and their styles,
the column/row/page metadata, the frozen-pane view, and the sheet-level overlays (merges, data
validations, conditional formattings, tables, the autofilter, protection). [`Worksheet.model`](./worksheet.md#worksheetmodel)
exports one; assigning it back reproduces that content. The getter and setter cover exactly the
same fields, so a `dst.model = src.model` round-trip drops none of it: an export field the import
ignored would silently lose data, the historical merge-loss failure this contract exists to
prevent. Both directions are driven from one field table (`core/worksheet-model.ts`), which the
compiler proves covers every field below, so adding a field here without wiring it fails the build.

The line between what belongs here and what does not is **workbook-independence**: a field earns
its place when its value means the same thing on any sheet of any workbook. That is the test a new
field is measured against, and applying it is what admitted the autofilter and the frozen-pane view
after each had been omitted for no stated reason (ADR-0005 §2).

Out of scope by design: content carrying workbook-level identity rather than pure sheet state.
That covers anchored and background images (their bytes live on the [`Workbook`](./workbook.md#workbook)), pivot tables
(their source references a live worksheet), threaded comments (their authors are ids into the
workbook's [`Workbook.persons`](./workbook.md#workbookpersons) registry, so a copied conversation would name an author the
destination has never heard of), and byte-preserved parts (charts, vector drawings, slicers) kept
verbatim for round-tripping. These stay with their source sheet; a model assignment neither copies
nor clears them.

```ts
interface WorksheetModel {
  state: WorksheetState['state'];
  tabColor: Color | undefined;
  properties: WorksheetProperties;
  outline: OutlineProperties;
  view: SheetView;
  pageSetup: PageSetup;
  printOptions: PrintOptions;
  pageMargins: PageMargins;
  headerFooter: HeaderFooter;
  rowBreaks: PageBreak[];
  columnBreaks: PageBreak[];
  columns: {index: number; properties: ColumnProperties}[];
  rows: {number: number; properties: RowProperties}[];
  cells: CellModel[];
  merges: string[];
  dataValidations: DataValidationEntry[];
  conditionalFormattings: ConditionalFormatting[];
  tables: TableOptions[];
  autoFilter: AutoFilter | undefined;
  protection: SheetProtection | undefined;
}
```

---

### `WorksheetProperties`

<sub>interface</sub>

Format defaults applied to every row/column that carries no explicit override.

```ts
interface WorksheetProperties {
  /** Height, in points, for rows with no explicit height. */
  defaultRowHeight?: number;
  /** Width, in character units, for columns with no explicit width. */
  defaultColWidth?: number;
}
```

---

### `WorksheetState`

<sub>interface</sub>

```ts
interface WorksheetState {
  /** Sheet visibility, as Excel models it. Defaults to `visible`. */
  readonly state: Visibility;
}
```
