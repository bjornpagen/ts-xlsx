// A worksheet: a sparsely-populated grid of cells addressed by A1 reference.
//
// Storage is sparse by construction: a spreadsheet is mostly empty, so cells
// materialise on first access and only occupied positions cost memory. Column and
// row metadata (widths, heights, visibility, outline grouping) are stored apart from
// the cell grid, because a column or row can carry formatting while holding no cells.

import {AuthoringError, quoted} from '../errors.ts';
import {tokenSet} from '../token-set.ts';
import {assertAxisInBounds, decodeCellRef, encodeAddress, tryDecodeCellRef} from './address.ts';
import {type AutoFilter, canonicalizeAutoFilter} from './autofilter.ts';
import {applyCellStyle, Cell, copyCellContent} from './cell.ts';
import {Column} from './column.ts';
import type {CommentThread} from './comment-thread.ts';
import {ConditionalFormattingOverlay} from './conditional-formatting-overlay.ts';
import type {ConditionalFormatting} from './conditional-formatting.ts';
import {DataValidationOverlay} from './data-validation-overlay.ts';
import type {DataValidation, DataValidationEntry} from './data-validation.ts';
import {type SheetSplice, translateFormula} from './formula-references.ts';
import {GridEdits} from './grid-edits.ts';
import type {
  AnchoredImage,
  AnchorPoint,
  ImageAnchor,
  ImageEditAs,
  PictureProperties,
} from './image.ts';
import {INTERNAL} from './internal.ts';
import {clearCoveredValues} from './merge.ts';
import type {HeaderFooter, PageBreak, PageMargins, PageSetup, PrintOptions} from './page-setup.ts';
import {type ParsedPivotTable, PivotTable, type PivotTableOptions} from './pivot-table.ts';
import type {PreservedWorksheetReference} from './preserved.ts';
import {
  deriveCredential,
  type SheetProtection,
  type SheetProtectionCredential,
  type SheetProtectionOptions,
} from './protection.ts';
import {Range, rangeFrom} from './range.ts';
import {buildRowCells, positionalPlacements, rowPlacements} from './row-input.ts';
import {Row} from './row.ts';
import type {CellContent, CellStyle, Color, Fill} from './style.ts';
import {Table, type TableOptions} from './table.ts';
import {UsedExtent} from './used-extent.ts';
import {type CellValue, isFormulaValue, isSharedFormulaValue} from './value.ts';
import {WorksheetComments} from './worksheet-comments.ts';
import {WorksheetMerges} from './worksheet-merges.ts';
import {WORKSHEET_MODEL_FACETS} from './worksheet-model.ts';
import {type PixelAnchor, WorksheetPictures} from './worksheet-pictures.ts';

/**
 * Whether a thing Excel can hide is showing: a sheet's tab, or the document window itself.
 *
 * One type for two schema enumerations. `ST_SheetState` and `ST_Visibility` are declared separately
 * in ECMA-376 and carry the same three tokens with the same meanings, and `veryHidden` means the same
 * thing in both: hidden, and not offered in the unhide list.
 */
export type Visibility = 'visible' | 'hidden' | 'veryHidden';

/** Narrow a raw `<sheet state>` or `<workbookView visibility>` token to a known {@link Visibility}. */
export const isVisibility = tokenSet<Visibility>({visible: true, hidden: true, veryHidden: true});

export interface WorksheetState {
  /** Sheet visibility, as Excel models it. Defaults to `visible`. */
  readonly state: Visibility;
}

/** Format defaults applied to every row/column that carries no explicit override. */
export interface WorksheetProperties {
  /** Height, in points, for rows with no explicit height. */
  defaultRowHeight?: number;
  /** Width, in character units, for columns with no explicit width. */
  defaultColWidth?: number;
}

/**
 * Placement of an outline's summary rows/columns. Excel's defaults are summary *below* the detail
 * rows and *right* of the detail columns; setting either to `false` inverts that placement so an
 * author who groups upward gets a file that honours it. An unset flag is omitted from the written
 * `<outlinePr>`, and an empty object emits no `<outlinePr>` at all.
 */
export interface OutlineProperties {
  summaryBelow?: boolean;
  summaryRight?: boolean;
}

/**
 * A worksheet's frozen-pane view. `state` `'frozen'` locks the top `ySplit` rows and left `xSplit`
 * columns in place while the rest scrolls; `'normal'` (the default) has no split and emits no
 * `<pane>`: writing a normal view leaves no leftover pane markup that would trip Excel's repair
 * prompt. An empty object is a normal view.
 */
export interface SheetView {
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

/**
 * Per-column formatting. A column may exist purely to carry these, with no cells. The style
 * facets are *defaults* for the column's cells: a cell that sets a facet of its own wins, but
 * one that leaves a facet unset inherits the column's, the same precedence Excel applies, and
 * symmetric with how a {@link RowProperties} fill defaults a row's cells.
 */
export interface ColumnProperties extends CellStyle {
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

/** A row handed to {@link Worksheet.addRow}: a positional array of cell values (a hole or `undefined`
 * leaves that column untouched), or an object keyed by column {@link ColumnProperties.key} whose
 * values land under the matching columns. */
export type RowInput = (CellValue | undefined)[] | Record<string, CellValue>;

/** Per-row formatting. A row may exist purely to carry these, with no cells. */
export interface RowProperties {
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

/**
 * One materialised cell in a {@link WorksheetModel}: its position, value, note, and every facet of its
 * formatting. Extends {@link CellContent} rather than {@link CellStyle} so the quote-prefix flag and
 * the named-style link travel with a model round-trip: they are written and read back like any other
 * facet, and leaving them off the tuple is what made a `dst.model = src.model` drop them.
 */
export interface CellModel extends CellContent {
  readonly row: number;
  readonly col: number;
  value: CellValue;
  note?: string | undefined;
}

/**
 * A serialisable snapshot of a worksheet's value and overlay content: its cells and their styles,
 * the column/row/page metadata, the frozen-pane view, and the sheet-level overlays (merges, data
 * validations, conditional formattings, tables, the autofilter, protection). {@link Worksheet.model}
 * exports one; assigning it back reproduces that content. The getter and setter cover exactly the
 * same fields, so a `dst.model = src.model` round-trip drops none of it: an export field the import
 * ignored would silently lose data, the historical merge-loss failure this contract exists to
 * prevent. Both directions are driven from one field table (`core/worksheet-model.ts`), which the
 * compiler proves covers every field below, so adding a field here without wiring it fails the build.
 *
 * The line between what belongs here and what does not is **workbook-independence**: a field earns
 * its place when its value means the same thing on any sheet of any workbook. That is the test a new
 * field is measured against, and applying it is what admitted the autofilter and the frozen-pane view
 * after each had been omitted for no stated reason (ADR-0005 §2).
 *
 * Out of scope by design: content carrying workbook-level identity rather than pure sheet state.
 * That covers anchored and background images (their bytes live on the {@link Workbook}), pivot tables
 * (their source references a live worksheet), threaded comments (their authors are ids into the
 * workbook's {@link Workbook.persons} registry, so a copied conversation would name an author the
 * destination has never heard of), and byte-preserved parts (charts, vector drawings, slicers) kept
 * verbatim for round-tripping. These stay with their source sheet; a model assignment neither copies
 * nor clears them.
 */
export interface WorksheetModel {
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

export class Worksheet {
  readonly name: string;
  /** 1-based workbook-assigned id, stable for the sheet's lifetime. */
  readonly id: number;
  state: WorksheetState['state'];

  /**
   * Colour of the sheet's tab, as an ARGB/theme {@link Color}. `undefined` leaves the tab its
   * default colour; the writer emits no `<tabColor>` for an uncoloured sheet, so a round-trip
   * never fabricates one.
   */
  tabColor: Color | undefined;

  /**
   * The sheet's VBA identity (`<sheetPr codeName>`), the name a macro means by `Sheet1`. Undefined
   * for a sheet in a workbook with no VBA project, which is what Excel writes for one.
   *
   * Preserved rather than modeled: nothing here reads it, but a `.xlsm` whose sheet code names are
   * dropped on a round trip has had the binding between its macros and its sheets cut. Deliberately
   * absent from {@link WorksheetModel}, which is the copy shape: a code name identifies *this* sheet
   * to the workbook's VBA project, so copying it onto a second sheet would give the project two
   * sheets answering to one name.
   */
  codeName?: string;

  /** Sheet-level format defaults. Mutate in place: `sheet.properties.defaultRowHeight = 20`. */
  readonly properties: WorksheetProperties = {};

  /**
   * Outline summary-position flags. Mutate in place: `sheet.outline.summaryBelow = false`. Empty
   * means unset: the writer emits no `<outlinePr>` and a round-trip never fabricates one.
   */
  readonly outline: OutlineProperties = {};

  /**
   * The sheet's frozen-pane view. Empty (a normal view) emits no `<pane>`. Use {@link freeze} and
   * {@link unfreeze} for the common cases, or mutate in place for finer control.
   */
  readonly view: SheetView = {};

  /**
   * Print-scaling and orientation. Mutate in place: `sheet.pageSetup.fitToPage = true`. Empty means
   * unset: the writer emits neither `<pageSetUpPr>` nor `<pageSetup>` and a round-trip never
   * fabricates them.
   */
  readonly pageSetup: PageSetup = {};

  /**
   * Print-toggle flags (`<printOptions>`): centring, and whether headings/gridlines print. Mutate in
   * place: `sheet.printOptions.gridLines = true`. Empty means unset. The writer emits no element and
   * a round-trip never fabricates one.
   */
  readonly printOptions: PrintOptions = {};

  /** Print margins. Mutate in place: `sheet.pageMargins.left = 0.5`. Empty means unset. */
  readonly pageMargins: PageMargins = {};

  /** Page header/footer text. Mutate in place: `sheet.headerFooter.oddHeader = '&C&"..."'`. */
  readonly headerFooter: HeaderFooter = {};

  /**
   * Horizontal page breaks (`<rowBreaks>`): each break's `id` is the last row before it, so
   * `sheet.rowBreaks.push({id: 3})` starts a new printed page at row 4. Mutate in place. A row splice
   * moves a break with the row after it. Empty means no row breaks and the writer emits no
   * `<rowBreaks>` element.
   */
  readonly rowBreaks: PageBreak[] = [];

  /**
   * Vertical page breaks (`<colBreaks>`): each break's `id` is the last column before it, so
   * `sheet.columnBreaks.push({id: 3})` starts a new printed page at column D. Mutate in place. A column
   * splice moves a break with the column after it. Empty means no column breaks and the writer emits
   * no `<colBreaks>` element.
   */
  readonly columnBreaks: PageBreak[] = [];

  // Row-major sparse storage: row index → (column index → cell). Keeping rows as the
  // outer key makes whole-row iteration cheap and mirrors how OOXML serializes
  // (`<row>` wrapping `<c>`).
  readonly #rows = new Map<number, Map<number, Cell>>();
  // Column and row metadata live apart from the grid so an empty-but-formatted line
  // (a hidden column, a tall header row with no data yet) costs no phantom cells.
  readonly #columns = new Map<number, ColumnProperties>();
  readonly #rowProperties = new Map<number, RowProperties>();
  // Tables and authored pivot tables are sheet-level overlays on the grid, not cell storage.
  readonly #tables: Table[] = [];
  readonly #pivotTables: PivotTable[] = [];
  // Pivot tables reconstructed from a loaded package (see io/xlsx/read-pivot.ts): a read-only,
  // inspection-only view distinct from #pivotTables. A loaded pivot round-trips by byte-preservation
  // (#preservedReferences), which stays its sole emission authority; this collection is never emitted,
  // so exposing it cannot double-emit. Empty for a sheet authored from scratch.
  readonly #loadedPivotTables: ParsedPivotTable[] = [];
  // The sheet's threaded conversations (see io/xlsx/threaded-comments.ts): unlike #loadedPivotTables, the
  // authority the writer serialises from: both the sheet's threadedComment part and the legacy fallback
  // comment that binds each cell to its conversation are derived from this list. Empty for a sheet with no
  // threaded comments.
  readonly #comments = new WorksheetComments(() => this.name);
  // The merged regions: the declared ranges, the rectangles the bounded ones cover, and the index
  // that keeps overlap-checking and covered-address resolution off a linear scan.
  readonly #merges = new WorksheetMerges();
  readonly #images = new WorksheetPictures({
    // A size a column or row does not set defers to the sheet default, then to Excel's own.
    columnWidth: (col) => this.#columns.get(col + 1)?.width ?? this.properties.defaultColWidth,
    rowHeight: (row) =>
      this.#rowProperties.get(row + 1)?.height ?? this.properties.defaultRowHeight,
  });
  // Worksheet-level references to package content the model does not interpret (a vector-shape
  // drawing, a header/footer image), captured verbatim on read so a round-trip re-emits them rather
  // than dropping them. Empty for a sheet authored from scratch.
  readonly #preservedReferences: PreservedWorksheetReference[] = [];
  // Data validations and conditional formattings are sheet-level overlays keyed by range, each owning
  // its own storage/cloning/lookup. See DataValidationOverlay and ConditionalFormattingOverlay.
  readonly #dataValidations = new DataValidationOverlay();
  readonly #conditionalFormattings = new ConditionalFormattingOverlay();
  // Sheet-level protection is a single overlay switch, absent until `protect` is called.
  #protection: SheetProtection | undefined;
  // What moves the formulas outside this sheet when a splice moves its lines: set by the workbook that
  // holds the sheet, absent on one no workbook does.
  #formulaHost: ((edit: SheetSplice) => void) | undefined;
  // The sheet's autofilter (range plus any per-column criteria), absent until one is set. A single
  // sheet-level overlay, distinct from a table's own autofilter; stored canonically so the
  // `<autoFilter>` element and the derived `_FilterDatabase` defined name always agree.
  #autoFilter: AutoFilter | undefined;

  // Structural-edit machinery (row/column splices), sharing this sheet's storage by reference. The
  // public spliceRows/spliceColumns/duplicateRow build the cells an insert introduces, then delegate
  // the shift arithmetic here. Wired in the constructor body, not a field initializer, so it never
  // depends on this field's declaration position relative to the storage fields above: a class field
  // initializer only sees an earlier-declared field's initial value, so if this stayed a field
  // initializer, reordering the class body could silently hand GridEdits a still-undefined map.
  readonly #edits: GridEdits;

  // The used-range extent, maintained as the grid is built rather than rescanned on every read. Wired
  // in the constructor body for the same reason #edits is: it holds the storage maps by reference.
  readonly #extent: UsedExtent;

  constructor(name: string, id: number, state: WorksheetState['state'] = 'visible') {
    this.name = name;
    this.id = id;
    this.state = state;
    this.#extent = new UsedExtent({
      rows: this.#rows,
      rowProperties: this.#rowProperties,
      columns: this.#columns,
      mergeRects: this.#merges.rects,
    });
    this.#edits = new GridEdits({
      rows: this.#rows,
      rowProperties: this.#rowProperties,
      columns: this.#columns,
      merges: this.#merges,
      tables: this.#tables,
      pivotTables: this.#pivotTables,
      images: this.#images.anchors,
      dataValidations: this.#dataValidations,
      conditionalFormattings: this.#conditionalFormattings,
      comments: this.#comments,
      rowBreaks: this.rowBreaks,
      columnBreaks: this.columnBreaks,
      sheetName: () => this.name,
      formulaHost: () => this.#formulaHost,
      autoFilter: {
        get: () => this.#autoFilter,
        set: (next) => {
          this.#autoFilter = next;
        },
      },
    });
  }

  /**
   * Get the cell at an A1 reference, creating it on first access. The reference must
   * name both a column and a row (`"B3"`); a whole-row or whole-column reference is
   * not a cell and is rejected.
   *
   * Addressing a cell covered by a merged region resolves to that region's master
   * (top-left) cell, mirroring how a spreadsheet treats the merge as one cell: a value
   * or style written through a covered address lands on the master, and reading a
   * covered address returns the master's. Only the master ever holds an independent
   * value, so the serialized sheet stays well-formed (no stray value on a covered cell).
   *
   * @throws {SyntaxError} if the reference does not resolve to a single cell.
   */
  getCell(reference: string): Cell {
    const {col, row} = decodeCellRef(reference);
    return this[INTERNAL].masterAt(row, col);
  }

  /** Whether a cell has been materialised at the given 1-based position. */
  hasCell(row: number, col: number): boolean {
    return this.#rows.get(row)?.has(col) ?? false;
  }

  /**
   * A handle on a 1-based column: its formatting, its cells, and its values. Cheap and stateless.
   * It creates neither cells nor a format record, so asking about a column costs nothing and does
   * not extend the used range. Writing through it (`getColumn(2).width = 12`) is what materialises
   * the record.
   *
   * @throws {RangeError} if the index is not a positive integer.
   */
  getColumn(index: number): Column {
    return new Column(this, index);
  }

  /**
   * A handle on a 1-based row: its formatting, its cells, and its values. Cheap and stateless. It
   * creates neither cells nor a format record, so asking about a row costs nothing and does not
   * extend the used range. Writing through it (`getRow(3).height = 20`) is what materialises the
   * record.
   *
   * @throws {RangeError} if the number is not a positive integer.
   */
  getRow(number: number): Row {
    return new Row(this, number);
  }

  /**
   * A handle on a rectangular block of cells: `getRange('B2:D5')`, or the same block by its
   * inclusive corners as `getRange(2, 2, 5, 4)`. Cheap and stateless like {@link getRow} and
   * {@link getColumn}: it creates no cells and does not extend the used range.
   *
   * Corners are stated **first and last, inclusive**, in either order, never as a start and a count.
   * That is the convention for every range-shaped accessor here, so the three axes cannot disagree
   * about what a pair of numbers means.
   *
   * A whole-row (`'1:1'`) or whole-column (`'A:A'`) reference is refused rather than accepted as a
   * million-cell block: OOXML states a whole-axis default in one attribute, and {@link getRow} /
   * {@link getColumn} are how you write it.
   *
   * @throws {SyntaxError} if the reference is unparseable, names another worksheet, or leaves an
   *   axis unbounded.
   * @throws {AuthoringError} if the numeric form is called with fewer than four corners.
   * @throws {RangeError} if a numeric corner is not a positive integer within the sheet's bounds.
   */
  getRange(reference: string): Range;
  getRange(top: number, left: number, bottom: number, right: number): Range;
  getRange(referenceOrTop: string | number, left?: number, bottom?: number, right?: number): Range {
    if (typeof referenceOrTop === 'string') return rangeFrom(this, referenceOrTop);
    // The overloads make a short numeric call unreachable from TypeScript, but this is a public API
    // and JavaScript reaches it. A `?? 0` here manufactured the missing corners, so `getRange(2, 2)`
    // failed with "row 0 is out of bounds" -- an error about a coordinate the caller never wrote,
    // pointing at the grid rather than at the call. Counting the corners says what went wrong.
    const corners = [left, bottom, right].filter((corner) => corner !== undefined).length + 1;
    if (left === undefined || bottom === undefined || right === undefined) {
      throw new AuthoringError(
        `getRange(top, left, bottom, right) needs all four corners; got ${corners}`,
      );
    }
    return new Range(this, referenceOrTop, left, bottom, right);
  }

  /**
   * The 1-based index of the last row carrying anything (data or its own formatting),
   * or 0 for an empty sheet. Spans gaps: a value in row 5 makes this 5 even if rows 2–4
   * are empty. This is the used-range extent, not a populated-row tally (see
   * {@link actualRowCount}).
   *
   * *Carrying anything* is deliberately wide: a value, a cell's own style facet, quote-prefix flag
   * or named-style link, a note, a row height or outline level, or a merge reaching down. Styling a
   * band of empty rows is how a template is laid out, so those rows are used and {@link addRow}
   * appends past them. A cell merely materialised by {@link getCell} and left untouched carries
   * nothing, so reading a far address never grows the sheet.
   */
  get rowCount(): number {
    return this.#extent.lastRow;
  }

  /** The number of rows that hold at least one non-empty cell, ignoring gaps and formatting-only rows. */
  get actualRowCount(): number {
    let count = 0;
    for (const cols of this.#rows.values()) {
      if (this.#rowIsPopulated(cols)) count++;
    }
    return count;
  }

  // Whether any cell materialised in a row holds a value: the tally test behind
  // {@link actualRowCount}, which asks how many rows have data, not how far the grid reaches.
  // Deliberately not the used-range test {@link UsedExtent} applies: a formatting-only row bounds the
  // used range but is not a populated row, and collapsing the two is what let an append land on a
  // row someone had styled.
  #rowIsPopulated(cols: Map<number, Cell>): boolean {
    for (const cell of cols.values()) {
      if (cell.value !== null) return true;
    }
    return false;
  }

  /**
   * The 1-based index of the last column carrying anything (a used cell or the column's own format
   * properties), or 0 for an empty sheet. The used-range width, mirroring {@link rowCount} for the
   * other axis, down to what *carrying anything* means: a value in column E makes this 5 even if
   * columns B–D are empty, and a column holding nothing but a styled empty cell is still used.
   */
  get columnCount(): number {
    return this.#extent.lastColumn;
  }

  /**
   * The sheet's used range as one handle: `A1` through the last row and column that carry
   * anything, or `undefined` when there is no rectangle to name.
   *
   * This is {@link rowCount} and {@link columnCount} said once, so a caller stops reassembling
   * `A1:${numberToColumn(sheet.columnCount)}${sheet.rowCount}` by hand. That is what an
   * {@link autoFilter} covering the whole sheet wants (`sheet.autoFilter = sheet.usedRange.address`),
   * and Excel writes exactly that ref for a filter it applies itself. A header-only ref filters
   * nothing, which is the bug this exists to make hard to write.
   *
   * It inherits both counts' definition of *used*, so it spans gaps (a value in `E5` and nothing
   * else still gives `A1:E5`) and includes a line carrying only its own formatting: a set column
   * width, an outline level, a merge reaching past the last value. `undefined` therefore means
   * strictly "no rectangle": an empty sheet, or one carrying only row formatting and no columns at
   * all (or the reverse), where an axis has no extent to bound the other against.
   *
   * Not the same thing as the `<dimension>` a written package records. That is the *tight* box,
   * top-left at the first used cell and formatting-only rows excluded, because Excel writes it to
   * describe where the data is, not what the grid spans. This handle is anchored at `A1`, because
   * a caller asking for the used range means the block to read, style or filter.
   */
  get usedRange(): Range | undefined {
    const bottom = this.rowCount;
    const right = this.columnCount;
    if (bottom === 0 || right === 0) return undefined;
    return new Range(this, 1, 1, bottom, right);
  }

  /** The columns carrying format properties, as handles, in ascending index order. */
  *columns(): IterableIterator<Column> {
    for (const index of [...this.#columns.keys()].sort((a, b) => a - b)) {
      yield new Column(this, index);
    }
  }

  /**
   * The rows to serialise, as handles, in ascending row order: the union of rows holding cells and
   * rows holding only metadata (a hidden or grouped row need carry no data). Mirrors how OOXML
   * serialises (`<row>` wrapping `<c>`) and is the writer's row surface.
   *
   * A handle yields its cells only when asked, so a pass that reads nothing but row attributes
   * never assembles a cell array it will not look at.
   */
  *rows(): IterableIterator<Row> {
    const numbers = new Set<number>([...this.#rows.keys(), ...this.#rowProperties.keys()]);
    for (const number of [...numbers].sort((a, b) => a - b)) {
      yield new Row(this, number);
    }
  }

  /**
   * Define a table over a range of this sheet. The table's shape invariants (a legal
   * name, at least one column, at least one row) are enforced here; conflicts with the
   * rest of the sheet (e.g. an overlapping merge) are the writer's concern.
   *
   * @throws {AuthoringError} if the name, columns, or geometry are invalid.
   */
  addTable(options: TableOptions): Table {
    const table = new Table(options, {
      // The guard the materialiser reads through must not create the cell it asks about, or touch the
      // extent, so it reads the grid directly rather than through `#cellAt`.
      holdsValue: (row, col) => this.#rows.get(row)?.get(col)?.value != null,
      writeCell: (row, col, value, style) => {
        const cell = this.#cellAt(row, col);
        cell.value = value;
        if (style !== undefined) applyCellStyle(cell, style);
      },
      // Insert one empty grid row at `row`; the splice re-pins this table (growing its data rows) and
      // shifts the totals row and everything below down by one.
      insertRow: (row) => this.spliceRows(row, 0, []),
    });
    this.#tables.push(table);
    return table;
  }

  /** The tables defined on this sheet, in definition order. */
  get tables(): readonly Table[] {
    return this.#tables;
  }

  /** The table with the given name (case-sensitive, the identifier Excel uses), or `undefined`.
   * A table read back from a file is fully hydrated: its rows can be read and appended to. */
  getTable(name: string): Table | undefined {
    return this.#tables.find((table) => table.name === name);
  }

  /**
   * Add a pivot table to this (destination) sheet, summarising a source sheet's data. The source is
   * read once, now, so the pivot is a snapshot: later edits to the source's values do not change it,
   * while a row or column splice of the source sheet moves its {@link PivotTable.sourceRef}. The
   * supported shape (one summed value field, at least one row and column field) is enforced here.
   *
   * @throws {AuthoringError} if the metric, fields, or source shape are unsupported.
   */
  addPivotTable(options: PivotTableOptions): PivotTable {
    const pivot = new PivotTable(options);
    this.#pivotTables.push(pivot);
    return pivot;
  }

  /** The pivot tables hosted on this sheet, in definition order. */
  get pivotTables(): readonly PivotTable[] {
    return this.#pivotTables;
  }

  /**
   * Pivot tables reconstructed from a loaded package, in the order the reader found them: a
   * read-only inspection view (source range, field roles, value field, aggregation). A pivot
   * authored on this sheet via {@link addPivotTable} does not appear here; a pivot loaded from a
   * file does not appear in {@link pivotTables}. The loaded pivots re-emit verbatim through
   * byte-preservation, so this collection is never itself serialised.
   */
  get loadedPivotTables(): readonly ParsedPivotTable[] {
    return this.#loadedPivotTables;
  }

  /**
   * Anchor a threaded conversation to a cell. This is Excel's modern review comment: an opening
   * message, its replies, and whether the discussion was marked resolved. Distinct from a cell's
   * legacy note ({@link Cell.note}), and mutually exclusive with one: Excel refuses to put both on
   * one cell, and a cell carrying both is written back as the conversation alone.
   *
   * Every message supplies its own {@link Comment.id} and {@link Comment.date}, and names its author by
   * {@link Comment.personId} into the workbook registry ({@link Workbook.addPerson}): the writer has no
   * clock and no id generator, so nothing here is invented and the same workbook always serialises to the
   * same bytes. An author id nobody registered is refused when the workbook is written, where the
   * registry is complete; a message with no `personId` is written as Excel writes an unknown author. Every id is normalised to the brace-wrapped upper-case GUID form the format requires, so a
   * `crypto.randomUUID()` is accepted as-is.
   *
   * Message ids must be unique **within this sheet**, because that is the scope in which they mean
   * anything: a reply names its thread by the head's id inside the sheet's own part, and the legacy
   * fallback comment binds its cell by the same id inside the sheet's own comments part. Two sheets reusing
   * one id is therefore harmless and is not rejected: Excel's ids happen to be globally unique, but
   * nothing resolves across a part boundary.
   *
   * @throws {SyntaxError} if the anchor does not resolve to a single cell, if any id is not a GUID, if a
   * message id is already used on this sheet, or if a mention's span is not a whole number the wire can
   * express.
   */
  addCommentThread(thread: CommentThread): void {
    this.#comments.add(thread);
  }

  /**
   * The threaded conversations on this sheet: Excel's modern review comments (author, timestamp,
   * replies, resolved state, `@mentions`). Empty for a sheet with none. Distinct from a cell's legacy note
   * ({@link Cell.note}).
   */
  get commentThreads(): readonly CommentThread[] {
    return this.#comments.threads;
  }

  /**
   * The conversation anchored to a cell, or `undefined` when that cell carries none. The reference is
   * canonicalized, so an absolute `"$B$2"` finds the same thread as `"B2"`; it names the *anchor* cell,
   * so a cell merely covered by the anchor's merged region is not a match.
   *
   * @throws {SyntaxError} if the reference does not resolve to a single cell.
   */
  commentThreadAt(reference: string): CommentThread | undefined {
    return this.#comments.at(reference);
  }

  /**
   * Anchor a workbook image (the id returned by {@link Workbook.addImage}) to this sheet. Two shapes:
   *
   * - **Two-cell**: `{tl, br}` spans the rectangle from the top-left grid point to the bottom-right,
   *   reflowing as the spanned cells resize. `editAs` (`oneCell` by default) tunes how it follows.
   * - **One-cell**: `{tl, ext}` pins the image at `tl` at a fixed pixel size that the grid never
   *   resizes. `ext` is in pixels and converts to EMUs internally.
   *
   * Grid points are 0-based (`{col: 0, row: 0}` is cell A1). A later row/column splice re-pins the
   * anchor to the same logical position.
   *
   * `properties` gives the picture alternative text and a title, crops it, or makes it a link; see
   * {@link PictureProperties}.
   */
  addImage(
    imageId: number,
    anchor: {readonly tl: AnchorPoint; readonly br: AnchorPoint; readonly editAs?: ImageEditAs},
    properties?: PictureProperties,
  ): void;
  addImage(
    imageId: number,
    anchor: {
      readonly tl: AnchorPoint;
      readonly ext: {readonly width: number; readonly height: number};
    },
    properties?: PictureProperties,
  ): void;
  addImage(imageId: number, anchor: PixelAnchor, properties?: PictureProperties): void {
    refuseImagesBesideKeptDrawing(this);
    this.#images.add(imageId, anchor, properties);
  }

  /**
   * Anchor an image with a pre-built model anchor in the model's own units (EMUs). This is the
   * low-level primitive {@link addImage} builds on and the reader uses to re-pin an image parsed from
   * a drawing part without a lossy pixel round-trip.
   */
  addImageAnchor(imageId: number, anchor: ImageAnchor, properties?: PictureProperties): void {
    refuseImagesBesideKeptDrawing(this);
    this.#images.addAnchor(imageId, anchor, properties);
  }

  /** Drop every anchor of the given workbook image from this sheet. The image stays registered on the
   * workbook (another sheet may still show it), so only this sheet's anchors are removed; the writer
   * then omits any media no sheet anchors any longer. */
  removeImage(imageId: number): void {
    this.#images.remove(imageId);
  }

  /** The images anchored to this sheet, in the order they were added. */
  get images(): readonly AnchoredImage[] {
    return this.#images.anchors;
  }

  /** Set this sheet's background image to a workbook image (the id {@link Workbook.addImage} returned).
   * The picture tiles behind the whole grid; it is not anchored to any cell. Passing a new id replaces
   * the previous background. */
  addBackgroundImage(imageId: number): void {
    this.#images.setBackground(imageId);
  }

  /** Remove this sheet's background image, if any. The image stays registered on the workbook. */
  removeBackgroundImage(): void {
    this.#images.setBackground(undefined);
  }

  /** The workbook image id set as this sheet's background, or `undefined` when it has none. */
  get backgroundImageId(): number | undefined {
    return this.#images.backgroundImageId;
  }

  /** The worksheet-level references to unmodeled package content preserved for round-tripping. */
  get preservedReferences(): readonly PreservedWorksheetReference[] {
    return this.#preservedReferences;
  }

  /**
   * Merge a range of cells (`"A1:B2"`). A range that overlaps an already-merged region is
   * rejected: Excel forbids overlapping merges and writes such geometry as a corrupt file.
   * Whole-row/column ranges (`"A:A"`) are unbounded, carry no rectangle, and are not overlap-checked.
   *
   * Any value already sitting in a covered non-anchor cell is discarded, keeping only the top-left
   * anchor's, exactly how Excel collapses a range on merge. Leaving it would emit a populated `<c>`
   * under the `<mergeCell>` ref, the geometry Excel opens with a repair prompt. Covered-cell styles
   * survive (a border spanning the merge is legal), so only the conflicting value is cleared.
   *
   * The range is stored in canonical form, `$` anchors dropped and the top-left corner first, so
   * merging `$B$2:$A$1` reports `A1:B2` through {@link merges}.
   *
   * @throws {SyntaxError} if the range is unparseable or names a worksheet: a merge belongs to the
   *   sheet it is made on, and a prefix naming another one used to be ignored.
   * @throws {AuthoringError} if the range overlaps an already-merged region.
   */
  mergeCells(range: string): void {
    // What the merge does to the *grid* stays here, because the grid is this class's: the region
    // slice hands back the rectangle and knows nothing about cells or extents.
    const rect = this.#merges.add(range);
    if (rect !== undefined) {
      this.#extent.noteMerge(rect);
      clearCoveredValues(this.#rows, rect);
    }
  }

  /** The merged ranges on this sheet, in canonical form (`A1:B2`), in the order they were added. */
  get merges(): readonly string[] {
    return this.#merges.ranges;
  }

  /**
   * The sheet's autofilter (its range plus any per-column criteria), or `undefined` when the sheet
   * carries none. Setting one turns on the header-row filter dropdowns Excel draws over the range;
   * the writer emits both the sheet's `<autoFilter>` element and the hidden `_FilterDatabase` defined
   * name Excel derives from it. Setting `undefined` clears the filter.
   *
   * A bare range string is the ergonomic common case: `sheet.autoFilter = 'A1:C10'` for dropdowns
   * with no active criteria; pass an {@link AutoFilter} object to narrow columns. Either way the
   * value is normalised on assignment (range to canonical `A1:C10` form) and the getter returns the
   * structured object. The range must be a bounded rectangle: a whole-row/column reference is not a
   * filterable region and is rejected.
   */
  get autoFilter(): AutoFilter | undefined {
    return this.#autoFilter;
  }

  set autoFilter(filter: string | AutoFilter | undefined) {
    this.#autoFilter = filter === undefined ? undefined : canonicalizeAutoFilter(filter);
  }

  /**
   * Remove a merged range previously added with {@link mergeCells}, returning whether such a merge
   * existed. The range matches however it is spelled (`$` anchors, corner order), since merges are
   * stored and compared in canonical form. The covering rectangle is dropped alongside it, so a cell
   * the merge had masked addresses independently again. The inverse of {@link mergeCells}.
   *
   * @throws {SyntaxError} if the range is unparseable or names a worksheet.
   */
  unmergeCells(range: string): boolean {
    const {existed, rectsChanged} = this.#merges.remove(range);
    // The extent is derived from the rectangles, so an unbounded merge going away leaves it correct.
    if (rectsChanged) this.#extent.invalidate();
    return existed;
  }

  /**
   * Attach a data validation to a target range (`"B2:B20"`, a whole column `"B2:B1048576"`, or a
   * space-separated `sqref` of several ranges). The rule is stored once against the range, not copied
   * per covered cell, so a whole-column dropdown stays a single entry. A cell inside the range reports
   * the rule through {@link dataValidationAt}.
   *
   * Pass `{extended: true}` to mark a rule that belongs in the 2009 extension form
   * (`<x14:dataValidation>`), the carrier Excel uses for a list source on another sheet and other
   * shapes the standard element cannot express. The reader sets it for a rule found in that form so a
   * round-trip writes it back there instead of silently corrupting the cross-sheet reference.
   */
  addDataValidation(sqref: string, rule: DataValidation, options: {extended?: boolean} = {}): void {
    this.#dataValidations.add(sqref, rule, options);
  }

  /** The data validations on this sheet, each bound to its target range, in insertion order. */
  get dataValidations(): readonly DataValidationEntry[] {
    return this.#dataValidations.entries;
  }

  /**
   * Attach a conditional formatting to a target range. `formatting.ref` is an OOXML `sqref`: one
   * range (`"A1:A10"`), a whole column, or several space-separated areas (`"A1:C1 A3:C3"`) sharing one
   * rule set. The block is stored once against the range, defensively copied so the getter never hands
   * back a reference into the caller's object.
   *
   * @throws {AuthoringError} when `formatting.ref` names no cells: empty, or no area of it decodes.
   */
  addConditionalFormatting(formatting: ConditionalFormatting): void {
    this.#conditionalFormattings.add(formatting);
  }

  /** The conditional formattings on this sheet, each bound to its target range, in insertion order. */
  get conditionalFormattings(): readonly ConditionalFormatting[] {
    return this.#conditionalFormattings.entries;
  }

  /**
   * The validation covering a cell, or `undefined` when none does. The first added rule whose range
   * contains the cell wins, mirroring how a spreadsheet resolves overlapping validations.
   */
  dataValidationAt(reference: string): DataValidation | undefined {
    const cell = tryDecodeCellRef(reference);
    return cell === undefined ? undefined : this.#dataValidations.at(cell.col, cell.row);
  }

  /**
   * Remove `count` rows starting at the 1-based `start`, then insert the given rows in their place.
   * Rows below the edit shift by `inserts.length - count`: a delete pulls the tail up, an insert
   * pushes it down, and doing both at once is a replace. Each inserted row takes either
   * {@link RowInput} shape, a positional array from column A or a key-addressed object, exactly
   * like {@link addRow}. A `count` larger than the rows present simply clears the tail; it never
   * silently becomes a no-op. Cells carry their full style to the shifted position, and merged ranges
   * shift with the rows they cover.
   *
   * Formulas move with the rows as Excel moves them: a reference to this sheet, in any sheet's formula,
   * a defined name, a data validation, a conditional format or a table's column formulas, follows the
   * row it names, and one to a deleted row becomes `#REF!`. An authored pivot drawing from this sheet
   * has its source range moved the same way. What the inserted rows carry is written against the sheet after the
   * edit and is not moved.
   *
   * @throws {RangeError} if `start` is not a positive integer or `count` is negative.
   * @throws {RangeError} if an inserted row would land past the last row of the grid. The sheet is
   *   left untouched, so this is a refused edit rather than half of one: a region pushed off the edge
   *   clamps and absorbs the loss, but content pushed off it is what Excel refuses outright.
   */
  spliceRows(start: number, count: number, ...inserts: RowInput[]): void {
    assertStartAndCount('splice', 'row', start, count);
    assertSpliceFits('row', start, inserts.length);
    const inserted = inserts.map((values, i) => buildRowCells(start + i, values, this.#columns));
    this.#edits.spliceRows(start, count, inserted);
    this.#afterStructuralEdit();
  }

  /**
   * Insert one row of `values` at the 1-based `pos`, shifting the rows at and below it down by one.
   * `values` takes either {@link RowInput} shape (positional array or keyed object), like
   * {@link addRow}. Shorthand for {@link spliceRows}`(pos, 0, values)`.
   *
   * @throws {RangeError} if `pos` is not a positive integer.
   */
  insertRow(pos: number, values: RowInput): void {
    this.spliceRows(pos, 0, values);
  }

  /**
   * Append a row of `values` after the last used row, returning the cells it materialised.
   * The append point is {@link rowCount}` + 1`, so the row lands below every row that holds
   * data or its own formatting, never overwriting existing content, unlike {@link insertRow},
   * which shifts and needs a position. Unlike {@link spliceRows}, appending shifts nothing, so
   * it never disturbs merges or the rows above.
   *
   * A row takes either shape. A positional array maps its values to columns from A, and a hole in a
   * sparse array (`['a', , 'c']`) leaves that column untouched. A keyed object lands its values
   * under the columns carrying the matching {@link ColumnProperties.key}.
   */
  addRow(values: RowInput): Cell[] {
    return this.addRows([values])[0] ?? [];
  }

  /**
   * Append several rows after the last used row in one call, returning the cells materialised
   * for each. The rows stack in order: the first lands at {@link rowCount}` + 1`, the next
   * directly below it, so a later row never collides with an earlier one even when both are
   * value-less. Each row is an array or a keyed object independently, so a mixed batch is fine.
   * The bulk form of {@link addRow}.
   */
  addRows(rows: RowInput[]): Cell[][] {
    const first = this.rowCount + 1;
    // Every position resolved and bounded before the first write, so a batch that runs off the grid
    // is refused whole. Writing as it went put `A1` in place before a 16,385-wide row threw.
    const planned = rows.map((values, i) => {
      const row = first + i;
      const placements = rowPlacements(values, this.#columns);
      for (const [col] of placements) {
        assertAxisInBounds('row', row);
        assertAxisInBounds('column', col);
      }
      return {row, placements};
    });
    return planned.map(({row, placements}) =>
      placements.map(([col, value]) => {
        const cell = this.#cellAt(row, col);
        cell.value = value;
        return cell;
      }),
    );
  }

  /**
   * Freeze the top `ySplit` rows and left `xSplit` columns in place; the rest of the sheet scrolls
   * beneath them. `freeze(1)` pins a header row; `freeze(0, 1)` pins the first column. Passing both
   * zero clears the freeze (equivalent to {@link unfreeze}).
   *
   * @throws {RangeError} if either split is a negative or non-integer count, or leaves no row or
   *   column of the grid to scroll. The view is left as it was.
   */
  freeze(ySplit = 1, xSplit = 0): void {
    if (!Number.isInteger(ySplit) || ySplit < 0 || !Number.isInteger(xSplit) || xSplit < 0) {
      throw new RangeError(
        `freeze splits must be non-negative integers; got ySplit=${ySplit}, xSplit=${xSplit}`,
      );
    }
    if (ySplit === 0 && xSplit === 0) {
      this.unfreeze();
      return;
    }
    // Computed before anything is assigned, because it is what bounds both splits: a throw after
    // `state` and the splits were set left a frozen view pointing at the previous freeze's pane.
    const topLeftCell = encodeAddress(xSplit + 1, ySplit + 1);
    this.view.state = 'frozen';
    this.view.xSplit = xSplit;
    this.view.ySplit = ySplit;
    this.view.topLeftCell = topLeftCell;
  }

  /** Clear any frozen split, returning the sheet to a normal (fully scrolling) view. */
  unfreeze(): void {
    this.view.state = 'normal';
    delete this.view.xSplit;
    delete this.view.ySplit;
    delete this.view.topLeftCell;
  }

  /**
   * Copy the row at the 1-based `start`, `options.count` times (default 1). With `options.insert`
   * (the default) the copies are inserted directly after the source, shifting the rows below, and
   * any merged range there, down by `count`; otherwise the copies *replace* the rows immediately
   * below without shifting. A destination row is not overlaid but wholly re-made, so a cell it held
   * in a column the source leaves empty is dropped, exactly as it would be with a shifting insert.
   *
   * Each copy is a faithful duplicate of the source: its cell values, its per-cell styles, and its
   * row properties (height, hidden, outline level, row fill). It carries no merge of its own, so a
   * range can be merged onto a duplicated row afterwards. A formula is copied as Excel copies a row,
   * its relative references moved down with it and its cached result dropped, since that was computed
   * over the source's cells; an inserted copy is taken from the source as the insert left it.
   *
   * @throws {RangeError} if `start` is not a positive integer or `count` is negative, or if a copy
   *   would land past the last row. The sheet is left untouched.
   */
  duplicateRow(start: number, options: {count?: number; insert?: boolean} = {}): void {
    const {count = 1, insert = true} = options;
    assertStartAndCount('duplicate', 'row', start, count);
    assertSpliceFits('row', start + 1, count);
    const sourceProperties = this.#rowProperties.get(start);
    const snapshot = (destRow: number): Map<number, Cell> => {
      const row = new Map<number, Cell>();
      for (const [col, cell] of this.#rows.get(start) ?? []) {
        const copy = new Cell(destRow, col);
        copyCellContent(cell, copy);
        copy.value = filledDown(cell.value, destRow - start);
        row.set(col, copy);
      }
      return row;
    };
    // A row's height and outline level live beside the grid, so neither the splice nor a plain
    // `#rows.set` carries them; they are copied onto each destination explicitly. A source with no
    // properties of its own clears the destination's, keeping "replace" true of the whole row.
    const copyProperties = (destRow: number): void => {
      if (sourceProperties === undefined) this.#rowProperties.delete(destRow);
      // A shallow copy suffices: every field of `RowProperties` is a primitive or a deeply-readonly
      // `Fill`, so the copies share no mutable state with the source.
      else this.#rowProperties.set(destRow, {...sourceProperties});
    };
    // Blank rows go in first and the copies are taken after, so a source formula the insert moved (one
    // reading the rows below it) is copied as it now reads, which is what Excel does.
    if (insert) {
      const blanks = Array.from({length: count}, () => new Map<number, Cell>());
      this.#edits.spliceRows(start + 1, 0, blanks);
    }
    for (let i = 1; i <= count; i++) this.#rows.set(start + i, snapshot(start + i));
    for (let i = 1; i <= count; i++) copyProperties(start + i);
    this.#afterStructuralEdit();
  }

  /**
   * Remove `count` columns starting at the 1-based `start`, then insert the given columns in their
   * place: the column analog of {@link spliceRows}. Columns to the right shift by
   * `inserts.length - count`, keeping their values and styles, and a merged range lying wholly to
   * the right of the edit re-anchors to its new columns. Each inserted column is an array of values
   * indexed by row (index 0 → row 1); an empty array inserts a blank column.
   *
   * Formulas move with the columns, by the rules {@link spliceRows} gives for rows.
   *
   * @throws {RangeError} if `start` is not a positive integer or `count` is negative.
   * @throws {RangeError} if an inserted column would land past the last column, or one of its values
   *   past the last row. The sheet is left untouched, so this is a refused edit rather than half of
   *   one: a region pushed off the edge clamps and absorbs the loss, but content pushed off it is
   *   what Excel refuses outright, and {@link addColumn} refuses the same argument identically.
   */
  spliceColumns(start: number, count: number, ...inserts: CellValue[][]): void {
    assertStartAndCount('splice', 'column', start, count);
    assertSpliceFits('column', start, inserts.length);
    this.#edits.spliceColumns(start, count, inserts);
    this.#afterStructuralEdit();
  }

  /**
   * Insert one column of `values` at the 1-based `pos`, shifting the columns at and right of it over
   * by one. `values` is an array of values indexed by row (index 0 → row 1), like
   * {@link addColumn}. Shorthand for {@link spliceColumns}`(pos, 0, values)`.
   *
   * @throws {RangeError} if `pos` is not a positive integer, or if the column would land past the
   *   last column or one of its values past the last row; see {@link spliceColumns}.
   */
  insertColumn(pos: number, values: CellValue[]): void {
    this.spliceColumns(pos, 0, values);
  }

  /**
   * Append a column of `values` after the last used column, returning the cells it materialised.
   * The append point is {@link columnCount}` + 1`, so the column lands right of every column that
   * holds data or its own formatting, never overwriting existing content, unlike {@link insertColumn},
   * which shifts and needs a position. Unlike {@link spliceColumns}, appending shifts nothing, so it
   * never disturbs merges or the columns to its left.
   *
   * `values` is an array indexed by row (index 0 → row 1); a hole or an explicit `undefined` leaves
   * that row untouched, mirroring {@link addRow}'s positional-array shape. That is the only shape a
   * column takes: the other {@link RowInput} form addresses columns by their
   * {@link ColumnProperties.key}, and a column's values are indexed by *row*, which carries no key,
   * so there is nothing on this axis for a keyed object to name.
   */
  addColumn(values: CellValue[]): Cell[] {
    return this.addColumns([values])[0] ?? [];
  }

  /**
   * Append several columns after the last used column in one call, returning the cells materialised
   * for each. The columns stack in order: the first lands at {@link columnCount}` + 1`, the next
   * directly right of it, so a later column never collides with an earlier one even when both are
   * value-less. The bulk form of {@link addColumn}.
   */
  addColumns(columns: CellValue[][]): Cell[][] {
    const first = this.columnCount + 1;
    // Planned whole before the first write, as `addRows` is: a batch that runs off the grid is refused
    // whole rather than leaving the cells it reached before the throw.
    const planned = columns.map((values, i) => {
      const col = first + i;
      const placements = positionalPlacements(values);
      for (const [row] of placements) {
        assertAxisInBounds('column', col);
        assertAxisInBounds('row', row);
      }
      return {col, placements};
    });
    return planned.map(({col, placements}) =>
      placements.map(([row, value]) => {
        const cell = this.#cellAt(row, col);
        cell.value = value;
        return cell;
      }),
    );
  }

  /**
   * A snapshot of this sheet's value and overlay content (see {@link WorksheetModel}). Reading it and
   * assigning it onto another sheet (`dst.model = src.model`) reproduces the source: merges, cells
   * and their styles, column/row metadata, tables, the autofilter, protection, the frozen-pane view,
   * and the page setup all survive, because the getter emits and the setter consumes exactly the same
   * fields. Identity (`name`, `id`) is not part of the model and is never touched by assignment; nor
   * are attached parts that carry workbook-level identity (images, pivots, threaded comments,
   * byte-preserved charts/drawings). See {@link WorksheetModel} for that boundary.
   */
  get model(): WorksheetModel {
    const model: Record<string, unknown> = {};
    for (const facet of WORKSHEET_MODEL_FACETS) model[facet.key] = facet.read(this);
    // TypeScript cannot follow an object built key by key; what makes this sound is that the
    // registry is proved exhaustive over `keyof WorksheetModel`; see the type assertion beside it.
    return model as unknown as WorksheetModel;
  }

  // What every structural edit ends with. The extent was derived from lines the edit just moved and
  // cannot tell which bound went, so it re-derives on the next read. The merge index needs no such
  // call: a splice replaces its rectangles through `replaceAll`, which resets it.
  #afterStructuralEdit(): void {
    this.#extent.invalidate();
  }

  // Empty every collection the model round-trips, so a subsequent replay leaves no residue from
  // whatever the sheet held before. Images, pivots, and byte-preserved parts carry workbook-level
  // identity, are excluded from the model, and so are deliberately left untouched here.
  #resetContent(): void {
    this.#rows.clear();
    this.#columns.clear();
    this.#rowProperties.clear();
    this.#merges.clear();
    this.#dataValidations.clear();
    this.#conditionalFormattings.clear();
    this.#tables.length = 0;
    this.#extent.reset();
  }

  // Assigning a model replaces this sheet's content wholesale: the sheet becomes the model, with no
  // residue from whatever it held before. The registry's declaration order is the application order,
  // and it is load-bearing: see WORKSHEET_MODEL_FACETS.
  //
  // The whole model is applied to a detached sheet first, and only a model that applies there in full
  // reaches this one. Replaying straight onto the reset sheet meant a facet that threw part-way (two
  // merges overlapping, a validation naming no cells, a cell off the grid) left it emptied and half
  // loaded. The rehearsal needs nothing a workbook owns, since no facet reaches past the sheet it is
  // handed, and every check a facet makes is against state the reset clears, so the replay cannot
  // fail where the rehearsal passed. It costs a second application, which a whole-sheet copy affords.
  set model(model: WorksheetModel) {
    const rehearsal = new Worksheet(this.name, this.id);
    for (const facet of WORKSHEET_MODEL_FACETS) facet.write(rehearsal, model);
    this.#resetContent();
    for (const facet of WORKSHEET_MODEL_FACETS) facet.write(this, model);
  }

  /**
   * Protect the sheet, making the per-cell `locked`/`hidden` flags enforceable. Without a
   * password the protection is a soft lock any consumer can lift; with one, the password is
   * salted and hashed on the spot (the plaintext is never retained) so lifting the protection
   * requires re-supplying it. `options` names which operations stay available to a user while
   * the sheet is protected; anything unspecified falls to Excel's default for that operation.
   *
   * Re-protecting replaces any prior protection; {@link unprotect} clears it.
   */
  protect(password?: string, options: SheetProtectionOptions = {}): void {
    const {spinCount, ...flags} = options;
    const protection: {flags: SheetProtection['flags']; credential?: SheetProtectionCredential} = {
      flags,
    };
    if (password !== undefined && password !== '') {
      protection.credential = deriveCredential(password, spinCount);
    }
    this.#protection = protection;
  }

  /** Remove any protection previously set by {@link protect}. */
  unprotect(): void {
    this.#protection = undefined;
  }

  /** The sheet's protection, or `undefined` if the sheet is unprotected. */
  get protection(): SheetProtection | undefined {
    return this.#protection;
  }

  #cellAt(row: number, col: number): Cell {
    const cols = this.#rows.get(row);
    let cell = cols?.get(col);
    if (cell === undefined) {
      // Built before anything is stored: the constructor is what bounds both axes, and storing the
      // row map first left a row off the grid behind a refused position, after which iterating or
      // writing the sheet threw on every call.
      cell = new Cell(row, col);
      if (cols === undefined) this.#rows.set(row, new Map([[col, cell]]));
      else cols.set(col, cell);
    }
    this.#extent.noteCell(row, col);
    return cell;
  }

  /**
   * The codec's channel into this sheet. See `core/internal.ts` for why these are not public
   * methods. Declared last so every private field it closes over is already in scope.
   */
  readonly [INTERNAL]: WorksheetInternals = {
    evictRow: (number) => {
      this.#rows.delete(number);
      this.#rowProperties.delete(number);
      this.#extent.invalidate();
    },
    addLoadedPivotTable: (pivot) => {
      this.#loadedPivotTables.push(pivot);
    },
    restoreCommentThreads: (threads) => {
      this.#comments.restore(threads);
    },
    addPreservedReference: (reference) => {
      this.#preservedReferences.push(reference);
    },
    setFormulaHost: (host) => {
      this.#formulaHost = host;
    },
    spliceFormulas: (edit) => {
      this.#edits.spliceFormulas(edit);
    },
    restoreProtection: (protection) => {
      this.#protection = protection;
    },
    cellAt: (row, col) => this.#cellAt(row, col),
    masterAt: (row, col) => {
      const master = this.#merges.masterOf(row, col);
      return this.#cellAt(master.row, master.col);
    },
    peekCell: (row, col) => this.#rows.get(row)?.get(col),
    rowPropertiesOf: (number) => this.#rowProperties.get(number),
    ensureRowProperties: (number) => {
      let properties = this.#rowProperties.get(number);
      if (properties === undefined) {
        properties = {};
        this.#rowProperties.set(number, properties);
      }
      this.#extent.noteDeclaredRow(number);
      return properties;
    },
    dropRowProperties: (number) => {
      // The extent only ever grows as positions are declared, so removing one has to make it
      // re-derive rather than be told: nothing else can know whether this row was the bound.
      if (this.#rowProperties.delete(number)) this.#extent.invalidate();
    },
    rowCells: (number) => {
      const cols = this.#rows.get(number);
      return cols ? [...cols].sort(([a], [b]) => a - b).map(([, cell]) => cell) : [];
    },
    columnPropertiesOf: (index) => this.#columns.get(index),
    ensureColumnProperties: (index) => {
      let properties = this.#columns.get(index);
      if (properties === undefined) {
        properties = {};
        this.#columns.set(index, properties);
      }
      this.#extent.noteDeclaredColumn(index);
      return properties;
    },
    dropColumnProperties: (index) => {
      if (this.#columns.delete(index)) this.#extent.invalidate();
    },
    columnCells: (index) => {
      const cells: Cell[] = [];
      for (const number of [...this.#rows.keys()].sort((a, b) => a - b)) {
        const cell = this.#rows.get(number)?.get(index);
        if (cell !== undefined) cells.push(cell);
      }
      return cells;
    },
  };
}

/**
 * Refuse a picture on a sheet whose drawing was kept whole from a file.
 *
 * A drawing holding a chart, shape or other object the model does not interpret is carried byte for
 * byte, and none of its anchors are modelled. A worksheet references one drawing, so a modelled picture
 * would need a second drawing the sheet cannot point at, and writing it would drop the kept one's
 * content without a word. Called wherever a picture can reach a sheet, and by the writer, which is the
 * guarantee for any path that skips the others.
 *
 * @throws {AuthoringError} if the sheet keeps such a drawing.
 */
export function refuseImagesBesideKeptDrawing(sheet: Worksheet): void {
  if (sheet.preservedReferences.some((reference) => reference.element === 'drawing')) {
    throw new AuthoringError(
      `sheet ${quoted(sheet.name)} keeps a drawing whose content (a chart, shape or other object) this ` +
        'library preserves but does not model, so a picture cannot be added to it: a sheet has one ' +
        'drawing, and writing the picture would drop what the kept one holds',
    );
  }
}

// The bounds contract every splice-shaped edit shares: a 1-based start and a non-negative count.
// The three call sites used to spell it out, so the messages differed only by interpolation, which
// is how messages drift. `verb` names the operation the caller offered, `axis` the line it edits.
// A cell value copied `offset` rows below its source: a formula's relative references move with it, and
// its cached result, computed over the source's cells, goes. A shared-formula clone stays a clone of the
// same master, which describes it at the new position as well; only the text a read resolved onto it
// moves.
function filledDown(value: CellValue, offset: number): CellValue {
  if (isFormulaValue(value)) return {formula: translateFormula(value.formula, 0, offset)};
  if (isSharedFormulaValue(value)) {
    return value.formula === undefined
      ? {sharedFormula: value.sharedFormula}
      : {sharedFormula: value.sharedFormula, formula: translateFormula(value.formula, 0, offset)};
  }
  return value;
}

function assertStartAndCount(
  verb: string,
  axis: 'row' | 'column',
  start: number,
  count: number,
): void {
  // The start is a position on the axis, so it is refused in the same words `getRow` and `getRange`
  // refuse one: `verb` used to prefix it, which named the call at the cost of making one mistake
  // answer three ways depending on which door the caller came through.
  assertAxisInBounds(axis, start);
  if (!Number.isInteger(count) || count < 0) {
    throw new RangeError(`${verb} count ${count} is invalid: it must be a non-negative integer`);
  }
}

// Where an edit's inserted lines land, `start` through `start + insertCount - 1`, checked before
// anything moves. An inserted line carrying content is refused by the `Cell` it builds, but one
// carrying none builds no cell: `spliceRows(1048576, 0, [], [])` left a row key past the grid, and
// duplicating a height-only last row did the same, so the sheet threw on every later read or write.
function assertSpliceFits(axis: 'row' | 'column', start: number, insertCount: number): void {
  if (insertCount > 0) assertAxisInBounds(axis, start + insertCount - 1);
}

/**
 * What the library's own machinery may do to a `Worksheet` that an author may not: reinstate state
 * read from a file, place a cell where a file or model says rather than where addressing would put
 * it, and release a row the streaming writer has finished with. Reached as `sheet[INTERNAL]`; see
 * `core/internal.ts`.
 */
export interface WorksheetInternals {
  /**
   * Drop a row's materialised cells and format properties, releasing its cell graph. The streaming
   * writer calls this the moment a row is serialised so peak memory stays bounded to the rows still
   * in flight rather than the whole sheet. Row *numbering* is the caller's concern: eviction lowers
   * {@link Worksheet.rowCount}, so an append-driven producer must track its own high-water mark
   * rather than lean on this sheet's used range.
   */
  evictRow(number: number): void;

  /**
   * Register a pivot table reconstructed from a loaded package: the reader's counterpart to
   * {@link Worksheet.addPivotTable}. This records an inspectable, read-only view of a pivot the
   * reader parsed from its OOXML parts; the pivot itself round-trips by byte-preservation, so
   * registering it here only makes it visible via {@link Worksheet.loadedPivotTables} and never
   * affects what the writer emits.
   */
  addLoadedPivotTable(pivot: ParsedPivotTable): void;

  /**
   * Reinstate the threaded conversations read from a file, in the order the reader found them,
   * replacing any already held. Their authors and mentioned people are already resolved against the
   * workbook registry, so a thread arrives self-contained.
   *
   * Not authoring: {@link Worksheet.addCommentThread} is that, and validates the anchor. These
   * threads are what a re-write emits, so what the reader hands over is what the file will say.
   */
  restoreCommentThreads(threads: readonly CommentThread[]): void;

  /**
   * Record a worksheet-level reference to package content the model does not interpret, so the writer
   * re-emits it verbatim. Called by the reader when it meets a `<drawing>` holding only vector shapes
   * or a `<legacyDrawingHF>` header/footer image.
   */
  addPreservedReference(reference: PreservedWorksheetReference): void;

  /**
   * Reinstate an already-derived protection state: the deserialization counterpart to
   * {@link Worksheet.protect}. A loaded `<sheetProtection>` carries its credential in finished agile
   * form (algorithm, hash, salt, spin count) with no recoverable plaintext password, so the reader
   * restores that credential verbatim rather than re-hashing.
   */
  restoreProtection(protection: SheetProtection): void;

  /**
   * Hand the sheet what moves the formulas beyond it, every other sheet's and the workbook's defined
   * names, when a splice of this sheet moves the lines they refer to. The workbook sets it on each sheet
   * it creates.
   */
  setFormulaHost(host: (edit: SheetSplice) => void): void;

  /** Move the references this sheet's formulas make to a spliced sheet, which here is another one. */
  spliceFormulas(edit: SheetSplice): void;

  /**
   * Materialise the cell at an exact 1-based position, creating it on first access. Unlike
   * {@link Worksheet.getCell} this performs no merge resolution: the cell returned is the one at
   * `(row, col)` even when a merged region covers it. Loading content is where that matters: a
   * model or a parsed file states where each value sits, and routing a covered value to its region
   * master mid-load would move it.
   */
  cellAt(row: number, col: number): Cell;
  /**
   * The cell materialised at a position, or `undefined` when nothing has been written there.
   *
   * Merge-blind and non-materialising, which is what separates it from `Worksheet.getCell`: that
   * resolves a covered address to its merge master, so `hasCell(row, col)` followed by a `getCell` of
   * the same position can hand back a *different* cell. A caller walking a rectangle got the master
   * once per covered position and never saw the covered cells at all, which made a `clearStyle` over a
   * merge clear the master repeatedly and the covered cells not once. A caller that wants merge
   * resolution keeps `getCell`; a caller enumerating what is actually stored wants this.
   */
  peekCell(row: number, col: number): Cell | undefined;

  /**
   * {@link cellAt} with merge resolution: the cell a write to this position lands on, materialised.
   *
   * What {@link Worksheet.getCell} does, minus the address. `getCell` is the public spelling and takes
   * an A1 reference, so a caller already holding a row and a column had to encode one and have it
   * decoded straight back. That round-trip costs about three times a positional read, and a
   * {@link Range} write performs one per cell: styling a 100x100 block spent 10,000 encodes and 10,000
   * regex decodes on positions it already had.
   */
  masterAt(row: number, col: number): Cell;

  /**
   * The store behind a {@link Row} or {@link Column} handle. These exist because the handles are
   * views rather than records: they hold a sheet and a position, and every read and write goes
   * through here to the one authoritative map. `…PropertiesOf` never fabricates, so reading a row
   * cannot extend the used range; `ensure…` is what a write calls, so the record appears exactly
   * when a value is set; `drop…` is what clearing the last field calls, so the record disappears
   * exactly when the line stops being formatted, which is what keeps the used range honest.
   */
  rowPropertiesOf(number: number): RowProperties | undefined;
  ensureRowProperties(number: number): RowProperties;
  dropRowProperties(number: number): void;
  /** The row's materialised cells in ascending column order. */
  rowCells(number: number): Cell[];
  columnPropertiesOf(index: number): ColumnProperties | undefined;
  ensureColumnProperties(index: number): ColumnProperties;
  dropColumnProperties(index: number): void;
  /** The column's materialised cells in ascending row order. */
  columnCells(index: number): Cell[];
}
