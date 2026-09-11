// Structural-edit machinery: the splice arithmetic that inserts or deletes whole rows and columns
// and keeps everything anchored to the grid moving in step: line metadata and page breaks, merged
// ranges, tables, anchored images, the coordinates a cell's value carries, and the range-bound
// overlays (data validations, conditional formats, comment threads, the autofilter). It is isolated
// from Worksheet because it is pure grid mechanics: it holds the sheet's storage containers by
// reference and mutates them in place, and touches none of the public cell API. Worksheet builds the cells an insert introduces, then hands
// the pre-built rows (or the raw column values) here for the shift.

import {
  boundedRect,
  decodeRange,
  encodeAddress,
  encodeRect,
  type GridRect,
  tryDecodeCellRef,
  tryDecodeRange,
} from './address.ts';
import {type AutoFilter, shiftAutoFilter} from './autofilter.ts';
import {Cell, copyCellContent} from './cell.ts';
import type {ConditionalFormattingOverlay} from './conditional-formatting-overlay.ts';
import {replaceContents} from './containers.ts';
import type {DataValidationOverlay} from './data-validation-overlay.ts';
import {type SheetSplice, spliceFormula, translateFormula} from './formula-references.ts';
import {type AxisSplice, isDeletedSpan, shiftIndex, shiftPoint, shiftRect} from './grid-shift.ts';
import {type AnchoredImage, type AnchorPoint, type ImageAnchor, isOneCellAnchor} from './image.ts';
import {INTERNAL} from './internal.ts';
import type {MergeRect} from './merge.ts';
import type {PageBreak} from './page-setup.ts';
import {type ParsedPivotTable, type PivotTable, splicePivotSource} from './pivot-table.ts';
import {positionalPlacements} from './row-input.ts';
import type {Table} from './table.ts';
import {
  type CellValue,
  isDataTableFormulaValue,
  isFormulaValue,
  isHyperlinkValue,
  isSharedFormulaValue,
} from './value.ts';
import type {WorksheetComments} from './worksheet-comments.ts';
import type {WorksheetMerges} from './worksheet-merges.ts';
import type {ColumnProperties, RowProperties} from './worksheet.ts';

/**
 * The sheet's autofilter, reached as a slot rather than held by reference like the containers beside
 * it: it is a single replaceable value, and a splice that deletes every filtered line clears it.
 */
export interface AutoFilterSlot {
  get(): AutoFilter | undefined;
  set(next: AutoFilter | undefined): void;
}

// The sheet's mutable storage, shared by reference with Worksheet. Never reassigned, only mutated in
// place, so the two views stay in sync through every splice.
interface GridStorage {
  readonly rows: Map<number, Map<number, Cell>>;
  readonly rowProperties: Map<number, RowProperties>;
  readonly columns: Map<number, ColumnProperties>;
  readonly merges: WorksheetMerges;
  readonly tables: Table[];
  readonly pivotTables: readonly PivotTable[];
  readonly loadedPivotTables: ParsedPivotTable[];
  readonly images: AnchoredImage[];
  readonly dataValidations: DataValidationOverlay;
  readonly conditionalFormattings: ConditionalFormattingOverlay;
  /** The sheet's name, which an unqualified reference in one of its formulas means. */
  readonly sheetName: () => string;
  /**
   * What moves the formulas outside this sheet that name it: every other sheet's, and the workbook's
   * defined names. `undefined` for a sheet no workbook holds.
   */
  readonly formulaHost: () => ((edit: SheetSplice) => void) | undefined;
  readonly comments: WorksheetComments;
  readonly rowBreaks: PageBreak[];
  readonly columnBreaks: PageBreak[];
  readonly autoFilter: AutoFilterSlot;
}

export class GridEdits {
  readonly #rows: Map<number, Map<number, Cell>>;
  readonly #rowProperties: Map<number, RowProperties>;
  readonly #columns: Map<number, ColumnProperties>;
  readonly #merges: WorksheetMerges;
  readonly #tables: Table[];
  readonly #pivotTables: readonly PivotTable[];
  readonly #loadedPivotTables: ParsedPivotTable[];
  readonly #images: AnchoredImage[];
  readonly #dataValidations: DataValidationOverlay;
  readonly #conditionalFormattings: ConditionalFormattingOverlay;
  readonly #comments: WorksheetComments;
  readonly #rowBreaks: PageBreak[];
  readonly #columnBreaks: PageBreak[];
  readonly #autoFilter: AutoFilterSlot;
  readonly #sheetName: () => string;
  readonly #formulaHost: () => ((edit: SheetSplice) => void) | undefined;

  constructor(storage: GridStorage) {
    this.#rows = storage.rows;
    this.#rowProperties = storage.rowProperties;
    this.#columns = storage.columns;
    this.#merges = storage.merges;
    this.#tables = storage.tables;
    this.#pivotTables = storage.pivotTables;
    this.#loadedPivotTables = storage.loadedPivotTables;
    this.#images = storage.images;
    this.#dataValidations = storage.dataValidations;
    this.#conditionalFormattings = storage.conditionalFormattings;
    this.#comments = storage.comments;
    this.#rowBreaks = storage.rowBreaks;
    this.#columnBreaks = storage.columnBreaks;
    this.#autoFilter = storage.autoFilter;
    this.#sheetName = storage.sheetName;
    this.#formulaHost = storage.formulaHost;
  }

  // Apply a delete-then-insert to the row grid: surviving rows below the edit shift by
  // `inserted.length - count`, deleted rows drop out, and the pre-built inserted rows land at `start`.
  // Row metadata, merged ranges and everything else anchored to the grid shift the same way, so a
  // formatting-only row, a covered merge or a dropdown stays aligned with the data it describes.
  spliceRows(start: number, count: number, inserted: Map<number, Cell>[]): void {
    const splice: AxisSplice = {axis: 'row', start, count, delta: inserted.length - count};
    this.#moveFormulaReferences(splice);
    const shifted = new Map<number, Map<number, Cell>>();
    for (const [row, cols] of this.#rows) {
      if (row < start) shifted.set(row, cols);
      else if (row >= start + count) {
        // Through `shiftIndex` like every other participant. Raw arithmetic here let an insert on a
        // sheet holding a cell in the last row hand `new Cell` a coordinate it asserts against, so the
        // splice died with a `RangeError` from inside the grid where a merge or a table would have
        // clamped. A row clamped onto the last one lands on whatever is already there, which is the
        // content Excel also loses when it pushes a row off the bottom.
        const dest = shiftIndex(row, splice);
        shifted.set(dest, this.#relocateRow(cols, dest));
      }
    }
    inserted.forEach((cols, i) => {
      shifted.set(start + i, this.#relocateRow(cols, start + i));
    });
    this.#rows.clear();
    for (const [row, cols] of shifted) this.#rows.set(row, cols);

    this.#shiftAnchored(splice, this.#rowProperties);
  }

  // Apply a delete-then-insert to the column grid: cells left of the edit stay, cells at or beyond the
  // deleted span shift by `inserts.length - count` carrying their content, and the inserted column
  // values materialise as fresh cells at `start`. Column metadata, merges, tables, images, the
  // coordinates cell values carry and the range-bound overlays re-anchor the same way.
  spliceColumns(start: number, count: number, inserts: CellValue[][]): void {
    const splice: AxisSplice = {axis: 'col', start, count, delta: inserts.length - count};
    this.#moveFormulaReferences(splice);
    // Built whole, then swapped in, the way `spliceRows` does it. Writing each row back inside the loop
    // meant a throw part-way left the sheet half-spliced: rows already visited shifted, the rest not,
    // with no way for the caller to act on the error. `Worksheet.spliceColumns` refuses an insert that
    // would land past the last column before calling here, and `new Cell` still refuses a value naming
    // a row past the last one while the inserts are built; swapping at the end is what keeps that a
    // refused edit rather than half of one.
    //
    // Refusing is the answer, not an omission, and it is deliberately not the clamp every *region* a
    // splice moves gets. Excel draws the same line: a region whose edge is pushed off the grid is
    // clamped and absorbs the loss, while content pushed off it makes Excel refuse the whole insert
    // ("can't insert new cells because it would push non-empty cells off the end of the worksheet").
    // Clamping here would be worse than either -- two inserted columns would stack on XFD, and the
    // caller would be told the edit succeeded. See
    // docs/knowledge/specs/a-splice-must-not-push-geometry-off-the-grid.md.
    const shiftedRows = new Map<number, Map<number, Cell>>();
    for (const [row, cols] of this.#rows) {
      const shifted = new Map<number, Cell>();
      for (const [col, cell] of cols) {
        if (col < start) {
          shifted.set(col, cell);
        } else if (col >= start + count) {
          // `shiftIndex` for the same reason the row axis uses it: `new Cell` asserts its coordinates,
          // so an unclamped destination past column XFD threw from inside the splice.
          const dest = shiftIndex(col, splice);
          const moved = new Cell(row, dest);
          copyCellContent(cell, moved);
          shifted.set(dest, moved);
        }
      }
      shiftedRows.set(row, shifted);
    }
    // A pass of its own, because an inserted column's values are indexed by row and the rows they
    // name need not exist yet. Nested inside the loop above, a value could only land on a row the
    // grid already held: `insertColumn(1, ['x','y','z'])` on an empty sheet wrote nothing at all, and
    // on a sheet holding only `A1` it wrote the first value and discarded the rest. `addColumn` goes
    // through `positionalPlacements` and materialises every row, so the two column-append paths
    // disagreed about the same argument. This is the column-axis mirror of the pre-built `inserted`
    // maps `spliceRows` receives.
    inserts.forEach((values, i) => {
      for (const [row, value] of positionalPlacements(values)) {
        const cell = new Cell(row, start + i);
        cell.value = value;
        let cols = shiftedRows.get(row);
        if (cols === undefined) {
          cols = new Map<number, Cell>();
          shiftedRows.set(row, cols);
        }
        cols.set(start + i, cell);
      }
    });
    for (const [row, cols] of shiftedRows) this.#rows.set(row, cols);
    this.#shiftAnchored(splice, this.#columns);
  }

  // The references formula text makes to this sheet, moved through a splice of it, in this sheet and
  // wherever else the workbook holds a formula. Run before any cell moves, for two reasons: a formula an
  // insert brings in was written against the grid after the splice and must not move again, and a
  // shared formula's clone is recovered from its master at the offset between the two as they stand now.
  #moveFormulaReferences(splice: AxisSplice): void {
    const edit: SheetSplice = {sheet: this.#sheetName(), splice};
    this.spliceFormulas(edit);
    this.#formulaHost()?.(edit);
  }

  /**
   * Move the references this sheet's formulas make to the sheet a splice changed, which may be this
   * sheet or another: every cell formula, and every data validation's and conditional format's formula.
   *
   * A shared formula is described once, by its master, and each clone is that text translated to where
   * the clone sits. A splice usually keeps that true, since the master, the clone and what they refer to
   * all move together, but not always: an insert between the cells one clone refers to and the cells its
   * master does grows one reference and not the other, and a delete can take the master and leave the
   * clone. A clone the moved master no longer describes stops sharing and becomes a formula of its own,
   * spelled as it was rewritten; the rest of the group keeps sharing.
   */
  spliceFormulas(edit: SheetSplice): void {
    const home = this.#sheetName();
    const {splice} = edit;
    const rewrite = (formula: string): string => spliceFormula(formula, home, edit);
    // A splice of another sheet moves no cell here, only what the formulas say.
    const cellsMove = home.toLowerCase() === edit.sheet.toLowerCase();
    const after = (cell: Cell): {col: number; row: number} => {
      if (!cellsMove) return cell;
      return splice.axis === 'row'
        ? {col: cell.col, row: shiftIndex(cell.row, splice)}
        : {col: shiftIndex(cell.col, splice), row: cell.row};
    };
    const taken = (cell: Cell): boolean => {
      const line = splice.axis === 'row' ? cell.row : cell.col;
      return cellsMove && isDeletedSpan(line, line, splice);
    };

    // Every master's text before and after, keyed by its address: a clone assigned without its own text
    // is recovered from the text it was translated from, not the rewritten one.
    const masters = new Map<string, {cell: Cell; formula: string; rewritten: string}>();
    for (const cols of this.#rows.values()) {
      for (const cell of cols.values()) {
        const value = cell.value;
        if (!isFormulaValue(value)) continue;
        masters.set(encodeAddress(cell.col, cell.row), {
          cell,
          formula: value.formula,
          rewritten: rewrite(value.formula),
        });
      }
    }
    for (const cols of this.#rows.values()) {
      for (const cell of cols.values()) {
        const value = cell.value;
        if (!isSharedFormulaValue(value)) continue;
        const master = masters.get(value.sharedFormula);
        if (master === undefined) {
          // An orphan the writer will report: its own text, if any, still moves.
          if (value.formula === undefined) continue;
          const rewritten = rewrite(value.formula);
          if (rewritten !== value.formula) cell.value = {...value, formula: rewritten};
          continue;
        }
        const source =
          value.formula ??
          translateFormula(master.formula, cell.col - master.cell.col, cell.row - master.cell.row);
        const rewritten = rewrite(source);
        const from = after(master.cell);
        const to = after(cell);
        const described =
          !taken(master.cell) &&
          translateFormula(master.rewritten, to.col - from.col, to.row - from.row) === rewritten;
        if (described) {
          if (value.formula !== undefined && rewritten !== value.formula) {
            cell.value = {...value, formula: rewritten};
          }
          continue;
        }
        cell.value =
          value.result === undefined
            ? {formula: rewritten}
            : {formula: rewritten, result: value.result};
      }
    }
    for (const {cell, formula, rewritten} of masters.values()) {
      const value = cell.value;
      if (rewritten !== formula && isFormulaValue(value))
        cell.value = {...value, formula: rewritten};
    }
    this.#dataValidations.mapFormulas(rewrite);
    this.#conditionalFormattings.mapFormulas(rewrite);
    for (const table of this.#tables) table[INTERNAL].rewriteFormulas(rewrite);
    for (const pivot of this.#pivotTables) pivot[INTERNAL].spliceSource(edit);
    // A loaded pivot is written from its preserved cache, which the writer moves by the same rule; its
    // view moves here so that it says what will be written.
    for (const [index, loaded] of this.#loadedPivotTables.entries()) {
      const {source} = loaded;
      if (source.kind !== 'worksheet' || source.inAnotherWorkbook) continue;
      const ref = splicePivotSource(source.sheet, source.ref, edit);
      if (ref !== source.ref)
        this.#loadedPivotTables[index] = {...loaded, source: {...source, ref}};
    }
  }

  // Everything anchored to the grid besides the cells, moved through one splice. Both axes end here, so
  // a participant is added once and cannot be moved on one axis and forgotten on the other.
  #shiftAnchored<T>(splice: AxisSplice, lineProperties: Map<number, T>): void {
    this.#shiftLineProperties(lineProperties, splice);
    this.#shiftPageBreaks(splice);
    this.#shiftMerges(splice);
    this.#shiftTables(splice);
    this.#shiftImages(splice);
    this.#reanchorValueReferences(splice);
    this.#shiftRangeBoundOverlays(splice);
  }

  // Re-anchor the four things bound to a range that live outside the cell grid: data validations,
  // conditional formats, comment threads, and the sheet's autofilter. Each owns its own arithmetic:
  // an overlay knows whether it holds a rectangle or a point, and the autofilter knows that its
  // criteria are addressed relative to its own left edge. This pass only routes the splice to them
  // and lets a deleted anchor take its entry with it.
  #shiftRangeBoundOverlays(splice: AxisSplice): void {
    this.#dataValidations.shift(splice);
    this.#conditionalFormattings.shift(splice);
    this.#comments.shift(splice);
    const filter = this.#autoFilter.get();
    if (filter !== undefined) this.#autoFilter.set(shiftAutoFilter(filter, splice));
  }

  // Rebuild a row's cells at a new row index. `Cell` fixes its position at construction, so a moved
  // row is a fresh set of cells at `destRow` carrying the originals' content.
  #relocateRow(cols: Map<number, Cell>, destRow: number): Map<number, Cell> {
    const moved = new Map<number, Cell>();
    for (const [col, cell] of cols) {
      if (cell.row === destRow) {
        moved.set(col, cell);
      } else {
        const copy = new Cell(destRow, col);
        copyCellContent(cell, copy);
        moved.set(col, copy);
      }
    }
    return moved;
  }

  // Move the grid coordinates a cell's value carries: the positions stored inside a value rather than
  // beside it, which are a shared-formula clone's master address, a hyperlink's clickable `range`, and
  // a data table's filled `ref` and input cells. Each names cells by position, so left behind it names
  // cells the splice moved away from, and the writer emits it as written. Formula text is moved before
  // any of this, by `spliceFormulas`.
  #reanchorValueReferences(splice: AxisSplice): void {
    for (const cols of this.#rows.values()) {
      for (const cell of cols.values()) {
        const value = cell.value;
        const moved = reanchoredValue(value, splice);
        if (moved !== value) cell.value = moved;
      }
    }
  }

  // Shift a line-metadata map (row properties keyed by row, or column properties keyed by column)
  // through a splice: entries before the edit stay, entries a delete swallowed whole drop, entries
  // after shift by `delta`. Mutates the map in place.
  //
  // Through `shiftIndex` and `isDeletedSpan` like every other participant. This was the last site
  // still doing the arithmetic by hand, and a hand-written shift does not clamp: a row height on the
  // last row plus an insert above it left a properties entry at 1048577, which made `rowCount` name
  // a row `new Row` refuses to construct, so iterating the sheet threw and the sheet could no longer
  // be written or inspected. The cells on that row clamped correctly, so a row and its own metadata
  // also came apart.
  #shiftLineProperties<T>(map: Map<number, T>, splice: AxisSplice): void {
    const shifted = new Map<number, T>();
    for (const [index, value] of map) {
      if (isDeletedSpan(index, index, splice)) continue;
      shifted.set(shiftIndex(index, splice), value);
    }
    map.clear();
    for (const [index, value] of shifted) map.set(index, value);
  }

  // Re-anchor the page breaks on the spliced axis. A break falls between line `id` and the line after
  // it and belongs to that later line, the one Excel reports as its location: the break moves with it,
  // and a delete of that line takes the break too, which is what Excel does to both. Breaks on the
  // other axis keep their `id`; `PageBreak` says why their extent is not moved either.
  #shiftPageBreaks(splice: AxisSplice): void {
    const breaks = splice.axis === 'row' ? this.#rowBreaks : this.#columnBreaks;
    const moved = breaks.flatMap((brk) => {
      const line = brk.id + 1;
      if (isDeletedSpan(line, line, splice)) return [];
      const id = shiftIndex(line, splice) - 1;
      return [id === brk.id ? brk : {...brk, id}];
    });
    replaceContents(breaks, moved);
  }

  // Re-anchor merged ranges through a row or column splice. A range wholly before the edit is
  // untouched; one wholly after shifts by `nInserts - count`; one whose covered rows/columns are
  // entirely deleted is dropped. A range straddling the cut is a genuinely ambiguous geometry: its
  // edges are clamped to the cut line as a best effort. Unbounded whole-row/column merges carry no
  // rectangle and pass through unchanged.
  #shiftMerges(splice: AxisSplice): void {
    const merges: string[] = [];
    const rects: MergeRect[] = [];
    for (const range of this.#merges.ranges) {
      const decoded = boundedRect(decodeRange(range));
      if (decoded === undefined) {
        merges.push(range);
        continue;
      }
      const rect = shiftRect(decoded, splice);
      if (rect === undefined) continue;
      rects.push(rect);
      merges.push(encodeRect(rect));
    }
    this.#merges.replaceAll(merges, rects);
  }

  // Re-pin the sheet's tables through a splice on the given axis, dropping any table a delete leaves
  // with no row to occupy. `Table` owns the shift arithmetic; the sheet only prunes the casualties.
  #shiftTables(splice: AxisSplice): void {
    const survivors = this.#tables.filter((table) =>
      splice.axis === 'row' ? table.shiftRows(splice) : table.shiftColumns(splice),
    );
    replaceContents(this.#tables, survivors);
  }

  // Re-pin anchored images through a splice. An anchor point moves like a merge edge: a point before
  // the cut stays, one at or after it shifts by `delta`, and one inside a deleted span clamps to the
  // cut line. Grid points are 0-based, so each is converted to the 1-based coordinate the shared
  // shift arithmetic uses and back. An anchor whose points both move keeps its size; an anchor
  // straddling the cut grows or shrinks, matching how Excel reflows a picture across inserted rows.
  #shiftImages(splice: AxisSplice): void {
    const shiftAnchor = (point: AnchorPoint): AnchorPoint => {
      const zeroBased = splice.axis === 'row' ? point.row : point.col;
      const shifted = shiftIndex(zeroBased + 1, splice) - 1;
      if (shifted === zeroBased) return point;
      return splice.axis === 'row' ? {...point, row: shifted} : {...point, col: shifted};
    };
    const moved: AnchoredImage[] = this.#images.map((image) => {
      const from = shiftAnchor(image.anchor.from);
      const anchor: ImageAnchor = isOneCellAnchor(image.anchor)
        ? {...image.anchor, from}
        : {...image.anchor, from, to: shiftAnchor(image.anchor.to)};
      return {...image, anchor};
    });
    replaceContents(this.#images, moved);
  }
}

// A cell value with the coordinates it carries moved through a splice, or the very same value when
// none of them moved, so an untouched cell is not reassigned.
function reanchoredValue(value: CellValue, splice: AxisSplice): CellValue {
  if (isSharedFormulaValue(value)) {
    // A clone stores its master's absolute address, and a stale one is a clone the writer rejects as
    // orphaned. A master in the deleted span clamps to the cut line like a merge edge, leaving a
    // genuinely orphaned clone the writer then reports legibly.
    const master = tryDecodeCellRef(value.sharedFormula);
    if (master === undefined) return value;
    const anchored =
      splice.axis === 'row'
        ? encodeAddress(master.col, shiftIndex(master.row, splice))
        : encodeAddress(shiftIndex(master.col, splice), master.row);
    return anchored === value.sharedFormula ? value : {...value, sharedFormula: anchored};
  }
  if (isHyperlinkValue(value) && value.range !== undefined) {
    const range = shiftedRange(value.range, splice);
    if (range === value.range) return value;
    // A range the delete took whole, around a cell that survived, describes nothing any more: the link
    // stays on its cell as an ordinary single-cell one.
    if (range === undefined) {
      const {hyperlink, text, tooltip} = value;
      return tooltip === undefined ? {hyperlink, text} : {hyperlink, text, tooltip};
    }
    return {...value, range};
  }
  if (isDataTableFormulaValue(value)) {
    // The filled range holds the table's own cell, so a cell that survived keeps a range that did.
    const ref = shiftedRange(value.ref, splice) ?? value.ref;
    const r1 = inputAfter(value.r1, value.r1Deleted, splice);
    const r2 = inputAfter(value.r2, value.r2Deleted, splice);
    if (ref === value.ref && r1.same && r2.same) return value;
    return {
      ...value,
      ref,
      ...(r1.ref === undefined ? {} : {r1: r1.ref}),
      ...(r2.ref === undefined ? {} : {r2: r2.ref}),
      ...(r1.deleted ? {r1Deleted: true} : {}),
      ...(r2.deleted ? {r2Deleted: true} : {}),
    };
  }
  return value;
}

// A range reference moved as a region: the text itself when it did not move or names no bounded range,
// re-spelled when it moved, and `undefined` when the delete took it whole.
function shiftedRange(ref: string, splice: AxisSplice): string | undefined {
  const decoded = tryDecodeRange(ref);
  const rect = decoded === undefined ? undefined : boundedRect(decoded);
  if (rect === undefined) return ref;
  const moved = shiftRect(rect, splice);
  if (moved === undefined) return undefined;
  return sameRect(moved, rect) ? ref : encodeRect(moved);
}

// A data table's input cell after a splice. One that survived moves with its line. One the delete took
// keeps the reference as it was written and is flagged deleted, which is what Excel 16.0 does whichever
// way the cell goes (its row, its column, or itself with the cells below shifting up), and the table
// then shows `#REF!`. A reference already flagged names no cell any more, so nothing moves it.
function inputAfter(
  ref: string | undefined,
  deleted: boolean | undefined,
  splice: AxisSplice,
): {readonly ref: string | undefined; readonly deleted: boolean; readonly same: boolean} {
  const cell = ref === undefined || deleted === true ? undefined : tryDecodeCellRef(ref);
  if (cell === undefined) return {ref, deleted: deleted === true, same: true};
  const moved = shiftPoint(cell, splice);
  if (moved === undefined) return {ref, deleted: true, same: false};
  const same = moved.row === cell.row && moved.col === cell.col;
  return {ref: same ? ref : encodeAddress(moved.col, moved.row), deleted: false, same};
}

function sameRect(a: GridRect, b: GridRect): boolean {
  return a.top === b.top && a.left === b.left && a.bottom === b.bottom && a.right === b.right;
}
