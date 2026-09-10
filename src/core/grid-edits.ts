// Structural-edit machinery: the splice arithmetic that inserts or deletes whole rows and columns
// and keeps everything anchored to the grid moving in step: line metadata, merged ranges, tables,
// anchored images, shared-formula clones, and the range-bound overlays (data validations, conditional
// formats, comment threads, the autofilter). It is isolated from Worksheet because it is pure grid
// mechanics: it holds the sheet's storage containers by reference and mutates them in place, and
// touches none of the public cell API. Worksheet builds the cells an insert introduces, then hands
// the pre-built rows (or the raw column values) here for the shift.

import {boundedRect, decodeRange, encodeAddress, encodeRect, tryDecodeCellRef} from './address.ts';
import {type AutoFilter, shiftAutoFilter} from './autofilter.ts';
import {Cell, copyCellContent} from './cell.ts';
import type {ConditionalFormattingOverlay} from './conditional-formatting-overlay.ts';
import {replaceContents} from './containers.ts';
import type {DataValidationOverlay} from './data-validation-overlay.ts';
import {type AxisSplice, isDeletedSpan, shiftIndex, shiftRect} from './grid-shift.ts';
import {type AnchoredImage, type AnchorPoint, type ImageAnchor, isOneCellAnchor} from './image.ts';
import type {MergeRect} from './merge.ts';
import {positionalPlacements} from './row-input.ts';
import type {Table} from './table.ts';
import {type CellValue, isSharedFormulaValue, type SharedFormulaValue} from './value.ts';
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
  readonly images: AnchoredImage[];
  readonly dataValidations: DataValidationOverlay;
  readonly conditionalFormattings: ConditionalFormattingOverlay;
  readonly comments: WorksheetComments;
  readonly autoFilter: AutoFilterSlot;
}

export class GridEdits {
  readonly #rows: Map<number, Map<number, Cell>>;
  readonly #rowProperties: Map<number, RowProperties>;
  readonly #columns: Map<number, ColumnProperties>;
  readonly #merges: WorksheetMerges;
  readonly #tables: Table[];
  readonly #images: AnchoredImage[];
  readonly #dataValidations: DataValidationOverlay;
  readonly #conditionalFormattings: ConditionalFormattingOverlay;
  readonly #comments: WorksheetComments;
  readonly #autoFilter: AutoFilterSlot;

  constructor(storage: GridStorage) {
    this.#rows = storage.rows;
    this.#rowProperties = storage.rowProperties;
    this.#columns = storage.columns;
    this.#merges = storage.merges;
    this.#tables = storage.tables;
    this.#images = storage.images;
    this.#dataValidations = storage.dataValidations;
    this.#conditionalFormattings = storage.conditionalFormattings;
    this.#comments = storage.comments;
    this.#autoFilter = storage.autoFilter;
  }

  // Apply a delete-then-insert to the row grid: surviving rows below the edit shift by
  // `inserted.length - count`, deleted rows drop out, and the pre-built inserted rows land at `start`.
  // Row metadata, merged ranges and everything else anchored to the grid shift the same way, so a
  // formatting-only row, a covered merge or a dropdown stays aligned with the data it describes.
  spliceRows(start: number, count: number, inserted: Map<number, Cell>[]): void {
    const splice: AxisSplice = {axis: 'row', start, count, delta: inserted.length - count};
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

    this.#shiftLineProperties(this.#rowProperties, splice);
    this.#shiftMerges(splice);
    this.#shiftTables(splice);
    this.#shiftImages(splice);
    this.#reanchorSharedFormulas(splice);
    this.#shiftRangeBoundOverlays(splice);
  }

  // Apply a delete-then-insert to the column grid: cells left of the edit stay, cells at or beyond the
  // deleted span shift by `inserts.length - count` carrying their content, and the inserted column
  // values materialise as fresh cells at `start`. Column metadata, merges, tables, images,
  // shared-formula clones and the range-bound overlays re-anchor the same way.
  spliceColumns(start: number, count: number, inserts: CellValue[][]): void {
    const splice: AxisSplice = {axis: 'col', start, count, delta: inserts.length - count};
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
    this.#shiftLineProperties(this.#columns, splice);
    this.#shiftMerges(splice);
    this.#shiftTables(splice);
    this.#shiftImages(splice);
    this.#reanchorSharedFormulas(splice);
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

  // Re-anchor shared-formula clones through a splice on the given axis. A clone stores its master's
  // absolute address; when the splice shifts the master, that stored address goes stale and the writer
  // would reject the clone as orphaned. Applying the same shift the grid used keeps each clone pointed
  // at its master's new cell. A master whose axis coordinate falls in the deleted span clamps to the
  // cut line like a merge edge: a genuinely orphaned clone the writer then reports legibly.
  #reanchorSharedFormulas(splice: AxisSplice): void {
    for (const cols of this.#rows.values()) {
      for (const cell of cols.values()) {
        const value = cell.value;
        if (!isSharedFormulaValue(value)) continue;
        const master = tryDecodeCellRef(value.sharedFormula);
        if (master === undefined) continue;
        const anchored =
          splice.axis === 'row'
            ? encodeAddress(master.col, shiftIndex(master.row, splice))
            : encodeAddress(shiftIndex(master.col, splice), master.row);
        if (anchored === value.sharedFormula) continue;
        const reanchored: SharedFormulaValue = {...value, sharedFormula: anchored};
        cell.value = reanchored;
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
      return {imageId: image.imageId, anchor};
    });
    replaceContents(this.#images, moved);
  }
}
