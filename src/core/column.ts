// A handle on one column of a worksheet: its formatting and its cells, reached by column index.
//
// The other axis of `Row`, and the same contract: a live view over the worksheet's stores, position
// fixed at construction, formatting created on write rather than on read. See `core/row.ts` for why
// it is a handle rather than a record.
//
// A column carries more than a row does: besides the geometry (width, visibility, outline) it holds
// the six `CellStyle` facets as *defaults* for its cells, which is why the mirror below is twice the
// length of the row's.

import {assertColumnInBounds, assertRowInBounds, numberToColumn} from './address.ts';
import {AxisHandle} from './axis-handle.ts';
import type {Cell} from './cell.ts';
import {type AssertNever, INTERNAL} from './internal.ts';
import type {Alignment, Border, Fill, Font, Protection} from './style.ts';
import type {CellValue} from './value.ts';
import type {ColumnProperties, Worksheet} from './worksheet.ts';

export class Column extends AxisHandle<ColumnProperties> {
  readonly #sheet: Worksheet;

  /** 1-based column index. Fixed for this handle's lifetime. */
  readonly index: number;

  /** @throws {RangeError} unless the index is an integer within Excel's column grid (1..16384). */
  constructor(sheet: Worksheet, index: number) {
    super();
    assertColumnInBounds(index);
    this.#sheet = sheet;
    this.index = index;
  }

  /** The column's letters (`"B"`): the spreadsheet-facing name for {@link index}. */
  get letter(): string {
    return numberToColumn(this.index);
  }

  /**
   * The column's format record if it has one, else `undefined`: a read that never fabricates, so a
   * serializer can ask every column it visits whether there are attributes to emit without giving
   * each one an empty record. Read-only on purpose: {@link width} and its siblings are how a column
   * is formatted, and they create the record on first write.
   */
  get properties(): Readonly<ColumnProperties> | undefined {
    return this.propertiesOf();
  }

  protected override propertiesOf(): ColumnProperties | undefined {
    return this.#sheet[INTERNAL].columnPropertiesOf(this.index);
  }

  protected override ensureProperties(): ColumnProperties {
    return this.#sheet[INTERNAL].ensureColumnProperties(this.index);
  }

  protected override dropProperties(): void {
    this.#sheet[INTERNAL].dropColumnProperties(this.index);
  }

  /**
   * Stable key naming this column so a keyed-object row (see {@link Worksheet.addRow}) can place a
   * value under it by name rather than position. In-memory only: never serialized to OOXML.
   */
  get key(): string | undefined {
    return this.read('key');
  }
  set key(key: string | undefined) {
    this.write('key', key);
  }

  /**
   * Column width in character units: digits of the workbook default font's maximum digit width,
   * so what one unit measures moves with that font. `undefined` leaves the sheet default in force.
   *
   * Not bounded here, for the same reason {@link Row.height} is not, and for a stronger one:
   * {@link MAX_COLUMN_WIDTH} bounds what Excel accepts being set, not what a file may hold. Excel
   * reads a wider column back unchanged and re-saves it verbatim, so a bound here would refuse a
   * width Excel itself preserves.
   */
  get width(): number | undefined {
    return this.read('width');
  }
  set width(width: number | undefined) {
    this.write('width', width);
  }

  /** Whether the column is hidden. */
  get hidden(): boolean | undefined {
    return this.read('hidden');
  }
  set hidden(hidden: boolean | undefined) {
    this.write('hidden', hidden);
  }

  /** Outline (grouping) depth; 0 or `undefined` means ungrouped. */
  get outlineLevel(): number | undefined {
    return this.read('outlineLevel');
  }
  set outlineLevel(outlineLevel: number | undefined) {
    this.write('outlineLevel', outlineLevel);
  }

  /** Whether this column is the collapsed summary of an outline group. */
  get collapsed(): boolean | undefined {
    return this.read('collapsed');
  }
  set collapsed(collapsed: boolean | undefined) {
    this.write('collapsed', collapsed);
  }

  /** Default fill for the column's cells that set none of their own. */
  get fill(): Fill | undefined {
    return this.read('fill');
  }
  set fill(fill: Fill | undefined) {
    this.write('fill', fill);
  }

  /** Default number format for the column's cells that set none of their own. */
  get numFmt(): string | undefined {
    return this.read('numFmt');
  }
  set numFmt(numFmt: string | undefined) {
    this.write('numFmt', numFmt);
  }

  /** Default font for the column's cells that set none of their own. */
  get font(): Font | undefined {
    return this.read('font');
  }
  set font(font: Font | undefined) {
    this.write('font', font);
  }

  /** Default border for the column's cells that set none of their own. */
  get border(): Border | undefined {
    return this.read('border');
  }
  set border(border: Border | undefined) {
    this.write('border', border);
  }

  /** Default alignment for the column's cells that set none of their own. */
  get alignment(): Alignment | undefined {
    return this.read('alignment');
  }
  set alignment(alignment: Alignment | undefined) {
    this.write('alignment', alignment);
  }

  /** Default protection flags for the column's cells that set none of their own. */
  get protection(): Protection | undefined {
    return this.read('protection');
  }
  set protection(protection: Protection | undefined) {
    this.write('protection', protection);
  }

  /**
   * The cell at a 1-based row number in this column, creating it on first access. Resolves through
   * merges exactly as {@link Worksheet.getCell} does.
   *
   * @throws {RangeError} if the row is not a positive integer.
   */
  getCell(row: number): Cell {
    // Straight to the merge-resolving lookup, as `Range` goes, rather than through an A1 string that
    // `getCell` would only decode again.
    assertRowInBounds(row);
    return this.#sheet[INTERNAL].masterAt(row, this.index);
  }

  /**
   * The column's materialised cells in ascending row order. Sparse: a row that never wrote to this
   * column has no cell here.
   */
  get cells(): readonly Cell[] {
    return this.#sheet[INTERNAL].columnCells(this.index);
  }

  /**
   * The column's values by position, index 0 being row 1. Sparse in the same way {@link cells} is:
   * a row with no cell in this column is a hole, which is what distinguishes "never written" from a
   * cell holding `null`.
   *
   * Assigning places each value it names and leaves every other row untouched, mirroring
   * {@link Row.values}: a hole or an explicit `undefined` skips that row, and a shorter array does
   * not clear the tail.
   */
  get values(): (CellValue | undefined)[] {
    const values: (CellValue | undefined)[] = [];
    for (const cell of this.cells) values[cell.row - 1] = cell.value;
    return values;
  }
  set values(values: (CellValue | undefined)[]) {
    values.forEach((value, index) => {
      if (value !== undefined) this.getCell(index + 1).value = value;
    });
  }
}

/**
 * Compile-time proof that {@link Column} mirrors every {@link ColumnProperties} field, including
 * the six inherited `CellStyle` facets, so a seventh facet reaches this handle the moment it joins
 * the tuple. See the counterpart on `Row` for why the mirror needs proving rather than reviewing.
 */
export type EveryColumnPropertyIsMirrored = AssertNever<
  Exclude<keyof ColumnProperties, keyof Column & keyof ColumnProperties>
>;
