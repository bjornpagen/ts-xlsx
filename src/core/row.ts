// A handle on one row of a worksheet: its formatting and its cells, reached by row number.
//
// A handle, not a record. `Worksheet` keeps the authoritative stores (the cell grid and the sparse
// map of per-row formatting) and a `Row` reads and writes straight through to them, so two handles
// on the same number always agree and neither can hold a stale copy. That is the whole reason this
// is not a snapshot: a row object that copied its cells out would be the shape of the merge-loss
// class of bug the model contract exists to prevent.
//
// Position is fixed at construction, exactly as `Cell` fixes `(row, col)`. `sheet.getRow(3)` means
// "row 3" and keeps meaning row 3: a splice that moves content past it does not carry the handle
// along, any more than it re-points a `Cell`.
//
// Formatting is created on write, never on read. Asking for `sheet.getRow(500)` costs nothing and
// does not extend the used range; assigning `height` is what materialises the record.

import {assertColumnInBounds, assertRowInBounds, columnToNumber} from './address.ts';
import {AxisHandle} from './axis-handle.ts';
import type {Cell} from './cell.ts';
import {type AssertNever, INTERNAL} from './internal.ts';
import type {Fill} from './style.ts';
import type {CellValue} from './value.ts';
import type {RowProperties, Worksheet} from './worksheet.ts';

export class Row extends AxisHandle<RowProperties> {
  readonly #sheet: Worksheet;

  /** 1-based row number. Fixed for this handle's lifetime. */
  readonly number: number;

  /** @throws {RangeError} unless the number is an integer within Excel's row grid (1..1048576). */
  constructor(sheet: Worksheet, number: number) {
    super();
    assertRowInBounds(number);
    this.#sheet = sheet;
    this.number = number;
  }

  /**
   * The row's format record if it has one, else `undefined`: a read that never fabricates, so a
   * serializer can ask every row it visits whether there are attributes to emit without giving each
   * one an empty record. Read-only on purpose: {@link height} and its siblings are how a row is
   * formatted, and they create the record on first write.
   */
  get properties(): Readonly<RowProperties> | undefined {
    return this.propertiesOf();
  }

  protected override propertiesOf(): RowProperties | undefined {
    return this.#sheet[INTERNAL].rowPropertiesOf(this.number);
  }

  protected override ensureProperties(): RowProperties {
    return this.#sheet[INTERNAL].ensureRowProperties(this.number);
  }

  protected override dropProperties(): void {
    this.#sheet[INTERNAL].dropRowProperties(this.number);
  }

  /**
   * Row height in points; `undefined` leaves the sheet default in force.
   *
   * Not bounded here, deliberately: {@link MAX_ROW_HEIGHT} is what Excel accepts *being set*, but
   * the schema puts no ceiling on `ht` and this setter is also how the reader loads a foreign
   * file, so refusing a taller row would mean refusing a file Excel opens clean. Check against the
   * constant when authoring: Excel silently clamps a taller row on read, so a height above it is
   * one you state and do not get.
   */
  get height(): number | undefined {
    return this.read('height');
  }
  set height(height: number | undefined) {
    this.write('height', height);
  }

  /** Whether the row is hidden. */
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

  /** Whether this row is the collapsed summary of an outline group. */
  get collapsed(): boolean | undefined {
    return this.read('collapsed');
  }
  set collapsed(collapsed: boolean | undefined) {
    this.write('collapsed', collapsed);
  }

  /** Background fill for the row's cells that carry no fill of their own. */
  get fill(): Fill | undefined {
    return this.read('fill');
  }
  set fill(fill: Fill | undefined) {
    this.write('fill', fill);
  }

  /**
   * The cell at a column in this row, creating it on first access. The column is a 1-based index
   * (`row.getCell(2)`) or its letters (`row.getCell('B')`).
   *
   * Resolves through merges exactly as {@link Worksheet.getCell} does: addressing a cell covered by
   * a merged region yields that region's master.
   *
   * @throws {RangeError} if the column is not a positive integer.
   * @throws {SyntaxError} if the column letters do not name a column.
   */
  getCell(column: number | string): Cell {
    const index = typeof column === 'number' ? column : columnToNumber(column);
    // Straight to the merge-resolving lookup, as `Range` goes. Encoding an A1 string only for
    // `getCell` to decode it again cost two conversions and checked nothing the guard does not.
    assertColumnInBounds(index);
    return this.#sheet[INTERNAL].masterAt(this.number, index);
  }

  /**
   * The row's materialised cells in ascending column order. Sparse: a column never written to has
   * no cell here, and the array is a fresh snapshot of *which* cells exist; the cells themselves
   * are the live ones.
   */
  get cells(): readonly Cell[] {
    return this.#sheet[INTERNAL].rowCells(this.number);
  }

  /**
   * The row's values by position, index 0 being column A. Sparse in the same way {@link cells} is:
   * a column with no cell is a hole, which is what distinguishes "never written" from a cell
   * holding `null`.
   *
   * Assigning places each value it names and leaves every other column untouched: a hole or an
   * explicit `undefined` skips that column, and a shorter array does not clear the tail. These are
   * {@link Worksheet.addRow}'s rules, deliberately: `values` is that same row shape addressed by
   * number rather than appended. To *replace* a row, including clearing what it held, splice it:
   * `sheet.spliceRows(n, 1, values)`.
   */
  get values(): (CellValue | undefined)[] {
    const values: (CellValue | undefined)[] = [];
    for (const cell of this.cells) values[cell.col - 1] = cell.value;
    return values;
  }
  set values(values: (CellValue | undefined)[]) {
    values.forEach((value, index) => {
      if (value !== undefined) this.getCell(index + 1).value = value;
    });
  }
}

/**
 * Compile-time proof that {@link Row} mirrors every {@link RowProperties} field. A field added to
 * the record without an accessor here resolves this to that field's name, which does not satisfy
 * `never`, so the build fails naming what is unreachable. Without it the mirror would rot silently:
 * the record would carry the new field, the codecs would read and write it, and the public handle
 * would simply never mention it.
 */
export type EveryRowPropertyIsMirrored = AssertNever<
  Exclude<keyof RowProperties, keyof Row & keyof RowProperties>
>;
