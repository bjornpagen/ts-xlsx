// Which `xf` record a cell's format resolves to, in one place for all three cell readers.
//
// A cell may inherit its format rather than state it: one with no format of its own takes its row's
// when the row is marked as carrying one, failing that its column's, and failing that the workbook's
// default format, xf 0. That is the order Excel applies, and it is not decoration: the value decoders
// read `numFmt` off the resolved style to decide whether a serial number is a *date*, so a cell whose
// format is inherited decodes to a different **type** depending on whether the resolution ran.
//
// The buffered and streaming XML readers and the BIFF12 reader each finished this by hand, and they
// disagreed on the last step: two fell back to xf 0 and the streaming reader to no style at all, so a
// workbook whose default format is a date read `2023-03-15` through `readXlsx` and `45000` through
// `readSheetRows`. The fallback is owned here, so no caller can decide it differently. The inputs are
// numbers because the two codecs spell them differently, XML as attributes and BIFF12 as record fields.

import type {XfStyle} from './xf-style.ts';

/**
 * The cell → row → column → xf 0 style resolution for one worksheet, fed in the order every reader
 * already visits a sheet: column formats before any cell refers to them, a row before its own cells.
 * Columns are 1-based, and a negative style index means "declares no format".
 */
export class CellStyleResolver {
  // An index rather than the resolved record: column formats are read before any cell asks for one,
  // and most of them are never asked for.
  readonly #columnStyle = new Map<number, number>();
  // The open row's format when its cells inherit it, else -1.
  #rowStyle = -1;

  /** Record the format a span of columns gives its cells. */
  noteColumnSpan(first: number, last: number, styleIndex: number): void {
    if (styleIndex < 0) return;
    for (let index = first; index <= last; index++) this.#columnStyle.set(index, styleIndex);
  }

  /**
   * Open a row. Its format reaches its cells only when the row is marked `customFormat`; without that
   * flag the format describes the row itself and is not inherited.
   */
  openRow(styleIndex: number, customFormat: boolean): void {
    this.#rowStyle = customFormat ? styleIndex : -1;
  }

  /** Close the row, so a cell outside one never inherits the last row's format. */
  closeRow(): void {
    this.#rowStyle = -1;
  }

  /**
   * The style a cell resolves to: its own, else its row's, else its column's, else xf 0. `undefined`
   * only when the index it resolves to names no record in `xfStyles`, a damaged or truncated table.
   */
  styleFor(
    col: number,
    cellStyleIndex: number,
    xfStyles: ReadonlyArray<XfStyle>,
  ): XfStyle | undefined {
    const index = this.#indexFor(col, cellStyleIndex);
    return xfStyles[index < 0 ? 0 : index];
  }

  /** Whether the cell, its row or its column declares a format, rather than the cell falling to xf 0. */
  declaresFormat(col: number, cellStyleIndex: number): boolean {
    return this.#indexFor(col, cellStyleIndex) >= 0;
  }

  #indexFor(col: number, cellStyleIndex: number): number {
    if (cellStyleIndex >= 0) return cellStyleIndex;
    if (this.#rowStyle >= 0) return this.#rowStyle;
    return this.#columnStyle.get(col) ?? -1;
  }
}
