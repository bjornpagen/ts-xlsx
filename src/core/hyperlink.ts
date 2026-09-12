// A sheet's hyperlinks, each bound to the cells it covers.
//
// OOXML keeps a link beside the grid, in the sheet's `<hyperlinks>` list, naming the cells it covers by
// reference, and the model keeps it the same way. It used to be a kind of cell value, a label with a
// destination, and that was wrong twice. It made the link a label, so a link over a number, a date or a
// formula, where Excel puts one as readily as over text, had no value to be and was dropped on read. And
// it pinned the link to its top-left cell, where Excel 16.0 treats it as a range: a row inserted inside
// the range grows it, a delete of its first row shrinks it rather than taking it, and a second link over
// one of its cells is kept beside it (`test/corpus/fixtures/excel-oracle/hyperlinks-through-edits.json`).

import {AuthoringError, quoted} from '../errors.ts';
import {
  boundedRect,
  decodeRange,
  encodeAddress,
  encodeRect,
  type GridRect,
  tryDecodeRange,
} from './address.ts';
import {replaceContents} from './containers.ts';
import {type AxisSplice, shiftRect} from './grid-shift.ts';

/** A hyperlink and the cells it covers. */
export interface Hyperlink {
  /** The cells the link covers: one cell (`'B2'`) or a rectangle of cells (`'D1:H1'`). */
  readonly ref: string;
  /**
   * Where the link goes: a URL or a path, or a `#`-prefixed place in this workbook such as
   * `#Summary!A1` or `#TaxRate`.
   */
  readonly target: string;
  /** The text shown when the pointer rests on the link. */
  readonly tooltip?: string;
}

/** The links on one sheet, in the order they were added, each with the rectangle it covers. */
export class HyperlinkOverlay {
  readonly #entries: Hyperlink[] = [];
  // The rectangle each entry covers, in step with `#entries`, so neither a lookup nor a splice decodes
  // a reference per link.
  readonly #rects: GridRect[] = [];

  /**
   * Add a link, replacing one over the same cells. The `ref` is stored in its canonical spelling.
   *
   * @throws {SyntaxError} if `ref` is not a reference.
   * @throws {AuthoringError} if `ref` names a whole row or column rather than cells.
   */
  add(link: Hyperlink): void {
    const rect = boundedRect(decodeRange(link.ref));
    if (rect === undefined) {
      throw new AuthoringError(
        `hyperlink range ${quoted(link.ref)} names a whole row or column, not cells`,
      );
    }
    const stored = storedLink(link, refOf(rect));
    // Removed and pushed rather than replaced in place: the replacement is the link added last, which
    // is what decides between it and a wider link covering the same cells.
    this.#removeRef(stored.ref);
    this.#entries.push(stored);
    this.#rects.push(rect);
  }

  /** Remove the link whose `ref` names exactly these cells, and report whether there was one. */
  remove(ref: string): boolean {
    const decoded = tryDecodeRange(ref);
    const rect = decoded === undefined ? undefined : boundedRect(decoded);
    return rect !== undefined && this.#removeRef(refOf(rect));
  }

  /** The links on the sheet, in the order they were added. */
  get entries(): readonly Hyperlink[] {
    return this.#entries;
  }

  /** The link covering the 1-based `col`/`row`, the last added of any that do, or `undefined`. */
  at(col: number, row: number): Hyperlink | undefined {
    for (let index = this.#entries.length - 1; index >= 0; index--) {
      const rect = this.#rects[index];
      if (
        rect !== undefined &&
        col >= rect.left &&
        col <= rect.right &&
        row >= rect.top &&
        row <= rect.bottom
      ) {
        return this.#entries[index];
      }
    }
    return undefined;
  }

  /**
   * Move every link through a row or column splice. A link grows when lines are inserted inside it and
   * shrinks when some of its lines are deleted, and one the delete takes whole goes with it.
   */
  shift(splice: AxisSplice): void {
    const entries: Hyperlink[] = [];
    const rects: GridRect[] = [];
    this.#entries.forEach((entry, index) => {
      const rect = this.#rects[index];
      const moved = rect === undefined ? undefined : shiftRect(rect, splice);
      if (moved === undefined) return;
      const ref = refOf(moved);
      entries.push(ref === entry.ref ? entry : storedLink(entry, ref));
      rects.push(moved);
    });
    replaceContents(this.#entries, entries);
    replaceContents(this.#rects, rects);
  }

  /** Give row `to` a copy of every link that lies wholly within row `from`, as a row copy does. */
  copyRow(from: number, to: number): void {
    const copies = this.#entries.flatMap((entry, index) => {
      const rect = this.#rects[index];
      return rect === undefined || rect.top !== from || rect.bottom !== from
        ? []
        : [storedLink(entry, refOf({...rect, top: to, bottom: to}))];
    });
    for (const copy of copies) this.add(copy);
  }

  /** Drop every link, leaving the overlay empty. */
  clear(): void {
    this.#entries.length = 0;
    this.#rects.length = 0;
  }

  #removeRef(ref: string): boolean {
    const index = this.#entries.findIndex((entry) => entry.ref === ref);
    if (index === -1) return false;
    this.#entries.splice(index, 1);
    this.#rects.splice(index, 1);
    return true;
  }
}

// A single cell is spelled as one, the way Excel writes a one-cell link's `ref`.
function refOf(rect: GridRect): string {
  return rect.top === rect.bottom && rect.left === rect.right
    ? encodeAddress(rect.left, rect.top)
    : encodeRect(rect);
}

// A copy holding exactly the link's own fields, so nothing a caller's object carried besides them is
// kept, and a caller mutating that object afterwards changes nothing here.
function storedLink(link: Hyperlink, ref: string): Hyperlink {
  return {
    ref,
    target: link.target,
    ...(link.tooltip === undefined ? {} : {tooltip: link.tooltip}),
  };
}
