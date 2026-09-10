// A sheet's merged regions: the declared range strings, the rectangles the bounded ones decode to,
// and the index that answers questions about them without walking the list.
//
// The last of `Worksheet`'s overlays to be lifted out (see "the two model classes delegate their
// state" in docs/architecture.md). It passes that section's test easily: it reaches nothing outside
// itself. The two things a merge does to the *grid* -- collapsing the values it covers, and widening
// the used extent -- stay on `Worksheet`, which owns the grid, and are driven by the rectangle
// {@link add} hands back. Pulling them in here would have given the slice two edges to buy one line.
//
// Three representations rather than one, and each earns its place. The range *strings* are what the
// sheet declares and what round-trips, including the unbounded `A:A` form that names no rectangle.
// The rectangles are what geometry is done against. The index is what keeps overlap-checking and
// covered-address resolution off a linear scan; see `merge-index.ts` for why that mattered.

import {AuthoringError, InternalError, quoted} from '../errors.ts';
import {boundedRect, decodeRange, encodeRect} from './address.ts';
import {replaceContents} from './containers.ts';
import {MergeIndex} from './merge-index.ts';
import type {MergeRect} from './merge.ts';

/** What removing a declared range did, which is two answers because a range and a rectangle are not
 * the same thing: an unbounded whole-row/column merge is declared but has no rectangle, so dropping
 * it leaves every geometry derived from the rectangles (the used extent, the index) still valid. */
export interface MergeRemoval {
  /** Whether a merge with this exact range string was declared. */
  readonly existed: boolean;
  /** Whether a rectangle went with it, so anything derived from the rectangles is now stale. */
  readonly rectsChanged: boolean;
}

/**
 * A merged range as the sheet stores it, with the rectangle it covers: `$` anchors dropped and the
 * corners top-left first, the spelling the writer emits and a splice produces. Stored as typed, a merge
 * made as `$A$1:$B$2` could not be removed as `A1:B2`, and stopped matching its own spelling once an
 * unrelated splice rewrote it.
 *
 * @throws {SyntaxError} if the range is unparseable or carries a sheet prefix. A `<mergeCell>` names a
 *   region of the sheet it sits on, so a prefix was either redundant or silently wrong.
 */
function canonicalMerge(range: string): {
  readonly canonical: string;
  readonly rect: MergeRect | undefined;
} {
  const decoded = decodeRange(range);
  if (decoded.sheetName !== undefined) {
    throw new SyntaxError(
      `merged range ${quoted(range)} names worksheet ${quoted(decoded.sheetName)}: a merge belongs to the sheet it is made on`,
    );
  }
  // `MergeRect` and the narrowed rectangle are the same four inclusive bounds, so the decode is
  // already the record this needs.
  const rect: MergeRect | undefined = boundedRect(decoded);
  return {canonical: rect === undefined ? decoded.dimensions : encodeRect(rect), rect};
}

export class WorksheetMerges {
  readonly #ranges: string[] = [];
  // Parallel to #ranges for the bounded entries only, so addressing a covered cell resolves to its
  // region's master without re-parsing the range string on every access, and a new merge can be
  // checked for overlap against the existing ones. An unbounded whole-row/column merge is still
  // declared but participates in neither.
  readonly #rects: MergeRect[] = [];
  readonly #index = new MergeIndex(this.#rects);

  /** The declared ranges, in the order they were added. Live, not a copy. */
  get ranges(): readonly string[] {
    return this.#ranges;
  }

  /**
   * The bounded regions as rectangles. Live, and its *identity* is stable for the sheet's lifetime:
   * {@link UsedExtent} holds this array by reference, so every rewrite here goes through
   * `replaceContents` rather than reassignment.
   */
  get rects(): readonly MergeRect[] {
    return this.#rects;
  }

  /**
   * Declare a merged range, returning the rectangle it covers, or `undefined` for an unbounded
   * whole-row/column range, which is declared and overlap-checks against nothing. The range is stored
   * in canonical form (see {@link canonicalMerge}).
   *
   * @throws {SyntaxError} if the range is unparseable or names a worksheet.
   * @throws {AuthoringError} if the range overlaps an already-merged region. Excel forbids
   *   overlapping merges and writes such geometry as a file it then offers to repair.
   */
  add(range: string): MergeRect | undefined {
    const {canonical, rect} = canonicalMerge(range);
    if (rect !== undefined) {
      if (this.#index.overlapping(rect) !== undefined) {
        throw new AuthoringError(
          `merged range ${quoted(range)} overlaps an existing merged region`,
        );
      }
      this.#rects.push(rect);
      this.#index.note(rect);
    }
    this.#ranges.push(canonical);
    return rect;
  }

  /**
   * Drop a declared range and, with it, the rectangle it covers. The inverse of {@link add}, and it
   * matches however the range is spelled, since both sides are compared in canonical form.
   *
   * @throws {SyntaxError} if the range is unparseable or names a worksheet.
   */
  remove(range: string): MergeRemoval {
    const {canonical, rect} = canonicalMerge(range);
    const index = this.#ranges.indexOf(canonical);
    if (index === -1) return {existed: false, rectsChanged: false};
    this.#ranges.splice(index, 1);
    if (rect === undefined) return {existed: true, rectsChanged: false};
    const {top, left, bottom, right} = rect;
    const at = this.#rects.findIndex(
      (r) => r.top === top && r.left === left && r.bottom === bottom && r.right === right,
    );
    // A bounded range and its rectangle are added and removed together, so a declared one always has
    // its rectangle; splicing at -1 would drop some other merge's.
    if (at === -1) {
      throw new InternalError(
        `merged range ${quoted(canonical)} was declared without its rectangle`,
      );
    }
    this.#rects.splice(at, 1);
    this.#index.invalidate();
    return {existed: true, rectsChanged: true};
  }

  /**
   * Resolve a position to the top-left of the region covering it, or to itself when none does. This
   * is what makes a write through a covered address land on the region's master.
   */
  masterOf(row: number, col: number): {row: number; col: number} {
    return this.#index.masterOf(row, col);
  }

  /** Replace every declared range and rectangle at once: what a structural splice re-anchors to. */
  replaceAll(ranges: readonly string[], rects: readonly MergeRect[]): void {
    replaceContents(this.#ranges, ranges);
    replaceContents(this.#rects, rects);
    this.#index.invalidate();
  }

  clear(): void {
    this.#ranges.length = 0;
    this.#rects.length = 0;
    this.#index.invalidate();
  }
}
