// Inclusive grid rectangles and the geometry a worksheet uses to reason about merged regions and the
// `sqref` ranges that overlays (data validations, conditional formats) apply to: decoding an OOXML
// `sqref` into containment rectangles, re-anchoring one through a splice, and collapsing the values a
// new merge covers.
//
// These used to be private methods on `Worksheet`, which split merge geometry across two files: the
// rectangle type and the overlap test lived here, while the operations that consume them lived there.
// Reasoning about "what does a merge do to the grid" meant reading both. The storage arrives as a
// parameter, so these stay pure functions of the rects and rows handed in. Resolving a covered
// position to its region's master is the one that did not stay a scan; it lives on `MergeIndex`.

import {
  boundedRect,
  decodeRange,
  encodeCornerRef,
  type GridRect,
  rectsOverlap,
  tryDecodeRange,
} from './address.ts';
import type {Cell} from './cell.ts';
import {type AxisSplice, shiftSpan} from './grid-shift.ts';

/** A merged region, as the {@link GridRect} every range-shaped thing in the library is. */
export type MergeRect = GridRect;

/**
 * Each merged range that overlaps a table, paired with the first table it overlaps. Excel forbids a
 * merge inside a table: the writer refuses that geometry and the reader drops the merge, the repair
 * Excel makes on load. Both used to spell the overlap test out by hand. The answer is a list of its
 * own, so the reader can unmerge while it walks it rather than the live `sheet.merges`. An unbounded
 * whole-row/column merge carries no rectangle and overlaps nothing.
 */
export function mergesOverlappingTables<T extends {readonly region: GridRect}>(
  merges: readonly string[],
  tables: readonly T[],
): {readonly merge: string; readonly table: T}[] {
  const overlapping: {merge: string; table: T}[] = [];
  if (tables.length === 0) return overlapping;
  for (const merge of merges) {
    const rect = boundedRect(decodeRange(merge));
    if (rect === undefined) continue;
    const table = tables.find((candidate) => rectsOverlap(rect, candidate.region));
    if (table !== undefined) overlapping.push({merge, table});
  }
  return overlapping;
}

/**
 * Drop any value already sitting in a merge's covered non-anchor cells, keeping only the top-left
 * anchor, the collapse Excel performs on merge. A leftover covered value would serialise as a
 * populated `<c>` under the range's `<mergeCell>` ref, the geometry that trips Excel's repair prompt.
 * Styles are untouched: a border spanning the merged region rides the covered cells.
 */
export function clearCoveredValues(rows: Map<number, Map<number, Cell>>, rect: MergeRect): void {
  // Both loops walk what is populated and test it against the rectangle, never the rectangle itself:
  // a file may declare `A1:A1048576`, and the work that costs must be the cells that exist, not the
  // area claimed. Walking the rect is how a few kilobytes of `<mergeCell>` buys minutes of CPU.
  for (const [row, cols] of rows) {
    if (row < rect.top || row > rect.bottom) continue;
    for (const [col, covered] of cols) {
      if (col < rect.left || col > rect.right) continue;
      if (row === rect.top && col === rect.left) continue;
      covered.value = null;
    }
  }
}

/** Decode an OOXML `sqref` (one or more space-separated ranges) into containment rectangles. A whole
 * column or row leaves one axis unbounded, so its missing edges open to `Infinity` rather than
 * clamping: a cell anywhere down the column still resolves inside it. An area that names no region
 * that can exist contributes no rectangle, the reader's rule for every other foreign attribute. */
export function decodeSqrefRects(sqref: string): MergeRect[] {
  const rects: MergeRect[] = [];
  for (const part of sqref.split(/\s+/)) {
    if (part === '') continue;
    const decoded = tryDecodeRange(part);
    if (decoded === undefined) continue;
    const {top, left, bottom, right} = decoded;
    rects.push({
      top: top ?? 1,
      left: left ?? 1,
      bottom: bottom ?? Infinity,
      right: right ?? Infinity,
    });
  }
  return rects;
}

/**
 * Re-anchor an OOXML `sqref` through a splice, the inverse of {@link decodeSqrefRects}. Returns
 * `undefined` when the splice deleted every area the `sqref` named: an empty `sqref` is not
 * writable, so the entry holding it goes too.
 *
 * Each space-separated area shifts on its own, and one the splice does not move is returned as the
 * *original text*. That matters for a file the library did not author: `B:B` and `B1:B1048576` decode
 * identically, so a re-encode would rewrite a foreign spelling and cost the byte-clean round trip.
 */
export function shiftSqref(sqref: string, splice: AxisSplice): string | undefined {
  const areas: string[] = [];
  for (const area of sqref.split(/\s+/)) {
    if (area === '') continue;
    const shifted = shiftSqrefArea(area, splice);
    if (shifted !== undefined) areas.push(shifted);
  }
  return areas.length > 0 ? areas.join(' ') : undefined;
}

function shiftSqrefArea(area: string, splice: AxisSplice): string | undefined {
  const decoded = tryDecodeRange(area);
  // An area this library cannot read is one it cannot move either, and the `sqref` it came from is
  // still the file's own text: return it untouched rather than dropping a region on a guess.
  if (decoded === undefined) return area;
  const {top, left, bottom, right} = decoded;
  const rowAxis = splice.axis === 'row';
  const [lo, hi] = rowAxis ? [top, bottom] : [left, right];
  // An area unbounded on the spliced axis covers every line of it, so the splice cannot move it: `B:B`
  // after a row insert is still `B:B`, never `B2:B1048577`. The other axis may stay unbounded, which
  // is why this moves a span rather than a rectangle: `1:5` must come back spelled without columns.
  if (lo === undefined || hi === undefined) return area;
  const moved = shiftSpan(lo, hi, splice);
  if (moved === undefined) return undefined;
  if (moved.lo === lo && moved.hi === hi) return area;
  const [tl, br] = rowAxis
    ? [encodeCornerRef(left, moved.lo), encodeCornerRef(right, moved.hi)]
    : [encodeCornerRef(moved.lo, top), encodeCornerRef(moved.hi, bottom)];
  // A single-cell area stays a single cell: `B5` must not come back as `B6:B6`.
  return area.includes(':') ? `${tl}:${br}` : tl;
}
