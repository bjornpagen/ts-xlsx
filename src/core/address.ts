// Address decoding: the foundational spreadsheet primitive.
//
// Every higher layer (cells, ranges, defined names, formulas) is ultimately
// addressed by a `col`/`row` pair, so this module is the first thing the rewrite
// builds and the bedrock everything else stands on. It is pure, allocation-bounded,
// and has no I/O or platform dependency.
//
// The honesty rule that drives the shape here: an axis that a reference does not
// mention is `undefined`, never a sentinel. A whole-row reference (`$1`) has no
// column; a whole-column reference (`$A:$A`) has no row. Legacy code let those
// absent axes decay into `NaN`/`"undefined"` and leak into serialized addresses
// (`"$undefined$1"`, `"NaN:NaN"`), the exact defect the corpus locks against.
//
// The decoders come in two arities and two temperaments. `decodeAddress`/`decodeCellRef`/
// `decodeRange` face a caller and throw, because a reference a caller supplied is a mistake at the
// call. `tryDecodeCellRef`/`tryDecodeRange` face a file and return `undefined`, because a file this
// library did not write is allowed to be wrong and losing one attribute beats losing the sheet.
// Read-side code uses the second pair and nothing else; the rule and the reasoning are `xml-scan.ts`'s.

import {quoted} from '../errors.ts';

/** Excel's column bounds: `A` (1) through `XFD` (16384). */
export const MAX_COLUMN = 16384;

/** Excel's row bound: 1 through 1048576. The other axis of {@link MAX_COLUMN}. */
export const MAX_ROW = 1048576;

/**
 * The same two bounds, zero-based: the last valid *index* on each axis.
 *
 * BIFF12 counts from zero throughout, so its reader and its formula decoder both need these, and both
 * were deriving them privately. Two derivations of one fact is one fewer than three, but it is still
 * the shape where a change to a limit reaches some of its consequences. The subtraction belongs beside
 * the number it subtracts from.
 */
export const MAX_COLUMN_INDEX = MAX_COLUMN - 1;
export const MAX_ROW_INDEX = MAX_ROW - 1;

// The numeric door into the grid. `numberToColumn`/`columnToNumber` bound a reference spelled in
// letters; these bound the same position spelled as a number, so `getColumn(16385)` and
// `getCell('XFE1')` refuse the same mistake with the same words. Native `RangeError` rather than
// `AuthoringError`: a single scalar out of range is what `errors.ts` reserves for native errors.
//
// "The same words" is the whole point and it had stopped being true. `Range` re-implemented the pair
// with a different message split into two throws, and `spliceRows` a third, so one mistake had three
// answers depending on which door the caller came through:
//
//     getRange(0, 1, 1, 1)  ->  "row 0 is out of bounds: rows start at 1"
//     getRow(0)             ->  "row 0 is out of bounds: Excel supports 1..1048576"
//     spliceRows(0, 1)      ->  "splice start 0 is out of bounds: rows start at 1"
//
// So the guard is axis-parameterised and every caller reaches it, rather than each spelling the two
// comparisons out with whatever bound it happened to have in hand.

/** The upper bound of each axis, so the guard below can be written once for both. */
const AXIS_BOUND = {row: MAX_ROW, column: MAX_COLUMN} as const;

/**
 * Refuse a coordinate that names no position on the grid.
 *
 * @throws {RangeError} unless `value` is an integer within the axis's bounds.
 */
export function assertAxisInBounds(axis: 'row' | 'column', value: number): void {
  const bound = AXIS_BOUND[axis];
  if (!Number.isInteger(value) || value < 1 || value > bound) {
    throw new RangeError(`${axis} ${value} is out of bounds: Excel supports 1..${bound}`);
  }
}

/** @throws {RangeError} unless `n` is an integer in `1..MAX_COLUMN`. */
export function assertColumnInBounds(n: number): void {
  assertAxisInBounds('column', n);
}

/** @throws {RangeError} unless `n` is an integer in `1..MAX_ROW`. */
export function assertRowInBounds(n: number): void {
  assertAxisInBounds('row', n);
}

/**
 * A rectangular block of the grid, as **inclusive** 1-based bounds on both axes.
 *
 * One declaration because inclusive-first/last is the convention every range-shaped thing in this
 * library follows, and three copies of a convention are three places it can drift. A merged region,
 * a table's extent and a {@link Range} handle are all this shape; what differs between them is what
 * the rectangle *means*, which is what their own names carry.
 */
export interface GridRect {
  /** 1-based row of the top edge. */
  readonly top: number;
  /** 1-based column of the left edge. */
  readonly left: number;
  /** 1-based row of the bottom edge, inclusive. */
  readonly bottom: number;
  /** 1-based column of the right edge, inclusive. */
  readonly right: number;
}

/** Whether two grid rectangles share at least one cell. */
export function rectsOverlap(a: GridRect, b: GridRect): boolean {
  return a.left <= b.right && b.left <= a.right && a.top <= b.bottom && b.top <= a.bottom;
}

/** A decoded single-cell reference. An axis the reference omits is `undefined`. */
export interface CellAddress {
  /** Canonical A1 form with `$` anchors stripped: e.g. `"B2"`, `"1"`, `"A"`. */
  readonly address: string;
  /** 1-based column, or `undefined` for a row-only reference (`$1`). */
  readonly col: number | undefined;
  /** 1-based row, or `undefined` for a column-only reference (`$A`). */
  readonly row: number | undefined;
}

/**
 * A decoded range reference. Corners are the min/max of the endpoints per axis;
 * an axis neither endpoint mentions (a whole-row or whole-column range) is
 * `undefined` on every corner and simply absent from `dimensions`.
 */
export interface RangeAddress {
  readonly top: number | undefined;
  readonly left: number | undefined;
  readonly bottom: number | undefined;
  readonly right: number | undefined;
  /** The originating sheet, present only when the reference carried one. */
  readonly sheetName?: string;
  readonly tl: CellAddress;
  readonly br: CellAddress;
  /** Canonical `tl:br` form: `"A1:B2"`, `"1:1"` (rows), `"A:A"` (columns). */
  readonly dimensions: string;
}

const SINGLE_REF = /^\$?([A-Z]*)\$?(\d*)$/;
// A leading `Sheet!` prefix: quoted (`'a''b'!`, doubled apostrophe escapes one) or
// bare (`Sheet1!`). Group 1 = quoted body, group 2 = bare name, group 3 = the rest.
const SHEET_PREFIX = /^(?:(?:'((?:[^']|'')*)')|([^'!]+))!(.*)$/;

/** Convert a 1-based column number to its letters (`1 → "A"`, `27 → "AA"`).
 *
 * @throws {RangeError} unless `n` is an integer in `1..MAX_COLUMN`. */
export function numberToColumn(n: number): string {
  // Through the guard rather than a second copy of its four lines: the note above
  // `assertColumnInBounds` says these two refuse the same mistake in the same words, and a call is
  // how that is true rather than how two literals happen to agree.
  assertColumnInBounds(n);
  let letters = '';
  let remaining = n;
  while (remaining > 0) {
    const digit = (remaining - 1) % 26;
    letters = String.fromCharCode(65 + digit) + letters;
    remaining = Math.floor((remaining - 1) / 26);
  }
  return letters;
}

// Uppercase only, one to three letters: the spelling Excel writes and the widest one the grid has a
// column for. `AAAA` is not a near-miss to be repaired, it is four letters where three fit.
const COLUMN_LETTERS = /^[A-Z]{1,3}$/;

/**
 * Convert column letters to a 1-based number (`"A" → 1`, `"AA" → 27`), or `undefined` when they name
 * no column that can exist: not letters at all, or past `XFD` (`"ZZZ"` is well-formed and is column
 * 18278, which the grid does not have). The tolerant half of the pair the module header describes,
 * for a caller reading a file rather than a caller's own argument.
 */
export function tryColumnToNumber(letters: string): number | undefined {
  if (!COLUMN_LETTERS.test(letters)) return undefined;
  let n = 0;
  for (let i = 0; i < letters.length; i++) {
    n = n * 26 + (letters.charCodeAt(i) - 64);
  }
  return n > MAX_COLUMN ? undefined : n;
}

/**
 * Convert column letters to a 1-based number (`"A" → 1`, `"AA" → 27`).
 *
 * @throws {RangeError} if the letters are malformed or name a column past `XFD`. The two are
 *   separate messages because a caller fixes them differently.
 */
export function columnToNumber(letters: string): number {
  const n = tryColumnToNumber(letters);
  if (n !== undefined) return n;
  if (!COLUMN_LETTERS.test(letters)) {
    throw new RangeError(`invalid column letters: ${quoted(letters)}`);
  }
  throw new RangeError(
    `column ${quoted(letters)} is out of bounds: Excel supports up to ${MAX_COLUMN} (XFD)`,
  );
}

/**
 * One corner of a range, as its reference text. An axis the corner omits contributes nothing, so a
 * whole-row corner spells `6` and a whole-column one spells `B`, rather than acquiring the other
 * axis from a sentinel. That is the honesty rule this module's header states, and letting an absent
 * axis decay into text is exactly how `"$undefined$1"` and `"NaN:NaN"` used to reach a file.
 *
 * The unbounded counterpart to {@link encodeAddress}, which names one cell and so requires both.
 */
export function encodeCornerRef(col: number | undefined, row: number | undefined): string {
  return `${col !== undefined ? numberToColumn(col) : ''}${row !== undefined ? row : ''}`;
}

/**
 * A bounded rectangle as its canonical `tl:br` A1 range, `"B2:D5"`. A one-cell rectangle still reads
 * as `"B2:B2"`, which is the form Excel writes and the form every consumer of these refs parses.
 *
 * The rendering half of {@link boundedRect}, and it belongs beside it for the same reason: ten sites
 * spelled `` `${encodeAddress(left, top)}:${encodeAddress(right, bottom)}` `` by hand, and the risk in
 * that expression is not its length but its argument order. `encodeAddress` takes column first and a
 * rectangle names its rows first, so the two orders are transposed with respect to each other at every
 * one of those sites.
 */
export function encodeRect(rect: GridRect): string {
  return `${encodeAddress(rect.left, rect.top)}:${encodeAddress(rect.right, rect.bottom)}`;
}

/**
 * A bounded rectangle as a formula's `ref` spells it: {@link encodeRect}'s `"B1:B3"`, except that one
 * cell is its bare address, `"B1"`. Excel writes an array formula over a single cell that way, and a
 * range that moves keeps the spelling it would be saved with.
 */
export function encodeRange(rect: GridRect): string {
  return rect.top === rect.bottom && rect.left === rect.right
    ? encodeAddress(rect.left, rect.top)
    : encodeRect(rect);
}

/** Build a {@link CellAddress} corner straight from optional numeric axes: the address string is
 * assembled from the parts we already hold, so no encode-then-decode round-trip is needed. An axis the
 * corner omits stays `undefined`; both absent yields the empty address (`""`). */
function makeCellAddress(col: number | undefined, row: number | undefined): CellAddress {
  return {address: encodeCornerRef(col, row), col, row};
}

/**
 * Decode a single cell/row/column reference into `{address, col, row}`. Anchoring
 * `$` signs are accepted and dropped; an absent axis is `undefined`.
 *
 * @throws {SyntaxError} if the reference mentions neither a column nor a row.
 * @throws {RangeError} if it names a column past `XFD` or a row outside `1..1048576`.
 */
export function decodeAddress(reference: string): CellAddress {
  const match = SINGLE_REF.exec(reference);
  if (!match) {
    throw new SyntaxError(`invalid cell reference: ${quoted(reference)}`);
  }
  const letters = match[1] ?? '';
  const digits = match[2] ?? '';
  if (letters.length === 0 && digits.length === 0) {
    throw new SyntaxError(`invalid cell reference: ${quoted(reference)}`);
  }
  const col = letters.length > 0 ? columnToNumber(letters) : undefined;
  const row = digits.length > 0 ? Number.parseInt(digits, 10) : undefined;
  // Bounded here, beside the column `columnToNumber` already bounds. A decoder that let `A0` through
  // handed every caller a row the grid does not have, and `getCell('A0')` stored that row before the
  // `Cell` constructor refused it, leaving a sheet no later read or write could get past.
  if (row !== undefined) assertRowInBounds(row);
  return {address: `${letters}${digits}`, col, row};
}

/**
 * A reference that names one cell, both axes present. The narrowing of {@link CellAddress} that
 * most callers actually want: `decodeAddress` is deliberately three-shaped because a bare row
 * (`$1`) and a bare column (`$A`) are legitimate references, but a cell is where a value lives, and
 * every caller that needs one was re-deriving that invariant by hand.
 */
export interface CellPosition {
  readonly col: number;
  readonly row: number;
}

/**
 * Decode a reference that must name a single cell. Anchoring `$` signs are accepted and dropped.
 *
 * @throws {SyntaxError} if the reference is unparseable, or parses but omits an axis (`"A"`, `"1"`).
 * @throws {RangeError} if it names a position off the grid (`"A0"`, `"XFE1"`).
 */
export function decodeCellRef(reference: string): CellPosition {
  const {col, row} = decodeAddress(reference);
  if (col === undefined || row === undefined) {
    throw new SyntaxError(
      `${quoted(reference)} is not a single-cell reference: it omits a column or row`,
    );
  }
  return {col, row};
}

// The tolerant half of the decoders, and the read side's only door to them.
//
// `xml-scan.ts` states the rule these honour: a file the library did not write is allowed to be
// wrong, and losing one attribute beats losing the sheet. Every other foreign scalar already has
// its tolerant reader (`numInteger`, `boolTristate`, `enumToken`); a reference had none, so a
// malformed `r`/`ref`/`sqref` in an untrusted package aborted the whole read with a native
// `RangeError` or `SyntaxError`, outside the `XlsxError` taxonomy one `catch` clause answers.
//
// "Tolerant" here means *a reference naming something that can exist*, not merely one that parses.
// Both axes are bounded inside `decodeAddress`, so the throwing decoders already refuse `A0` and
// `XFE1`, and these are nothing more than a `catch` around them: the rule is written once, and the
// two temperaments cannot disagree about where the grid ends.
//
// The authoring decoders keep throwing. `getCell('A0')` from a caller is a mistake at the call, and
// the taxonomy deliberately reserves a native error for a single scalar out of range.

/**
 * {@link decodeCellRef} for a reference that came out of a file rather than out of a caller:
 * `undefined` for anything that does not name one cell that can exist: a range, a bare row or
 * column, a position off the grid (`A0`, `XFE1`), or outright garbage. A foreign producer writes
 * all of them, and none is worth throwing over when the reading code's answer is simply "then
 * there is nothing here".
 */
export function tryDecodeCellRef(reference: string): CellPosition | undefined {
  try {
    return decodeCellRef(reference);
  } catch {
    return undefined;
  }
}

/**
 * Narrow a decoded range to a {@link GridRect}, or `undefined` when either axis is unbounded.
 *
 * A whole-column reference (`A:A`) and a whole-row one (`1:1`) are legal range references whose
 * corners are partly absent, so every caller that needs a real rectangle has to test all four. Doing
 * that in seven places is the other half of the convention `GridRect` already exists to state once:
 * the shape is declared here, and so is the one narrowing that produces it.
 *
 * A caller wanting to say something *different* about each axis (a `Range` refuses `A:A` and `1:1`
 * with two distinct messages) still tests them separately, and one that only cares about the axis it
 * is splicing still tests only that one. Neither is this question.
 */
export function boundedRect(range: RangeAddress): GridRect | undefined {
  const {top, left, bottom, right} = range;
  if (top === undefined || left === undefined || bottom === undefined || right === undefined) {
    return undefined;
  }
  return {top, left, bottom, right};
}

/**
 * {@link decodeRange} for a reference that came out of a file: `undefined` for anything that does
 * not name a region that can exist. The sibling of {@link tryDecodeCellRef} on the other arity: a
 * `ref` or one area of a `sqref` is as likely to be malformed as a cell's `r`, and the reader's
 * answer to both is the same one.
 *
 * An axis neither endpoint mentions stays `undefined`, exactly as in `decodeRange`: a whole-column
 * range is unbounded, not unreadable, and a caller that needs a bounded rectangle says so itself.
 */
export function tryDecodeRange(reference: string): RangeAddress | undefined {
  try {
    return decodeRange(reference);
  } catch {
    return undefined;
  }
}

/**
 * The rectangle `reference` names when its top-left cell is (`col`, `row`), or `undefined` for a
 * reference that does not parse, is unbounded, names a sheet, or starts at another cell. That is the
 * shape of the range an array formula fills: Excel states the formula on the cell the range starts at,
 * so a `ref` read from a file and one handed to the writer are both held to it.
 */
export function tryDecodeAnchoredRange(
  reference: string,
  col: number,
  row: number,
): GridRect | undefined {
  const range = tryDecodeRange(reference);
  if (range === undefined || range.sheetName !== undefined) return undefined;
  const rect = boundedRect(range);
  return rect !== undefined && rect.left === col && rect.top === row ? rect : undefined;
}

/**
 * Decode a range reference (`A1:B2`, `$1:$1`, `Sheet1!$A:$A`) into its corners and
 * canonical dimensions. A single reference collapses to a degenerate range whose
 * corners coincide.
 *
 * @throws {SyntaxError} if an endpoint is unparseable.
 * @throws {RangeError} if an endpoint names a column past `XFD` or a row outside `1..1048576`.
 */
export function decodeRange(reference: string): RangeAddress {
  const prefix = SHEET_PREFIX.exec(reference);
  let sheetName: string | undefined;
  let body = reference;
  if (prefix) {
    const quoted = prefix[1];
    sheetName = quoted !== undefined ? quoted.replace(/''/g, "'") : prefix[2];
    body = prefix[3] ?? '';
  }

  const parts = body.split(':');
  const start = decodeAddress(parts[0] ?? '');
  const end = parts.length > 1 ? decodeAddress(parts[1] ?? '') : start;

  const cols = [start.col, end.col].filter((v): v is number => v !== undefined);
  const rows = [start.row, end.row].filter((v): v is number => v !== undefined);
  const left = cols.length > 0 ? Math.min(...cols) : undefined;
  const right = cols.length > 0 ? Math.max(...cols) : undefined;
  const top = rows.length > 0 ? Math.min(...rows) : undefined;
  const bottom = rows.length > 0 ? Math.max(...rows) : undefined;

  const tl = makeCellAddress(left, top);
  const br = makeCellAddress(right, bottom);

  return {
    top,
    left,
    bottom,
    right,
    ...(sheetName !== undefined ? {sheetName} : {}),
    tl,
    br,
    dimensions: `${tl.address}:${br.address}`,
  };
}

/**
 * Encode a 1-based `col`/`row` pair into its canonical A1 address (`"B2"`).
 *
 * Both axes go through the shared guard. The column already did, through `numberToColumn`; the row
 * checked only its lower bound in a message of its own, so `encodeAddress(1, 1048577)` produced an
 * address naming a row Excel has no reference for while `encodeAddress(16385, 1)` refused.
 */
export function encodeAddress(col: number, row: number): string {
  assertRowInBounds(row);
  return `${numberToColumn(col)}${row}`;
}

// A number an R1C1 reference or an A1 row may carry: digits naming a line the axis has.
function isLineNumber(digits: string, bound: number): boolean {
  const n = Number(digits);
  return n >= 1 && n <= bound;
}

const A1_NAME = /^([A-Za-z]{1,3})(\d+)$/;
const R1C1_NAME = /^(?:R(\d*)C(\d*)|R(\d*)|C(\d*))$/i;
// An R1C1 row or column number followed by a letter or underscore. Digits cannot follow it: the number
// takes them all, and `R1048577` is a row past the grid rather than row 104857 followed by a 7.
const R1C1_PREFIX = /^(?:R(\d+)|C(\d+))[A-Za-z_]/i;

/**
 * Does Excel read this name as a reference, where it expects a name? Such a name must be quoted as a
 * sheet prefix (`'R1C1'!A1`), and cannot name a table at all.
 *
 * Four shapes, case-insensitive, all found by asking Excel (16.0 build 20326) rather than read off a
 * specification, which states the rule as "not the same as a cell reference, and not `C` or `R`" and
 * leaves the edges to the implementation:
 *
 * - an A1 cell **on the grid**: `T1`, `a01`, `XFD1048576`, but not `XFE1` or `A0`;
 * - a whole R1C1 reference, each number optional and on the grid: `R`, `C`, `RC`, `R1C`, `R1C1`;
 * - an R1C1 row or column number followed by a letter or underscore: `R1X`, `C1X`, `R1C1_`, but
 *   not `R1.5`, nor `RCX`, which carries no number;
 * - `TRUE` or `FALSE`.
 *
 * Every name tried both ways was quoted in a formula exactly when a table by that name made Excel
 * offer to repair the package, so the two rules are one.
 */
export function nameReadsAsReference(name: string): boolean {
  const a1 = A1_NAME.exec(name);
  if (a1 !== null) {
    const [, letters = '', digits = ''] = a1;
    if (tryColumnToNumber(letters.toUpperCase()) !== undefined && isLineNumber(digits, MAX_ROW)) {
      return true;
    }
  }
  const r1c1 = R1C1_NAME.exec(name);
  if (r1c1 !== null) {
    const [, row, column, rowOnly, columnOnly] = r1c1;
    const on = (digits: string | undefined, bound: number) =>
      digits === undefined || digits === '' || isLineNumber(digits, bound);
    if (on(row ?? rowOnly, MAX_ROW) && on(column ?? columnOnly, MAX_COLUMN)) return true;
  }
  const prefix = R1C1_PREFIX.exec(name);
  if (prefix !== null) {
    const [, row, column] = prefix;
    if (row !== undefined ? isLineNumber(row, MAX_ROW) : isLineNumber(column ?? '', MAX_COLUMN)) {
      return true;
    }
  }
  return /^(?:TRUE|FALSE)$/i.test(name);
}
