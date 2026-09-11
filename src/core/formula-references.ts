// The references a formula's text makes, moved: to where a shared formula's clone sits, or through a
// row or column splice.
//
// A spreadsheet fills a formula down or across a range by storing it once on a master cell and marking
// the rest as shared clones. Reading a clone means recovering the master's formula shifted to the
// clone's position: relative references move by the row and column offset, absolute (`$`-anchored)
// parts stay put. That is `translateFormula`.
//
// A splice already moves every coordinate the model stores as data: cells, merges, tables, anchors,
// overlays. Formula text is the one place a coordinate is spelled rather than stored, so it needs a
// reader of its own: `SUM(A1:A10)` above an inserted row has to become `SUM(A1:A11)` or it sums cells
// the author never chose, and nothing about the file says so.
//
// The rules are Excel's, read back through `Range.Formula` after the same edits made over COM (Excel
// 16.0 build 20326, `test/corpus/fixtures/excel-oracle/formula-references-through-a-splice.json`):
//
// - An insert moves every line at or after it, absolute (`$A$5`) or relative alike: a splice is a
//   change to the grid, not a copy of the formula. A range grows when the insert lands inside it.
// - A delete turns a reference it takes into `#REF!`, keeping a sheet prefix in front of it
//   (`'S1'!#REF!`), and shrinks a range it cuts into. A reference pushed past the grid's last line is
//   `#REF!` too, while a range edge already on that line stays there.
// - Whole columns do not move with rows, nor whole rows with columns.
// - Only a reference to the spliced sheet moves: an unqualified one in a formula on that sheet, or one
//   qualified with its name anywhere. A 3-D span (`S1:S3!A5`), an external reference (`[1]S1!A5`), a
//   string, and a structured reference are left as they are.

import {MAX_COLUMN, MAX_ROW, numberToColumn, tryColumnToNumber} from './address.ts';
import {scanFormula, skipOpaque} from './formula-scan.ts';
import type {AxisSplice} from './grid-shift.ts';
import {REF_ERROR} from './value.ts';

// The three reference shapes a shift moves, tried at each position in this order: a whole-column
// range (`A:C`), a whole-row range (`1:5`), then a single cell (`A1`). Each axis takes an optional `$`.
// Column letters are uppercase-only because Excel stores them that way, so a lowercase defined name is
// never mistaken for one, and row digits stop at seven (Excel's last row is 1048576).
//
// The lookbehind rejects a reference glued to a preceding name character or '.', so the `A1` inside
// `_xlfn.A1` or a defined name `FOO_A1` is left alone. The lookahead rejects one continued by a name
// character, opening a call `(`, or preceding a sheet `!`: a token before `!` is the sheet name
// (`Q1!A1`), not a cell. A sheet-qualified reference still shifts, because the `!` before it is not a
// name character. Applied per code run, where opaque regions such as a `"10:30"` literal are gone.
//
// Only the token directly before `!` is a sheet. `Q1:Q4!B2` is not a 3-D span over two sheets named
// like cells, which would have to be quoted as a whole (`'Q1:Q4'!B2`): Excel reads it as the range from
// cell Q1 to `Q4!B2`, so Q1 shifts here too.
//
// One pattern rather than a pass per shape, so a shifted result is never scanned again, and `A1:B2`
// still resolves as two cells: a range alternative needs a bare letter run or a bare digit run on both
// sides of its colon. Cells used to be the only shape, so `A:A` and `1:1` never moved.
const SHIFTABLE_REFERENCE =
  /(?<![A-Za-z0-9_.])(?:(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})|(\$?)([0-9]{1,7}):(\$?)([0-9]{1,7})|(\$?)([A-Z]{1,3})(\$?)([0-9]{1,7}))(?![A-Za-z0-9_.!(])/g;

/**
 * Shift every relative reference in a formula by `colDelta` columns and `rowDelta` rows, leaving
 * absolute (`$`-anchored) axes fixed. This is how a shared-formula clone recovers its own formula from
 * the master's: a master `A1*2` shared one row down reads back as `A2*2`, `$A$1*B1` shared one row and
 * one column across as `$A$1*C2`, and `COUNTIF(A:A,A1)` shared one column right as `COUNTIF(B:B,B1)`.
 * Whole-column and whole-row ranges shift at each relative end, as cells do. String literals,
 * single-quoted sheet names, and bracketed structured references are copied verbatim, and a
 * sheet-qualified reference shifts while its sheet name is untouched. Function names and defined names
 * are not reference-shaped, so they pass through.
 */
export function translateFormula(formula: string, colDelta: number, rowDelta: number): string {
  if (colDelta === 0 && rowDelta === 0) return formula;
  // A reference the grid cannot hold becomes `#REF!` on either axis, which is what Excel writes for the
  // same shift; the decode is tolerant so that answer is reachable at all. The deltas come from a
  // file's own shared-formula geometry, so this is a read path, and an odd file aborting the whole
  // sheet with an error outside the library's taxonomy is not an answer. Both axes used to do exactly
  // that, in opposite ways. The column went through `columnToNumber`, which threw a bare `RangeError`
  // before any guard could speak: the pattern matches three letters, so `ZZZ1` (column 18278) reached
  // it, and a reference past XFD failed where a shift past XFD resolved. The row axis just did the
  // arithmetic and emitted `A0` or `A-4`, which is not a reference at all.
  const column = (anchor: string, letters: string): string | undefined => {
    const decoded = tryColumnToNumber(letters);
    if (decoded === undefined) return undefined;
    const col = anchor === '$' ? decoded : decoded + colDelta;
    return col < 1 || col > MAX_COLUMN ? undefined : `${anchor}${numberToColumn(col)}`;
  };
  const row = (anchor: string, digits: string): string | undefined => {
    const index = anchor === '$' ? Number(digits) : Number(digits) + rowDelta;
    return index < 1 || index > MAX_ROW ? undefined : `${anchor}${index}`;
  };
  const range = (first: string | undefined, last: string | undefined): string =>
    first === undefined || last === undefined ? REF_ERROR : `${first}:${last}`;

  return scanFormula(formula, (code) =>
    code.replace(
      SHIFTABLE_REFERENCE,
      (
        _match,
        firstColumnAnchor: string | undefined,
        firstColumn: string | undefined,
        lastColumnAnchor: string | undefined,
        lastColumn: string | undefined,
        firstRowAnchor: string | undefined,
        firstRow: string | undefined,
        lastRowAnchor: string | undefined,
        lastRow: string | undefined,
        cellColumnAnchor: string | undefined,
        cellColumn: string | undefined,
        cellRowAnchor: string | undefined,
        cellRow: string | undefined,
      ) => {
        // An anchor group is `''`, never `undefined`, whenever its alternative matched, so the
        // `?? ''` only ever meets an alternative that did not.
        if (firstColumn !== undefined && lastColumn !== undefined) {
          return range(
            column(firstColumnAnchor ?? '', firstColumn),
            column(lastColumnAnchor ?? '', lastColumn),
          );
        }
        if (firstRow !== undefined && lastRow !== undefined) {
          return range(row(firstRowAnchor ?? '', firstRow), row(lastRowAnchor ?? '', lastRow));
        }
        const shiftedColumn = column(cellColumnAnchor ?? '', cellColumn ?? '');
        const shiftedRow = row(cellRowAnchor ?? '', cellRow ?? '');
        return shiftedColumn === undefined || shiftedRow === undefined
          ? REF_ERROR
          : `${shiftedColumn}${shiftedRow}`;
      },
    ),
  );
}

/** A splice as a formula sees it: which sheet's lines moved, and how. */
export interface SheetSplice {
  /** The spliced sheet's name, matched in any case. */
  readonly sheet: string;
  readonly splice: AxisSplice;
}

// A reference at a token start, in the four shapes a splice moves: a cell range, a cell, a column
// range, a row range. Column letters are uppercase only, as Excel stores them, so a lowercase name is
// never read as one. What follows may not continue a name, open a call, or name a sheet.
const REFERENCE =
  /(?:(\$?)([A-Z]{1,3})(\$?)(\d{1,7}):(\$?)([A-Z]{1,3})(\$?)(\d{1,7})|(\$?)([A-Z]{1,3})(\$?)(\d{1,7})|(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})|(\$?)(\d{1,7}):(\$?)(\d{1,7}))(?![\p{L}\p{N}_.(![])/uy;

const NAME_CHAR = /[\p{L}\p{N}_.\\]/u;

/**
 * Rewrite a formula's references for a splice of `edit.sheet`. `home` is the sheet the formula belongs
 * to, whose name an unqualified reference means; `undefined` for a formula that has none, such as a
 * workbook-scoped defined name, whose references to a sheet are always qualified.
 *
 * Returns the formula unchanged, as the same string, when nothing in it moved.
 */
export function spliceFormula(
  formula: string,
  home: string | undefined,
  edit: SheetSplice,
): string {
  const target = edit.sheet.toLowerCase();
  const homeIsTarget = home !== undefined && home.toLowerCase() === target;
  let out = '';
  let i = 0;
  const n = formula.length;
  // Whether the reference at `i` may be one: not glued to a name or number before it.
  const atTokenStart = (index: number): boolean =>
    index === 0 || !NAME_CHAR.test(formula[index - 1] ?? '');

  // Copy the reference at `i`, rewritten when `moves` says it names the spliced sheet. Returns whether
  // one was there.
  const reference = (moves: boolean): boolean => {
    REFERENCE.lastIndex = i;
    const match = REFERENCE.exec(formula);
    if (match === null) return false;
    out += moves ? moved(match, edit.splice) : match[0];
    i = REFERENCE.lastIndex;
    return true;
  };

  while (i < n) {
    const ch = formula[i] ?? '';
    if (ch === '"') {
      const end = skipOpaque(formula, i);
      out += formula.slice(i, end);
      i = end;
      continue;
    }
    if (ch === '[') {
      // A structured reference, or an external workbook's index in front of a bare sheet name: either
      // way nothing up to the reference that follows is this workbook's grid.
      const end = skipOpaque(formula, i);
      out += formula.slice(i, end);
      i = end;
      const name = nameRun(formula, i);
      if (name > i && formula[name] === '!') {
        out += formula.slice(i, name + 1);
        i = name + 1;
        reference(false);
      }
      continue;
    }
    if (ch === "'") {
      const end = skipOpaque(formula, i);
      const quoted = formula.slice(i + 1, end - 1).replaceAll("''", "'");
      out += formula.slice(i, end);
      i = end;
      if (formula[i] !== '!') continue;
      out += '!';
      i += 1;
      // A quoted prefix spans sheets (`'S1:S3'`) or names another workbook (`'[1]S1'`) by what it
      // holds.
      reference(!quoted.includes(':') && !quoted.includes('[') && quoted.toLowerCase() === target);
      continue;
    }
    if (atTokenStart(i) && NAME_CHAR.test(ch)) {
      const end = nameRun(formula, i);
      const next = formula[end];
      if (next === '!') {
        out += formula.slice(i, end + 1);
        const name = formula.slice(i, end);
        i = end + 1;
        reference(name.toLowerCase() === target);
        continue;
      }
      if (next === ':') {
        // `S1:S3!A5` is a 3-D span, whose sheets a splice of one of them does not move.
        const last = nameRun(formula, end + 1);
        if (last > end + 1 && formula[last] === '!') {
          out += formula.slice(i, last + 1);
          i = last + 1;
          reference(false);
          continue;
        }
      }
      if (reference(homeIsTarget)) continue;
      out += formula.slice(i, end);
      i = end;
      continue;
    }
    if (atTokenStart(i) && ch === '$' && reference(homeIsTarget)) continue;
    out += ch;
    i += 1;
  }
  return out === formula ? formula : out;
}

function nameRun(formula: string, start: number): number {
  let j = start;
  while (j < formula.length && NAME_CHAR.test(formula[j] ?? '')) j += 1;
  return j;
}

// A reference matched by REFERENCE, moved: its text with the moved lines spelled in, anchors kept, or
// `#REF!` when the splice took it or pushed it off the grid.
function moved(match: RegExpExecArray, splice: AxisSplice): string {
  const g = (index: number): string => match[index] ?? '';
  // Each end as its column anchor, letters, row anchor and digits. A whole-line range has no pair for
  // the other axis, and a single cell is a range whose ends are one.
  const ends =
    match[2] !== undefined
      ? [
          [g(1), g(2), g(3), g(4)],
          [g(5), g(6), g(7), g(8)],
        ]
      : match[10] !== undefined
        ? [[g(9), g(10), g(11), g(12)]]
        : match[14] !== undefined
          ? [
              [g(13), g(14), '', ''],
              [g(15), g(16), '', ''],
            ]
          : [
              ['', '', g(17), g(18)],
              ['', '', g(19), g(20)],
            ];
  const rows = splice.axis === 'row';
  const lines = ends.map(([, letters = '', , digits = '']) => (rows ? digits : letters));
  // Whole columns do not move with rows, nor whole rows with columns.
  if (lines.includes('')) return match[0];
  const [first = 0, last = first] = lines.map((line) => (rows ? Number(line) : columnOf(line)));
  const edges = movedEdges(
    Math.min(first, last),
    Math.max(first, last),
    splice,
    rows ? MAX_ROW : MAX_COLUMN,
  );
  if (edges === undefined) return REF_ERROR;
  const placed = first <= last ? edges : [edges[1], edges[0]];
  return ends
    .map(([columnAnchor = '', letters = '', rowAnchor = '', digits = ''], i) => {
      const at = placed[i] ?? 0;
      return rows
        ? `${columnAnchor}${letters}${rowAnchor}${at}`
        : `${columnAnchor}${numberToColumn(at)}${rowAnchor}${digits}`;
    })
    .join(':');
}

// Column letters as a number, or one past the grid for letters naming no column, so a reference to a
// column that cannot exist is never moved onto one that can.
function columnOf(letters: string): number {
  return tryColumnToNumber(letters) ?? MAX_COLUMN + 1;
}

// Where a range's two edges land, or `undefined` when the delete took every line between them. An edge
// the delete took moves to the cut: the bottom up to the last line before it, the top down to the first
// line after it, which for a splice that also inserts is the first line after the inserted block, since
// Excel's delete and insert are two edits and the insert moves a top already on the cut. A bottom pushed
// past the grid stops on its last line; a top pushed past it leaves nothing, which is how a single cell
// pushed off the grid, a range whose edges are one line, becomes `#REF!`.
function movedEdges(
  lo: number,
  hi: number,
  splice: AxisSplice,
  bound: number,
): [number, number] | undefined {
  if (hi > bound) return [lo, hi];
  const {start, count, delta} = splice;
  const end = start + count;
  if (lo >= start && hi < end) return undefined;
  const top = lo < start ? lo : lo < end ? end + delta : lo + delta;
  const bottom = hi < start ? hi : hi < end ? start - 1 : hi + delta;
  if (top > bound) return undefined;
  return [top, Math.min(bottom, bound)];
}
