// Which array formulas a sheet cannot hold as they stand: the range an array formula fills may hold no
// other cell's formula, and may share no cell with another array formula's range.
//
// Excel 16.0 (build 20326) offers to repair every package breaking either rule, for a legacy
// Ctrl+Shift+Enter formula and a dynamic-array one alike: another formula in a member cell, whether plain
// or an array of its own, and two ranges sharing a cell, whether or not either range holds the other's
// own cell (`test/corpus/fixtures/excel-oracle/array-formula-ranges.json`). A member cell holding a value
// is how Excel stores what the formula produced, and breaks neither rule.
//
// The writer refuses such a sheet, naming the first conflict (`arrayRangeConflicts`). A reader repairs
// one as Excel repairs it (`arrayRangeRepair`): Excel keeps the array formula and takes the other.

import {type GridRect, tryDecodeAnchoredRange} from './address.ts';
import {MergeIndex} from './merge-index.ts';
import {
  type CellValue,
  isArrayFormulaValue,
  isDataTableFormulaValue,
  isFormulaValue,
  isSharedFormulaValue,
} from './value.ts';

/**
 * Where a cell holding `value` places a formula, or `undefined` when the value is no formula. A shared
 * formula's clone and a data table's cell count: each is a formula Excel will not find inside another
 * array formula's range.
 */
export function formulaPlacement(
  address: string,
  col: number,
  row: number,
  value: CellValue,
): FormulaPlacement | undefined {
  if (isArrayFormulaValue(value)) return {address, col, row, arrayRef: value.ref};
  if (isFormulaValue(value) || isSharedFormulaValue(value) || isDataTableFormulaValue(value)) {
    return {address, col, row};
  }
  return undefined;
}

/** A cell holding a formula of any kind, with the range it fills when it is an array formula. */
export interface FormulaPlacement {
  readonly address: string;
  readonly col: number;
  readonly row: number;
  /** The array formula's `ref`, as its value states it; absent for any other formula. */
  readonly arrayRef?: string | undefined;
}

/** An array formula whose range holds another formula's cell, or shares a cell with a range kept. */
export interface ArrayRangeConflict {
  readonly array: FormulaPlacement;
  /** The address of the other formula in the range, or of the array formula whose range it overlaps. */
  readonly other: string;
  readonly kind: 'formula-inside' | 'ranges-overlap';
}

/**
 * The array formulas that cannot stand, deciding in reading order: an array formula conflicts when
 * another formula's cell lies in its range, or when its range shares a cell with the range of an
 * array formula decided before it that did not conflict. The array formulas left over hold ranges
 * that share no cell with each other and hold no formula but their own, so reading each conflict as
 * its plain formula leaves a sheet the writer accepts.
 *
 * A `ref` that names no bounded range starting at its own cell fills no cells here; the reader never
 * produces one and the writer refuses it on its own terms.
 */
export function arrayRangeConflicts(placements: Iterable<FormulaPlacement>): ArrayRangeConflict[] {
  const all = [...placements];
  // Every formula's cell as a one-cell region, so a range asks the banded index which of them it holds.
  const cellRects: GridRect[] = [];
  const placementOfCell = new Map<GridRect, FormulaPlacement>();
  const arrays: {readonly placement: FormulaPlacement; readonly range: GridRect}[] = [];
  for (const placement of all) {
    const cell = {
      top: placement.row,
      left: placement.col,
      bottom: placement.row,
      right: placement.col,
    };
    cellRects.push(cell);
    placementOfCell.set(cell, placement);
    const range =
      placement.arrayRef === undefined
        ? undefined
        : tryDecodeAnchoredRange(placement.arrayRef, placement.col, placement.row);
    if (range !== undefined) arrays.push({placement, range});
  }
  arrays.sort((a, b) => a.placement.row - b.placement.row || a.placement.col - b.placement.col);

  const cells = new MergeIndex(cellRects);
  const keptRanges: GridRect[] = [];
  const kept = new MergeIndex(keptRanges);
  const placementOfRange = new Map<GridRect, FormulaPlacement>();
  const conflicts: ArrayRangeConflict[] = [];
  for (const {placement, range} of arrays) {
    const inside = formulaInside(range, cells);
    if (inside !== undefined) {
      const other = placementOfCell.get(inside)?.address ?? '';
      conflicts.push({array: placement, other, kind: 'formula-inside'});
      continue;
    }
    const overlapped = kept.overlapping(range);
    if (overlapped !== undefined) {
      const other = placementOfRange.get(overlapped)?.address ?? '';
      conflicts.push({array: placement, other, kind: 'ranges-overlap'});
      continue;
    }
    keptRanges.push(range);
    placementOfRange.set(range, placement);
    kept.note(range);
  }
  return conflicts;
}

/** What Excel's repair takes from a sheet whose array formulas cannot stand. */
export interface ArrayRangeRepair {
  /** Formulas whose cell lies in the range of an array formula kept: each loses its whole cell. */
  readonly cellsRemoved: FormulaPlacement[];
  /**
   * Array formulas whose own cell lies in no range kept but whose range shares a cell with one: each loses
   * its formula and keeps the value it cached.
   */
  readonly formulasRemoved: FormulaPlacement[];
}

/**
 * What Excel's repair takes from the sheet, deciding in reading order: an array formula keeps its range
 * unless that range shares a cell with a range kept before it. A formula of any kind whose cell lies in a
 * kept range loses the cell; an array formula whose range shares a cell with a kept range, its own cell
 * outside it, loses its formula and keeps its value. What is left holds no conflict.
 *
 * Excel 16.0 (build 20326) repaired B1:B3 holding a plain formula, or an array formula of its own, in B2
 * by keeping B1's range and removing B2's "cell information", and B1:C2 beside A2:B3 by keeping B1's range
 * and removing A2's "formula", its value left (`test/corpus/fixtures/excel-oracle/array-formula-ranges.json`).
 * An array formula's own cell is the first of its range in reading order, so a range is always decided
 * before any cell it holds, and a range whose formula was taken claims none.
 */
export function arrayRangeRepair(placements: Iterable<FormulaPlacement>): ArrayRangeRepair {
  const ordered = [...placements].sort((a, b) => a.row - b.row || a.col - b.col);
  const keptRanges: GridRect[] = [];
  const kept = new MergeIndex(keptRanges);
  const cellsRemoved: FormulaPlacement[] = [];
  const formulasRemoved: FormulaPlacement[] = [];
  for (const placement of ordered) {
    const {col, row} = placement;
    if (kept.overlapping({top: row, left: col, bottom: row, right: col}) !== undefined) {
      cellsRemoved.push(placement);
      continue;
    }
    const range =
      placement.arrayRef === undefined
        ? undefined
        : tryDecodeAnchoredRange(placement.arrayRef, col, row);
    if (range === undefined) continue;
    if (kept.overlapping(range) !== undefined) {
      formulasRemoved.push(placement);
      continue;
    }
    keptRanges.push(range);
    kept.note(range);
  }
  return {cellsRemoved, formulasRemoved};
}

// A formula's cell in `range` other than the range's own top-left cell, which is the array formula
// itself. The range less that cell is the rows below it and the rest of its own row, so those two
// rectangles are asked; either may be empty.
function formulaInside(range: GridRect, cells: MergeIndex): GridRect | undefined {
  const below: GridRect = {...range, top: range.top + 1};
  const beside: GridRect = {...range, bottom: range.top, left: range.left + 1};
  return (
    (below.top <= below.bottom ? cells.overlapping(below) : undefined) ??
    (beside.left <= beside.right ? cells.overlapping(beside) : undefined)
  );
}
