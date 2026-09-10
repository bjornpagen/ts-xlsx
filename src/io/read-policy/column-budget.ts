// The ceiling on how much work one sheet's column records may cost between them.
//
// Every worksheet reader clamps a single column run to `MAX_COLUMN`, which bounds one record's loop
// and nothing else: the *number* of runs is unbounded, and each may span the whole grid, so the cost
// is the product. 154 KB of worksheet XML holding 2,000 full-grid `<col>` spans cost the buffered
// reader 12.3 s; extrapolated against the 512 MiB inflate ceiling that is hours of CPU from a package
// compressing to a few hundred KB. The zip-bomb guard cannot see it, because after inflation the
// payload really is small: this is a CPU bomb, not an allocation one.
//
// It binds `BrtColInfo` exactly as it binds `<col>`, which is why it sits below both codecs. While it
// lived in the XML codec the binary reader had no budget at all, and 800 full-width hidden runs, 16 KB
// of part, took it most of a second, linear in how many runs a file cares to write.

import {MAX_COLUMN} from '../../core/address.ts';

// Four times the whole grid. Excel writes disjoint spans covering at most `MAX_COLUMN` columns in
// total, so no legitimate file comes near this, while the worst case stays a fraction of a second.
const MAX_COLUMN_RECORD_TOUCHES = MAX_COLUMN * 4;

/**
 * A per-sheet allowance of column touches, spent by each column run the reader applies.
 *
 * A run past the allowance is truncated, and one starting past it is dropped, rather than the read
 * being refused: that is the stance every reader already takes on an out-of-grid run, and it means a
 * hostile file loses formatting it could not have meant while a real one is untouched.
 */
export class ColumnRecordBudget {
  #remaining = MAX_COLUMN_RECORD_TOUCHES;

  /** The last column of `[first, last]` this sheet can still afford, or `undefined` when none is. */
  take(first: number, last: number): number | undefined {
    if (this.#remaining <= 0 || first > last) return undefined;
    const affordable = Math.min(last, first + this.#remaining - 1);
    this.#remaining -= affordable - first + 1;
    return affordable;
  }
}

/**
 * The columns of a declared run `[first, last]`, one-based, that a sheet actually applies, or
 * `undefined` for a run that applies to none: one starting past the grid, an empty one, or one the
 * budget can no longer afford.
 *
 * Beside the budget because the two decisions are one decision. The XML readers each had their own
 * copy, and the copies had already drifted: one tested `first > MAX_COLUMN` and the other did not,
 * which happened not to matter only because `budget.take` returns `undefined` when `first > last`, an
 * accident of that method's contract rather than an agreement between the readers.
 */
export function clampColumnSpan(
  first: number,
  last: number,
  budget: ColumnRecordBudget,
): {first: number; last: number} | undefined {
  if (first > MAX_COLUMN) return undefined;
  // Clamped to the format's ceiling rather than refused: `<col max="99999999">` is a file Excel
  // opens, and an unclamped loop would materialise 16.7 million column records before dying.
  const affordable = budget.take(first, Math.min(last, MAX_COLUMN));
  return affordable === undefined ? undefined : {first, last: affordable};
}
