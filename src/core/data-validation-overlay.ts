// The data-validation overlay a Worksheet owns: an insertion-ordered list of range-bound rules, plus
// the decoded rectangles a point-in-rect lookup ({@link DataValidationOverlay.at}) tests against. Kept
// as its own class, the sibling to {@link GridEdits} that owns splice arithmetic, rather than inline
// on Worksheet, since a validation's storage (a rule plus its decoded ranges) is a self-contained unit
// Worksheet only ever adds to, reads, or clears wholesale.

import {AuthoringError, quoted} from '../errors.ts';
import {replaceContents} from './containers.ts';
import {
  cloneDataValidation,
  type DataValidation,
  type DataValidationEntry,
} from './data-validation.ts';
import type {AxisSplice} from './grid-shift.ts';
import {decodeSqrefRects, type MergeRect, shiftSqref} from './merge.ts';

export class DataValidationOverlay {
  readonly #entries: DataValidationEntry[] = [];
  readonly #rects: {rects: readonly MergeRect[]; rule: DataValidation}[] = [];

  /**
   * Attach a rule to a target range (`"B2:B20"`, a whole column `"B2:B1048576"`, or a space-separated
   * `sqref` of several ranges). The rule is stored once against the range, not copied per covered
   * cell, so a whole-column dropdown stays a single entry. A cell inside the range reports the rule
   * through {@link at}.
   *
   * Pass `{extended: true}` to mark a rule that belongs in the 2009 extension form
   * (`<x14:dataValidation>`), the carrier Excel uses for a list source on another sheet and other
   * shapes the standard element cannot express. The reader sets it for a rule found in that form so a
   * round-trip writes it back there instead of silently corrupting the cross-sheet reference.
   *
   * @throws {AuthoringError} when `sqref` names no area at all. A rule attached to nothing covers no
   * cell and is written back as the same unreadable text, so it is a mistake worth surfacing at the
   * call. The reader does not reach this: it drops such an entry at its own boundary, where a foreign
   * file's malformed attribute is supposed to be dropped.
   */
  add(sqref: string, rule: DataValidation, options: {extended?: boolean} = {}): void {
    const rects = decodeSqrefRects(sqref);
    if (rects.length === 0) {
      throw new AuthoringError(`data validation range ${quoted(sqref)} names no cells`);
    }
    // One defensive copy, shared by the serialisable entry and the lookup index, so the getter never
    // hands back a reference into the caller's object.
    const stored = cloneDataValidation(rule);
    const entry: DataValidationEntry = {sqref, rule: stored};
    if (options.extended) entry.extended = true;
    this.#entries.push(entry);
    this.#rects.push({rects, rule: stored});
  }

  /** The data validations on this sheet, each bound to its target range, in insertion order. */
  get entries(): readonly DataValidationEntry[] {
    return this.#entries;
  }

  /**
   * The validation covering the 1-based `col`/`row`, or `undefined` when none does. The first added
   * rule whose range contains the cell wins, mirroring how a spreadsheet resolves overlapping
   * validations.
   */
  // A linear scan over every rule and its decoded rectangles, and it stays one. Measured on a sheet
  // of whole-column dropdowns read cell by cell: a hit costs ~0.3 us whatever the rule count, since
  // the first match returns, and the worst case, a miss that has to test all 400 rectangles of 200
  // two-area rules, costs ~3 us. A caller reading 100k cells against that sheet spends 0.3 s, which
  // is not where a spreadsheet library's time goes. An index would have to be an interval tree over
  // the rectangles, never a different container for the entries: their insertion order is what makes
  // "first added rule wins" true, so it is load-bearing rather than incidental.
  at(col: number, row: number): DataValidation | undefined {
    for (const {rects, rule} of this.#rects) {
      for (const rect of rects) {
        if (col >= rect.left && col <= rect.right && row >= rect.top && row <= rect.bottom) {
          return rule;
        }
      }
    }
    return undefined;
  }

  /**
   * Re-anchor every rule through a row or column splice, so a dropdown stays on the cells it was
   * attached to rather than on whatever moved into their place. A rule whose every target area fell
   * inside a deleted span is dropped with them.
   */
  shift(splice: AxisSplice): void {
    const entries: DataValidationEntry[] = [];
    const rects: {rects: readonly MergeRect[]; rule: DataValidation}[] = [];
    for (const entry of this.#entries) {
      const sqref = shiftSqref(entry.sqref, splice);
      if (sqref === undefined) continue;
      entries.push({...entry, sqref});
      // The decoded rectangles are what `at()` answers from, so they are re-derived here rather than
      // shifted alongside: a stale index would report the pre-splice geometry.
      rects.push({rects: decodeSqrefRects(sqref), rule: entry.rule});
    }
    replaceContents(this.#entries, entries);
    replaceContents(this.#rects, rects);
  }

  /**
   * Rewrite every rule's formula operands through `rewrite`, keeping a rule whose operands did not
   * change as the same object. A numeric operand is not a formula and passes through.
   */
  mapFormulas(rewrite: (formula: string) => string): void {
    for (const [index, entry] of this.#entries.entries()) {
      const formulae = entry.rule.formulae;
      if (formulae === undefined) continue;
      const next = formulae.map((operand) =>
        typeof operand === 'string' ? rewrite(operand) : operand,
      );
      if (next.every((operand, i) => operand === formulae[i])) continue;
      const rule: DataValidation = {...entry.rule, formulae: next};
      this.#entries[index] = {...entry, rule};
      const indexed = this.#rects[index];
      if (indexed !== undefined) this.#rects[index] = {rects: indexed.rects, rule};
    }
  }

  /** Drop every validation, leaving the overlay empty. */
  clear(): void {
    this.#entries.length = 0;
    this.#rects.length = 0;
  }
}
