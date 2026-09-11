// The conditional-formatting overlay a Worksheet owns: an insertion-ordered, defensively-copied list
// of range-bound rule sets. Kept as its own class, the sibling to {@link DataValidationOverlay}, so
// Worksheet delegates the collection's storage and cloning rather than managing the array itself.

import {AuthoringError, quoted} from '../errors.ts';
import {type ConditionalFormatting, cloneConditionalFormatting} from './conditional-formatting.ts';
import {replaceContents} from './containers.ts';
import type {AxisSplice} from './grid-shift.ts';
import {decodeSqrefRects, shiftSqref} from './merge.ts';

export class ConditionalFormattingOverlay {
  readonly #entries: ConditionalFormatting[] = [];

  /**
   * Attach a conditional formatting to a target range. `formatting.ref` is an OOXML `sqref`: one
   * range (`"A1:A10"`), a whole column, or several space-separated areas (`"A1:C1 A3:C3"`) sharing one
   * rule set. The block is stored once against the range, defensively copied so the getter never hands
   * back a reference into the caller's object.
   *
   * @throws {AuthoringError} when `formatting.ref` names no area at all. A rule set attached to
   * nothing formats no cell and is written back as the same unreadable text, so it is a mistake worth
   * surfacing at the call. The reader does not reach this: it drops such a block at its own boundary.
   */
  add(formatting: ConditionalFormatting): void {
    if (decodeSqrefRects(formatting.ref).length === 0) {
      throw new AuthoringError(
        `conditional formatting range ${quoted(formatting.ref)} names no cells`,
      );
    }
    this.#entries.push(cloneConditionalFormatting(formatting));
  }

  /** The conditional formattings on this sheet, each bound to its target range, in insertion order. */
  get entries(): readonly ConditionalFormatting[] {
    return this.#entries;
  }

  /**
   * Re-anchor every rule set through a row or column splice, so a highlight keeps covering the cells
   * it was written for. A rule set whose every target area fell inside a deleted span goes with them.
   */
  shift(splice: AxisSplice): void {
    const entries: ConditionalFormatting[] = [];
    for (const entry of this.#entries) {
      const ref = shiftSqref(entry.ref, splice);
      if (ref !== undefined) entries.push({...entry, ref});
    }
    replaceContents(this.#entries, entries);
  }

  /**
   * Rewrite every rule's formulae, and the value of each scale anchor whose type is `formula`, through
   * `rewrite`, keeping a rule set nothing in which changed as the same object.
   */
  mapFormulas(rewrite: (formula: string) => string): void {
    for (const [index, entry] of this.#entries.entries()) {
      let changed = false;
      const rules = entry.rules.map((rule) => {
        const formulae = rule.formulae?.map((operand) =>
          typeof operand === 'string' ? rewrite(operand) : operand,
        );
        const cfvo = rule.cfvo?.map((anchor) =>
          anchor.type === 'formula' && typeof anchor.value === 'string'
            ? {...anchor, value: rewrite(anchor.value)}
            : anchor,
        );
        const formulaeMoved =
          formulae?.some((operand, i) => operand !== rule.formulae?.[i]) ?? false;
        const cfvoMoved =
          cfvo?.some((anchor, i) => anchor.value !== rule.cfvo?.[i]?.value) ?? false;
        if (!formulaeMoved && !cfvoMoved) return rule;
        changed = true;
        return {
          ...rule,
          ...(formulae === undefined ? {} : {formulae}),
          ...(cfvo === undefined ? {} : {cfvo}),
        };
      });
      if (changed) this.#entries[index] = {...entry, rules};
    }
  }

  /** Drop every conditional formatting, leaving the overlay empty. */
  clear(): void {
    this.#entries.length = 0;
  }
}
