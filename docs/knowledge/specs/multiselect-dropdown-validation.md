# Multi-select dropdown validation: a request the format cannot natively satisfy

Cluster: data-validation

## Scenario

A user wants a cell where the reader can pick **several** values from a predefined list and have them
accumulate in the cell, say comma-separated, rather than a normal list dropdown that replaces the
cell content with a single chosen value. They ask the library to emit such a "multi-select dropdown".

> Spec note, not a corpus case: there is nothing to assert implementation-blind, because a
> multi-select dropdown is **not a native spreadsheet-format feature**. Standard list data validation
> is single-select, and the accumulate-on-pick behavior in real spreadsheets is implemented with an
> event-handler macro (VBA `Worksheet_Change`) in a macro-enabled workbook, not with any writable
> validation XML. Recording the constraint and the portable primitive we _should_ nail is the durable
> value.

## Desired behavior

- **Set the expectation honestly.** The library cannot emit a self-contained, macro-free
  multi-select dropdown that works on open in Excel, because the format has no such construct. Public
  docs and types should make that clear rather than implying a `multiSelect: true` flag could exist.
- **Nail the portable primitive instead.** First-class, well-typed **single-select list validation**
  is the thing every consumer of this request actually shares underneath. Inline literal lists,
  cross-sheet range sources, and defined-name sources must all round-trip cleanly, already covered by
  corpus cases. A robust single-select API is the honest 90% of the ask.
- **If a multi-select affordance is ever offered**, it must be an explicitly opt-in, clearly-labelled
  macro-emitting path (a macro-enabled `.xlsm` with an injected change handler), never silent, never
  in a plain `.xlsx`, and never presented as a native validation type.

## Open questions

- Is emitting a macro-enabled workbook with an injected change handler ever in scope, or permanently
  out? It collides with the no-unsafe-eval and minimal-surface stances and with macro-security
  concerns.
- Would a **helper that pre-populates a helper column plus a single-select list** cover the common
  intent, pick-from-list with one value per row, without any macro?
- Documentation placement: a "format limitations" section so this recurring request has a canonical
  answer.

Related: `list-validation-inline-formula-length-limit`, `no-unsafe-eval-csp-compatible`,
`sheet-protection-password-hash-compatibility` (macro and security posture),
`cross-sheet-list-validation-x14`.
