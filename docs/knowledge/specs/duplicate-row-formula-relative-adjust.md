# Duplicating a row carries and translates formulas

Cluster: formulas

## Scenario

A user duplicates a worksheet row containing formula cells, expecting spreadsheet-app behavior:
the copied cells keep their formulas, and relative references shift by the row offset while
absolute (`$`-anchored) references stay fixed. A row where a cell holds `=A1+1`, duplicated one
row down, should produce `=A2+1` in the copy. Today the duplicate operation carries over values
and styles only, so formulas are lost, replaced by static computed values or blanks, and no
reference translation happens. Users want the same for column duplication, and expect formulas
elsewhere that reference the shifted region to update when rows are inserted.

## Desired behavior

A row-duplication operation should, in addition to copying values and styles, copy each source
cell's formula into the corresponding destination cell, translating relative references by the
`(destRow − srcRow)` delta and leaving absolute (`$column` and `$row`) components unchanged.
Shared-formula groups and array formulas in the source should degrade to well-defined per-cell
formulas in the copy rather than being dropped. A parallel column-duplication operation applies
the same rule, translating by the column delta.

Prior art: desktop spreadsheet applications perform relative and absolute reference translation on
copy-and-insert; the A1-relative addressing model with `$` anchors is the OOXML standard. The
fork already models per-cell formula text and shared formulas.

## Open questions

- **Scope of reference updates.** Should duplicating or inserting rows also rewrite formulas
  _elsewhere_ in the sheet, and on other sheets, that reference cells at or below the insertion
  point, the way a live spreadsheet would? That is a much larger reference-maintenance feature
  than local translation of the copied cells and should probably be a separate, opt-in
  capability with its own spec.
- **Edge overflow.** Translated references that would run off the sheet edge, such as row 0. Error,
  clamp, or produce a `#REF!`-style result?
- **Cached results.** Should the duplicate recompute cached formula result values or leave them
  stale until the next calc?
- **Cross-sheet references.** For `Sheet2!A1` in a copied formula, translate the local coordinates
  only and leave the sheet qualifier intact.
