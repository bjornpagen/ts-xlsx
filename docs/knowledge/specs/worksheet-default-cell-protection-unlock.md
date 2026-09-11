# Worksheet-default cell protection: unlock all cells without per-cell style bloat

Cluster: styles

## Scenario

A common template pattern ships a protected worksheet: a few header rows stay locked, but every
other cell must be editable by the end user. In OOXML, cell protection (the `locked` flag) only takes
effect once the worksheet itself is protected, and every cell is `locked` by default. To invert that
today, locking a few cells and unlocking the rest, the only route is to iterate over a guessed
rectangle of thousands of rows and columns setting each cell's protection to unlocked. That is both a
correctness hazard, since the guessed bounds may be too small and leave cells unexpectedly locked, or
too large and inflate the sheet, and a size hazard: writing an explicit unlocked style to thousands of
cells materializes a style record and a per-cell style reference for each, bloating the file
dramatically. The user wants either a single operation to unlock all cells, or a way to author a
worksheet whose cells are unlocked by default so only the intentionally locked cells carry deviating
metadata.

> Spec note, not a corpus case: this is a missing authoring capability, expressing a non-default
> protection baseline efficiently, not a malformed-output bug with a current behavior to assert.
> Sheet-level protection authoring and password-hash compatibility are already covered by corpus
> cases; this note is the umbrella policy for the _default cell protection_ axis and its size
> contract. It becomes a corpus case once the surface exists and a case can assert that unlocking N
> default-matching cells produces zero per-cell style deviations.

## Desired behavior

- **Worksheet-level default protection.** A caller can declare "this worksheet's cells are unlocked by
  default", and the inverse, as a worksheet or default-style property, without touching every cell.
  Locking and unlocking specific cells, rows, or columns then expresses only the _exceptions_ to that
  default.

- **The default belongs in the default cell format, not per cell.** OOXML models cell protection via
  the `<protection locked="…" hidden="…"/>` element inside a cell format (`xf`) record in `styles.xml`,
  and it only matters once `<sheetProtection>` is present. The `locked` default is `true` at the
  format level, so a workbook whose _default_ cell format carries `locked="0"` yields an all-unlocked
  sheet with **zero** per-cell metadata. The library must encode a worksheet's default protection in
  that default or normal format, and the sheet's default row and column formatting, the way Excel does,
  so the common case costs nothing per cell.

- **No per-cell style explosion for default-matching cells.** Unlocking, or locking, N cells that all
  match the worksheet default must not produce N distinct per-cell style references. It must produce
  none, because those cells already inherit the default. Only cells whose protection _deviates_ from
  the default carry their own style index. This is the load-bearing size contract: it is what makes a
  200k-cell "unlock everything but the header" sheet small instead of enormous.

- **Round-trip stability.** Reading a file that uses a non-default protection baseline, editing, and
  writing back must preserve the default-and-exception structure rather than flattening it into
  explicit per-cell protection on every cell.

## Open questions

- Surface shape: a `worksheet.protection` authoring object with a `defaultLocked`, or
  `cellsLockedByDefault`, flag, against a method like `worksheet.unlockAllCells()` that sets the
  default format? The declarative default is more honest and composes with the exceptions, and a
  helper method can wrap it.
- Interaction with row- and column-level default formatting: unlocking a whole column should ride the
  column's default format, not fan out to its cells. Define which level "wins" when a cell, its row,
  its column, and the sheet default all specify protection.
- Whether `hidden`, the formula-hiding protection flag, follows the same default-and-exception model as
  `locked`, since they share the `<protection>` element.
- Should applying an explicit protection equal to the current default be a no-op, emitting nothing, so
  round-tripping and idempotent edits don't reintroduce per-cell bloat?

Related: `workbook-structure-protection-authoring`, `sheet-protection-password-hash-compatibility`,
`set-style-over-cell-range`, `style-dedup-value-based-and-cell-add-style`,
`cellstylexfs-named-style-fill-roundtrip`.
