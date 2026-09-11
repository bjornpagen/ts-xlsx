# Defined names must be able to carry a per-sheet scope

## The scenario

A workbook can declare the **same defined name on more than one sheet**, each scoped
to its own sheet. A common pattern: `Sheet1` defines `Input1` pointing at `Sheet1!$A$1` and
`Sheet2` also defines `Input1` pointing at `Sheet2!$B$2`, as two distinct, sheet-scoped
names. In Excel these are independent: a formula on `Sheet1` that mentions `Input1`
resolves to `Sheet1`'s definition, and the same text on `Sheet2` resolves to `Sheet2`'s.
Excel stores this with a `localSheetId` on the defined-name entry; a name with no
`localSheetId` is workbook-global.

## The limitation this must not reproduce

The legacy model keys defined names by **name alone**, a single map from name to
ranges. It has no place to record a scope, so it cannot represent two same-named
names with different sheet scopes. Assigning the name on two sheets does not error.
Instead both cells are folded into **one workbook-global name with two ranges**
(`Input1` pointing at `[Sheet1!$A$1, Sheet2!$B$2]`). The consequences reported by users:

- The intended per-sheet scope is lost, and the surviving name is workbook-global.
- A formula meant to reference the _local_ `Input1` becomes ambiguous and can
  resolve to the wrong sheet's cell.
- Because there is only one entry, editing or round-tripping can drop one sheet's
  definition, so "only the last instance is recorded" is a real failure mode on the
  read/merge path even though the naive in-memory assignment happens to keep both.

## Desired behaviour

- The model represents a defined name as a **(name, scope)** pair, where scope is
  either workbook-global or a specific worksheet, mirroring OOXML's `localSheetId`.
- Two sheets may each declare `Input1` scoped to themselves without collision, and both
  survive read then write.
- Name resolution is scope-aware: a lookup from a worksheet prefers that sheet's
  scoped name, then falls back to the workbook-global name.
- The public API exposes scope on assignment and read-back, so
  `cell.names` or a defined-names view reports the scoped name that actually applies,
  not a merged global aggregate.

## Root cause (legacy)

Defined names are stored keyed by name, as a cell-matrix keyed by the name string,
with no scope dimension. There is no representation for `localSheetId`, so scoped
names from a file are flattened into workbook-global entries on read, and same-named
scoped assignments merge on write.

## Open questions for the rebuild

- API shape for declaring a scoped name: `worksheet.defineName(name, ref)` for
  a sheet-scoped name against `workbook.defineName(...)` for a global one, or a `scope`
  option on a single call.
- How `cell.names` should report when both a global and a local name cover the same
  cell. List both, or prefer local?
- Interaction with data-validation lists that reference a named range: the scope must
  survive so the dropdown still resolves in the written file (see
  [[defined-names-tolerate-non-address-tokens]]).
