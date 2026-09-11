# A formula-cell value must be constructible from just the formula (no mandatory date-system flag)

Cluster: types

## Scenario

A caller authoring a cell with a formula constructs the value object using the public formula-cell
value type. The type declares the epoch-selection flag (1900 against 1904 date system) as a
**required** field, so TypeScript forces the caller to supply a workbook-level concern on every
individual formula cell, even though nothing about that flag belongs to a single cell. The natural
authoring shape, a formula string, optionally a cached result, and nothing else, does not type-check,
and callers resort to casts or filler values to satisfy the compiler.

> Spec note, not a corpus case: this is a public-type-surface ergonomics decision, not a runtime bug,
> and there is no serialization to baseline. The durable value is the contract: which fields of the
> formula-cell value type are genuinely required to author one. The type surface is verified by
> type-level tests over the public API, not the runtime adapter.

## Desired behavior

- **The date-system flag is not required to author a formula cell.** The 1900-against-1904 base is a
  property of the _workbook_ (`workbookPr/@date1904`), governing how all serial dates in the package
  are interpreted, and it is not a per-cell attribute. The formula-cell authoring type must therefore
  let a caller construct a value from just the `formula` string, optionally a cached `result`, and
  nothing else.
- **If retained for read-back, it is optional and defaulted.** Should the value type still expose the
  flag, so a value read back from a workbook can carry the base that was in effect, it is
  optional and defaults from the workbook's date system, never a field the author must provide.
- **The required and optional split reflects the format, not implementation convenience.** Required:
  `formula`. Optional: `result`, shared-formula linkage (`shareType`, `ref`, `sharedFormula`), and any
  read-back-only metadata. The types _are_ the docs: a caller should be able to see, from the type
  alone, the minimal shape needed to write a formula cell.

## Open questions

- Does the same over-required-field problem exist on the _other_ cell value types (rich-text,
  hyperlink, shared-formula, error)? Audit the whole cell-value union for fields marked required that
  are really workbook-level or read-back-only, and fix them as a class.
- Read-back shape: when a formula cell is read from a 1904 workbook, does the returned value need to
  carry the base at all, or is the workbook the single source of truth the consumer already has access
  to? Prefer not duplicating workbook state onto every cell.
- Distinct authoring and read-back types: is it cleaner to expose a narrow _input_ type, what you must
  supply to write, separate from a wider _output_ type, what a read surfaces, rather than one type
  doing both?

Related: `column-definition-type-is-partial-on-write`, `public-types-node-stream-portability`,
`date-value-timezone-conversion`.
