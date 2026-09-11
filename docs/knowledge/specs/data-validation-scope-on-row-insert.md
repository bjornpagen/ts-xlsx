# Data validation scope when inserting rows into a table

Cluster: data-validation

## Scenario

A user builds a table whose body cells carry a data validation, say a dropdown list, and cell
protection. After opening the file, they insert a new row inside the table by right-clicking a row
header in the spreadsheet application. The newly inserted row picks up the data validation and
protection of the row above it. The user reports this as unexpected inheritance and asks the library
to prevent it.

> Spec note, not a corpus case: the inheritance the reporter observed is performed by the desktop
> application's UI when it inserts a row, so a file-writing library cannot control what the application
> does after the file is opened and there is nothing to baseline against current library behavior.
> The durable value is defining what the _library's own_ row-insert operations do with validation
> ranges, which is a real and separate question.

## Desired behavior

The load-bearing distinction is between what the spreadsheet _application_ does on a UI row-insert,
which is out of our control, and what _this library's_ row-insert and splice operations do to the
validation `sqref` ranges. The library must have a defined, predictable policy for the latter:

- **Data validations are anchored to explicit cell ranges (`sqref`), independent of table geometry.**
  When the library inserts or splices rows, it must adjust those ranges the same way it adjusts other
  range-anchored artifacts (merged cells, table refs, print areas, anchored images, all already
  locked by corpus cases). A validation covering `B2:B10` with a row inserted at row 5 should extend to
  `B2:B11`; a validation on rows entirely below the insertion point shifts down; one entirely above is
  untouched. The validation must never silently drop, duplicate, or leave a dangling range that
  triggers a repair prompt.

- **Inserting a row does not fabricate a new validation the source data did not have.** Extending an
  existing `sqref` to cover an inserted row inside its span is correct; inventing a validation on a row
  that was outside every validation range is not. "Inherit the row above's formatting" is a UI
  convenience, not a file-format guarantee, so the library's programmatic insert should default to
  extending existing ranges only, not copying validations onto genuinely new rows unless the caller
  asks.

- **Round-trip stability.** Reading a file with table-scoped validations, inserting or removing rows,
  and writing back must produce validation ranges that still open without an Excel repair prompt and
  still cover the intended cells.

## Open questions

- Should a programmatic "insert row into table" offer an explicit option to inherit or not inherit the
  neighbouring row's validation and protection, mirroring the UI convenience but under caller control?
- When a validation range's top or bottom edge coincides exactly with the insertion point, does the
  inserted row fall inside or outside the range? Define the boundary rule (inclusive-grow or
  shift-only) and lock it.
- Does the table's own structure (the `tableColumns` or autofilter ref) participate, or are validations
  purely `sqref`-driven and independent of whether the range happens to sit inside a table?

Related: `data-validation-survives-template-roundtrip`, `splice-rows-preserves-merged-cells`,
`splice-rows-updates-table-and-image-refs`, `data-validation-whole-column-range-writes-single-sqref`.
