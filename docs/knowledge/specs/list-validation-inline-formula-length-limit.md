# Inline list validation silently breaks past the 255-character formula limit

Cluster: data-validation

## Scenario

A user attaches a list-type data validation and supplies the allowed options as an **inline
comma-separated formula string**, say `formulae: ['"opt1,opt2,opt3,…"']`. The spreadsheet format
enforces a hard limit of **255 characters** on the inline drop-down formula string for a single cell.
When the joined string exceeds 255 characters, the written file is _structurally valid_ but the
application silently refuses to show the drop-down: the cell has a validation attached, yet no visible
list. Because nothing warns the author, the failure looks like a library bug when it is actually an
application-format constraint.

The known-good alternative has no length limit: move the options into a **range**, on any sheet, and
reference it, as `formulae: ['=Sheet2!$A$2:$A$21']` or `formulae: ['$D$5:$F$5']`.

> Spec note, not a corpus case: the written output is structurally valid, so there is no malformed
> serialization to assert against. The defect is a silent usability cliff plus an undecided design
> choice between warning, throwing and documenting. The 255-char fact and the warn-on-overflow
> requirement are durable product guidance.

## Desired behavior

- **Do not let an author cross the limit silently.** When an inline list-validation formula string
  would exceed the 255-character limit, the library surfaces it, at minimum a clear diagnostic and
  under a strict or validate mode a thrown error, rather than emitting a validation that opens as an
  invisible no-op.
- **Point at the fix.** The diagnostic names the working alternative: move the options to a range and
  reference it, cross-sheet or same-sheet, which the format allows without a length cap.
- **The range-reference form is first-class and lossless.** A list validation whose source is a range
  reference round-trips faithfully, already locked by corpus cases, so the recommended escape hatch
  is solid.
- The 255-char limit applies to the **inline literal form only**; range and defined-name sources are
  unbounded and are the right tool for long option sets.

## Open questions

- Warn or throw: is a console diagnostic enough for the default path, with a strict mode that
  throws, or should authoring an over-limit inline list always fail fast?
- Should the library **auto-spill** an over-limit inline list into a hidden helper range and reference
  it automatically, or is that too much implicit magic, since it adds a sheet or range the author did
  not ask for?
- Is the 255 count measured in UTF-16 code units, code points, or bytes for non-ASCII options, and
  does the target application count the surrounding quotes and separators?

Related: `multiselect-dropdown-validation`, `cross-sheet-list-validation-x14`,
`list-validation-value-source-forms-roundtrip`, `list-validation-defined-name-and-cross-sheet-range-source`,
`whole-column-data-validation-bounded-memory`.
