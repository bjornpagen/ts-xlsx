# Producing .xlsx files Excel opens without a "repair" prompt

Cluster: writer fidelity

Reported as "Microsoft Excel needs to recover XLSX before opening".

## Problem

A file written by the library can be well-formed XML, open cleanly in LibreOffice and Google Sheets, yet still make Microsoft Excel display _"Excel found unreadable content... Do you want to recover the contents of this workbook?"_ On recovery Excel logs entries like `Repaired Records: Cell information from /xl/worksheets/sheet1.xml`, `Worksheet properties from /xl/workbook.xml`, `String properties from /xl/sharedStrings.xml`, or `Table from /xl/tables/table1.xml`, then strips or rewrites the offending part. Excel's recovery log names only the _part_, never the _cause_, so users debug blindly.

This is not one bug. It is the absence of a validation and sanitization layer that enforces Excel's structural rules at write time, plus a couple of concrete correctness defects. The durable goal: **if a workbook can be represented and Excel accepts it, we write it faithfully; if it cannot, we fail loudly at write time, or sanitize with an explicit documented policy, and never silently emit a file Excel must repair.**

## Excel constraints a writer must respect

Each was independently observed causing the repair prompt.

- **Worksheet name**: max 31 characters; may not contain `\ / ? * [ ] :`; may not be blank; must be unique, case-insensitively. Violations give a `Worksheet properties` repair.
- **Table name**: must be unique across the _entire workbook_, not just the worksheet; must begin with a letter or underscore; may not contain spaces, dashes, or other non-identifier characters; may not collide with a cell reference like `C1` or `P1`. Violations give a `Table` repair.
- **Table column headers**: must be unique within a single table. Duplicate header text, say two `Fri/24` columns, gives a `Table` repair. Note that real Excel silently de-duplicates by appending a suffix; a faithful writer should do the same or reject.
- **Cell text length**: a single cell value may not exceed 32767 characters, or Excel gives a `Cell information` or `String properties` repair.
- **Data validation dropdown list**: when the combined length of an inline list exceeds Excel's limit of about 255 characters, Excel repairs.
- **Numeric and date values**: `NaN`, whether from `parseFloat(null)`, `Number(null)`, a division producing NaN, or a null in a date column, must never be serialized as a cell value or a coordinate. See the defects below.
- **Formulas**: the separator is always `,` regardless of locale, and the library must not localize `,` to `;`. That is user confusion rather than a library defect, but worth documenting.

## Concrete correctness defects worth regression coverage

1. **Whole-column print area corrupts to NaN.** A defined print-area range with column-only bounds, `$A:$F`, round-trips (read then write) to `$ANaN:$FNaN`. The empty row component is being coerced to a number, yielding `NaN`. Correct behavior: whole-column, and by symmetry whole-row, ranges must survive a round-trip unchanged and never produce `NaN` in an address. This is a real, reproducible parser and serializer bug and belongs in the address-decoding corpus.
2. **Empty rich-text run with a font produces an invalid shared string.** A cell whose value is `{ richText: [{ text: '', font: { bold: true } }, { text: 'test' }] }` emits a `sharedStrings.xml` entry Excel repairs (`String properties`). Correct behavior: an empty run must be dropped or serialized in a form Excel accepts, and a written rich-text string must always be a valid, Excel-openable shared string.
3. **Dropped-but-referenced part dangles a reference. (Fixed.)** A no-op load-then-save of a workbook with an external-workbook link dropped the `xl/externalLinks/externalLink1.xml` part, and its `<externalReferences>` registration, while _keeping_ the `[1]…`-using formulas and defined names, leaving a dangling external reference Excel repairs on open. This was the single structural break behind the repair prompt on a rich real-world workbook, and a referential-integrity sweep of the re-emitted package found no other. Fixed by preserving external links through the workbook preserved-reference net. The general lesson: the writer regenerates rels, content-types and `workbook.xml` from the model, so any part it neither models nor preserves is dropped, and if a _retained_ element still points at it, the result is a dangling reference. See `external-workbook-reference-formulas`; locked by the `external-workbook-link-survives-roundtrip` corpus case. calcChain, customXml and sharedStrings also drop on this path but are referenced by nothing that survives, so Excel tolerates them: a dropped part only forces a repair when something still points at it.

## Prior art

- Excel specifications and limits (Microsoft support) are the authoritative list of the numeric, length and naming ceilings above.
- Real Excel silently sanitizes many of these, de-duplicating table columns and truncating over-long strings, rather than rejecting. That is a reasonable model for our writer's default policy, provided the sanitization is documented and optionally strict.

## Open questions

- Policy: for each constraint, do we throw at write time, sanitize-and-warn, or sanitize silently to match Excel? Leaning: sanitize-and-warn by default, with a strict mode that throws. Naming and uniqueness violations that would silently drop user data, such as duplicate table names, should probably throw.
- Where does validation live: at mutation time, so `addTable` or `addWorksheet` rejects a bad name immediately and gives a good stack trace, or only at serialization time? Mutation-time gives the best developer experience; serialization-time is the last-resort safety net. Likely both.
- Do we surface a machine-readable diagnostics list (part plus cause) so callers can inspect what was sanitized, instead of only console warnings?
