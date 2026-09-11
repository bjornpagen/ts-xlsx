# Custom table styles: a cross-part claim, and what may be refused

Cluster: styles

## Scenario

A table style is a named, reusable look, a branded header row and a chosen stripe colour, that a table
applies to itself by putting that name in `tableStyleInfo/@name`, exactly as it would name one of
Excel's built-in gallery styles. A workbook can define its own beside the gallery.

`Workbook.addTableStyle({name, elements})` is that API. Each element names a region
(`ST_TableStyleType`) and carries a `DifferentialStyle`, interned into the same shared `<dxfs>` table
conditional formatting uses, so a header row and a total row painted alike cost one entry.

## Why this needed Excel Desktop, not just the oracle

The claim spans three parts: the **table part** names a style, the **styles part** defines it, and the
**dxf table** backs each of its elements. Every one of those can be individually schema-valid while
the whole says nothing to Excel and the table renders completely unstyled, a file that opens clean
and looks wrong, which no schema check and no round-trip of our own can catch.

So it was put to Excel directly, reading `DisplayFormat` (the _effective_ format: direct formatting
plus table style plus conditional formatting) on a table whose cells carry no formatting of their own:

- the authored style appears in the workbook's own table-style gallery;
- the header row reads `DisplayFormat.Font.Bold = true` while `Font.Bold` is `false`, so the bold comes
  from the style, not from the cells;
- `firstRowStripe` with `size="2"` bands both data rows, confirming band width is honoured;
- the file opens with no repair.

Recorded in `test/corpus/fixtures/excel-oracle/authored-table-style-renders.json` (Excel 16.0 build
20131). That observation _seeds_; the corpus case locks it with the structural half, every element's
`dxfId` resolved through the emitted dxf table and the table's name resolved against the emitted
definitions (ADR 0012's seed-plus-lock split).

## What is refused, and why refusing is right here

Two things throw at `addTableStyle`:

- **An empty name.** A table references its style by name; an unnamed definition is unreachable.
- **A `size` on anything but the four stripe types** (`firstRowStripe`, `secondRowStripe`,
  `firstColumnStripe`, `secondColumnStripe`), or a `size` that is not a positive integer. ECMA-376
  restricts band width to those four and Excel ignores it elsewhere.

Both produce a file Excel opens without complaint and then quietly does nothing with, the failure
class that never gets found. The same reasoning the ARGB normaliser uses: reject at the boundary
where the caller can still see the cause.

## What is deliberately _not_ refused

**`TableStyleInfo.name` is not validated.** A table may name a style nothing defines, and this
library will write it.

The alternative was to check the name against the built-in gallery plus the workbook's own custom
styles. Rejected, for two reasons that point the same way: the built-in list grows with Excel, so a
name from a newer version would be wrongly refused; and a reader must never make a file Excel opens
unreadable, so a writer that threw would make round-tripping such a file impossible. With no
diagnostics channel in the library, the only choices were "throw" and "accept", and accepting is the
one that cannot break a working file. **If a warning channel is ever added, this is the first thing
that should use it.**

## Also not modelled

`defaultTableStyle` and `defaultPivotStyle` on the `<tableStyles>` container are **preserve-only**. They
tell Excel which style to pre-select for a table the _user_ inserts later; every table this library
writes states its own `tableStyleInfo`, so nominating a default would change nothing about the file's
appearance. Preserved faithfully, not authorable.

## Element application order

Elements are applied in the order ECMA-376 fixes, not the order they were authored: whole table,
column stripes, row stripes, last column, first column, header row, total row, then the four corner
cells. A row stripe therefore wins over a column stripe, and both win over whole-table formatting.
Worth knowing when a colour appears not to take. Pivot styles have their own longer order.

## Where this lives

`src/core/table-style.ts` (the model and `checkTableStyle`), `Workbook.addTableStyle`,
`StyleRegistry.addTableStyle` (`src/io/xlsx/styles.ts`).

Related: `authored-custom-table-style-renders`, `custom-table-style-definition-survives-roundtrip`,
`table-style-none-produces-unstyled-table`.
