# ADR 0042: A hyperlink belongs to the sheet, not to a cell's value

**Status:** Accepted 2026-09-12

## Context

A hyperlink was a kind of cell value: `{hyperlink, text, tooltip?, range?}`, a destination with the
label the cell shows. Assigning one made the cell a link; reading a file folded each `<hyperlink>`
element onto the value of the cell at the top-left of its `ref`.

OOXML does not store a link that way. The cell holds its value as any cell does, and the link is a
`<hyperlink ref="…">` element in a sheet-level list beside the grid. The mismatch cost data in two
ways, both confirmed against Excel 16.0 (build 20326) and recorded in
`test/corpus/fixtures/excel-oracle/hyperlinks-through-edits.json`:

- **A link over anything but text had nowhere to go.** Excel puts a link on a number, a date, a
  formula or a boolean as readily as on a label, and leaves the value alone. A label-shaped value
  could keep either the number or the link, so the reader kept the number and dropped the link.
  Every linked total in a report lost its link on the first save.
- **A link is a range, not a cell.** A link over `A2:A4` grows to `A2:A5` when a row is inserted at 3.
  When row 2, its first row, is deleted, Excel shrinks the link to `A2:A4` and keeps it; a link riding
  on its top-left cell went with that cell. And a link added over `B1` inside a link over `A1:C1` is
  kept beside it, which one link per cell cannot say.

## Decision

**A hyperlink is sheet state, held in a range-bound overlay like validations and conditional formats.**

- `Hyperlink` is `{ref, target, tooltip?}`, published from `/core`. `ref` is a cell or a rectangle,
  stored in its canonical spelling; `target` is a URL or path, or a `#`-prefixed place in the workbook.
- `Worksheet.addHyperlink(link)` adds one, replacing a link over the same `ref`;
  `removeHyperlink(ref)`, `hyperlinks` and `hyperlinkAt(reference)` read and remove them. Where two
  links cover a cell, the one added last is the one it opens. A whole-row or whole-column `ref` is
  refused, since a link covers cells.
- `HyperlinkOverlay` (`core/hyperlink.ts`) moves every link through a splice with `shiftRect`, so an
  insert inside grows it, a partial delete shrinks it, and a delete of all of it takes it.
  `duplicateRow` gives each copy the source row's own links, as Excel's row copy does. The model
  snapshot carries the links as a facet of its own.
- `HyperlinkValue`, `isHyperlinkValue` and `ValueType.Hyperlink` are deleted. A cell's label is its
  plain or rich-text value.
- The writer serialises the overlay in insertion order, which is also what a re-read produces, so the
  link that wins a shared cell is the same after a save. The reader adds every `<hyperlink>` whose
  `ref` names cells, whatever the cells hold. The streaming writer no longer carries links across a
  row's eviction, because nothing about a link lives in a row; `WorksheetStreamWriter.addHyperlink`
  may cover a row already committed.

## Alternatives

**A `hyperlink` property on `Cell`, beside `note`.** Closer to the old API, and it would have kept
links on non-text cells. It cannot express the other half: a link anchored to a cell is lost with that
cell, and a cell can hold one link where Excel keeps two over it. Notes are modelled on the cell
because a note does belong to one cell; a link does not.

**Keep `HyperlinkValue` for text and add the overlay for everything else.** Two representations of one
thing, with the reader choosing between them by what the cell happens to hold, and every consumer
checking both.

## Consequences

- **Breaking.** `cell.value = {text: 'Docs', hyperlink: url}` becomes `cell.value = 'Docs'` and
  `sheet.addHyperlink({ref: cell.address, target: url})`. A value read from a file is now the label
  itself, never an object carrying the destination.
- **Not mirrored:** Excel's Clear Contents removes a link from the cells it clears. Assigning a cell the
  empty value is not that command, and leaves a link covering the cell in place; `removeHyperlink` is
  the explicit way.
- **Not modelled:** the `display` attribute Excel writes beside some links. It repeats text Excel shows
  from the cell, and a read ignores it.
