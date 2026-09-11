# A number format declared on a column must reach that column's cells

Cluster: styles

## Scenario

An author declares columns up front and gives a column a number format (say a percentage `0.00%`)
as part of its style, then appends rows whose values fall into that column. The expectation is that
each written cell inherits the column's number format so the value renders formatted, with 0.12 shown
as "12.00%". The report was framed as streaming-specific, with streamed cells written under default
or General formatting and ignoring the column-level numFmt, but probing shows the column-declared
numFmt fails to reach the data cells in **both** the streaming and the buffered writers when the
column is declared via the columns array and rows are added by key. So the durable requirement is
broader than streaming: a column's declared number format must reach its cells regardless of write
path.

> Spec note, not a corpus case: the corpus already locks per-cell numFmt survival
> (`custom-numfmt-string-roundtrips-verbatim`, `date-value-written-as-serial-not-text`). This note
> records the distinct, currently-unmet requirement that a _column-level_ numFmt propagate to cells,
> so it can be designed and then locked once the propagation model is decided. Capturing it as a note
> avoids pinning a fragile assertion against today's ambiguous column-style application.

## Desired behavior

- A `numFmt` declared on a column applies to the data cells in that column: a value placed in the
  column is written with a style referencing that number format, so it renders formatted rather than
  as General.
- Behavior is **identical between the streaming and buffered writers** for the same column and numFmt
  declaration, and neither silently drops it.
- The column-level format composes predictably with a per-cell numFmt: an explicit cell numFmt wins;
  otherwise the column's applies.

## Open questions

- Propagation model: does a column style materialize onto each written cell's style record at write
  time, or is it stored once as a column style the reader resolves per cell? This affects file size
  and the styles table.
- Does declaring the column via the columns array behave the same as `getColumn(n).numFmt = …`? Today
  they may differ; the fork should make one predictable path.
- Interaction with the header row, which should not inherit a data-oriented numFmt like a percentage.

Related: `custom-numfmt-string-roundtrips-verbatim`, `column-style-does-not-force-hidden`,
`column-level-value-type`, `streaming-write-memory-and-shared-strings-tradeoff`.
