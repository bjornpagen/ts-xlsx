# The reader must not crash on a foreign-shaped rich-text shared string

Cluster: types

## Scenario

A workbook contains a cell whose value is rich text, multiple styled runs, stored in the shared
strings table, including content that represents a hyperlink expressed as rich text within a shared
string. When the reader parses that shared string it must accumulate the `<r>` runs into a rich-text
value object. A brittle parser assumes its accumulator is already an object, but for certain
foreign-generated shapes it is handed an empty string primitive instead, so setting the run
collection onto it throws `Cannot create property 'richText' on string ''`. The result is that an
otherwise valid workbook fails to open at all.

> Spec note, not a corpus case: probing shows the library reads its _own_ rich-text shared strings,
> including runs with fonts, without throwing, and the concatenated text and runs surface correctly.
> So the crash depends on a specific foreign-generator shape, an empty-string accumulator seeded
> before the runs, such as a leading empty `<t/>` or a hyperlink-as-rich-text quirk, that is not
> reproducible from the library's own output. The durable value is the robustness requirement and the
> exact failure signature; promote to a fixture-backed corpus case once a reliably-reproducing foreign
> sample is captured, the same posture as `rich-text-cell-value-writer-robustness` on the write side.

## Desired behavior

- **Reading a shared string with rich-text runs never throws**, regardless of how the `<si>` was
  shaped by a foreign generator, whether the entry leads with an empty text node, mixes a plain
  `<t>` with `<r>` runs, or carries a hyperlink expressed as rich text. The run accumulator must be
  initialized as the correct container up front, not assumed to already be one, and never
  `value.richText = …` on a value that may still be a primitive string.
- **The cell surfaces its content after read.** The affected cell reads back as a rich-text value
  whose concatenated run text equals the intended string, so downstream consumers see the text rather
  than a load failure.
- **Foreign-generator tolerance is the frame.** This is one instance of the broader rule that the
  reader must tolerate structurally-valid-but-unusual OOXML from non-Excel producers without an
  unguarded property access bringing down the whole open.

## Open questions

- What are the concrete foreign shapes that seed an empty-string accumulator? Capture at least one as
  a fixture (hyperlink-as-rich-text, or a leading empty `<t/>` before runs) to lock the regression.
- Should a malformed run, missing `<t>` or empty, contribute an empty string to the concatenation,
  or be skipped? Define it so the read is deterministic.
- Does the same accumulator-initialization bug affect inline strings (`<is>` on the cell) as well as
  shared strings? Audit both paths.

Related: `rich-text-cell-value-writer-robustness`, `streaming-read-resolves-shared-strings`,
`hyperlink-display-text-can-be-rich-text`, `html-fragment-to-rich-text-cell-value`.
