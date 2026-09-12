# Phonetic guides (furigana) survive a read and a save

Cluster: xlsx-io

## Scenario

Japanese Excel records the reading of text typed through the input method beside the text itself: a
string holding `漢字` also carries the furigana `かんじ`. The reading is what Japanese Excel sorts such
text by and what the `PHONETIC()` function returns, and a user can show it above the cell. A workbook
of names or addresses read by the library and saved again comes back with its text intact but every
reading gone.

In the package the reading is part of the string, not of the cell. A `CT_Rst` (a shared string's
`<si>`, an inline string's `<is>`, a note's `<text>`) holds, after its base text or runs, one `<rPh>`
per annotated span, `sb` and `eb` giving the span's start and end in the base text and a `<t>` its
reading, then a `<phoneticPr>` naming the font the reading is drawn in (`fontId`, an index into the
stylesheet's fonts), its script (`type`) and its alignment. Whether the reading is shown is a separate
flag: `ph` on a `<c>` and on a `<row>`, `phonetic` on a `<col>`. A binary workbook carries the same data
in the phonetic tail of a `RichStr`.

> Spec note, not a corpus case: the reader already keeps a reading out of the text it annotates
> (`phonetic-run-not-read-as-cell-text` locks that), so nothing is corrupted. What is missing is a model
> for the reading, which is an API decision with no recorded demand yet; the open questions below are
> that decision. It becomes a corpus case once a reading has somewhere to live, against a fixture Excel
> saves with readings set through `Range.Phonetics` and shown through `Range.Phonetic.Visible`.

## Desired behavior

- A string read with phonetic runs keeps them, each span and its reading, and its `<phoneticPr>`, and a
  save writes them back in the same container they were read from.
- The show-reading flags on cells, rows and columns survive a save.
- Two cells with the same text but different readings do not share one shared-string entry, and two
  cells with the same text and reading still do.
- Changing a cell's text does not leave a reading describing text the cell no longer holds.
- A reading's font survives the stylesheet being re-interned on write, since `fontId` is an index the
  writer renumbers.

## Open questions

- Where a reading lives. A field on the string value makes every furigana-bearing cell in a Japanese
  workbook something other than a plain `string`, which every consumer testing `typeof value` would
  trip over. A facet on the cell, beside `note`, keeps the value a string but has to say what happens
  to the reading when the value is reassigned: dropped, as the fourth bullet above wants, or kept.
- Whether a rich-text value's reading rides on the value (it is already an object) while a plain
  string's rides on the cell, or both on the cell.
- How `<phoneticPr fontId>` is modelled: resolved to a `Font` on read and interned on write, or kept as
  an index into the preserved font table.
- Whether a note's body, which is plain text by model, keeps a reading at all.
- The streaming writer, which interns strings row by row, and the binary reader, which leaves the
  `RichStr` phonetic tail unread today.

Related: `phonetic-run-not-read-as-cell-text`.
