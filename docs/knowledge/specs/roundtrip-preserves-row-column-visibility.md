# A read→write round-trip must preserve row/column visibility

## The scenario

A user opens an existing `.xlsx`, makes no edits or a single-cell edit, and writes it
back. The output refuses to open in Excel and LibreOffice, which reject it,
while a lenient in-browser viewer still renders it. The user traced the difference to
**every row and column being written as hidden**, and manually re-setting visibility made
the file open again. A no-op round-trip must never turn a visible sheet into a wholly
hidden one, and must never emit a document mainstream consumers reject.

## Desired behaviour

- Row and column **visibility (`hidden`) round-trips faithfully**: a row or column visible
  on read is visible on write, and a genuinely hidden one stays hidden. No path defaults
  visibility to hidden.
- A sheet that was not wholly hidden on input is never serialized as wholly hidden.
- The written package is accepted by all mainstream consumers (Excel, LibreOffice,
  Sheets), not only lenient viewers, so this is also a structural-validity concern
  rather than only a visibility-flag concern (see
  [[excel-repair-on-open-structural-constraints]]).

## Root cause (legacy)

Unconfirmed. The reported symptom is "all rows and columns hidden after a no-op
round-trip," but it is not yet isolated whether the true defect is a visibility flag
mis-serialized on the read-then-write path, say a default `hidden` leaking onto every
row and column model, or a separate structural corruption that merely _manifested_ as an
all-hidden sheet in the viewer that could open it. The two have different fixes.

## Open questions for the rebuild

- Is the defect visibility mis-serialization, or a structural package corruption that
  presents as "all hidden"? Disambiguating needs a reproducing input file, and the
  original was only ever shared via a since-dead link, so this is recorded as design
  intent rather than a corpus lock.
- Where does a default `hidden` value come from on the row and column model, and can the
  model distinguish "explicitly visible" from "unset" so a round-trip never invents a
  hidden flag?
- Should the writer refuse to emit a wholly-hidden workbook, since Excel treats a workbook
  with no visible sheet as invalid, or warn? This overlaps the write-time validation
  policy in [[excel-repair-on-open-structural-constraints]].
