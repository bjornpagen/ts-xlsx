# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project adheres to
[Semantic Versioning](https://semver.org/) — see
[ADR-0015](docs/decisions/0015-publishing-name-semver-and-first-version.md) for the
versioning policy, including why the first published version is `1.0.0` rather than a
`0.x` series.

This file tracks changes from its introduction forward. Earlier history — the full
ExcelJS-to-`ts-xlsx` rewrite — is recorded in `git log` and the [ADR series](docs/decisions/).

## [Unreleased]

### Added

- **`SheetProtection.legacyPasswordHash`** carries the legacy 16-bit `password="XXXX"` hash a file
  protected a sheet with, verbatim, so that protection survives a save. The library never derives one
  from a password; `protect` still writes only the agile credential.

- **`DataValidation.suppressDropDown` and `DataValidation.imeMode`**, with the `DataValidationImeMode`
  union and its `isDataValidationImeMode` guard on `/core`. The two attributes Excel writes as
  `showDropDown` and `imeMode` had no field, so a list validated without its dropdown arrow got the
  arrow back on save and an input mode was dropped. Both now round-trip in the standard and the
  extended form; `suppressDropDown` is named for what `showDropDown="1"` does, which is hide the arrow.

### Changed

- **BREAKING: `Workbook.tableStyles.styles` holds `{name, xml}` definitions rather than bare
  fragments.** An authored table style replaces the preserved definition it shares a name with, and
  the writer found that definition by pulling `name="…"` back out of its bytes: a file that quoted the
  attribute `name='…'` kept both definitions under one name. The name is now read by the scanner when
  the file is, and carried beside the fragment as `TableStyleDefinition.name`; read `.xml` where a
  fragment string was read before.

- **BREAKING: `StreamedRow` can no longer be constructed directly.** Get one from
  `WorksheetStreamWriter.addRow` or `addRows`, which was always the only meaningful source. A
  constructed row named a number its writer never handed out, so `new StreamedRow([], sheet, 5).commit()`
  evicted row 5 from the model and wrote nothing for it.

- **`decodeRange` refuses a row outside `1..1048576` with a `RangeError`**, as it already refused a
  column past `XFD`. `decodeRange('A1:B1048577')` used to return a range naming a row that cannot
  exist, which every caller then had to notice for itself.

- **BREAKING: conditional-format and data-validation numbers the schema cannot hold are refused at
  write.** A conditional format's `priority`, `rank` or `stdDev` that is not an integer (or a negative
  `rank`), and a `NaN` cfvo value, which was written as `val="NaN"`, now make `writeXlsx` throw
  `AuthoringError`; an authored `NaN` priority used to spread to every automatic priority after it. A
  data validation's non-finite numeric operand throws too, rather than being dropped, which silently
  turned a two-bound rule into a one-bound one. `Worksheet.addConditionalFormatting` refuses a range
  that names no cells.

### Fixed

- **A refused cell position left the sheet unreadable and unwritable.** `getCell('A0')`,
  `getCell('A1048577')` and `addTable({ref: 'A0', …})` threw, but only after storing the row they
  were refusing, so `rows()` and `writeXlsx` threw `row 0 is out of bounds` on every call from then
  on. The position is now refused before anything is stored.

- **A splice or duplicate that lands past the last line is refused before anything moves.**
  Inserting empty rows at row 1048576 (`spliceRows(1048576, 0, [], [])`) or duplicating a
  height-only last row left a line past the grid behind with the same result, a replacing
  `duplicateRow` near the bottom wrote some copies before throwing, and `spliceColumns(16384, 0, [],
  [])` reported success for columns that cannot exist.

- **`freeze`, `addRows` and `addColumns` no longer apply half of a refused call.** A `freeze` with
  nothing left to scroll set the frozen state before throwing, and a row wider than the grid wrote
  its first cells before throwing.

- **A shared formula over a whole column or row read back aimed at the master's lines.** A clone
  recovers its text by shifting the master's references, and only `A1`-shaped ones moved, so
  `COUNTIF(A:A,A1)` filled one column right read back as `COUNTIF(A:A,B1)` where Excel shows
  `COUNTIF(B:B,B1)`, and a save stored that. Column ranges (`A:C`) and row ranges (`1:5`) now shift
  as cells do, each end only when it is relative, and a range pushed off the grid becomes `#REF!`.

- **Furigana was read as part of the text it annotates.** Japanese Excel stores a phonetic run
  (`<rPh>`) beside a string's base text, in shared strings, inline strings and note bodies alike, and
  every reader gathered each `<t>` it met, so a cell holding `漢字` read as `漢字かんじ` and a save
  wrote that back. Phonetic runs are now skipped wherever a string is read; the reading itself is
  dropped rather than modelled.

- **A hyperlink over a number, date, boolean or formula erased the value.** A link folds into its
  cell as a `HyperlinkValue` whose label is text, so reading a linked non-text cell replaced `42` with
  an empty label, and a save wrote an empty string where the number was. The value is now kept and
  the link dropped. Keeping both needs hyperlinks modelled beside cells rather than inside their
  values, which is a separate change.

- **A sheet protected by the legacy password hash was saved unguarded, and a sheet with no `sheet`
  attribute was saved locked.** The 16-bit hash had no field, so a `<sheetProtection>` carrying
  `password="CC3D"` was rewritten without it and anyone could unprotect the sheet. And `sheet`
  defaults to false, but its absence read as protected, so a save locked a sheet the file had left
  open. Both now read as the file states.

- **A cell added to an open streamed row through `getCell` was dropped when the row committed.** The
  streaming writer rendered the cells `addRow` had returned and then evicted the whole row from the
  model, so `addRow(['a', 'b'])` followed by `getCell('D1').value = 'x'` wrote A1 and B1 only, and a
  hyperlink or note set that way went with it. A commit now renders the row as the model holds it.

- **Adding a picture to a sheet whose drawing holds a chart or shape dropped the chart or shape.** Such
  a drawing is kept byte for byte with none of its anchors modelled, and a worksheet references one
  drawing, so the writer planned a second one for the picture and referenced only that. `addImage`,
  `addImageAnchor`, `importImages` and the writer now refuse with an `AuthoringError` naming the sheet.

- **`readSheetRows` and `readWorkbookStream` decoded a cell with no format under no format at all.**
  Every cell falls back to the workbook's default format, xf 0, and `readXlsx` applied it, so in a
  workbook whose default format is a date a bare cell read as a `Date` through `readXlsx` and as a
  serial such as `45000` through the streaming readers. All three cell readers now resolve a cell's
  format through one rule. A streamed cell still reports a `style` only when it, its row or its column
  declares one.

- **Removing a VBA module could overwrite the project's `PROJECT` stream, and a module sharing a name
  with a stream in another storage did not parse.** Streams inside `vbaProject.bin` were found by bare
  name anywhere in the container, but [MS-OVBA] puts `dir` and the module streams in the `VBA` storage,
  `PROJECT` at the root and each UserForm's streams in a storage of its own, so names repeat. A module
  named `f` beside a UserForm, or one named `PROJECT`, failed with a compression signature error, and in
  the other directory order `removeVbaModule` wrote the module's bytes over the root `PROJECT`,
  destroying every module declaration. Streams are now read and replaced by their storage path.

- **Removing a VBA module from a project with non-ASCII `PROJECT` text threw, and rewrote the stream's
  line endings.** The stream was decoded and re-encoded to drop the module's lines, and under Shift_JIS,
  GBK, Big5 and EUC-KR the encoder can only invert ASCII, so a Japanese `Description` made
  `removeVbaModule` throw `VbaAuthorError` over text the caller never touched; a stream mixing CRLF and
  LF also came back with one ending throughout. The lines are now cut out of the bytes, and every other
  byte stays as the file wrote it.

- **`addVbaReference` with a path of a few hundred thousand characters overflowed the call stack.** A
  path longer than the 32,767 characters a Windows path can hold is now refused with `VbaAuthorError`,
  and the reference records are built as bytes.

- **A worksheet model assignment that failed part-way left the destination emptied and half loaded.**
  `dst.model = model` reset the sheet before replaying the model, so overlapping merges or a data
  validation naming no cells threw with the sheet's own content already gone and part of the model
  applied. The model is now applied to a scratch sheet first; an assignment that fails throws and
  leaves the destination as it was.

- **A row or column splice left a hyperlink's clickable range and a data table's ranges behind.** A
  splice moved a cell but not the coordinates inside its value: a link over merged `D1:H1` moved to
  `D2` with its `range` still `D1:H1`, so the writer emitted the hyperlink over the old cells and a
  re-read put an empty link on `D1`. A hyperlink's `range` and a data table's `ref`, `r1` and `r2` now
  move with the cell, as a shared-formula clone's master address already did.

- **An automatic page break was saved as a manual one, and page breaks stayed put through a splice.**
  The writer marked every `<brk>` `man="1"` and counted all of them as manual, and the reader dropped
  `min`, so a producer's automatic breaks hardened into author-set ones on the first save. Breaks now
  keep `man` and `min`, and a break authored without `man` is still written as manual. `rowBreaks` and
  `columnBreaks` also sat outside the row and column splices, so inserting a row above a break left it
  on the old line; a break now moves with the line after it and goes when that line is deleted, as
  Excel does. The `PageBreak` documentation said a break precedes line `id`, where Excel counts `id`
  as the last line before the break (a break above row 20 is `id: 19`); it now says so.

- **A picture anchored between cells without `editAs` stopped resizing with them after a save, and a
  drawing holding an absolutely anchored or linked picture lost it.** The schema default for a
  two-cell anchor's `editAs` is `twoCell`, but an absent attribute read as none and was written back as
  this library's authoring default, `oneCell`. And the reader decided whether a drawing was fully
  modeled from a list of content to skip that missed `absoluteAnchor` and a linked `<a:blip r:link>`,
  so such a drawing was rewritten from its other pictures alone. An absent `editAs` now reads as
  `twoCell`, and a drawing is modeled only when every anchor holds exactly one embedded picture;
  anything else keeps the drawing whole.

- **An error value whose code was not an Excel error wrote markup into the sheet.** A cell's error code,
  and a formula's cached error result, went into `<v>` unescaped on the claim that the codes are a closed
  set, which nothing checked: `{error: '</v></c><evil/>'}`, smuggled past the type, closed the cell and
  wrote an element of its own. A code outside `ERROR_CODES` is now refused with `AuthoringError`.

- **A table's totals row wrote a malformed formula for a column name with special characters.** The
  `SUBTOTAL` a totals cell materialises interpolated the column name bare, so `Price [USD]` wrote
  `SUBTOTAL(109,T[Price [USD]])` and `it's #1` wrote `SUBTOTAL(109,T[it's #1])`. The reference is now
  spelled by Microsoft's structured-reference rules: double brackets for a name with a special
  character, and `'` before `[`, `]`, `#`, `'` and `@`. Formula scanning also honours that escape, so an
  escaped bracket no longer hides the rest of a formula from `_xlfn.` prefixing.

- **An authored theme colour was stored as spelled, so it resolved to garbage.** `setTheme` checked a
  colour and then kept the raw string: `#FF0000` read back from `themeColors` as `#FF0000` and
  `resolveColor({theme: 4})` returned `FF#FF0000` (`FFNANNANNAN` with a tint), and `80FF0000` resolved
  to ten digits. `themeColors` and `themeOverrides` now return six uppercase hexadecimal digits
  whatever the authored spelling or the theme part's casing, and `resolveColor` resolves a
  caller-built scheme slot that is not six hex digits to `undefined`.

- **Future-function prefixes came from an incomplete and partly wrong list.** Excel stores FILTER, SORT
  and PY as `_xlfn._xlws.FILTER`: formula text read from such a file kept `_xlws.FILTER`, a rewrite
  dropped its `_xlfn.`, and an authored FILTER or SORT was written as `_xlfn.FILTER`. The list now
  follows [MS-XLSX]'s future-function table plus the functions Excel writes with `_xlfn.` that the table
  does not list yet. CEILING.MATH, FLOOR.MATH, FORECAST.LINEAR, the FORECAST.ETS family, GROUPBY,
  PIVOTBY, PERCENTOF, the REGEX functions, TRIMRANGE, IMAGE and STOCKHISTORY were written bare, and so
  were SINGLE and ANCHORARRAY, the forms Excel stores `@` and a spill reference in. ISO.CEILING was
  prefixed although Excel writes it bare. A LET or LAMBDA parameter referenced in a different case
  (`LET(x,1,X+1)`) is now recognised, and written in its declared spelling as Excel writes it.

- **A time of day on 1900-02-28 was written as the non-existent 1900-02-29.** The phantom-leap-day
  correction compared the instant rather than the day, so noon on 1900-02-28 became serial 60.5, and
  reading 59.5 then writing it back changed the stored value.

- **Date formats rendered every meridiem as `AM`/`PM` and read a literal `s` as seconds.** `h:mm A/P`
  rendered `9:30 AM` where Excel shows `9:30 A`; the short form now keeps each letter's case as Excel
  does (`a/p` renders `a`), while the long form stays in capitals however it is spelled, which is also
  what Excel displays. An escaped or quoted `s` before `.0` no longer renders fractional seconds. CSV
  output changes for these format codes.

- **Conditional formats kept numbers and ranges from a file that the schema cannot hold.** A foreign
  `priority="1.5"`, `rank="2.5"` or `stdDev="-1.5"` was written back verbatim, and a block whose `sqref`
  was missing or did not decode was written back as `sqref=""` or as its unreadable text. Each is now
  dropped on read.

- **A gradient fill's non-finite `degree` or inset was omitted instead of refused**, because it was
  written behind a truthiness test and `NaN` is falsy; it now throws `AuthoringError` at write, as
  every other unwritable number does. Fills are also interned by the XML they produce rather than by a
  separate field list, so two fills written identically (a solid fill that states the default
  background, a gradient with `degree: 0`) no longer take two entries in the stylesheet.

- **A table with a nameless column was read shifted.** The column was skipped with nothing checking
  the count, so every later column moved one place left, the range shrank by a column, and a totals
  formula landed on the column before its own. A table whose column count differs from its `ref` width
  is now dropped whole, as one with an unreadable `ref` already was.

- **Optional booleans read an unrecognised token as `true`.** `totalsRowShown` and the
  `tableStyleInfo` banding flags on a table, `summaryBelow`/`summaryRight`, `fitToPage` and
  `showGridLines` on a sheet, and a conditional format's `aboveAverage` read `"yes"` as on, so the
  writer turned it into `"1"`. Such a token is now dropped.


## [3.1.0] — 2026-09-04

Twenty-one fixes, most of them on the read path, and most of them cases where a file this
library did not write came back as a plausible workbook rather than as an error. A worksheet
that leaves `r` off its cells read as blank. A package that points at its own parts through the
relationship graph, which OPC says it may, read as no workbook at all. A workbook binding the
relationships namespace to a prefix other than `r` loaded every sheet empty. None of the three
threw anything.

**Read the four breaks below before upgrading.** They land under a minor version number, which
is not what [ADR-0015](docs/decisions/0015-publishing-name-semver-and-first-version.md) §2 asks
for: each is a public API change and the policy calls for a major. This is the second time the
project has made that call at release time, after 2.1.0, and for the same reason. None of the
four is a redesign, and each closes a surface that was wrong rather than changing one that was
right. But the version number will not warn you the way it is supposed to, so this list has to.

- **`CsvWriteOptions.dateFormat` takes an Excel number-format code, not a moment.js token set.**
  Only formats carrying a time change what they emit, because `mm` was minutes and is now
  months.
- **`WorkbookStreamWriter.commit()` resolves with `undefined`** when the writer was given a sink
  and `stream` was never touched. Supply no sink, or touch `writer.stream`, and the bytes come
  back exactly as before.
- **`WorksheetStreamWriter` can no longer be constructed,** and `flushRow`/`flushedSheet` are no
  longer callable. `WorkbookStreamWriter.addWorksheet` is unchanged and is where a sheet writer
  comes from.
- **`Table.shiftRows` and `Table.shiftColumns` take one `AxisSplice`** instead of three loose
  numbers: `table.shiftRows(3, 0, 2)` is now
  `table.shiftRows({axis: 'row', start: 3, count: 0, delta: 2})`.

### Added

- **The 1904 date system, carried on `Workbook.dateEpoch`.** A workbook declares which calendar
  its serials count from, and nothing in this tree had ever read `<workbookPr date1904="1"/>`.
  Every date cell in such a workbook came back four years and a day early, through the buffered
  reader, the streaming reader and the BIFF12 reader alike; writing one back dropped the
  declaration while re-emitting the serials unchanged, so the next reader applied the other
  calendar and the document silently changed meaning. `DateEpoch` is `1900 | 1904` rather than a
  boolean, because the epoch year is what every conversion actually needs, and it is a required
  parameter on `dateToSerial`, `serialToDate` and `coerceDateSerial`, so no call site can forget
  which workbook it is converting for. `WorkbookStreamWriter` takes it at construction: it
  serialises each row as that row is committed, so a system changed part-way through would leave
  the rows before the change counting from a different day than the rows after it.

### Changed

- **BREAKING: `CsvWriteOptions.dateFormat` is an Excel number-format code, not a moment.js token
  set.** It is the same vocabulary `Cell.numFmt` holds, so `writeCsv(wb, {dateFormat: cell.numFmt})`
  now renders what the cell shows. It did not: the CSV writer had its own case-sensitive token table
  in which the month is `MM` and lowercase `mm` is minutes, so passing this library's own format
  codes emitted `2024-45-dd` for every date cell, with no throw and no warning. The break is in the
  *meaning* of the option rather than its type, so a consumer passing `"MM/DD/YYYY"` gets a different
  string instead of a compile error -- but the common formats are unaffected, because the tokens
  differ from the codes only in case and a format code is case-insensitive. The formats that change
  are those containing a time, where `mm` moves from minutes to months. See
  [ADR-0041](docs/decisions/0041-one-date-format-vocabulary.md).

- **BREAKING: `WorkbookStreamWriter.commit()` resolves with `undefined` when the writer was given a
  sink and `stream` was never touched.** The package went to the sink; retaining every chunk to
  concatenate them at the end cost a second full-size copy of the archive, which is the memory a
  caller passes `{filename}` or `{stream}` to avoid. A caller who supplies no sink, or who touches
  `writer.stream`, gets the bytes exactly as before.

- **BREAKING: `WorksheetStreamWriter` can no longer be constructed, and its `flushRow`/`flushedSheet`
  are no longer callable.** All three were the streaming writer's own plumbing on a published class:
  the constructor took a `StyleRegistry`, `flushedSheet()` returned a `FlushedSheet`, and between them
  they named seven internal types no entry barrel exports, so a consumer could hold a value whose type
  they could not write. A caller receives the sheet writer from `WorkbookStreamWriter.addWorksheet`,
  which is unchanged and is the only way it was ever meant to be obtained.

- **BREAKING: `Table.shiftRows` and `Table.shiftColumns` take one splice descriptor instead of three
  loose numbers.** `table.shiftRows(3, 0, 2)` is now
  `table.shiftRows({axis: 'row', start: 3, count: 0, delta: 2})`, and the new `AxisSplice` type is
  published from `/core`. `start`, `count` and `delta` are three `number`s in an order no call site
  can be read against, and transposing two of them is invisible to the compiler, to the linter and to
  Excel: the workbook that comes out is well-formed, with its merges, dropdowns, highlights, comment
  anchors and images re-anchored to cells the author never chose. Named fields cannot be transposed.

- **`encodeAddress` now refuses a row outside the sheet, as it already refused a column.** It bounded
  one axis and let the other through, so `encodeAddress(1, 0)` returned `"A0"` and
  `encodeAddress(1, 2 ** 31)` an address no reader can decode. Both now throw a `RangeError`. Every
  coordinate check in the model also states its bound the same way, so the three messages the API
  used to give for one mistake are now one message that names the axis, the value, and the limit.

- **Seventeen types a consumer could reach and not name are now published, and the eighteen missing
  token guards with them.** `TableColumn.totalsRowFunction` had the type `TotalsRowFunction`, which
  no entry exported, so the value was reachable and the type was not; the same held for
  `CellContent`, `DateEpoch`, `AxisHandle`, `CellPosition`, `TableGrid` and `PreservedTheme`. Every
  closed token union `/core` publishes now publishes its narrowing guard too (`isBorderStyle`,
  `isFillPatternType`, `isVisibility` and fifteen more): two were published and eighteen were not,
  with no rule saying why. Both are additive, and both are now gated.

- **The workbook part is read in one pass, not seven.** Its protection, its `<workbookPr>`, its
  window view, its sheet list, its defined names, its `<pivotCaches>` registry and its
  `<externalReferences>` registry were seven readers with seven scans between them, on the same shape
  `readSheet` had already been given single-pass treatment for. All seven now see the whole part
  before the first sheet is read, which is also what puts the date system in place before the first
  cell decode rather than leaving that to a call sited above the loop and a comment asking it to stay
  there. A `RecordReader` also builds its `DataView` on first numeric read rather than in its
  constructor, and one is constructed per BIFF12 record in both the worksheet and the styles parser.

### Fixed

- **A worksheet that leaves `r` off its cells read back as a workbook full of blanks.** `r` is
  `use="optional"` on both `sml:CT_Row` and `sml:CT_Cell`: a producer may rely on document position
  instead, and several non-Excel writers do, so the nth `<row>` of `<sheetData>` is row n and the nth
  `<c>` of a row is that row's nth column. Excel opens such a file without comment. This reader
  required the attribute and lost every cell of every row, silently, through the buffered reader and
  the streaming one alike; stripping `r=` from a package this library had just written reproduced it.
  An absent `r` is legal, an `r` naming a cell that cannot exist (`A0`, `junk!!`) is malformed, and
  the two are no longer answered the same way.

- **A package that locates its parts through the relationship graph, as OPC allows, read as no
  workbook.** `xl/workbook.xml`, `xl/sharedStrings.xml` and `xl/styles.xml` are where Excel puts
  those parts, not where the format says they live; the reader consulted the graph for the theme and
  for the VBA project and nowhere else. The three failures are graded. A workbook part it cannot find
  opens nothing. A pool it cannot find is silent, and every `t="s"` cell reads as the empty string. A
  stylesheet it cannot find is worse than silent, because the date test reads `numFmt` off the
  resolved style to tell `45000` from `2023-03-15`, so an empty style table changes cell *types* with
  no error anywhere. Part names are also compared case-insensitively now, as OPC requires: a document
  at `XL/Workbook.xml` was reported as not a workbook, and an `<Override PartName>` cased differently
  from its zip entry fell through to `application/octet-stream` for a part being preserved verbatim.
  `readXlsbPackage` takes the resolved document path too, rather than assuming one.

- **A workbook binding the relationships namespace to a prefix other than `r` loaded every sheet
  empty.** Six sites indexed the attribute map with the literal `r:id`. A prefix is a local nickname
  and the rest of the reader knows it, matching on stripped local names, which is what made these few
  so damaging: the whole file reads perfectly and one feature vanishes, with nothing thrown and
  nothing logged. One site already hedged with `attrs['r:embed'] ?? attrs.embed`, which is the tell
  that this had been noticed once and never generalised, and which caught only the unprefixed
  spelling.

- **Five readers returned a plausible wrong value rather than failing.** An empty pooled string
  written `<si/>` committed no entry, so every later ordinal shifted by one and a `t="s"` cell past it
  resolved to its neighbour's string. `RunAccumulator` latched "inside a container" on the open of a
  self-closing one, which then never ended, so a `<t>` outside every container was absorbed as that
  container's text. A `t="s"` cell whose `<v>` was present and empty went through a bare `Number()`,
  and `Number('')` is 0, so the cell resolved to the first pooled string. A Strict-mode `t="d"` cell
  with unparseable text became a Date whose time is `NaN`, which satisfies every guard, survives into
  the model and serialises back as the literal `Invalid Date`. And the CSV reader discarded a bare CR
  instead of ending the row on it, splicing the next row's first field onto the last and losing every
  row boundary in a classic-Mac file.

- **A `<t>` nobody closed was answered with the previous cell's text.** `RunAccumulator.beginContainer`
  resets every field a string container owns, and the text capture was not one of them, so markup
  that opens a `<t>` and is then truncated leaves the capture armed across the container boundary and
  a stray `</t>` in the next container is answered with the previous one's buffer:
  `<is><t>orphan</is><is>loose</t></is>` read as `orphanloose`.

- **The two readers disagreed about one document, and the two writers about one model.** The
  streaming half of each pair is documented as producing what the buffered half produces, and in four
  places it did not, with no error on either side. OOXML lets a cell inherit its format, and the
  buffered reader resolved cell then row then column while the streaming reader read the cell's own
  `s` alone, so a cell under a date-formatted column came back as a `Date` from one reader and as
  `45000` from the other. A streamed hyperlink kept its visible label and lost what it pointed at,
  and a streamed note vanished with no comments part emitted at all, because both are serialised
  outside the `<row>` and were released with it. And `translateFormula` gave two different wrong
  answers on the read path, where the deltas come from a file's own shared-formula geometry: the
  column axis threw a bare `RangeError` that aborted a whole sheet read, the row axis emitted `A0` or
  `A-4`, which is not a reference. Both axes answer `#REF!` now, which is what Excel writes.

- **Seven paths let a file the library did not write choose its own cost.** A `<mergeCell>` cost the
  height it declared rather than the rows that exist, so 16,000 non-overlapping full-column merges,
  about 570 KB of XML that zips to a few kilobytes, bought roughly eight minutes of CPU, invisible to
  the inflate cap because after inflation the payload really is small. A finite negative
  `outlineLevel` hung the writer forever, walking straight past a non-finite gate that had reasoned
  about exactly that hazard. The MS-OVBA encoder rescanned its whole 4096-byte back-window for every
  output byte, 95 ms per 4 KB chunk of unmatched data, on the path `removeVbaModule` and
  `addVbaReference` take to recompress a `dir` stream that arrived in an `.xlsm`; a hash chain over
  three-byte prefixes brings a megabyte of unmatched data from about 10 s to 0.14 s, byte-identical
  on real VBA source. `MODULES_COUNT` was patched on an invariant stated only in a comment, so a
  crafted `dir` put both patch writes inside an unrelated record's payload and a declared count of
  zero underflowed to `0xFFFF`. The CFB reader validated two of three header layout fields, so a
  crafted mini-stream cutoff returned the wrong module source, and a directory entry's size read as
  its low 32 bits truncated an oversized stream silently. A duplicate zip entry name won last-write
  rather than being refused, which is the whole of a "same file, two meanings" attack. And the
  encoder accumulated into the `number[]` its own inverse's docstring rejects for costing several
  times its byte count in memory.

- **Six more paths spent more than they cost to write.** The CFB FAT was assembled by believing the
  header's sector counts, with ids free to repeat, so 1 MB of crafted container cost 272 MB of heap
  and a 10 MB project was an OOM rather than a typed failure; a file cannot hold more FAT sectors
  than it holds sectors, and that is the bound now. The directory's sibling tree is red-black and so
  logarithmic in depth, but a hostile file is under no such obligation, and linking every entry to
  its predecessor gave an uncatchable `RangeError` out of `Workbook.removeVbaModule` and
  `editXlsxVbaRemoveModule`. The decompressor's 64 MiB ceiling was per call, which bounds one bomb
  and nothing else: at the measured 230:1 amplification a 30 MB project of a hundred modules reached
  several gigabytes with every individual call well under its own limit, so one budget now runs
  across the whole project. A chunk could decompress past the 4096 bytes [MS-OVBA] 2.4.1.3.6 permits,
  and `bitCount` past the 12 that 2.4.1.3.19.3 bounds it to. Nothing bounded the number of `<col>`
  elements, each free to span the grid, so 154 KB of worksheet XML cost the buffered reader 13 s and
  the streaming one 1.7 s; a per-sheet work budget, four times the whole grid, brings both under
  150 ms. And the two package-level VBA edit functions handed raw caller bytes to an uncapped
  `unzipSync` rather than to the shared inflater every other reader goes through.

- **A VBA stream whose sector chain ended before its declared size read back as a well-formed
  prefix.** The compound-file reader collected what the chain offered and returned it, so a module's
  source came back truncated with nothing downstream able to tell -- the same silent-truncation
  failure the 4-GiB size check a few lines above it already refused. Such a chain now raises
  `VbaParseError` naming the chain and both byte counts.

- **Seven edits through the public API produced a sheet nothing could read back.** `insertColumn` and
  `spliceColumns` ran their insert pass inside the loop over the rows the grid already held, so
  inserting into an empty sheet wrote nothing whatsoever and reported no error, while `addColumn`,
  handed the identical array, materialised every row of it. Line metadata was the last splice
  participant doing its arithmetic by hand, and a hand-written shift does not clamp: a height on the
  last row plus an insert above it left a properties entry at 1048577, after which iterating the
  sheet threw and it could no longer be written or inspected, after a legal public call. A table's
  anchor was checked against the grid and the far corner derived from it was not, so `XFC1` with
  three columns reached XFE and the writer emitted a `<table ref>` naming rows that cannot exist. The
  conditional-formatting deep copy stopped one level short of `font`, `fill` and `border`, none of
  which is flat, so a stored rule shared all three with the caller. `translateFormula`'s `#REF!` guard
  could not be reached for a three-letter column like `ZZZ1`, which threw a bare `RangeError` first.
  And `getRange(2, 2)` manufactured the two missing corners as zeros and then complained about row 0,
  a coordinate the caller never wrote.

- **Seven structural edits lost state the writer would have emitted.** A cell carries eight facets of
  formatting and the tuple driving every copy listed six, so a row splice, a column splice, a
  `duplicateRow` or a `dst.model = src.model` dropped the quote-prefix flag and the link to a named
  cell style while the number format came through: inserting a row above a sheet turned a
  leading-apostrophe `'007` back into an unprefixed cell. The cell grid was the one splice participant
  doing raw arithmetic where merges, tables, images, overlays and shared-formula anchors all clamped
  through `shiftIndex`, so an insert on a sheet holding a cell in the last row or column threw a
  `RangeError` from inside the splice, and on the column axis left the sheet half-shifted.
  `Table.shiftColumns` asked neither question `shiftRows` asks, so a splice deleting a table's every
  column left the table alive, carrying the names of columns that no longer exist, and the writer
  emitted it. Clearing a line's last property deleted the key and left the record, so
  `getRow(500).height = 20` followed by `= undefined` kept `rowCount` at 500 forever. `Range` fetched
  by address, and fetching a covered address resolves it to the merge master, so a range over a merge
  reported the master twice and the covered cell never, while `clearStyle` cleared the master twice.
  And `addPivotTable` reached its source through a materialising accessor bounded by the used extent,
  so a source holding 22 cells plus one lone value far down a column grew to 150,000 cells, in
  189 ms, and they stayed on the sheet.

- **Eight write sites stood outside the escaping boundary, three of them re-emitting untrusted
  input.** `[Content_Types].xml` interpolated a part path and a content type raw, and a preserved part
  feeds both from the source package, so a crafted content type closed the attribute and injected a
  second `<Override>` for `xl/workbook.xml`, while a bare `&` alone made the part unparseable.
  `<sheetProtection>` re-emitted its agile-hash credential the same way, where its workbook-level
  counterpart already escaped correctly. Two preserved relationship targets were escaped twice, so
  `xl/slicers/s&1.xml` came back out as `Target="../slicers/s&amp;amp;1.xml"` and resolved to a part
  that is not in the package, silently orphaning the slicer. A totals-row function and a custom-filter
  operator were escaped instead of checked, which turns a bogus token into a well-formed document
  Excel still rejects rather than a refusal at the call. And a defined name whose scope names no sheet
  emitted `localSheetId="-1"`, an `xsd:unsignedInt` Excel offers to repair; that lookup also matched
  its scope exactly where `defineName` had validated it case-insensitively, so a scope accepted at
  authoring time could not be resolved at write time.

- **Adding a second pivot table to a workbook that already had one overwrote a part.** Both wanted
  `xl/pivotTables/pivotTable1.xml`, and preserved parts are emitted last, so the old pivot's bytes
  landed on the new one's path, the new pivot's sheet relationship pointed at data it was never built
  from, and `[Content_Types].xml` declared three PartNames twice, which violates OPC M2.5 and makes
  Excel offer to repair the file. The cause was a stale comment claiming the writer never generates a
  pivot table or its caches. The writer's part map is now a class that refuses a second part on a
  path, checked at each of the twenty write sites rather than trusted, and it carries no prototype,
  because a preserved part keeps its zip entry name from an untrusted package and one called
  `__proto__` was silently dropped while re-pointing the map's prototype at its bytes.

- **A streamed collapsed outline group rendered expanded when a cell in its summary row held the
  text ` collapsed="1"`.** The attribute is decided after the row is serialised, and the patcher was
  reading the rendered markup -- which also holds cell text -- to ask whether it was already there.

- **A theme override could be dropped, mis-spliced, or silently discard the `panose` metric beside
  the typeface it replaced.** The three edits were regular expressions over XML; they now splice at
  offsets a scanner found, so a theme with nothing overridden comes back byte for byte.

- **A number format carrying a tab, line feed or carriage return was written raw into an XML
  attribute**, where a conforming parser normalises all three to a space, so Excel read back a
  format code the author did not write.

- **A `.xlsb` whose `PtgNum` bytes decode to an infinity produced the formula text `INFINITY`**,
  which the writer then emitted as a package Excel reports as damaged; and a formula whose last
  token ran past the end of its own `rgce` aborted the entire workbook read rather than costing that
  one formula and keeping its cached result.

- **A corrupt package could raise `AuthoringError`, `SyntaxError` or `RangeError` out of
  `readXlsx`.** A sheet named twice, named nothing, named `a/b` or named at 32 characters is now
  repaired the way Excel repairs it; a table or defined name the model refuses is dropped along with
  its feature. The two native errors were the worse half: `catch (e) { if (e instanceof XlsxError) }`
  could not see them.

- **A character reference naming a code point XML 1.0 cannot carry (`&#1;`, `&#xD800;`) was decoded
  into the model and refused on the next write**, so a load-and-resave of a hostile file failed and
  blamed the caller.

- **`workbook.properties.created` set to an Invalid Date threw a bare `RangeError: Invalid time
  value`**, naming neither the property nor the document; a year outside 0000-9999 did not throw at
  all and wrote a timestamp no `dcterms:W3CDTF` admits.

- **The CSV writer accepted a delimiter the CSV reader refuses.** `{delimiter: '||'}` produced text
  this codec could not read back, and `{delimiter: ''}` quoted every field and emitted a file with
  no separators in it.

- **`TableStyleElement.size` set to `NaN` was written into an `xsd:unsignedInt`** as the four
  letters.


## [3.0.0] — 2026-08-30

Appending rows and reading a sheet are the two things every caller does, and both were
carrying a quadratic. Appending 16,000 rows in a loop went from 15 seconds to 82 ms, a
20,000-row sheet reads in 1.03 s instead of 1.75 s, and a workbook of 20,000 merges loads in
65 ms instead of 1.37 s. Not one observable value changes with any of them.

**One break, and it is small.** `duplicateRow` now copies the source row's own properties
(height, hidden flag, outline level, row fill) onto every copy, and `{insert: false}` replaces
the destination row rather than overlaying it. If you relied on a duplicate coming back at the
default height, or on a destination cell surviving in a column the source leaves empty, that
changes. Everything else here is a fix or an internal change.

### Changed

- **BREAKING: `duplicateRow` carries the source row's height, hidden flag, outline level and row
  fill onto every copy, and states that `{insert: false}` replaces the destination rather than
  overlaying it.** Both were undecided rather than decided: the doc promised "a faithful duplicate
  of the source's values and per-cell styles", which is narrower than what a caller duplicating a
  tall grouped row expects, and neither path copied the row's own properties, so a duplicated header
  came back at the default height. The overwrite path also wrote straight into the grid, which meant
  a destination cell in a column the source leaves empty survived in one copy mode and not the
  other. A copy is now the whole row in both modes: values, per-cell styles, and row properties,
  with a source that declares no properties clearing the destination's.

- **Appending is no longer quadratic in the number of appended lines.** `rowCount` and
  `columnCount` walked every cell on every read, and `addRow`, `addRows` and the streaming writer
  each read one of them once per line they append. Ten columns wide: 16,000 rows through `addRow`
  in a loop, 15,039 ms to 82 ms; the same rows through `addRows`, 85 ms to 66 ms; 8,000 rows
  streamed with `useSharedStrings`, 1,056 ms to 15 ms. The streamed case is the sharpest, because
  bounding memory on a large sheet is that writer's whole reason to exist. The extent is not
  cached: a caller holding a `Cell` can style or clear it without the sheet hearing about it, and
  a stale used range is worse than a slow one, since it is what lets an append land on a row
  someone had prepared. What is maintained is the structure the sheet does observe, confirmed
  against the live grid at a cost of one row.

- **Merged regions are indexed by row band rather than scanned.** `mergeCells` rejected an
  overlapping region by walking every region on the sheet, and the reader calls it once per
  `<mergeCell>` in the part, so loading a file's merges was quadratic on an untrusted path,
  reachable with a few megabytes of XML. 20,000 merges: 1,368 ms to 65 ms. `getCell` paid the
  same scan resolving a covered address to its region's master and now rides the same index.
  Memory is one entry per region, so a sheet of whole-column merges (legal, disjoint, cheap to
  write) cannot turn the index itself into the cost.

- **A worksheet part is parsed once, not five times.** Reading a sheet drove five full SAX passes
  over the same XML: the body, then hyperlinks, standard validations, extended validations and
  conditional formattings. A plain data sheet paid all five in full, because the cost is
  structural rather than proportional to what the scans find. A 20,000-row sheet with none of the
  four features present: 1,749 ms to 1,030 ms.

- **The VBA writer's mini stream stays as bytes.** Every sub-cutoff stream in a project (`dir`,
  `PROJECT`, `PROJECTwm`, and each module's compressed source) was spread into a `number[]` and
  converted back at layout time, so writing a real macro project built hundreds of thousands of
  boxed numbers to produce bytes the caller had already handed over as bytes. The spread goes
  with it, and with it an argument-count ceiling bounded only by the mini cutoff happening to be
  small.

- **The size budgets run under `pnpm run verify` and the pre-push hook, not only at publish
  time.** A budget the publish step alone checks is a surprise rather than a tripwire: the
  `/customui` entry spent a release 3 KB over its budget while every gate anyone actually ran
  stayed green. It is now the ninth gate of `verify --full`, and `/customui` is back under it at
  12.8 KB, after the XML scanner was split from the traversals that had accumulated beside it.

### Fixed

- **The VBA decompressor accumulates into a growable `Uint8Array` rather than a `number[]`, and
  enforces its ceiling on every write.** The bomb guard admits 64 MiB, and that many boxed array
  slots cost several times their byte count in real memory before the conversion to bytes, so a
  hostile `.xlsm` could amplify a small container into a much larger allocation than the bound
  suggests. The bound itself was always correct; the representation was not the one for it. Growth
  is now capped at the ceiling too, so the limit bounds the allocation rather than being checked
  after it.

- **A lookup table keyed by text taken out of a file no longer inherits a prototype.** Six tables
  in the tree were plain object literals indexed by an untrusted string, so about a dozen
  attacker-chosen keys resolved to a function off `Object.prototype`, at sites that all detect a
  miss with `??` or `=== undefined`. The reach was widest at entity decoding, which runs over
  every text node and every attribute value of every part of every package read: `&constructor;`
  decoded to the source text of `Object`, contradicting the documented promise that an
  unrecognised entity is left verbatim. `imageContentType('constructor')` returned a function into
  a `[Content_Types].xml` attribute, a `PROJECT` line reading `constructor=Foo` published `Object`
  as a VBA module kind, and a zip entry named `__proto__` re-pointed the inflated part map's
  prototype at its own bytes. Nothing was ever written through any of them, so this is the read
  side of that mistake rather than prototype pollution.

- **A macro-enabled workbook whose office document is not at `xl/workbook.xml` no longer reads as
  macro-free.** `editVbaProject` resolved the `vbaProject` relationship through a second resolver
  that prefixed `xl/` instead of resolving against the part that declared it, and that resolver
  never collapsed `..`. For a workbook one directory deeper the project key answered to no part in
  the package, so it was silently treated as absent, on the path whose whole purpose is not losing
  someone's macros. It now uses the resolver every other reader path already used.

- **A streaming write that fails destroys its streams instead of abandoning them.**
  `WorkbookStreamWriter.commit()` ran without a `try`/`finally`, so a throw from the package build
  (where a value OOXML cannot spell is refused, which happens routinely) rejected the promise and
  left every stream open forever. A caller following the documented `writer.stream.pipe(out)`
  idiom waited on a stream that would never end, and a caller-supplied sink was neither ended nor
  awaited, so an outer `await finished(sink)` hung on a writer that had already announced its
  failure. The failure now reaches every stream as a failure, by `destroy(err)` rather than
  `end()`, because ending would claim that a truncated archive is a complete package.

- **A row or column carrying only its own formatting counts toward the used range.** `rowCount`
  promised "the last row carrying anything (data or its own formatting)" and measured values
  alone, so a pre-formatted band, which is how anyone lays out a template, was invisible to it and
  `addRow` appended onto the styled row rather than below it. The row's styles survived, which is
  what made it quiet: the caller's row 2 and row 3 were now one row. `addColumn` failed identically
  on the other axis, and `usedRange` handed `autoFilter` a ref that under-covered the sheet it was
  meant to span. `actualRowCount` deliberately keeps the value-only reading, since how many rows
  hold data is a different question, and a cell materialised by `getCell` and left untouched still
  carries nothing, so reading a far address cannot grow the sheet.

- **An unregistered image id is refused once, with a message naming the sheet and the role.** Two
  places checked it, and the one that could actually fire carried the worse message: "a worksheet
  anchors image id 999", which named neither. It now reads `sheet "Sales" anchors image id 999,
  which is not registered on the workbook`, or `sets background image id 999` for the other role.

- **`StreamedSheetReader.merges` hands back a copy, as its sibling `hiddenColumns` already did.**
  Beyond the ownership question, `rows()` assigns a fresh array at the start of each iteration, so
  a caller who kept the returned array across a second pass was holding a detached snapshot of the
  first, with nothing to tell them so.

## [2.1.0] — 2026-08-28

The package root bundles for a browser now, and that is the change the rest of this release
orbits.

**Read the five breaks below before upgrading.** They land under a minor version number, which
is not what [ADR-0015](docs/decisions/0015-publishing-name-semver-and-first-version.md) §2 asks
for: each of them is a public API change and the policy calls for a major. None is a redesign,
and each is a place the library was accepting something it could not honour. But the version
number will not warn you the way it is supposed to, so this list has to.

- **`WorkbookStreamWriter`, `WorksheetStreamWriter`, `StreamedRow` and their options now come
  from `@shbernal/ts-xlsx/node`.** Importing them from the package root no longer resolves.
  That import is the only thing that has to change: the classes, their sink options and
  `writer.stream` are untouched.
- **`CsvWriteOptions.encoding` is `CsvEncoding`, not Node's `BufferEncoding`.** `base64` and
  `hex` are gone, having never been output encodings for a text format. The bytes written for
  every remaining spelling are unchanged.
- **An out-of-enumeration token is refused at the write and dropped at the read.** Page setup,
  data validation, conditional formatting, sheet and window visibility, and two-cell image
  anchors now throw `AuthoringError` naming the offending value, where such a value used to be
  interpolated raw into the document.
- **`ConditionalFormattingRule`'s `type`, `operator`, `timePeriod` and `iconSet` are closed
  unions,** not `string`. A rule type the library does not model in depth still round-trips.
- **A `NaN` or an infinity is refused wherever it can reach an OOXML numeric attribute,**
  instead of producing a package Excel reports as damaged.

Everything else here is a fix or an internal change. The sections below carry the detail.

### Changed

- **BREAKING: the package entry no longer reaches a Node built-in, and the streaming writer moved
  to `@shbernal/ts-xlsx/node`.** Three Node modules used to sit on the graph reachable from
  `src/index.ts`, so a browser build pulled in `node:crypto`, `node:fs` and `node:stream` whether
  or not the page ever called them: a bundler resolves imports, not call graphs. Webpack reported a
  module it could not resolve and Vite externalised them with a warning, and this repository's own
  site could only bundle the library by aliasing all three to a stub. `WorkbookStreamWriter`,
  `WorksheetStreamWriter`, `StreamedRow` and their options are now published from
  `@shbernal/ts-xlsx/node` and nowhere else, which is the only import that has to change; the class
  itself, its sink options and `writer.stream` are untouched. Every other entry, the root specifier
  included, now bundles for a browser with nothing to configure, and
  `scripts/check-browser-safe.ts` walks the module graph on every run to keep it that way
  ([ADR-0040](docs/decisions/0040-the-browser-boundary-is-an-entry-point.md)).
- **BREAKING: `CsvWriteOptions.encoding` is `CsvEncoding`, not Node's `BufferEncoding`.** The new
  exported union names the eight spellings a CSV consumer actually asks for (`utf8`, `utf-8`,
  `utf16le`, `utf-16le`, `ucs2`, `ucs-2`, `latin1`, `ascii`) and drops `base64` and `hex`, which
  were never output encodings for a text format. The bytes are unchanged: a test asserts each
  encoding byte-for-byte against what `Buffer.from` produced. The public API no longer depends on
  `@types/node` to state this option's type.
- **Sheet-protection passwords are hashed without `node:crypto`,** by this library's own SHA-512
  (`src/sha512.ts`), with the salt from the platform's `crypto.getRandomValues`. Credentials are
  identical: the same algorithm, the same salt length, the same default 100000 spins, checked
  against Node's implementation in tests. `Worksheet.protect(password)` therefore works in a
  browser, and costs about 0.8 s at the default spin count against `node:crypto`'s 0.4 s.
- **`readCsv` and `writeCsv` no longer touch `Buffer`,** using `TextDecoder`/`TextEncoder` and two
  small encoders instead, so the CSV functions run in a tab as the documentation always said.
- **BREAKING: an out-of-enumeration token is refused at the write and dropped at the read.** A
  `<pageSetup>` orientation or page order, a `<dataValidation>` type/operator/error style, a
  `<cfRule>` type/operator/time period/icon set/anchor type, a sheet tab or document window
  visibility, and a two-cell image anchor's edit mode each go through the same check the style
  facets have always had: a value outside the enumeration throws `AuthoringError` naming the value.
  Such a value was previously interpolated raw, so a caller reaching past the types with untyped
  JavaScript, a `JSON.parse`, or a cast could close the attribute and the element behind it and
  choose the rest of the document. Reading is the mirror: a foreign token in a file is dropped
  rather than stored, so what the reader accepts is always something the writer can write back.
- **BREAKING: `ConditionalFormattingRule`'s four enumerated fields are unions, not `string`.**
  `type`, `operator`, `timePeriod` and `iconSet` now have the closed types ECMA-376 gives them
  (`ConditionalFormattingType`, `ConditionalFormattingOperator`, `CfTimePeriod`, `IconSetType`, all
  exported). A rule type the library does not model in depth still round-trips, exactly as before;
  a token no Excel would open no longer does. A file carrying one loses that rule on read rather
  than carrying it to a write that would refuse it.
- **BREAKING: a number that OOXML cannot spell is now refused at the write, everywhere.** A
  `NaN` or an infinity reaching `pageSetup.scale`/`fitToWidth`/`fitToHeight`/`paperSize`, a page
  break's `id`/`max`, a `Color`'s `theme`/`tint`/`indexed`, a conditional-formatting rule's
  `rank`/`stdDev`, an image anchor's extent, rotation or grid point, the workbook window rect, or a
  row or column `outlineLevel` throws `AuthoringError` instead of writing `NaN` into an
  `xsd:double` or `xsd:unsignedInt` attribute. Those calls used to "succeed" and produce a package
  Excel reports as damaged, so the failure moved to the moment the mistake was made. `numberText`
  already took this stance for a font size; the rest of the write path now shares it, and the
  shared numeric-attribute helper `attr` is renamed `numAttr` to sit beside `boolAttr`.

- **`DEFAULT_THEME_COLOR_SCHEME` is typed as a complete record, not a partial one.** It declares
  every `ThemeColorSlot` and always did, so indexing it no longer yields `string | undefined` and no
  longer needs an assertion at the use site. Nothing about the value changed.

### Fixed

- **A row or column insert no longer pushes a region off the edge of the grid.** A whole column is
  written bounded, `B1:B1048576`, which is the spelling Excel itself writes, so a validation, a
  conditional format, an autofilter or a merge anchored to the last row had its bottom edge shifted
  past the end of the grid by an insert above it: `sqref="B1:B1048578"`, naming rows that cannot
  exist. `OpenXmlValidator` passes such a package; Excel meets it with the "we found a problem with
  some content" repair prompt, where the same workbook with its edges inside the grid opens clean.
  Every re-anchoring pass now clamps to the axis's last line, which is what Excel does to the same
  region on the same edit. See
  `docs/knowledge/specs/a-splice-must-not-push-geometry-off-the-grid.md`.

- **A streamed `getCell` into an already-committed row is refused instead of emitting a duplicate
  row.** An eager `WorksheetStreamWriter` renders a committed row and releases its cells; addressing
  a cell in that row re-materialised it in the model, and the commit then wrote a second `<row>`
  carrying the same number, leaving the sheet saying two contradictory things about one row. It now
  throws `AuthoringError` at the call, in the same shape as the shared-formula refusal beside it.
  Rows never committed stay writable as before.

- **A self-closing `<t/>`, `<text/>`, `<xm:f/>` or `<totalsRowFormula/>` can no longer leave a text
  capture latched.** Nine parsers gathered an element's character data across the open/text/close
  events without honouring the documented fact that a self-closing element fires no close. Nothing
  was observably wrong: the next open happened to clear the buffer before anything read it. That was
  the order of two resets, not a property anyone chose, on a path that reads untrusted input. The
  capture is now one shared state machine that takes `selfClosing`, so it is structural. It also
  scopes each capture to the element that opened it, so a nested element inside a document property
  no longer ends that property's text early.

- **A conditional-formatting block with no rules is omitted rather than emitted empty.**
  `CT_ConditionalFormatting` requires at least one `<cfRule>`, so `addConditionalFormatting({ref,
  rules: []})` used to write a schema-invalid element.
- **A non-finite row `outlineLevel` no longer hangs the writer.** The scan that derives which
  summary rows terminate a fully-collapsed group walks outward comparing outline levels; against
  `-Infinity` every comparison held, so the walk ran off the sheet and never returned. It is now
  refused with the rest of its family, before the walk starts.

## [2.0.0] — 2026-08-27

A major, because four documented breaks land together. None of them is a redesign: each is a
place where the library was quietly doing the wrong thing and the honest fix is visible to a
caller.

- **Seven scalar-validation throws became native errors.** `catch (e) { if (e instanceof
  XlsxError) }` no longer catches them. See *Changed*, first entry, for the table.
- **`Workbook.authoredThemeXml()` is now the `Workbook.themeOverrides` getter,** and
  `parseThemeColorScheme` moved from the `/core` subpath to `/xlsx`. The package root is
  unaffected either way.
- **A row or column addressed past the grid now throws `RangeError`** where it used to return,
  matching what the letter spelling has always done. Reading such a file is unaffected.
- **A character XML cannot represent now throws when written anywhere but a cell value.** It
  used to return a package Excel reports as damaged.

Two more changes are additive but worth reading before upgrading: `WorksheetModel` gained
`view`, which only affects code that builds a model literal by hand, and the `_xHHHH_` escape is
now decoded on read, so a foreign file holding one reads back as the character rather than as
the seven-character literal.

### Added

- **`GridRect` and `MentionRef`: the two shapes the library was declaring more than once.**
  `GridRect` is the four inclusive 1-based bounds every range-shaped thing in the library carries;
  `MergeRect` and `TableRegion` are now aliases of it and `Range` is declared to satisfy it, so a
  function written over one rectangle takes all of them. `MentionRef` is an `@mention` as a file
  spells it, and `Mention` is that plus the identity the id resolved to. Both are exported from
  `/core`. No existing type changed shape.

- **`Workbook.exportImages(sheet)` / `Workbook.importImages(sheet, images)` — carry a sheet's
  pictures to another workbook.** An anchored image holds a media *id* into one workbook's registry,
  and that id names a different picture, or none, in the next. Copying a sheet has therefore always
  left its images behind: `dst.model = src.model` is a semantic copy by design
  ([ADR-0005](docs/decisions/0005-worksheet-model-is-semantic-only.md)), and there was no affordance
  for the bytes. Now there is one, and carrying a sheet whole is the two together:

  ```ts
  destination.model = source.model;
  destinationWorkbook.importImages(destination, sourceWorkbook.exportImages(source));
  ```

  `exportImages` resolves a sheet's anchored images and its background out of the media registry into
  a workbook-independent `WorksheetImages` value; `importImages` registers those bytes in the
  destination and re-anchors against the ids they land on. Registration is content-addressed, so a
  picture already held is re-used rather than stored twice — twenty sheets sharing one logo cost one
  media part. An import replaces the destination sheet's pictures rather than appending to them, the
  same direction a `model` assignment goes.

  `exportImages` throws `AuthoringError` for a sheet whose image ids this workbook never registered —
  which is what a sheet belonging to *another* workbook looks like from here. That pairs with the
  existing write-time check, so an anchor holding a foreign media id is refused by name at both ends
  instead of being emitted as a drawing relationship pointing at media that was never written.

  This covers floating anchors and the sheet background. Header/footer images are byte-preserved
  parts and ride the preserved-reference machinery; charts, vector drawings, pivots and slicers stay
  round-trip-only ([ADR-0014](docs/decisions/0014-charts-shapes-slicers-are-round-trip-only-for-1-0.md)) —
  their bytes carry the source sheet's identity (a pivot table's workbook-unique name, a slicer
  cache's `tabId`, a chart series' `Sheet1!$A$1`), so they cannot be relocated by copying bytes.

### Fixed

- **A malformed frozen-pane split no longer surfaces from the writer.** A `<pane>` whose `xSplit`
  or `ySplit` spelled a fraction, a negative, or a word was stored verbatim, though `freeze()`
  refuses all three. The read succeeded and the *serializer* then threw a `RangeError` naming a
  column the file never mentioned. A split that is not a non-negative integer is now dropped where
  it is read, leaving the rest of the sheet view intact.

- **A `<col>` span wider than the sheet no longer hangs the reader.** `<col min="1"
  max="99999999"/>` is one line of XML that named more columns than the format has, and the
  buffered reader walked it verbatim: roughly 16.7 million column records, twenty-five seconds,
  and then `RangeError: Map maximum size exceeded`. A denial of service on a one-line input.
  The span is now clamped to the last real column and an element wholly outside the grid is
  dropped, which is what the streaming reader already did with the same element. A `<row r>`
  past the last row is dropped for the same reason, in both readers: the streaming reader used
  to hand the out-of-grid number straight to the consumer as `StreamedRow.number`, so a data
  read could be given an address no `getCell` will accept. It now drops the row, as the
  buffered reader does. An `<r>` names one row and nothing can be folded onto it, so it is
  dropped rather than clamped: clamping would move a foreign file's formatting onto the last
  real row.

- **A lone surrogate in a CSV field is refused instead of silently becoming U+FFFD.**
  `writeCsv`'s UTF-8 encode substituted the replacement character for an unpaired surrogate, so
  the file was written, nothing failed, and the character was gone. A CSV field is plain text
  with no escape convention to hide it in, unlike a cell value's `_xHHHH_`, so refusing at the
  encode step is the only honest answer: it now throws, naming the code point and its offset.
  `writeCsvText` still returns the surrogate intact, because a JS string carries it losslessly,
  and the UTF-16 path still encodes it through.

  One spelling inconsistency is settled alongside it: `encoding: 'utf-8'` used to take the
  non-UTF-8 branch and get no byte-order mark, unlike `'utf8'`. Both spellings now behave the
  same, BOM and refusal alike.

- **An authored theme typeface reads back decoded, matching what the writer escaped.**
  `parseThemeFontScheme` handed back the raw `<a:latin typeface>` attribute text, so a face
  carrying an `&` came back as `&amp;`: authoring `Ampersand & Co` and reading it back yielded
  `Ampersand &amp; Co`, and a second write would have doubled the escape. The reader was the
  half that disagreed with the writer, and it now decodes like every other attribute reaching
  the model. The colour scheme beside it was never exposed, because a hex slot cannot carry an
  entity; the font scheme is the one place in the part where free text lives.

- **A control character in a string no longer produces a package Excel reports as damaged.**
  The writer escaped `& < > " '` and nothing else, so every other character reached the file
  byte for byte — including the C0 controls XML 1.0 forbids outright and the U+FFFE/U+FFFF
  noncharacters. A cell value holding U+0001 emitted a sheet part Microsoft's
  `OpenXmlValidator` rejects. Strings arriving from a database column, a CSV field, or a user
  form are exactly where these turn up.

  A **cell value** is now escaped with SpreadsheetML's `_xHHHH_` convention, which is what
  Excel writes and what Excel reads back, so the character survives the round-trip instead of
  breaking the file. That covers inline strings, pooled strings, rich-text runs, legacy note
  text, and the cached result of a string formula.

  A lone surrogate rode along in the same fix. It never produced an invalid package — the
  UTF-8 encoder substituted U+FFFD for it, so the file validated and the character was
  already gone. It is now escaped and preserved like the rest.

  Because `_xHHHH_` now carries meaning, a cell value that legitimately contains the literal
  text `_x0041_` has its underscore escaped as `_x005F_x0041_`, or it would read back as the
  letter `A`. Text that only resembles an escape (`_`, `_x`, `_xZZZZ_`) is untouched.

  The reader undoes the same convention, which is what makes the round-trip claim above true
  and also changes how *foreign* files read. A workbook Excel authored with a control character
  in a cell previously read back as the literal seven-character text `_x0001_`; it now reads
  back as the character. That applies to inline strings, the shared-strings pool, rich-text
  runs, legacy note text, and the cached result of a string formula, in both the buffered and
  the streaming reader.

  The decode is one left-to-right pass, so `_x005F_x0041_` reads as the literal `_x0041_`
  rather than collapsing to `A`, and the grammar is exact: `_x041_`, `_x00041_` and `_xZZZZ_`
  are ordinary text and stay that way. Every one of those outcomes was measured against Excel
  Desktop before it was implemented, and the probe workbook is committed as a corpus fixture
  (`docs/knowledge/specs/spreadsheetml-xhhhh-escape-is-decoded-on-read.md`).

  Two more carriers of human-typed prose have since been measured and joined the group: a
  **threaded comment's message** and a **print header/footer definition**. Both had been guessed
  the other way and refused such a character on write. Excel decodes and re-emits the escape in
  each, verbatim to the cell-text grammar, so both now escape on write and decode on read. Two
  consequences for callers: a header or a threaded message may carry any character a cell may,
  and a foreign file whose header reads `&C_x0041_` now loads as `&CA` rather than as the
  seven-character literal.

- **A frozen pane no longer disappears when a sheet is copied through `model`.**
  `WorksheetModel` was missing `view`, so `dst.model = src.model` reproduced the cells,
  merges, tables, autofilter and page setup — and silently unfroze the header row. The
  frozen/split pane now rides the model like every other sheet-level field, and clearing it
  works in the same direction: assigning a normal-view model over a frozen sheet unfreezes it
  rather than leaving a stale pane.

  This is additive to `WorksheetModel`. Code that *constructs* a model literal by hand rather
  than reading one from a sheet must now supply `view` (`{}` is a normal view); code that does
  the usual `dst.model = src.model` is unaffected.

  The boundary is now a rule rather than a list
  ([ADR-0005 amendment](docs/decisions/0005-worksheet-model-is-semantic-only.md)): **a field
  belongs in the model when its value means the same thing on any sheet of any workbook.** By
  that test threaded comments stay out — a comment's author is an id into the *workbook's*
  `persons` registry — and they are now named on the out-of-scope side instead of being absent
  from both lists.

- **A foreign file spelling an OOXML boolean `"false"` is read correctly.** `xsd:boolean`
  permits `true`/`false` alongside `1`/`0`. Excel writes the digit, so five readers that
  tested the attribute by hand against `'0'` had never been caught reading the long spelling
  as *true*: a `<col customWidth="false">` gained a width it does not have, a
  `<row customHeight="false">` a height, an x14 data bar's `gradient="false"` stayed a
  gradient, and a `cfRule aboveAverage="false"` read as above-average. All five now go
  through the reader's existing `boolPresent`, which has handled both spellings all along.
  Files written by Excel are unaffected; files from other producers are the point.

### Changed

- **Six scalar-validation sites now throw native errors, as `errors.ts` says they should.** The
  taxonomy draws the line explicitly: one argument out of range, unparseable, or the wrong type is
  a native `RangeError`/`SyntaxError`/`TypeError`, and `AuthoringError` starts where a *composite*
  is inconsistent. Six sites were on the wrong side of it, so `catch (e) { if (e instanceof
  XlsxError) }` caught some argument mistakes and not others, with no rule a caller could predict.

  | What | Was | Now |
  | --- | --- | --- |
  | `Workbook.setDefaultFont` size / empty name | `AuthoringError` | `RangeError` |
  | `Workbook.addTableStyle` band `size` | `AuthoringError` | `RangeError` |
  | `readCsv` `delimiter` length | `AuthoringError` | `RangeError` |
  | A table name's length | `AuthoringError` | `RangeError` |
  | A table name that is not an Excel identifier | `AuthoringError` | `SyntaxError` |
  | A theme colour that is not `RRGGBB` | `AuthoringError` | `SyntaxError` |
  | An ARGB colour that is not 6 or 8 hex digits | `AuthoringError` | `SyntaxError` |

  Every message string is unchanged, and no other throw moved: the composite claims stay where they
  were (a table style with no name, an element carrying a `size` it cannot have, a table whose
  columns do not span its range). Code branching on the message or on `instanceof Error` is
  unaffected; code branching on `XlsxError` for these seven needs the native type instead.

- **`Workbook.authoredThemeXml()` is now `Workbook.themeOverrides`.** The old method handed back
  theme part *text*, which made the model the place that knew how a theme is spelled. The getter
  returns the colour slots and typefaces {@link setTheme} authored, or `undefined` when none were,
  and the serializer composes them onto the part it is about to write. Same result in the file;
  a caller who needs the text can compose it with `applyThemeOverrides`.

- **`parseThemeColorScheme` moved from `/core` to `/xlsx`,** which is where the part it parses is
  read, and is joined there by `parseThemeFontScheme` and `DEFAULT_THEME_XML`. Importing from the
  package root is unaffected; a subpath import needs the new one.

- **An authored theme typeface is escaped properly.** The theme writer carried its own attribute
  escape, weaker than the library's: it handled `& < > "` and left the apostrophe, the newline, the
  carriage return and the tab raw, and it did not refuse a character XML 1.0 cannot represent. A
  typeface carrying any of those produced a malformed part. It now uses the same escape as every
  other attribute in the package. A custom number format code, whose escape *deliberately* differs
  (a bare apostrophe round-trips), gained the representability guard it was also missing.

- **A row or column addressed by number is now bounded by the grid, as one addressed by letters
  already was.** `getCell('XFE1')` threw and `getColumn(16385)` did not, so whether the library
  refused a position outside the spreadsheet grid depended on how the caller spelled it, and the
  numeric spelling let the model hold, and the writer emit, a package Excel must repair. `Row`,
  `Column` and `Cell` now reject an index outside `1..1048576` / `1..16384` with a `RangeError`
  worded the same way the letter path words it.

  This is a new throw on calls that used to return. Code placing content at a synthetic index past
  the grid was producing a file Excel rejects, so the failure has moved to the call that causes it.
  Reading is unaffected: a file declaring an out-of-grid position is bounded on read rather than
  refused, so no file that opened before stops opening.

- **Error messages no longer contain em dashes.** Where a message used ` — ` to weld two
  clauses together it now uses a colon, a semicolon, or a pair of commas, whichever the
  sentence wanted: `column 16385 is out of bounds: Excel supports 1..16384` rather than
  `column 16385 is out of bounds — Excel supports 1..16384`. No message changed meaning, none grew,
  and the error *types* and `code` values are untouched, so anything branching on the
  taxonomy is unaffected. Code that matched on message text is not: match on the class or
  on `code` instead, which is what they are for.

- **Writing a character XML cannot represent into anything but a cell value now throws
  `AuthoringError`.** Only cell values have the `_xHHHH_` escape; a sheet name, a defined
  name, a table column name, a formula, a document property, a print header and a threaded
  comment do not. There is no faithful way to write U+0001 into a sheet tab — `Sheet_x0001_A`
  is a different name, not the one that was asked for — so the writer refuses, naming the
  code point and its offset.

  This is a new throw on a call that used to return. It used to return a file Excel reports
  as damaged, so the failure has moved earlier and got louder rather than appearing from
  nowhere. Tab, LF, CR and U+007F are valid XML 1.0 and are unaffected; so are the C1
  controls, which only XML 1.1 forbids.

- **The OOXML conformance oracle is now the shared `ooxml-validate` package**, and the
  repository-owned .NET tool behind it is gone
  ([ADR-0033](docs/decisions/0033-the-ooxml-oracle-is-a-shared-package.md)). Nothing in
  `src/` moved and nothing a consumer installs changed — this is development and CI
  tooling, as it always was.

  The point is that `ts-pptx` used to carry a *different* validator on a different Open XML
  SDK version, so two sibling projects were enforcing two rule sets while both calling it
  "Microsoft's validator". They now share one oracle, one pin, and one report contract.

  For contributors: **there is no .NET requirement any more.** The package fetches a
  prebuilt, checksum- and provenance-verified binary on first use and caches it, so
  `pnpm run validate:ooxml file.xlsx` and `pnpm run test:ooxml` work from a plain dev
  install. `global.json` and CI's `setup-dotnet` step are gone. Both spellings work —
  `pnpm run validate:ooxml file.xlsx` and `pnpm run validate:ooxml -- file.xlsx` — which is
  why the dependency floor is `ooxml-validate` 0.0.3.

  The baseline did not move. This repo was already validating at `Microsoft365` against
  `DocumentFormat.OpenXml` 3.5.1, so `test/ooxml-validation/allowed-errors.json` is
  unchanged and still empty.

- **A truncated VBA project fails closed instead of reading as zeros.** The `vbaProject.bin`
  reader's little-endian integer reads were open-coded in four modules, and the shared shape
  indexed past the end of the buffer without noticing: `undefined | (undefined << 8)` is `0`,
  so a truncated compound file could parse as a structure full of valid-looking zeros. Every
  caller happened to bounds-check first, so no known file was misread — but that is the wrong
  place for the check in the one subsystem that parses a blob straight out of an untrusted
  `.xlsm`. The bound now belongs to the read itself, which raises `VbaParseError` naming the
  offset and the buffer length.

## [1.3.1] — 2026-08-11

`1.3.0` was set in `package.json` and cut no release, so this is the first published version of
the line and carries what was staged for it as well as the fix below.

### Added

- **`SheetView.showGridLines` turns the on-screen grid off**, the one sheet-view attribute an
  authoring consumer reaches for first and the only facet of `<sheetView>` this library could not
  express. A workbook built as a *deliverable* — its own fills, its own borders, a title band —
  reads as a spreadsheet rather than a document while Excel's grey grid shows through it, and
  there was no way to say so:
  [`PrintOptions.gridLines`](docs/api/page-setup.md) is the neighbouring question about *printing*
  and Excel exposes the two as separate checkboxes because the answers differ.

  ```js
  sheet.view.showGridLines = false;
  ```

  Only an explicit `false` writes an attribute, and reading one back records only
  `showGridLines="0"`. Excel's default is on, so "unset" and "on" are the same state: recording
  `true` would make every re-written sheet fabricate an attribute its source never carried, which
  is the round-trip noise `<pane>` is already careful to avoid. Both arms of `sheetViewsXml` carry
  it, so a frozen sheet can hide its grid too.

  Found by the library's first real authoring consumer, which hides the grid on all seven sheets
  of the two workbooks it ships. The remaining `<sheetView>` booleans (`showRowColHeaders`,
  `showRuler`, `showFormulas`, `showZeros`, `rightToLeft`) are still unmodelled — see
  `docs/knowledge/specs/sheetview-boolean-flags-and-showformulas.md`, which called this set out and
  is now one item shorter.

- **`WorkbookProperties.title` and `.company`**, the two remaining fields of Excel's File ▸ Info
  panel a deliverable actually sets. Both were silently unwritable: a caller could assign neither,
  and a file that carried them lost them on a round-trip.

  ```js
  wb.properties.title = 'Planning Ateliers';
  wb.properties.company = 'Acme & Co';
  ```

  `title` is `dc:title` in `docProps/core.xml`, emitted **before** `dc:creator` because
  `cp:coreProperties` is a schema *sequence* and Excel repairs a file whose children are out of
  order. `company` is `<Company>` in `docProps/app.xml` — the one document property OOXML keeps
  outside the core part, which is why `appPropsXml` now takes the properties at all. Both are
  omitted entirely when unset, and both now read back.

### Fixed

- **Writing a workbook is now reproducible: the same model always produces the same bytes.** A zip
  entry stores a modification time, `fflate` defaults it to `Date.now()`, and so every package this
  library wrote was stamped with the moment it ran. Two writes of an unchanged workbook differed in
  a handful of bytes per entry and in nothing else — invisible to Excel, and corrosive to everything
  around the file. A committed `.xlsx` deliverable churned in `git diff` on every regeneration, so
  the diff stopped being read and a real change to it had somewhere to hide; nothing downstream
  could cache on output bytes; and a byte-comparison gate could not tell "the writer changed" from
  "the clock moved", which is the same as not having one.

  Every entry is now stamped 2001-01-01 12:00 — see
  [ADR 0032](docs/decisions/0032-package-output-is-reproducible.md) for why that date and why there
  is no option to go back to the clock. All four writing paths are covered: `writeXlsx`,
  `writeXlsxAsync`, `WorkbookStreamWriter`, and the package-level VBA edits, which re-zip after
  splicing and had nothing to preserve because unzipping drops the original times. The streamed
  container takes the stamp as a field on the entry rather than as an option, which is exactly the
  kind of difference that leaves one call site behind — hence a test that decodes the DOS date out
  of every entry rather than only comparing two writes, which would pass within the same
  two-second bucket regardless.

  Consequences worth stating plainly: package bytes changed once, at this release, so a stored hash
  of a previously written file no longer matches — nothing about the content did. `writeXlsx` and
  `writeXlsxAsync` are now byte-identical, which pins them to the same compression settings in a way
  comparing inflated parts never could. And a file manager listing the archive's contents shows
  2001-01-01 against each entry, the same trade every reproducible-build toolchain makes; the file's
  own filesystem timestamp is untouched.

  Found by the library's first authoring consumer, which commits its generated workbooks.
  [ADR 0024](docs/decisions/0024-async-is-one-writer-not-a-mirrored-pair.md) had met this while
  comparing the two writers, declined to pin `mtime` for a test's convenience, and left the real
  question open; a consumer answered it.

## [1.2.0] — 2026-08-09

No exported symbol changes in this release. What changes is what the tarball contains and what
one error says: the package now ships a skill for reporting a defect back here, because the
agent that hits one is working in someone else's repository and has no reason to believe filing
is its job. Alongside it, the published tarball is 42% smaller — the same code with 438 KB of
implementation comments no longer addressed to anyone.

### Added

- **The package ships a skill for reporting a bug upstream, and `InternalError` prints where to
  send one.** Most code using this library is written by an agent, and an agent that hits a
  library defect writes a workaround: that is the rational move from where it stands, since the
  workaround unblocks its user today and the issue tracker belongs to a repository it is not in.
  So the report never happens, the corpus never gets the case, and the next consumer rediscovers
  the same defect. Three layers answer three different reasons the report dies.
  `skills/ts-xlsx-upstream/` is now in `files`, so `npx skills add
  ./node_modules/@shbernal/ts-xlsx` works offline and always matches the installed version. The
  skill covers triage (is this ours or your file's?) and spends most of its length on reducing a
  failure to a script that *builds its own input* — a spreadsheet in a real project holds
  salaries and customer lists and the tracker is public, so the rule is synthesize or describe in
  prose, never redact. It files without interrupting you once the reproduction stands on its own,
  and passes `--repo shbernal/ts-xlsx` on every `gh` call, because `gh` in a consumer's checkout
  defaults to the consumer's tracker and would file our bug there.
- **`InternalError` messages now carry the tracker URL**, appended by the constructor rather than
  by its six throw sites, so a site added later cannot forget it. `'internal'` is the one code in
  the taxonomy that already declares whose bug it is, which makes an unconditional "report this"
  a statement of fact rather than a nag — and a stack trace is the only artefact of that failure
  that reaches whoever is debugging it, so a pointer anywhere else is one they must already be
  looking for. The other three codes deliberately get no banner: malformed input is the routine
  outcome for a library reading untrusted files, and a report pointer on every corrupt file
  teaches the reader to skip the line, including the one time it always means something. They get
  the test in prose instead, on the `XlsxErrorCode` TSDoc and therefore in `dist/index.d.ts` —
  report it when the file opens cleanly in Excel but not here, because that combination means the
  gap is ours.

### Changed

- **The published tarball is 42% smaller, and nothing was removed to make it so.** `build` was
  one `tsc` pass emitting both halves of `dist/`, and it kept comments — necessarily, because
  that is what puts JSDoc into the `.d.ts` hovers. The cost was that the same flag also shipped
  every implementation comment in the tree into the runtime JS: measured, 47% of it, 438 KB
  addressed to nobody. `removeComments` is whole-emit, so a single pass could only ever give up
  one audience or the other. The build is now two passes that split by audience —
  `tsconfig.build.json` emits JS with the prose stripped, `tsconfig.build.dts.json` emits
  declarations with it kept. Runtime JS falls from 918 KB to 480 KB and the tarball from 392 KB
  to 227 KB gzipped, while all 112 `.d.ts` files are byte-identical to the single-pass output —
  every one of the 237 KB of JSDoc across 90 files still reaches an editor hover. No module left
  any entry's closure, no signature changed, and the corpus's 838 behaviours run green against
  the emitted `dist/`.
- **The per-entry size budgets are halved, which makes them stricter rather than looser.** They
  were measured against JS that was ~47% comment, so a codec crossing a module boundary — the
  failure the tripwire exists to catch — could have arrived inside a release's ordinary comment
  churn without moving them. Rebaselined onto comment-free emit, they measure code.
- **The `npm-publish` environment is now checked rather than provisioned, and this repository
  holds no credential at all.** `environment.yml` wrote the environment through an
  administration endpoint, which `GITHUB_TOKEN` may not reach, so it needed a PAT —
  `ENV_ADMIN_TOKEN`. That secret was never created: the workflow failed on its own first step
  every day of its life, and the reviewer removal it was supposed to apply was done by hand
  instead. Restoring the secret would have put a long-lived credential able to rewrite the gate
  that decides who may publish next to a publish path that deliberately carries none. Reading
  an environment needs only read access, so the write was dropped: the workflow now fails when
  the live environment differs from what the file declares — weekly, on any edit to the
  declaration, and on demand. Drift is detected rather than corrected, which is the half that
  was load-bearing. Nothing about the published package changes. See
  [ADR-0026](docs/decisions/0026-releasing-is-a-github-release-and-npm-follows.md).

## [1.1.0] — 2026-08-08

The first release since 1.0.0 to add API surface. Six additions, each of them something a
caller was otherwise assembling by hand: a worksheet lookup that throws instead of returning
`undefined`, the sheet's extent as a single range, one plain-text rendering of a cell value,
the grid's measured geometry ceilings, a wrapped-line count for authoring a row height, and
the `CellValue` type guards — which the reader and writer had always used but the barrel never
exported. Nothing is removed and no existing signature changes. The CSV writer now shares the
value renderer rather than keeping a near-copy of it, which shifts two of its behaviours; both
are recorded below.

### Added

- **`Workbook.requireWorksheet` — the total counterpart of `getWorksheet`.** Same lookup, by name
  (case-insensitive) or numeric id, but a miss throws an `AuthoringError` naming every sheet the
  workbook does have instead of returning `undefined` that flows on into a `?.` chain and fails
  several steps later with nothing left to say. A lookup miss is a typo, a stale template or a
  renamed tab, and all three are answered by seeing the real names. The CSV writer now uses it, so
  its "no worksheet named …" message gains that listing and loses the "to write as CSV" suffix.
- **`estimateWrappedLines` — how many lines a wrapped cell takes.** A row that states no height
  records no geometry, so its height is whoever opens the file's answer. Excel's answer is an
  auto-fit computed on open — measured, and prompt: it saturates at `MAX_ROW_HEIGHT`, so past ~28
  lines it stops answering, and a consumer that does not implement it (this library's reader
  included) has no height at all. Writing one settles the geometry, and needs a line count. Counts
  characters against a character-unit width — the same unit a column's `width` is in, so no font
  metric is assumed and the deferred metric-table question stays deferred. Exact for a monospaced
  face that wraps mid-word, and against Excel a shade low (5 lines where Excel laid out 6), which
  the doc states. A hard break opens a line of its own; a zero or non-finite width is a `RangeError`
  rather than an `Infinity` that would land in a row height. The auto-fit measurements are in
  [`docs/knowledge/specs/rows-with-no-stated-height-are-autofitted-on-open.md`](docs/knowledge/specs/rows-with-no-stated-height-are-autofitted-on-open.md).
- **`MAX_ROW_HEIGHT` and `MAX_COLUMN_WIDTH` — the grid's geometry limits, measured.** The
  companions to `MAX_ROW`/`MAX_COLUMN`: how large a line may be set, where those two bound where
  a cell may be. `ht` and `width` are a bare `xsd:double` in the schema, so the ceiling is Excel's
  own, and it is not the one Microsoft's specifications table publishes — Excel Desktop accepts a
  row height of 409.5 and refuses 409.6, against a documented "409 points". Column width is exactly
  255. Both were measured over COM rather than quoted; the probe and its numbers are in
  [`docs/knowledge/specs/grid-geometry-limits-are-excels-not-the-schemas.md`](docs/knowledge/specs/grid-geometry-limits-are-excels-not-the-schemas.md).
  They bound *assignment*, not a file, and the two behave differently once one exceeds them: Excel
  opens an over-limit package clean, silently clamping a row to 409.6 (a tick above what it lets
  you set) while honouring a `width` of 1000 and re-saving it verbatim. Nothing enforces them at
  either end, therefore — `Row.height` and `Column.width` are the reader's path into a foreign
  file as well as an author's, so a bound that threw would refuse files Excel opens. There is
  deliberately no `DEFAULT_COLUMN_WIDTH` beside them: the default width follows the workbook's
  default font (8.43 for Calibri 11, 8.09 for Aptos Narrow 11), so a constant would be a wrong
  answer wearing a right one's name.
- **`Worksheet.usedRange` — the sheet's extent as a handle.** `rowCount` and `columnCount` said
  once: `A1` through the last row and column carrying anything, or `undefined` when the sheet spans
  no rectangle. It replaces the `` `A1:${numberToColumn(sheet.columnCount)}${sheet.rowCount}` ``
  every caller was assembling by hand, and it is what an auto-filter over a whole sheet wants —
  `sheet.autoFilter = sheet.usedRange.address`, where a header-only ref yields dropdowns that
  filter nothing. Anchored at `A1` and inheriting both counts' definition of *used*, so it is not
  the tight `<dimension>` box a written package records; the doc comment states the difference.
- **`cellValueToText` and `Cell.text` — one plain-text rendering of a value, for everyone.**
  `cellValueToText` is total over `CellValue`: the empty cell and an invalid `Date` give `""`, a
  boolean gives Excel's `TRUE`/`FALSE`, an error its literal, rich text its runs concatenated, a
  hyperlink its label, and any of the three formula kinds the text of its cached result. It is the
  value's text, not the cell's *displayed* text — no number format is applied, because the format
  lives on the style. `cell.text` is the same answer for the cell you are holding. The CSV writer
  now renders its fields through it rather than through a private near-copy, so a CSV field and
  `cell.text` cannot disagree about the same cell; `dateFormat`/`dateUTC` remain a CSV-only
  deviation. One consequence of the merge: a data-table formula's cached result now reaches a CSV
  field, where it used to render as empty.
- **The `CellValue` type guards are public.** `isErrorValue`, `isFormulaValue`,
  `isSharedFormulaValue`, `isDataTableFormulaValue`, `isRichTextValue` and `isHyperlinkValue` have
  always existed — the reader and the writer discriminate the union with them — but the `/core`
  barrel published only `detectValueType`. A consumer holding a `CellValue` therefore had no
  narrowing primitive at all: `detectValueType` classifies but does not narrow, so reading
  `.richText` off a value meant hand-rolling `'richText' in value` (which admits far more than the
  guard does) or an `as` cast. Now exported, each with the narrowing target pinned by a type-level
  contract.

### Changed

- **A release now publishes unattended.** The `npm-publish` deployment environment required a
  human to approve the job before it could reach the registry. On a single-maintainer project
  that was never a second pair of eyes — the only login that could approve was the one that had
  just cut the release — so it delayed every publish to re-ask a decision already made. Removed
  in `environment.yml`; deployments are still restricted to `v*` tags, which is now the whole
  gate. Nothing about the published package changes, and the OIDC trusted-publishing identity is
  untouched: authentication was never what the approval provided. See
  [ADR-0026](docs/decisions/0026-releasing-is-a-github-release-and-npm-follows.md) for what that
  trades away.

## [1.0.3] — 2026-08-05

Nothing in the library's behaviour changed: `src/` differs from 1.0.2 only in doc comments.
What reaches a consumer is the JSDoc the package ships — which now states what each documented
error is — and `dist/` built by a different compiler.

### Changed

- **The toolchain is on TypeScript 7.** `typescript@^7.0.2` — the native Go compiler — replaces
  the 6.0 line that [ADR-0008](docs/decisions/0008-typescript-6-upgrade.md) settled on, now that
  the docs generator no longer needs the printer API that 7 does not ship. The two scripts using
  the compiler programmatically moved to `typescript/unstable/*`; the type gate got ~4.5× faster.
  `src/` is untouched and the published API is unchanged. The emitted JS in `dist/` now quotes
  import specifiers with single quotes rather than double — the emitter's choice, no behavioural
  difference, and the full corpus passes against the new output. See
  [ADR-0028](docs/decisions/0028-typescript-7-adoption.md).

- **Contributor tooling: authored text may no longer contain characters that render as
  something other than what they mean.** `src/core/range.ts` carried a literal U+0000 byte
  as a sentinel where the six-character escape was meant. The two are identical to the
  compiler, and opposite to every text tool downstream: grep answered `Binary file … matches`
  instead of the matching lines, so the file silently dropped out of searches while still
  appearing to have been searched — which is how its two `@throws` tags survived the audit
  below. The sentinel's spelling changed, not its value. `scripts/check-source-text.ts` joins
  the invariants gate, refusing C0 controls other than tab/LF/CR, DEL, and bidirectional
  overrides (CVE-2021-42574) across `src`, `scripts`, `test`, `tools` and `docs`. Nothing in
  the published package changes.

### Fixed

- **The API reference now says what each documented error is, and no longer truncates the
  description.** Every `**Throws**` line in `docs/api/` was missing its error type: the generator
  matched the `{ErrorType}` slot only to strip it, so the reader was told a throw happens but
  never told what is thrown. Separately, 25 tags across 10 modules were written
  `@throws {@link SomeError}` — TypeScript parses the braces after `@throws` as a type
  expression, a `{@link …}` is not one, and the parse ran past the close brace and ate the rest
  of the comment. Those descriptions were missing from editor hovers too, not only from the
  generated pages; six of them rendered as the bare text `{`. The tags now use the brace-slot
  spelling the other 45 already used, the type is rendered rather than discarded, and `gen-docs`
  fails the build on a slot that is not a type name so the shape cannot return. Links
  mid-sentence were never affected and are unchanged.

  Two further faults in the same generator are fixed, both `docs/api/` only. A class member
  rendered its description and none of its tags, so ~40 `@throws` — including every one on
  `Workbook` and `Worksheet`, the two pages a caller is most likely to open — reached the
  reference never; members now get their own block, as top-level functions always had. And
  `{@link Target}` was parsed and then flattened to a bare code span, leaving the reference
  with zero links outside its index; 433 now resolve to the page and heading that documents
  the target, and `gen-docs` fails before writing anything if one would dangle.

- **A publish rehearsal now fails on a rejected identity.** `npm publish --dry-run` demotes a
  failed OIDC token exchange to a warning and exits `0`, so the rehearsal reported success for
  the one failure it exists to catch. That is not hypothetical: the 2026-07-30 rehearsal went
  green while its exchange returned `404`, which is how a misconfigured trusted publisher
  survived the check and cost two version numbers before anyone read the log. `publish.yml`
  now reads the verbose log and fails the job when the exchange was rejected — or when it was
  never attempted at all, the shape that made 1.0.1 fail `ENEEDAUTH`. Nothing in the library
  changed; `src/` is untouched.

  (1.0.2 itself reached npm on 2026-08-05, with provenance, once the publisher's `environment`
  field was corrected on npmjs.com. The workflow needed no change for it.)

## [1.0.2] — 2026-07-29

### Changed

- **Still nothing in the library — `src/` remains byte-identical to 1.0.0.** 1.0.1 was
  tagged and released on GitHub but **never reached npm**: the publish workflow failed
  `ENEEDAUTH`, having dropped `setup-node`'s `registry-url` on the incorrect theory that it
  would make npm send a placeholder token instead of exchanging an OIDC one. It is in fact
  what tells npm which registry to authenticate against, and without it the exchange never
  starts. Restored.

  1.0.2's own first attempt then failed too, for a second and unrelated reason — this one
  outside the repository. The trusted publisher configured on npmjs.com named
  `environment.yml`, the workflow that *provisions* the deployment environment, where it had
  to name the environment itself, `npm-publish`. npm answers a rejected identity with a 404,
  which reads as "no such package" and hides which claim failed to match. With the publisher
  corrected on npm, this tag publishes on a re-run; nothing here changed for it.

  The tag and release for 1.0.1 are left standing rather than rewritten: a published claim
  that turned out to be wrong is corrected in the open, not deleted. Install 1.0.2.

## [1.0.1] — 2026-07-29 · never published to npm

### Changed

- **Nothing in the library. This release exists to prove the release path.** `src/` is
  byte-identical to 1.0.0 — `npm diff` between the two will show only the version field.
  What changed is how a release reaches you: publishing a GitHub release now publishes the
  package, authenticated by a short-lived token minted from the workflow's OIDC identity
  rather than by any credential stored in the repository, and the result carries a
  provenance attestation tying it to the commit and the run that built it. 1.0.0 was
  published by hand and has no such attestation.

  A patch bump with no behavior change is the honest way to exercise that path end to end:
  the alternative is discovering a broken publish on a release that actually matters. See
  [ADR-0026](docs/decisions/0026-releasing-is-a-github-release-and-npm-follows.md).

## [1.0.0] — 2026-07-29

The first published release. `ts-xlsx` is an independent hard fork of ExcelJS, rebuilt from
the ground up in strict TypeScript; it carries no backwards-compatibility guarantee with its
ancestor — see [migrating from ExcelJS](docs/migrating-from-exceljs.md). The **BREAKING**
markers below describe changes made against the unreleased development line while this
section accumulated, not against any previously published version; there is no earlier
release of this package to break. From here forward the project follows
[SemVer](https://semver.org/) strictly — see
[ADR-0015](docs/decisions/0015-publishing-name-semver-and-first-version.md).

Charts, vector shapes, slicers and legacy form controls are **round-trip-only** in 1.0: a
workbook carrying them survives a load/edit/save byte-faithfully, but there is no API to
author a new one ([ADR-0014](docs/decisions/0014-charts-shapes-slicers-are-round-trip-only-for-1-0.md)).

### Added

- **`writeXlsxAsync` — the same package, deflated off the calling thread.** `writeXlsx` spends the
  whole cost of DEFLATE on the caller's thread, which on a large workbook means the event loop does
  not tick for seconds. The async writer shares part-building with the sync one and differs only in
  handing the part map to `fflate`'s worker-backed zip: every part compresses to identical bytes.
  Measured on a ~42 MB part map, a single large sheet takes the same wall-clock but the longest stall
  falls from the entire write to ~17 ms, and a twenty-sheet workbook finishes ~2.4× sooner because
  its parts deflate in parallel.

  There is deliberately no `readXlsxAsync` to match: reading is dominated by XML parsing and model
  building, which no worker can take, and the reader's zip-bomb ceiling depends on counting output
  between synchronous input slices. The asymmetry is the decision, not an omission —
  [ADR-0024](docs/decisions/0024-async-is-one-writer-not-a-mirrored-pair.md).

- **BREAKING: `Row` and `Column` are real classes, and `getRow`/`getColumn` return them.** They were
  formatting bags: `sheet.getRow(2)` handed back a `RowProperties` record with no way to reach the
  row's cells, so every row-oriented consumer hand-wrote an address-encoding loop, and `getRow` and
  `addRow` disagreed about what a row even was. A handle now carries both — `row.getCell('B')`,
  `row.cells`, `row.values` alongside `row.height`/`hidden`/`outlineLevel`/`collapsed`/`fill`, and the
  column equivalent including `key`, `width` and the six style facets it defaults for its cells.

  Formatting reads and writes exactly as before (`sheet.getRow(2).height = 20`,
  `Object.assign(sheet.getColumn(1), {key, width})`), and destructuring the iterators is unchanged
  (`for (const {number, cells} of sheet.rows())`). What breaks: the declared return type,
  `sheet.rowProperties(n)` / `sheet.columnProperties(n)` (use `getRow(n).properties` /
  `getColumn(n).properties`, which no longer fabricate either), and spreading a handle —
  `{...sheet.getRow(2)}` is no longer that row's properties.

- **`getRow` and `getColumn` no longer extend the used range.** They created a format record on
  access, so merely *asking* about row 500 made `rowCount` 500 and put an empty record in the
  worksheet model. The record is now created on first write. Reading a `<row r="5"/>` that states no
  attributes likewise leaves nothing behind, which is the honest reading of an element that says
  nothing.

- **`Cell.setRichText` — rich-text runs that inherit the cell's font.** A run's `<rPr>` is a
  *complete* character format: a facet it omits falls back to the workbook default font, **not** to
  the cell's. Verified against Excel — a cell set to Courier New 16 whose first run carries only
  `<b/>` renders that run in the workbook default face at the default size. So authoring
  `{bold: true}` on a run beside a styled cell silently loses the typeface, and the only fix was to
  restate the whole font on every run.

  `cell.setRichText([{text: 'Note:', font: {bold: true}}, {text: ' the rest'}])` composes each run
  over the cell's own font, per facet. Assigning `cell.value` directly is unchanged and stays the
  bare path, for a caller who wants a run that deliberately falls back to the workbook default —
  which is the format's rule, so the writer still emits exactly the facets a run carries.

- **`Range` — style a rectangular block in one call.** `sheet.getRange('A2:A33').border = {...}`, or
  `getRange(2, 1, 33, 1)` by inclusive corners. The third handle beside `Row` and `Column`, with the
  same contract: constructing one creates nothing, `addresses()` walks the block as a generator, and
  `cells` reports only what already exists. The six style facets plus a composing `style` accessor and
  `clearStyle()` mirror `Cell`'s semantics exactly — assigning a facet replaces that facet, assigning
  `style` composes facet by facet — so there is no second convention to learn.

  Writing materialises every position in the block, because a styled-but-valueless cell is the only
  way an empty cell renders with a fill; a uniformly styled block still collapses to one shared style
  entry. The cost is bounded by construction: `A:A` and `1:1` are refused, pointing at
  `getColumn`/`getRow`, which state a whole-axis default in one attribute instead of a million cells.
  A block overlapping a merged region restyles the region's master rather than stranding a style on a
  covered cell.

- **The workbook's default font is a first-class part of the model.** `Workbook.setDefaultFont`
  authors the face, size and colour every cell with no font of its own renders in — **empty cells
  included** — merging like `setTheme`, so `setDefaultFont({size: 14})` keeps the resolved face.
  `Workbook.defaultFont` reports the resolved, complete result, and `Workbook.declaredDefaultFont`
  reports what a source package stated (`undefined` when it stated nothing).


- **Seven subpath entry points, and `"sideEffects": false`.** `@shbernal/ts-xlsx/core`, `/xlsx`,
  `/xlsb`, `/csv`, `/vba`, `/customui` and `/errors` are published alongside the bare package name,
  which still exports everything it did. Additive — nothing moves or breaks. With a bundler the
  bare name remains the right default (`sideEffects: false` now lets it prune per symbol, which
  beats any subpath); reach for a subpath when there is no bundler, or when the module graph should
  state the dependency. `/errors` carries every error class the library throws and costs 12 KB, so
  classifying a failure never loads a parser; `/core` is 332 KB against the package's 902 KB.
  `/xlsx` is only marginally cheaper than everything, because `readXlsx` sniffs the bytes and
  dispatches a binary package to the BIFF12 reader — see the table in the README, and
  [ADR-0023](docs/decisions/0023-subpath-entry-points-and-disjoint-barrels.md) for why the split
  is shaped this way.

- **`tools/vba-compiler`** — an offline build tool that produces genuinely compiled, source-matched VBA
  p-code by driving a real headless Excel (VBIDE). Emits a `vbaProject.bin` (attach via
  `Workbook.vbaProjectBytes`) or a whole edited `.xlsm`. Windows + licensed Excel only; never in CI.


- **BREAKING: every deliberate failure now descends from one `XlsxError`, and carries a `code`.**
  `catch (e) { if (e instanceof XlsxError) … }` is the whole answer to "was that this library?" —
  previously five typed classes shared no ancestor, and the model's own validation threw bare `Error`
  distinguishable only by string-matching the message. `error.code` is the coarse branch
  (`'unsupported-format'` | `'malformed-input'` | `'authoring'` | `'internal'`); `error.name` and
  `instanceof` remain the exact one. See "How a failure is reported" in
  [docs/architecture.md](docs/architecture.md).

  New classes: `AuthoringError` (a document that cannot exist — 50 sites that used to throw bare
  `Error`), `PackageReadError` (a refused inflate, previously recognised by a message prefix),
  `XmlParseError` (malformed markup, previously a native `SyntaxError` that escaped `readXlsx`
  indistinguishable from a caller's own), `XlsxParseError`, and `InternalError` (an invariant of ours
  that did not hold). The five existing classes — `UnsupportedFormatError`, `XlsbParseError`,
  `VbaParseError`, `VbaAuthorError`, `CustomUiParseError` — keep their names, messages and fields and
  gain the ancestry.

  What did **not** change: scalar argument validation stays native `RangeError` / `SyntaxError` /
  `TypeError`, and every error message is byte-identical. The break is the *type* of a caught error,
  which matters if you catch `SyntaxError` around XML parsing or switch on `constructor`.

- **`Workbook.addTableStyle({name, elements})` — custom table styles are authorable.** A workbook can
  now define its own named table styles beside Excel's built-in gallery, and a table reaches one by
  putting that name in `TableStyleInfo.name`. Each element names a region (`wholeTable`, `headerRow`,
  `firstRowStripe`, … — all 28 of `ST_TableStyleType`) and carries a `DifferentialStyle`, interned into
  the same shared `<dxfs>` table conditional formatting uses, so two elements painted alike cost one
  entry. A stripe may set its band width with `size`. Authoring a name a source file already defined
  overrides that definition rather than adding an ambiguous second one.

  Verified against Excel Desktop, not just the schema: a table style is a cross-part correspondence
  every part of which can be valid while the table still renders unstyled. Excel registers the
  authored style in the workbook's gallery and paints from it — see
  `docs/knowledge/specs/custom-table-styles.md`.

  An unnamed style, or a `size` outside the four stripe types, throws — both otherwise produce a file
  Excel opens cleanly and then ignores. `TableStyleInfo.name` itself stays unvalidated, deliberately;
  the type documents why.

- **`Workbook.setTheme({colors, fonts})` — the workbook's palette is authorable.** A colour picked
  from a spreadsheet's theme row is written as `theme="4"`, a reference resolved at render time, so
  setting `accent1` restyles every cell, chart and table style that follows the theme at once — the
  only way to recolour a workbook without touching a cell. Any subset of the twelve colour-scheme
  slots and either of the two typefaces can be set, and calls merge.

  It generates *over* the existing theme rather than replacing it. The format scheme — the gradient,
  line and effect styles a designer authored — rides through untouched, a slot left unnamed keeps its
  source encoding (`dk1`/`lt1` stay `<a:sysClr>`, so they still follow the viewer's window colours),
  and a theme that references a picture keeps that relationship. `Workbook.themeColors` and
  `themeFonts` report the effective theme. A malformed colour throws at the setter, because Excel does
  not report one — it renders the slot as flat black.

- **`Workbook.resolveColor(color)` — a themed or indexed colour now resolves to a concrete ARGB.**
  A `Color` read from a file often carries no colour at all, only a reference: `{theme: 4}` into the
  workbook theme's scheme, or `{indexed: 2}` into the legacy 64-entry palette, either optionally with
  a `tint`. Resolution follows the workbook's *own* theme and its own custom `<indexedColors>` when it
  declares one, and applies the tint last. `Workbook.themeColors` exposes the scheme it resolves
  against. Two things it deliberately does not do: `indexed="64"`/`65` (the system foreground and
  background) resolve to `undefined` rather than to invented black and white, and nothing is written
  back into the model — the `Color` keeps the encoding its file used, so a round-trip still emits
  `theme="4" tint="0.4"` and the cell keeps its link to the theme.

  Note the index order: `theme="0"` is `lt1` and `theme="1"` is `dk1`, which is *not* the order the
  slots appear in the theme part. Verified against Excel Desktop; see
  `docs/knowledge/specs/theme-color-index-order.md`.

### Changed

- **BREAKING: the styles part's font 0 is the workbook's own default font, not an assumed Calibri.**
  The writer used to splice a `Calibri 11` constant into `<fonts>` with no workbook input, and two
  things followed from that.

  A themed workbook rendered every unstyled cell in the wrong face: `setTheme({fonts: {minor}})`
  wrote the theme part correctly and could not reach a cell, because font 0 went on claiming
  `scheme="minor"` — *I am the theme's body face* — while naming Calibri outright, and Excel resolves
  the explicit name. Working around it meant setting `font` on every column and naming the face in
  every rich-text run. And reading a package whose font 0 was Aptos Narrow and writing it back
  *unmodified* replaced the declared default with Calibri, re-adding the real face as a redundant
  custom entry — so populated cells still rendered right while empty cells, and the metric every
  character-unit `<col width>` is expressed in, quietly changed.

  Font 0 now resolves from the workbook: an authored default font, else an authored theme body face,
  else the file's own font 0, else the theme's. `scheme` and `family` are derived rather than
  copied — carried while the resolved face really is the theme's body face and dropped when it is
  not, which is what Excel itself writes — and the emitted entry always states a size and a colour,
  the absence foreign readers report as a "missing default font".

  What breaks: a workbook with an authored theme font, or read from a package whose font 0 was not
  Calibri, now emits different bytes. A plain `new Workbook()` is byte-identical.
  [ADR-0025](docs/decisions/0025-the-default-font-is-declared-not-assumed.md).

- **BREAKING: a corrupt or truncated package is a `PackageReadError`, not an unsupported format.** A
  `PK`-headed archive the zip layer rejects used to surface as `UnsupportedFormatError` with format
  `'unknown'`, whose message reads *"not a valid .xlsx package: no OOXML workbook part was found"* —
  a check that never ran, since nothing inflated. It now carries `code: 'malformed-input'` alongside
  the zip-bomb refusal, which is what the taxonomy already said it was: the container is the right
  kind of thing and we cannot unpack it. Code branching on `format === 'unknown'` for a truncated
  file must catch `PackageReadError` instead.

  Two messages get honest with it. A non-ZIP blob now says the input is not a ZIP rather than blaming
  a missing part, and *"no OOXML workbook part was found"* is left to the one case where it is true —
  a package that inflated and carries neither `xl/workbook.xml` nor `xl/workbook.bin`. Unchanged: the
  zip library's own text is discarded, never folded into the message and never attached as `cause`,
  because it can name internals or an absolute filesystem path.


- `removeVbaModule` and `addVbaReference` (and their `Workbook`/`editXlsxVba*` wrappers) no longer reset
  the `_VBA_PROJECT` stream — that reset crashed the VBA load on a project with real p-code. They now
  leave it byte-for-byte untouched; the `dir` stream carries the structural change. These edits are
  retained precisely because they never touch a module's compiled p-code.

- **Contributor tooling: the gate set has one name.** `node scripts/verify.ts` runs every gate
  concurrently and is what `pnpm test`, lefthook's `pre-push` and the turn-boundary hook all
  invoke; `--quick` is the inner loop and `--cached` exits immediately when the working tree is
  byte-for-byte the one that last passed. The corpus runner gained `--case`/`--json` and now
  prints only what needs attention (`--verbose` for the old listing), and the OOXML validator
  builds on demand instead of paying `dotnet run`'s project re-evaluation per call. Nothing in
  the published package changes. See
  [ADR 0022](docs/decisions/0022-verification-is-one-cached-parallel-entrypoint.md).

### Removed

- **BREAKING: the pure-TS VBA source-authoring API is gone.** `writeVbaProject`,
  `Workbook.setVbaProject`, `editVbaModuleSources`, `Workbook.setVbaModuleSource`,
  `editXlsxVbaModuleSource`, `editXlsxVbaModuleSources`, `addVbaModule`, `Workbook.addVbaModule`, and
  `editXlsxVbaAddModule` are removed, along with the `VbaModuleSource`/`VbaProjectSpec` types. They
  emitted modules with no compiled p-code on the theory that Excel recompiles from source on open —
  which it does not: such files either fail to load ("Invalid data format") or silently run stale code.
  See [ADR 0019](docs/decisions/0019-vba-authoring-needs-real-pcode-recompile-cookie-retracted.md).

### Fixed

- **A workbook's theme part is no longer destroyed on round-trip.** Reading a file and writing it
  straight back replaced its theme with the default Office one, so every `theme="n"` colour and
  `scheme="major|minor"` font in the file silently re-rendered in the wrong brand — a branded
  workbook came back in Office blue. The source theme is now preserved verbatim, reached through the
  workbook's `.../theme` relationship rather than the conventional `xl/theme/theme1.xml` path, and
  re-emitted with the parts it references (a picture used as a themed fill) so its `r:embed` does not
  dangle. Exposed on the model as `Workbook.themePart` / `restoreThemePart`, opaque preserved XML in
  the same spirit as `restoreDifferentialStyles`.

- **Custom table styles and recent-colour swatches are no longer dropped on round-trip.** A workbook's
  `<tableStyles>` definitions and the default table/pivot styles it nominates were discarded when the
  stylesheet was regenerated, so a table asking for a custom style by name was left referencing
  nothing and rendered completely unstyled — a file that opens clean and looks wrong.
  `<colors><mruColors>` (the "Recent Colors" swatches) went the same way. Both are now preserved
  verbatim and exposed as `Workbook.tableStyles` / `restoreTableStyles` and `Workbook.mruColors` /
  `restoreMruColors`. Each `tableStyleElement`'s `dxfId` keeps resolving because the differential-style
  table is re-emitted at its original indices, and the namespace prefixes Excel stamps on a table style
  (`xr9:uid`) are re-declared on the stylesheet root rather than left dangling.

[Unreleased]: https://github.com/shbernal/ts-xlsx/compare/v3.1.0...HEAD
[3.1.0]: https://github.com/shbernal/ts-xlsx/compare/v3.0.0...v3.1.0
[3.0.0]: https://github.com/shbernal/ts-xlsx/compare/v2.1.0...v3.0.0
[2.1.0]: https://github.com/shbernal/ts-xlsx/compare/v2.0.0...v2.1.0
[2.0.0]: https://github.com/shbernal/ts-xlsx/compare/v1.3.1...v2.0.0
[1.3.1]: https://github.com/shbernal/ts-xlsx/compare/v1.2.0...v1.3.1
[1.2.0]: https://github.com/shbernal/ts-xlsx/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/shbernal/ts-xlsx/compare/v1.0.3...v1.1.0
[1.0.3]: https://github.com/shbernal/ts-xlsx/compare/v1.0.2...v1.0.3
[1.0.2]: https://github.com/shbernal/ts-xlsx/compare/v1.0.1...v1.0.2
[1.0.1]: https://github.com/shbernal/ts-xlsx/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/shbernal/ts-xlsx/releases/tag/v1.0.0
