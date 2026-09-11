# Reading and writing the binary XLSB format

Cluster: io / formats

## Scenario

A user generating very large spreadsheets wants an alternative to the XML-based `.xlsx` format:
the binary `.xlsb` (SpreadsheetML Binary, BIFF12) form. The motivation is size and speed, since a
binary workbook stores cell data as packed records rather than verbose XML, so it is typically much
smaller on disk and faster to parse and emit for large datasets. The request is for an additional
serialization format, write and by extension read, not a bug: "produce a valid `.xlsb` that
Excel opens, and read one back."

> **Status: read is implemented, formulas included; write is not.** The read half is assertable and
> asserted by two corpus cases, `xlsb-binary-workbook-reads-like-its-xlsx-twin` for values, styles and
> geometry, and `xlsb-formula-token-streams-decode-to-formula-text` for formulas and defined names.
> Both are backed by one workbook Excel saved in *both* forms, so the XML twin is an independent oracle
> for what the binary must decode to. The write half remains a spec note. The decisions taken while
> implementing read are recorded under "Scope decisions" below.

## Desired behavior

- **XLSB is a second serializer and parser over the same logical model, not a bolt-on.** The workbook
  model (sheets, cells, styles, defined names, tables…) is independent of on-disk encoding; XLSX
  (XML) and XLSB (binary) are two codecs over it. Designing this cleanly means a pluggable
  container and part-codec layer where a part is `{contentType, encode(model), decode(bytes)}`, the
  same split that lets the streaming writer emit either form.
- **A produced `.xlsb` opens in Excel unmodified**, and a foreign-generated `.xlsb` reads back into
  the same model an `.xlsx` would, so a caller can convert between the two without loss of the
  supported feature set.

## Format notes / prior art

- XLSB uses the **same OPC/ZIP package** and relationship graph as XLSX, but the sheet, workbook,
  sharedStrings and styles parts are **binary BIFF12 record streams** (`BrtRow`, `BrtCellRk`,
  `BrtCellIsst`, …) rather than XML. Part names and content types differ, since worksheet parts are
  `.bin` and the workbook is `xl/workbook.bin`.
- The canonical reference is the open **[MS-XLSB]** specification. SheetJS/`xlsx` implements XLSB
  read and write and is the strongest cross-check for record encodings and edge cases: RK number
  encoding, inline against shared strings, formula token streams.
- **Numeric encoding quirks:** `RK` values pack a 30-bit integer or a truncated float with a
  ×100-divisor flag, and formulas are stored as **Ptg token streams**, not text, which is
  significantly more work than XLSX's plain formula strings.

## Scope decisions

- **Read before write.** *Taken.* Reading foreign `.xlsb` files is higher-value and lower-risk than
  writing them, and shipped first; the motivating large-workbook use case then pairs write support with
  the streaming writer, since binary record streams stream well.
- **Feature subset first:** *taken.* Values, shared strings and styles for read, then formulas and
  defined names, still deferring tables, pivots and rich formatting.
- **The container layer is shared, not duplicated.** *Taken.* `.xlsb` and `.xlsx` are the same OPC/ZIP
  package with the same relationship graph, so the bounded inflater, the magic-byte probe, the OPC and
  rel resolution, and the resolved-style-table shape are one implementation used by both codecs
  (`src/io/xlsx/sniff-format.ts`, `read-opc.ts`, `read-styles.ts`). Only the part parsers differ
  (`src/io/xlsb/`).
- **One public entry, auto-detecting.** *Taken.* `readXlsx` detects the serialisation from the package,
  by which office-document part is present, rather than from a file extension, and dispatches, so a
  caller handed a file never branches on its format. `readXlsb` is also exported for a caller that
  already knows what it holds. The `UnsupportedFormatError` `'xlsb'` branch survives only where a
  *particular* entry point still cannot take one: the row streamer, which is built on the XML worksheet
  parser.
- **Style parity in the first cut is full, not partial.** *Taken.* Number formats, fonts, fills,
  borders, alignment and protection all decode, because the records are fixed-layout and stopping
  halfway would have cost more in explanation than in code. The one exception is the gradient fill: its
  stop array is the only `BrtFill` field with no Excel-authored sample to check against, so it is
  dropped rather than guessed.
- **Where the binary states what XML omits, the binary reading drops it.** *Taken, and load-bearing.*
  BIFF12 writes every field on every record (a bottom vertical alignment, a locked cell, a General
  number format, a row's height, a pattern fill's automatic colour sentinels) where XML writes only
  what differs from the default. A reader that carries all of it through produces a *similar* model,
  not the *same* one. The rule is that each such field is compared against its default and dropped when
  it matches, which is what makes the corpus case's model-equality assertion hold.
- **Sheet protection reads from its two records, and every flag is inverted.** *Taken.* A protected
  sheet carries `BrtSheetProtection` (535): a 16-bit legacy hash, then sixteen `Bool32`s whose first is
  "protected" and whose other fifteen say an operation is **allowed**, the opposite sense to the XML
  attributes. A password adds `BrtSheetProtectionIso` (678): a spin count, the same flags, then the hash,
  the salt and the algorithm name. Both layouts, the flag order and the sense were read off files Excel
  saved as `.xlsb` and `.xlsx` from one workbook per setting, and the twin fixture carries a sheet
  protected each way. A flag matching its XML default is dropped under the rule above.
- **Formulas are decoded to text on read.** *Taken.* Of the three options, decoding to text, storing the
  token stream opaquely, or recomputing, only decoding gives the *same* model the XML reader produces,
  which is the point of the whole exercise. It is implemented as a stack machine over the postfix stream
  (`src/io/xlsb/formula.ts`), with the built-in function table transcribed from [MS-XLS] 2.5.198.17
  (`ptg-functions.ts`). Two properties of the format make the reconstruction exact rather than
  approximate: Excel stores the author's **parentheses explicitly** (`PtgParen`), so no precedence
  arithmetic is needed and `(1+2)*3` cannot decay into `1+2*3`; and a fixed-arity call carries no
  argument count, so the arity table is what says which operands belong to which call.
- **An undecodable token drops the formula, never the cell.** *Taken.* A token stream is only
  self-describing while every token's length is known, so continuing past an unrecognised one would
  desynchronise the walk and emit confident nonsense. The decoder returns nothing instead, and the cell
  keeps the cached result Excel stored beside the formula, which is exactly what the reader surfaced
  before the decoder existed. The same rule drops a defined name whose target will not decode.
- **The `_xlfn.` placeholder names are not defined names.** *Taken.* Excel registers a hidden,
  function-flagged `BrtName` for every post-2007 function a workbook calls (`_xlfn.TEXTJOIN`), and a
  call to one is a "user defined" call whose name comes from that entry. The XML form persists no
  `<definedName>` for them, so carrying them onto `Workbook.definedNames` would make the two readings
  of one workbook disagree. They are filtered out of the model but still **counted** for lookup, since
  a `PtgName` cites a position in the unfiltered list.
- **Only a self-contained externals table resolves.** *Taken.* A 3-D reference names its sheet through
  an index into `BrtExternSheet`, whose entries name a *supporting book* plus a span of its sheets.
  Rather than enumerate every record type that can open a supporting book, and risk miscounting into a
  reference that names the **wrong** sheet, anything in the externals block other than the single
  `BrtSupSelf` disqualifies the table, and 3-D references then decode to nothing. A workbook with no
  external links, the ordinary case, declares exactly one supporting book: itself.

## Open questions

- **Shared formulas lose only their grouping.** A spreadsheet fills a formula down a column by storing
  it once and marking the rest as clones. The XML form records that grouping (`<f t="shared" si=…>`);
  Excel's binary form **does not**, since it writes each cell's own formula out in full and emits no
  `BrtShrFmla` at all. So a clone reads back with the same formula *text* either way, and only the
  pointer to the master differs. That is a fact about the format, not a gap in the reader: there is
  nothing in the file to recover the grouping from. A `BrtShrFmla` from another producer, whose member
  formulas use the position-relative `PtgRefN` and `PtgAreaN` tokens, is not decoded, since no
  Excel-authored sample exists to verify the layout against, so those cells keep their cached results.
- **Token classes left undecoded**, each dropping the formula to its cached result: the extended
  (`PtgElfXxx`) family that spells a structured table reference, which belongs with table support;
  `PtgNameX`, a name in another workbook; and the precomputed-reference tokens other than `PtgMemArea`
  (`PtgMemErr`, `PtgMemFunc`, `PtgMemNoMem`), again for want of a sample to pin their layouts.
- **Whitespace inside a formula is dropped.** Excel records the spaces an author typed around a token
  as a `PtgAttr` hint. Reattaching one to the right operand is not something a postfix walk can do, so
  the formula reads back without it, and Excel redisplays it identically either way.
- **Features beyond values, styles and formulas.** Tables, pivots, conditional formatting and images
  are read from `.xlsx` and not yet from `.xlsb`; each is its own record family to decode, and each
  should extend the twin-comparison corpus case rather than assert against a hand-written expectation.
- **Rich text.** A pooled `RichStr`'s per-run formatting is skipped and the flattened text is read. The
  runs are `{character index, font index}` pairs over the style sheet's font table, so this is a small
  slice once the rich-text model is wired to it.
- **Row streaming.** `readSheetRows` and `readWorkbookStream` are XML-only. The binary cell table streams
  at least as well, since it is already a flat record run, but the streaming reader's state machine is
  built on XML events.
- **Write.** Untouched. Full BIFF12 record coverage plus Ptg *encode* is a large, self-contained
  sub-project.
- **Security:** *addressed for read.* Binary record parsing of untrusted input carries the same
  bounded-allocation and zip-bomb defenses as the XML path (see `bounded-memory-large-workbook-read`,
  `lean-zip-container-strategy`), plus the per-record length check specific to BIFF12: a record's
  payload is a *view* onto the already-inflate-capped part, never a buffer sized from the declared
  length, and a length that overruns the part is rejected rather than clamped. Length-prefixed strings
  check their byte count against the record before materialising a character.

Related: `bounded-memory-large-workbook-read`, `lean-zip-container-strategy`,
`unified-streaming-and-buffered-io`, `unsupported-input-format-typed-error`.
