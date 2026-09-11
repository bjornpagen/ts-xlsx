# Architecture

How `ts-xlsx` is built and why. `CLAUDE.md` is the constitution (the principles every
change answers to); this document describes the library as it stands and the working
agreements that keep it coherent. Point-in-time decisions live under
[`docs/decisions/`](decisions/) as ADRs.

## Origin

`ts-xlsx` is a hard fork of [ExcelJS](https://github.com/exceljs/exceljs), which had
gone effectively unmaintained while still serving tens of millions of downloads a month.
Two assets were trapped in that project, and they were of opposite kinds. One was
knowledge: thousands of hours of hard-won understanding of how real-world `.xlsx` files
behave, scattered across hundreds of issues and PRs. The other was code, a weakly-typed,
callback-flavored tree with a rotting dependency graph. The strategy was to separate
them, harvest the knowledge into a durable, implementation-blind form, then rebuild the
code from scratch against it.

That harvest is complete. Every credible bug, reproduction, and edge case became a
[regression corpus](../test/corpus/) case; the legacy tree is deleted; the runtime
dependency is now [`fflate`](https://github.com/101arrowz/fflate) alone. What remains is
a modern, strict-TypeScript library whose correctness is pinned by the corpus it carried
across.

## The corpus defines correct behavior

The [regression corpus](../test/corpus/) is what pins this library down. Each case
encodes "correct behavior" as implementation-blind assertions that run against any
implementation through a thin adapter, so a behavior, once captured, can never silently
regress. This is why the corpus outlived the rewrite. It was written against the
_behavior_ and not the code, so it validated the new implementation the same way it
indicted the old one.

The rule that follows: when in doubt, add a case. A bug without a corpus case is a bug
that will return. A missing feature is best reported as a corpus case so it is fixed once
and never regresses.

## Test topology

Tests live in two places on purpose, because they are two different kinds of test with
opposite contracts. Where a test goes is decided by _what it is allowed to know_, not by
tidiness:

- **Co-located unit tests, `src/**/*.test.ts`.** White-box. Each sits next to the module
  it exercises (`address.ts` ↔ `address.test.ts`), imports src internals freely, and moves
  or dies with that module under refactor. Co-location keeps the test honest about one unit
  and makes an untested module visible at a glance. Run by `test:src`.
- **Shared test support, `src/**/*.test-support.ts`.** The plumbing several co-located
  tests need and none of them owns: unzipping a written package, reaching one part, patching
  a part and reading the package back (`io/xlsx/package.test-support.ts`). It earns its own
  suffix rather than living in a `.test.ts` file, which `node --test` would try to run, or in
  a production module, which would ship it. `tsconfig.build.json` excludes the suffix, so a
  new one is invisible to `dist/` without further wiring. One rule holds it together: an
  accessor here fails loudly on a part that is not there. A helper that answers a missing
  part with an empty string makes every negative assertion built on it pass for the wrong
  reason. The rule reaches below parts: an element looked up inside one goes through
  `elementIn`, and a patch that changes nothing fails, since both are the same empty
  string arriving by another route.
- **The regression corpus, `test/corpus/`.** Black-box and implementation-blind (see
  above): cases reach the implementation _only_ through the adapter and must never import a
  src internal, because that blindness is the whole reason the corpus outlived the rewrite.
  It is a behavioral spec, not a test of any module. Run by `corpus`.
- **External oracle tests, `test/ooxml-validation/`.** These validate emitted packages
  against the independent `OpenXmlValidator`, via the shared `ooxml-validate` package
  (ADR-0002). Different toolchain (a .NET binary this repo neither builds nor pins) and
  different cadence (`test:ooxml`, not in the default `test`).

The wall matters: the `test/` trees earn their separation by being forbidden from reaching
into src the way a co-located unit test may. Put a white-box test in `src/`; keep `test/`
for the blind corpus and the external oracles. A "corpus" case that imports a src internal
has quietly stopped being implementation-blind, and `corpus-blind:check`
(`scripts/check-corpus-blind.ts`, in the `invariants` gate) refuses it: a case may import
`test/corpus/case.ts`, `test/corpus/untyped.ts` and `node:` built-ins, and nothing else. The same gate
holds the adapter to one route into the library: only `test/corpus/adapters/ts-xlsx/runtime.ts` names
`src/` by path, because it is where `CORPUS_TARGET=dist` swaps in the emitted JS, and every other
adapter module imports its bindings and types through it.

## Spec and schema reference

Correctness is defined by an external standard, so the ground truth lives in the repo
next to the code that answers to it:

- [`.claude/skills/ooxml-lookup/`](../.claude/skills/ooxml-lookup/) holds the ECMA-376
  schema (Transitional and Strict) as a local SQLite graph behind a query CLI, vendored for
  offline, deterministic reference while implementing. Ask it what may go inside an element
  and in what order, what attributes a type takes, and what values those accept, rather
  than hand-joining XSDs:
  `node .claude/skills/ooxml-lookup/scripts/ooxml.mjs children x:c`. It is _reference_,
  not a validator. Conformance validation stays with the independent `OpenXmlValidator`
  oracle (ADR-0002). Repo-only; never published.
- [`docs/knowledge/specs/`](knowledge/specs/) holds hand-authored, implementation-blind
  behavior notes from the harvest.
- Microsoft Learn MCP (`.mcp.json`) is grounded search over Microsoft's Open
  Specifications ([MS-XLSX] et al.) for the Excel-specific deltas the standard omits.

See ADR-0007 for why the static standard is pinned in the tree while the evolving prose is
an MCP, and ADR-0034 for why that pinned form is now a queryable graph rather than the raw
XSD set it used to be.

## Module layout

The source tree under [`src/`](../src/) is strict-TypeScript, ESM-only, and build-free on
the dev/test path (Node runs the `.ts` sources directly via type-stripping; `tsc` is the
type _checker_ and, for publishing, the emitter). The domain decomposition, in dependency
order:

| Area            | Role                                                                                                                           |
| --------------- | ------------------------------------------------------------------------------------------------------------------------------ |
| errors          | the failure taxonomy every layer throws through (`src/errors.ts`), below all of them                                           |
| core model      | `Workbook` / `Worksheet` / `Row` / `Column` / `Cell`, addresses and styles: the in-memory document                             |
| xml             | escaping, emission and a hostile-input-safe SAX reader (`src/xml/`), with no spreadsheet knowledge                             |
| opc container   | ZIP inflation under a bound, magic-byte sniffing, the relationship graph and part paths (`src/io/opc/`)                        |
| resolved format | `XfStyle`, what applying an xf to a cell means, and which xf a cell resolves to, shared by both codecs (`src/io/style/`)       |
| read policy     | the rules every reader obeys about a foreign file and no codec owns: read repair and the column budget (`src/io/read-policy/`) |
| xlsx read/write | OOXML parse and serialize; the hardest, highest-value code in the tree                                                         |
| xlsb read       | the binary BIFF12 serialisation of the same model, read-only so far (`src/io/xlsb/`)                                           |
| streaming       | bounded-memory row streaming, both reads and an incremental workbook writer                                                    |
| csv             | a thin, optional entry point, never coupled to the xlsx core                                                                   |
| vba             | native read/author/edit of a macro-enabled workbook's `vbaProject.bin` (`src/vba/`)                                            |

That order is a real constraint, not a description: `scripts/check-layering.ts` (a gate in
`verify --full`) fails the build on an import that runs up it. The rules it carries are that
`src/errors.ts` reaches nothing, `src/xml/` reaches nothing above it, `src/core/` never reaches a
serialisation, `src/io/opc/`, `src/io/style/` and `src/io/read-policy/` sit below every codec, and
the two codecs are peers, so `src/io/xlsb/` may not import `src/io/xlsx/`. Co-located tests are
exempt, since a test import is not a dependency of the graph we ship. Shared code that tempts a codec
to reach sideways belongs in `opc` (the container), `style` (the resolved format) or `read-policy`
(what a reader does with a file the model would refuse); that is what those directories are for. A
rule left inside one codec is a rule the other does not follow: the BIFF12 reader went without read
repair and without a column budget for as long as both lived in `src/io/xlsx/`.

### The two model classes delegate their state, they do not accumulate it

`Worksheet` and `Workbook` are the two classes everything else hangs off, so both would grow
without bound if every feature simply added a private field beside the getter that reads it. Past
about a thousand lines that is no longer a class you can read: the fields are scattered through the
file, and there is no point at which you can see what the object _is_.

Both push cohesive slices of state into their own objects and keep the public accessors in front of
them. `Worksheet` holds `DataValidationOverlay`, `ConditionalFormattingOverlay`, `GridEdits`,
`UsedExtent` (`core/used-extent.ts`), `WorksheetMerges` (`core/worksheet-merges.ts`, which owns the
`MergeIndex` in turn), `WorksheetPictures` (`core/worksheet-pictures.ts`) and `WorksheetComments`
(`core/worksheet-comments.ts`); `Workbook` holds `WorkbookVbaProject` (`core/workbook-vba.ts`),
`WorkbookTheme` (`core/workbook-theme.ts`), `WorkbookStyleTables` (`core/workbook-styles.ts`) and
`WorkbookMedia` (`core/workbook-media.ts`).
The public surface does not move: an accessor stays on the model class, keeps its name, its type
and its full doc comment, and becomes a one-line delegation. The doc comment staying put is not
incidental, since it is what `scripts/gen-docs.ts` reads and what a consumer sees; the slice
carries implementation notes only.

What makes a slice a slice is how little it touches outside itself. The VBA project reaches exactly
one thing, the preserved-reference list, which is handed to it. The style tables reach nothing: the
six of them (`<dxfs>`, the named cell styles, the two colour lists, the table-style block and the
definitions a caller authors) are one thing said six ways, a table read out of `styles.xml` and
handed back to the writer, each the target of an index held elsewhere in the file.

The theme reaches one, and it is instructive: resolving a colour needs the workbook's custom indexed
palette as well as the theme scheme, but that palette is also the writer's source for
`<indexedColors>` and is filled in by the reader, so it belongs with the style tables rather than
with the theme. It is passed in as a narrow accessor over that slice. Had the palette moved _into_
the theme, the rest of the styles-table state would have followed it and the result would be a
colour-and-styles overlay, which is not a slice of anything. When a candidate slice has more than one
or two such edges, that is the signal it is not one.

The media block is the case where that test cut a candidate in half rather than rejecting it. The
_registry_ (the images, and the content index that makes re-registering an identical picture a hash
rather than a walk) reaches nothing, and is `WorkbookMedia`. `exportImages`/`importImages` reach five
different things on `Worksheet` and stayed on `Workbook`, reading through the slice: moving them would
have relocated that coupling rather than removed it. The same cut runs through `WorksheetMerges`,
which owns the regions and hands back the rectangle a new merge covers, while the two things a merge
does to the _grid_ (collapsing the values it covers, widening the used extent) stay on the class that
owns the grid. A slice that had to be handed the row storage to buy one line at the call site would
have been paying two edges for it.

A collection accessor on either class hands back the _live_ array, not a copy, and that is a
decision rather than an omission. These are views onto a document that is still being edited: a
caller holding `workbook.worksheets` across an `addWorksheet` should see the sheet that was just
added, exactly as it would from a re-read, and a copy would quietly turn every such reference into a
snapshot taken at a moment the caller did not choose. It is also the cheap answer on paths that read
them per sheet or per row. What `readonly` buys is a compile-time refusal, and it is honest about
being only that: an untyped caller can push onto `merges` and leave `MergeIndex` believing something
the list no longer says. That is the same class of hazard as reaching into any other internal, and
the library does not defend against it here.

`StreamedSheetReader.merges` (`io/xlsx/read-rows.ts`) does copy, and the reason is not that one is
more careful than the other. Its array is _reassigned_ on each pass over the part, so a caller
holding the live one across a second iteration would be holding a detached snapshot without ever
having been told - the very thing the model's live array is not. Copying is what makes the detachment
explicit at the one place it can happen. The rule, then, is about lifetime rather than about trust:
an accessor on the owner of the state returns the state; an accessor on a transient view over
somebody else's scan returns a copy.

`Worksheet` is still over the thousand lines after those two, and deliberately so. What is left on
it is the grid and the things that reach into the grid constantly: tables and pivots materialise
header and totals rows and re-pin themselves through `GridEdits` on every splice, so lifting them
would produce a tables-and-grid overlay, which is the theme example above with a different name.

What that materialising _is_, though, belongs to the table, and lives there. `Worksheet.addTable`
hands the new table a `TableGrid`, the three-call channel a registered table holds into the grid
(does this cell hold a value, write this cell, open a row here), and the table fills its own header
and totals cells as the last act of construction. The knowledge that an empty header row is
corruption Excel repairs on open, that a totals aggregate is a `SUBTOTAL` under a code from
`TOTALS_ROW_SUBTOTAL_CODE`, and that a `custom` total is the column's own stored formula, is the
table's, not the sheet's. The round-trip guard is the load-bearing part: reading a workbook
re-registers every table after its cells are loaded, so only an _empty_ cell may be filled. A cell
holding rich text, a style, or text that drifted from the column name is authoritative, and
clobbering it would be silent data loss on every file that has a table, with no schema error to
catch it. That is why the channel asks whether a cell holds a value rather than handing over the
sheet: a materialiser that could reach `#cellAt` would create the very cells it was asking about. The
page-layout fields (`view`, `pageSetup`, `printOptions`, `pageMargins`, `headerFooter`, and the two
break lists) are plain mutable objects with no accessors and no behaviour, so there is nothing to
delegate and grouping them would change the public API to no end. The line count is the symptom the
rule watches for, not the rule; a slice that is not one costs more than the lines it removes.

`UsedExtent` is the narrowest of them: its only edge is the four storage collections it is handed,
and it exists because the used range is read far more often than the grid changes shape. Deriving it
by scanning every cell is what made appending quadratic, since each of the three appenders (`addRows`,
`addColumns`, and the streaming writer's row numbering) reads it once per line it appends. Caching
the answer is not available: a caller holding a `Cell` styles or clears it without the sheet hearing
about it, so a remembered extent goes stale invisibly, and a wrong used range is worse than a slow
one because it is what lets an append land on a row someone had prepared. What the extent keeps
instead is the _structure_, which the sheet does observe: the highest row key, the highest
materialised column, and the highest line declared by a row height, a column width or a merge. Those
bound the extent from above, the top line is the one that was just appended in the case that matters,
and confirming it against the live grid costs a single row. A bound may overstate, which costs a
scan; it may never understate, which is why every edit that can pull the grid inward marks the bounds
stale rather than adjusting them.

`MergeIndex` is the same shape one step further, and its design turns on what a hostile file can
ask for. Overlap-checking a new merged region against every existing one made loading them quadratic
(20,000 regions took 1.4 seconds against 58 ms for 2,000), and the reader calls `mergeCells` once per
`<mergeCell>` in the part, so that scan sat on an untrusted path. The obvious index, a bucket per row
holding every region covering it, is the one that cannot ship: a sheet of whole-column merges is
legal, disjoint and cheap to write, and would claim a bucket entry per row per region. So the index
buckets a region under its _top_ row alone, in bands, and keeps the height of the tallest region on
the sheet; a region overlapping a query must begin within that height above it, which bounds the
bands worth visiting. Memory is one entry per region, and the price is that a sheet of tall regions
widens the window until the query is the scan it replaced, never worse. Resolving a covered address
to its region's master rides the same index, which takes that scan off `getCell` as well.

`GridEdits` owns that splice arithmetic for _everything_ anchored to the grid, which is a wider set
than the cell rows: line metadata, merges, tables, anchored images and the coordinates a cell's value
carries (a shared-formula master, a hyperlink's clickable range, a data table's ranges) move with the
cells, and so do the four things bound to a range that live outside the cell grid entirely:
data validations, conditional formats, comment threads and the sheet's autofilter. The invariant
is the general one, not a list that happened to be complete once: an overlay left behind re-points a
dropdown or a highlight rule at cells nobody chose, and the writer emits that without complaint. The
single coordinate rule they all share is `core/grid-shift.ts`. The edit is one value there,
`AxisSplice`, rather than the loose `(axis, start, count, delta)` a call site cannot be read against:
three of the four are `number`, and a transposed pair is invisible to the compiler, to the linter and
to Excel -- the workbook it produces is well-formed and aimed at the wrong cells. Every participant
answers the same two questions through that module -- where does this land, and did the delete
swallow it whole? -- and the projections built out of them (a span, a rectangle, a point) live beside
them, so the axis ternary is written once instead of once per overlay. A range the delete swallowed
whole takes its entry with it rather than clamping onto the cut line, because dropping a rule is
legible and silently re-aiming one is not.

Formula text is the one participant that reaches past the sheet. A formula spells its coordinates
rather than storing them, so `core/formula-references.ts` reads the text and moves each reference to
the spliced sheet by Excel's rules, and a reference to that sheet can sit in any sheet's formulas and
in a defined name. So the splice runs a formula pass before any cell moves: the sheet rewrites its own
cells, validations and conditional formats, then hands the edit to the workbook through a host it was
given at creation, which passes it to every other sheet and to the defined names. Before the move
matters twice: what an insert brings in was written against the grid after the edit and must not move
again, and a shared-formula clone is recovered from its master at the offset they have now. A clone the
rewritten master no longer describes becomes a formula of its own, and the writer groups what is still
shared. A table's column formulas are formula text too and move in the same pass, and so does an
authored pivot's source range, the one reference Excel never turns into `#REF!`: a delete that takes
all of it leaves it as written.

Preserved parts spell references too, a chart's series and a loaded pivot's cache source among them,
and Excel moves those as well. The model holds them as bytes, and `core/` may not parse XML, so the
workbook keeps a journal of every splice (`workbook[INTERNAL].splices()`) and the writer replays it
over each such part as it plans the preserved parts (`io/xlsx/preserved-splices.ts`), editing at the
offsets the scanner found. Replaying the whole journal over the bytes as they were read is sound for
two reasons the model guarantees rather than assumes: a preserved part only enters a workbook as it is
read, before anyone can splice it, and a sheet's name, which each entry is keyed by, never changes. A
loaded pivot's model view moves in the formula pass by the same function the writer calls,
`splicePivotSource`, so the view and the bytes cannot disagree about where the source went. What the
replay does not reproduce is Excel re-deriving a series whose name or categories a delete takes: the
reference the delete took becomes `#REF!`, as a formula's would. A hyperlink's in-workbook location (`S1!B5`) is left as written on purpose:
Excel 16.0 leaves it too, even when rows are inserted above the cell it names.

The row and column axes are deliberately not mirror images, and where they diverge is a decision
rather than a gap someone forgot to close. A row takes either input shape, a positional array or an
object keyed by `ColumnProperties.key`; a column takes only the positional one, because a column's
values are indexed by _row_ and a row carries no key for the other shape to name. `duplicateRow` has
no column counterpart for a mechanical reason rather than a matter of taste: a row insert carries
pre-built `Cell`s, so a duplicate keeps the source's per-cell styles, while a column insert carries
raw `CellValue`s that materialise fresh cells, so the same verb on that axis would silently drop the
styles it claims to be copying. What is missing is the machinery, not the verb, and adding it means
giving `spliceColumns` a cell-shaped insert path first.

### Inside `src/io/xlsx/`: three kinds of module, deliberately flat

Thirty-odd files in one directory reads like something nobody got round to organising, and the
obvious tidy-up into `read/`, `write/` and `shared/` is wrong here. The modules fall into three
kinds, and the split would cut across the most cohesive of them:

- **Feature modules, both directions in one file.** `comments.ts`, `tables.ts`, `images.ts`,
  `hyperlinks.ts`, `data-validation.ts`, `conditional-formatting.ts`, `threaded-comments.ts`,
  `sheet-properties.ts` and `font-xml.ts` each export a `parseX` for the reader beside an `xXml` for
  the writer. That pairing is the point: the two halves share one feature's element names and must
  agree with each other, and a round-trip is exactly the claim that they do. Splitting each into two
  files would double the count while moving the two functions that have to stay in step into different
  directories. The pull is the other way, and it is the pull to leave a reader in whichever module
  happened to call it first: `<pageSetup>` was written in one file and read in another, and
  `<font>` was written in the stylesheet writer and read inside the 744-line style-table reader,
  which charged every consumer of a rich-text run for that whole reader.
- **The read pipeline.** `read.ts` and everything prefixed `read-*`, plus its private helpers
  (`cell-accumulator.ts`, `cell-value.ts`, `read-rich-runs.ts`). The `read-` prefix is the convention;
  `pivot-read.ts` and `shared-strings-read.ts` were the two files spelling it the other way round, and
  `rich-runs.ts` was one spelling it neither way.
- **The write pipeline.** `write.ts`, `write-stream.ts`, the `*-xml.ts` serialisers, and the
  write-side services (`styles.ts`'s interning registry, `shared-strings.ts`, `package-plan.ts`).

The invariant worth having is that the read pipeline never reaches into the write pipeline. It
holds, and `color-xml.ts` is why it took work: `parseColor` sat in the write-side style _table_, so
the style reader and the worksheet reader both imported the writer to decode a `<color>`. Reading
and writing that element are one concern with two directions, and they now live together in a module
either side may use.

That invariant is documented rather than gated, because a gate for it cannot be honest.
`check-layering.ts` matches directories, and a rule derived from reachability defeats itself: the
moment a read module imports a write module, that module becomes reachable from the read roots and
so classifies as _shared_, which is precisely the label that makes the check pass. A declared list
of write-pipeline modules would work but goes stale in silence, which is the failure these checkers
exist to prevent. See ADR 0030.

Cell formatting is one named tuple, not six loose fields. `CellStyle` in `core/style.ts`
holds the six OOXML direct-format facets (`fill`, `numFmt`, `font`, `border`, `alignment`,
`protection`), and every style-bearing shape derives from it rather than re-declaring the
tuple: a cell model, a column's defaults, a table column, a differential (conditional)
format, a named style. `CELL_STYLE_FACETS`, derived from `Record<keyof CellStyle, true>`, is
the single facet list the copy loops walk, so adding a facet is a one-line change the compiler
forces every consumer to honour. Applying a style splits by target: `applyCellStyle` drives a
`Cell`'s per-property setters, `assignStyleFacets` copies plain records. Two helpers, because a
cell and a bag of fields take writes differently.

The same pattern governs a sheet's snapshot one level up. `WorksheetModel` is what
`dst.model = src.model` carries, and a field the getter emits but the setter ignores loses data
silently, which is the merge-loss failure that contract exists to prevent. Both directions are
therefore driven from one table, `WORKSHEET_MODEL_FACETS` in `core/worksheet-model.ts`, where each
field declares its read and its write side by side along with the clone strategy that field needs
(`{...spread}`, `replaceContents`, replay through the authoring API). Declaration order is the
order a model assignment applies: cells are placed before any merge exists, so a covered cell's
value lands where the model says instead of being routed to a region master mid-load. The registry
is proved exhaustive over `keyof WorksheetModel` at compile time, so a field added without a facet
is a build error that names the field. An assignment is applied to a detached scratch sheet first
and reaches its target only once every facet has applied there, so a model that throws part-way
leaves the target as it was.

### The harness is code, and is held to the same rules as the library

Four gates ask `scripts/module-graph.ts` what a module imports, and its comment-stripper is a
hand-rolled character scanner: the right shape, since a parser for four gates would cost more than it
saves, and exactly the shape that gets one case wrong quietly. It did. `const a = /don't/;` opened a
phantom string at the apostrophe, so a commented-out import on the next line read as a real one. That
degrades toward a false alarm, which is the safe direction, and a gate that can fail for reasons
unrelated to the code is still how a gate loses its reader. `scripts/module-graph.test.ts` runs in its
own `test:harness` gate and covers the scanner's edges.

Which file a tool runs over is `scripts/targets.ts`, once. The unit-suite glob had three copies, the
lint targets two, the formatter's targets carried a comment claiming they tracked the tsconfig include
lists and had already drifted, and the coverage exclusions were a partial copy of
`tsconfig.build.json`'s. `package.json` cannot import a constant, which is why `lint` and `test:src`
are now thin `node scripts/…` wrappers, the shape `format` and `build` already had.

### A table is only a single source of truth if the other copy is derived from it

An exhaustiveness proof covers omission from _the table_. It says nothing about a consumer that
re-enumerates the same set beside it, and that is where these tables have actually drifted. Three
rules follow from the cases fixed so far.

**A format-specific reading of a shared list belongs to the format, but the list is still shared.**
`ALIGNMENT_FACETS` proves that the XML reader and the XML writer agree on the seven `<alignment>`
facets; the BIFF12 reader restated all seven and their default-omission rules by hand, so an eighth
facet would have compiled, lit up in XML both ways, and silently vanished from every `.xlsb`. The
codec now walks the shared list and supplies only its own bit layout, as a `Record` keyed by
`Alignment` so the compiler asks for the eighth entry too. What stays in `core/` is the _set_;
each format brings its own reading of it. Masks and shifts do not go in the model.

**A closed enumeration a binary format indexes needs the ordered list, not a second table.**
`ST_PatternType` was written out three times: the union, the guard's lookup table, and the BIFF12
index array. Only the first two were checked against each other. `FILL_PATTERNS_IN_SCHEMA_ORDER` is
now the list; the guard derives from it (`tokenSetOf`) and the codec indexes it. A list can name
_fewer_ members than its union where the `Record<T, true>` shape cannot, so it carries an
`AssertNever` proof of the other direction explicitly. That is the price of needing an order.

**Two things a file must agree on must be computed once, not twice identically.** The classic
data-bar element and its x14 extension describe one bar and must show the same anchors or Excel
repairs the sheet; both halves used to derive the defaults for themselves. Likewise a
`Begin`/`End` record pair, declared as one triple rather than as an entry in a start table and a
matching entry in an end table, because a mistyped pairing across two maps reinstates exactly the
"an `EndFonts` cannot close `<fills>`" hazard the tracker exists to remove. And an arity recorded
beside its function name rather than in a map keyed by name, so "no arity recorded" stops being
representable at all. It used to read as "variadic", which made a `PtgFunc` pop the wrong operands.

The same reasoning applies to a _predicate_ several layers share. `isRelType` in `rel-type.ts` is a
leaf importing nothing, because the reader and the writer both ask whether a relationship Type names
a part class, and putting the answer beside the reader pulled the whole read path (the bounded
inflater included) into the writer's module closure. The per-entry size budget is what noticed.

It sat in `io/opc/` while only the codecs asked it, and that was the wrong home the moment two other
layers did. `core/workbook-vba.ts` finds the macro project among a workbook's preserved references
and `customui/ribbon.ts` recognises a ribbon part; neither may import `src/io`, so both spelled the
suffix test by hand, three sites carrying the separator gotcha with nowhere to say why the slash was
load-bearing. The question that settles the home is whether a relationship Type is a concept the
model may own, and it is: a `Workbook` already holds preserved references keyed by their Type, so the
string is model state and only the reading of it was missing. Reading a URI's last segment is no more
a serialisation concern than rendering a number as hex, so it moved to a root leaf beside `hex.ts`,
under the same layering rule.

Not everything the codecs need to do to the model belongs in its public API. Pushing preserved bytes
back into a `Workbook`, restoring a loaded sheet's hashed protection credential, placing a cell at an
exact position without resolving merges, evicting a row the streaming writer has already serialised:
about fifteen operations exist for a codec's benefit and put the model in states no authoring path
can produce. They were public members. They shipped in the `.d.ts`, they appeared in the generated
reference, and the model class _was_ the codec's mutation interface. They now hang off one symbol key
in `core/internal.ts` (`sheet[INTERNAL].restoreProtection(…)`), which no entry barrel exports, so the
boundary is the module graph rather than a naming convention. No layering rule guards who may import
that module: `package.json` maps only the eight subpaths, so the symbol is already unreachable from
outside the package, and a rule would only police `src/core` against itself.

The same key draws the same boundary one layer up, around the streaming writer rather than the model.
`WorksheetStreamWriter`'s constructor took the writer's `StyleRegistry` and its `flushedSheet()`
returned its `FlushedSheet`, so seven writer-internal types were named by a published class's
signature and a consumer could hold values whose types the package does not export. They carried
`@unpublished`, which records a decision without changing anything: `new WorksheetStreamWriter(…)`
still compiled, and nobody should ever call it, because the sheet writer comes from
`WorkbookStreamWriter.addWorksheet`. The constructor is now `private` with a static
`WorksheetStreamWriter[INTERNAL].create`, and the flush plumbing is on the instance channel, so the
two members are gone from the `.d.ts` and from the API reference. `StreamedRow` is built the same way,
through `StreamedRow[INTERNAL].create`, for a sharper reason than tidiness: a row names a number its
writer hands out, and a caller-built one could commit a number the writer had not, evicting that row
from the model with nothing written for it.

The channel takes two shapes on purpose. `Workbook` and `Worksheet` carry a symbol-keyed _object_ of
operations, one allocation per book or sheet, which is nothing. `Cell` gets a symbol-keyed accessor
_pair_ instead, because a per-instance channel object on the one class allocated in the millions is a
real cost for state most cells never carry. Despite the origin it is not the _codec_ channel: the
model's own `dst.model = src.model` setter reaches through it for exact-position cell placement, and
`Row`/`Column` reach through it for the per-line stores. It is the internal channel, and core is
allowed to use it.

`Row` and `Column` (`core/row.ts`, `core/column.ts`) are handles, not records. `Worksheet`
keeps the authoritative stores, the row-major cell grid and the two sparse maps of per-line
formatting. A handle holds nothing but the sheet and a position, reading and writing straight
through. That is deliberate: a row object that copied its cells out would be the shape of the
merge-loss bug the model contract exists to prevent, and two handles on the same number could
disagree. Position is fixed at construction, the rule `Cell` already follows: a splice that moves
content past `sheet.getRow(3)` does not carry the handle along, any more than it re-points a `Cell`.
Formatting is created on write and never on read, so `getRow(500)` costs nothing and does not extend
the used range; the format record appears when a value is set. The handles reach the stores through
`sheet[INTERNAL]`, which is the same channel the codecs use.

Each handle mirrors its record's fields as accessors, five for a row and eleven for a column (the
geometry plus the six inherited `CellStyle` facets), so `row.height = 20` is the flat, discoverable
path rather than a hop through a properties bag. That mirror is proved complete at compile time the
same way the model registry is: a field added to `RowProperties` or `ColumnProperties` with no
accessor fails the build naming the field, because otherwise the record would carry it, the codecs
would read and write it, and the public handle would simply never mention it.

Those sixteen pairs are spelled out rather than generated from a facet list, and the decision was
taken deliberately once the list-driven form was costed. The facet lists elsewhere in `core/`
(`CELL_STYLE_FACETS`, `WORKSHEET_MODEL_FACETS`) drive _loops_ over state a caller never names;
generating accessors is a different thing, and its price is paid twice. `gen-docs.ts` reads class
members, so each property would still need a declared member carrying its doc comment, leaving only
the four-line body to save. And the members would have to be installed at runtime rather than
declared, which trades a surface the compiler checks for one it merely believes. That is the
`AssertNever` proof above spending its own guarantee. A slice that is not one costs more than the
lines it removes, and so does an abstraction.

The plumbing _underneath_ those accessors was costed separately, and the answer split. Four members
were byte-identical between the two handles modulo which coordinate they name: the private read and
write helpers, and the `values` getter and setter. The first pair moved; the second did not.

The read and write pair is `AxisHandle` (`core/axis-handle.ts`), which `Row` and `Column` both
extend. What made it worth having is not the lines it saves but the rule it states: writing
`undefined` clears the field, and clearing the _last_ field takes the record with it, because the
used extent derives its bounds from which lines have a record. That rule is subtle enough that
stating it twice is how it drifts, and it did: an emptied record used to pin `rowCount` at a row
nothing formatted any more. A base class is also the allocation-free shape, which the free-function
alternative was not: that needs the pair of stores the handle reads through, `peek` (which never
fabricates) and `ensure` (which materialises on first write), as an object the handle holds, and a
handle is constructed on every `getRow`/`getColumn` and once per step of `rows()`/`columns()`, so
iterating twenty thousand rows and reading one property measured about 17% slower.

`values` stayed duplicated, and that is where the `gen-docs.ts` constraint bites: it reads class
members, so `values` moving to a base would drop out of the reference unless `Row` and `Column`
redeclare it, at which point nothing is shared. Sharing it would also need an abstract
cell-by-position accessor and an abstract position-of-cell reader to bridge `cell.col` against
`cell.row`, which is more indirection than the six lines it would replace. `axis-handle.ts` says so
from its own side.

The same test was applied to `DataValidationOverlay` and `ConditionalFormattingOverlay` and reached
the same answer for a different reason. About fifteen lines are byte-identical between them: an
`add` that clones, an `entries` getter, a `shift` that maps every entry through `shiftSqref` and
drops what comes back `undefined`, and a `clear`. A shared base needs a type parameter, an accessor
for whichever field holds the sqref, and an injected clone function, and even then the two are not
symmetric: `DataValidationOverlay` also maintains a decoded-rectangle index that `at()` answers
point lookups from and that `shift` must re-derive, and the conditional-formatting overlay has
neither. A base class would carry one subclass's collection discipline while the other's sat outside
it, which is the slice-that-is-not-one test failing again.

If that is ever taken up, the honest shape is not a base class but a free function both `shift`
methods call, `shiftSqrefEntries(entries, refOf, withRef, axis, start, count, delta)`. That is the
part which is genuinely one rule, and it is the part where a drift would silently re-aim a
validation or a highlight at cells nobody chose.

**Costed and declined once, on the grounds that the rule was better placed elsewhere.** The rule at
risk is "a `sqref` whose every area the splice deleted takes its entry with it", and that rule lives
in `shiftSqref`, not in either caller: `src/core/merge.test.ts` now pins it there, along with the
byte-clean guarantee that an unmoved area comes back as its own text. What the free function would
extract from the callers is three lines, and it would need a `refOf`/`withRef` pair only because the
two spell the field differently (`sqref` against `ref`), while the validation overlay would still
keep its loop to re-derive the rectangle index. More indirection than it removes; taken up only if a
third `sqref`-bound overlay appears.

The xlsx reader and writer, the two largest pieces here, are each a cluster rather than a
monolith, split along the OOXML package's own divisions so a change touches one part:

- **read** (`src/io/xlsx/`): `read-styles.ts` (`styles.xml`), `read-worksheet.ts` (one sheet),
  with `read.ts` keeping `readXlsx` and the workbook-level wiring. `read-rich-runs.ts` owns the
  `<r>`/`<rPr>`/`<t>` element machine every `CT_Rst` reader shares (a worksheet's inline strings, the
  shared-strings pool, and a note's body), taking the container name (`is`, `si` or `text`) as a
  constructor argument, since that is the only thing that differs between them, and keeping a
  phonetic run's `<t>` out of the text in all three; `cell-accumulator.ts` owns
  the per-cell gathering state machine the buffered and streaming readers both drive (ADR-0004).
  Both are _machines_, not bags of state calls, and for one reason: a grammar two readers spell out
  separately is kept in step by convention rather than by mechanism, and the drift it admits reads
  the same content two ways depending on which encoding the producer happened to choose. Every other
  parser that gathers an element's text across open/text/close events does it through `TextCapture`
  in `src/xml/xml-read.ts`, which exists because nine hand-rolled copies each had to remember that a
  self-closing `<x/>` fires no close and a latch nothing closes is a latch that eats the next
  element's text. `capturedText` is its pull shape, for a parser whose whole job is reading a handful
  of text elements out of a part; a parser that interleaves capture with per-element state of its own
  stays bespoke. The same rule reaches past the grammar to the _decisions_ the two readers make about
  a row and a column: `row-position.ts` owns where a `<row>` sits when it declares no `r`, and
  `takeColumnSpan` (`column-span.ts`, in front of `clampColumnSpan` and `ColumnRecordBudget` in
  `src/io/read-policy/column-budget.ts`) owns which columns a `<col min max>` reaches. Both
  were a line each, copied into both readers, and both copies had drifted: one reader inferred a
  missing row number positionally while the other dropped the row's formatting, and one tested a span
  starting past the grid while the other relied on the budget's contract to catch it by accident. A
  decision small enough to inline is exactly the size that drifts unnoticed.
- **write** (`src/io/xlsx/`): `package-plan.ts` (the part-graph plan layer), `workbook-xml.ts`
  and `worksheet-xml.ts` (the serialisers), `relationships.ts` (the SpreadsheetML relationship-type
  vocabulary), with `write.ts` keeping `writeXlsx` and the `buildPackageParts` orchestrator.

Neither cluster owns the container it rides in. `src/io/opc/` holds what is true of _any_ OOXML
package: `inflate.ts` (the bounded inflater), `sniff-format.ts` (the magic-byte probe and the
typed rejection), `read-opc.ts` (resolving relationships and walking a part closure), `rels.ts`
(emitting a `.rels` part, escaping its own attributes rather than asking each caller to),
`part-paths.ts`, and the package namespaces. Beneath even that,
`src/xml/` holds escaping, emission and the SAX reader.

### The write boundary is where a value stops being a value and becomes bytes

A `PageSetup`, a `Color`, a `ConditionalFormattingRule` and a dozen shapes like them are plain data
the model stores verbatim, with no accessor to validate through. That is deliberate: making them
classes to catch a mistake would change the public surface of every one to guard a moment that has
not happened yet, and a `Color` that never reaches a file never had a problem. So the check belongs
at the single point where the value becomes bytes, and `src/xml/xml.ts` states all three forms of it:

- **A number is refused if the format cannot spell it.** `assertWritableNumber`, in `src/errors.ts`
  rather than in the serialiser, because a number has two spellings here and they sit on opposite
  sides of the `core`/`xml` boundary: `numberText` / `numAttr` write an attribute, and
  `formulaNumberLiteral` in `core/formula.ts` writes a formula literal, which the BIFF12 codec
  produces as readily as the XML one. Every numeric attribute in OOXML is `xsd:double`,
  `xsd:unsignedInt` or a bounded flavour, and no lexical space has a form for a NaN or an infinity,
  so writing one produces a package Excel reports as damaged. The refusal is also exported on its own
  because a value can be unwritable and still be _read_ on the way to the bytes: the collapsed-row
  scan compares outline levels to walk a group, and against `-Infinity` every comparison holds and
  the walk never ends.
- **A date is refused if `dcterms:W3CDTF` cannot spell it.** `assertWritableDate`, for the same
  reason and with the same shape. An Invalid Date is truthy and is an instance of `Date`, so it
  passes every guard short of this one; a year outside 0000-9999 does not throw at all and writes ISO
  8601's expanded `+275760-09-13T…` notation, which no `xsd:dateTime` admits.
- **A token from a closed enumeration is checked, not escaped.** `checkedToken`, against the
  `is*` guard in `src/core/` that owns that enumeration. Escaping a bogus token yields a well-formed
  document Excel still rejects, which buries the mistake in the file instead of raising it at the
  call.
- **A free string is escaped.** `escapeAttr` / `escapeText` / `textAttr`, plus
  `escapeSpreadsheetText` for the measured set of elements where Excel decodes `_xHHHH_`. A
  character XML cannot carry at all is refused, because there is no faithful representation and a
  sheet named `Sheet_x0001_A` is a different name rather than a rendering of the one asked for.

The read side is the same grammar with one asymmetry, and stating it is what makes the pair
trustworthy. `enumToken`, `numFinite`, `numInteger` and the `bool*` family in `xml-scan.ts` _drop_
what they cannot read, where the writer throws. A file the library did not write is allowed to be
wrong, and losing one attribute beats losing the sheet; a value an author supplied is a mistake at
the call. Both halves lean on one guard per enumeration, so what the reader accepts is always
something the writer can write back. Where that leaves a model fragment unwritable anyway, the
reader drops the fragment: a `<cfRule>` whose type is not in `ST_CfType` is dropped whole rather than
half-read, because `type` is the attribute every other field is read relative to.

Two rules bind the _emitter_ rather than the value it is handed, and both were learned the same way.
**Content never decides how its container is serialised.** The one row attribute that cannot be known
when a streamed row is rendered, `collapsed="1"`, is patched in afterwards; the patcher used to ask
the rendered markup whether the attribute was already there, and that markup also carries cell text
in which `escapeText` leaves a double quote verbatim, so a cell whose value was the literal string
` collapsed="1"` suppressed the attribute on the row containing it. What a row declared is carried as
a field beside its XML, never re-read out of it. **An edit to a part is made at offsets a scanner
found, never by a pattern.** `elementRange` in `xml-read.ts` is that primitive: it returns source
offsets, so everything outside the range it names is spliced through byte for byte -- the whitespace,
the comments, the attribute order, the prefix the source chose -- which is a guarantee no
re-serialisation can make. It replaced three `<container>[\s\S]*?</container>` regular expressions
in `theme-xml.ts`, the one file whose own header condemns that pattern, on a path a preserved source
theme's bytes reach. The package-level VBA edit (`edit-vba.ts`) removes a stale signature's
relationship and content-type override the same way, matching each element by its decoded attributes,
so a single-quoted `Id` or a differently cased `PartName` is not left naming a deleted part. It finds
parts through the reader's case-folding accessors and writes back under the package's own spelling.

A **reference** is a foreign scalar like any other, and has its own tolerant pair in
`core/address.ts`: `tryDecodeCellRef` and `tryDecodeRange`, returning `undefined` where
`decodeCellRef` / `decodeRange` throw. Read-side code uses those two and nothing else, so a
malformed `r`, `ref` or `sqref` costs the element that carried it and never the sheet: the cell is
skipped, and the validation, hyperlink, table, filter or note is dropped. They answer "can this
name something that exists", not merely "does this parse", because `A0` and a row past the last
parse cleanly and then throw at the grid, which is the same abort one step later. Bounds are part
of the question, so they are part of the answer. An axis a range omits stays unbounded rather than
unreadable: `A:A` is a legitimate reference, and a caller needing a bounded rectangle checks the
corners itself.

That leaves each feature to decide what "drop" means for it, and both readings are in the tree.
A table is dropped _whole_ when its `ref` is unreadable, for the `<cfRule>` reason: the anchor is
the coordinate every other field is read relative to. It is dropped whole too when its column count
differs from the `ref` width, since the model lays columns out one per `<tableColumn>` from the anchor
and a missing one would shift every column after it. A `sqref` is dropped _per area_, because its
areas are independent and one unreadable area says nothing about its neighbours. Where the
authoring path shares the code, the guard stays on the authoring side and the reader filters before
reaching it, so `addDataValidation` still refuses a range naming no cells while a file carrying one
simply loses that validation.

## How a failure is reported

Every error the library raises deliberately descends from `XlsxError` (`src/errors.ts`), so one
`catch` clause answers "was that us?" without naming a class. Two levels of branch sit under it,
chosen so neither is redundant with the other: `code` is the _kind_ of failure, `name` (and
`instanceof`) is exactly which one. Several classes share a code on purpose. A code in 1:1
correspondence with the classes would carry nothing the class did not already carry.

| code                 | what the caller does about it                                      | classes                                                                                                                         |
| -------------------- | ------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------- |
| `unsupported-format` | try a different reader, or reject the input                        | `UnsupportedFormatError` (its `format` field says _which_ unsupported input)                                                    |
| `malformed-input`    | reject the file, which is either broken or hostile                 | `PackageReadError`, `XmlParseError`, `XlsxParseError`, `XlsbParseError`, `CsvParseError`, `VbaParseError`, `CustomUiParseError` |
| `authoring`          | fix the calling code, which described a document that cannot exist | `AuthoringError`, `VbaAuthorError`                                                                                              |
| `internal`           | report it, because an invariant of ours did not hold               | `InternalError`                                                                                                                 |

Scalar argument validation stays outside the taxonomy: an index out of range, an unparseable
reference, a value of the wrong type are native `RangeError` / `SyntaxError` / `TypeError`, because
that is what those types are for. The line is composite against scalar: `getColumn(0)` is a
`RangeError`, while a table that names a column twice is an `AuthoringError`. A layer that wraps a
lower-level failure passes it as `cause` rather than flattening it into the message, except where
the lower layer's text is itself the hazard (a zip library's message can name an absolute path, so
`sniff-format.ts` replaces rather than wraps it).

The reader is held to the other half of that line: **a file is never allowed to raise an
`authoring` failure, and never a native one.** A model method validates what it is given, and a
reader that hands it a file-derived name is asking the model to judge the _file_ through a guard
written about the _caller_. `src/io/read-policy/read-repair.ts` is the single seam where that is
resolved, for both codecs,
and it offers exactly two answers. `repairSheetName` rewrites a name the way Excel repairs it, for
the case where the rule is a naming rule and dropping the thing would cost more than the name (a
sheet takes its cells, its position, and every `localSheetId` indexing past it with it). `admitting`
runs the call and answers `undefined` where the model refuses, for the case where the name _is_ the
identity: a table called `1 bad`, a defined name with no name. It catches `AuthoringError`,
`RangeError` and `SyntaxError` and nothing else, so an `XlsxError` from a layer below and an
`InternalError` of ours both keep their identity, and it is the only place on the read path that
swallows a refusal at all, which makes the set of constructs a corrupt file may silently lose a list
of call sites rather than a habit.

The same rule binds the scanner, one layer down: **the reader may only produce values the writer can
serialise.** XML 1.0's `Char` production is stated once, in `src/xml/xml-chars.ts`, below both halves
of the codec so that neither imports the other; the writer refuses what it names and the reader
declines to decode it. A character reference to an unrepresentable code point is left as the `&#1;`
the file wrote, on the same grounds as one naming no code point at all; a raw one has no verbatim
form to keep and is dropped. Without that bound a hostile file read cleanly and threw
`AuthoringError: cannot write U+0001 at offset 2` on the _next save_, which is the same taxonomy
violation arriving by a longer route.

### What an audit of the read path has already checked

A read-only audit of `src/` in September 2026 looked for the classic parser failures and found none.
Recorded here so the next one spends its budget elsewhere, and so that a change to any of these is
recognisable as a change to a property somebody checked rather than to an implementation detail.

- **Entity expansion and XXE are structurally absent**, not mitigated. The scanner knows five
  entities by name, leaves an unknown one verbatim, and skips a `<!DOCTYPE>` by balancing brackets.
  There is nothing to expand.
- **Path traversal in a relationship target is clamped.** `resolveRelativePart` pops `..` without
  going below the root, for an absolute target as for a relative one, and a part is a key in a map,
  never a filesystem path.
- **The inflate bound consults no declared size.** The counter aborts on bytes actually produced, so
  a header that lies about its uncompressed length buys nothing.
- **`<dimension ref>` and `<row spans>` drive no preallocation.** Both are hints a file chooses, and
  a reader that sized an array from one would let a few bytes of XML ask for an arbitrary allocation.
- **The attribute scan is linear.** It was not until September 2026: `parseAttributes` was a global
  regex that backtracked over a run of name characters no `=` followed from every position inside
  the run, so one junk token in a tag cost the square of its length, and a megabyte of one letter,
  a kilobyte zipped, cost minutes in either reader. It is now a hand-written scan that resumes where
  a dropped token ended and finds a closing quote with one `indexOf`, which bounds every character to
  a constant number of visits. `xml-scan.test.ts` counts those visits and holds the scan to the
  regex's answers.
- **Formula text decoded from BIFF12 is bounded.** In XML a formula is never longer than the bytes it
  came from; in BIFF12 a five-byte `PtgName` or 3-D reference stands for a name or sheet of any
  length, and a 3,599-byte stream citing a 1 MiB name 600 times threw a native `RangeError` out of
  `readXlsb`. Every decoded result passes through one `push`, which drops a formula past four times
  Excel's 8,192-character limit and keeps its cached value. A byte above `0x7f` names no token, rather
  than being masked onto an operand as it was.
- **Attribute coercion is one vocabulary.** No hand-rolled `attr === '1'` or bare `Number(attr)`
  survives on the read path; `xml-attrs.ts` is the whole of it, and its answer to an unreadable value
  is always `undefined`.

There is deliberately no "not implemented yet" code. Every candidate turned out to be an
unreachable exhaustiveness guard, and the one real feature gap, that a binary `.xlsb` cannot be
row-streamed, is already reported through `UnsupportedFormatError`'s `format` branch.

That gap makes the two read entry points answer the same bytes differently, and the asymmetry is
deliberate. Handed an `.xlsb` package, `readXlsx` dispatches to the BIFF12 codec and returns a
workbook, because both codecs build the same model and a caller who asked for a workbook gets one.
`readSheetRows` refuses, because row streaming is built on the XML worksheet parser and the binary
cell table has no streaming path: the honest answer to "stream me these rows" is that this format
cannot be streamed yet, not a silent buffer of the whole workbook under a streaming name. The two
share their package opening (`openSpreadsheetPackage`), so the divergence is one branch on whether
the package carries an XML office document, stated once in each of the two functions that hold a
different opinion about it.

## Two serialisations, one model

`.xlsb` is not a second library bolted on; it is a second codec over the same `Workbook`. The two
formats share an OPC/ZIP container, a relationship graph, and a style model, and differ only in how
the office-document parts are spelled: XML in `.xlsx`, BIFF12 record streams in `.xlsb`. The code
follows that split exactly, and the directory layout states it. The bounded inflater, magic-byte
probe and OPC/relationship resolution live in `src/io/opc/`, and the resolved-format table
(`XfStyle`, its built-in number formats, `applyXfToCell`, and the cell → row → column → xf 0
resolution every cell reader drives) in `src/io/style/`. Both sit _above_
the codecs rather than inside either. Only the part parsers live apart in `src/io/xlsb/`:
`record-stream.ts` (the record framing), `primitives.ts` (RkNumber, length-prefixed strings,
colours), `formula.ts` and `ptg-functions.ts` (the Ptg token stream a binary formula is stored as,
decoded to the text `<f>` would have carried), then a per-part parser mirroring its XML counterpart.

`readXlsx` detects which serialisation a package holds and dispatches, so a caller never branches on
format. The property that keeps the two honest is asserted, not assumed: the corpus reads one workbook
Excel saved in _both_ forms and requires the two models to be identical. Anything the binary states
that XML omits must therefore be dropped on the binary side: a bottom vertical alignment, a locked
cell, a row restating the sheet's default height, a hatch fill's automatic-colour sentinels. That is
where most of the subtlety in that reader lives.

Namespace URIs and ext-URI GUIDs are registered once, split by which layer owns them: the
package-level URIs (`.rels`, content types, the relationship vocabulary) in
`src/io/opc/namespaces.ts`, the SpreadsheetML and extension ones in `src/io/xlsx/namespaces.ts`.
Every relationship the writer generates is recorded in a `RelationshipLedger`
(`src/io/xlsx/package-plan.ts`) at the moment its id is handed out: one ledger per owning part (each
sheet, each drawing, the workbook, the package root, and each link of a pivot's part chain). The
ledger returns the id, and its record is the `.rels` part, so a part with an empty ledger has no
`.rels` at all. Id prefixes were once re-derived by hand-summing every prior part's count, which
silently collides two parts onto one id when a prefix drifts; later the ids came from a counter while
the relationships were listed again, in another order, by a separate renderer. Ids are now unique by
construction, never recomputed by arithmetic, and never listed twice. `assertRelationshipsWired` in
the xlsx test support checks the result on every package the round-trip helpers write: each cited id
is declared, each relationship reaches a part, and each is cited or found by its type.

The public API is eight curated entry barrels under [`src/entries/`](../src/entries/), one
per subpath the package publishes (`/core`, `/xlsx`, `/xlsb`, `/csv`, `/node`, `/vba`,
`/customui`, `/errors`), plus [`src/index.ts`](../src/index.ts), which unions seven of them so
the bare package name still carries everything a browser can run. Each symbol is listed in
exactly one entry, so the root barrel is a union of `export *` lines rather than a second list
to keep in step.

`/node` is the one the root barrel leaves out, and the reason is its imports rather than its
size. It carries the streaming writer, the only public surface reaching a Node built-in
(`node:fs`, `node:stream`); unioning it would put those on every browser consumer's graph for a
symbol they never named. `scripts/check-browser-safe.ts` walks the graph from the other seven
entries and fails on any `node:` specifier or Node-only global, and a `browser` condition in
`package.json` resolves `/node` to a stub whose classes throw by name
([ADR-0040](decisions/0040-the-browser-boundary-is-an-entry-point.md)).

That disjointness is load-bearing rather than tidy: `export *` does not report an ambiguous
re-export, it silently drops the name, so a symbol exported from two entries would vanish from
the root specifier with no diagnostic anywhere. `scripts/check-entries.ts` (a gate in
`verify --full`) fails the build on a duplicate, on an entry `package.json` does not publish, and
on a published subpath whose module is gone. It is also why the _whole_ failure taxonomy is
exported from `/errors` and nowhere else: a container-level failure belongs to no single codec,
since `readXlsx` and `readXlsb` both raise `UnsupportedFormatError`, so putting the classes with the
codecs would have forced exactly the duplication the union cannot survive. That entry costs 3 KB,
so classifying a failure never loads a parser.

### The size budgets are per entry, and they are a `verify` gate

`scripts/size-budget.ts` measures the whole emitted `dist/**/*.js` against a total, and each
published subpath against its own budget by walking that entry's static imports to a closure. The
per-entry numbers are the ones that catch anything: the total cannot see a codec acquiring a value
import of something it previously needed only as a type, or the model reaching into a parser, since
neither changes the number of bytes in the tarball. The closure is a lower bound on any bundler's
answer, because `sideEffects: false` lets a bundler prune _within_ those modules and never add to
them, which is what makes an over-budget reading a statement about the module graph rather than
about minification.

It runs inside `verify --full`, not only under `prepublishOnly`, and that placement is the whole
lesson of the one breach that happened. `/customui` is a ribbon reader that wanted three symbols
from the XML layer; every traversal helper added to the module those three lived in was charged to
its closure, and it sat 3 KB over a 16 KB budget across a release while every gate a developer or a
pre-push hook ran stayed green. A budget only the publish step checks is not a tripwire, it is a
surprise. Measuring it needs an emitted artifact, so the gate builds first
(`scripts/build.ts`, the single definition of the build that `package.json` also names); the build
is a few seconds and runs inside the corpus gate's shadow.

Budgets are tripwires, not targets. Raise one deliberately, with the same eyes a dependency addition
would get, and say in the commit _what_ the entry gained. The alternative the `/customui` breach
actually called for was the module boundary rather than the number: see
[ADR-0004](decisions/0004-xml-read-path.md)'s 2026-08-30 update.

The barrels are curated, not exhaustive: modelled core-feature types (autofilter, page setup,
sheet views, defined names, image options) are public, and internal helper functions stay off
them. `src/vba/index.ts` and `src/customui/index.ts` are _internal_ barrels that the model and
the codecs import; the public `/vba` and `/customui` faces are deliberately narrower. The two
halves of the streaming API are symmetric in reach but asymmetric in what needed naming.
Everything the streaming _writer_ exposes is public: its workbook/worksheet/row handles are
classes, its options are interfaces, and nothing structural is left un-named. The streaming
_reader_'s per-row/cell output stays inferred-structural rather than a named commitment.
Streaming is not its own subpath: measured, it reaches every module `/xlsx` does plus three, and
an entry that costs what the codec costs is an alias rather than a boundary.

`"sideEffects": false` is declared, and it is true. No module in `src/` mutates anything at
import time, so a bundler may drop an unused one whole. `scripts/size-budget.ts` measures each
subpath's static-import closure against its own budget, which is what notices a boundary being
crossed: a codec acquiring a value-import of something it previously needed only as a type moves
one of those numbers and leaves the package total untouched. `scripts/smoke-dist.ts` resolves
every subpath through the package name, which is self-reference and the only thing in the repo
that exercises the `exports` map at all, then asserts `/core` reaches no serialisation and
`/errors` reaches nothing but itself.

Those numbers say one thing about the model that is worth stating rather than leaving to be
rediscovered: roughly half of `/core`'s weight is the VBA codec and the ribbon parser, because
`core/workbook.ts` imports `parseVbaProject`, `addVbaReference` and `removeVbaModule` at _value_
level. The layering rule permits that deliberately, since a workbook models a VBA project and its
ribbon, so those types are part of the document, and the measurement is what shows the price of
letting the operations live there too. It is left alone knowingly. Removing it means a split the
model does not import at value level, which breaks `Workbook`'s VBA methods, and it would pay off
only for a consumer importing `/core` with no bundler at all; with a bundler, `sideEffects: false`
already prunes what such a consumer never calls. Revisit it when that consumer turns out to exist
rather than on the strength of the number alone. Until then `./core`'s budget is what notices the
weight growing further.

The [API reference](api/README.md) is generated straight from the root barrel, so it cannot
describe a shape the compiler wouldn't accept.

## The website

`www/` is a VitePress site, and this tree is its source. `www/docs/` is a generated mirror of
`docs/`, git-ignored and rewritten whole on every build, so a page is authored once, here,
where the gates already read it. `docs/docs.json` carries the reading order; a group either
names its pages or declares a `tree` that takes the rest of a directory, and every `.md`
under `docs/` has to be claimed exactly once. Titles come from the first `#` heading and
descriptions from the first paragraph that reads as a sentence, so no page carries
frontmatter.

Drift between the two is a build error rather than a warning, which is the whole reason the
mirror is generated: an unreachable page, a manifest entry with no file, a link that resolves
to nothing, a missing heading, or markdown the site cannot compile each fail
`node www/scripts/check.ts`, which runs in the `invariants` gate. Every runnable `ts` block
under `docs/guide/` is executed by `www/scripts/check-samples.ts` in the same gate, against
this commit's own `src/` rather than against whatever version happens to be installed.

The site is served at <https://shbernal.github.io/ts-xlsx/> and is generated in CI, never
committed: `.github/workflows/site.yml` builds it on every pull request and publishes it only
from `master`. ADR-0039 records the arrangement, including the size budget declined on
purpose.

## The VBA subsystem

Macro-enabled workbooks (`.xlsm`/`.xltm`) carry their VBA as a single opaque part,
`vbaProject.bin`, an OLE2 / Compound File ([MS-CFB]) container holding Office
run-length-compressed ([MS-OVBA]) module source and p-code. OOXML treats it as a binary
blob referenced by a workbook relationship; the ZIP package is otherwise identical to a plain
`.xlsx`. `src/vba/` is the self-contained, dependency-free subsystem that reads that blob and
applies pure-TS structural edits natively, with no `fflate` and no runtime dependency, just bytes.
Authoring or editing module _source_ is not here: it needs genuinely compiled p-code only a real
Excel can emit, so it lives in the offline `tools/vba-compiler` (see below, and ADR 0019).

**Preservation is the safety floor, and every VBA feature is additive over it.** Read captures
`vbaProject.bin` (and its relationship/content-type closure, including a sibling
`vbaProjectSignature`) as a `PreservedWorkbookReference`, and the writer re-emits those bytes
verbatim, so a load/edit/save of an `.xlsm` keeps its macros with no VBA-specific code on the
common path. Everything below layers onto that guarantee; none of it can desync the two
representations, because a _read-and-not-re-authored_ project is still emitted from the
preserved bytes alone.

The subsystem is built as encode/decode pairs over two formats plus a project layer:

- **CFB container.** `cfb.ts` reads and `cfb-writer.ts` writes. The reader walks the
  header, FAT, directory and mini-FAT, and reconstructs the whole storage/stream tree so the
  edit path can re-emit it with one stream swapped. The writer emits a v3 container whose
  storages are the name-ordered balanced red-black tree a _navigating_ host (Excel) needs, not
  just the linear scan our own reader would accept. Streams are found and replaced by storage path
  (`VBA/dir`, the root `PROJECT`), never by bare name, because a module and a UserForm's storage can
  each repeat a name another storage already uses.
- **MS-OVBA compression.** `ms-ovba.ts` (`decompressContainer` / `compressContainer`) is the
  chunked copy-token/literal-run codec Office uses for module source and the `dir` stream. It
  is _not_ deflate. The compressor's contract is that its output re-expands byte-for-byte.
- **Project layer.** `project.ts` decodes the `VBA/dir` and module streams into a typed view;
  `project-editor.ts` splices structural edits (remove module, add reference) into an existing
  project by splicing bytes, never by decoding and re-encoding a stream; `dir-records.ts` holds the
  `dir` record grammar, the walk that reads it and the one function that writes a record; and
  `vba-encoding.ts` holds VBA name validation.
  `codepage.ts` handles the project code page (MBCS, not latin1) in both directions, and
  `errors.ts` holds the two failure types. There is deliberately no from-source synthesizer;
  see below.

**The reader is hostile-input-facing (CLAUDE.md §3).** Every CFB sector index, chain, and
stream size is bounds-checked and cycle-guarded; the header's layout fields (both shifts, the
mini-stream cutoff) are checked against the values [MS-CFB] fixes rather than believed, because a
crafted cutoff routes every stream through the other allocator and both are bounds-checked enough
to hand back bytes nobody wrote. Every MS-OVBA back-reference is validated and total output is
bomb-capped. A malformed project fails closed with `VbaParseError`, never a crash, hang, or
unbounded allocation, and each guard is pinned by a crafted-malformed fixture.

The authoring side fails closed with `VbaAuthorError` on a contract violation (over-long or
duplicate stream name, unrepresentable character) rather than emitting a silently broken container.
It is _not_ fed only our own bytes: `removeVbaModule` and `addVbaReference` re-encode a `dir` stream
that arrived in the file, so the encoder is bounded in time like a parser, and the invariants the
editors patch against (MODULES_COUNT preceding every module block, and being non-zero) are checked
rather than assumed.

The public API layers by fidelity and intent, each slice fail-closed:

- **Read.** `Workbook.vbaProject: VbaProject | undefined` parses the preserved bytes _lazily_
  and memoises. Modules expose `name`, `streamName`, `kind` (the full procedural / document /
  class / designer classification), and decompressed `source`. A read never perturbs what the
  writer emits (ADR 0016).
- **Attach, replace, strip.** `Workbook.vbaProjectBytes` is a get/set accessor pair over the
  raw blob. A set is validated by `parseVbaProject` _before_ any state change, so a bad blob is
  rejected whole, and replacing or removing drops the old bytes' now-stale signature. This is also
  how an authored project is installed: the offline `tools/vba-compiler` produces a compiled
  `vbaProject.bin`, and a consumer attaches it here, in pure TS with no Office at runtime.
- **Structural edits, pure TS.** `removeVbaModule` and `addVbaReference` (with their `Workbook`
  and package-level `editXlsxVba*` wrappers) splice the original `.bin`. They edit only the `dir`
  stream (and, for a removal, `PROJECT`/`PROJECTwm`) and leave every module stream and
  `_VBA_PROJECT` byte-for-byte. They are safe _precisely because_ they never touch a module's
  compiled p-code: the `dir` stream, authoritative for the module/reference list, carries the
  change (ADR 0018/0019).
- **Author or edit module source, offline and not in the shipped library.** This is done by
  `tools/vba-compiler`, which drives a real headless Excel through the VBIDE object model to emit
  genuinely compiled, source-matched p-code (a `vbaProject.bin`, or a whole edited `.xlsm` for
  document code-behind). It is a Windows+Excel build tool, never in CI, whose output seeds
  committed fixtures (ADR 0019).

**Why source authoring cannot be pure TS.** Excel does not recompile VBA from source on open. A
module runs the compiled p-code (PerformanceCache) it ships, and the source is only recompiled
when a human opens the VBE. A `vbaProject.bin` synthesized from source alone, with no p-code or
mismatched p-code, either throws "Invalid data format" or _silently runs stale code_, so "opens
clean" is not proof of correctness. Only a real Excel can produce runnable p-code, hence the
offline compiler. Authoring artifacts are verified with `execute-verdict.ps1`, which opens with
macros enabled and _runs_ a known authored macro, not merely by an open verdict (ADR 0019).
_Executing_ macros in-process stays out of scope forever. That needs a live host, and this is a
document tool, not a VBA interpreter (ADR 0013).

## Tech decisions

The stack is deliberately small and each choice is recorded as an ADR under
[`docs/decisions/`](decisions/):

- **Runtime and no-build dev path.** ADR-0001.
- **Toolchain.** oxlint for rules and oxfmt for layout, with type-aware rules from tsgolint,
  ADR-0036; `node --test` over Vitest, hand-rolled type-level
  tests. ADR-0029. The `test/` and `scripts/` trees are TypeScript held to the same strict bar
  as `src/`, gated by `typecheck:test`, ADR-0011. All of those gates run as one concurrent,
  content-cached command (`node scripts/verify.ts`), ADR-0022.
- **Zip and XML write path.** `fflate`, plus a hand-written SAX reader with bounded allocation
  on every parser path. ADR-0003.
- **Docs generated from the types.** ADR-0006.
- **The website generated from the docs tree.** ADR-0039.
- **Spec reference.** Vendored OOXML schemas plus Microsoft Learn MCP. ADR-0007.
- **VBA subsystem.** Read view, ADR-0016; pure-TS structural edits, ADR-0018/0019; source
  authoring moved to the offline `tools/vba-compiler` after the "recompile cookie" premise was
  retracted, ADR-0019 (ADRs 0017/0018's from-source mechanism is retracted). Excel as a test
  oracle, ADR-0013.

## Working agreements

- **Preserve provenance as knowledge, not as a link.** Capture the real-world scenario a
  bug taught us. That survives; an upstream issue number does not. Durable artifacts
  (corpus cases, spec notes, commit messages) never cite upstream numbers; the commit
  that lands a change is its account of record.
- **Security- and correctness-first.** Every parser path is hostile-input-facing: no
  unbounded allocation, no zip-bomb naïveté. Entities are decoded but never expanded;
  inflation is bounded by a running output counter, not any declared size.
- **Work is bounded by what a file contains, not by what it declares.** The declared-size rule
  above has a CPU twin, and it is the one that keeps being rediscovered: a loop written over a
  declared _region_ costs whatever the region says, and a region is free to be the whole grid. The
  shape is always the same. `<col min max>` spans the sheet 16,384 columns at a time and any number
  of elements may do so, which is what `ColumnRecordBudget` bounds; `<mergeCell A1:A1048576>` is a
  few bytes that used to cost thirty milliseconds each, so 16,384 of them (~570 KB of XML, a few KB
  zipped) were eight minutes of CPU. Neither is visible to the inflate cap, because after inflation
  the payload really is small. Prefer the rewrite over the budget where one exists: walking the
  populated cells and testing each against the rectangle is a strict improvement for legitimate
  files too, and needs no number anyone has to justify. Where the work is genuinely proportional to
  a declared count, it gets a named budget with its reasoning attached, and a test that counts the
  work rather than timing it.
- **A hostile file reaches the encoders too.** "Our own bytes" is a property of a value's _origin_,
  not of the direction it is travelling: an edit reads a foreign file, mutates a structure inside
  it, and writes it back, so every re-encode on an edit path is fed bytes the caller did not write.
  `Workbook.removeVbaModule` recompresses a `dir` stream out of a `.xlsm`, which is why the MS-OVBA
  encoder carries a time bound (a hash chain over three-byte prefixes rather than a rescan of the
  whole back-window) and why the `dir` editors _check_ the record ordering they patch against
  rather than asserting it in a comment.
- **A part is found through the relationship that names it, never by its conventional path.**
  `xl/workbook.xml`, `xl/sharedStrings.xml`, `xl/styles.xml` and the rest are where Excel puts those
  parts, not where OPC says they live: the package names its office document in `_rels/.rels` and its
  pool and stylesheet in the workbook's own `.rels`, and a conforming producer may point them
  anywhere. The conventional path stays as a _fallback_ for a package whose rels graph is damaged,
  never as the first question. What makes this worth a rule rather than a bug report is how the two
  halves fail: a workbook part not found is loud, while a pool not found reads every `t="s"` cell as
  the empty string and a stylesheet not found leaves every cell without a number format, which
  changes a date cell's _type_, silently, because the date test reads `numFmt` off the resolved style.
- **A Strict package is read into Transitional at the container, because every write is Transitional.**
  ISO/IEC 29500 Strict is the same vocabulary under `purl.oclc.org` namespaces, with two relationship
  types renamed and DrawingML percentages written `60%` where Transitional stores `60000`. The parts
  the model reads are matched by local name and written from the model, so Strict never reaches them;
  a part carried through as bytes kept its Strict spelling, though, inside a package whose
  relationships said Transitional, and the Open XML SDK could not load such a theme at all.
  The translation happens at the two places a package's own spelling enters the model: a relationship
  Type as the container's `parseRelationshipRecords` reads it (`io/opc/strict-relationships.ts`), and
  a preserved part's bytes as the XML reader captures them (`io/xlsx/strict-parts.ts`: namespace
  declarations, a graphic frame's `uri`, and the percentage attributes, re-rendering only the tags
  that change). The second is the XML codec's rather than the container's because a binary workbook
  has no Strict form, and the size budget measured what `/xlsb` would have paid for the XML machinery
  it needs: 11.5 KB. Both tables come out of the schema graph rather
  than from memory: the namespace pairs are every vocabulary declared under both profiles, and the
  percentages every attribute whose Strict type is a DrawingML, chart or diagram percentage, with a
  chart's in whole percent and the others in thousandths, as Transitional types them. Nothing past the
  container needs to know a package was ever Strict, which is why the reader's own spelling check for
  `extendedProperties` could go.
- **Part-name lookups fold ASCII case; the package's own spelling is what gets written back.** OPC
  compares part names case-insensitively, so `/XL/Workbook.XML` and `/xl/workbook.xml` name one part,
  and a package cased differently from the references to it used to read as a package with no
  workbook at all. The fold lives once, at the boundary: `packageAccessors` tries the exact key and
  falls back to a folded map, and `contentTypeResolver` folds its `<Override PartName>` keys the way
  it already folded `<Default Extension>`. It is a _lookup_ rule only. The inflated part map keeps
  the package's own spelling, which is the name a preserved part is re-emitted under, so folding
  never rewrites what a round trip hands back.
- **A name inside an error message goes through `quoted()`.** `src/errors.ts` exports it, every
  layer may import it, and it is the only spelling: the tree had grown three (`"…"`, `'…'`,
  `JSON.stringify`), split by directory rather than by intent, and only the last survives a name that
  itself contains a quote. A sheet may be called `Q1 "draft"` and a part path may hold a newline; under
  either literal spelling the message that reports it is one whose reader cannot tell where the name
  ends, which on a library reading untrusted input is a decoy rather than a diagnostic. Messages stay
  lowercase throughout, sentence punctuation and all. `scripts/check-error-messages.ts` is
  the mechanism, because the convention is invisible at the throw site (both spellings produce
  identical output today, so review cannot see the difference) and a rule with no mechanism decayed
  into nine inline calls, two of them in a file that did both and one of those four lines from a
  `quoted()` in the same function. oxlint ships no `no-restricted-syntax`, so the gate is a script:
  every `JSON.stringify` in `src/` is either `quoted`'s own implementation, `core/range.ts`'s
  replacer form, or a finding. `errors.ts` owns the message _skeletons_ for the same reason it owns
  the spelling -- `invalidToken` for the OOXML-enumeration refusal, `unrepresentable` for a character
  the target format cannot encode -- because both of those were being thrown on two sides of a
  layering boundary and the response had been to transcribe the sentence.
- **A lookup table on a parser path is a `Map`, or an object with no prototype.** A plain object
  literal indexed by a string the file supplies answers about a dozen attacker-chosen keys with a
  _function_: `constructor`, `toString`, `valueOf`, `hasOwnProperty` and the rest of
  `Object.prototype`. Every miss check spelled `?? ` or `=== undefined` then reads that as a hit.
  This is not prototype pollution (nothing is written), it is the read side of the same mistake, and
  it lands where the format's own vocabulary is open-ended: an entity name, a media extension, a
  `PROJECT` keyword, a relationship type's final segment, a part path resolved out of a relationship
  target. Prefer the `Map`, which makes the property impossible rather than merely absent;
  `Object.create(null)` is for the accumulator that must stay an object because its type is an index
  signature (the SAX reader's attribute map, the inflated part map). A table whose keys come from a
  fixed regex character class or from our own writer is not on this path and stays an object literal.
  No rule in this toolchain gates it: oxlint carries no `security/detect-object-injection`
  equivalent, and `no-prototype-builtins` catches the `obj.hasOwnProperty(k)` call rather than the
  `obj[k]` read. So this paragraph is the gate, and the corpus's
  `undefined-entity-in-cell-text-stays-verbatim` is the lock on the site that reaches furthest.
- **On a round-tripping surface, ask whether input is safe to _write back_.** Bounding what a
  parser will hold is only half of it. What the reader accepts, the writer re-emits, so a value
  a foreign part carries can leave _our_ output invalid, and one invalid attribute is enough for
  Excel to offer to repair the feature away. A wire bound therefore lives in the parser, the
  authoring verb, _and_ the serialiser, so the serialiser cannot emit an illegal value however the
  model was populated (`MENTION_OFFSET_MAX` is the worked example; see
  `knowledge/specs/threaded-comments-and-the-legacy-fallback.md`).
- **Two parts that are two halves of one representation derive from one variable.** Some features
  are only coherent as a pair. A threaded comment's conversation part and the legacy fallback
  `<comment>` that binds a cell to it are each _invisible in Excel_ without the other, even though
  either alone validates clean. The writer computes such a pair from a single source (`write.ts`
  derives both from one `threads` value) so neither can be emitted without the other by
  construction, rather than by a rule someone has to remember.
- **A part family graduates from preserved to modeled in one change, never both at once.**
  Byte preservation (`core/preserved.ts`) is the sole emission authority for what it covers, so a
  _read view_ over preserved bytes is safely additive (`Workbook.vbaProject`, `Workbook.customUI`,
  `loadedPivotTables`). But the moment a serialiser emits a part from the model, that rel type
  must leave `isPreservedSheetRelType`/`isPreservedWorkbookRelType` in the same commit or the
  package carries it twice.
- **Narrow foreign tokens; never trust them into the model.** An enumerated attribute
  read from a file is admitted only through a type guard that recognises the known union
  members (the pattern is `isCustomFilterOperator` in `core/autofilter.ts`); an
  unrecognised token is _dropped_, leaving the facet unset, rather than cast in with `as`.
  The reader's posture is _skip, never guess_: a malformed token yields absence, not a
  bogus value the rest of the code will trust. See ADR-0004 for the read path this serves.
  Nothing gates this the way `check-layering.ts` gates the import graph, and the closest
  candidate is declined on the record: `no-unsafe-type-assertion` finds 50 in `src/`, 45 of
  them `arr[i] as T` escapes that `noUncheckedIndexedAccess` and the ban on `x!` force, and it
  cannot tell one of those from a token. This rule is enforced by review, so a guard is what a
  reviewer looks for; the enumerations each carry one (`isCustomFilterOperator`, `isVisibility`,
  `isThemeColorSlot`, `isDeclarablePivotSourceKind`), keyed off a `Record` over the union where
  the compiler can then refuse an omitted member.
- **No half-migrations on main.** Each change is fully green, meaning typed, linted, tested and
  corpus-passing, and leaves the tree better than it found it.
