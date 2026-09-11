# ADR 0004: The XML read path is a lean, hand-written SAX pull parser

**Status:** Accepted (2026-07-12) · Phase 3 reader slice

## Context

ADR 0003 stood up the writer on `fflate` and emitted XML by direct string assembly,
deliberately **deferring the XML _parser_ choice to the reader slice**. That slice is
now here: almost every remaining corpus capability (`roundtripWorkbook`,
`readFixtureReport`, styles/defined-name/merge read-back, roughly 48 cases) asserts through
a package that is _read back_, so the model cannot light them up without a reader.

The deferred decision was `fast-xml-parser` (a DOM-building dependency) against a lean
hand-written SAX. Two forces from the constitution (`CLAUDE.md`) decide it:

1. **A spreadsheet reader parses untrusted input.** "No unbounded allocation, no
   zip-bomb naïveté, no eval-shaped surprises." A DOM parser materialises the _entire_
   part as an object tree before the reader sees a single cell, which is an allocation
   multiplier an attacker controls, exactly the shape we must not ship.
2. **The dependency tree stays small, modern, and clean**, a founding reason for the
   fork. A parser is not a dependency we want to own transitively when the subset of
   XML that OOXML uses is small and regular.

There is also a build-order force: the eventual **streaming** reader (large files,
`.eachRow`) fundamentally needs a pull/SAX model, not a DOM. Building SAX now produces the
same component that path will extend, not throwaway work.

## Decision

- **The reader is a lean, hand-written SAX pull parser** (`src/xml/xml-scan.ts`; see the
  2026-08-30 update below, and note it lived at `src/io/xlsx/xml-read.ts` when this was written),
  no XML-library dependency. It emits open/text/close events in a single O(n) pass with
  no recursion; the OOXML reader (`src/io/xlsx/read.ts`) consumes them and builds only
  the model, so peak memory tracks real content, not document structure.
- **Entities are decoded, never expanded.** Only the five predefined entities
  (`&amp; &lt; &gt; &quot; &apos;`) and numeric character references (bounded to valid
  Unicode) are recognised. DTDs and `<!ENTITY>` definitions are ignored outright, so
  _billion-laughs_ entity expansion and XXE external-entity resolution are structurally
  impossible. Not mitigated, absent.
- **Inflate is bounded** at the reader's entry (`readXlsx`): a cap on total declared
  uncompressed size rejects the naïve zip bomb before the parser runs.
- **`fflate` serves both directions.** Reading uses `unzipSync` (with the size filter);
  the write path already used `zipSync`. One zip library, as ADR 0003 set out.

## Consequences

- **Positive:** zero new dependencies for the read path; the hostile-input guards live
  in code we own and test; the SAX component is what the streaming reader will extend;
  escaping (write) and entity-decoding (read) are each one small audited place.
- **Negative / deferred:** we own XML edge cases (attribute quoting, CDATA, comments,
  processing instructions, `xml:space`), covered by unit tests.
- **Revisit when:** a real-world file exercises an XML construct the lean parser does
  not cover (record it as a corpus fixture first).

## Update (2026-07-13): the inflate bound no longer trusts declared sizes

The original slice bounded inflation by the zip's _declared_ uncompressed size, and flagged
the gap: a header-lying bomb (declares small, inflates large) slips past, and worse,
trusting the declared size to preallocate lets an attacker force a large allocation from a
few compressed bytes. Both are now closed by `src/io/opc/inflate.ts`: the package is fed to
fflate's streaming unzip in bounded slices, the decompressor grows its output from the bytes
it _actually_ produces, and a running counter aborts the moment real output crosses the cap.
Declared sizes are consulted for nothing. `maxUncompressedBytes` now bounds produced output,
not header claims. This is the first slice of the streaming reader; the same streaming-inflate
component is what an eventual row-streaming (`.eachRow`) read path extends.

## Update (2026-07-13): the pull parser and the streaming row reader

The SAX parser was push-only: `parseXml(source, handlers)` drove callbacks over the whole string
in one loop. A callback cannot `yield`, so a reader that must _emit_ incrementally could not sit on
top of it. The scan loop is now extracted into a generator, `xmlEvents(source)`, that yields
`open`/`text`/`close` events; `parseXml` is a thin push adapter over it, so every existing call site
is byte-for-byte unchanged (the corpus proves it) while a pull consumer can now drive the parse.

On that component sits the first streaming _read_ API: `readSheetRows(data, options)`
(`src/io/xlsx/read-rows.ts`), a generator that yields one worksheet's rows in order as plain
`{number, cells}` records, retaining only the row in hand rather than materialising the whole
`Workbook`. Value decoding is shared with the buffered reader through one module
(`src/io/xlsx/cell-value.ts`), so a cell streamed one row at a time decodes identically to the same
cell read as part of a full workbook, and the divergence such a split would otherwise invite is
closed by construction. This slice still inflates the package whole (bounded as above) and reads
shared strings and styles as whole parts, both legitimately document-sized; a later slice can make
the inflate itself per-part lazy on the same pull primitive.

## Update (2026-07-20): the reader consumes the pull parser through a small helper vocabulary

The hand-written SAX was right, but the _consumption_ side had drifted into three habits worth
naming so they don't return: ~30 no-op `onText(){}`/`onClose(){}` stubs, ~20 inline
`=== '1' || === 'true'` boolean idioms in three incompatible forks, and every parser hand-coding
both a self-closing branch _and_ an on-close branch for the same element. The read path now consumes
events through a settled set of helpers built on the same primitive, so a parser states only what it
means:

- **`SaxHandlers.onText`/`onClose` are optional.** A handler declares only the events it consumes.
- **`openElements(xml, ...localNames)`** (`xml-read.ts`) is a pull generator over `xmlEvents` for the
  "scan opens, read attributes" readers, collapsing an accumulator-threaded-through-a-closure into a
  plain `for..of` that reads `attrs` directly.
- **`closeEmptyElements(events, names)`** expands a self-closing `<x/>` whose local name is in `names`
  into open+close, so a formatted-but-empty element commits once in `onClose`. It is a **per-name
  opt-in** (`ReadonlySet<string>`), deliberately _not_ a blanket "synthesize a close for every empty
  tag" mode: text-bearing elements (`<f/>`, `<v/>`, `<t/>`) commit _captured text_ on close, so a
  synthesized close would misread them. A self-closing shared-formula clone `<f t="shared" si=".."/>`
  would set `hasFormula` and be mis-read as a shared-formula _master_. Only elements safe to run
  on-close-when-empty are listed.
- **Boolean attributes** go through `boolPresent` / `boolStrict` / `boolTristate`; **numeric operands**
  through `coerceNumericLiteral` (strict decimal regex `^-?\d+(?:\.\d+)?$`, so a non-canonical
  spelling like `1E5` is kept _verbatim_, the round-trip-faithful and more conservative read of
  foreign input); a leading `=` through `stripFormulaEquals`, which returns _unescaped_ text, leaving
  the caller to escape for its target.

**Enumerated attributes are narrowed, never trusted.** A union-typed attribute token is admitted only
through a guard that recognises the known members; an unrecognised one is dropped rather than cast in
with `as` (see the _narrow foreign tokens_ working agreement in `docs/architecture.md`). This is a
behaviour change on _malformed_ input only. Valid tokens are unchanged, so the byte corpus stays
green, and the drops are pinned by unit tests.

**Generator gotcha (load-bearing).** A sub-parser that drains a slice of a _shared_ event generator
must pull with `.next()` (or a helper generator that `return`s), **never** `for..of` plus `break`.
Breaking out of a `for..of` calls the generator's `.return()`, which terminates the shared stream for
every later sub-parser. The style-table driver's `until(events, container)` is the canonical shape:
one `xmlEvents` pass, each section's sub-parser draining its own container's slice, no re-scan.

## Update (2026-07-20): one cell-gathering state machine, two finalisers

The earlier update shared value _decoding_ across the buffered and streaming readers. Cell
_gathering_, the per-cell `<c>`/`<f>`/`<v>`/`<is>` state each reader accumulated, was still
re-implemented twice and free to drift. It is now one class, `CellAccumulator`
(`src/io/xlsx/cell-accumulator.ts`, modelled on the `read-rich-runs.ts` run accumulator): both readers
drive the same `beginCell`/`beginFormula`/`setFormula`/`setValue`/`appendText` calls, so a cell
can never be _gathered_ differently depending on which reader saw it.

Finalisation, by contrast, deliberately diverges, and that split is the point. The buffered reader
calls `finalize`, which resolves shared-formula masters/clones and data-table cells and opens rich
`<r>` runs; the streaming reader calls `decode`, the plain-decode subset, which must **not** resolve
shared formulas (it surfaces the clone's own cached result) and must **not** open rich runs (a
streamed inline string flattens to text). `finalize`'s plain path is itself routed through `decode`,
so there is exactly one `RawCell`-build and one decode. Sharing gathering while forking finalisation
is what lets the two readers stay honest to their different contracts without duplicating the fragile
part. A malformed non-empty `<c r>` is handled in `beginCell` for _both_ readers, closing another way
the two could have diverged.

**It is dropped there, not thrown.** This paragraph said "throws" until 2026-09-04, and it was the
first thing an agent read before touching the path: it cited a behaviour the reader does not have and
must not have, so an agent trusting it would have "restored" a bug. A `<c>` whose `r` names no cell
that can exist (`A0`, `ZZZZ1`, `junk!!`) is treated as one with no address at all, which is the
stance every other unreadable foreign attribute gets, and the stance this whole document argues for:
a file this library did not write is allowed to be wrong, and losing one cell beats losing the sheet.
The thing that had to be shared between the two readers was the _decision_, not the throw, and it is:
`cell-accumulator.ts` makes it once.

## Update (2026-09-11): a streamed rich string keeps its runs

The update above has the streaming reader open no rich `<r>` runs, so a streamed inline string
flattens to its text. That covered the inline string and never the pooled one: the shared-string
reader reads runs for both readers, and Excel pools rich text, so the same runs streamed as
`"bold and not"` when a producer inlined them and as a `RichTextValue` when it pooled them. The
streamed value depended on an encoding choice the author never made, which is the drift the shared run
machine exists to prevent, and it broke the streaming reader's own promise of a value identical to
the one `readXlsx` produces.

The streaming reader now reads runs, so the flattening is gone rather than extended to pooled strings.
Extending it would have been consistent too, but it discards formatting and breaks consumers all the
same. `CellAccumulator` lost its `richRuns` option with it. `RunAccumulator` keeps `readRuns`, because
a note's body is plain text by model and is still read with runs off. The other divergence stands:
`decode` still surfaces a shared-formula clone's cached result rather than resolving it.

## Update (2026-09-04): two conventions here have no machine check

Both are held by review, and saying so is the point: a convention a reader assumes is gated is one
they stop checking.

**The `read-` prefix** (ADR-0030) is enforced by review and nothing else. It has slipped once
already: `rich-runs.ts` was read-only and unprefixed for as long as it existed, and is
`read-rich-runs.ts` now.

**`elementSubtrees` is a deliberate second scanner** over the same text, and the two must agree to the
character about where an element ends. That is why `markupAt` and `tagAt` are exported rather than
private: the classification is shared even though the walks are not. The seam that would make the
agreement structural is to give `xmlEvents` the source offsets it already computes, so
`elementSubtrees` becomes a consumer doing `source.slice(…)` rather than a parallel scan. That costs
two numbers per event on the hot worksheet path, so it is a measurement rather than a refactor, and
it has not been made. Until it is, the agreement is maintained by hand and this is where that is
written down.

## Update (2026-08-30): the scanner and the ways of driving it are two modules

The parser grew a second half. Beside `xmlEvents` there are now the pull generators a reader writes
a `for..of` over (`openElements`, `capturedText`), the push adapter and the multi-pass driver that
lets five readers of the worksheet part share one scan of it, the verbatim subtree capture
(`elementSubtrees`, itself a second scanner over the same text), and `TextCapture`. Those are ways
of _driving_ a scan, and they had accumulated in the same file as the scan.

They are now `src/xml/xml-read.ts`, and the scanner - the events, and the readings of one
attribute's text that go with them - is `src/xml/xml-scan.ts`. Nothing changed about the parse; the
seam is where a helper stops answering "what does this markup say" and starts walking a document.

What forced the question was a bundle budget. `/customui` is a ribbon reader that wants `xmlEvents`,
`localName` and `boolStrict` and none of the traversals, so every helper added to the far half was
charged to that entry's closure: it drifted 3 KB over its 16 KB budget on growth it will never call.
That is precisely the drift the per-entry budgets of [ADR 0023](./0023-subpath-entry-points-and-disjoint-barrels.md)
exist to surface, and the honest fix was the module boundary rather than the number. `elementSubtrees`
is why the scanner exports `markupAt` and `tagAt` rather than keeping them private: the two scanners
must agree to the character about where an element ends, which is the property the shared
classification was introduced to guarantee.

The split stopped one module short. The scanner kept the readings of an attribute's _value_ as well
as the events, and half of those are spreadsheet-flavoured rather than XML: `numInteger` and
`numFinite` with their floors, `coerceNumericLiteral`, `enumToken`, and the `_xHHHH_` cell-text
escape. They are now `src/xml/xml-attrs.ts`. The three OOXML booleans stayed with the scanner, and
that is where the line falls rather than at "everything that reads a value": a `<button enabled="0">`
in a ribbon part is the same boolean a `<sheetPr>` attribute is, so the booleans are not a
spreadsheet's question. It is also the only line that pays. With `boolStrict` on the far side,
`/customui` would have gone on loading the new module for that one import and the split would have
moved bytes without removing any, which is what the first measurement said. As drawn, `/customui`
dropped a kilobyte and went from 0.7 KB of headroom to 2.5 KB.
