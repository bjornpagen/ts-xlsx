# ADR 0003: Zip container is `fflate`; the write path emits XML directly

**Status:** Accepted (2026-07-12) · Phase 3 writer slice

## Context

The core in-memory model landed (`value`/`cell`/`worksheet`/`workbook`). The next
build-order step is the `.xlsx` writer, because almost every remaining corpus
capability (`inspectPackage`, `roundtripWorkbook`, …) round-trips through a written
package, so the model cannot light those cases up on its own. Standing up the writer
forces two dependency decisions that `STRATEGY.md` left open:

1. **The zip container.** An `.xlsx` is an OPC zip. Legacy used `archiver` (write) and
   `unzipper`/`jszip` (read), three overlapping, dated packages, part of why we
   forked.
2. **XML.** Legacy used `saxes`/`sax` for reading and hand-rolled string building for
   writing.

## Decision

- **Zip: `fflate`.** A single, tiny (~8 kB), zero-dependency, actively-maintained,
  ESM-first library that does both deflate _and_ inflate. It replaces
  `archiver` + `unzipper` + `jszip` outright. The writer uses `zipSync`; the reader
  (later slice) will use `fflate`'s streaming inflate, which is where the
  hostile-input guards (bounded output, no zip-bomb naïveté) will live.
- **Writing XML needs no parser.** OOXML parts are emitted by direct string assembly
  through a tiny escaping/serialisation helper (`src/io/xlsx/xml.ts`). Correct
  escaping of text and attributes is the only hard requirement, and it is security-
  relevant (an unescaped `<`/`&`/`"` produces a malformed package a consumer rejects),
  so it is centralised and unit-tested rather than sprinkled inline.
- **The XML _parser_ choice stays deferred** to the reader slice (`fast-xml-parser`
  against a lean SAX). The writer does not need it, and deferring keeps this slice to
  exactly one new package (`fflate`).

## Consequences

- **Positive:** one modern dependency replaces three legacy ones; the same library
  serves both write and read; the write path has no XML-library dependency at all;
  escaping is one audited place.
- **Negative / deferred:** the reader's parser decision is still open (tracked for the
  next slice); `fflate`'s `zipSync` is synchronous, which is acceptable for the buffered
  writer, and the streaming writer slice can move to its async/stream API later.
- **Revisit when:** the streaming writer needs incremental/async zipping, or the
  reader slice benching shows a different inflate is materially better.

## Update (2026-07-20): the streaming writer, and a settled attribute-serialisation vocabulary

The deferred streaming _writer_ is built (`src/io/xlsx/write-stream.ts`):
`WorkbookStreamWriter`/`WorksheetStreamWriter`/`StreamedRow` author a package incrementally,
flushing each committed row and evicting its cell graph so peak memory tracks the rows in flight,
not the whole sheet. It uses `fflate`'s streaming `Zip`/`ZipDeflate` with incremental CRC-32, the
async zip API this ADR anticipated, and `commit()` settles on the sink's `finish`/`error` rather
than hanging. Everything it exposes is public (see the barrel note in `architecture.md`).
`useSharedStrings` is offered but **disables** eager per-row flushing, because a whole-workbook
string pool defeats the memory bound the streaming writer exists to provide. That is a deliberate,
documented trade rather than a silent one.

Direct string assembly stayed the right call, but the ad-hoc `? '1' : '0'` / `? 1 : 0` attribute
spellings scattered across the serialisers were unified into `boolAttr(name, value?)` and numeric
`attr(name, value?)` in `xml.ts`. Each returns a leading-space ` name="…"` or `''`, so an emitter
states an attribute once and absence costs nothing. One subtlety worth keeping: when two passes emit
paired markup for the same logical element (a data-bar rule's classic body and its `<extLst>` x14
extension), they must agree on the shared synthetic key by _identity_, not by walking the rule list
in lock-step. A `ReadonlyMap<rule, guid>` built once and read by both passes replaces the pair of
parallel counters whose only guarantee was that they iterated in the same order, which is the
fragile kind of coupling that survives until someone reorders one pass.
