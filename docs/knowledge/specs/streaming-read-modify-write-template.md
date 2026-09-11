# Streaming read-modify-write: edit a large template without materializing it in memory

Cluster: streaming

## Scenario

A developer has a large pre-built spreadsheet that serves as a template. They want to load it, make
targeted edits (replace a cell value, append rows to an existing sheet, tweak a header), and write a
new file, all **without holding the whole workbook in memory**, using the streaming reader on the
input side and the streaming writer on the output side.

Today the streaming APIs are one-directional building blocks: a streaming reader that yields rows,
and a streaming writer that accepts appended rows. Nothing connects _reading_ an existing package to
_re-emitting_ a modified one via streaming. The only workaround is to fall back to the fully-buffered
in-memory path, reading the whole file, mutating, and writing the whole buffer, which defeats the
purpose and does not scale to the very templates that motivated the request.

> Spec note, not a corpus case: there is no streaming read-modify-write path to assert against, and
> the community workaround is exactly the buffered path this capability is meant to avoid. The durable
> value is the desired capability and its design tradeoffs.

## Desired behavior

- A supported way to **stream a package in, apply edits, and stream a modified package out** with
  memory bounded by the working set, the rows and parts currently in flight, not by the file size.
- **Passthrough by default.** Parts and rows the caller does not touch are copied through
  byte-for-byte, or re-emitted losslessly, rather than fully re-serialized, so an unedited sheet
  costs streaming I/O rather than a full model rebuild.
- **Targeted edits are correct.** Replacing a cell value, restyling a cell, or appending rows to an
  existing sheet produces a valid package; edits that change structure (insert or delete rows, add a
  sheet) are either supported with correct reference and shared-string fixups or rejected with a clear
  error rather than silently corrupting the file.
- **Workbook-scoped state stays consistent.** Shared strings, styles, and relationship ids referenced
  by passed-through parts remain valid even when some parts are rewritten, with no dangling style
  index or shared-string id.

## Prior art / observed workarounds

- Read the entire file into the buffered workbook, mutate, write the whole buffer. Correct but
  memory-unbounded, the anti-pattern this request exists to replace.
- Manual zip surgery, editing one entry and repackaging, preserves memory but is fragile: it bypasses
  shared-string and style reconciliation and breaks the moment an edit touches interned state.

## Open questions

- Granularity of the edit API: a row-transform callback over the streamed input, a
  sparse patch keyed by address, or a "reopen sheet N and append" handle?
- Which structural edits are in scope for v1: value and style patches only, or row insert and delete
  with reference rewriting?
- Passthrough fidelity against normalization: copy unedited parts verbatim, or re-emit them through
  the writer, risking incidental reformatting, to guarantee a single canonical output shape?
- Interaction with the memory-release model already captured for pure streaming writes.

Related: `streaming-write-per-sheet-memory-release`, `streaming-writer-row-commit-backpressure`,
`bounded-memory-large-workbook-read`, `streaming-read-resolves-shared-strings`,
`foreign-file-read-modify-write-preserves-validity`, `worksheet-clone-and-cross-workbook-merge`.
