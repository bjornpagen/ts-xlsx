# Bounded memory in the streaming writer: row commit must expose backpressure

Cluster: streaming

Reported as "Failure on writing large amount of data to excel even though using streams".

## Desired behavior

The streaming workbook writer must let callers export arbitrarily large datasets in bounded memory. Adding and committing rows one at a time and then committing the workbook should never grow heap usage proportional to the total number of rows.

## The problem observed in the reference implementation

Even with the streaming writer and `addRow(...).commit()` per row, plus `useSharedStrings` and `useStyles`, memory grows without bound and the process OOMs at large row counts, with reports ranging from tens of thousands of heavily-styled rows to millions of plain rows. A committed row is serialized and pushed into the internal zip or output stream, but committing returns synchronously and never waits for that stream to be consumed by the underlying sink. When the producer commits rows faster than the sink drains, the pending output buffers, and the row objects they still reference, accumulate. The independently-rediscovered workaround is to `await` a zero-delay timer every few rows, which yields the event loop so the stream can flush and keeps memory flat. That a manual yield fixes it is the diagnosis: the writer ignores backpressure.

## Contract this fork should provide

- Committing a row must expose backpressure. Either `commit()` returns a promise or awaitable that resolves once the row's serialized bytes have been accepted by, meaning drained into, the downstream sink, or the writer exposes the standard "wait for drain" signal so a producer loop can pause. A caller that awaits per-row commit must run in memory bounded by a small window, not by total row count.
- The writer must not retain references to already-committed rows or cells. Once a row is committed and flushed, it and its related objects, with style records already interned and shared strings already interned, must be eligible for garbage collection. Interned and deduplicated tables (shared strings, styles) may legitimately grow with the number of _distinct_ values, but not with the number of rows.
- No artificial event-loop yield should be required by callers to achieve bounded memory. The correct, documented usage, awaiting each commit or respecting the drain signal, must be sufficient on its own.
- Heavy per-cell styling must not defeat this: styles should be interned so that N rows sharing a style hold one style record, not N.

## Prior art / notes

- Node streams already define this via the writable `drain` event and the boolean return of `write()`. A modern implementation should surface that instead of forcing users to `setTimeout(0)`.
- Some users worked around it by chunked or batched DB reads, which reduces peak DB memory but does not fix the writer-side retention; others reported the yield-based workaround only delayed the crash. Both point to the same missing backpressure contract rather than a caller mistake.

## Open questions

- Should per-row `commit()` be awaitable by default, breaking the old fire-and-forget shape, or should backpressure be surfaced via a separate drain-await API? Given the fork's no-compatibility stance, an awaitable commit is the clearer, safer default.
- What is the target memory ceiling to assert in a benchmark, say peak RSS independent of row count for a fixed-width, fixed-style dataset?
- How should shared-strings and style interning growth be bounded or documented so users understand what memory _is_ expected to scale with?
