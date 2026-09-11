# ADR 0024: Async is one writer, not a mirrored pair

**Status:** Accepted (2026-07-29) · I/O surface · the deferred `mtime` question was answered by [ADR 0032](./0032-package-output-is-reproducible.md) (2026-08-11)

## Context

`readXlsx`/`writeXlsx` are synchronous and the README sells that as a differentiator. But
`zipSync` on a large workbook blocks the event loop for seconds, which is a real problem for a
server generating a workbook per request, and `WorkbookStreamWriter.commit()` is already `async`,
so the obvious move was a symmetric `readXlsxAsync`/`writeXlsxAsync` pair over `fflate`'s
worker-backed API, keeping the sync functions as the default.

Measuring first changed the shape of the answer. Deflating a ~42 MB part map with `fflate`'s
callback `zip()` against `zipSync`, one process per case, `monitorEventLoopDelay` for the stall:

| workbook shape  | `zipSync` | `zip()` async | event-loop stall under async |
| --------------- | --------- | ------------- | ---------------------------- |
| one large sheet | 1408 ms   | 1417 ms       | max **17 ms**                |
| twenty sheets   | 1402 ms   | **581 ms**    | max 41 ms                    |

Under `zipSync` the loop does not tick at all for the whole duration. So the async writer buys
responsiveness always, and wall-clock only when there are several parts to deflate in parallel,
since one sheet means one worker means the same total time.

The read side does not mirror this, for two independent reasons.

**The compression is not the cost.** `readXlsx` inflates and then runs the SAX parse and model
build synchronously. Moving inflation to a worker leaves the majority of the work on the caller's
thread, so a `readXlsxAsync` would advertise a non-blocking read and then block for most of its
duration. That is a promise the library would not keep, the same reasoning that deleted
`UnsupportedFeatureError` before it shipped in the error-taxonomy work.

**It would regress the zip-bomb ceiling.** `io/opc/inflate.ts` derives its guarantee from feeding
compressed input in 16 KiB slices and checking a running _output_ counter between them, which bounds
worst-case overshoot to one slice's expansion (~16 MiB at DEFLATE's ~1032:1 ceiling). With
`AsyncUnzipInflate` the counter sits on the calling thread while a worker keeps producing, and
`fflate` exposes no way to enforce a cap inside the worker. The cap would become advisory with
unbounded lag, on the reader's primary hostile-input path.

One premise behind the symmetric pair also turned out to be wrong. `commit()` being `async` is not
evidence that the synchronous contract is already inconsistent: it awaits a Node `Writable`, for
backpressure and finish, which is I/O and not CPU offload. Streaming to a sink is inherently
asynchronous; deflating a buffer is not.

## Decision

1. **Synchronous stays the default and the documented shape.** `readXlsx`, `writeXlsx`, `readCsv`,
   `writeCsv` and the `.xlsb` reader are unchanged.

2. **Add `writeXlsxAsync`, and nothing else.** It shares `buildPackageParts` with `writeXlsx` and
   differs only in handing the part map to `fflate`'s worker-backed `zip()`. Every part compresses
   to identical bytes; the two archives differed only in the per-entry timestamp, and since
   [ADR 0032](./0032-package-output-is-reproducible.md) pinned that, not at all.

3. **No `readXlsxAsync`**, for the two reasons above. Callers who need a non-blocking read run the
   whole read in a worker, which is a documentation answer.

4. **The asymmetry is stated, not hidden.** `writeXlsxAsync`'s doc comment says why there is no
   read counterpart, because a reader who notices the gap will otherwise assume it is an oversight.

5. **Failures propagate exactly as the sync path's do.** `AuthoringError` from part-building arrives
   as a rejection rather than a throw; a zip-layer failure, including an environment that cannot
   spawn a worker, propagates unwrapped, as it already does from `writeXlsx`. Wrapping it in
   `InternalError` would have been false: that class documents "no caller can provoke one", and a
   worker-hostile environment is provocable.

## Consequences

- The published API gains one function and one asymmetry. A symmetric pair would have read
  better in the API reference and been worse in fact.
- The event-loop win is available to the buffered writer without touching the streaming writer,
  which already solves a different problem (bounded memory, not thread occupancy).
- Reading a large workbook still blocks. That is now a stated limitation with a stated workaround
  rather than an unexamined default.
- If a non-blocking read is ever wanted in-library, the path is a worker returning a
  structured-cloneable model, not an async inflater. There is a per-sheet `WorksheetModel` today
  but no workbook-level equivalent, so that is a feature, not a packaging change.

## Alternatives rejected

- **A symmetric `readXlsxAsync`/`writeXlsxAsync` pair.** The original plan. Rejected on the two
  read-side findings above: it would under-deliver on its central promise and weaken the zip-bomb
  bound to buy it.
- **Async-first, demoting sync.** Rests on the premise that the sync story was already broken by
  `commit()`, which conflates I/O-async with CPU-offload-async. It would also make the common
  small-workbook case worse for no gain.
- **Sync-only, workers as the whole answer.** Defensible, but it leaves a measured multi-second
  stall on the table for a five-line function over a boundary that was already clean.
- **Pinning `mtime` so the two writers are byte-identical.** Tempting while testing, but it changes
  `writeXlsx`'s output and belongs to a separate question, whether `.xlsx` output should be
  reproducible at all, which it is not today for the same reason. _(Answered yes by
  [ADR 0032](./0032-package-output-is-reproducible.md), on a consumer's evidence rather than a
  test's convenience; the two writers are byte-identical as a side effect.)_
