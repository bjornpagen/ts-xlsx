# Streaming writer must deliver bytes to the sink incrementally

Cluster: streaming

## Scenario

A server streams a large workbook to an HTTP client by handing the streaming writer the response
object as its output stream, setting `Transfer-Encoding: chunked`, and committing rows. The intent
is that the browser receives the file progressively: first bytes soon after the request, then
chunk after chunk as rows are committed. What is reported instead: the client sits at "0 B" for a
long time and then fails with a network error, and nothing arrives until, or unless, the whole
document has been generated. It works for small data, which fits in buffers and flushes at the end,
and fails for large data, which either exhausts memory first or exceeds the client's patience or
timeout before any byte is sent.

> Spec note, not a corpus case: "bytes reach the client progressively over real HTTP" is a
> latency and delivery property that a durable in-process behavior case cannot faithfully assert,
> since there is no client, no wire, and no time-to-first-byte to measure without a flaky timing
> harness. The requirement is recorded here; a soak or latency check belongs in a dedicated perf
> harness.

## Desired behavior

- Committing rows to a streaming writer whose sink is a live writable, an HTTP response or a socket,
  must **push serialized bytes downstream as it goes**, so a consumer observes a low
  time-to-first-byte and a steady flow, not a single flush after the entire package is assembled.
- Incremental delivery must not depend on the caller inserting manual event-loop yields; correct
  documented usage, awaiting each row commit and respecting the drain signal, is sufficient. This is
  the delivery-latency face of the same writer-to-sink contract whose _memory_ face is
  `streaming-writer-row-commit-backpressure`.
- Where the OOXML and zip container genuinely force some trailing structure to be written last, such
  as the zip central directory, that unavoidable tail must be small and bounded, and the bulk of each
  sheet's row data must stream as produced rather than being buffered whole.
- Back-pressure and progressive delivery compose: a slow client throttles the producer, giving bounded
  memory, while still receiving a continuous stream, giving bounded latency.

## Open questions

- How much of the perceived "nothing until the end" is the zip layer buffering entries against the
  writer ignoring drain? A reduction that streams a single large sheet to a slow sink would separate
  the two.
- Should the writer document, and test in a perf harness, a time-to-first-byte target for a live
  sink, analogous to the bounded-memory target for the streaming reader?
- Is a flush hint needed at part boundaries, at the end of each committed sheet, so consumers see
  steady progress, or does respecting drain suffice?

Related: `streaming-writer-row-commit-backpressure`, `streaming-writer-per-sheet-memory-release`,
`browser-safe-io-boundary`.
