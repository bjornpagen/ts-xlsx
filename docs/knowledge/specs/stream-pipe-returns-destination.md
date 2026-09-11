# The library's writable streams must honor the Node stream contract (pipe returns the destination)

Cluster: streaming

## Scenario

A caller drives the streaming writer and pipes its output into another stream, using the standard
Node idiom that chains off the destination:

```js
workbook.stream
  .pipe(fs.createWriteStream('/tmp/sheet.xlsx'))
  .on('finish', () => console.log('written'));
```

or composes it with `stream.pipeline(source, transform, destination, cb)`. Node's `Readable.pipe(dest)`
is contractually required to **return `dest`** so these forms work. The library's internal stream
buffer (`StreamBuf`) overrides `pipe` but returns `undefined`, so `.pipe(dest).on('finish', …)`
throws ("Cannot read properties of undefined") and `stream.pipeline` mis-wires, so the finish handler
is never attached and the write appears to hang or silently drop.

> The pipe-return-value defect is now pinned by the corpus case
> `streaming-writer-stream-pipe-returns-destination`: the streaming `WorkbookWriter` exposes its own
> output stream as a public property, so a case constructs the writer, pipes that stream into a
> `PassThrough`, and asserts `pipe(dest) === dest`, a baseline fail today since it returns
> `undefined`, while a control confirms the bytes still flow. This spec remains the broader design
> requirement: the durable value is that **every** stream object the public API hands back, not only
> the writer's output stream, behaves like a real Node stream, with backpressure, `finish`, `end` and
> `error` events, and `pipe` returning the destination, so the contract holds wherever the rewrite
> exposes a stream.

## Desired behavior

- **Every writable or readable stream the public API exposes conforms to the Node stream contract.** In
  particular `pipe(dest)` returns `dest`, so `writer.stream.pipe(out).on('finish', …)` chains and
  `stream.pipeline(...)` wires source to destination correctly.
- **Backpressure and `finish`, `end` and `error` events propagate** as a consumer expects from a native
  stream, so the streaming writer drops into existing Node stream plumbing without special-casing.
- Prefer building on the platform's stream primitives, or a thin, spec-faithful wrapper, rather than a
  bespoke buffer that re-implements and then diverges from the stream contract. See
  `public-types-node-stream-portability` for the type-surface half of the same portability goal, and
  `web-streams-io-surface` for the cross-platform stream direction.

## Open questions

- Which internal buffering the streaming writer needs is genuinely custom, and which is expressible
  with `stream.PassThrough` or `Transform`, whose contract conformance is free.
- Whether the rewrite exposes Node `Readable` and `Writable`, Web Streams, or both at the pipe site,
  which ties into `web-streams-io-surface` and `unified-streaming-and-buffered-io`.

Related: `public-types-node-stream-portability`, `web-streams-io-surface`,
`unified-streaming-and-buffered-io`, `streaming-writer-incremental-http-delivery`.
