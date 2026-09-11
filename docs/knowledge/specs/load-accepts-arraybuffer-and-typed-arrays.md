# The load API must accept ArrayBuffer and typed arrays, not only Node Buffer

## The scenario

In a browser a user reads a file input into memory and hands the bytes to the workbook
loader. The idiomatic browser paths (`File.arrayBuffer()`, `FileReader`'s
`readAsArrayBuffer`, `fetch(...).then(r => r.arrayBuffer())`) all yield an
`ArrayBuffer`, or a `Uint8Array` view over one. Passing that directly to the loader
fails with _"Chunk must be one of type String, Buffer or StringBuf"_, because the
load path only accepts a Node `Buffer` or string. The user must know to wrap the bytes
in `Buffer.from(...)` first, a Node-ism that has no place in browser code and that
drags a `Buffer` polyfill into the bundle.

This is the primary browser load path, so the friction hits nearly every browser
consumer.

## Desired behaviour

- The public load API accepts, uniformly, the byte containers a caller actually has:
  `ArrayBuffer`, `Uint8Array` and other typed-array or `DataView` views, Node
  `Buffer`, and `Blob` asynchronously. It normalizes to bytes internally.
- No caller needs `Buffer.from(...)` or any Node-only shim to load from bytes in the
  browser.
- The accepted input union is part of the typed public API, so a wrong input type
  is a compile-time error with a clear expected-type list, not a runtime string throw
  from deep in the zip layer.

## Root cause (legacy)

The loader funnels input through a Node stream and `StringBuf` write path whose chunk
guard admits only `string`, `Buffer` and `StringBuf`. An `ArrayBuffer` or `Uint8Array`
falls through to the guard's throw. There is no normalization step that turns arbitrary
byte containers into the internal representation.

## Open questions for the rebuild

- Exact accepted input union for the modern TS API, and whether to also accept a
  `ReadableStream<Uint8Array>` for streaming loads.
- Whether `Blob` support is first-class, with an async read, or left to the caller to resolve
  to an `ArrayBuffer`.
- How this composes with the planned zip-layer replacement: the new unzip layer should
  take `Uint8Array` as its native input so the normalization is a thin, zero-copy view
  rather than a buffer copy. The legacy `StringBuf` machinery is being deleted with the
  toolchain rebuild regardless.
- This is the input-type facet of the browser story; the filesystem-against-buffer boundary
  is covered separately in [[browser-safe-io-boundary]].
