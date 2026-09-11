# The filesystem-path reader is Node-only and must fail clearly elsewhere

Cluster: streaming

## Scenario

A developer wants to open a spreadsheet in the browser, or another non-Node runtime, and reaches for
the convenience entry point that takes a filesystem path, the most discoverable "open a file"
method. In the browser there is no filesystem: the underlying path and stat machinery is unavailable,
so the call blows up deep inside the library with an opaque internal error, an undefined-property read
on a file-existence flag, that gives no hint the real problem is "this method only works under Node."
The correct approach is to obtain the file's bytes, say from a file input via an `ArrayBuffer`, and
hand them to the buffer-loading API, which is the environment-neutral entry point.

> Spec note, not a corpus case: the defect is a discoverability and error-quality cliff plus a design
> choice between absent-in-browser-builds and a guarded runtime error, not a malformed-output behavior.
> The durable value is the constraint and how the API should surface it.

## Desired behavior

- The path-based read entry point is understood as inherently **Node-only**, since it needs filesystem
  access. Invoked without that capability it fails fast with a **clear, actionable** error that
  names the constraint and points at the fix, such as _"reading from a filesystem path requires a
  Node.js environment; in the browser, load the file's bytes with the buffer-loading API instead"_,
  never an internal undefined-property dereference.
- The **buffer and array-buffer loading path is the documented, primary way to open a file** so users
  do not default to the path reader. Correct use is the easy path. An analogous buffer loader
  exists for CSV.

## Open questions

- Should the path reader be **entirely absent from browser builds**, tree-shaken or not exported, so
  the mistake is a compile-time type error rather than a runtime crash? For a TypeScript-first
  library this is the cleanest outcome, with the type surface exposing filesystem readers only under
  Node.
- If it remains present at runtime, what is the canonical guard? Feature-detect the fs capability
  up front and throw a named error type?
- Documentation: ensure the buffer-loading entry point is prominently the primary "open a file"
  method.

Related: `load-accepts-arraybuffer-and-typed-arrays`, `image-by-filename-is-node-only`,
`public-types-node-stream-portability`, `write-buffer-return-type-contract`.
