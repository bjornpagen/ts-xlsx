# Writing to an in-memory buffer must return an honestly-typed binary container

Cluster: types

## Scenario

A TypeScript user writes a workbook to memory and receives a value whose declared type does not
match its runtime shape. The legacy surface declares the returned value as extending `ArrayBuffer`,
but at runtime it is a Node `Buffer` in Node, or a `Uint8Array` in the browser. The mismatch
forces every caller to cast, `as Buffer` or `as unknown as Uint8Array`, to do anything useful:
pass it to `Buffer.from`, write it to a stream, or hand its `.buffer` to a worker `postMessage`
transfer list. The misdeclaration is not merely noise: treating the value as a raw `ArrayBuffer`,
as its type invites, can corrupt the output, and transferring it between worker threads fails
because the declared type hides the real backing `ArrayBuffer`.

> Spec note, not a corpus case: the runtime bytes are already correct, and the defect lives entirely
> in the public type surface, so it is pinned by a type-level test (`expectTypeOf` or tsd-style) on
> the write-to-buffer signature rather than by a runtime behavior assertion. The durable value is the
> return-type contract and the design choice behind it.

## Desired behavior

- The buffer-writing API returns a value whose static type **precisely** matches its runtime value,
  with no cast required for ordinary use: wrapping in `Buffer.from`, reading `.byteLength`, handing
  `.buffer` to a transfer list.
- The return type is a real binary **view**, honestly typed. A platform-neutral `Uint8Array` is the
  strongest candidate: it is the common denominator between Node and browser, exposes `.buffer` for
  structured-clone and `postMessage` transfer, and wraps trivially with `Buffer.from(...)` in Node.
- The surface must **not** declare `interface Buffer extends ArrayBuffer`, or any equivalent. A
  `Buffer` or `Uint8Array` is a view _over_ an `ArrayBuffer`, not a subtype of it, and declaring
  inheritance breaks structural typing and invites the file-corrupting misuse above.
- A type-level test pins the exact return type so a future refactor cannot silently regress it back
  to an `ArrayBuffer`-shaped lie.

## The load side and modern TypeScript targets

The mirror-image contract applies to the **load-from-buffer** entry point, and both sides must hold
across modern TS library targets. Under ES2024+ the Node `Buffer` type became generic over its
backing storage (`Buffer<ArrayBufferLike>`), so a hand-authored `declare interface Buffer extends
ArrayBuffer {}` not only lies about the write return but also makes a plain `Buffer` read from disk
or the network fail to satisfy the load signature without a cast.

- The load API accepts **whatever a real Node buffer read is**: `Buffer`, `Buffer<ArrayBufferLike>`,
  `Uint8Array`, `ArrayBuffer`, and ideally `Blob` and `ArrayBufferView` where sensible, with no
  user-side casting, across CommonJS and ESM/NodeNext resolution.
- The write API returns a value the caller hands directly to `fs.writeFile`, `res.send`, and the like,
  without narrowing.
- The fix root: do not ship a bespoke top-level `Buffer` declaration. Type against standard
  library and `node:*` types and let `Buffer`'s own generic definition flow through, then pin both
  directions with type-level tests across the supported target matrix.

## Open questions

- Return `Uint8Array` uniformly across runtimes, which is portable and honest and keeps the Node
  `Buffer` global out of browser-targeted type surfaces, or a Node `Buffer` under Node for ecosystem
  parity? This note leans `Uint8Array`; Node callers wrap with
  `Buffer.from(result.buffer, result.byteOffset, result.byteLength)` without copying.
- If a Node-only convenience that returns `Buffer` is desirable, it should be a distinct, explicitly
  named entry point, not a runtime-dependent return type on one method.

Related: `public-types-node-stream-portability`, `load-accepts-arraybuffer-and-typed-arrays`,
`path-reader-is-node-only-clear-error`.
