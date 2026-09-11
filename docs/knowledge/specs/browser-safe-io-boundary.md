# Browser-safe I/O boundary

Cluster: browser

Reported as "Cannot read property 'F_OK' of undefined".

> **Built, 2026-08-28.** The bundle-time half of this note is done and gated. `node:crypto` and
> `Buffer` are gone from the library entirely (`src/sha512.ts`, `toBase64`, `TextEncoder`), the
> streaming writer is published from `@shbernal/ts-xlsx/node` with a `browser` condition resolving
> it to a stub that throws by name, and `scripts/check-browser-safe.ts` walks the graph from every
> browser-facing entry on every run so no Node built-in or Node-only global can return to it. See
> ADR 0040. What remains open is the forward half named at the end of this note: a browser-native
> streaming sink over Web Streams (`browser-streaming-workbook-write`, `web-streams-io-surface`),
> which would let the writer itself run in a tab rather than being absent from it.

## Problem

The file-path I/O methods (read from or write to a path on disk) depend on Node's `fs` module, including `fs.constants.F_OK` for existence checks. When the library is bundled for the browser, `fs` is either absent or a bundler-provided empty stub, so `fs.constants` is `undefined`. Any code path that dereferences `fs.constants.F_OK` throws a cryptic `TypeError: Cannot read property 'F_OK' of undefined` from inside the minified bundle. The user has no way to tell that the actual issue is "filesystem paths are not supported in the browser."

This also affects downstream libraries that wrap the file-path API, such as render-from-template helpers that call a path-based writer under the hood: they surface the same opaque crash to their users.

## Desired behavior

- The library must have a clear, documented boundary between environment-agnostic I/O (buffers and streams) and Node-only path-based I/O.
- Buffer and stream based read and write, producing or consuming an in-memory `ArrayBuffer`, `Uint8Array` or stream, must work in the browser with no `fs` dependency reachable at runtime.
- Path-based methods (read or write by filesystem path) are explicitly Node-only. When invoked in an environment without a real filesystem, they must fail fast with an actionable error naming the constraint ("Filesystem path I/O is not available in this environment; use the buffer/stream API") rather than crashing on an internal `undefined` dereference.
- Ideally the Node-only path methods live behind an import or subpath that a browser bundle does not pull in, so `fs` never ends up in the browser bundle at all and the buffer path has zero Node built-in dependencies.

## Prior art

- Common pattern: a core module with no Node built-ins plus a thin Node adapter exposing `readFile(path)` and `writeFile(path)` layered on top of the buffer API.
- Bundlers used to auto-polyfill Node core modules (webpack 4) and now do not (webpack 5), which is why environment-dependent code that "worked" in older toolchains now surfaces as `undefined` module shims.

## Bundle-time requirement: no Node builtins reachable from the browser entry

The boundary must hold at **bundle time**, not just runtime. Bundlers (webpack 5, Vite, Rollup) no
longer auto-polyfill Node core modules, so any code statically reachable from the browser entry that
`import`s `fs`, `path` or Node-form `stream` makes the build fail with "dependency not found: fs",
and merely instantiating a workbook is enough to pull it in. The requirement:

- No code reachable from the browser entry point statically imports a Node-only module. Filesystem
  and stream conveniences that need Node live behind a **Node-only conditional export** the browser
  condition never resolves, so `fs` never enters the browser bundle and no bundler config is needed.
- The browser entry ships **first-class TypeScript types identical in quality to the Node entry**,
  which exist by construction from the source rather than as a separately hand-authored `.d.ts` that
  drifts.
- Correct `package.json` `exports` and `browser` conditions drive the split so consumers get the right
  API automatically.

## Runtime requirement: no unguarded Node-_global_ access on a hot path

The boundary is not only about `import`ing Node modules. It also covers touching Node-only _globals_
(`process`, `Buffer`, `global`) in code that runs in the browser. Upstream regressed here when a
runtime environment-detection helper read `process.version`, splitting the version string, on a hot
path reached by ordinary in-memory API use. In a browser `process` is undefined, so _merely creating a
workbook or setting a cell_ threw a cryptic `TypeError` ("Cannot read property 'split' of
undefined"), and the library was unusable in the browser on its first ordinary call, with nothing
pointing at the real cause. The requirement:

- **Core in-memory operations touch no Node-only global.** Building a workbook, adding worksheets,
  reading and writing cell values, and serializing to or from a buffer must not reference `process`,
  `Buffer`, `global`, `fs`, or `path` in a way that executes on import or on first API use where those
  are undefined. Feature and environment detection must be guarded (`typeof process !== 'undefined' && …`)
  or removed in favor of a build-time condition.
- If a value from a Node global is genuinely needed (a version gate, a `Buffer` fast path), it is
  accessed behind an existence guard with a browser-safe fallback, never assumed present.

The **streaming** reader and writer are the sharpest case of this boundary: they depend on Node `fs`
and `stream`, and their CSV code paths transitively pull `fs` too, so a browser bundle that reaches
them fails with "dependency not found: fs" and, at runtime, an undefined streaming namespace
(`stream.xlsx` giving _"Cannot read property 'xlsx' of undefined"_). The requirement: the streaming
symbols must be **absent from the browser entry, or throw a precise typed error** ("streaming write
is not available in this environment; use the document writer or a Web Streams sink"), never
present-but-broken. The forward path is either a browser-native streaming sink over Web Streams (see
`web-streams-io-surface`, `browser-streaming-workbook-write`) or a cleanly streaming-free browser
build; either way no Node core module leaks into the browser graph, including transitive CSV paths.

## Open questions

- Should path-based methods be a separate entry point (subpath export) or a runtime-guarded no-op that throws? A separate entry point keeps `fs` out of browser bundles entirely and is the cleaner break.
- What is the canonical browser write API name and return type: ArrayBuffer, Uint8Array, or something Blob-friendly?
- Should the environment guard detect the missing filesystem lazily, on call, or fail at import time in a browser build?

Related: `no-global-polyfill-in-browser-bundle`, `path-reader-is-node-only-clear-error`,
`public-types-node-stream-portability`, `write-buffer-return-type-contract`.
