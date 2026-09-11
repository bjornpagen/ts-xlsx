# ADR 0040: The browser boundary is an entry point, and a gate walks the graph to prove it

**Status:** Accepted 2026-08-28

## Context

Three Node built-ins were reachable from `src/index.ts`: `node:crypto`, through the
sheet-protection password hash in `src/core/protection.ts`; `node:fs` and `node:stream`,
through the streaming writer in `src/io/xlsx/write-stream.ts`. Reaching only `readXlsx`,
`writeXlsx` and `Workbook` still reached `node:crypto`, because `Worksheet.protect()` needs it.

None of that is a runtime problem in Node, and none of it is reached by anything a browser
would call. It is a _bundling_ problem, and the distinction is the whole point: a bundler
resolves imports, not call graphs. Webpack 5 stopped auto-polyfilling Node core modules, so a
browser build reported a module it could not resolve; Vite externalised the three with a
warning and a stub that throws only when called. Either way the failure arrived at a consumer
who had written nothing more exotic than `new Workbook()`.

This repository's own website is the case in point. `www/` bundles the library into a page,
and it could only do that by aliasing `node:crypto|fs|stream` to a hand-written module that
throws. That workaround lived in the site's Vite config, was documented at length in the
browser guide as "the gap you will hit today", and told every reader to write the same
workaround themselves.

`docs/knowledge/specs/browser-safe-io-boundary.md` had asked for the opposite arrangement
since the harvest: no Node built-in statically reachable from a browser entry, Node-only
conveniences behind an export condition a browser never resolves, and no Node-only _global_
(`process`, `Buffer`) on a path ordinary in-memory use reaches. Nothing had been built for it.

## Decision

**The boundary is drawn by imports, and it is an entry point.** `@shbernal/ts-xlsx/node` is a
new subpath entry carrying the streaming writer, which is the only public surface that needs a
Node built-in. `src/index.ts` unions every entry barrel except that one, so the root specifier
(the specifier nearly every consumer writes) reaches no Node built-in at all.

**The other two built-ins were removed rather than relocated,** because the code behind them
did not need Node, only a hash and an encoder:

- `src/sha512.ts` is a from-scratch SHA-512 replacing `node:crypto`. Web Crypto was the obvious
  candidate and is unusable here: `crypto.subtle.digest` is asynchronous and OOXML's agile
  hashing chains 100000 digests, so adopting it would have made `Worksheet.protect()` async for
  every caller. The salt still comes from the platform, through the `crypto.getRandomValues`
  that browsers and Node both carry. Measured, the spin loop costs about 0.8 s against
  `node:crypto`'s 0.4 s: a factor of two on one authoring call, for a library that bundles.
- `Buffer` left the CSV codec: `TextDecoder` on the read side, `TextEncoder` plus two small
  loops on the write side, byte-for-byte what `Buffer.from` produced (a test asserts exactly
  that, for every encoding). `CsvWriteOptions.encoding` was `BufferEncoding`, so the public API
  also stopped depending on `@types/node`, and stopped offering `base64` and `hex` as though
  they were output encodings for a text format.

**A browser condition resolves `/node` to a module that throws by name.** `package.json` maps
the subpath to `entries/node-unavailable.ts` under `browser`, whose classes throw
"WorkbookStreamWriter is not available in this environment: … use writeXlsx", while the `types`
condition still resolves to the real declarations. A browser build that imports the Node entry
therefore links something that says what happened, rather than failing on `node:fs` or dying as
`undefined is not a function` inside a minified chunk.

**The boundary is a gate, not a comment.** `scripts/check-browser-safe.ts` walks the module
graph from all eight browser-facing entries and fails on any `node:` specifier and on any
Node-only global, with the file, the line and the entry that reaches it. It also fails if
`/node` reaches _no_ Node built-in, because an entry that carries nothing has stopped being a
boundary. `scripts/check-entries.ts` gained two rules: the root barrel must not union `/node`,
and the browser stub's value exports must equal the entry's. `scripts/smoke-dist.ts` re-checks
the same property on the emitted JavaScript, which is the only form a consumer's bundler sees.

## Consequences

- **`WorkbookStreamWriter` moved.** `import {WorkbookStreamWriter} from '@shbernal/ts-xlsx'`
  no longer resolves; it is `'@shbernal/ts-xlsx/node'`. This is a breaking change and a
  deliberate one: the alternative is every browser consumer paying for a symbol they never
  named. Nothing about the class itself changed: the sink options, `writer.stream` and the
  pipe contract the corpus locks are all as they were.
- **ADR 0023 said streaming would not get its own entry point,** on the measured ground that
  its closure is the codec's closure plus three modules, so an entry costing what the codec
  costs is an alias rather than a packaging boundary. That reasoning was about _size_ and it
  still holds; this entry exists for a different reason, and the size figures confirm it is not
  a size boundary (`/node` is 358 KB against `/xlsx`'s 490 KB).
- **The site stopped needing a workaround.** The Vite alias and `www/node-absent.ts` are gone,
  and the browser guide's "the gap you will hit today, and what to do about it" section is
  replaced by "nothing to configure".
- **Every entry that reaches `/core` grew by about 5 KB,** the cost of carrying a SHA-512
  implementation and a base64 encoder instead of importing them. `/customui` went over its
  budget by 0.7 KB on the strength of that, and the budgets were re-baselined to the tenth of
  headroom their comment describes.
- **What is still open:** a browser-native streaming _writer_ (`browser-streaming-workbook-write`,
  `web-streams-io-surface`) is unaffected by this record. When one exists it belongs on `/xlsx`
  with the rest of the codec, and `/node` keeps whatever genuinely needs a filesystem.
