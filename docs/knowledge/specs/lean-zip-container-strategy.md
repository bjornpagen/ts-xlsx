# A single lean ZIP container layer over native DEFLATE

Cluster: security

## Scenario

The `.xlsx` format is a ZIP of XML parts. A user bundling for the browser observes that the zip
toolchain (a JS zip reader plus separate archiver and unzipper libraries used server-side) dominates
bundle size, well over a hundred kilobytes gzipped, and drags in unmaintained transitive
dependencies with known advisories and unclear licensing. Modern runtimes expose native DEFLATE
streams, so the container layer could be a thin ZIP-framing shim over the platform rather than three
bundled third-party zip libraries. Replacing the rotting zip toolchain is one of the fork's founding
motivations.

> Spec note, not a corpus case: `roundtripWorkbook` and `readFixtureReport` already exercise the
> container end-to-end, so the _behavior_ is locked regardless of implementation. The durable value
> is the container strategy and its hard-input safety limits.

## Desired behavior

- The ZIP container layer, the only place raw DEFLATE and the ZIP central-directory format are
  handled, is a **single, small, internal module**, not three heavyweight external deps (a JS zip
  reader for the browser plus archiver and unzipper for Node streaming).
- Compression and decompression of entries **prefers the platform's native DEFLATE**: the standard
  `CompressionStream`/`DecompressionStream` Web API, available in modern browsers and current Node,
  for raw deflate, with Node's `zlib` as an equivalent fallback.
- The library owns only the **ZIP framing** (local file headers, central directory, data descriptors,
  CRC32, and ZIP64 where large parts require it) as typed, tested code, a bounded and well-specified
  amount, since the ZIP APPNOTE format is fixed. Native streams do raw DEFLATE but do **not**
  understand the ZIP archive format, so framing must be implemented.
- **One code path** for read and write (isomorphic) instead of divergent browser and server zip
  stacks. The goals served: a drastically smaller browser bundle, elimination of unmaintained,
  vulnerable or unlicensed transitive zip deps, and a single audited place.
- **Hostile-input bounded**: streaming decompression never trusts declared uncompressed sizes. Cap
  max entry count, max total inflated bytes, and max compression ratio for zip-bomb resistance, and
  stream-inflate on demand rather than buffering the whole archive.

## Open questions

- Ship our own ZIP framing, or vendor one audited, tiny, permissively-licensed helper, measured
  against bundle size and audit cost?
- Is a single isomorphic path (native streams everywhere, `zlib` only as a Node fallback) achievable,
  or is a thin runtime shim still needed?
- Exact hard-input limits (max entries, max inflated bytes, max ratio) enforced at the container layer.
- Confirm whether `CompressionStream` `'deflate-raw'` or `'deflate'` framing matches what ZIP entries
  require.
- The minimum runtime floor (a modern Node major plus evergreen browsers) implied by native streams is
  acceptable for a clean-break modern library.

Related: `bounded-memory-large-workbook-read`, `no-unsafe-eval-csp-compatible`,
`browser-safe-io-boundary`, `minimal-audit-clean-dependency-tree`.
