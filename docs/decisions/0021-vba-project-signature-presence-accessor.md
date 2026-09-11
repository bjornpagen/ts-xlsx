# ADR 0021: The VBA project signature is readable as _presence_, not verified as _validity_

**Status:** Accepted (2026-07-24) · VBA signature read slice · completes the `isSigned` accessor
deferred by ADR 0016 (§Consequences) and `docs/knowledge/specs/xlsm-macro-preservation.md`'s open
questions.

## Context

A code-signed VBA project carries one or more sibling signature parts reached from
`xl/_rels/vbaProject.bin.rels`: the legacy `vbaProjectSignature.bin`, and its `agile` (V2) and `V3`
successors ([MS-OFFMACRO2]; V3 added by KB5000676 to close a tampering hole). These are **already**
preserved byte-for-byte. The closure walk (`capturePartClosure`) is purely relationship-driven, so it
carries any part reachable from `vbaProject.bin` regardless of relationship type, and the content-types
writer keeps each part's own type (ADR 0016 §Consequences fixed the `<Default>`-collapse bug that had
mis-typed a signature part). The bytes are also correctly **dropped** whenever the project is
replaced or edited (ADRs 0017/0018), because a signature over the old bytes cannot vouch for new ones.

What was missing was a _read_ of that captured state. ADR 0016 left an `isSigned` accessor "cleanly
sourceable from the preserved closure once a consumer needs it." Per CLAUDE.md §3, exposing a read of
already-captured data is a bounded, low-risk slice rather than speculative infrastructure, so it is built
now without waiting for a named consumer.

## Decision

1. **Presence, not verification.** `Workbook.vbaProjectSigned: boolean` reports whether a signature
   blob is _attached_; `Workbook.vbaProjectSignatures: readonly VbaProjectSignature[]` exposes each
   signature's generation (`'legacy' | 'agile' | 'v3'`) and its raw bytes, passed through verbatim.
   The library does **not** parse the PKCS#7/CMS structure, validate the certificate chain, or extract
   signer identity. That is a materially larger effort (ASN.1/CMS, trust-store semantics) with real
   security-review weight and no consumer. The TSDoc and the naming say "a signature is attached,"
   never "this signature is valid." A caller who needs cryptographic validation hands `bytes` to an
   external verifier.

2. **A read over the preserved closure, with no memoisation and no new source of truth.** The accessor
   walks the VBA project's preserved reference: it finds each signature part reached by a signature
   relationship off `vbaProject.bin` and returns a defensive copy of its bytes. It is computed on each
   access rather than cached, because the closure is small and already in memory, and recomputing
   sidesteps a cache that every signature-dropping mutation (`vbaProjectBytes` replace, `removeVbaModule`,
   `addVbaReference`) would otherwise have to invalidate. Consequently the drop behavior (ADRs
   0017/0018) is now observable in-memory: a replaced project reads `vbaProjectSigned === false`.

3. **Detection keys off the relationship Type's final path segment, not the full URI.** The three
   generations carry different year segments (`.../office/2006/...`, `.../2014/...`, `.../2020/...`),
   but their final segments (`vbaProjectSignature`, `vbaProjectSignatureAgile`, `vbaProjectSignatureV3`)
   are stable and are what name the generation. Matching the segment is year-agnostic, so no exact
   URI needs pinning, and it mirrors the codebase's existing suffix-matching idiom
   (`relType.endsWith('/vbaProject')`). A future scheme this map does not know stays **preserved but
   unreported**, a conservative failure rather than a wrong positive, and preservation is unaffected
   either way.

4. **Multi-generation is in scope because the mechanism already generalises.** Since the closure walk
   preserves every generation for free and segment-matching recognises all three, `vbaProjectSignatures`
   reports each one rather than narrowing to legacy-only.

## Consequences

- **Positive:** the spec's `isSigned` open question is closed. `wb.vbaProjectSigned` answers "is it
  signed?" (matching Excel's own `Workbook.VBASigned`), and `wb.vbaProjectSignatures` hands the raw
  blobs and generations to a caller who wants to verify externally, with zero new dependencies and no
  risk to preservation, since it is a pure read and the corpus shows zero writer-path drift.
- **The verification boundary is explicit.** Presence against validity is stated in the API TSDoc, this
  ADR, and the spec note, so the accessor cannot be mistaken for a trust decision.
- **Test coverage.** Legacy detection is fixture-verified end-to-end (read, then accessor). Agile and V3
  are exercised via synthetic fixtures wired with their real 2014/2020 relationship URIs, proving the
  year-agnostic segment match across generations. The corpus carries no real Office-produced agile or V3
  file, so that remains the one un-pinned edge, bounded by the fact that preservation carries such parts
  regardless of whether the accessor names them.
- **Revisit when:** a consumer needs actual cryptographic verification (CMS parse, chain validation,
  signer identity). That is a distinct, security-heavy slice and gets its own ADR. This one
  deliberately stops at presence.
