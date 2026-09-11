# ADR 0016: The VBA project is readable through a typed view; authoring stays deferred

**Status:** Accepted (2026-07-22) · VBA read slice (Phase 1) · **Amended 2026-07-23 by ADR 0017 §2.3d.**
Decisions 2 and 5 below (preservation is the _sole_ emission authority; authoring is out of scope) no
longer hold. `Workbook.setVbaProject({modules})` now synthesizes a `vbaProject.bin` from module source,
making the workbook an emission authority for authored macros. The read view (decisions 1, 3, 4) is
unchanged, and preservation remains the emission authority for a project that was _read_ and not
re-authored. See ADR 0017 for the authoring design and its real-Excel verification.

## Context

Macro-enabled workbooks (`.xlsm`) carry their VBA as an opaque `vbaProject.bin`, an
OLE2/Compound File ([MS-CFB]) holding Office run-length-compressed ([MS-OVBA]) module
source and p-code. Until now the library handled it **only** opaquely: read captured the
part as a `PreservedWorkbookReference` (bytes, content type, relationship closure) and the
writer re-emitted it byte-for-byte, so a `.xlsm` survived a load/edit/save with macros
intact but **nothing could look inside the blob** through the public API. ADR-0014 and
`docs/knowledge/specs/xlsm-macro-preservation.md` named "expose the VBA project bytes to
callers, or only pass them through opaquely?" as an open question, and parked it pending
evidence that reading inside was small and safe.

A throwaway prototype in the sibling repo `../cyberbenchmark-analysis` supplied that
evidence: a dependency-free CFB plus MS-OVBA plus `dir`/`PROJECT` parser that decoded a real
10-module project byte-faithfully against oletools' `olevba`. The question was no longer
"is this feasible?" but "productionise it under the quality bar."

## Decision

1. **The VBA project is exposed as a read-only, typed _view_, not a new source of truth.**
   `Workbook.vbaProject: VbaProject | undefined` parses the already-preserved
   `vbaProject.bin` **lazily** and memoises it; `VbaProject` carries the project code page
   and `readonly VbaModule[]` (`name`, `streamName`, `kind`, decompressed `source`).
   Barrel-exported alongside `parseVbaProject` and `VbaParseError`.

2. **Preservation stays the sole emission authority.** There is **no** write path from
   `VbaProject` back to bytes. Editing a workbook re-emits the original blob unchanged, so
   VBA _read_ cannot regress macro preservation. It is strictly additive and cannot desync
   the two representations, because only one of them is ever serialised. This mirrors the
   worksheet invariant at `src/core/worksheet.ts` (`#preservedReferences` "stays its sole
   emission authority; this collection is never emitted").

3. **The VBA parser is treated as hostile-input-facing (CLAUDE.md §3).** Every CFB sector
   index, chain, and stream size is bounds-checked and cycle-guarded; MS-OVBA output is
   bomb-capped and every copy-token back-reference validated. A malformed project fails
   closed with a `VbaParseError`, never a crash, hang, or unbounded allocation. Each guard
   is a test with a crafted-malformed fixture.

4. **`kind` is the full four-way classification.** `MODULETYPE` only splits procedural
   (`0x21`) from non-procedural (`0x22`); the `PROJECT` stream's `Document=`/`Class=`/
   `BaseClass=` lines refine `0x22` into `document`/`class`/`designer`. Verified against a
   real project carrying both document and class modules.

5. **Authoring stays out of scope, exactly as ADR-0014 holds for charts/shapes/slicers.**
   Phase 1 is read-only. Writing source back into a valid `vbaProject.bin` (a CFB writer,
   the encode side of MS-OVBA, the full `dir`/`PROJECT` record set, and the
   recompile-from-source cookie) is a substantial multi-part feature with no forcing
   consumer today. Building it speculatively is the premature abstraction CLAUDE.md §4 warns
   against.

## Consequences

- **Positive:** the spec's "expose the bytes?" open question is closed _with evidence_.
  `workbook.vbaProject?.modules[*].source` reads macro source through a precisely-typed
  API, with zero new dependencies and no risk to the preservation guarantee.
- **Corrected on the way in, with a real bug found and fixed.** The plan had leaned toward an
  `isSigned` accessor on the assumption that a signed `.xlsm` lost its `vbaProjectSignature`
  part on round-trip. A reproduction disproved that: the signature is a sibling part reached
  from `xl/_rels/vbaProject.bin.rels`, so the closure walk already carries it. But the
  reproduction surfaced a _different_ real bug. The content-types writer collapsed all
  same-extension binary preserved parts to one `<Default>`, mis-typing a
  `vbaProjectSignature.bin` sitting next to a `vbaProject.bin` (both `.bin`, different
  types). Fixed in `workbook-xml.ts` (per-part `<Override>` for any preserved binary part
  whose type differs from its extension default) and corpus-locked in
  `preserved-parts.test.ts`. `isSigned` remains deferred for want of a consumer, but is now
  cleanly sourceable from the preserved closure if one appears.
- **Negative / deferred:** callers can _read_ macros but not _create or edit_ them. The
  forward map, in value-to-cost order, each gated on a forcing consumer:
  attach-an-external-blob authoring (a thin wrapper over preservation); first-class authoring
  (source to a valid `.bin`); a `customUI`/ribbon round-trip audit; and, adjacent rather than
  VBA, a formula-evaluation engine for callers who mean "recompute the numbers," not "run the
  macros." Executing VBA is **never** a library feature, since it needs a live host; ADR-0013
  frames Excel automation as a test oracle, not a runtime dependency.
- **Revisit when:** a concrete consumer needs to author or edit macros. Pick up the
  attach-blob path first, since it is cheapest, then first-class authoring as its own slice.
  No new ADR is needed unless the shape of the decision changes.
