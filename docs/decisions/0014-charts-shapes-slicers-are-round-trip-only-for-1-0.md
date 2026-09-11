# ADR 0014: Charts, vector shapes, slicers, and form controls stay round-trip-only for 1.0

**Status:** Accepted (2026-07-21) · Phase 4 scope slice

## Context

Four package-part families the model does not parse, namely embedded charts, DrawingML vector
shapes, table/pivot slicers, and legacy form controls, are all handled the same way
today: `PreservedPart`/`PreservedWorksheetReference`/`PreservedWorkbookReference`
(`src/core/preserved.ts`) capture their bytes, content type, and relationship closure
verbatim on read, and the writer re-emits them unchanged with relationships rewired to
fresh, collision-proof part paths. A load that only edits unrelated cells reproduces the
chart, shape, slicer, or control faithfully. None of the four can be **authored**. There
is no API to add a new chart, draw a shape, or declare a slicer from scratch, and a workbook
that never had one still can't get one from this library.

This is not an oversight. Four backlog specs already sit under `docs/knowledge/specs/`,
written during the harvest, each describing the first-class authoring model in detail:
`embedded-chart-read-write.md`, `drawingml-shapes-authoring-and-roundtrip.md`,
`form-controls-roundtrip-preserved.md`, and the passthrough half locked by
`chart-parts-survive-template-roundtrip.md` / `preserve-drawing-shapes-on-roundtrip.md`.
Each spec explicitly separates "passthrough preservation" (done) from "first-class model"
(not started) and lists open design questions, including preset-geometry scope,
text-in-shapes, combo-chart support, and the shared color/anchor abstraction, that have no
answer yet because no consumer has forced one. Slicers have no dedicated spec but follow the
same generic mechanism as charts and shapes and the same reasoning applies.

ADR-0005 already set precedent for a closely related question, whether to carry
attached-part byte closures on `WorksheetModel`, and declined it while deferring the
_authoring_ question generally: "a faithful whole-sheet/package copy primitive is
deferred pending a consumer... it waits for a real use-case to force its shape." This ADR
extends that same reasoning from the model-copy API to the public authoring API.

## Decision

1. **Round-trip fidelity is the 1.0 bar for charts, shapes, slicers, and form controls.**
   A workbook containing any of the four, loaded and re-saved with unrelated edits, must
   reproduce them byte-faithfully. This already holds today via the preserved-parts
   mechanism and is corpus-locked.

2. **Native authoring of the four is out of scope for 1.0, not merely unscheduled.**
   Building any of them is a substantial, multi-part feature (a typed model, an authoring
   API, a serializer, and per each spec's open questions real design work) with no
   forcing consumer today. Building it speculatively would be exactly the premature
   abstraction CLAUDE.md §4 warns against: designing a shape/chart vocabulary nobody has
   asked to use yet.

3. **The backlog specs stand as the scoping for when authoring is picked up.** They are
   not superseded by this ADR and remain the design starting point. Picking up any one
   of them is a normal feature slice at that point, not a decision that needs a new ADR
   first.

4. **The public API must make the boundary discoverable, not implicit.** A caller
   inspecting a loaded workbook can already see `Worksheet.preservedReferences` and
   `Workbook`'s preserved workbook references, with the `relType` naming what was preserved
   (chart, drawing, slicer, vbaProject, …), so a caller can detect "this workbook has a
   chart I can't touch" programmatically. What was missing was the same information
   stated in prose: the README and migrating-from-exceljs guides did not say this
   explicitly, leaving it to be discovered by absence. Both now state it (see
   Consequences).

## Consequences

- **Positive:** the 1.0 scope line is explicit and defensible rather than an implicit gap
  someone has to reverse-engineer from a missing method; the backlog specs are affirmed
  as live design documents rather than orphaned notes; no speculative chart/shape API
  ships ahead of a real consumer.
- **Negative / deferred:** a caller who needs to _create_ a chart, shape, slicer, or form
  control cannot do so with this library at 1.0, because round-trip preservation is the
  ceiling. This is a real capability gap against desktop Excel and against what some ExcelJS
  users relied on ([exceljs/exceljs#141](https://github.com/exceljs/exceljs/issues/141), "Chart
  support", 33 reactions at the 2026-07-09 fork-time capture).
- **Revisit when:** a concrete consumer needs to author one of the four. At that point,
  pick up the matching spec under `docs/knowledge/specs/`, resolve its open questions
  against that consumer's actual need rather than speculatively, and build it as a normal
  feature slice. No new ADR is required unless the shape of the decision changes.
