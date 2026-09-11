# ADR 0005: `WorksheetModel` stays a semantic value; attached parts are out of scope

**Status:** Accepted (2026-07-18) · Phase 3 core slice

## Context

`Worksheet.model` is the sheet-to-sheet transfer contract: `dst.model = src.model`
snapshots a sheet and reproduces it elsewhere. Its documentation claimed the snapshot
was "everything that defines the sheet" and that the round-trip "drops nothing", a
promise that exists to prevent a historical silent-data-loss bug (a getter field the
setter ignored).

That promise was false. The getter carried the sheet's value and overlay content
(cells, styles, column/row/page metadata, merges, data validations, conditional
formattings, tables, protection) but silently omitted every _attached part_: anchored
and background images, the autofilter, authored pivots, loaded pivots, and the
byte-preserved references (charts, vector drawings, slicers) captured for
round-tripping. A `dst.model = src.model` on a loaded sheet quietly dropped all of them.

The parked question was whether to make the model _faithful_ by carrying those parts,
concretely whether to put opaque preserved-part **byte closures** (and loaded-pivot
records) onto the public model so the "drops nothing" claim became true.

## Decision

1. **`WorksheetModel` remains a semantic, serialisable value type.** It carries pure
   sheet state; it does not carry attached parts that hold _workbook-level identity_.
   Those parts cannot be made correct inside a value snapshot: image bytes live on the
   `Workbook`, a pivot's source references a live `Worksheet` instance, and preserved
   parts are opaque package bytes with their own relationship graph. None of that
   survives a copy that only sees one sheet's state. Cramming byte blobs into the model
   would break its character (the types are the docs) without being correct
   cross-workbook. **Byte-closures-on-the-model is declined, not merely deferred.**

2. **The autofilter is in scope and was added.** Unlike the attached parts, an
   autofilter is pure sheet-level state, a range plus criteria, workbook-independent
   (its `_FilterDatabase` defined name is derived by the writer at emit time, not
   stored), in the exact category the model already carries (`dataValidations`,
   `conditionalFormattings`). Its omission was an unprincipled gap, so `autoFilter` now
   rides the getter and setter like the other overlays.

3. **The docs are made honest.** The `WorksheetModel` and `model`-getter comments now
   scope the "drops nothing" claim to value plus overlay content and name the boundary
   explicitly: attached parts carrying workbook-level identity stay with their source
   sheet, and a model assignment _neither copies nor clears_ them. The silent trap
   becomes a documented boundary.

4. **A faithful whole-sheet/package copy primitive is deferred pending a consumer.**
   Transferring attached parts (a chart, a pivot, an image) between sheets would be a
   _new_ public API, not a change to `model`. Nothing in the codebase needs it yet, and
   building it now would be speculative abstraction (CLAUDE.md §4). It waits for a real
   use-case to force its shape.

## Consequences

- **Positive:** the model stays a clean, serialisable value with no opaque bytes; the
  autofilter is no longer silently dropped; the transfer boundary is documented rather
  than discovered by data loss.
- **Negative / deferred:** `dst.model = src.model` still does not transfer images,
  pivots, or byte-preserved charts/drawings between sheets. That is now stated in the
  contract, but a consumer who expected a deep clone must copy those parts by other
  means, and no such API yet exists.
- **Revisit when:** a concrete consumer needs to copy a sheet's attached parts (chart,
  pivot, image) to another sheet or workbook. At that point design the dedicated copy
  primitive around that use-case rather than retrofitting byte closures onto
  `WorksheetModel`.

---

## Amendment (2026-08-25): the frozen-pane view joins the model; threaded comments do not

Analysing the deferred copy primitive (§4) turned up a second field in exactly the position
`autoFilter` had been in. `Worksheet.view`, the frozen/split pane, was absent from
`WorksheetModel` for no stated reason, so `dst.model = src.model` reproduced every other
sheet-level field and silently unfroze the header row. That is the same silent-loss shape this
ADR exists to prevent, and §2's reasoning already decides it: a pane is a pair of split counts
and an anchor cell, workbook-independent, in the category the model already carries. It now
rides the registry, and `worksheet-model-preserves-frozen-pane` locks it through to the written
`<pane>`.

The recurrence is the lesson. §2 fixed one gap by name and left the _test_ implicit, so the next
field in the same position went unnoticed until someone looked. The test is now written into the
`WorksheetModel` doc comment as the thing a new field is measured against: **a field belongs in
the model when its value means the same thing on any sheet of any workbook.**

Applying that test also settles `commentThreads`, which had been in neither list. A threaded
comment names its author by an id into the workbook's `persons` registry (`Workbook.persons`,
emitted as `xl/persons/person.xml` from the _workbook_, not the sheet), so a copied conversation
would name an author the destination workbook has never heard of. It fails the test and is now
named explicitly on the out-of-scope side rather than being absent from both.

The boundary this ADR draws is unchanged; it is now stated as a rule instead of a list.
