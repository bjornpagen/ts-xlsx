# Extended (x14) conditional-formatting rules must round-trip with their formula and extended dxf

Cluster: conditional-formatting

## Scenario

Modern Excel stores certain conditional-formatting rules in a worksheet **extension block** rather
than the classic conditional-formatting collection. These extended rules can be of type
`expression`, where the trigger condition is a formula carried in a nested formula element, and their
applied formatting is described by a **differential-formatting (dxf) child that lives inside the
extension namespace**, an _extended_ dxf, distinct from the classic dxf used by ordinary
conditional formatting. When a workbook containing such an extended expression rule is read and
written back, two things must survive: the rule's condition formula, and the extended dxf that
describes the visual formatting. If the extended dxf is dropped on write, the rule is preserved
structurally but becomes a **visual no-op**, so Excel opens the file with the rule present but applying
no formatting at all.

> Spec note rather than a corpus case: faithfully exercising this needs a fixture authored by Excel
> that carries the extension-namespace conditional-formatting block, and the harvested attachment for
> this report is not available; the current writer has no authoring path for extended rules, so a
> spec-built fixture cannot reproduce it either. The durable requirement is recorded here and should
> be promoted to a corpus case under `roundtripFixturePackageParts`, or an extended-CF inspection
> facet, once a real extension-namespace fixture surfaces. The unmodeled-part-passthrough principle
> already captured for drawings, VML and pivot applies directly.

## Desired behavior

- An extended conditional-formatting rule of type `expression` **preserves its condition formula**
  through a read/write round-trip.
- The rule **preserves its extended differential-formatting (dxf) child**, so the applied formatting
  is never silently dropped into a no-op.
- On write, the differential-formatting node is emitted **in the extension namespace** as an extended
  dxf, not omitted, and not confused with or downgraded to the classic conditional-formatting dxf.
- Classic and extended rule families should merge into a **single conceptual model on read**, so the
  formula on an extended expression rule is handled the same way as the formula on a classic rule and
  callers see one coherent conditional-formatting surface regardless of where a rule was stored.

## Open questions

- Should extended rules the library does not fully model be **passed through** verbatim on write,
  preserving the raw extension block, until a first-class model exists, so nothing is dropped in the
  interim?
- Which extended rule types beyond `expression` (data bars with extended options, icon sets, color
  scales with extension attributes) share this extension-block storage and need the same treatment?
  Concretely observed for **icon sets with per-icon overrides, "mixed" or "custom" icons**: an
  `iconSet` rule whose individual threshold entries each draw from a _different_ icon set or icon id
  than the rule's base set (one arrow, one symbol, one "no icon") is an x14-extension form. Both
  the buffered writer and the streaming writer emit only the classic `<iconSet>` element and drop the
  per-icon overrides entirely, so the custom icons never reach the file and every threshold falls back
  to the base set's uniform icon. This is one model gap surfacing through two writers, not a
  streaming-specific regression: the extension-namespace authoring path is simply absent. The
  first-class, or pass-through, extension model must carry each threshold's `{iconSet, iconId}`
  override, including the "no icon" selection, and emit it identically on both write paths.
- How are the extension-block dxf records reconciled with the classic `dxfs` table on write: as
  separate tables, or a merged one with namespace-aware emission?
- The **copy path** loses the extension too: reading a sheet's conditional formattings and replaying
  a rule onto another sheet (`newSheet.addConditionalFormatting(rule)`) drops the x14 extension,
  because the read-side model exposes only the classic rule fields (`type`, `operator`, `formulae`,
  …) and not the extended-namespace payload. Preserving x14 on a read, re-add and write copy needs the
  same first-class, or pass-through, extension model. The copy is not a separate bug, it is the same
  model gap surfacing through the authoring API rather than a raw round-trip.

Related: `conditional-format-numfmt-roundtrip`, `conditional-formatting-duplicate-values-roundtrip`,
`pivot-table-round-trip-preservation`, `foreign-file-read-modify-write-preserves-validity`.
