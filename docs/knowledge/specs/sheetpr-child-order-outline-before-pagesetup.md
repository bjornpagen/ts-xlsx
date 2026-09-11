# The sheetPr child order: outlinePr must precede pageSetUpPr

Cluster: xlsx-io

## Scenario

A worksheet is configured with both page-setup fit-to-page scaling and outline summary properties,
say summary rows placed above their detail rather than below. The worksheet's sheet-properties
element (`<sheetPr>`) must emit its outline-properties child (`<outlinePr>`) before its
page-setup-properties child (`<pageSetUpPr>`), the order the CT_SheetPr schema sequence requires. If
they are emitted in the wrong order the file is rejected as corrupt by Excel, which offers only to
recover the workbook and then drops the affected sheet. Setting either property alone produces a
valid file, and the corruption arises only from their combination, a sign of a child-ordering bug.

> Spec note, not a corpus case. Update: the writer **does** now emit `<outlinePr summaryBelow="0"
summaryRight="0"/>` from the outline-summary settings, and those flags round-trip, locked
> by the `worksheet-outline-summary-position-round-trips` case. What remains for _this_ note is the
> **child-order** guard: when `<outlinePr>` and `<pageSetUpPr>` are emitted together they must follow
> the CT_SheetPr sequence (`tabColor`, `outlinePr`, `pageSetUpPr`), and a wrong order corrupts the
> file. That combined-emission ordering is the residual value here; the earlier assumption that
> `<outlinePr>` was never emitted is now stale.

## Desired behavior

- Writing a worksheet with **both** fit-to-page scaling and an outline summary property produces a
  valid, non-corrupt worksheet part.
- Within `<sheetPr>`, the children follow the CT_SheetPr sequence, in particular **`<outlinePr>`
  before `<pageSetUpPr>`**, where the full order is `tabColor`, `outlinePr`, `pageSetUpPr`.
- The fit-to-page setting round-trips, with `<pageSetUpPr>` recording `fitToPage`, alongside the
  outline summary setting.
- **The summary-placement settings are precisely typed on the public worksheet-properties surface.**
  The worksheet properties type must fully type an `outlineProperties` field with its two booleans,
  one for whether outline summary rows appear _below_ their detail rows rather than above, one for
  whether summary columns appear to the _right_ of their detail columns rather than left, which map
  to the `<outlinePr>` attributes `summaryBelow` and `summaryRight`. Reports of the runtime honoring
  these values while the public TypeScript declaration omitted `outlineProperties`, forcing an
  untyped index-access workaround, are the type-surface face of the same outline feature: the types
  _are_ the docs, so a supported outline property that is settable at runtime but absent from the
  type is a defect. Absence of the setting corresponds to the format default, summary below and to
  the right.

## Open questions

- Resolved: `summaryBelow` and `summaryRight` **do** now surface as an `<outlinePr>` child and
  round-trip (see `worksheet-outline-summary-position-round-trips`). The remaining gap is authoring a
  worksheet that emits `<outlinePr>` and `<pageSetUpPr>` **together** so the child-order guard can be
  exercised.
- The child-order assertion uses the worksheet XML element-order technique the
  `comment-and-table-coexist-on-same-sheet` case uses for the drawing, legacyDrawing and tableParts
  sequence.

Related: `comment-and-table-coexist-on-same-sheet`, `row-outline-collapsed-flag-belongs-on-summary-row`,
`column-width-and-pagesetup-roundtrip-fidelity`.
