# Form controls: preserve on round-trip, model where feasible

Cluster: round-trip-fidelity

Reported as "Worksheet with form controls".

## Problem

Worksheet **form controls** (drop-down and combo-box lists, checkboxes, option buttons, spin buttons, scroll bars, group boxes, buttons) are not represented in the workbook model. Reading a file that contains them and writing it back out drops them completely. Any cell-link a control drives is lost, so downstream formulas that read those linked cells break.

## How form controls are encoded (OOXML)

A single sheet with form controls typically ships these parts, all present in the sample fixture for this scenario:

- `xl/ctrlProps/ctrlProp*.xml`, one per control, holding the control's semantic properties: object type (`checkbox`, `drop`/list, `spin`, `scroll`, `radio`, `buttonface`), `fmlaLink` (linked cell), `fmlaRange` (source list range for drop-downs), `fmlaTxbx`, checked state, min/max/inc/page for spinners, `dropLines`, and so on.
- `xl/drawings/vmlDrawing*.vml`, the legacy VML shapes that anchor the controls on the grid (`<x:ClientData ObjectType="...">` with anchor, `FmlaLink`, `FmlaRange`, `Checked`, `Sel`, `DropStyle`, `Val`, `Min`, `Max`).
- `xl/drawings/drawing*.xml`, DrawingML `<xdr:...>` control shapes referencing the ctrlProps by relationship.
- Worksheet `xl/worksheets/sheet1.xml` wiring: a `<controls>` block, where each `<mc:AlternateContent>` gives a `<control>` with `r:id` pointing at a ctrlProp, plus a `<controlPr>` referencing the drawing, and a `<legacyDrawing r:id="...">` pointing at the VML.
- Relationships in `xl/worksheets/_rels/sheet1.xml.rels` and `xl/drawings/_rels/drawing1.xml.rels` tying it all together.

## Desired behavior

Tiered, so partial support is still valuable:

1. **Minimum, lossless round-trip (highest priority).** A file with form controls that is read and written back without the user touching the controls must retain every control, its cell links, list ranges, and state. Practically this means the unknown and legacy parts (`ctrlProps/*`, `vmlDrawing*.vml`, control drawing parts, the `<controls>` and `<legacyDrawing>` worksheet elements, and their rels) must be preserved verbatim rather than stripped. This alone fixes the "formulas break after save" complaint.
2. **Read model.** Expose controls per worksheet with a typed shape: control type, anchor or position, linked cell address, source list range for drop-downs, current value or checked state, and spinner bounds. Read-only first.
3. **Authoring (later).** Allow creating and editing common controls: a drop-down list bound to a range with a cell link, a checkbox bound to a cell. This is a larger design and can follow the read and round-trip work.

## Prior art / notes

- The same class of "unmodeled part gets discarded on write" failure has bitten other embedded content, such as media and images. A general principle worth adopting: parts the model does not understand should be carried through untouched on round-trip rather than dropped, so fidelity degrades gracefully.
- Data-validation dropdowns (`<dataValidation type="list">`) are a _different_ feature from form-control drop-downs and are unrelated here. Don't conflate them.

## Open questions

- Do we implement generic unknown-part pass-through as the round-trip mechanism, or model form controls explicitly? Pass-through is cheaper and fixes the reported breakage; explicit modeling is needed for the read and author tiers. Likely: pass-through first, explicit model incrementally.
- How to keep pass-through relationships valid if a worksheet is otherwise heavily rewritten, meaning r:id and part-name stability.
- Scope of control types to model explicitly in tiers 2 and 3, against leaving them as pass-through only.

## Fixture

An .xlsx with a full set of form controls (12 ctrlProps, VML legacy drawing, control drawing, worksheet `<controls>` and `<legacyDrawing>` wiring) is available for a round-trip-preservation test.
