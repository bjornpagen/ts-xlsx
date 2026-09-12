# Charts survive a load-modify-save of a template

Cluster: round-trip-fidelity

## Scenario

A user opens a spreadsheet authored in a desktop application that already contains one or
more charts, uses the library to change unrelated cell values ("fill the template"), and
writes the result out. On open, every chart is gone. Charts are backed by package parts the
library does not parse: the chart definition XML, chart colors and style XML, their
relationship files, and the drawing that anchors the chart to the sheet. Because those parts
are neither parsed into the object model nor preserved verbatim, they are silently omitted from
the output, so the charts vanish.

## Desired behavior

Two layers, separable in scope.

**1. Passthrough preservation (near-term, high value).** When reading a package, parts and
relationships the library does not model (chart XML, chart colors and style XML, chartEx, drawings
that anchor charts, their `.rels`, and the per-sheet `printerSettings` parts that chart-bearing
report templates commonly carry) should be retained verbatim and re-emitted on write,
with their relationships and content-type overrides intact. A load-modify-save that only touches
cell values must produce an output whose charts still render. This needs no chart data model,
only faithful passthrough of opaque parts plus correct rewiring of the relationship graph
(sheet → drawing → chart → colors/style) and the `[Content_Types].xml` overrides.

**2. First-class chart model (long-term).** Parse chart parts into a typed, editable object
model so charts can be created and mutated programmatically. Much larger effort; out of scope
for the passthrough fix and can land independently.

## Root cause

Charts are described by `chart{N}.xml`, `colors{N}.xml`, `style{N}.xml` and
`chart{N}.xml.rels`, referenced from a drawing part, itself referenced from the worksheet's
rels. The legacy library parses none of these, and its writer reconstructs the package purely
from its object model rather than preserving unrecognized inbound parts, so anything unmodeled
is dropped on save.

## Open questions

- Scope of "unknown part" passthrough: whitelist chart and drawing part types specifically, or
  generic verbatim retention of any inbound part not owned by a modeled feature? Generic
  retention is more robust but risks re-emitting stale parts when a modeled feature such as an
  image drawing shares the same drawing part that also hosts a chart. A drawing part can mix
  charts and images, so partial ownership must be handled without corrupting the shared XML.
- Relationship-id and part-name collision handling when the library also adds its own new parts
  (images, tables) to a workbook that already carries preserved chart parts. For a picture on the
  sheet a chart's drawing sits on, the answer is recorded by `picture-joins-a-kept-drawing`: the
  picture joins the kept drawing, its relationship and shape ids numbered past the drawing's own.
- Interaction with the streaming write path, which builds the package incrementally and may have
  no place to stash preserved parts.
- Whether preserved parts should be exposed read-only in the public API, as an inventory of
  retained-but-unmodeled parts, so callers can detect that a workbook contains features the
  library won't touch.

Related notes: `form-controls-roundtrip-preserved`, `pivot-table-round-trip-preservation`,
`excel-repair-on-open-structural-constraints`, all instances of the same principle: parts the
model does not own must survive a round-trip rather than being dropped.
