# Native DrawingML shapes must be authorable, enumerable, and round-trip-preserved

Cluster: images / drawing

## Scenario

A user wants to place vector shapes (rectangles, ellipses, arrows, braces, lines) onto a worksheet,
not raster images but Office preset-geometry drawings, with a fill color, an outline or stroke, and a
rotation, positioned by cell anchor. They also want existing shapes in a foreign-generated workbook
to survive a read/write round-trip rather than being dropped. Today the library models images
(raster media with cell anchors) but has no concept of a native drawing shape, so shapes authored
elsewhere are lost and there is no API to add one.

> The _preserve-on-round-trip_ half is already locked as a corpus case
> (`vector-shape-drawing-survives-roundtrip`, currently known-open): a no-op load-then-save must not
> drop an `xdr:sp` shape. This note captures the broader, still-unbuilt **authoring** API and the
> read-back model that make shapes first-class rather than merely passed-through.

## Desired behavior

- A worksheet-level API to **add** a shape and to **enumerate** existing shapes, parallel to the
  image API. Shape input: a preset geometry kind, an optional rotation in degrees, a fill (color plus
  opacity), and a stroke or outline (color, opacity, weight). Placement reuses the existing
  cell-anchor model: a one-cell anchor with an explicit extent, or a two-cell from/to anchor.
- On write, emit DrawingML: a shape (`sp`) with non-visual shape properties, `spPr` containing a
  preset geometry (`prstGeom`), a 2D transform (`xfrm` with rotation), a fill
  (`solidFill`/`noFill` via `srgbClr`/`schemeClr`), and a line (`ln` with fill or no-fill and weight),
  plus a style reference block. The shape is registered in the sheet's drawing part and drawing
  relationships exactly like an image anchor.
- On read, parse the same structure back into shape objects so incoming shapes are enumerable and
  preserved on round-trip.

## Prior art

A shape vocabulary observed working in practice: `line`, `rect`, `roundRect`, `ellipse`, `triangle`,
`rightArrow`, `downArrow`, `leftBrace`, `rightBrace`, a subset of the OOXML preset-geometry
(`ST_ShapeType`) enumeration. Fill and stroke each accept a color (RGB hex or a theme scheme color)
plus opacity; stroke also accepts a weight. Rotation is authored in degrees and stored in the
DrawingML 60000ths-of-a-degree unit.

## Open questions

- Preset-geometry scope: expose the full `ST_ShapeType` enum or a curated subset? Full-fidelity
  round-trip argues for accepting and preserving any `prst` value even if the convenience API surfaces
  only common ones.
- Shape text bodies (`txBody`): support text inside shapes in a first cut, or preserve-only on
  round-trip?
- Color model: shape fill and stroke should reuse the workbook's existing color abstraction (RGB plus
  theme/scheme) rather than a shape-local color type; opacity maps onto the alpha channel.
- Anchoring: confirm both one-cell (`tl` plus `ext`) and two-cell (`from`/`to`) anchors, matching image
  placement, and that `editAs` and anchor-movement semantics are handled.
- Foreign-file tolerance: reading a workbook whose drawings mix images and shapes must not corrupt or
  drop the image anchors already supported.

Related: `vector-shape-drawing-survives-roundtrip` (the locked preserve-on-round-trip behavior),
`form-controls-roundtrip-preserved`, `image-anchor-emu-from-real-column-geometry`,
`in-cell-rich-value-images`.
