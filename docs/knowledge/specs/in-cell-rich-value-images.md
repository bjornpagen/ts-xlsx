# Reading in-cell ("Place in Cell") rich-value images

Cluster: images

## Scenario

Excel's "Place in Cell" feature stores an image _inside a cell value_ rather than as a
floating drawing anchored over the grid. This uses the rich-value mechanism: the cell carries a
value-metadata (`vm`) index that points through the workbook metadata part to a rich-value record,
which references the image via a rich-value relationship to a media part. The legacy `#VALUE!`
error is stored in the cell as a fallback for applications that don't understand rich values.

A reader that only understands drawing-anchored images (twoCell and oneCell anchors) surfaces none of
these: `getImages()` returns nothing, and the image-bearing cells look empty or like errors. The
image is silently invisible to the caller.

## Desired behavior

- Reading a workbook whose cells use in-cell rich-value images surfaces those images, each
  **associated with its host cell** such as C2, not with a drawing-anchor position.
- Resolution follows the real chain: cell `vm` index, then workbook metadata, then the rich-value
  record, then the rich-value relationship, then the media part, yielding the correct image bytes for
  each cell.
- The legacy `#VALUE!` error fallback stored in the cell must not mask the presence of the in-cell
  image. The model should expose the image and treat the error as the documented fallback.
- This is distinct from drawing-anchored images: the API should make clear which images are
  in-cell, part of a cell's value, and which are floating, anchored over the grid.

## Open questions

- The public model shape for an in-cell image: a distinct cell value kind, such as an image value,
  or a parallel per-cell image accessor. It must round-trip.
- Write-side authoring of in-cell images, creating a "Place in Cell" image: in scope, or read and
  preserve only for now?
- Hostile-input hardening: the metadata to rich-value to media indirection is attacker-influenced, so
  bound the indirection depth and validate indices, and a crafted file must not force unbounded work
  or dangling-reference crashes.
- Interaction with `getImages()` (drawing-anchored) so callers can enumerate both kinds without
  conflating them.

Related: `image-by-filename-is-node-only`, `image-anchor-emu-from-real-column-geometry` (the
drawing-anchored image model this extends beyond).
