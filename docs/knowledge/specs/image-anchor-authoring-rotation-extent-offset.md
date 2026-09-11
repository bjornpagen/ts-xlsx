# Authoring an image anchor: rotation, explicit extent, and sub-cell offset are first-class inputs

Cluster: images

## Scenario

A caller anchors an image onto a worksheet and wants to place it exactly the way an interactive
editor would: rotated to an angle, sized to an explicit extent rather than the image's intrinsic
pixel size, and nudged by a fractional offset within its anchor cell. The natural expectation is that
the add-image and anchor API accepts these as authoring inputs, so that programmatically rebuilding a
sheet, for instance duplicating a sheet by re-adding its images through the public API rather than
cloning internal model objects, reproduces the original placement faithfully.

Today the anchor authoring API is narrower than the drawing model it serializes into. Rotation is
preserved on a raw read-then-write round-trip (see `image-rotation-preserved-on-roundtrip`) but cannot
be _set_ through the add-image path, so a caller who assembles an image with the public API gets an
un-rotated, intrinsically-sized, cell-aligned picture even when the drawing format can express far
more. Extent (an explicit width and height for the picture, independent of the source bitmap's pixel
dimensions and DPI) and a sub-cell offset (the fine EMU offset from the anchor cell's top-left) are
likewise expressible in the OOXML drawing anchor but absent from the authoring inputs.

> Spec note, not a corpus case: this is a write-side API gap, not a malformed-output bug from
> a data file. Round-trip _preservation_ of rotation, extent and offset is already a corpus concern;
> the durable value here is the authoring contract, which placement properties a caller can specify
> when adding an image, which is a Phase 3 API-shape decision rather than a current-behavior assertion.
> It becomes a corpus case once the rewrite's add-image API accepts these inputs and a case can
> author a rotated, extent-sized, offset image and assert the serialized `<a:xfrm>` transform.

## Desired behavior

- **Rotation is an authoring input.** The add-image and anchor API accepts a rotation angle and emits
  it as the picture shape's `<a:xfrm rot="…">` transform (OOXML 1/60000-degree units). Setting rotation
  on write mirrors the rotation that is already preserved on read, so author-then-write and
  read-then-write agree.
- **Extent is specifiable independently of intrinsic size.** A caller can give an explicit picture
  extent (width and height) rather than inheriting the source bitmap's pixel size, and the extent
  serializes in EMU, DPI-independent, consistent with
  `image-pixel-extent-converts-to-emu-independent-of-dpi`.
- **Sub-cell offset is specifiable.** The anchor accepts a fractional or EMU offset from its top-left
  anchor cell, so an image can sit partway into a cell rather than snapping to the cell corner,
  consistent with `image-anchor-fractional-offset-respects-cell-size`.
- **Re-adding is lossless.** Rebuilding a sheet by re-adding images through the public API, not by
  cloning internal objects, reproduces rotation, extent, and offset, so "duplicate this sheet"
  written in terms of the public API does not silently flatten placement.

## Open questions

- The precise units and shape of the authoring inputs: rotation in degrees, converted internally, or
  raw 1/60000-degree units; extent in pixels, EMU, or a cell-relative fraction; offset in pixels or
  EMU. The public API should favor human-legible units and convert at the boundary.
- Whether these properties belong on the `addImage` or media definition, on the anchor passed to the
  worksheet's add-image call, or split between them, with media as bytes and anchor as placement and
  rotation.
- How rotation composes with a two-cell (from/to) range anchor, where the bounding box rather than a
  single extent defines size; rotation of a range-anchored image rotates within that box.

Related: `image-rotation-preserved-on-roundtrip`, `image-anchor-fractional-offset-respects-cell-size`,
`image-pixel-extent-converts-to-emu-independent-of-dpi`, `image-anchor-emu-from-real-column-geometry`,
`add-image-source-contract`, `image-embedded-in-cell-vs-floating-anchor`.
