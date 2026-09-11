# Distinguish a cell-embedded ("in-cell") image from a floating anchored drawing

Cluster: images

## Scenario

A user adds an image and anchors it to a cell range, say C2:E6. On export the image floats _over_ the
cells as an independent drawing object: it is not owned by any single cell, is not returned by
cell-value reads, and is not sorted, filtered or moved with the row it visually sits on. The user
expected the picture to be embedded _in_ the cell, to behave as the cell's value, live inside the
cell's bounds, and travel with the cell through sort, filter, and row insert and delete. These are two
fundamentally different features, and today only the floating one exists.

> Spec note, not a corpus case: true in-cell images are a new capability, a rich-value and
> cell-metadata feature, not a bug in current behavior, so there is nothing to baseline yet. The
> floating-anchor behavior is already covered by the image-anchor cases. The durable value is naming
> the distinction and the design questions; a corpus case follows once, and if, rich-value in-cell
> images are implemented, at which point the adapter grows a capability to read an image-valued cell.

## Desired behavior

Distinguish the two ways an image can live in a sheet, and let callers choose:

- **Floating / overlay image (today's behavior).** A drawing object anchored to the sheet grid via a
  one-cell or two-cell anchor. It sits over the cells, is not a cell value, and, depending on its
  "move and size with cells" property (`editAs`), may or may not reflow when rows or columns are
  inserted, deleted, or resized. `addImage(id, 'C2:E6')` produces this, and the range only positions
  the overlay's anchors. This mode deserves explicit, documented coverage of the `editAs` and
  "move and size with cells" property, which is the usual source of "why does my picture float, or not
  move with the cell" confusion.
- **In-cell / embedded image.** The image is the _value_ of exactly one cell, which is Excel's "Place
  in Cell" and IMAGE-function rich-value cells. It is owned by and clipped to that cell, moves and
  resizes with it, participates in sort, filter, insert and delete like any value, and reads back as
  the cell's value. On the OOXML side this is not an `xdr:drawing` two-cell anchor at all. It is a rich
  value in the rich-data and cell-metadata parts (`xl/richData/*`, cell metadata `vm` attributes, and a
  rels-linked image), a distinct feature from classic drawings.
- **Native / intrinsic-size placement (a missing floating mode).** Even within the floating family,
  callers routinely want "place this picture _here_ at its own dimensions", pinned to a single
  top-left cell and rendered at the image's intrinsic pixel size, without stretching to span a range.
  Today the range-anchor helpers force a two-corner (from plus to) anchor, so supplying only a top-left
  position is not enough and the picture is distorted to fill the spanned cells. This maps to an OOXML
  **`oneCellAnchor`** whose `ext` (width and height in EMU) is derived from the image's decoded
  intrinsic size, or an `absoluteAnchor` for a fully grid-independent fixed size. Intrinsic dimensions
  come from the image bytes (PNG IHDR, JPEG SOFn, GIF header) converted to EMU via the image DPI, where
  the default 96 DPI gives 9525 EMU per pixel. The add-image call should accept just a top-left anchor
  plus an optional explicit size, defaulting the size to the decoded intrinsic dimensions when only the
  anchor is given, and preserving aspect ratio when only one dimension is supplied. Reading a file that
  already uses `oneCellAnchor` or `absoluteAnchor` must preserve that anchor mode and size on write,
  not silently rewrite it to a stretched two-cell anchor.

- **Reading floating images is a first-class, enumerable affordance, not just a write feature.**
  Opening a file that contains a picture floating over the grid, a caller must be able to _enumerate_
  the worksheet's drawing images and, for each, recover the media payload (bytes plus content
  type or extension) and the anchor: which cells it is pinned to, the anchor _kind_
  (one-cell, two-cell or absolute), and the offset. Read-back must preserve the anchor kind and size
  so a subsequent write round-trips it faithfully rather than silently rewriting a one-cell or
  absolute anchor into a stretched two-cell one. This read enumeration across anchor variants,
  returning media plus top-left cell coordinates, is already locked by
  `worksheet-images-enumerated-across-anchor-variants`. The durable point here is that "how do I read
  the floating images out of a file?" is answered by that enumeration, and it must work for
  foreign-authored files, not only ones this library wrote.

- **Images should be queryable by the cell or row they are anchored to, not only enumerable as a flat
  list.** A recurring real-world task, building a JSON export from a product-catalog sheet where each
  row carries a thumbnail, is "given this row or cell, which picture belongs to it?". Today a caller
  can enumerate every image with its anchor extent but must hand-roll the anchor-to-cell matching. The
  reading API should offer an ergonomic lookup: given a worksheet and a row or cell, return the
  images anchored there, each with its media payload (bytes plus content type or extension) and anchor
  metadata (from and to cells, anchor kind, offsets). The load-bearing decision is what "belongs to" a
  cell means for a two-cell (movable, resizable) anchor spanning a range: match only the top-left
  anchor cell, or any overlapped cell? Absolute (grid-independent) anchors have no owning cell, so a
  per-cell query must define how they surface, either excluded or returned only by a whole-sheet
  enumeration. Multiple images may overlap the same cell, so the lookup returns a collection. The
  anchor extents themselves are already an inspectable fact (see
  `worksheet-images-enumerated-across-anchor-variants`); the gap is the keyed query API, not new
  parsing, so a helper such as `worksheet.imagesAnchoredAt(row, col)` or a caller-built index over the
  enumeration.

Prior art: classic implementations support only floating drawing anchors, one-cell "over" a cell or
two-cell "over" a range, with the `editAs` flag. Helpers that "add an image over a range" still produce
a floating two-cell overlay, not a true cell-embedded rich value. Genuine in-cell images require the
rich-value and cell-metadata machinery.

## Open questions

- Scope: commit to reading and writing true rich-value in-cell images (rich-data parts plus cell
  metadata), or first just clarify and name the existing floating anchor modes and the "move and size
  with cells" behavior?
- API shape: how does a caller request an in-cell image rather than a floating one? A distinct method,
  or an anchor option like `anchor: 'inCell'` targeting a single cell, rather than overloading the
  range-anchor call?
- Read-back: an in-cell image should surface as the owning cell's value, so define the value shape
  (image ref plus alt text) and how it coexists with a formula or number-format on the same cell.
- Fallback for older Excel lacking rich-value support, where Excel writes a `#VALUE!`-style cached
  value plus the rich value: do we emit that compatibility shim?

Related: `image-range-anchor-edit-as-mode-honored`, `anchored-image-sppr-transform-detaches-in-libreoffice`,
`fractional-image-anchor-positioning`, `header-footer-image-authoring`, `public-type-surface-matches-runtime`.
