# Carrying image-bearing content between workbooks must copy the media or fail loudly

Cluster: images

> **Resolved (2026-08-25).** `Workbook.exportImages(sheet)` and `Workbook.importImages(sheet, images)`
> are the copy affordance, and a corpus case
> (`sheet-images-carry-between-workbooks`) locks both legs of the contract below. The open
> questions are answered: the affordance is a **dedicated pair on `Workbook`**, not a `model`
> field, because a model is a serialisable value and an image is bytes on the workbook, which is the
> boundary ADR-0005 draws. So carrying a sheet whole is `dst.model = src.model` followed by
> `dstWb.importImages(dst, srcWb.exportImages(src))`. Registration is **content-addressed**, so a
> picture already held is re-used at its existing id rather than stored twice. **Scope is floating
> anchors and the sheet background.** A header or footer image is a byte-preserved part
> (`legacyDrawingHF`) and rides the preserved-reference machinery instead, and in-cell rich-value
> images do not exist in the model yet. The raw-`model`-splice failure named below is also no longer
> an opaque `undefined` dereference: `planMedia` throws an `AuthoringError` naming the offending
> image id, so an anchor that reaches the writer still holding a foreign id fails loudly, as the
> contract demands.

## Scenario

A user copies content, a worksheet, a range or a drawing, from one workbook into another, and the
copied images render as broken in the destination file. The cause is structural: an image in OOXML is
a two-part thing, a drawing or anchor in the worksheet that references an image by a workbook-scoped
media id, plus the actual media bytes and the relationship that binds them. Copying only the
drawing or anchor, or the cell model, into a different workbook carries a reference to an image id that
does not exist there, so the writer emits a package with a dangling drawing relationship pointing at
media that was never registered: a silently broken image.

## Desired behavior

- **Copying image-bearing content transparently copies the media.** When a drawing or anchor, or a
  worksheet or range containing one, is copied into a different workbook, the underlying image media
  bytes and their relationships are registered in the destination workbook and the anchor is rebound
  to the destination's media id, so the image renders rather than breaks.
- **Or fail loudly, never silently.** If the library cannot or does not copy the media, for example on
  a low-level model copy that only moves the anchor, an anchor referencing a media id absent from the
  destination workbook must produce a clear, actionable error at write time, naming the missing image,
  rather than emitting a package with a dangling drawing relationship that a consumer shows as a
  broken image.
- **Round-trip integrity of the destination package.** After a cross-workbook copy, every drawing
  relationship in the written package resolves to a real media part with a unique rel id, the same
  packageParts and rel-id invariant the table and comment coexistence cases assert, and re-reading the
  destination surfaces the image anchored where it was placed.
- **A raw `model` splice cannot carry the media, and no longer fails opaquely when it doesn't.** As
  originally reported, transplanting an image-bearing worksheet's serialized `model` into a
  _different_ workbook (`dstWorkbook.addWorksheet(...).model = JSON.parse(JSON.stringify(srcSheet.model))`,
  the workaround users reached for absent a copy API) threw `TypeError: Cannot read properties of
undefined (reading 'name')` at **write time**, because the anchor referenced a workbook-scoped media
  id that did not exist in the destination. That is now an `AuthoringError` naming the image id, and
  the model no longer carries anchors at all (ADR-0005), so the splice is a content-only copy by
  construction and `importImages` is how the pictures follow.

## Open questions

None. The affordance, the de-duplication policy and the scope are all settled in the resolution
note above. In-cell rich-value images remain unmodelled, but that is
`image-embedded-in-cell-vs-floating-anchor`'s question rather than this one's.

Related: `add-image-source-contract`, `image-embedded-in-cell-vs-floating-anchor`,
`streaming-write-add-image`, `comment-and-table-coexist-on-same-sheet`.
