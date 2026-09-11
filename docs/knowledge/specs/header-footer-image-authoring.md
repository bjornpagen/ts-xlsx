# Authoring images in page headers and footers

Cluster: images

## Scenario

Beyond preserving existing header and footer images on round-trip, covered by the
`header-footer-image-survives-roundtrip` corpus case, users want to _create_ them: place a logo
in the left, center or right section of a page header or footer. There is no authoring API for this
today.

## Desired behavior

Expose an authoring API to place a workbook-registered image into a worksheet header or
footer: register the image at the workbook level, then attach it to a sheet's header or footer
section with an id and explicit size (`addHeaderFooterImage(imageId, { position: 'R', width:
'15pt', height: '15pt' })` or similar), while the header/footer string carries the picture token
in the chosen section (`&L&G`, `&C&G` or `&R&G`).

On write this must emit:

- the `&G` token in the correct section of the `headerFooter` element,
- a `<legacyDrawingHF>` relationship on the worksheet,
- a VML drawing part containing a `v:shape` of type `_x0000_t75` whose `v:imagedata` references
  the image via that VML part's own `.rels`,
- the correct content-type and override entries and image media part.

## Open questions

- **Coexistence with cell-comment VML.** A comment on the same sheet also emits a VML drawing,
  and relationship-id allocation across the two VML drawings must not collide. This is exactly where
  the reference implementation repeatedly broke. See `header-footer-image-survives-roundtrip`.
- **Cross-generator tolerance.** LibreOffice does not support header and footer images, so scope the
  guarantee to Excel and WPS and degrade gracefully elsewhere.
- **Sizing units.** Accept points only, or also EMU and pixels? Pick one canonical unit and document
  conversions.
- **Naming.** How the authoring API names header against footer and left, center and right sections,
  consistently with the `&L`/`&C`/`&R` and `&G` tokens.

Related: `header-footer-image-survives-roundtrip` (the preservation case), and the general
principle in `chart-parts-survive-template-roundtrip` that unmodeled parts survive a round-trip.
