# Image anchor offsets must derive from real column/row geometry

## The scenario

Images are positioned in a sheet by an anchor whose coordinates can be fractional:
`tl.col = 1.5` means "the middle of column B", `tl.row = 2.25` means "a quarter of the
way down row 3". The public API exposes this fractional coordinate, and OOXML stores it
as a whole cell index plus a sub-cell offset in EMU (`<xdr:col>` plus `<xdr:colOff>`). The
translation between the two must use the _actual_ geometry of the referenced column and
row: a wide column means "halfway across" is a large EMU offset, and a narrow one a small
offset. Users also resize columns and rows after anchoring and expect the picture to track.

## The defect this must fix

The legacy anchor converts a column width to EMU with an ad-hoc factor
(`Math.floor(width * 10000)`) and falls back to a fixed constant (`640000` EMU) for any
column whose width was not explicitly set. Neither matches Excel's real width-to-EMU
formula (character units to pixels to EMU). The observable consequences, all confirmed by
reproduction:

- A **wider** custom column can yield a **smaller** half-way offset than a default
  column. Concretely: halfway across a width-38 column serializes to `colOff = 190000`,
  while halfway across a _default_-width column serializes to `colOff = 320000`, so the
  wide column gets the smaller offset, which is backwards.
- Because default columns use a constant rather than the real default width, an image
  over unset columns is positioned against a fictional width.
- The picture lands left of where the caller asked, and the error grows with the
  mismatch between the ad-hoc factor and the true geometry.

Whole-integer anchors are unaffected, since the offset is zero and the image sits on the
cell boundary, and must stay that way.

## Desired behaviour

- Fractional coordinate `c` maps to `nativeCol = floor(c)` and
  `colOff = round((c - floor(c)) * widthEMU(nativeCol))`, where `widthEMU` is the
  referenced column's real width in EMU, using Excel's character-width to pixel to EMU
  formula, and symmetrically for rows using real row height in points to EMU.
- Default (unset) columns and rows use Excel's real default width and height, not a constant.
- Offsets are consistent under monotonic width changes: a wider column always yields a
  larger offset for the same fraction.
- The anchor resolves geometry from the worksheet so that a resize before serialization
  is reflected, or the behaviour is explicitly documented if frozen at add time.

## Related lossy round-trip: image scale

A separate but adjacent defect: a picture the author scaled, say to 40% of native
size, changes scale after a read-then-write round-trip, observed drifting to ~78%. The
extent is stored as pixels via a fixed `EMU_PER_PIXEL_AT_96_DPI = 9525` factor, so a
picture authored at non-96-DPI or explicitly scaled does not survive the pixel
round-trip, and `Math.floor` truncates each direction. The reader should preserve the
raw EMU extent (`<xdr:ext cx cy>`) and only convert to pixels at the public API
boundary, so an unedited image re-serializes with a byte-identical extent.

## Open questions for the rebuild

- The exact EMU-per-character-width formula and Excel's default column width in EMU.
- Whether offsets recompute at write time, tracking later resizes, or freeze at add
  time. The former matches user expectation but couples the anchor to the live sheet.
- Whether to add the convenience anchor corners users ask for (`tr` and `bl`) alongside
  `tl` and `br`.
- How extent preservation (scale) interacts with two-cell anchors that carry no `ext`,
  where the extent is implied by the cell span, since the two representations must not fight.
- A file we repair on write must still be one Excel opens without a repair prompt (see
  [[excel-repair-on-open-structural-constraints]]).
