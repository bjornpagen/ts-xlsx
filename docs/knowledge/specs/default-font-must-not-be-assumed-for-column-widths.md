# The default font must be honored, not assumed, when computing column widths

Cluster: styles

## Scenario

A workbook's default (normal-style) font is not Calibri 11, say Calibri 72, or a different
face entirely. Column widths in the spreadsheet format are expressed in _character units_ of the
**Maximum Digit Width** of the workbook's default font, so the pixel width of a column depends on
what that default font actually is. A library that hardcodes an assumed Calibri-11 metric when
reading or writing column widths miscomputes them for any workbook whose real default font
differs: columns come out too narrow or too wide, and a load-modify-save of such a template
visibly breaks its layout.

## Desired behavior

- The width model must be anchored to the workbook's **declared** default font, the normal
  cell-style font in `styles.xml`, not to a baked-in Calibri-11 assumption.
- The library must not unconditionally inject its own assumed default font ahead of, or in place
  of, the default font a file already declares. Reading a workbook whose default font is
  non-standard and writing it back unmodified must preserve that default font.
- Column `width` definitions and the `sheetFormatPr` baseline (`defaultRowHeight`,
  `defaultColWidth`, `baseColWidth`) must survive a read-write round-trip unchanged.
- Per-cell and per-column style-index references must stay consistent across the round-trip, so
  styled columns still point at their intended styles.

## Prior art / format facts

- OOXML column width is `width = (charCount * MDW + 5) / MDW` rounded, where MDW is the Maximum
  Digit Width in pixels of the workbook default font at the workbook default size. Change the
  default font and every character-unit width changes meaning.
- `sheetFormatPr@defaultColWidth`, `baseColWidth` and the default row height are likewise tied to
  the default font metrics.

## Open questions

- ~~How do we represent "the workbook default font" as a first-class, typed part of the model so
  it is neither lost nor silently overwritten on write?~~ **Answered** (ADR-0025): `Workbook`
  carries `declaredDefaultFont` (what a file stated), `setDefaultFont` (authoring), and
  `defaultFont` (the resolved, complete result the writer emits as font 0). With nothing authored a
  file's own font 0 passes through verbatim, so a declared face is never overwritten. Locked by the
  `workbook-default-font-is-declared-not-assumed` corpus case.
- **Still open, and deliberately so:** do we measure MDW from actual font metrics, which requires
  font metric data for common faces, or carry width as an opaque round-trippable value and only
  interpret it when a caller asks for pixels? Round-tripping the raw value avoids the metric problem
  for the common open-modify-save path; true pixel-accurate layout needs the metrics.

  Widths stay opaque for now. Fixing the default font changed what the units _mean_ without changing
  the units, and read-write stays faithful as long as we neither reinterpret nor recompute them,
  so nothing about that fix forces the metric question. It is only forced by a caller who wants
  pixels, or by `autoFit`.

- **Still open, and downstream of the above:** what is the fallback when a declared default font is
  unknown to our metric table? Assume a documented default while preserving the declared font name
  for re-emission? Moot until there is a metric table.
