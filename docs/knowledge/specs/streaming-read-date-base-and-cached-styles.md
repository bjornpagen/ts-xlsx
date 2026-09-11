# Streaming read of date cells: cached/compact styles

Cluster: streaming

## Scenario

A large workbook is read row-by-row with the streaming reader, and its date cells resolve less
context than a buffered read would.

**Cached / compact style mode.** When style resolution runs in a cache or compact mode, with styles
interned and looked up by index rather than fully materialized per cell, the streaming reader's
date-typing path can fail to resolve a cell's number format at all, degrading date cells to raw
serials or crashing, where the buffered read succeeds.

> Spec note, not a corpus case here: the _core_ streaming-date defect, a date-formatted numeric cell
> streaming as a raw serial because styles are not applied when typing, is already locked as a
> known-open corpus case (`streaming-read-applies-date-format`). What this note carries is the root
> cause that case does not exercise: the cached or compact style mode. It needs a workbook exercised
> under both buffered and cached-style streaming reads, plus a streaming-read harness that surfaces
> per-cell resolved values under a chosen style mode, which is new adapter surface.

The 1904 date base was the other half of this note and is now built: `workbookPr/@date1904` is read
once into `Workbook.dateEpoch`, both readers hand it to the shared cell decoder, and the writer
declares it again. The corpus case `workbook-date-system-1904` locks it against Excel-authored
fixtures in both serialisations. The ordering hazard the open questions raised is answered rather
than avoided: the workbook part is read before any sheet, in both readers, because the date system is
an input to every cell decode.

## Desired behavior

- **Streaming date typing does not depend on full style materialization.** Resolving a cell's number
  format, to decide date against number, works whether styles are fully materialized or held in a
  cache or compact form. A date cell reads as a date in every style mode, and no mode crashes or
  silently downgrades a date to a raw serial.
- **Parity is the invariant.** For the same workbook, streaming and buffered reads agree on both the
  _type_ (date) and the _value_ of every date-formatted cell.

## Open questions

- Does the compact or cache style mode intern the number-format _code_, which is recoverable, or only a
  resolved style _object_, which may omit the numFmt needed for date detection? The fix differs
  accordingly.

Related: `streaming-read-applies-date-format` (the default-style face of the same defect),
`streaming-reader-styles-option-and-defaults`, `streaming-read-styles-before-cells`,
`date-value-timezone-conversion`, `xlsx-date-detection-control`.
