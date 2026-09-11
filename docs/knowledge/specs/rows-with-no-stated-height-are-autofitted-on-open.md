# A row with no stated height is auto-fitted on open, and the auto-fit saturates

Cluster: worksheet

## Scenario

A row that carries no `ht` or `customHeight` records no geometry at all. Something has to decide how
tall it is, and for Excel Desktop that something is Excel's own auto-fit. `estimateWrappedLines`
(`src/core/text-metrics.ts`) exists so an author can state a height instead, and its first
justification, written before it was measured, was that Excel's auto-fit is _lazy and
incomplete_: that on a sheet of wrapped multi-thousand-character cells, bands of the grid open
blank until they are clicked.

That claim is wrong, in the same way and for the same reason as the one corrected in
`grid-geometry-limits-are-excels-not-the-schemas`: it was an inference about rendering, never an
observation. This note records what was actually measured, and the narrower reasons to state a
height that survive it.

## Measured behavior: the auto-fit happens, on open, and is not lazy

Excel Desktop (Microsoft 365, build 16.0.20131, Windows, workbook default font Aptos Narrow 11,
`StandardHeight = 14.5`). A ladder fixture, one wrapped cell per row in a column of stated width
40, row _n_ holding a prose string of the length below, no `ht` anywhere in the package, opened
over COM and read back with `tools/excel-oracle/read-geometry.ps1`:

| Characters in the cell | `RowHeight` read back | Lines (height / 14.5) | `estimateWrappedLines(text, 40)` |
| ---------------------- | --------------------- | --------------------- | -------------------------------- |
| 200                    | 87.0                  | 6                     | 5                                |
| 500                    | 188.5                 | 13                    | 13                               |
| 1000                   | 377.0                 | 26                    | 25                               |
| 1500 and up (to 32000) | **409.5**             | 28 (saturated)        | 38 and up                        |

Two facts. The auto-fit is real and immediate, since a headless automation open, which paints
nothing, already has the heights, so it is not "at paint time" in any sense a caller can observe. And
it **saturates at 409.5**, the same `MAX_ROW_HEIGHT` a user is allowed to set; from roughly 1100
characters at this width upward, every row is that height and the text below the 28th line is
clipped. Note the asymmetry with the file-read path in the geometry note, which clamps to 409.6.

## Measured behavior: no blank bands

Three fixtures, every column wrapped, no row heights, opened in interactive Excel Desktop through
the `excel-gui-automation` skill's window library and navigated by keyboard, screenshotting the
window rect with no settle delay after each keystroke:

| Fixture   | Cells                                   | Package | First window | Result              |
| --------- | --------------------------------------- | ------- | ------------ | ------------------- |
| 400 x 6   | ~2000 chars, width 100                  | 0.6 MB  | 0.9 s        | every frame painted |
| 2000 x 10 | ~1500 chars, width 60                   | 3.7 MB  | 1.1 s        | every frame painted |
| 6000 x 10 | ~1500 chars, width 60 (~90M characters) | 5.9 MB  | 1.1 s        | every frame painted |

Navigation covered the cases where a deferred layout would show: a settle burst from first paint,
Go To jumps to `A800`, `J1500`, `A1999` and back to `A2`, ten `{PGDN}`s at 60 ms intervals, and
`^{END}` to the far corner. Every screenshot taken 0 ms after the keystroke, including
`R6000C10`, the far corner of the 90M-character sheet, shows its text. No band opened blank, no
frame filled in on the following capture, and selecting a cell changed nothing about what was
already drawn. The one visible consequence of a tall auto-fitted row is the one the geometry note
already records: a row taller than the viewport puts content below the fold.

The probe cannot exclude a blank frame shorter than the capture latency, a few tens of ms. It can
exclude the claim as stated, which was about bands that stay blank until clicked.

## What survives as a reason to state a height

- **The saturation is a ceiling on the auto-fit, not just on assignment.** Past about 28 lines Excel
  stops answering the question, so an author who wants a _chosen_ geometry, exactly the lines that
  matter, or a deliberate 3-line band with the rest clipped, has to write one, and needs a line
  count to write it.
- **The file records nothing.** The auto-fit is Excel's answer, computed at open and not persisted
  unless Excel saves. Any consumer that does not implement it, and this library's reader reports
  `row.height === undefined`, as does anything paginating or rendering from the package, has
  no height to work from.
- **A stated height is deterministic.** An unstated one is a function of the opening application's
  default font, which `default-font-must-not-be-assumed-for-column-widths` shows is not a constant:
  14.5 points per line here, 15 for a Calibri-11 workbook.

## The estimate is a lower bound

The table above is also the first measurement of `estimateWrappedLines` against Excel: 5 where
Excel laid out 6, 13 where Excel laid out 13, 25 where Excel laid out 26. It under-counts, which is
the direction that costs a line of text, for two compounding reasons. Excel breaks at word
boundaries where a character count breaks mid-word, and the usable width is about 0.64 character
units narrower than the stated `width` (a stated 40 reads back as `ColumnWidth = 39.36`, the same
offset the geometry note records on every width).

Making the count word-aware would close most of that gap for prose and is not obviously right: it
would trade an exact answer for a monospaced face and a hard-wrapping renderer for a closer
approximation of one specific application. Left open deliberately, and recorded here so the direction
of the error is known rather than discovered in a clipped report.

## Provenance

`source: excel-desktop-verification`, the tier ADR-0013 describes. One Excel build, one host.

The height table is reproducible from the repo:
`pwsh -NoProfile -File tools/excel-oracle/read-geometry.ps1 -Path <ladder.xlsx> -Rows 13 -NoResave`
against a fixture built as described above. The blank-band half needed the interactive tier, since
headless COM paints nothing and cannot answer a question about painting, and was taken by
driving a visible Excel with the `excel-gui-automation` skill's `xl-window-lib.ps1`
(`Set-ExcelForeground` plus `Save-WindowScreenshot`, window-rect-scoped) plus `SendKeys` navigation.
That half is a recorded observation, not reproducible from the repo, and the fixtures were scratch
files under `.tmp/`.
