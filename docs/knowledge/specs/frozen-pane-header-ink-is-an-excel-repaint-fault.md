# A frozen pane's header ink can go missing until the cell is clicked, and no package can prevent it

Cluster: worksheet

## Scenario

The shape this library writes for almost every report (a banner row, a header row beneath it,
`freeze(2)` over the pair, and a table whose `autoFilter` puts a dropdown button on every header
cell) renders with header text missing on some columns. The affected cells hold their values,
since the formula bar shows them, and only the ink is absent. Clicking a blank header paints it and
blanks a different one, so the set of drawn headers changes but its size does not.

The reflex on a report like this is to suspect the package: a malformed `<pane>`, a font colour
that collapses into its fill, a header clipped by a row height too small for its wrapped text.
All three were checked and all three are wrong. This note records the measurement so the next
agent does not re-derive it, and more usefully records the control that settles it, because
the fault is not reproducible by any means this repo can automate.

## The control: an Excel-authored package reproduces it

The decisive test is not "does our file do this" but "does a file _Excel wrote_ do this".
`tools/excel-oracle/observe.ps1` opened the reported workbook over COM and re-saved it, giving
`openThrew: false` and `repaired: false`, so Excel took the package as given rather than recovering
it, producing a copy whose every byte is Excel's own serialization of the same content.

**The fault reproduces in that package**, confirmed by the reporter on the machine that showed it
originally. Nothing this library emits is in the causal path: the same content, serialized by the
application that renders it, renders the same way.

## Measured behavior: the package is right, and synthetic navigation cannot see the fault

Excel Desktop (Microsoft 365, version 16.0, build 20228.0, Windows). The workbook: two sheets
frozen at `ySplit="2"` (one also `xSplit="1"`), 13 and 18 columns, each a table with an
`autoFilter` across the header row, header cells bold 9pt `FFFFFFFF` on a solid `FF009EE0` fill,
row height 34 with `wrapText`.

| Hypothesis                     | Check                                                                                                      | Result                                                                                                                                                                                                                                                  |
| ------------------------------ | ---------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Malformed `<pane>`             | Diff ours against Excel's re-save                                                                          | **Identical.** `<pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/>`, and the four-pane `xSplit="1" ySplit="2"` variant, come back byte-for-byte. Excel adds only empty `<selection>` elements for the panes it does not sit in. |
| Ink collapsing into the fill   | Sample the rendered pixels in the frozen header band                                                       | Fill `#009EE0`, ink `#FFFFFF`, exactly what `styles.xml` states, at every position captured.                                                                                                                                                            |
| Header clipped by row height   | `estimateWrappedLines` per header against its column width, less 2 character units for the dropdown button | Widest label needs 2 lines; 34pt holds 2. No header is clipped.                                                                                                                                                                                         |
| Fault visible under automation | 9 capture positions: COM `ScrollColumn`, `{RIGHT}` walks, `%{PGDN}`/`%{PGUP}` paging, both directions      | White ink present in the header band at **every** position it was on screen (2.35% to 4.73% of band pixels). Never blank.                                                                                                                               |

## Why the automation cannot see it, and why that is not a refutation

The symptom is _stateful_: clicking one header paints it and unpaints another. No property of a
package can make two cells mutually exclusive, because a file describes cells, not a budget of them.
That signature belongs to the repaint path, and the repaint path is the one thing synthetic
navigation does not exercise faithfully. A COM `ScrollColumn` assignment and a `SendKeys` keystroke
each force a clean, settled repaint; the fault lives in incremental invalidation of the frozen region
under real scrolling.

This is the same boundary `rows-with-no-stated-height-are-autofitted-on-open` states about its own
"no blank bands" result: that note excluded blank bands _on an unfrozen sheet, under synthetic
navigation_, and explicitly could not exclude a frame shorter than its capture latency. The two
notes do not conflict. This one adds the case that method cannot reach, a frozen pane with a
human scrolling it.

## What follows for the writer

**Nothing.** There is no change to the emitted package that is known to help, and the control above
is the reason: the fault survives Excel's own serialization. An agent tempted to "fix" this by
nudging the pane element, dropping `wrapText`, restating the row height, or moving the autoFilter
should read the control result first. Every one of those changes would be a guess dressed as a
fix, and would leave a permanent distortion in the writer paying for a fault it does not cause.

The mitigation is client-side and belongs in advice to the reader, not in the file:

- **Disable hardware graphics acceleration** (File → Options → Advanced → Display). This is the
  standard remedy for Office repaint faults; the setting lives at
  `HKCU\Software\Microsoft\Office\16.0\Common\Graphics\DisableHardwareAcceleration`, and its
  absence means the default, enabled, is in force.
- Update the display driver, or test in Excel safe mode to rule out an add-in.

## Provenance

`source: excel-desktop-verification`, the tier ADR-0013 describes. One Excel build, one host.

The package half is reproducible from the repo: re-save any frozen, autofiltered workbook with
`pwsh -NoProfile -File tools/excel-oracle/observe.ps1 -Path <wb.xlsx>` and diff
`xl/worksheets/*.xml` against the input. The rendering half needed the interactive tier, since
headless COM paints nothing, and was taken by driving a visible Excel through the
`excel-gui-automation` skill (`Set-ExcelForeground` plus window-rect-scoped `Save-WindowScreenshot`),
then sampling the header band's pixels. That half is a recorded observation, not reproducible from
the repo, and its fixtures were scratch files under `.tmp/`.

**The fault itself is a user report, not a repo measurement.** What this repo measured is that the
package is correct and that the fault survives Excel's own re-save. Anyone revisiting this should
start from the control, not from the writer.

Related: `rows-with-no-stated-height-are-autofitted-on-open`,
`grid-geometry-limits-are-excels-not-the-schemas`, `dark-mode-repaints-authored-cell-colors`.
