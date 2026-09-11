# A row's height and a column's width are bounded by Excel, not by the schema

Cluster: worksheet

## Scenario

A caller sizing a line it is about to write, a header band or a row wrapping a long paragraph,
needs to know how large it may go. The format is no help: `CT_Row@ht` and `CT_Col@width` are both
a bare `xsd:double` in `sml.xsd`, with no `maxInclusive` on either. The ceiling is Excel's, and it
turns out to be a ceiling on _setting_ a size rather than on having one, which is the two halves of
this note. What Excel refuses from a user is not what it refuses from a file.

Microsoft's published _Excel specifications and limits_ table states "Row height: 409 points" and
"Column width: 255 characters". The first of those is rounded.

## Measured behavior: what Excel accepts being set

Excel Desktop (Microsoft 365, Windows, workbook default font Aptos Narrow 11), driven over COM
against a blank workbook, assigning then reading back:

| Assigned `RowHeight` | Result                                                             |
| -------------------- | ------------------------------------------------------------------ |
| 409                  | accepted, reads back 409                                           |
| 409.4                | accepted, reads back 409.4                                         |
| **409.5**            | **accepted, reads back 409.5**                                     |
| 409.6                | refused, "Unable to set the RowHeight property of the Range class" |
| 410, 500             | refused                                                            |

| Assigned `ColumnWidth` | Result                       |
| ---------------------- | ---------------------------- |
| **255**                | **accepted, reads back 255** |
| 255.4, 255.5, 255.9    | refused                      |
| 256, 300               | refused                      |

So the row-height ceiling is 409.5 and the published 409 is a rounding of it; the column-width
ceiling is exactly integral at 255, which is why one constant carries a fraction and the other
does not. `MAX_ROW_HEIGHT` and `MAX_COLUMN_WIDTH` (`src/core/limits.ts`) hold these, and
`src/core/limits.test.ts` pins them with a comment pointing back here, so a future reader who
"corrects" 409.5 to 409 against the specifications page fails that test and finds this note.

The same probe recorded `StandardWidth = 8.09` and `StandardHeight = 14.5` for that workbook,
against the familiar 8.43 and 15 of a Calibri-11 one. That is the second fact here, and the reason
no `DEFAULT_COLUMN_WIDTH` constant sits beside the two ceilings: the default width is a function
of the workbook's default font, not a property of the format. See
`default-font-must-not-be-assumed-for-column-widths`.

## Measured behavior: what Excel does with a file that exceeds them

The ceilings above are what the object model and the UI refuse. A _package_ is read by a different
path, which refuses nothing. Fixtures written by this library carrying over-limit `ht` and `width`,
opened in interactive Excel Desktop on the same build and host, then read back over COM and re-saved
by Excel itself:

| In the file     | Interactive open | COM reads back | Excel's own re-save writes |
| --------------- | ---------------- | -------------- | -------------------------- |
| `ht="5000"`     | clean            | 409.6          | `ht="409.6"`               |
| `ht="600"`      | clean            | 409.6          | `ht="409.6"`               |
| `ht="409.6"`    | clean            | 409.6          | `ht="409.6"`               |
| `ht="409.5"`    | clean            | 409.5          | `ht="409.5"`               |
| `width="1000"`  | clean            | 999.36         | `width="1000"`             |
| `width="300"`   | clean            | 299.36         | `width="300"`              |
| `width="255.4"` | clean            | 254.73         | `width="255.36328125"`     |

"Clean" is the full interactive verdict: no repair prompt, no `[Repaired]` in the title, no
`error*.xml` written beside the file, no format-mismatch warning. An all-legal control file was
run through the same classifier and also came back clean, so the verdict distinguishes.

Two asymmetric facts fall out.

**A row height is silently clamped on read, to 409.6**, a tick _above_ the 409.5 the setter
accepts. 409.6 points is 8192 twentieths of a point, that is 2^13, so the file-read path saturates
the field it parses into where interactive assignment refuses to reach it. Excel then persists the
clamp: its own save writes `ht="409.6"`, and the authored 5000 is gone.

**A column width is not clamped at all.** `width="1000"` survives Excel's own save byte-identical.
The one rewrite, `255.4` to `255.36328125`, is `65373/256`, ordinary quantization to 1/256 of a
character unit, not a cap. The -0.64 offset in the COM readback is the normal round-trip and not
evidence of clamping: the control file shows it on legal widths too (30 reads 29.36, 12 reads
11.36), and 1000 reading 999.36 fits the same rule.

Rendering is cosmetic in both directions, and the earlier claim here that such a sheet opens with
blank unpainted bands was an inference, not an observation. It does not. A row taller than the
viewport puts its own text below the fold, since cells are bottom-aligned by default, and leaves the
row-number gutter blank; a column wider than the viewport draws its data correctly and only loses
its header letter. Nothing a scroll or a click does not resolve, and no data is lost either way.

## Why nothing enforces them

`Row.height` and `Column.width` accept values above these ceilings, deliberately, and so does the
writer. The setters are not an authoring-only path: `read-worksheet.ts` loads a foreign
`ht` or `width` straight through them, so a bound that threw would refuse to read files Excel opens
clean, and, for a width, files Excel round-trips more faithfully than we would. The constants are
published so a caller can check before writing, and the two setters' doc comments point at them.

Refusing at the _writer_ instead was considered and rejected on the same evidence. An
`AuthoringError` above `MAX_ROW_HEIGHT` would decline to emit a file Excel opens without complaint
and treats as 409.6; above `MAX_COLUMN_WIDTH` it would be plainly wrong, since that limit does not
bind a file at all.

So an authoring mistake, a 5000-point row, stays undetected by the library. Its consequence is now
known and small: the row is clamped to 409.6 by whoever opens it. The remaining alternative, an
internal restore path for the reader so the public setter can validate, buys a diagnostic for
that one case at the cost of a second way to set the same property, and is worth revisiting only
if the mistake turns out to be common.

## Provenance

`source: excel-desktop-verification`. One Excel build, one host, state-observable via COM
assignment, the tier that ADR-0013 describes. Not run in CI and not reproducible without Excel;
the numbers are locked by the unit test instead.

The file-read half additionally needed the interactive tier, because "opens clean" is precisely
what headless COM cannot report: `DisplayAlerts = $false` suppresses the repair modal, so an
automation-only run could confirm that nothing threw but never that nothing was asked. It was run
through the `excel-gui-automation` skill's `open-verdict.ps1`, one file at a time, paired with an
all-legal negative control so a classifier that always answered "clean" would have been caught.
The clamped and preserved values and the re-save come from a COM readback of `RowHeight`,
`ColumnWidth` and `SaveAs` over the same fixtures. That half is reproducible:
`pwsh -NoProfile -File tools/excel-oracle/read-geometry.ps1 -Path <file.xlsx>` is the probe that
took it, and re-running it against a fixture carrying `ht="5000"` or `width="1000"` reproduces both
tables. The interactive half is not reproducible from the repo and stays a recorded observation.
