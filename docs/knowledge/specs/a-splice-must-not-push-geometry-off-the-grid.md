# A splice must not push geometry off the grid

Cluster: tables

## Scenario

A worksheet's regions are commonly anchored to the last row. A whole-column dropdown is written
`B1:B1048576`, not `B:B`, because that bounded spelling is what Excel itself writes; a full-height
autofilter and a full-height conditional format are the same shape. Insert a row anywhere above
them and every edge below the insert wants to move down. The bottom ones have nowhere to go: the
grid ends at row 1048576, and the column axis ends at XFD.

Before this was fixed, `shiftIndex` applied no ceiling, so `spliceRows(900, 0, …)` over that sheet
produced `sqref="B1:B1048578"`, `mergeCell ref="C1048577:D1048578"` and `autoFilter ref="A1:A1048578"`:
references naming rows that cannot exist. The question this note answers is what such a file does in
Excel, and what Excel does to the same regions on the same edit, because the second is the answer
worth copying.

## Measured: what Excel does with a file whose geometry runs off the grid

The schema oracle passes it. `OpenXmlValidator` (Open XML SDK 3.5.1, `Microsoft365` profile) reports
the package **valid**: `ST_Sqref` and `ST_Ref` are string types, and neither carries a row bound the
validator can check against.

Excel does not. Excel Desktop (Microsoft 365, Windows, build 16.0.20228), one over-limit region per
file so the verdict names a culprit:

| File | Automation open (`tools/excel-oracle/observe.ps1`) | Interactive open (`open-verdict.ps1`) |
| --- | --- | --- |
| `sqref="B1:B1048578"` (validation) | throws, no workbook | **repair-prompt** |
| `ref="C1048577:D1048578"` (merge) | throws, no workbook | not run |
| `autoFilter ref="A1:A1048578"` | throws, no workbook | not run |
| control: the same three regions bounded at 1048576 | opens, `repaired: false`, cells read back | **clean** |

The interactive dialog is the ordinary corruption one: *"We found a problem with some content in
'over-validation.xlsx'. Do you want us to try to recover as much as we can?"* The control was run
through the same classifier and came back `clean` with no dialog, so the verdict distinguishes. This
is the whole workbook's content put at the user's mercy by an edit that touched none of it, which is
why this is a defect rather than a judgement call about a region's size.

## Measured: what Excel does to the same regions on the same edit

The reverse direction, and the one that decides the fix. A workbook carrying one region was opened
over COM and given `Rows("900:901").Insert()`, then saved by Excel and read back as XML:

| Region present | Excel's insert | What Excel then wrote |
| --- | --- | --- |
| validation `B1:B1048576` | succeeds | `sqref="B1:B1048576"`, the bottom edge left on the last row |
| autofilter `A1:A1048576` | succeeds | `ref="A1:A"`, Excel's own canonical spelling, with no over-limit row |
| merge `C1048575:D1048576` | **refused** | *"Microsoft Excel can't insert new cells because it would push non-empty cells off the end of the worksheet…"* |
| control: neither, data in `A1`/`A5` | succeeds | nothing to report |

So Excel has two answers, and which it gives depends on whether what would be pushed off is a
*region* or *content*. A region is clamped: its bottom edge stays where it is and the region absorbs
the loss. Content is not clamped, and rather than damage it Excel refuses the whole insert.

## What this library does

`shiftIndex` (`core/grid-shift.ts`) takes the axis and clamps its result to that axis's last line, so
every re-anchoring pass in the model inherits the rule from one place rather than six. That matches
Excel exactly for the region case, which is the common one: a whole-column validation comes back
`B1:B1048576` from either program.

For the content case this library already refuses, and by a different mechanism: relocating a cell
constructs a `Cell` at the destination, and `assertRowInBounds` throws a `RangeError` before the row
map is replaced, so a splice that would push a cell off the end fails with the sheet unchanged.
Excel's message is friendlier; the outcome is the same. An inserted line with no content builds no
`Cell`, so it cannot lean on that: `spliceRows`, `spliceColumns` and `duplicateRow` check where their
inserted lines land before anything moves. Before they did, `spliceRows(1048576, 0, [], [])` and a
duplicate of a height-only last row each left a line key past the grid, and the sheet threw on every
read or write after that.

The column axis is decided by the same sentence, and it is worth writing down because the two
answers look inconsistent side by side. `spliceColumns(16384, 0, a, b)` refuses with `column 16385 is
out of bounds` and leaves the sheet untouched, while a validation or an autofilter on the same sheet
would have been clamped. That is the region/content split, not a gap: the second inserted column is
content with nowhere to go, and the only clamp available would stack it on top of the first at XFD
and report success. An inserted column whose values reach past row 1048576 is refused for the same
reason, and `addColumn` refuses that argument identically, so the two column-append doors agree.
Locked by the two `an insert with nowhere to put a column` / `an inserted column reaching past the
last row` tests in `src/core/grid-edits.test.ts`.

That leaves one deliberate deviation, in a merge occupying the very last rows. Excel refuses the
insert; we clamp, so `C1048575:D1048576` becomes `C1048576:D1048576`, one row shorter. Refusing at
that point is worse for a library: the splice has already been accepted for the cells and every other
region, so the caller would face a half-legal edit, and the merge is the only participant that would
have raised the objection. A region one row shorter is a legible loss in a file that opens; the
alternative measured above is a file that does not.

## Provenance

`source: excel-desktop-verification`. One Excel build, one host, per ADR-0013.

The reproducible half is `tools/excel-oracle/observe.ps1` against a package this library writes, plus
`ooxml-validate` for the schema verdict; both re-run from the repo. The insert half was taken with a
throwaway COM script (open, `Rows("900:901").Insert()`, `SaveAs`, read the part back) and reproduces
in a few lines against any workbook carrying those regions. The interactive verdict came from the
`excel-gui-automation` skill's `open-verdict.ps1`, paired with the all-legal control above, and is a
recorded observation rather than a reproducible artifact.

The lock is `test/corpus/cases/splice-holds-geometry-inside-the-grid.case.ts` plus the unit tests in
`src/core/grid-shift.test.ts` and `src/core/grid-edits.test.ts`; none of them needs Excel.
