# A rich-text run's font is all-or-nothing: it inherits nothing from its cell

Cluster: styles

## Scenario

A cell is styled with a typeface and size, say Courier New 16, and its text is then split into
rich-text runs so one phrase can be bolded. The natural way to author that is to give the bold run
the one facet it changes:

```ts
cell.font = {name: 'Courier New', size: 16};
cell.value = {richText: [{text: 'Note:', font: {bold: true}}, {text: ' the rest'}]};
```

The bolded phrase renders in the _wrong face at the wrong size_. Callers hit this while building
themed workbooks and conclude the library is dropping the cell font on rich-text cells; the
workaround is to restate the full font on every single run.

## The format's rule, verified

**It is not a bug.** A run's `<rPr>` (CT_RPrElt) is a _complete_ character-format element. A facet it
omits does not fall through to the cell's font. It falls back to the **workbook default font**, font
id 0 of the styles part.

Observed directly in Excel Desktop 16.0 over COM, reading `Range.Characters(i, 1).Font`, the
_rendered_ font rather than the stored markup. A cell whose own font is Courier New 16, whose first
run carries only `<b/>` and whose second carries a full `<rPr>`:

| characters                                   | rendered                                                   |
| -------------------------------------------- | ---------------------------------------------------------- |
| run 1 (`<rPr><b/></rPr>`)                    | **Calibri 11, bold**, the workbook default, not the cell's |
| run 2 (full `<rPr>`)                         | Courier New 16                                             |
| a plain control cell with the same cell font | Courier New 16                                             |

The control is what rules out a malformed file: the cell font _is_ being applied where there are no
runs. `Range.Font.Name` on the rich cell reports the empty string, Excel's "mixed" answer,
confirming each run carries its own format.

**The fallback is font 0 specifically, not a hardcoded Calibri.** Repeating the probe on a workbook
whose default font was authored as Georgia 14 rendered the bare run as _Georgia 14 bold_. So
`Workbook.defaultFont` is what an unspecified run facet resolves to, and changing it changes how
every bare run renders.

Corroborating evidence in Excel's own output: every `<rPr>` in an Excel-authored `sharedStrings.xml`
is complete, restating `<sz>`, `<color>`, `<rFont>`, `<family>` and `<scheme>` on _every_ run,
including a run whose only difference from its neighbour is the absence of `<b/>`. Excel never relies
on inheritance here because there is none to rely on.

## Desired behavior

- The writer must keep emitting exactly the facets a run carries. Silently merging the cell's font
  into every run would break a caller who deliberately wants a bare run, and would inflate the shared
  string table with a full font on runs that need none.
- The library provides an **explicit, opt-in** way to author the inheriting shape callers expect:
  `Cell.setRichText(runs)` composes each run's partial font over the cell's own, per facet, so a facet
  the run names wins and one it omits comes from the cell. Assigning `cell.value` directly stays the
  bare path.
- A cell that names no font of its own needs no composition: an omitted facet already falls back to
  the workbook default, which is exactly what such a cell renders in. Its runs pass through unchanged
  rather than being stamped with a fabricated base.

## Notes

- This is why the workbook default font matters twice over. Before it was modelled, _every_ bare run
  in a themed workbook rendered in Calibri no matter what the theme said, the same root cause as the
  unstyled-cell symptom (`default-font-workbook-worksheet-level`, ADR-0025), reaching text that
  _does_ have a value.
- Cell-level alignment is the adjacent question and behaves differently. See
  `rich-text-cell-alignment-composition`.
