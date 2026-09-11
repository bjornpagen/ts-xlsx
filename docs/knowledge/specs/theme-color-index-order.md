# `theme="n"` does not index the order the theme part is written in

Cluster: styles

## Scenario

A cell states its colour as `<color theme="1"/>`. Resolving that means indexing the theme part's
`<a:clrScheme>`. ISO/IEC 29500 §20.1.6.2 tabulates the scheme's child sequence explicitly:

| Sequence index | 0     | 1     | 2     | 3     | 4–9                 | 10      | 11         |
| -------------- | ----- | ----- | ----- | ----- | ------------------- | ------- | ---------- |
| Element        | `dk1` | `lt1` | `dk2` | `lt2` | `accent1`–`accent6` | `hlink` | `folHlink` |

That is the order the elements appear in the XML. It is **not** the order SpreadsheetML's `theme="n"`
indexes them. Excel swaps each dark/light pair:

| `theme="n"` | 0     | 1     | 2     | 3     | 4–9                 | 10      | 11         |
| ----------- | ----- | ----- | ----- | ----- | ------------------- | ------- | ---------- |
| Scheme slot | `lt1` | `dk1` | `lt2` | `dk2` | `accent1`–`accent6` | `hlink` | `folHlink` |

Take the spec's listing at face value and every workbook reads inverted: `theme="1"`, which the
stylesheet's own default font carries, resolves to `lt1`, so ordinary black body text comes back
white. The failure is silent and total, and it affects the two most-referenced slots in any file.

## How this was settled

Not by reasoning, but by asking Excel Desktop. Twelve cells were given
`Interior.ThemeColor = MsoThemeColorIndex 1..12`, Excel saved the file, and the emitted `theme="n"`
was read out of `styles.xml`. The same cells were then reopened and `Interior.Color` read back,
giving the concrete RGB Excel renders each index as, against a theme whose twelve slots were all
distinct.

The RGB readback is the decisive half. The names alone settle nothing: Excel's object model calls
index 1 `msoThemeColorDark1`, and that is itself swapped relative to the theme part's `<a:dk1>`, since
`msoThemeColorDark1` writes `theme="0"`, which renders `FFFFFF`, which is `lt1`. Two swaps in
opposite directions is exactly the shape of trap that reading either document alone walks into.

Recorded in `test/corpus/fixtures/excel-oracle/theme-color-index-order.json` (Excel 16.0 build
20131). One build, one host, the usual Tier-3 caveat (ADR 0013).

## Two more corners in the same area

- **`dk1` and `lt1` are almost always `<a:sysClr>`, not `<a:srgbClr>`.** A sysClr's `val` is an
  operating-system colour _name_ (`windowText`), not a value, and its `lastClr` attribute records what
  the authoring application last resolved that name to, which is the only thing a consumer on a
  different system can use. A reader that only understands `srgbClr` resolves nothing for the two
  most-used slots in the scheme.
- **`indexed="64"` and `indexed="65"` are not colours.** They are the system foreground and
  background, with no fixed value. `indexed="64"` in particular sits on the `bgColor` of essentially
  every solid fill Excel writes, so resolving it to black repaints every one of them. They are
  reported unresolved; what "automatic" should look like is the consumer's decision, not ours.

## Where this lives

`THEME_COLOR_SLOTS` in `src/core/theme.ts` is the mapping, in index order. `resolveColor`
(`src/core/color-resolution.ts`, reached as `Workbook.resolveColor`) applies it.

Related: `theme-and-indexed-colors-resolve-to-concrete`,
`indexed-palette-colors-resolve-to-concrete.md`, `theme-color-font-backed-by-theme-part`,
`foreign-theme-part-survives-roundtrip`.
