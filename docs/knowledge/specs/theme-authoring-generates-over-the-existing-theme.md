# Authoring a theme generates over the existing one, and stops at the format scheme

Cluster: styles

## Scenario

"Give me our brand colours" is a theme-level request, not a cell-level one. A colour picked from
Excel's theme row is written as `theme="4"`, a _reference_ resolved at render time, so setting
`accent1` restyles every cell, chart and table style that follows the theme at once, and is the only
way to recolour a workbook without touching a single cell. `Workbook.setTheme({colors, fonts})` is
that surface.

## Why generation is a splice, not a rebuild

A theme part has three blocks. Only two of them are anybody's business here:

| Block            | Authorable                   | Why                                                            |
| ---------------- | ---------------------------- | -------------------------------------------------------------- |
| `<a:clrScheme>`  | yes                          | twelve colours; the thing people actually ask for              |
| `<a:fontScheme>` | yes (major/minor latin face) | a font that says `scheme="minor"` names no typeface of its own |
| `<a:fmtScheme>`  | **no**                       | gradient stops, line styles, effect styles                     |

Nobody hand-authors `fillStyleLst` gradient stops from a spreadsheet API, and modelling them is a
large surface for approximately zero demand. More importantly, regenerating the part would _replace_ a
designer's format scheme with the Office default, a silent downgrade of the file. So authoring
rewrites the two blocks it understands, inside the part it was given, and leaves everything else byte
for byte.

The base is the workbook's own theme when it has one (see `foreign-theme-part-survives-roundtrip`),
otherwise the library's default. That answers the question the preservation work left open: a
preserved theme and an authored override are not in conflict, because the preserved part _is_ the
base the override is applied to. The theme's own relationships, such as a picture used as a themed
fill, survive for the same reason: nothing about the part's identity or its rels changes.

## An unauthored slot keeps its encoding, not just its value

`dk1` and `lt1` are almost always written as `<a:sysClr val="windowText" lastClr="000000"/>`: they
_follow the viewer's system colours_, and `lastClr` only records what the authoring machine last
resolved them to. Re-serialising an unauthored slot from its parsed value would turn that into
`<a:srgbClr val="000000"/>` and pin the workbook to one machine's window colours. So a slot the caller
did not name is re-emitted as its verbatim source element.

An authored slot is always `<a:srgbClr>`: a caller supplying `dk1: '1A1A1A'` is asking for that exact
colour, not for a system colour that happens to resolve there today.

## Two smaller decisions

- **Colours are validated at the setter, not at write time.** A malformed value does not error in
  Excel, since the slot renders as flat black, so the library has to, and doing it far from the call
  that supplied it is no use. `RRGGBB`, `#RRGGBB` and `AARRGGBB` are all accepted, with the alpha
  dropped because DrawingML's `<a:srgbClr val>` has no alpha channel, and anything else throws.
- **`setTheme` merges.** Branding one accent leaves the other eleven alone, and two calls accumulate.
  A replace-everything setter would make "change accent1" require restating all twelve slots.

## Not modelled

The colour scheme's `name` attribute, what Excel's theme gallery labels it, is left as the source
wrote it. Deliberate: it is a UI label with no rendering effect, and the API is already the honest
description of what changed.

## Where this lives

`applyThemeOverrides` in `src/io/xlsx/theme-xml.ts`; `Workbook.setTheme`, `themeColors`,
`themeFonts` and `themeOverrides` on `src/core/workbook.ts`. The model carries what was authored;
the writer composes it onto the part at plan time, so a preserved theme's closure is unaffected
(`planPreservedParts`, `src/io/xlsx/package-plan.ts`).

Related: `authored-theme-palette-reaches-cells`, `foreign-theme-part-survives-roundtrip`,
`theme-color-index-order.md`, `theme-and-indexed-colors-resolve-to-concrete`.
