# A worksheet-level (and workbook-level) default style, mapped to OOXML defaulting

Cluster: styles

## Scenario

A user building a worksheet with many columns wants to set a base look (font, fill, alignment,
number format, border) **once**, at the worksheet or workbook level, rather than applying the same
style object to every column or cell. Today a uniform look across a large sheet means either looping
over all columns and cells to assign the identical style or reaching for OOXML's default column and
row formatting, which the API does not expose ergonomically. A default-style facility would let
authors declare "unless overridden, cells in this sheet use this style" and then specify only the
deviations.

> Spec note, not a corpus case: this is a missing authoring surface, not a current-behavior bug. It is
> the umbrella over two facet-specific notes already recorded, a default _font_
> (`default-font-workbook-worksheet-level`) and default _cell protection_
> (`worksheet-default-cell-protection-unlock`), generalizing the same defaulting mechanism to the
> whole style surface. It becomes corpus-covered once the surface exists and a case can assert a
> default-styled sheet renders correctly with no per-cell style stamping.

## Desired behavior

- **Declare a default style once.** Expose a way to set a default style (font, fill, alignment,
  numFmt, border) scoped to a worksheet, applied as the base for every cell, column and row that does
  not specify its own. Cell- and column-level styles override the worksheet default, and the worksheet
  default overrides any workbook-wide default. This subsumes the existing default-font ask and the
  default-protection ask as facets of one mechanism.

- **Map to OOXML defaulting, do not materialize per cell.** The format already models this: the
  `cellXfs` index 0 plus the "Normal" cell style is the document-wide default; `<sheetFormatPr>`
  carries default row height and base column width; and `<cols>` `<col>` spans carry a default style
  and width. A well-designed API surfaces a workbook-level default feeding the Normal style and
  default `cellXf`, and a worksheet-level default mapping to `<col>` default spans and
  `<sheetFormatPr>`. So the feature is largely a mapping onto existing defaulting rules, **not**
  stamping the same style id onto every cell. Writing a default-styled sheet must stay O(deviations),
  not O(cells): a worksheet whose default is a custom font must not emit a per-cell style reference
  for every cell that merely inherits it. It shares the size contract of
  `worksheet-default-cell-protection-unlock`.

- **Round-trip stability.** Reading a file that uses a non-default base style surfaces it through the
  same API used to set it, so read-modify-write preserves the default rather than flattening it into
  explicit per-cell styles.

## Open questions

- Granularity: worksheet-level only, or also a workbook-level default? The request is phrased at
  worksheet level, but the underlying Normal-style default is workbook-wide, so likely both, with
  worksheet overriding workbook.
- Which facets are defaultable at each level? Is a default _border_ or _fill_ meaningful at the
  `<col>` and `<sheetFormatPr>` level, or only font, numFmt and alignment, with fill and border
  requiring a real per-column style span?
- Precedence order when workbook default, worksheet default, column style, row style, and cell style
  all touch the same facet: define and lock the resolution chain.
- How this composes with value-based style dedup (`style-dedup-value-based-and-cell-add-style`), where
  the default is the base that per-cell deviations are diffed against.

Related: `default-font-workbook-worksheet-level`, `worksheet-default-cell-protection-unlock`,
`style-dedup-value-based-and-cell-add-style`, `set-style-over-cell-range`,
`column-declared-numfmt-reaches-cells`, `default-font-must-not-be-assumed-for-column-widths`.
