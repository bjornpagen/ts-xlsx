# Indexed-palette colors must resolve to a concrete color, not surface a raw index

Cluster: styles

## Scenario

A caller inspects a cell's fill, font or border color and finds, instead of a usable `argb` value, a
bare integer `indexed` reference, say a `bgColor` carrying `indexed="64"` with no RGB. They have no
way to turn that number into an actual color. This happens because OOXML permits several color
encodings (direct ARGB, theme-color-plus-tint, and a legacy 64-entry **indexed palette** carried
over from the old BIFF/`.xls` format) and the library hands back the raw indexed reference without
resolving it. Files produced by older tools, or by converters that emit the legacy palette, are full
of these.

> **Shipped.** `Workbook.resolveColor` resolves `indexed`, `theme`, and `tint` against the workbook's
> own palette and theme, and `theme-and-indexed-colors-resolve-to-concrete` locks it. What follows is
> the contract as it was specified, annotated with how each point was answered.

## Desired behavior

- **Every color the public model exposes is resolvable to a concrete ARGB** without the caller knowing
  the encoding. An `indexed` reference resolves through the palette, the standard 64-entry default or
  a workbook-supplied `<indexedColors>` override, to an ARGB value the caller can read and render.
- **The encoding is not lost.** Resolving to ARGB for consumption must not erase the fact that the
  source was indexed or themed, so a color carries both its resolved ARGB and its original reference
  and a round-trip can re-emit the same encoding rather than rewriting every color as literal ARGB,
  bloating the styles table and changing the file's semantics.
- **The default palette is built in.** The legacy 64-entry palette, including the special
  `indexed="64"` and `65` "automatic" foreground and background system colors, is known to the library,
  so a file that references it without shipping a custom `<indexedColors>` block still resolves
  correctly.
- **A workbook-level custom indexed palette is honored.** When the styles part declares its own
  `<indexedColors>`, those entries override the defaults for that workbook on both read and write.

## How it shipped

- **The public shape.** Resolution is a _method_, not a field: `Workbook.resolveColor(color)` returns
  an 8-hex ARGB or `undefined`. `Color` is untouched, so "resolved ARGB" and "original encoding" cannot
  drift apart and there is no union for a caller to narrow. Choosing a method over a `.argb` getter on
  `Color` also keeps the resolution _context_, meaning which workbook's theme and palette, explicit,
  since the same `Color` value resolves differently in two workbooks.
- **A custom palette replaces the built-in one wholesale**, rather than overlaying it. That is what
  `<indexedColors>` means, and Excel writes all 64 entries whenever it writes any. An index past the
  end of a short custom palette therefore resolves to nothing rather than falling through to a
  built-in entry the workbook explicitly overrode away.
- **`indexed="64"` and `65` resolve to `undefined`.** They name the system foreground and background,
  which have no fixed value; `64` sits on the `bgColor` of essentially every solid fill Excel writes,
  so resolving it to black would repaint them all. What "automatic" should look like is the consumer's
  decision.
- **Alpha.** The spec tabulates each palette entry with a leading `00`. That byte is an artefact of the
  32-bit colour records the palette came from, not transparency, so a resolved `indexed` or `theme`
  colour comes back fully opaque. An explicit `rgb` keeps whatever alpha the file stated.
- **Never down-converted on write.** Nothing converts an authored ARGB into a palette entry. Indexed
  is a read-tolerance and round-trip-fidelity concern only.
- **Theme resolution composes on the same principle**, and carries its own trap, which has its own
  note: see `theme-color-index-order.md`. `tint` is applied last, on whatever the base resolved to.

Related: `custom-indexed-color-palette-roundtrip`, `theme-and-indexed-colors-resolve-to-concrete`,
`theme-color-index-order.md`, `theme-color-font-backed-by-theme-part`,
`public-type-surface-matches-runtime`.
