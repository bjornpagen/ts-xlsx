# Workbook Styles

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `DifferentialStyle`

<sub>type</sub>

A differential style (OOXML CT_Dxf): formatting laid *over* whatever a cell already carries. Only
the facets present override; the rest of the cell's own style shows through. It carries the subset
of the cell-style facets (see [`CellStyle`](./styles.md#cellstyle)) a `<dxf>` can express: font, number format, fill,
and border.

Differential styles live in one workbook-level table (`<dxfs>`) that several features index into:
a conditional-formatting rule's highlight format, and a table style's per-element formatting
(`<tableStyleElement dxfId="…">`). They are interned and shared, so two features asking for the
same formatting land on one entry.

Not every facet reaches every consumer. As a **table style element**, Excel applies only the font,
fill, and border: its own object model exposes `Font`, `Interior`, and `Borders` on a table style
element and nothing for a number format, so a `numFmt` set here is carried faithfully through a
round-trip but has no visible effect. The type is left whole rather than split, because the same
value is legitimately reused across both consumers and narrowing it would only move the surprise.

```ts
type DifferentialStyle = Pick<CellStyle, 'font' | 'numFmt' | 'fill' | 'border'>;
```

---

### `NamedCellStyle`

<sub>type</sub>

A named cell style: the OOXML `cellStyleXfs`/`cellStyles` layer. A spreadsheet applies a built-in
or custom style (e.g. "Normal", "Accent1") whose visual facets live in this shared, named layer
rather than on each cell's direct format; a cell links to it and inherits any facet the direct
format leaves unset. The facets are a cell's own (see [`CellStyle`](./styles.md#cellstyle)); `name` is the style's
display name and `builtinId` its Excel gallery index when it is a built-in style.

```ts
type NamedCellStyle = Readonly<CellStyle> & {
  readonly name?: string;
  readonly builtinId?: number;
};
```

---

### `TableStyleDefinition`

<sub>interface</sub>

One preserved `<tableStyle>` definition: its bytes, and the `name` a table's style info matches it by.

The name travels beside the fragment rather than being read back out of it. An authored style
replaces the preserved definition it shares a name with, and a pattern pulling `name="…"` out of
the bytes did not see `name='…'`, which is the same attribute in single quotes: the writer emitted
both definitions under one name, the duplicate [`Workbook.addTableStyle`](./workbook.md#workbookaddtablestyle) promises cannot happen.
The reader takes the name from the scanner that read the element, decoded, whatever its quoting.

```ts
interface TableStyleDefinition {
  /** The style's `name`, decoded; empty for a fragment that declared none. */
  readonly name: string;
  /** The `<tableStyle>…</tableStyle>` element exactly as the file spelled it. */
  readonly xml: string;
}
```

---

### `TableStyleNamespace`

<sub>interface</sub>

One namespace declaration a preserved `<tableStyle>` fragment depends on.

```ts
interface TableStyleNamespace {
  readonly prefix: string;
  readonly uri: string;
  /** Whether the source listed this prefix in the stylesheet's `mc:Ignorable`. */
  readonly ignorable: boolean;
}
```

---

### `TableStyleTable`

<sub>interface</sub>

The `<tableStyles>` block of a styles part: the custom table/pivot style definitions a file
declares, and the two gallery names it nominates as the default for a new table and a new pivot.

Each entry of `styles` is one `<tableStyle>…</tableStyle>` fragment kept verbatim, for the
same reason a `<dxf>` is: a `tableStyleElement`'s `dxfId` indexes the differential-style table,
which the writer re-emits **at its original indices**, so the references stay valid without
reparsing anything. That index-stability is load-bearing: renumbering the dxf table would
silently re-point every preserved table style at a different format.

The two default names are ordinary strings, not fragments: they are re-escaped on write, so they
are held decoded.

```ts
interface TableStyleTable {
  readonly styles: readonly TableStyleDefinition[];
  readonly defaultTableStyle?: string | undefined;
  readonly defaultPivotStyle?: string | undefined;
  /**
   * The namespace prefixes the verbatim {@link styles} fragments use, mapped to their URI and to
   * whether the source marked the prefix ignorable (`mc:Ignorable`).
   *
   * Carrying a fragment verbatim carries its *prefixes* too. Excel stamps a revision id
   * (`xr9:uid="{…}"`) on every `<tableStyle>` it writes, so a fragment re-emitted under a
   * `<styleSheet>` that declares only the default namespace is not namespace-well-formed, and no
   * consumer can parse the part at all, which is a far louder failure than the dropped table style
   * this preservation exists to prevent. The writer re-declares each prefix on `<styleSheet>` and
   * re-states the ignorable ones, exactly as the source did.
   */
  readonly namespaces?: readonly TableStyleNamespace[];
}
```
