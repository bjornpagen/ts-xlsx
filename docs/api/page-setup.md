# Page Setup

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `HeaderFooter`

<sub>interface</sub>

Page header/footer text, one string per page class. Excel only honours the even- and
first-page variants when the writer also sets the gating flags (`differentOddEven`,
`differentFirst`); the writer derives those from which variants are present. An empty
object means the element is omitted entirely.

```ts
interface HeaderFooter {
  oddHeader?: string;
  oddFooter?: string;
  evenHeader?: string;
  evenFooter?: string;
  firstHeader?: string;
  firstFooter?: string;
}
```

---

### `isPageOrder`

<sub>const</sub>

Narrow a raw `<pageSetup pageOrder>` token to a known [`PageOrder`](./page-setup.md#pageorder).

```ts
const isPageOrder: (value: string) => value is PageOrder
```

---

### `isPageOrientation`

<sub>const</sub>

Narrow a raw `<pageSetup orientation>` token to a known [`PageOrientation`](./page-setup.md#pageorientation).

```ts
const isPageOrientation: (value: string) => value is PageOrientation
```

---

### `PageBreak`

<sub>interface</sub>

A page break (`<brk>`). It falls after line `id` and before line `id + 1`, rows for a row break and
columns for a column break, which is how Excel counts it: a break Excel shows above row 20 is saved
as `id="19"`. `min` and `max` bound its extent across the other axis, zero-based, and Excel writes
only the whole axis (`max` 16383 for a row break). `man` marks an author-set break rather than one
a producer recorded as automatic. The model keeps whatever the source carried.

A row or column splice moves a break with the line after it and drops the break when that line is
deleted, which is what Excel Desktop does: driven over COM, a manual break above row 10 moved above
row 11 when a row was inserted at row 1, and went away when row 11 was then deleted, and a column
break moved the same way. The extent is not moved, because the only one Excel writes is the whole
axis, and a splice should leave that whole rather than shorten it by the lines it removed.

```ts
interface PageBreak {
  /** The last row (or column) before the break. */
  readonly id: number;
  /** The break's near extent across the other axis, zero-based, if the source declared one. */
  readonly min?: number;
  /** The break's far extent across the other axis, zero-based, if the source declared one. */
  readonly max?: number;
  /** Whether the break is manual. A break without it is written as a manual one: Excel stores no other kind. */
  readonly man?: boolean;
}
```

---

### `PageMargins`

<sub>interface</sub>

Print margins, in inches. OOXML's `<pageMargins>` requires all six to be present, but
the model stores only what the caller set; the writer fills the untouched ones with
valid defaults. An empty object means the element is omitted entirely.

```ts
interface PageMargins {
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
  header?: number;
  footer?: number;
}
```

---

### `PageOrder`

<sub>type</sub>

The order pages are numbered and printed in across a sheet wider and taller than one page.

```ts
type PageOrder = 'downThenOver' | 'overThenDown';
```

---

### `PageOrientation`

<sub>type</sub>

Paper orientation, as `<pageSetup orientation>` carries it.

`ST_Orientation` has a third member, `default`, which means "whatever the printer decides" and is
indistinguishable from the attribute being absent. The model spells that absence as an unset field,
so a file carrying `default` reads back with no orientation and writes back without the attribute.

```ts
type PageOrientation = 'portrait' | 'landscape';
```

---

### `PageSetup`

<sub>interface</sub>

Print-scaling and orientation settings. These map onto two OOXML elements: `fitToPage` is the
`<pageSetUpPr>` flag (a `<sheetPr>` child) that switches Excel from fixed-zoom to fit-to-page
scaling, while the rest are `<pageSetup>` attributes. Excel honours `scale` only when `fitToPage`
is off and the `fitToWidth`/`fitToHeight` page counts only when it is on, but the model carries
whatever the author set: an unset field is omitted so a round-trip never fabricates one. An
empty object emits neither element.

```ts
interface PageSetup {
  /** Switch to fit-to-page scaling. Emitted as `<pageSetUpPr fitToPage="1">`. */
  fitToPage?: boolean;
  /** Pages wide to fit onto; `0` means "unbounded" (fit only by height). */
  fitToWidth?: number;
  /** Pages tall to fit onto; `0` means "unbounded" (fit only by width). */
  fitToHeight?: number;
  /** Fixed print zoom as a percentage; Excel honours it only when `fitToPage` is off. */
  scale?: number;
  /** Paper orientation. */
  orientation?: PageOrientation;
  /** Order pages are numbered/printed in across a multi-page sheet. */
  pageOrder?: PageOrder;
  /**
   * Paper size as Excel's 1-based enumeration index (e.g. `9` = A4, `1` = US Letter). Carried as an
   * opaque integer: the model does not map it to physical dimensions, only preserves whatever the
   * author or source file set.
   */
  paperSize?: number;
  /**
   * The printer-settings blob a source file bound to this sheet's `<pageSetup>` via an `r:id`
   * relationship, held verbatim. Excel stores the platform-specific `DEVMODE` (paper tray, duplex,
   * DPI, …) in this opaque binary part; the model does not interpret it, only round-trips the exact
   * bytes so re-writing a file that carried one does not silently drop the user's print configuration.
   */
  printerSettings?: Uint8Array;
}
```

---

### `PrintOptions`

<sub>interface</sub>

Print-toggle flags from the `<printOptions>` element. Each maps to a boolean OOXML attribute that
defaults false, except `gridLinesSet`, which defaults true and gates whether `gridLines` is
honoured. The model stores only what the source or caller set, so an unset flag is omitted and a
round-trip never fabricates one; an empty object emits no element at all.

```ts
interface PrintOptions {
  /** Centre the printed content horizontally on the page. */
  horizontalCentered?: boolean;
  /** Centre the printed content vertically on the page. */
  verticalCentered?: boolean;
  /** Print the row and column headings (the `1,2,3…` / `A,B,C…` gutters). */
  headings?: boolean;
  /** Print the cell gridlines. */
  gridLines?: boolean;
  /** Whether the `gridLines` flag is authoritative; when `false`, Excel ignores `gridLines`. */
  gridLinesSet?: boolean;
}
```
