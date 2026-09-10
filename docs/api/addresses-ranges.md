# Addresses & ranges

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `CellAddress`

<sub>interface</sub>

A decoded single-cell reference. An axis the reference omits is `undefined`.

```ts
interface CellAddress {
  /** Canonical A1 form with `$` anchors stripped: e.g. `"B2"`, `"1"`, `"A"`. */
  readonly address: string;
  /** 1-based column, or `undefined` for a row-only reference (`$1`). */
  readonly col: number | undefined;
  /** 1-based row, or `undefined` for a column-only reference (`$A`). */
  readonly row: number | undefined;
}
```

---

### `CellPosition`

<sub>interface</sub>

A reference that names one cell, both axes present. The narrowing of [`CellAddress`](./addresses-ranges.md#celladdress) that
most callers actually want: `decodeAddress` is deliberately three-shaped because a bare row
(`$1`) and a bare column (`$A`) are legitimate references, but a cell is where a value lives, and
every caller that needs one was re-deriving that invariant by hand.

```ts
interface CellPosition {
  readonly col: number;
  readonly row: number;
}
```

---

### `columnToNumber`

<sub>function</sub>

Convert column letters to a 1-based number (`"A" → 1`, `"AA" → 27`).

```ts
function columnToNumber(letters: string): number;
```

**Throws:** `RangeError` if the letters are malformed or name a column past `XFD`. The two are
separate messages because a caller fixes them differently.

---

### `decodeAddress`

<sub>function</sub>

Decode a single cell/row/column reference into `{address, col, row}`. Anchoring
`$` signs are accepted and dropped; an absent axis is `undefined`.

```ts
function decodeAddress(reference: string): CellAddress;
```

**Throws:** `SyntaxError` if the reference mentions neither a column nor a row.
**Throws:** `RangeError` if it names a column past `XFD` or a row outside `1..1048576`.

---

### `decodeRange`

<sub>function</sub>

Decode a range reference (`A1:B2`, `$1:$1`, `Sheet1!$A:$A`) into its corners and
canonical dimensions. A single reference collapses to a degenerate range whose
corners coincide.

```ts
function decodeRange(reference: string): RangeAddress;
```

**Throws:** `SyntaxError` if an endpoint is unparseable.
**Throws:** `RangeError` if an endpoint names a column past `XFD` or a row outside `1..1048576`.

---

### `encodeAddress`

<sub>function</sub>

Encode a 1-based `col`/`row` pair into its canonical A1 address (`"B2"`).

Both axes go through the shared guard. The column already did, through `numberToColumn`; the row
checked only its lower bound in a message of its own, so `encodeAddress(1, 1048577)` produced an
address naming a row Excel has no reference for while `encodeAddress(16385, 1)` refused.

```ts
function encodeAddress(col: number, row: number): string;
```

---

### `GridRect`

<sub>interface</sub>

A rectangular block of the grid, as **inclusive** 1-based bounds on both axes.

One declaration because inclusive-first/last is the convention every range-shaped thing in this
library follows, and three copies of a convention are three places it can drift. A merged region,
a table's extent and a [`Range`](./range.md#range) handle are all this shape; what differs between them is what
the rectangle *means*, which is what their own names carry.

```ts
interface GridRect {
  /** 1-based row of the top edge. */
  readonly top: number;
  /** 1-based column of the left edge. */
  readonly left: number;
  /** 1-based row of the bottom edge, inclusive. */
  readonly bottom: number;
  /** 1-based column of the right edge, inclusive. */
  readonly right: number;
}
```

---

### `MAX_COLUMN`

<sub>const</sub>

Excel's column bounds: `A` (1) through `XFD` (16384).

```ts
const MAX_COLUMN: 16384
```

---

### `MAX_ROW`

<sub>const</sub>

Excel's row bound: 1 through 1048576. The other axis of [`MAX_COLUMN`](./addresses-ranges.md#maxcolumn).

```ts
const MAX_ROW: 1048576
```

---

### `numberToColumn`

<sub>function</sub>

Convert a 1-based column number to its letters (`1 → "A"`, `27 → "AA"`).

```ts
function numberToColumn(n: number): string;
```

**Throws:** `RangeError` unless `n` is an integer in `1..MAX_COLUMN`.

---

### `RangeAddress`

<sub>interface</sub>

A decoded range reference. Corners are the min/max of the endpoints per axis;
an axis neither endpoint mentions (a whole-row or whole-column range) is
`undefined` on every corner and simply absent from `dimensions`.

```ts
interface RangeAddress {
  readonly top: number | undefined;
  readonly left: number | undefined;
  readonly bottom: number | undefined;
  readonly right: number | undefined;
  /** The originating sheet, present only when the reference carried one. */
  readonly sheetName?: string;
  readonly tl: CellAddress;
  readonly br: CellAddress;
  /** Canonical `tl:br` form: `"A1:B2"`, `"1:1"` (rows), `"A:A"` (columns). */
  readonly dimensions: string;
}
```
