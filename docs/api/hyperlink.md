# Hyperlink

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `Hyperlink`

<sub>interface</sub>

A hyperlink and the cells it covers.

```ts
interface Hyperlink {
  /** The cells the link covers: one cell (`'B2'`) or a rectangle of cells (`'D1:H1'`). */
  readonly ref: string;
  /**
   * Where the link goes: a URL or a path, or a `#`-prefixed place in this workbook such as
   * `#Summary!A1` or `#TaxRate`.
   */
  readonly target: string;
  /** The text shown when the pointer rests on the link. */
  readonly tooltip?: string;
}
```
