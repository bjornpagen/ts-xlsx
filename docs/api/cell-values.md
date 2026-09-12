# Cell values

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `CellValue`

<sub>type</sub>

Everything a cell's value can be. `null` is the empty cell.

```ts
type CellValue =
  | null
  | number
  | string
  | boolean
  | Date
  | ErrorValue
  | FormulaValue
  | SharedFormulaValue
  | DataTableFormulaValue
  | RichTextValue;
```

---

### `cellValueToText`

<sub>function</sub>

The plain text of any cell value, total over [`CellValue`](./cell-values.md#cellvalue), so a caller reading a sheet
whose cells it did not write never has to switch on the union itself.

This is the *value's* text, not the cell's *display* text: a number renders as JavaScript
renders it, with no number format applied (`0.1 + 0.2` is `"0.30000000000000004"`, a currency
cell has no currency sign), because the format lives on the style and this function is given
only the value. What each kind yields:

- the empty cell (`null`) and an invalid `Date` → `""`, the two ways a cell has no text
- a boolean → `"TRUE"` / `"FALSE"`, Excel's own literals rather than JavaScript's
- a `Date` → a full ISO-8601 timestamp
- an error → its literal, e.g. `"#REF!"`, the same string the grid shows
- rich text → every run concatenated ([`richTextToPlain`](./cell-values.md#richtexttoplain))
- any of the three formula kinds → the text of the *cached result*, and `""` when the cell
  carries no cached result: the formula source is not text the sheet ever displayed

```ts
function cellValueToText(value: CellValue): string;
```

---

### `coerceCellValue`

<sub>function</sub>

Normalise a raw assignment into a stored [`CellValue`](./cell-values.md#cellvalue). `undefined` becomes the
empty cell (`null`); every other kind is validated by [`detectValueType`](./cell-values.md#detectvaluetype). The
model never rewrites one value *kind* into another (a numeric-looking string stays a
string). The single exception is formula text, which is canonicalised to the OOXML
stored form (no leading `=`) so round-trips are idempotent regardless of how the
caller supplied it.

```ts
function coerceCellValue(value: CellValue | undefined): CellValue;
```

**Throws:** `TypeError` if the value is not a recognised cell-value shape.

---

### `DataTableFormulaValue`

<sub>interface</sub>

A cell computed by a What-If-Analysis data table (`<f t="dataTable">`), the OOXML formula kind that
fills a range by re-evaluating a model against a grid of substituted input cells. The library does
not evaluate it; it preserves the declaration so a read-modify-write cycle re-emits it verbatim
rather than silently dropping the data-table kind.

```ts
interface DataTableFormulaValue {
  readonly shareType: 'dataTable';
  /** The range the data table fills, e.g. `'B2:B5'`. */
  readonly ref: string;
  /** Whether the table substitutes two inputs (a 2-D data table) rather than one. */
  readonly dataTable2D?: boolean;
  /** For a 1-D table, whether the input runs along the row rather than down the column. */
  readonly dataTableRow?: boolean;
  /** The first (row) input-cell reference. */
  readonly r1?: string;
  /** The second (column) input-cell reference, present for a 2-D table. */
  readonly r2?: string;
  /**
   * Whether the cell {@link r1} named has been deleted. Excel keeps the reference as it was written and
   * sets this flag, and the table then shows `#REF!`; a row or column delete that takes the input cell
   * does the same here.
   */
  readonly r1Deleted?: boolean;
  /** Whether the cell {@link r2} named has been deleted, as {@link r1Deleted} is for {@link r1}. */
  readonly r2Deleted?: boolean;
  readonly result?: FormulaResult;
}
```

---

### `detectValueType`

<sub>function</sub>

Classify a value into its observable [`ValueType`](./cell-values.md#valuetype). This is total over
[`CellValue`](./cell-values.md#cellvalue): every legal value has exactly one type. A `Date` is a date even
when its time is `NaN` (an invalid date is still a date-typed cell); serialization,
not the model, decides what to do with it.

```ts
function detectValueType(value: CellValue): ValueType;
```

---

### `ERROR_CODES`

<sub>const</sub>

The Excel error literals a cell (or formula result) can carry: the spellings a typed cell's `<v>`
holds and Excel reads back as that error.

Not every error Excel displays is one. `#SPILL!`, `#CALC!`, `#FIELD!`, `#BLOCKED!`, `#CONNECT!`,
`#UNKNOWN!` and `#PYTHON!` are stored as `#VALUE!` with a rich value naming the real error beside
it, and a cell carrying one of them literally makes Excel offer to repair the package, so none of them
is here. A file's such cell reads as the `#VALUE!` its `<v>` states. `#BUSY!` and `#GETTING_DATA` are
literals like the classic seven (Excel 16.0 build 20326).

```ts
const ERROR_CODES: readonly ["#N/A", "#REF!", "#NAME?", "#DIV/0!", "#NULL!", "#VALUE!", "#NUM!", "#GETTING_DATA", "#BUSY!"]
```

---

### `ErrorCode`

<sub>type</sub>

```ts
type ErrorCode = (typeof ERROR_CODES)[number];
```

---

### `ErrorValue`

<sub>interface</sub>

An in-cell error, e.g. `{error: '#REF!'}`.

```ts
interface ErrorValue {
  readonly error: ErrorCode;
}
```

---

### `FormulaResult`

<sub>type</sub>

The cached result a formula carries: any scalar, a date, or an error.

```ts
type FormulaResult = number | string | boolean | Date | ErrorValue;
```

---

### `FormulaValue`

<sub>interface</sub>

A cell whose value is computed by its own formula.

```ts
interface FormulaValue {
  readonly formula: string;
  readonly result?: FormulaResult;
}
```

---

### `isDataTableFormulaValue`

<sub>function</sub>

Whether a value is a What-If-Analysis data-table formula ([`DataTableFormulaValue`](./cell-values.md#datatableformulavalue)).

```ts
function isDataTableFormulaValue(value: CellValue): value is DataTableFormulaValue;
```

---

### `isErrorCode`

<sub>function</sub>

Whether a string is one of Excel's canonical error literals.

```ts
function isErrorCode(text: string): text is ErrorCode;
```

---

### `isErrorValue`

<sub>function</sub>

Whether a value is an in-cell error ([`ErrorValue`](./cell-values.md#errorvalue)). The narrowing counterpart of
`detectValueType(value) === ValueType.Error`: use this one when the branch goes on to read
`.error`, and [`detectValueType`](./cell-values.md#detectvaluetype) when it dispatches over all eight kinds at once.

```ts
function isErrorValue(value: CellValue): value is ErrorValue;
```

---

### `isFormulaValue`

<sub>function</sub>

Whether a value is a cell's own formula ([`FormulaValue`](./cell-values.md#formulavalue)): a master, or a formula
belonging to no shared group. A shared-formula clone is **not** one of these; see
[`isSharedFormulaValue`](./cell-values.md#issharedformulavalue). Both report as `ValueType.Formula`, so a caller that means "any
formula-shaped cell" wants [`detectValueType`](./cell-values.md#detectvaluetype), not this.

```ts
function isFormulaValue(value: CellValue): value is FormulaValue;
```

---

### `isRichTextValue`

<sub>function</sub>

Whether a value is composed of formatted runs ([`RichTextValue`](./cell-values.md#richtextvalue)). This is the test to
make before [`richTextToPlain`](./cell-values.md#richtexttoplain), which accepts nothing else.

```ts
function isRichTextValue(value: CellValue): value is RichTextValue;
```

---

### `isSharedFormulaValue`

<sub>function</sub>

Whether a value is a clone participating in a shared formula ([`SharedFormulaValue`](./cell-values.md#sharedformulavalue)).

```ts
function isSharedFormulaValue(value: CellValue): value is SharedFormulaValue;
```

---

### `RichTextRun`

<sub>interface</sub>

One formatted run of a rich-text value.

```ts
interface RichTextRun {
  readonly text: string;
  readonly font?: Font;
}
```

---

### `richTextToPlain`

<sub>function</sub>

Flatten a rich-text value to its plain text by concatenating every run's text in order. This is the
text a consumer that cannot render per-run formatting (a CSV field, a pivot cache entry) sees, and
the string a rich cell reads as when its formatting is discarded.

```ts
function richTextToPlain(value: RichTextValue): string;
```

---

### `RichTextValue`

<sub>interface</sub>

A value composed of independently-formatted text runs.

```ts
interface RichTextValue {
  readonly richText: readonly RichTextRun[];
}
```

---

### `SharedFormulaValue`

<sub>interface</sub>

A cell that participates in a shared formula: a clone of a master formula cell filled across a
range. `sharedFormula` is the master cell's address (e.g. `'B1'`); the master itself is a plain
[`FormulaValue`](./cell-values.md#formulavalue). On read, the clone's own formula is the master's translated to the clone's
position and `result` is the clone's cached value; on write, the clones of a master collapse into
OOXML's shared-formula grouping.

```ts
interface SharedFormulaValue {
  readonly sharedFormula: string;
  /** The master's formula translated to this cell's position. Filled in on read; a clone assigned by
   * a caller carries only `sharedFormula`, and the writer recovers the formula from the master. */
  readonly formula?: string;
  readonly result?: FormulaResult;
}
```

---

### `ValueType`

<sub>const</sub>

The observable kind of a cell's value. Both formula shapes report as `Formula`.

```ts
const ValueType: { readonly Null: 'null'; readonly Number: 'number'; readonly String: 'string'; readonly Boolean: 'boolean'; readonly Date: 'date'; readonly Error: 'error'; readonly Formula: 'formula'; readonly RichText: 'richText'; }
```
