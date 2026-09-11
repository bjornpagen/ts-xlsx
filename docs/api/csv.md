# CSV

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `CsvEncoding`

<sub>type</sub>

A byte encoding [`writeCsv`](./csv.md#writecsv) can produce, spelled the way Node's `Buffer` spells it and
meaning the same bytes.

Named here rather than taken from Node's `BufferEncoding`, which is what this option used to be:
that type is part of `@types/node`, so it made a browser consumer's public API surface depend on
Node's types, and it offered `base64` and `hex` as if they were output encodings for a text
format. The list is what a CSV consumer actually asks for.

```ts
type CsvEncoding =
  | 'ascii'
  | 'latin1'
  | 'ucs-2'
  | 'ucs2'
  | 'utf-8'
  | 'utf-16le'
  | 'utf16le'
  | 'utf8';
```

---

### `CsvReadOptions`

<sub>interface</sub>

```ts
interface CsvReadOptions {
  /** Field separator; defaults to a comma. A single character. */
  readonly delimiter?: string;
  /** Treat the first line as a header and drop it, leaving only data rows. */
  readonly headers?: boolean;
  /** Per-field transform replacing the default type coercion; receives the raw string and its
   * 0-based column index. */
  readonly map?: (value: string, index: number) => CellValue;
  /** Name for the single worksheet produced; defaults to `"Sheet1"`. */
  readonly sheetName?: string;
}
```

---

### `CsvWriteOptions`

<sub>interface</sub>

```ts
interface CsvWriteOptions {
  /** Which worksheet to write; defaults to the first. A name matching no sheet throws rather than
   * silently emitting an empty file. */
  readonly sheetName?: string;
  /** Field separator; defaults to a comma. A single character other than a quote, CR or LF. */
  readonly delimiter?: string;
  /** Line separator between rows; defaults to `"\n"`. Not empty, and containing neither the field
   * delimiter nor a quote. A field that contains it is quoted. */
  readonly rowDelimiter?: string;
  /**
   * An Excel number-format code (e.g. `"yyyy-mm-dd"`, `"d mmm yy hh:mm"`) for Date cells; without it
   * a Date renders as a full ISO-8601 timestamp.
   *
   * The same vocabulary as {@link Cell.numFmt}, so `writeCsv(wb, {dateFormat: cell.numFmt})` renders
   * what the cell would show. It used to be a moment.js-style token set, which is case-sensitive and
   * spells the month `MM` and the minute `mm`: passing this library's own format codes to it produced
   * `2024-45-dd` with no throw and no warning. See ADR 0041.
   */
  readonly dateFormat?: string;
  /** Render Date cells in UTC rather than the runner's local time. */
  readonly dateUTC?: boolean;
  /** Byte encoding for {@link writeCsv}; defaults to `"utf8"`. */
  readonly encoding?: CsvEncoding;
  /** Prepend a UTF-8 byte-order mark (applies only to UTF-8); defaults to `true` for UTF-8. */
  readonly bom?: boolean;
  /** Per-field transform replacing the default value rendering; receives the cell's value (`null`
   * for an unpopulated column) and its 0-based column index. Quoting (commas, quotes, newlines) is
   * still applied to the returned text. */
  readonly map?: (value: CellValue, index: number) => string;
}
```

---

### `readCsv`

<sub>function</sub>

Parse CSV text (or UTF-8 bytes) into a workbook holding a single worksheet.

```ts
function readCsv(input: string | Uint8Array, options: CsvReadOptions = {}): Workbook;
```

**Throws:** [`CsvParseError`](./csv-errors.md#csvparseerror) if a record has more fields than a worksheet has columns, or the data
records outnumber a worksheet's rows.

---

### `writeCsv`

<sub>function</sub>

The CSV bytes of one worksheet in the requested encoding, with a UTF-8 BOM by default.

```ts
function writeCsv(workbook: Workbook, options: CsvWriteOptions = {}): Uint8Array;
```

**Throws:** [`AuthoringError`](./errors.md#authoringerror) if a field holds an unpaired surrogate and the encoding is UTF-8, which
cannot represent one. The alternative is a silent U+FFFD substitution.
**Throws:** `RangeError` on a delimiter [`writeCsvText`](./csv.md#writecsvtext) refuses.

---

### `writeCsvText`

<sub>function</sub>

The logical CSV text of one worksheet: no BOM, no byte encoding. Line N holds row N, so an empty
row before the last populated one is an empty line.

```ts
function writeCsvText(workbook: Workbook, options: CsvWriteOptions = {}): string;
```

**Throws:** `RangeError` if `delimiter` or `rowDelimiter` is one the output could not be split on.
