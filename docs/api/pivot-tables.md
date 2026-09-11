# Pivot tables

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `isDeclarablePivotSourceKind`

<sub>const</sub>

Narrow a raw `<cacheSource type>` token to a [`PivotSourceKind`](./pivot-tables.md#pivotsourcekind) a file may declare.

```ts
const isDeclarablePivotSourceKind: (value: string) => value is "consolidation" | "external" | "scenario" | "worksheet"
```

---

### `ParsedPivotField`

<sub>interface</sub>

One field in a loaded pivot's cache catalogue, in declared order; the pivot refers to it by index.

```ts
interface ParsedPivotField {
  readonly name: string;
}
```

---

### `ParsedPivotSource`

<sub>interface</sub>

Where a pivot cache draws its rows from. `kind` names the source type; `sheet` and
`ref` locate the range only when it is `worksheet` and are empty strings otherwise, so a
consumer can tell a genuinely non-worksheet source apart from a worksheet source that failed to
parse (the former reports its `kind`, the latter stays `worksheet` with empty coordinates).

```ts
interface ParsedPivotSource {
  readonly kind: PivotSourceKind;
  readonly sheet: string;
  readonly ref: string;
  /**
   * Whether {@link sheet} and {@link ref} name a range in another workbook, which the cache reaches
   * through a relationship. A row or column splice of this workbook moves the source of a pivot drawing
   * from one of its own sheets, as Excel does, and never one of these.
   */
  readonly inAnotherWorkbook: boolean;
}
```

---

### `ParsedPivotTable`

<sub>interface</sub>

The semantic model reconstructed from a loaded pivot's `pivotTableDefinition` and its
`pivotCacheDefinition` (see `io/xlsx/read-pivot.ts`). Field roles are indices into `fields`;
`metric` is the aggregation the value field applies. This mirrors the authoring model's shape
without requiring the source sheet it was built from, so a pivot loaded from a package is
inspectable data rather than an opaque preserved blob. It is a read-only view: the writer emits a
loaded pivot from its preserved parts, not from this model, so exposing it never double-emits. A
splice of the source sheet moves the source range in both, by `splicePivotSource`.

```ts
interface ParsedPivotTable {
  readonly name: string;
  readonly cacheId: string;
  readonly source: ParsedPivotSource;
  readonly fields: readonly ParsedPivotField[];
  readonly rowFields: readonly number[];
  readonly columnFields: readonly number[];
  /** Index into {@link fields} of the aggregated field, or -1 when no `<dataField>` was declared. */
  readonly valueField: number;
  readonly valueFieldName: string;
  /** The `<dataField>`'s own caption ("Average of Amount"), which Excel shows on the data column. */
  readonly valueCaption: string;
  readonly metric: PivotMetric;
}
```

---

### `PivotCacheField`

<sub>interface</sub>

One field of the pivot cache. An axis field (row or column) carries a `sharedItems` catalogue its
records reference by index; any other field stores its values inline in the records. Either way the
cache describes which kinds of value the field holds, because Excel reads the catalogue against that
description: a catalogue of numbers that does not say it holds numbers opens with the repair prompt.

```ts
interface PivotCacheField {
  readonly name: string;
  readonly sharedItems: readonly PivotItem[] | null;
  /** Whether any value is a string. */
  readonly containsString: boolean;
  /** Whether any value is missing. */
  readonly containsBlank: boolean;
  /** The field's numbers, summarised, or `null` when it holds none. */
  readonly numeric: PivotNumericSummary | null;
}
```

---

### `PivotItem`

<sub>type</sub>

One distinct value in a cache field's shared-items catalogue, or an inline record cell. A
`blank` is a missing source value, serialised as `<m/>` rather than an empty string.

```ts
type PivotItem =
  | {readonly kind: 'string'; readonly value: string}
  | {readonly kind: 'number'; readonly value: number}
  | {readonly kind: 'blank'};
```

---

### `PivotMetric`

<sub>type</sub>

The aggregation a pivot's value field applies. These are OOXML's `ST_DataConsolidateFunction`
names verbatim, so a metric doubles as its `<dataField subtotal="…">` value. Excel performs the
aggregation itself on refresh; the writer only records which function to apply.

```ts
type PivotMetric =
  | 'sum'
  | 'count'
  | 'countNums'
  | 'average'
  | 'max'
  | 'min'
  | 'product'
  | 'stdDev'
  | 'stdDevp'
  | 'var'
  | 'varp';
```

---

### `PivotNumericSummary`

<sub>interface</sub>

The range and integrality of the numbers a cache field holds, which Excel records beside them.

```ts
interface PivotNumericSummary {
  readonly allInteger: boolean;
  readonly min: number;
  readonly max: number;
}
```

---

### `PivotRecordCell`

<sub>type</sub>

One cell of a cache record: an index into a shared-items catalogue, or an inline value.

```ts
type PivotRecordCell = {readonly kind: 'index'; readonly index: number} | PivotItem;
```

---

### `PivotSourceKind`

<sub>type</sub>

The kind of data a pivot cache draws from, mirroring OOXML's `ST_SourceType`. Only `worksheet`
carries a [`ParsedPivotSource.sheet`](./pivot-tables.md#parsedpivotsource)/[`ParsedPivotSource.ref`](./pivot-tables.md#parsedpivotsource); every other kind draws from
data the reader does not model (an external connection, a range consolidation, or a scenario), and
`unknown` covers a `type` the file declares that is none of these.

```ts
type PivotSourceKind = 'worksheet' | 'external' | 'consolidation' | 'scenario' | 'unknown';
```

---

### `PivotTable`

<sub>class</sub>

A pivot table built over a source sheet's data. Construction reads the source once and computes
the full cache (fields + records) and the axis-field wiring the renderer needs; nothing here
touches XML.

Supported shape: at least one row field and one column field, each source field on at most one
axis and at most once, over a header row whose names are unique ignoring case; and exactly one value
field, aggregated by any [`PivotMetric`](./pivot-tables.md#pivotmetric) (`sum` by default), which may also be an axis field. An
unsupported request throws at authoring time rather than emitting a corrupt file.

```ts
class PivotTable {
  readonly metric: PivotMetric;
  readonly sourceSheetName: string;
  readonly cacheFields: readonly PivotCacheField[];
  readonly records: readonly (readonly PivotRecordCell[])[];
  readonly rowFields: readonly number[];
  readonly columnFields: readonly number[];
  readonly valueField: number;
  get sourceRef(): string;
  get valueFieldName(): string;
}
```

**Members**

#### `PivotTable.rowFields`

```ts
readonly rowFields: readonly number[];
```

Indices into `cacheFields` of the row-axis, column-axis, and value fields.

#### `PivotTable.sourceRef`

```ts
get sourceRef(): string;
```

The `A1:C4` source range: the header row through the last data row, across the field columns.

It moves with a row or column splice of the source sheet, as Excel moves a pivot's source: an
insert inside it grows it and a delete shrinks it, while a delete that takes the whole range leaves
it as it was. The cache captured at construction does not change; Excel rebuilds it from this range
when it opens the file.

#### `PivotTable.valueFieldName`

```ts
get valueFieldName(): string;
```

The value field's header name, used to label the aggregated data column ("Sum of Amount").

---

### `PivotTableOptions`

<sub>interface</sub>

How a pivot table is authored: a source sheet and the header names that drive each axis.
`rows`/`columns`/`values` name columns by their header text in the source's first row, ignoring
case, as the headers themselves must be unique ignoring case.

```ts
interface PivotTableOptions {
  readonly source: Worksheet;
  readonly rows: readonly string[];
  readonly columns: readonly string[];
  readonly values: readonly string[];
  readonly metric?: PivotMetric;
}
```
