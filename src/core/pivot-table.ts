// A pivot table authored from a model.
//
// A pivot summarises a source range: its distinct field values become row/column axes and a value
// field is aggregated across them. OOXML splits that into three parts: a `pivotCacheDefinition`
// (the field catalogue), a `pivotCacheRecords` (a copy of the source rows, with axis-field cells
// swapped for indices into the catalogue), and a `pivotTableDefinition` (the layout on the
// destination sheet). This module owns the *semantic* computation of all three; the OOXML rendering
// lives in `io/xlsx/pivot.ts`.
//
// The source data is captured when the pivot is added: the model reads the source sheet's cells
// once, here, so the pivot is a stable snapshot independent of later edits to the source's values. The
// source *range* is the exception: a splice of the source sheet moves it, as it moves any reference.

import {AuthoringError, InternalError, quoted} from '../errors.ts';
import {tokenSet} from '../token-set.ts';
import {encodeRect} from './address.ts';
import {type SheetSplice, spliceFormula} from './formula-references.ts';
import {INTERNAL} from './internal.ts';
import {
  type CellValue,
  isErrorValue,
  isFormulaValue,
  isHyperlinkValue,
  isRichTextValue,
  isSharedFormulaValue,
  REF_ERROR,
  richTextToPlain,
} from './value.ts';
import type {Worksheet} from './worksheet.ts';

/** The aggregation a pivot's value field applies. These are OOXML's `ST_DataConsolidateFunction`
 * names verbatim, so a metric doubles as its `<dataField subtotal="…">` value. Excel performs the
 * aggregation itself on refresh; the writer only records which function to apply. */
export type PivotMetric =
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

const PIVOT_METRICS: Record<PivotMetric, true> = {
  sum: true,
  count: true,
  countNums: true,
  average: true,
  max: true,
  min: true,
  product: true,
  stdDev: true,
  stdDevp: true,
  var: true,
  varp: true,
};

/** Narrow a raw `subtotal` attribute (or any string) to a known {@link PivotMetric}. */
const isPivotMetric = tokenSet<PivotMetric>(PIVOT_METRICS);

/** Map an OOXML `<dataField subtotal="…">` value back to its metric. The attribute is absent for
 * `sum` (Excel's implicit default), so `undefined` reads as `sum`; an unrecognised value also reads
 * as `sum` rather than throwing, because reconstructing an existing file is a lenient operation.
 * The strict rejection of unknown metrics belongs on the authoring path, not the read path. */
export function pivotMetricFromSubtotal(subtotal: string | undefined): PivotMetric {
  if (subtotal === undefined) return 'sum';
  return isPivotMetric(subtotal) ? subtotal : 'sum';
}

/** One field in a loaded pivot's cache catalogue, in declared order; the pivot refers to it by index. */
export interface ParsedPivotField {
  readonly name: string;
}

/** The kind of data a pivot cache draws from, mirroring OOXML's `ST_SourceType`. Only `worksheet`
 * carries a {@link ParsedPivotSource.sheet}/{@link ParsedPivotSource.ref}; every other kind draws from
 * data the reader does not model (an external connection, a range consolidation, or a scenario), and
 * `unknown` covers a `type` the file declares that is none of these. */
export type PivotSourceKind = 'worksheet' | 'external' | 'consolidation' | 'scenario' | 'unknown';

// Keyed by the union minus `unknown`: that is this library's word for a token it did not recognise
// and is never one a file declares, so admitting it here would let the very token the guard exists
// to catch through.
/** Narrow a raw `<cacheSource type>` token to a {@link PivotSourceKind} a file may declare. */
export const isDeclarablePivotSourceKind = tokenSet<Exclude<PivotSourceKind, 'unknown'>>({
  worksheet: true,
  external: true,
  consolidation: true,
  scenario: true,
});

/** Where a pivot cache draws its rows from. {@link kind} names the source type; {@link sheet} and
 * {@link ref} locate the range only when it is `worksheet` and are empty strings otherwise, so a
 * consumer can tell a genuinely non-worksheet source apart from a worksheet source that failed to
 * parse (the former reports its {@link kind}, the latter stays `worksheet` with empty coordinates). */
export interface ParsedPivotSource {
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

/** The semantic model reconstructed from a loaded pivot's `pivotTableDefinition` and its
 * `pivotCacheDefinition` (see `io/xlsx/read-pivot.ts`). Field roles are indices into {@link fields};
 * {@link metric} is the aggregation the value field applies. This mirrors the authoring model's shape
 * without requiring the source sheet it was built from, so a pivot loaded from a package is
 * inspectable data rather than an opaque preserved blob. It is a read-only view: the writer emits a
 * loaded pivot from its preserved parts, not from this model, so exposing it never double-emits. A
 * splice of the source sheet moves the source range in both, by {@link splicePivotSource}. */
export interface ParsedPivotTable {
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

/** How a pivot table is authored: a source sheet and the header names that drive each axis.
 * `rows`/`columns`/`values` name columns by their header text in the source's first row, ignoring
 * case, as the headers themselves must be unique ignoring case. */
export interface PivotTableOptions {
  readonly source: Worksheet;
  readonly rows: readonly string[];
  readonly columns: readonly string[];
  readonly values: readonly string[];
  readonly metric?: PivotMetric;
}

/** One distinct value in a cache field's shared-items catalogue, or an inline record cell. A
 * `blank` is a missing source value, serialised as `<m/>` rather than an empty string. */
export type PivotItem =
  | {readonly kind: 'string'; readonly value: string}
  | {readonly kind: 'number'; readonly value: number}
  | {readonly kind: 'blank'};

/** The range and integrality of the numbers a cache field holds, which Excel records beside them. */
export interface PivotNumericSummary {
  readonly allInteger: boolean;
  readonly min: number;
  readonly max: number;
}

/** One field of the pivot cache. An axis field (row or column) carries a `sharedItems` catalogue its
 * records reference by index; any other field stores its values inline in the records. Either way the
 * cache describes which kinds of value the field holds, because Excel reads the catalogue against that
 * description: a catalogue of numbers that does not say it holds numbers opens with the repair prompt. */
export interface PivotCacheField {
  readonly name: string;
  readonly sharedItems: readonly PivotItem[] | null;
  /** Whether any value is a string. */
  readonly containsString: boolean;
  /** Whether any value is missing. */
  readonly containsBlank: boolean;
  /** The field's numbers, summarised, or `null` when it holds none. */
  readonly numeric: PivotNumericSummary | null;
}

/** One cell of a cache record: an index into a shared-items catalogue, or an inline value. */
export type PivotRecordCell = {readonly kind: 'index'; readonly index: number} | PivotItem;

const BLANK: PivotItem = {kind: 'blank'};

/**
 * A pivot table built over a source sheet's data. Construction reads the source once and computes
 * the full cache (fields + records) and the axis-field wiring the renderer needs; nothing here
 * touches XML.
 *
 * Supported shape: at least one row field and one column field, each source field on at most one
 * axis and at most once, over a header row whose names are unique ignoring case; and exactly one value
 * field, aggregated by any {@link PivotMetric} (`sum` by default), which may also be an axis field. An
 * unsupported request throws at authoring time rather than emitting a corrupt file.
 */
export class PivotTable {
  readonly metric: PivotMetric;
  readonly sourceSheetName: string;
  #sourceRef: string;
  readonly cacheFields: readonly PivotCacheField[];
  readonly records: readonly (readonly PivotRecordCell[])[];
  /** Indices into {@link cacheFields} of the row-axis, column-axis, and value fields. */
  readonly rowFields: readonly number[];
  readonly columnFields: readonly number[];
  readonly valueField: number;

  constructor(options: PivotTableOptions) {
    const metric = options.metric ?? 'sum';
    // `Object.hasOwn` rather than `isPivotMetric`: the guard would narrow `metric` to `never` in this
    // branch, and the message it throws needs the offending value. The check earns its place against
    // untyped callers, for whom the declared `PivotMetric` is not a check at all.
    if (!Object.hasOwn(PIVOT_METRICS, metric)) {
      throw new AuthoringError(
        `unsupported pivot metric ${quoted(metric)}: expected one of ${Object.keys(PIVOT_METRICS).join(', ')}`,
      );
    }
    this.metric = metric;

    const source = options.source;
    const columnCount = source.columnCount;
    const lastRow = source.rowCount;
    if (columnCount < 1 || lastRow < 2) {
      throw new AuthoringError('a pivot source needs a header row and at least one data row');
    }

    // The source is read positionally, never through `getCell`. `getCell` materialises, so reading a
    // rectangle bounded by the used extent created a cell at every position in it: a source holding 22
    // cells plus one lone value far down a column grew to 150,000 cells on the source sheet, in 189 ms,
    // and they stayed there. That is exactly the cost cliff `core/range.ts` refuses to permit for a
    // range, and a pivot has no more right to it. A position with no cell is blank, which `scalarOf`
    // already answers for `null`.
    const internals = source[INTERNAL];
    const valueAt = (row: number, col: number): CellValue =>
      internals.peekCell(row, col)?.value ?? null;

    const fields = discoverFields(valueAt, columnCount);
    const roles = resolveRoles(options, fields);
    this.rowFields = roles.rowFields;
    this.columnFields = roles.columnFields;
    this.valueField = roles.valueField;
    const axisFields = new Set<number>([...roles.rowFields, ...roles.columnFields]);

    // `discoverFields` guarantees at least one, so the span runs from its first to its last column.
    // An absent one here would be this class's own bug, not the caller's, which is what separates
    // `InternalError` from the `AuthoringError`s above: a caller cannot produce this state.
    const firstField = fields[0];
    const lastField = fields.at(-1);
    if (firstField === undefined || lastField === undefined) {
      throw new InternalError(
        'pivot field list is empty after discoverFields guaranteed it is not',
      );
    }
    this.sourceSheetName = source.name;
    this.#sourceRef = encodeRect({
      top: 1,
      left: firstField.col,
      bottom: lastRow,
      right: lastField.col,
    });

    // Read the source body once, field by field, so the same scan feeds both the shared-items
    // catalogues and the records that reference them.
    const dataRowCount = lastRow - 1;
    const columnScalars = fields.map((field) => {
      const scalars: PivotItem[] = [];
      for (let row = 2; row <= lastRow; row++) scalars.push(scalarOf(valueAt(row, field.col)));
      return scalars;
    });

    const catalogues: (Map<string, number> | null)[] = fields.map(() => null);
    this.cacheFields = fields.map((field, fieldIndex) => {
      const scalars = columnScalars[fieldIndex] ?? missingColumn(fieldIndex);
      const kinds = {
        containsString: scalars.some((scalar) => scalar.kind === 'string'),
        containsBlank: scalars.some((scalar) => scalar.kind === 'blank'),
        numeric: numericSummary(scalars),
      };
      if (axisFields.has(fieldIndex)) {
        const items: PivotItem[] = [];
        const catalogue = new Map<string, number>();
        for (const scalar of scalars) {
          const key = itemKey(scalar);
          if (!catalogue.has(key)) {
            catalogue.set(key, items.length);
            items.push(scalar);
          }
        }
        catalogues[fieldIndex] = catalogue;
        return {name: field.name, sharedItems: items, ...kinds};
      }
      return {name: field.name, sharedItems: null, ...kinds};
    });

    // Both lookups are resolved once per field rather than once per cell. The column's scalars were
    // being fetched (and its bounds re-asserted) inside the per-row-per-field loop, which is rows x
    // fields repetitions of an answer that does not vary with the row.
    const byField = fields.map((_field, fieldIndex) => ({
      scalars: columnScalars[fieldIndex] ?? missingColumn(fieldIndex),
      catalogue: catalogues[fieldIndex],
    }));
    const records: (readonly PivotRecordCell[])[] = [];
    for (let row = 0; row < dataRowCount; row++) {
      records.push(
        byField.map(({scalars, catalogue}, fieldIndex): PivotRecordCell => {
          const scalar = scalars[row];
          if (scalar === undefined) {
            throw new InternalError(
              `pivot record row ${row} is out of range for field ${fieldIndex}: every column was ` +
                'scanned for the same dataRowCount above, so this index is always in range',
            );
          }
          if (!catalogue) return scalar;
          // Every scalar was catalogued in the shared-items pass above, so this always hits.
          const index = catalogue.get(itemKey(scalar));
          if (index === undefined)
            throw new InternalError('pivot record references an uncatalogued item');
          return {kind: 'index', index};
        }),
      );
    }
    this.records = records;
  }

  /**
   * The `A1:C4` source range: the header row through the last data row, across the field columns.
   *
   * It moves with a row or column splice of the source sheet, as Excel moves a pivot's source: an
   * insert inside it grows it and a delete shrinks it, while a delete that takes the whole range leaves
   * it as it was. The cache captured at construction does not change; Excel rebuilds it from this range
   * when it opens the file.
   */
  get sourceRef(): string {
    return this.#sourceRef;
  }

  /** What a splice does to the pivot's source; see `core/internal.ts`. */
  readonly [INTERNAL]: PivotTableInternals = {
    spliceSource: (edit) => {
      this.#sourceRef = splicePivotSource(this.sourceSheetName, this.#sourceRef, edit);
    },
  };

  /** The value field's header name, used to label the aggregated data column ("Sum of Amount"). */
  get valueFieldName(): string {
    const field = this.cacheFields[this.valueField];
    if (field === undefined) {
      throw new InternalError(
        `pivot valueField index ${this.valueField} is out of range: resolve() validated it against ` +
          'the same fields array cacheFields was built from',
      );
    }
    return field.name;
  }
}

/** One column of the source that a header names, and where it sits. */
interface SourceField {
  readonly name: string;
  readonly col: number;
}

// Every non-blank header cell in row 1 defines a field, in ascending column order. A cache names each
// field once, and Excel compares those names ignoring case: two headers that differ only in case make
// it repair the package just as an exact repeat does. Refused rather than renamed the way a table's
// columns are, because the caller names pivot fields by header text, and a silent rename would leave
// the second column unreachable by the name the caller can see in the sheet.
function discoverFields(
  valueAt: (row: number, col: number) => CellValue,
  columnCount: number,
): SourceField[] {
  const fields: SourceField[] = [];
  const seen = new Map<string, string>();
  for (let col = 1; col <= columnCount; col++) {
    const name = textOf(scalarOf(valueAt(1, col)));
    if (name === '') continue;
    const clash = seen.get(name.toLowerCase());
    if (clash !== undefined) {
      throw new AuthoringError(
        `the pivot source header ${quoted(name)} repeats ${quoted(clash)}: pivot field names must ` +
          'be unique, ignoring case',
      );
    }
    seen.set(name.toLowerCase(), name);
    fields.push({name, col});
  }
  if (fields.length === 0) throw new AuthoringError('the pivot source header row is empty');
  return fields;
}

/** Which discovered field plays each role, as indices into the field list. */
interface FieldRoles {
  readonly rowFields: readonly number[];
  readonly columnFields: readonly number[];
  readonly valueField: number;
}

// Bind the caller's field *names* to positions in the source. Every refusal here is the caller's
// mistake, so every one is an AuthoringError naming the field and the role it was asked to play.
//
// A field sits on at most one axis, once: `<pivotField axis>` holds a single axis, so a field listed
// under both `<rowFields>` and `<colFields>`, or twice under one, contradicts its own declaration and
// Excel repairs the package. The value field is the exception that is not one: aggregating a field that
// is also an axis ("Count of Name" by Name) is an ordinary pivot, and the writer flags it as both.
//
// A name binds its header ignoring case, which is the same comparison `discoverFields` holds the
// headers unique under, so it can never bind two. Matched exactly, `name` over a `Name` header was
// refused as not a header of a sheet that visibly has one.
function resolveRoles(options: PivotTableOptions, fields: readonly SourceField[]): FieldRoles {
  const resolve = (role: string, name: string): number => {
    const key = name.toLowerCase();
    const index = fields.findIndex((field) => field.name.toLowerCase() === key);
    if (index < 0) {
      throw new AuthoringError(
        `pivot ${role} field ${quoted(name)} is not a column header in the source sheet`,
      );
    }
    return index;
  };
  if (options.rows.length === 0)
    throw new AuthoringError('a pivot table needs at least one row field');
  if (options.columns.length === 0)
    throw new AuthoringError('a pivot table needs at least one column field');
  const [valueName, ...extraValues] = options.values;
  if (valueName === undefined || extraValues.length > 0) {
    throw new AuthoringError('a pivot table needs exactly one value field');
  }
  const axisRoles = new Map<number, string>();
  const resolveAxis = (role: string, names: readonly string[]): number[] =>
    names.map((name) => {
      const index = resolve(role, name);
      const earlier = axisRoles.get(index);
      if (earlier === role) {
        throw new AuthoringError(`pivot ${role} field ${quoted(name)} is named more than once`);
      }
      if (earlier !== undefined) {
        throw new AuthoringError(
          `pivot field ${quoted(name)} is named as both a ${earlier} and a ${role} field: a field ` +
            'sits on one axis',
        );
      }
      axisRoles.set(index, role);
      return index;
    });
  return {
    rowFields: resolveAxis('row', options.rows),
    columnFields: resolveAxis('column', options.columns),
    valueField: resolve('value', valueName),
  };
}

// columnScalars[fieldIndex] is always present: it was built by mapping the same `fields` array every
// index is drawn from. Stated once, as the fallback of the two `??` that read it.
function missingColumn(fieldIndex: number): never {
  throw new InternalError(
    `pivot field index ${fieldIndex} is out of range for columnScalars: it was built from the same fields array`,
  );
}

/** A stable dedup key for a shared item: kind-tagged so the number `1` and the string `"1"` differ. */
function itemKey(item: PivotItem): string {
  switch (item.kind) {
    case 'string':
      return `s:${item.value}`;
    case 'number':
      return `n:${item.value}`;
    case 'blank':
      return 'b';
  }
}

/** The summary of a field's numbers, or null when it holds none. Strings and blanks beside them do not
 * change it: Excel records the range of a mixed field's numbers as it does a pure one's. */
function numericSummary(scalars: readonly PivotItem[]): PivotNumericSummary | null {
  let min = Infinity;
  let max = -Infinity;
  let allInteger = true;
  let sawNumber = false;
  for (const scalar of scalars) {
    if (scalar.kind !== 'number') continue;
    sawNumber = true;
    if (scalar.value < min) min = scalar.value;
    if (scalar.value > max) max = scalar.value;
    if (!Number.isInteger(scalar.value)) allInteger = false;
  }
  return sawNumber ? {allInteger, min, max} : null;
}

/** The string form of a shared item, used for header names (a blank header contributes no field). */
function textOf(item: PivotItem): string {
  return item.kind === 'blank' ? '' : String(item.value);
}

/**
 * Reduce any cell value to the scalar a pivot cache can hold: a number, a string, or a blank. Only
 * finite numbers stay numeric (a NaN would corrupt the cache); every other kind is flattened to its
 * displayed text so hostile or exotic source content can never throw or leak an object into the XML.
 *
 * Deliberately not `cellValueToText`, close as the two look: this classifies rather than renders
 * (a number must reach the cache *as a number*), and its no-throw promise is the opposite of that
 * function's, which rejects a value outside the union rather than quietly caching a blank.
 */
function scalarOf(value: CellValue): PivotItem {
  if (value === null) return BLANK;
  switch (typeof value) {
    case 'number':
      return Number.isFinite(value) ? {kind: 'number', value} : BLANK;
    case 'string':
      return {kind: 'string', value};
    case 'boolean':
      return {kind: 'string', value: value ? 'TRUE' : 'FALSE'};
    default:
      break;
  }
  if (value instanceof Date) return {kind: 'string', value: value.toISOString()};
  if (isRichTextValue(value)) return {kind: 'string', value: richTextToPlain(value)};
  if (isHyperlinkValue(value)) {
    return {
      kind: 'string',
      value: typeof value.text === 'string' ? value.text : richTextToPlain(value.text),
    };
  }
  if (isErrorValue(value)) return {kind: 'string', value: value.error};
  if (isFormulaValue(value) || isSharedFormulaValue(value)) {
    return value.result === undefined ? BLANK : scalarOf(value.result);
  }
  return BLANK;
}

/**
 * A pivot's worksheet source range on `sheet`, moved through a splice as Excel moves it: grown by an
 * insert inside it, shrunk by a delete, and left as it was by a delete that takes every row or every
 * column of it, which Excel does not turn into `#REF!`. A splice of another sheet leaves it too. The one
 * rule for an authored pivot, a loaded pivot's view, and the preserved cache the writer edits.
 */
export function splicePivotSource(sheet: string, ref: string, edit: SheetSplice): string {
  if (sheet.toLowerCase() !== edit.sheet.toLowerCase()) return ref;
  const moved = spliceFormula(ref, sheet, edit);
  return moved === REF_ERROR ? ref : moved;
}

/** What the library's own machinery may do to a {@link PivotTable}; reached as `pivot[INTERNAL]`. */
export interface PivotTableInternals {
  /** Move the source range through a row or column splice, when the splice is of the source sheet. */
  spliceSource(edit: SheetSplice): void;
}
