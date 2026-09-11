// The cell value model: the second bedrock primitive after addressing.
//
// A cell holds exactly one value, and that value's *type* is observable and drives
// everything downstream (serialization, number-format application, formula results).
// The honest shape here is a discriminated union: a value is either a JS primitive
// (null / number / string / boolean / Date) or one of the structural OOXML value
// shapes (error, formula, shared formula, rich text, hyperlink). There is no
// stringly-typed sentinel and no silent coercion between kinds: a numeric-looking
// string stays a string, because the caller's chosen type is the source of truth.

import type {Font} from './style.ts';

/** The observable kind of a cell's value. Both formula shapes report as `Formula`. */
export const ValueType = {
  Null: 'null',
  Number: 'number',
  String: 'string',
  Boolean: 'boolean',
  Date: 'date',
  Error: 'error',
  Formula: 'formula',
  RichText: 'richText',
  Hyperlink: 'hyperlink',
} as const;

export type ValueType = (typeof ValueType)[keyof typeof ValueType];

/**
 * The Excel error literals a cell (or formula result) can carry: the spellings a typed cell's `<v>`
 * holds and Excel reads back as that error.
 *
 * Not every error Excel displays is one. `#SPILL!`, `#CALC!`, `#FIELD!`, `#BLOCKED!`, `#CONNECT!`,
 * `#UNKNOWN!` and `#PYTHON!` are stored as `#VALUE!` with a rich value naming the real error beside
 * it, and a cell carrying one of them literally makes Excel offer to repair the package, so none of them
 * is here. A file's such cell reads as the `#VALUE!` its `<v>` states. `#BUSY!` and `#GETTING_DATA` are
 * literals like the classic seven (Excel 16.0 build 20326).
 */
export const ERROR_CODES = [
  '#N/A',
  '#REF!',
  '#NAME?',
  '#DIV/0!',
  '#NULL!',
  '#VALUE!',
  '#NUM!',
  '#GETTING_DATA',
  '#BUSY!',
] as const;

export type ErrorCode = (typeof ERROR_CODES)[number];

/**
 * The error a spreadsheet puts in place of a reference that has nowhere left to point.
 *
 * Named here rather than in the two formula codecs that produce it. Both of them -- the A1 renderer
 * in `core/formula.ts` and the BIFF12 one in `io/xlsb/formula.ts` -- had a private literal of their
 * own, spelled the same and named differently, for a value {@link ERROR_CODES} already publishes and
 * `isErrorCode` already recognises.
 */
export const REF_ERROR: ErrorCode = '#REF!';

/** An in-cell error, e.g. `{error: '#REF!'}`. */
export interface ErrorValue {
  readonly error: ErrorCode;
}

/** One formatted run of a rich-text value. */
export interface RichTextRun {
  readonly text: string;
  readonly font?: Font;
}

/** A value composed of independently-formatted text runs. */
export interface RichTextValue {
  readonly richText: readonly RichTextRun[];
}

/** A hyperlink cell: a URL plus the text (plain or rich) shown in the cell. */
export interface HyperlinkValue {
  readonly hyperlink: string;
  readonly text: string | RichTextValue;
  readonly tooltip?: string;
  /** The clickable extent (`'D1:H1'`) when the link spans a range whose top-left corner is this
   * cell. Absent for an ordinary single-cell link. The destination and label live on the top-left
   * cell; `range` records how far Excel highlights the clickable area so it survives a round-trip. */
  readonly range?: string;
}

/** The cached result a formula carries: any scalar, a date, or an error. */
export type FormulaResult = number | string | boolean | Date | ErrorValue;

/** A cell whose value is computed by its own formula. */
export interface FormulaValue {
  readonly formula: string;
  readonly result?: FormulaResult;
}

/**
 * A cell that participates in a shared formula: a clone of a master formula cell filled across a
 * range. `sharedFormula` is the master cell's address (e.g. `'B1'`); the master itself is a plain
 * {@link FormulaValue}. On read, the clone's own formula is the master's translated to the clone's
 * position and `result` is the clone's cached value; on write, the clones of a master collapse into
 * OOXML's shared-formula grouping.
 */
export interface SharedFormulaValue {
  readonly sharedFormula: string;
  /** The master's formula translated to this cell's position. Filled in on read; a clone assigned by
   * a caller carries only `sharedFormula`, and the writer recovers the formula from the master. */
  readonly formula?: string;
  readonly result?: FormulaResult;
}

/**
 * A cell computed by a What-If-Analysis data table (`<f t="dataTable">`), the OOXML formula kind that
 * fills a range by re-evaluating a model against a grid of substituted input cells. The library does
 * not evaluate it; it preserves the declaration so a read-modify-write cycle re-emits it verbatim
 * rather than silently dropping the data-table kind.
 */
export interface DataTableFormulaValue {
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
  readonly result?: FormulaResult;
}

/** Everything a cell's value can be. `null` is the empty cell. */
export type CellValue =
  | null
  | number
  | string
  | boolean
  | Date
  | ErrorValue
  | FormulaValue
  | SharedFormulaValue
  | DataTableFormulaValue
  | RichTextValue
  | HyperlinkValue;

const ERROR_SET: ReadonlySet<string> = new Set(ERROR_CODES);

// Whether a value is a non-null object carrying `key`. The object-shaped {@link CellValue} kinds are all
// discriminated by the presence of a single property, so every guard below narrows through this one test.
function hasKey<K extends string>(value: unknown, key: K): value is Record<K, unknown> & object {
  return typeof value === 'object' && value !== null && key in value;
}

/**
 * Whether a value is an in-cell error ({@link ErrorValue}). The narrowing counterpart of
 * `detectValueType(value) === ValueType.Error`: use this one when the branch goes on to read
 * `.error`, and {@link detectValueType} when it dispatches over all nine kinds at once.
 */
export function isErrorValue(value: CellValue): value is ErrorValue {
  return hasKey(value, 'error');
}

/**
 * Whether a value is a cell's own formula ({@link FormulaValue}): a master, or a formula
 * belonging to no shared group. A shared-formula clone is **not** one of these; see
 * {@link isSharedFormulaValue}. Both report as `ValueType.Formula`, so a caller that means "any
 * formula-shaped cell" wants {@link detectValueType}, not this.
 */
export function isFormulaValue(value: CellValue): value is FormulaValue {
  // A shared-formula clone resolved on read carries both its master address (`sharedFormula`) and the
  // translated `formula`; it is a SharedFormulaValue, so exclude it here to keep the two kinds distinct.
  return hasKey(value, 'formula') && !('sharedFormula' in value);
}

/** Whether a value is a clone participating in a shared formula ({@link SharedFormulaValue}). */
export function isSharedFormulaValue(value: CellValue): value is SharedFormulaValue {
  return hasKey(value, 'sharedFormula');
}

/** Whether a value is a What-If-Analysis data-table formula ({@link DataTableFormulaValue}). */
export function isDataTableFormulaValue(value: CellValue): value is DataTableFormulaValue {
  return hasKey(value, 'shareType') && value.shareType === 'dataTable';
}

/**
 * Whether a value is composed of formatted runs ({@link RichTextValue}). This is the test to
 * make before {@link richTextToPlain}, which accepts nothing else.
 */
export function isRichTextValue(value: CellValue): value is RichTextValue {
  return hasKey(value, 'richText');
}

/**
 * Whether a value is a hyperlink ({@link HyperlinkValue}). Note that its `text` is itself either
 * a string or a {@link RichTextValue}, so reading the label out means one more narrowing.
 */
export function isHyperlinkValue(value: CellValue): value is HyperlinkValue {
  return hasKey(value, 'hyperlink');
}

/**
 * Flatten a rich-text value to its plain text by concatenating every run's text in order. This is the
 * text a consumer that cannot render per-run formatting (a CSV field, a pivot cache entry) sees, and
 * the string a rich cell reads as when its formatting is discarded.
 */
export function richTextToPlain(value: RichTextValue): string {
  return value.richText.map((run) => run.text).join('');
}

/**
 * The plain text of any cell value, total over {@link CellValue}, so a caller reading a sheet
 * whose cells it did not write never has to switch on the union itself.
 *
 * This is the *value's* text, not the cell's *display* text: a number renders as JavaScript
 * renders it, with no number format applied (`0.1 + 0.2` is `"0.30000000000000004"`, a currency
 * cell has no currency sign), because the format lives on the style and this function is given
 * only the value. What each kind yields:
 *
 * - the empty cell (`null`) and an invalid `Date` → `""`, the two ways a cell has no text
 * - a boolean → `"TRUE"` / `"FALSE"`, Excel's own literals rather than JavaScript's
 * - a `Date` → a full ISO-8601 timestamp
 * - an error → its literal, e.g. `"#REF!"`, the same string the grid shows
 * - rich text → every run concatenated ({@link richTextToPlain})
 * - a hyperlink → its label, never its destination
 * - any of the three formula kinds → the text of the *cached result*, and `""` when the cell
 *   carries no cached result: the formula source is not text the sheet ever displayed
 */
export function cellValueToText(value: CellValue): string {
  return classify(value, TO_TEXT);
}

/**
 * Classify a value into its observable {@link ValueType}. This is total over
 * {@link CellValue}: every legal value has exactly one type. A `Date` is a date even
 * when its time is `NaN` (an invalid date is still a date-typed cell); serialization,
 * not the model, decides what to do with it.
 */
export function detectValueType(value: CellValue): ValueType {
  return classify(value, TO_TYPE);
}

// What each {@link ValueType} is carried by. Keyed by the type rather than listed, so a tenth kind
// added to `ValueType` is a compile error here and in every visitor below, naming the kind.
interface ValuePayload {
  null: null;
  number: number;
  string: string;
  boolean: boolean;
  date: Date;
  error: ErrorValue;
  formula: FormulaValue | SharedFormulaValue | DataTableFormulaValue;
  richText: RichTextValue;
  hyperlink: HyperlinkValue;
}

/** One answer per {@link ValueType}, given the narrowed value that kind is carried by. */
type ValueVisitor<R> = {[K in ValueType]: (value: ValuePayload[K]) => R};

/**
 * The one ladder that decides which kind a cell value is.
 *
 * Its order is load-bearing and not obvious: the *outer* shape wins, so a hyperlink whose label is
 * rich text is a hyperlink and not rich text, and a formula carrying a cached result is a formula and
 * not whatever the result is. Deciding that twice -- which is what naming the kind and rendering it
 * used to do, each with a comment pointing at the other -- is two chances to get it right and one
 * ordering that nothing checks. Here the callers supply only the per-kind answer.
 */
function classify<R>(value: CellValue, visit: ValueVisitor<R>): R {
  if (value === null) return visit.null(value);
  switch (typeof value) {
    case 'number':
      return visit.number(value);
    case 'string':
      return visit.string(value);
    case 'boolean':
      return visit.boolean(value);
    default:
      break;
  }
  if (value instanceof Date) return visit.date(value);
  if (isHyperlinkValue(value)) return visit.hyperlink(value);
  if (isFormulaValue(value) || isSharedFormulaValue(value) || isDataTableFormulaValue(value)) {
    return visit.formula(value);
  }
  if (isRichTextValue(value)) return visit.richText(value);
  if (isErrorValue(value)) return visit.error(value);
  return unsupportedValue(value);
}

const TO_TYPE: ValueVisitor<ValueType> = {
  null: () => ValueType.Null,
  number: () => ValueType.Number,
  string: () => ValueType.String,
  boolean: () => ValueType.Boolean,
  date: () => ValueType.Date,
  error: () => ValueType.Error,
  formula: () => ValueType.Formula,
  richText: () => ValueType.RichText,
  hyperlink: () => ValueType.Hyperlink,
};

const TO_TEXT: ValueVisitor<string> = {
  null: () => '',
  number: (value) => String(value),
  string: (value) => value,
  boolean: (value) => (value ? 'TRUE' : 'FALSE'),
  date: (value) => (Number.isNaN(value.getTime()) ? '' : value.toISOString()),
  error: (value) => value.error,
  // Every formula kind carries a `result` of the same optional shape; one recursion renders it.
  formula: (value) => (value.result === undefined ? '' : cellValueToText(value.result)),
  richText: richTextToPlain,
  hyperlink: (value) => (typeof value.text === 'string' ? value.text : richTextToPlain(value.text)),
};

/**
 * The verdict "this is not a cell value at all", raised from one place so that every function
 * total over {@link CellValue} rejects the same inputs with the same message.
 */
function unsupportedValue(value: CellValue): never {
  throw new TypeError(`unsupported cell value: ${describe(value)}`);
}

/** Whether a string is one of Excel's canonical error literals. */
export function isErrorCode(text: string): text is ErrorCode {
  return ERROR_SET.has(text);
}

function describe(value: unknown): string {
  if (typeof value === 'object' && value !== null) {
    return `object with keys [${Object.keys(value).join(', ')}]`;
  }
  return String(value);
}

/**
 * Normalise a raw assignment into a stored {@link CellValue}. `undefined` becomes the
 * empty cell (`null`); every other kind is validated by {@link detectValueType}. The
 * model never rewrites one value *kind* into another (a numeric-looking string stays a
 * string). The single exception is formula text, which is canonicalised to the OOXML
 * stored form (no leading `=`) so round-trips are idempotent regardless of how the
 * caller supplied it.
 *
 * @throws {TypeError} if the value is not a recognised cell-value shape.
 */
export function coerceCellValue(value: CellValue | undefined): CellValue {
  if (value === undefined) return null;
  // detectValueType throws on an unrecognised object shape, so this both validates
  // and gives a precise error at the assignment site rather than deep in serialization.
  detectValueType(value);
  if (isFormulaValue(value)) {
    const formula = stripLeadingEquals(value.formula);
    return formula === value.formula ? value : {...value, formula};
  }
  // A shared formula's `sharedFormula` is a master cell address, not formula text, so there is no
  // leading `=` to canonicalise; it passes through as given.
  return value;
}

/**
 * OOXML stores a formula's text in `<f>` without the UI's leading `=`. Strip a single
 * leading `=` so the stored form is canonical and a strict consumer (Google Sheets,
 * WPS), which rejects a stored formula beginning with `=`, accepts the file.
 */
function stripLeadingEquals(formula: string): string {
  return formula.startsWith('=') ? formula.slice(1) : formula;
}
