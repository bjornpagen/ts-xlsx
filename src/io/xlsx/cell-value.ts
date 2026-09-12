// Decoding a worksheet cell's on-disk `<c>` payload into a model {@link CellValue}.
//
// This is the single value-decoding surface both readers share: the buffered reader
// (`./read.ts`) and the streaming row reader (`./read-rows.ts`). Keeping it in one place is
// what guarantees a cell read one row at a time decodes identically to the same cell read as
// part of a whole workbook. A divergence here would be a silent data bug in exactly one path.

import {coerceDateSerial, type DateEpoch, parseDateText} from '../../core/date.ts';
import {unmangleFunctions} from '../../core/formula.ts';
import {
  type CellValue,
  type ErrorCode,
  type FormulaResult,
  isErrorCode,
  type RichTextRun,
  type RichTextValue,
} from '../../core/value.ts';
import {decodeSpreadsheetText, numFinite, numInteger} from '../../xml/xml-attrs.ts';
import {boolTristate} from '../../xml/xml-scan.ts';

/**
 * One entry of the shared-strings pool. A `<si>` built from a bare `<t>` is a plain string; a `<si>`
 * built from `<r>` runs is rich text, so a `t="s"` cell can resolve to either kind, and rich text
 * that Excel pooled reads back with its per-run formatting intact rather than flattened to text.
 */
export type SharedString = string | RichTextValue;

/** The raw, still-textual pieces of a `<c>` element the SAX pass has gathered. */
export interface RawCell {
  /** The `t` attribute (`s`, `str`, `inlineStr`, `b`, `e`, `d`, or '' for a number). */
  readonly type: string;
  readonly hasFormula: boolean;
  readonly formula: string;
  /** The range an `<f t="array">` fills when it starts at this cell, as Excel spells a `ref`. */
  readonly arrayRef?: string | undefined;
  /** Whether the cell's `cm` points at cell metadata marking a dynamic array. */
  readonly dynamicArray?: boolean;
  /** The error the cell's `vm` names through a rich value, which is the error a `t="e"` cell holds. */
  readonly valueError?: ErrorCode | undefined;
  /** Whether a `<v>` was present, `<v/>` included. */
  readonly hasValue: boolean;
  readonly valueText: string;
  /** Whether an `<is>` was present, `<is/>` included. */
  readonly hasInlineString: boolean;
  readonly inlineText: string;
  /** The formatted runs of a rich inline string, when the `<is>` held `<r>` elements rather than a
   * bare `<t>`. Absent (or empty) for a plain inline string, which decodes to `inlineText`. */
  readonly richTextRuns?: readonly RichTextRun[];
}

/**
 * Decode a gathered cell into its model value. A formula cell becomes a `{formula, result?}`
 * object (the on-disk `_xlfn.`/`_xlpm.` mangling stripped back to the readable name), or an array
 * formula over its range, dynamic when the cell's metadata says so; a plain
 * numeric cell under a date number format becomes a {@link Date}; everything else decodes by its
 * `t` type. `numFmt` is the cell's resolved number-format code, used only for date detection, and
 * `epoch` the workbook's date system, which is what a serial under such a format counts from.
 * `definedNames` is the workbook's defined names as `definedNameKeys` spells them, which decide whether
 * a function passed as a value sheds its `_xleta.`.
 */
export function decodeCellContent(
  raw: RawCell,
  sharedStrings: readonly SharedString[],
  numFmt: string | undefined,
  epoch: DateEpoch,
  definedNames: ReadonlySet<string>,
): CellValue {
  if (raw.hasFormula) {
    const stored = unmangleFunctions(raw.formula, definedNames);
    const result = raw.hasValue
      ? decodeFormulaResult(raw.type, raw.valueText, numFmt, epoch, raw.valueError)
      : undefined;
    const cached = result === undefined ? {} : {result};
    if (raw.arrayRef !== undefined) {
      return {
        shareType: 'array',
        formula: stored,
        ref: raw.arrayRef,
        ...(raw.dynamicArray === true ? {dynamic: true} : {}),
        ...cached,
      };
    }
    return {formula: stored, ...cached};
  }
  // An inline string built from `<r>` runs is rich text: surface its runs rather than flattening
  // them to the concatenated `inlineText` a plain string would decode to.
  if (raw.type === 'inlineStr' && raw.richTextRuns !== undefined && raw.richTextRuns.length > 0) {
    return {richText: raw.richTextRuns};
  }
  const value = decodeValue(raw, sharedStrings);
  // A number stored under a date format is a date serial: surface it as a Date so a written date
  // round-trips as a date, not a bare number.
  return coerceDateSerial(value, numFmt, epoch);
}

// A cell Excel shows as blank reads as `null`, whatever its `t` claims. The two string types are text
// once their carrier is present, so `<v/>` under `t="str"` and `<is/>` under `t="inlineStr"` are empty
// strings; every other type is only its `<v>` text, and an empty `<v>` is the same blank as a missing
// one (Excel 16.0 build 20326, `test/corpus/cases/typed-cell-without-value-reads-blank.case.ts`). A
// token the type has no reading for (`<v>2</v>` under `t="b"`, a pool index past the end) is a package
// Excel refuses to open, so there is no answer to match, and it reads as no value too.
function decodeValue(raw: RawCell, sharedStrings: readonly SharedString[]): CellValue {
  const {valueText} = raw;
  switch (raw.type) {
    case 'inlineStr':
      // Already decoded per `<t>` as it was gathered; the accumulator owns that seam.
      return raw.hasInlineString ? raw.inlineText : null;
    case 'str':
      return raw.hasValue ? decodeSpreadsheetText(valueText) : null;
    case 'd':
      // A Strict-mode (ISO/IEC 29500 Strict) date cell stores an ISO 8601 value directly, not a
      // serial. Parse it literally, since an ISO date is UTC, so it reads as the date it states rather
      // than a 1900-epoch serial the transitional decoder would fabricate from the text.
      return parseDateText(valueText);
    case 's': {
      // A `t="s"` cell indexes the shared pool; the entry is a plain string or, when Excel pooled a
      // rich value, a {@link RichTextValue} whose runs surface here rather than being flattened.
      // Through the integer grammar, because `Number('')` is 0 and would resolve an empty `<v>` to the
      // first pooled string.
      const index = numInteger(valueText, 0);
      return index === undefined ? null : (sharedStrings[index] ?? null);
    }
    case 'b':
      return boolTristate(valueText) ?? null;
    case 'e':
      if (valueText === '') return null;
      return errorOf(valueText, raw.valueError);
    default:
      // Not a bare `Number()`, which reads an unparseable token as `NaN`, a number that satisfies every
      // guard downstream and that the writer then refuses to emit, and an empty `<v>` as zero.
      return numFinite(valueText) ?? null;
  }
}

/** Decode a formula's cached `<v>` result by its `t` type, coercing a numeric result under a date
 * `numFmt` to a {@link Date} exactly as a bare numeric cell is, so a date-valued formula result
 * (e.g. `TODAY()`) reads back as a Date, not a serial. Shared by the buffered reader's shared-formula
 * clone resolution, which caches a result the same way a plain formula cell does. */
export function decodeFormulaResult(
  type: string,
  valueText: string,
  numFmt: string | undefined,
  epoch: DateEpoch,
  valueError?: ErrorCode,
): FormulaResult | undefined {
  return coerceDateSerial(decodeResult(type, valueText, valueError), numFmt, epoch);
}

// A typed error cell's error. Excel stores an error it has no literal for as `#VALUE!` and names the real
// one through the cell's `vm`, which therefore wins over the `<v>`. A code this library does not list
// keeps its text: a producer writing `#PYTHON!` literally wrote data, not an error Excel reads.
function errorOf(valueText: string, valueError: ErrorCode | undefined): FormulaResult {
  if (valueError !== undefined) return {error: valueError};
  return isErrorCode(valueText) ? {error: valueText} : valueText;
}

// The formula-result subset of `decodeValue`: a cached result is only ever a string, boolean,
// error, or number, never a shared-string index, inline string, or Strict-mode date, so this
// handles just those cases rather than the full cell-value grammar.
//
// `undefined` is "no cached result", and an empty `<v>` is one, whatever `t` says. With calculation
// set to manual, so it shows the cache, Excel 16.0 (build 20326) showed an empty cached result under
// `t="b"`, `t="e"` and a number exactly as it showed a formula with no `<v>`: an empty string, not
// FALSE, an error or 0 (`test/corpus/fixtures/excel-oracle/formula-empty-cached-result.json`). Under
// `t="str"` a present `<v>` is text, as it is on a plain cell.
function decodeResult(
  type: string,
  valueText: string,
  valueError: ErrorCode | undefined,
): FormulaResult | undefined {
  switch (type) {
    case 'str':
      // The cached result of a string formula is a cell value, and Excel escapes and decodes it as
      // one. Verified on this host, a `<v>` of `_x0041_` under `t="str"` reads back as `A`. The
      // sibling `<v>` types are not text: a number, a boolean, and an error code have no `_` in
      // their grammars, so only this branch decodes.
      return decodeSpreadsheetText(valueText);
    case 'b':
      return boolTristate(valueText);
    case 'e':
      if (valueText === '') return undefined;
      return errorOf(valueText, valueError);
    default:
      // Narrowed exactly as a plain numeric cell is: an unparseable cached result is no cached
      // result, which is the state the cell would have reached anyway on the next write.
      return numFinite(valueText);
  }
}
