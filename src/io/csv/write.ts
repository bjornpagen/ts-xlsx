// CSV serialization: the flat-text sibling of the XLSX writer.
//
// A worksheet is a rectangle of typed cells; CSV is that rectangle flattened to delimited
// text. The lossy direction (styles, formulas-as-formulas, multiple sheets) is inherent to the
// format, so this writer makes the honest choices explicit: one selected sheet, each row sized to
// its own populated extent (never clamped to a sibling row's width), a formula rendered as its
// cached result, a Date rendered by a caller-supplied format or a full ISO-8601 timestamp. What a
// value reads as is `cellValueToText`'s answer, not a private one: a CSV field and `cell.text`
// disagreeing about the same cell would be a bug in one of them.
//
// `writeCsvText` yields the logical text; `writeCsv` encodes it to bytes and, for UTF-8 (the
// default), prepends a byte-order mark so a consumer such as Excel detects the encoding and does
// not mangle non-ASCII on open. The BOM is a byte-level marker, not part of the logical text.
//
// The split is also where the one non-obvious refusal lives: a lone surrogate is a perfectly good
// JavaScript string and has no UTF-8 encoding, so it survives `writeCsvText` and is refused by
// `writeCsv`. See {@link assertEncodable}.

import type {Cell} from '../../core/cell.ts';
import {formatSerialDate} from '../../core/date-format.ts';
import {
  type CellValue,
  cellValueToText,
  isDataTableFormulaValue,
  isFormulaValue,
  isSharedFormulaValue,
} from '../../core/value.ts';
import type {Workbook} from '../../core/workbook.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {AuthoringError, unrepresentable} from '../../errors.ts';
import {assertDelimiter, assertRowDelimiter} from './delimiter.ts';

/**
 * A byte encoding {@link writeCsv} can produce, spelled the way Node's `Buffer` spells it and
 * meaning the same bytes.
 *
 * Named here rather than taken from Node's `BufferEncoding`, which is what this option used to be:
 * that type is part of `@types/node`, so it made a browser consumer's public API surface depend on
 * Node's types, and it offered `base64` and `hex` as if they were output encodings for a text
 * format. The list is what a CSV consumer actually asks for.
 */
export type CsvEncoding =
  | 'ascii'
  | 'latin1'
  | 'ucs-2'
  | 'ucs2'
  | 'utf-8'
  | 'utf-16le'
  | 'utf16le'
  | 'utf8';

export interface CsvWriteOptions {
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

const UTF8_BOM = Uint8Array.of(0xef, 0xbb, 0xbf);

/**
 * The logical CSV text of one worksheet: no BOM, no byte encoding. Line N holds row N, so an empty
 * row before the last populated one is an empty line.
 *
 * @throws {RangeError} if `delimiter` or `rowDelimiter` is one the output could not be split on.
 */
export function writeCsvText(workbook: Workbook, options: CsvWriteOptions = {}): string {
  const sheet = selectSheet(workbook, options.sheetName);
  const delimiter = options.delimiter ?? ',';
  assertDelimiter(delimiter);
  const rowDelimiter = options.rowDelimiter ?? '\n';
  assertRowDelimiter(rowDelimiter, delimiter);

  // A line's position is its row number. `rows()` skips a row with nothing in it, so a gap is filled
  // with empty lines; otherwise A1 and A3 wrote `a\nc` and `c` read back in row 2. A row holding only
  // formatting (a hidden or resized row) writes nothing of its own: the gap fill gives it its empty
  // line when a later row needs one, and past the last row with cells it adds no trailing lines. The
  // output therefore depends on the cells alone, never on formatting a CSV cannot show.
  const lines: string[] = [];
  for (const row of sheet.rows()) {
    let width = 0;
    const byColumn = new Map<number, Cell>();
    for (const cell of row.cells) {
      byColumn.set(cell.col, cell);
      if (cell.col > width) width = cell.col;
    }
    if (width === 0) continue;
    const fields: string[] = [];
    for (let column = 1; column <= width; column++) {
      const value = byColumn.get(column)?.value ?? null;
      const text = options.map ? options.map(value, column - 1) : csvFieldText(value, options);
      fields.push(quoteField(text, delimiter, rowDelimiter));
    }
    while (lines.length < row.number - 1) lines.push('');
    lines.push(fields.join(delimiter));
  }
  return lines.join(rowDelimiter);
}

/**
 * The CSV bytes of one worksheet in the requested encoding, with a UTF-8 BOM by default.
 *
 * @throws {AuthoringError} if a field holds an unpaired surrogate and the encoding is UTF-8, which
 * cannot represent one. The alternative is a silent U+FFFD substitution.
 * @throws {RangeError} on a delimiter {@link writeCsvText} refuses.
 */
export function writeCsv(workbook: Workbook, options: CsvWriteOptions = {}): Uint8Array {
  const text = writeCsvText(workbook, options);
  const encoding = options.encoding ?? 'utf8';
  const utf8 = isUtf8(encoding);
  if (utf8) assertEncodable(text);
  const body = encode(text, encoding);
  const wantBom = options.bom ?? utf8;
  if (!wantBom || !utf8) return body;

  const out = new Uint8Array(UTF8_BOM.length + body.length);
  out.set(UTF8_BOM, 0);
  out.set(body, UTF8_BOM.length);
  return out;
}

// Node spells the same encoding two ways, and the BOM is owed to both.
function isUtf8(encoding: CsvEncoding): boolean {
  return encoding === 'utf8' || encoding === 'utf-8';
}

/**
 * The text's bytes in one of the {@link CsvEncoding} spellings, byte for byte what `Buffer.from`
 * produced when this went through Node: UTF-16 is little-endian code units with no BOM, and the
 * two byte-narrow encodings are the low byte of each code unit, `ascii` included, which Node
 * documents as equivalent to `latin1` in this direction rather than as a 7-bit filter.
 *
 * Written out because `Buffer` is a Node global, and a module the browser entry can reach may not
 * touch one (`scripts/check-browser-safe.ts`). `TextEncoder` covers UTF-8 and is the only encoder
 * the platform offers; the other two are a loop each.
 */
function encode(text: string, encoding: CsvEncoding): Uint8Array {
  if (isUtf8(encoding)) return new TextEncoder().encode(text);
  if (encoding === 'ascii' || encoding === 'latin1') {
    const bytes = new Uint8Array(text.length);
    for (let i = 0; i < text.length; i++) bytes[i] = text.charCodeAt(i) & 0xff;
    return bytes;
  }
  const bytes = new Uint8Array(text.length * 2);
  const view = new DataView(bytes.buffer);
  for (let i = 0; i < text.length; i++) view.setUint16(i * 2, text.charCodeAt(i), true);
  return bytes;
}

/**
 * A lone surrogate (half of an astral pair, typically what is left when a string was sliced through
 * the middle of one) has no UTF-8 encoding. `Buffer.from` does not say so: it substitutes U+FFFD
 * and returns bytes that look perfectly well-formed, so the character is gone and nothing failed.
 *
 * The XLSX writer can carry one because a cell value has the `_xHHHH_` convention to hide it in; a
 * CSV field is plain text with no escape to reach for, so refusing is the only honest answer left.
 *
 * Checked for UTF-8 alone. UTF-16 writes the code unit through verbatim and loses nothing, and the
 * byte-narrow encodings (`latin1`, `ascii`) mangle every non-ASCII character by the caller's own
 * explicit choice, and a surrogate is not a special case there.
 */
function assertEncodable(text: string): void {
  const error = unrepresentable(
    text,
    SURROGATES,
    'it is an unpaired surrogate, which UTF-8 cannot encode and CSV has no escape for',
  );
  if (error !== undefined) throw error;
}

// The `u` flag makes the pattern match code points, so a well-formed pair is one unit that no
// surrogate range can match and this means exactly "a surrogate that is not part of a pair".
const SURROGATES = /[\u{D800}-\u{DFFF}]/u;

function selectSheet(workbook: Workbook, name: string | undefined): Worksheet {
  if (name === undefined) {
    const first = workbook.worksheets[0];
    if (first === undefined) throw new AuthoringError('workbook has no worksheet to write as CSV');
    return first;
  }
  return workbook.requireWorksheet(name);
}

// A field is the value's plain text, with one CSV-only deviation: `dateFormat`/`dateUTC` let a
// caller render dates in something other than ISO-8601. That reaches inside a formula's cached
// result too, which is why the recursion is here rather than delegated wholesale.
function csvFieldText(value: CellValue, options: CsvWriteOptions): string {
  if (value instanceof Date && options.dateFormat !== undefined) {
    return formatSerialDate(value, options.dateFormat, options.dateUTC ?? false);
  }
  if (isFormulaValue(value) || isSharedFormulaValue(value) || isDataTableFormulaValue(value)) {
    return value.result === undefined ? '' : csvFieldText(value.result, options);
  }
  return cellValueToText(value);
}

// CR and LF are quoted whatever the row delimiter is, because a reader splits records on them either
// way; a custom row delimiter is quoted as well, or `a|b` under `rowDelimiter: '|'` reads as two rows.
function quoteField(field: string, delimiter: string, rowDelimiter: string): string {
  if (
    field.includes(delimiter) ||
    field.includes(rowDelimiter) ||
    field.includes('"') ||
    field.includes('\n') ||
    field.includes('\r')
  ) {
    return `"${field.replace(/"/g, '""')}"`;
  }
  return field;
}
