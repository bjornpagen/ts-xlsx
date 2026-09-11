// CSV parsing: flat delimited text back into a one-sheet workbook.
//
// The hard part of reading CSV is not splitting fields; it is deciding a field's *type* without
// corrupting data. The rules here are deliberate and lossless-by-default:
//   - An empty field is the empty cell (`null`); a whitespace-only field is a string, never the
//     number 0 that `Number("   ")` would silently produce.
//   - A numeric-looking field becomes a number only when a double holds it exactly: at most 15
//     significant digits, and within the safe-integer range. A 16-digit card number or a
//     20-digit account number is kept as its original string, so no digits are lost.
//   - Only a strictly-formatted ISO date (`YYYY-MM-DD`, optional time) naming a real calendar day
//     becomes a Date; padded ids and dash-codes such as `2020-00001` or `1-3`, and `2024-02-30`,
//     stay strings.
// A caller can override coercion wholesale with `map` (e.g. the identity function to keep every
// field a raw string, preserving leading zeros).

import {MAX_COLUMN, MAX_ROW} from '../../core/address.ts';
import type {CellValue} from '../../core/value.ts';
import {Workbook} from '../../core/workbook.ts';
import {assertDelimiter} from './delimiter.ts';
import {CsvParseError} from './errors.ts';

export interface CsvReadOptions {
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

/**
 * Parse CSV text (or UTF-8 bytes) into a workbook holding a single worksheet.
 *
 * @throws {CsvParseError} if a record has more fields than a worksheet has columns, or the data
 *   records outnumber a worksheet's rows.
 */
export function readCsv(input: string | Uint8Array, options: CsvReadOptions = {}): Workbook {
  // `ignoreBOM` means "do not strip it", so a leading BOM survives the decode and `stripBom` stays
  // the single place that knows about one, for text and bytes alike.
  const text = stripBom(
    typeof input === 'string' ? input : new TextDecoder('utf-8', {ignoreBOM: true}).decode(input),
  );
  const delimiter = options.delimiter ?? ',';
  assertDelimiter(delimiter);

  // The header record takes no row of the sheet, so it does not count against the row limit.
  let rows = parseCsvRows(text, delimiter, MAX_ROW + (options.headers ? 1 : 0));
  if (options.headers) rows = rows.slice(1);

  const coerce = options.map ?? defaultCsvCoerce;
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet(options.sheetName ?? 'Sheet1');
  sheet.addRows(rows.map((fields) => fields.map((field, index) => coerce(field, index))));
  return workbook;
}

function stripBom(text: string): string {
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

// A character-scan parser. A quote opens a quoted field only as the field's first character, the
// way Excel reads one: `5'10"` is a height, and taking its quote as an opener swallowed every
// delimiter and row break up to the next quote in the file. Inside a quoted field a doubled quote is
// one quote, and text after the closing quote is literal (`"ab"cd` reads as `abcd`, as in Excel). A
// row ends on LF, CRLF or a bare CR. A final trailing newline does not yield a spurious empty row.
//
// The grid bounds are enforced during the scan rather than after it, so a hostile file of millions of
// delimiters is refused at the first field past the limit instead of being allocated whole first.
function parseCsvRows(text: string, delimiter: string, maxRecords: number): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  let atFieldStart = true;

  const endField = (): void => {
    if (row.length === MAX_COLUMN) {
      throw new CsvParseError(
        `CSV record ${rows.length + 1} has more than ${MAX_COLUMN} fields, the most columns a worksheet holds`,
      );
    }
    row.push(field);
    field = '';
    atFieldStart = true;
  };
  const endRow = (): void => {
    endField();
    if (rows.length === maxRecords) {
      throw new CsvParseError(
        `CSV has more than ${MAX_ROW} data records, the most rows a worksheet holds`,
      );
    }
    rows.push(row);
    row = [];
  };

  for (let i = 0; i < text.length; i++) {
    const ch = text.charAt(i);
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += ch;
      }
      continue;
    }
    if (ch === '"' && atFieldStart) {
      inQuotes = true;
      atFieldStart = false;
    } else if (ch === delimiter) {
      endField();
    } else if (ch === '\n') {
      endRow();
    } else if (ch === '\r') {
      // A CR before an LF is the first half of a CRLF pair and the LF ends the row. A CR alone
      // is a classic-Mac line ending and ends the row on its own: dropping it would splice the
      // next row's first field onto this row's last and lose every row boundary in the file,
      // which is silent corruption rather than the lossless read this module promises.
      if (text[i + 1] !== '\n') endRow();
    } else {
      field += ch;
      atFieldStart = false;
    }
  }

  const trailingNewline = text.endsWith('\n') || text.endsWith('\r');
  if (!(trailingNewline && row.length === 0 && field === '')) endRow();
  return rows;
}

const NUMERIC = /^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/;
const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2})(?::(\d{2}))?)?$/;

// Every decimal of up to 15 significant digits survives a trip through a double and back; at 16,
// some do not. Excel keeps numbers to the same 15.
const MAX_EXACT_SIGNIFICANT_DIGITS = 15;

function defaultCsvCoerce(field: string): CellValue {
  if (field === '') return null;
  const numeric = NUMERIC.exec(field);
  if (numeric) {
    const value = Number(field);
    const exact =
      Number.isFinite(value) &&
      Math.abs(value) <= Number.MAX_SAFE_INTEGER &&
      significantDigits(numeric[1] ?? '') <= MAX_EXACT_SIGNIFICANT_DIGITS;
    return exact ? value : field;
  }
  const iso = ISO_DATE.exec(field);
  if (iso) {
    const date = isoToDate(iso);
    if (date !== null) return date;
  }
  return field;
}

// Leading and trailing zeros only place the decimal point, which costs a double no precision.
function significantDigits(mantissa: string): number {
  return mantissa.replace('.', '').replace(/^0+/, '').replace(/0+$/, '').length;
}

function isoToDate(match: RegExpExecArray): Date | null {
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const hour = match[4] === undefined ? 0 : Number(match[4]);
  const minute = match[5] === undefined ? 0 : Number(match[5]);
  const second = match[6] === undefined ? 0 : Number(match[6]);
  if (hour > 23 || minute > 59 || second > 59) return null;
  // `setUTCFullYear` rather than `Date.UTC`, which reads years 0 to 99 as 1900 plus the year. Both
  // roll an impossible day forward (`2024-02-30` becomes March 1), so the parts are read back: a date
  // that does not return them unchanged was not a calendar day.
  const date = new Date(0);
  date.setUTCFullYear(year, month - 1, day);
  date.setUTCHours(hour, minute, second, 0);
  const real =
    date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day;
  return real ? date : null;
}
