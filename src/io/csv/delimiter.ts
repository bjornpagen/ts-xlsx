// What a CSV delimiter may be, for both halves of the codec.
//
// The reader refused a multi-character delimiter and the writer accepted anything, in a module whose
// header is entirely about lossless round-tripping. So `readCsv(writeCsv(wb, {delimiter: "||"}))`
// threw on text this same codec had just produced, and `writeCsv({delimiter: ""})` was worse than a
// throw: `field.includes("")` is true of every field, so every field was quoted and the output had no
// separators in it at all -- a file that parses as one column and loses nothing visibly.
//
// One validator, called from both entry points, is the only arrangement in which the two halves
// cannot disagree about what a delimiter is.

import {quoted} from '../../errors.ts';

// The characters the parser gives a meaning of their own: a quote opens a quoted field, and CR and LF
// end a record. As a field delimiter each one is ambiguous with that meaning. With `"` as the
// delimiter, `a"b"c` read back as the single field `abc`.
const RESERVED = new Set(['"', '\r', '\n']);

/**
 * Refuse a field delimiter the codec cannot honour.
 *
 * A `RangeError` rather than an `AuthoringError`: this is a single scalar out of range, which
 * `src/errors.ts` reserves for the native types on the grounds that they exist for exactly that.
 *
 * @throws {RangeError} if the delimiter is not exactly one character, or is a quote, CR or LF.
 */
export function assertDelimiter(delimiter: string): void {
  if (delimiter.length !== 1) {
    throw new RangeError(`CSV delimiter must be a single character, got ${quoted(delimiter)}`);
  }
  if (RESERVED.has(delimiter)) {
    throw new RangeError(
      `CSV delimiter ${quoted(delimiter)} is reserved: a quote opens a quoted field, and CR and LF end a row`,
    );
  }
}

/**
 * Refuse a row delimiter the writer cannot keep apart from the data.
 *
 * Only the writer takes one: the reader recognises LF, CRLF and a bare CR, and a record boundary it
 * cannot see is a choice the caller makes for a consumer other than this codec. What is refused here
 * is output no consumer can split. An empty delimiter joins every row into one line. One that
 * contains the field delimiter reads as an extra field at every row boundary, and one that contains a
 * quote collides with the quoting that is supposed to protect a field from it.
 *
 * @throws {RangeError} if the row delimiter is empty, or contains the field delimiter or a quote.
 */
export function assertRowDelimiter(rowDelimiter: string, delimiter: string): void {
  if (rowDelimiter === '') {
    throw new RangeError('CSV row delimiter must not be empty');
  }
  if (rowDelimiter.includes(delimiter) || rowDelimiter.includes('"')) {
    throw new RangeError(
      `CSV row delimiter ${quoted(rowDelimiter)} must contain neither the field delimiter ${quoted(delimiter)} nor a quote`,
    );
  }
}
