import {XlsxError} from '../../errors.ts';

/**
 * Thrown when CSV text parses but cannot become a worksheet: a record with more fields than a sheet
 * has columns, or more records than it has rows.
 *
 * CSV has no malformed markup to reject, since any text splits into fields somehow, so the only way a
 * file fails is by not fitting the grid. That is still the file's fault rather than the caller's,
 * which is why it is this error and not the `RangeError` the grid raises for an authored cell out of
 * bounds.
 */
export class CsvParseError extends XlsxError {
  override readonly name = 'CsvParseError';
  override readonly code = 'malformed-input';
}
