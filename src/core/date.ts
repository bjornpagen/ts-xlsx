// Excel date serials: the bridge between a JS `Date` and OOXML's numeric storage.
//
// A spreadsheet never stores a date as such: it stores a number (the "serial") and a
// number format that tells the viewer to render that number as a date. What a serial
// counts from is the workbook's own declaration, not a constant: in the 1900 system
// serial 1 is 1900-01-01, and in the 1904 system serial 0 is 1904-01-01. The catch in the
// first is a bug Excel has carried since 1985 and every consumer must reproduce for
// compatibility: the 1900 date system counts a phantom 1900-02-29 that never existed
// (serial 60), so serials at or below 59 sit one calendar day later than a naive
// "days since the epoch" offset would place them. Round-tripping a date means reproducing
// that quirk faithfully in both directions, or imported pre-March-1900 dates land a day
// early - and reading the workbook's declared system, or every date in a 1904 file lands
// four years and a day early instead.
//
// Dates here are treated as UTC wall-clock: a serial carries no timezone, so we convert
// against a UTC epoch in both directions. A `Date` written and read back through this
// module is exact; the caller owns any timezone interpretation before it reaches here.

const MS_PER_DAY = 86_400_000;

// Serial 0 in the 1900 date system is nominally 1899-12-30. The phantom leap day lives
// at serial 60, so the offset correction pivots there.
const EPOCH_1900_UTC = Date.UTC(1899, 11, 30);
const PHANTOM_SERIAL = 60;

// Serial 0 in the 1904 date system is 1904-01-01, and the system carries no phantom leap day: it
// was Excel for Macintosh's way of not inheriting the Lotus bug. A workbook declares which system
// it counts in, and the two are 1462 days apart, so reading a 1904 file as a 1900 one puts every
// date four years and a day early. Verified against Excel Desktop: 2023-03-15 is serial 45000 in a
// 1900 workbook and 43538 in a 1904 one.
const EPOCH_1904_UTC = Date.UTC(1904, 0, 1);

/**
 * Which date system a workbook counts its serials in: the `date1904` flag of `<workbookPr>`.
 *
 * Carried as the epoch year rather than as a boolean because it is what every conversion here
 * actually needs, and a boolean would mean a `? 1904 : 1900` at each of them: five chances to write
 * the ternary backwards, on a value whose wrongness is silent by construction (a date read under the
 * wrong system is still a perfectly good date, four years off).
 */
export type DateEpoch = 1900 | 1904;

/**
 * The number format applied to a `Date` cell that carries no explicit format of its own,
 * so the value renders, and reads back, as a date rather than a bare serial number.
 */
export const DEFAULT_DATE_NUMFMT = 'yyyy-mm-dd';

/**
 * Convert a JS `Date` to its Excel serial under the workbook's date system, reproducing the
 * phantom-leap-day quirk (1900 only) so the value renders on the calendar date Excel would show.
 * Fractional serials carry the time of day. The caller must reject a non-finite (invalid) date
 * before here.
 */
export function dateToSerial(date: Date, epoch: DateEpoch): number {
  if (epoch === 1904) return (date.getTime() - EPOCH_1904_UTC) / MS_PER_DAY;
  const days = (date.getTime() - EPOCH_1900_UTC) / MS_PER_DAY;
  // Days at or below the phantom (1900-02-28 is day 60 from the nominal epoch) are shifted
  // one earlier to skip the fake 1900-02-29 that Excel counts at serial 60. The whole day decides,
  // not the instant: compared with its time of day, noon on 1900-02-28 (day 60.5) escaped the shift
  // and landed on the day that does not exist.
  return Math.floor(days) <= PHANTOM_SERIAL ? days - 1 : days;
}

/**
 * Convert an Excel serial back to a UTC `Date` under the workbook's date system, accounting in the
 * 1900 system for the phantom 1900-02-29 so serial 1 reads as 1900-01-01 (not 1899-12-31) and
 * consecutive serials map to consecutive days.
 */
export function serialToDate(serial: number, epoch: DateEpoch): Date {
  if (epoch === 1904) return new Date(EPOCH_1904_UTC + serial * MS_PER_DAY);
  const dayOffset = serial < PHANTOM_SERIAL ? serial + 1 : serial;
  return new Date(EPOCH_1900_UTC + dayOffset * MS_PER_DAY);
}

/**
 * Parse a date-bearing text field (a Strict-mode `t="d"` cell value, a core-property timestamp) to a
 * `Date`, or `null` when it carries nothing parseable.
 *
 * The `null` is the point. `new Date('not-a-date')` is a `Date` whose time is `NaN`: it satisfies
 * `instanceof Date` and every guard downstream, survives into the model, and reaches serialisation,
 * where it writes back as the literal string `Invalid Date` and loses the value for good. Dropping it
 * at the boundary is the only reading that cannot lie.
 */
export function parseDateText(text: string): Date | null {
  if (text === '') return null;
  const date = new Date(text);
  return Number.isNaN(date.getTime()) ? null : date;
}

/**
 * Whether a number-format code renders its value as a date or time. A format is a date
 * format when, once its non-formatting sections are removed (bracketed color/locale/
 * condition directives, quoted literals, and escaped characters) any of the date/time
 * placeholder letters (`y m d h s`) remain. So `"$"#,##0.00` and `0.00%` are not dates
 * while `yyyy-mm-dd`, `dd/mm/yyyy`, and `[$-409]mmmm d, yyyy` are.
 */
export function isDateFormat(code: string): boolean {
  const stripped = code
    .replace(/\[[^\]]*\]/g, '')
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '');
  return /[ymdhs]/i.test(stripped);
}

/**
 * Surface a number stored under a date format as a `Date`, and leave everything else alone.
 *
 * OOXML has no date type: a date is a number plus a number format that renders it as one, so this
 * one test is what separates `45000` from `2023-03-15` across every reader. Only a plain number
 * qualifies; a string, a boolean or a formula result of another kind under a date format keeps its
 * own kind.
 *
 * One function rather than three copies of the test, because the corpus asserts the rule is identical
 * across serialisations: the `.xlsb` reader, the `.xlsx` cell decoder and the cached formula-result
 * decoder must all answer the same, and three spellings of one rule are three chances to disagree.
 * The workbook's {@link DateEpoch} rides along for exactly that reason: it is the one input to this
 * test that is not the cell's, and the three used to agree on the wrong answer for a 1904 file.
 */
export function coerceDateSerial<T>(
  value: T,
  numFmt: string | undefined,
  epoch: DateEpoch,
): T | Date {
  return typeof value === 'number' && numFmt !== undefined && isDateFormat(numFmt)
    ? serialToDate(value, epoch)
    : value;
}
