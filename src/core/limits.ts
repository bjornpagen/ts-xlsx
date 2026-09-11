// The limits Excel enforces that a caller may want to check before authoring: what a row or column can
// be sized to, and what a sheet or table may be named.
//
// `MAX_ROW`/`MAX_COLUMN` stay in `address.ts`, where they belong: those are structural, in that a
// reference past XFD is not a reference, so the addressing code that refuses them is the code that
// defines them. Everything here is a limit on *content* the addressing layer has no view of.
//
// The name limits arrived here late, and the gap they left is worth naming: they were module-private
// constants and bare literals inside the two validators, so the only way to ask the library what a
// legal sheet or table name is was to author one and catch the throw. A caller that wants to offer a
// user a rename box needs the answer before the throw, not from it.
//
// The two geometry ceilings bound how big a line can be *set*, and who enforces them is what makes
// them different from the addressing bounds: those are refused outright by the model, while these are
// limits on *assignment* in Excel's own UI and object model. They are not limits on what a package may
// carry. The schema types `ht` and `width` as a bare `xsd:double` with no ceiling on either, and
// Excel opens an over-limit file clean, with no repair prompt and no repair log, so a reader that threw
// on one would refuse a file Excel accepts. Nothing here is enforced, therefore, on either the
// read or the write path; these are for a caller that wants the size it states to be the size
// Excel uses, and the doc comments on `Row.height`/`Column.width` point here for that reason.
//
// What Excel does with an over-limit file is not symmetric, and the difference is why only one of
// these two numbers describes a value that survives: a row is silently clamped, a column is not.
// Both ceilings and both file behaviours were measured rather than quoted: Microsoft's own
// specifications page rounds the row-height ceiling to "409 points", and Excel takes 409.5. See
// docs/knowledge/specs/grid-geometry-limits-are-excels-not-the-schemas.md for the probes.

/**
 * The tallest row Excel accepts being set to, in points: it refuses 409.6 and takes 409.5. A row
 * asked to hold more wrapped text than this cannot grow to fit it, and the overflow is simply not
 * shown.
 *
 * A file may state more, and stating more loses the value rather than the file. Excel opens such a
 * package without complaint and silently clamps the row to 409.6, a tick *above* what it lets you
 * assign, being 8192 twentieths of a point and so the width of the field it is read into, then
 * writes 409.6 back on its next save. Check against this constant to keep a stated height from
 * quietly becoming a different one.
 */
export const MAX_ROW_HEIGHT = 409.5;

/**
 * The widest column Excel accepts being set to, in character units of the workbook's default font.
 * Excel refuses 255.4, so unlike the row-height ceiling this one is exactly integral.
 *
 * It is also the weaker of the two ceilings: it binds assignment only, and not a file at all.
 * Excel honours a `width` of 1000 read from a package, renders the column at it, and round-trips
 * it verbatim through its own save, where an over-limit row height is clamped away. So a width
 * above this is a column no Excel user could have produced by dragging, not a value at risk.
 *
 * Character units, not points or pixels: a width is a count of digits of the default font's
 * *maximum digit width*, which is why there is no companion `DEFAULT_COLUMN_WIDTH` constant here.
 * The width a column takes when it states none is a function of that font: the familiar 8.43 holds
 * for Calibri 11 and not for a workbook whose normal style says otherwise (Excel reports 8.09 for
 * Aptos Narrow 11). `sheet.properties.defaultColWidth` is what a file declares, and
 * docs/knowledge/specs/default-font-must-not-be-assumed-for-column-widths.md is why assuming a
 * value for it is a bug rather than a shortcut.
 */
export const MAX_COLUMN_WIDTH = 255;

/**
 * The longest sheet name Excel accepts, in UTF-16 code units. A longer one is refused outright rather
 * than truncated: a truncated name silently collides with its neighbours.
 */
export const MAX_SHEET_NAME_LENGTH = 31;

/**
 * The characters Excel forbids anywhere in a sheet name. A name may also not begin or end with an
 * apostrophe, which this pattern does not express because the position is what makes it illegal: a
 * sheet-qualified reference quotes the name with apostrophes, so one at either edge cannot be told
 * from the quoting.
 */
export const INVALID_SHEET_NAME_CHARS = /[*?:\\/[\]]/;

/** The longest table name Excel accepts, in UTF-16 code units. */
export const MAX_TABLE_NAME_LENGTH = 255;

/**
 * Excel's table-name grammar: start with a letter, underscore, or backslash; every later character a
 * letter, digit, period, or underscore. Unicode letters and digits are allowed.
 *
 * The pattern is the grammar only. Excel also refuses a name that reads as a reference: an A1 cell on
 * the grid (`T1`), an R1C1 reference (`R`, `C`, `RC`, `R1C1`, `R1X`), or `TRUE`/`FALSE`, in any case.
 * No regular expression states that honestly, because "on the grid" is a bound on the column letters
 * and the row number, so `Worksheet.addTable` checks it separately and a name passing this pattern
 * can still be refused.
 */
export const TABLE_NAME_PATTERN = /^[\p{L}\\_][\p{L}\p{N}._]*$/u;
