// Minimal XML serialisation helpers for the write path.
//
// Writing OOXML needs only correct escaping and well-formed structure, so no XML library is on the
// write path; the reader's own scanner is `xml-scan.ts` (ADR 0004). Escaping is the one hard, security-relevant requirement, since an unescaped `<`,
// `&`, or `"` produces a malformed package a consumer rejects, so it lives here,
// audited once, rather than sprinkled through the part emitters.
//
// A second class of character is harder than `&`: the ones XML 1.0 has no syntax for at
// all. `&#1;` is not an escape for U+0001, it is another way of spelling the same illegal
// document. So there are two behaviours here, and which one applies is decided by the
// destination rather than by the character.
//
// **Prose the user typed** is escaped `_xHHHH_`. That is SpreadsheetML's own convention, and
// where it applies Excel both writes it and decodes it on the way back in. Which elements are
// in that group is a measured fact, not a schema one, and each was settled over COM against
// Excel Desktop: a `<t>` inline string, the cached `<v>` of a `t="str"` formula cell, a
// threaded comment's `<text>`, and a `<headerFooter>` child's print header text, which all read
// back as the decoded character and are written back out re-escaped by Excel itself. So
// `textElement`, the `t="str"` result, the threaded-comment body and the header/footer children
// get `escapeSpreadsheetText`.
//
// **Everything else is refused.** The schema is no help in drawing that line: `ST_Xstring`
// types a sheet name and a defined name just as it types `<t>`, so following the type alone
// would have us escape a sheet name too, and a workbook whose tab reads `Sheet_x0001_A` is
// not a faithful rendering of the name the author asked for, it is a different name. Where
// there is no faithful representation the honest answer is `AuthoringError`, which is the
// stance `numberText` below already takes on a non-finite number for the same reason. That
// covers every structural string: sheet and defined names, formulas, table column names,
// document properties, and relationship targets.
//
// The two comment systems land on the same side of that line by two different routes, which
// is why the test is *does Excel decode it here* rather than *does this read as prose*. A
// legacy note's body is a `<t>` in a `CT_Rst`, the very type the convention is defined on. A
// threaded comment's `<text>` is a different element in the 2018 extension namespace with no
// documented escape at all; it took a measurement to put it here, and it decodes with the
// same closed grammar and the same single pass. A print header was refused on that reading
// until it too was measured: its `&`-prefixed section codes look like the only in-band syntax
// it has, but Excel decodes `_xHHHH_` in a header exactly as it does in a cell and writes a
// control character back out escaped, so the two syntaxes share one string without either
// knowing about the other.
//
// Refusing is a new throw on a path that used to "succeed" by producing a file Excel reports
// as damaged, so the failure moved earlier and got louder, which is the whole trade.

import {
  assertWritableNumber,
  AuthoringError,
  codePointHex,
  invalidToken,
  unrepresentable,
} from '../errors.ts';
import {XML_UNREPRESENTABLE, XML_UNREPRESENTABLE_GLOBAL} from './xml-chars.ts';

const TEXT_ESCAPES: Record<string, string> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
};

const ATTR_ESCAPES: Record<string, string> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&apos;',
  '\n': '&#10;',
  '\r': '&#13;',
  '\t': '&#9;',
};

/**
 * Why XML refuses a character, as the clause {@link unrepresentable} appends after the offset.
 * Declared beside the guard rather than inline so the reader's own refusal can quote the same
 * sentence if it ever needs to.
 */
const XML_REFUSAL =
  'XML 1.0 has no representation for it, and the _xHHHH_ escape that would carry it is a ' +
  'convention of cell values only';

/**
 * Refuse a string that XML cannot carry, naming the character and where it is so the author
 * can find it in a value they never inspected. These arrive from a database column or a CSV
 * field, not from a literal in the calling code.
 *
 * Every escape in this module runs it, including {@link escapeFormatCode}, whose one deliberate
 * divergence from {@link escapeAttr} is about the apostrophe and not about this.
 *
 * @throws {AuthoringError} naming the code point and its offset.
 */
function assertRepresentable(value: string): void {
  const error = unrepresentable(value, XML_UNREPRESENTABLE, XML_REFUSAL);
  if (error !== undefined) throw error;
}

/** Escape a string for use as XML element text. */
export function escapeText(value: string): string {
  assertRepresentable(value);
  return value.replace(/[&<>]/g, (ch) => TEXT_ESCAPES[ch] as string);
}

/** Escape a string for use inside a double-quoted XML attribute value. */
export function escapeAttr(value: string): string {
  assertRepresentable(value);
  return value.replace(/[&<>"'\n\r\t]/g, (ch) => ATTR_ESCAPES[ch] as string);
}

/**
 * {@link escapeAttr} for a number format code: the same escapes minus the apostrophe.
 *
 * A format code can legitimately contain `"` (a quoted literal like `"$"`), `<` and `&`, and Excel
 * writes one with bare apostrophes. `'` is not markup-significant inside a double-quoted attribute,
 * so leaving it keeps a round-tripped code byte-identical to the source; that divergence is the whole
 * reason this exists.
 *
 * It is built by *omitting* one entry from {@link ATTR_ESCAPES} rather than by listing the escapes it
 * does want, which is how it used to be written, in `io/xlsx/style-elements.ts`. An independently
 * maintained list does not stay a one-character divergence: that one silently dropped the tab, line
 * feed and carriage return escapes as well, and XML 1.0 3.3.3 requires a parser to normalise all
 * three in an attribute value to a space, so Excel read back a format code with a space where the
 * author wrote a tab. Our own reader preserved the raw character, so a round trip through this
 * library returned it unchanged and no test here could see the loss.
 */
export function escapeFormatCode(code: string): string {
  assertRepresentable(code);
  return code.replace(/[&<>"\n\r\t]/g, (ch) => ATTR_ESCAPES[ch] as string);
}

/**
 * Escape a cell value: the `_xHHHH_` convention first, then XML's own escapes.
 *
 * The order of the two replacements is load-bearing. Once `_xHHHH_` decodes to something, a
 * value that legitimately reads `_x0041_` would come back as `A`, so the underscore is
 * escaped first, and only where it begins a sequence that would otherwise decode, leaving
 * every other underscore alone. Doing that after the character escape would re-escape the
 * `_x005F_` it had just introduced.
 *
 * `escapeText` runs last and sees no unrepresentable character left, so its guard is a
 * standing proof that the escape was complete rather than a second check of the same thing.
 *
 * Reversed by `decodeSpreadsheetText` in `./xml-attrs.ts`, whose single left-to-right pass is what
 * makes the `_x005F_` step above reversible. Change either and read its comment first.
 */
export function escapeSpreadsheetText(value: string): string {
  return escapeText(
    value
      .replace(/_(x[0-9A-Fa-f]{4}_)/g, '_x005F_$1')
      .replace(
        XML_UNREPRESENTABLE_GLOBAL,
        (ch) => `_x${codePointHex(ch.codePointAt(0) as number)}_`,
      ),
  );
}

/**
 * Whether an element's text must be wrapped with `xml:space="preserve"` to survive a
 * round-trip. Leading/trailing whitespace is otherwise collapsed by consumers, so a
 * string cell value that begins or ends with a space needs the marker.
 */
function needsSpacePreserve(value: string): boolean {
  return value.length > 0 && (value !== value.trim() || /[\n\r\t]/.test(value));
}

/**
 * A `<t>` text element carrying an escaped string, marked `xml:space="preserve"` when its
 * whitespace would otherwise be collapsed. Shared by every string-bearing element (a plain
 * inline string cell, a rich-text run, a pooled string, a note's body) so all decode
 * identically on the way back.
 */
export function textElement(value: string): string {
  const space = needsSpacePreserve(value) ? ' xml:space="preserve"' : '';
  return `<t${space}>${escapeSpreadsheetText(value)}</t>`;
}

/**
 * A boolean attribute rendered with a leading space (` name="1"` / ` name="0"`), or '' when the value
 * is undefined. OOXML booleans serialise as 1/0; emitting the explicit `="0"` lets a writer force a
 * flag off against a consumer's default, while an unset (undefined) flag stays out of the element
 * entirely: the two-state-plus-absent contract every flag writer here shares.
 */
export function boolAttr(name: string, value: boolean | undefined): string {
  return value === undefined ? '' : ` ${name}="${value ? 1 : 0}"`;
}

/**
 * A free-string attribute rendered with a leading space (` name="a &amp; b"`), or '' when the value
 * is undefined. The third of the trio with {@link boolAttr} and {@link numAttr}: the value is prose
 * the caller chose, so it is escaped rather than checked.
 *
 * A token from a closed OOXML enumeration is the other case and goes through {@link checkedToken}
 * instead. Escaping a bogus token would produce a well-formed document that Excel still rejects,
 * which hides the mistake in the file rather than raising it at the call.
 */
export function textAttr(name: string, value: string | undefined): string {
  return value === undefined ? '' : ` ${name}="${escapeAttr(value)}"`;
}

/**
 * Refuse a token from a closed OOXML enumeration that the writer would otherwise emit verbatim, and
 * return it unchanged when it belongs. The public types already forbid an out-of-contract value, so
 * this fires only for one smuggled past them by an untyped caller, a `JSON.parse`, or a cast.
 *
 * The writer must never serialise such a value: it would be schema-invalid OOXML that Excel
 * sometimes tolerates yet the library's own reader, which narrows every such token through the same
 * guard, discards on read-back. Refusing at the write boundary keeps the two symmetric, garbage out
 * refused exactly as garbage in, so a value the writer accepts is always one that round-trips.
 *
 * @throws {AuthoringError} naming the value and the enumeration.
 */
export function checkedToken(
  value: string,
  isValid: (candidate: string) => boolean,
  kind: string,
): string {
  if (!isValid(value)) {
    throw invalidToken(kind, value);
  }
  return value;
}

/**
 * A numeric attribute rendered with a leading space (` name="42"`), or '' when the value is undefined
 * so that a count an author never set stays out of the element rather than fabricating a default.
 *
 * The value goes through {@link numberText}, so a non-finite one is refused here rather than written
 * as `NaN`. Every attribute on this path is `xsd:double` or `xsd:unsignedInt`, neither of which has a
 * spelling for it, so the alternative is a package Excel reports as damaged.
 *
 * @throws {AuthoringError} when the value is not finite.
 */
export function numAttr(name: string, value: number | undefined): string {
  return value === undefined ? '' : ` ${name}="${numberText(value)}"`;
}

/**
 * An integer attribute rendered with a leading space (` name="3"`), or '' when the value is undefined.
 *
 * For an `xsd:int` or `xsd:unsignedInt` attribute, whose lexical space has no spelling for a fraction
 * any more than for a NaN: `priority="1.5"` and `rank="-1"` make a package the schema rejects, so the
 * value is refused here rather than written. `min` is the floor the attribute's type implies, `0` for
 * an unsigned one.
 *
 * @throws {AuthoringError} when the value is not finite, not a safe integer, or below `min`.
 */
export function intAttr(
  name: string,
  value: number | undefined,
  min = -Number.MAX_SAFE_INTEGER,
): string {
  if (value === undefined) return '';
  assertWritableNumber(value);
  if (!Number.isSafeInteger(value) || value < min) {
    const floor = min === -Number.MAX_SAFE_INTEGER ? '' : ` of at least ${min}`;
    throw new AuthoringError(`cannot write ${name}="${value}": expected an integer${floor}`);
  }
  return ` ${name}="${value}"`;
}

/**
 * A numeric attribute written only when it differs from its schema default, with a leading space, or
 * '' when the value is undefined or equals `defaultValue`.
 *
 * The number is checked before it is compared, because the two answer different questions: whether
 * the value can be spelled, and whether it is worth spelling. Tested for truthiness instead, `NaN` is
 * falsy and was omitted as if it were a default of zero, so a gradient fill with `degree: NaN` was
 * written as an ordinary gradient rather than refused.
 *
 * @throws {AuthoringError} when the value is not finite.
 */
export function nonDefaultNumAttr(
  name: string,
  value: number | undefined,
  defaultValue: number,
): string {
  if (value === undefined) return '';
  const text = numberText(value);
  return value === defaultValue ? '' : ` ${name}="${text}"`;
}

/**
 * Refuse a `Date` OOXML cannot spell, naming the property that carries it.
 *
 * Two ways a `Date` fails to have a `dcterms:W3CDTF` / `xsd:dateTime` spelling, and neither is
 * caught by the types. An **Invalid Date** is truthy and is an instance of `Date`, so it passes
 * every guard short of this one and reaches `toISOString()`, which throws a bare
 * `RangeError: Invalid time value` -- outside this taxonomy, and naming neither the property nor
 * the document it was being written into. A year **outside 0000-9999** has no four-digit form, so
 * `toISOString()` falls back to ISO 8601's expanded `+275760-09-13T…` notation, which is not in
 * the lexical space of either type; that one does not throw at all, it writes a package Excel
 * offers to repair.
 *
 * @throws {AuthoringError} naming the property.
 */
export function assertWritableDate(date: Date, property: string): void {
  const time = date.getTime();
  if (Number.isNaN(time)) {
    throw new AuthoringError(`cannot write ${property}: it is an Invalid Date`);
  }
  const year = date.getUTCFullYear();
  if (year < 0 || year > 9999) {
    throw new AuthoringError(
      `cannot write ${property}: the year ${year} is outside the 0000-9999 range a W3CDTF ` +
        'timestamp can spell',
    );
  }
}

/**
 * A finite number serialises as its shortest round-trippable decimal; a non-finite one
 * has no OOXML numeric representation, so the writer refuses it rather than emit `NaN`.
 */
export function numberText(value: number): string {
  assertWritableNumber(value);
  return String(value);
}

export const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
