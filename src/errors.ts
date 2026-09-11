// The failure taxonomy every layer throws through.
//
// Typed error classes grew up one per subsystem that happened to want one, sharing no ancestor, so a
// caller could not ask "did *this library* fail?" without naming all of them, while the model's own
// validation threw bare `Error`, distinguishable only by string-matching the message. For a library
// whose input is untrusted files, "was it my file or my call?" is a first-class question, and it had
// no answer.
//
// Two levels of branch, deliberately non-redundant:
//   - `code` says what *kind* of failure this is: the question a caller usually has.
//   - `name` (and `instanceof`) says exactly which one: the question a caller occasionally has.
// A code shared by several classes is the design, not an oversight: a `code` in 1:1 correspondence
// with the classes would carry nothing the class did not already carry.
//
// Scalar argument validation stays *outside* this taxonomy. See {@link AuthoringError} for where the
// line falls and why.
//
// This module sits below every layer that throws (`src/xml`, `src/core`, `src/io`, `src/vba`,
// `src/customui` all import it), so it reaches for nothing above it; `scripts/check-layering.ts`
// enforces that. Its one import is `./hex.ts`, which is a leaf for the same reason and exists so that
// the six sites that render a number as hex, this one included, share the pad width and the case.

import {hex} from './hex.ts';

/**
 * What kind of failure an {@link XlsxError} reports. This is the branch most callers want, and it is
 * deliberately coarse: the four answers are the four different things a caller would *do* next.
 *
 * - `'unsupported-format'`: the input is not a container this library reads at all (a legacy `.xls`,
 *   a blob that is not a spreadsheet). Nothing is wrong with the file; it is the wrong file *for us*.
 * - `'malformed-input'`: a part we do read is corrupt or does not conform to its specification. The
 *   file is broken, or hostile.
 * - `'authoring'`: the caller described a document that cannot exist. The bug is in the calling code.
 * - `'internal'`: an invariant this library maintains did not hold. It should be unreachable; if it
 *   fires, the bug is ours.
 *
 * There is deliberately no "not implemented yet" code. Every candidate for one turned out to be an
 * unreachable exhaustiveness guard (so: `'internal'`), and the one genuine feature gap (a binary
 * `.xlsb` cannot be row-streamed) is already reported by {@link UnsupportedFormatError}'s `format`
 * branch. A code with no throw site would be a promise the library does not keep.
 *
 * **Which of these is worth reporting upstream.** `'internal'` always is, and says so at runtime.
 * `'unsupported-format'` and `'malformed-input'` are worth reporting when the file in hand opens
 * cleanly in Excel: that combination means we are the ones who cannot read it, which is a gap, not
 * a corrupt input. `'authoring'` is worth reporting only if the document it refused is one a real
 * workbook can express. See `skills/ts-xlsx-upstream` for how to file one.
 */
export type XlsxErrorCode = 'unsupported-format' | 'malformed-input' | 'authoring' | 'internal';

/**
 * The common ancestor of every error this library raises deliberately. `catch (e) { if (e instanceof
 * XlsxError) … }` is the one-line answer to "was that us?", and {@link code} then says which kind of
 * failure it was without naming a single subclass.
 *
 * Abstract on purpose: a failure always has a kind, so there is never a reason to throw the base.
 * Every subclass fixes {@link code} to a literal, which makes the class hierarchy a discriminated
 * union, so narrowing on `error.code` narrows the type.
 *
 * The constructor is inherited from `Error`, so every subclass accepts `{cause}`; layers that wrap a
 * lower-level failure are expected to pass it rather than flatten it into the message.
 */
export abstract class XlsxError extends Error {
  abstract readonly code: XlsxErrorCode;
}

/**
 * Thrown when the caller describes a document that cannot exist: a pivot table with no row field, a
 * table whose columns do not span its range, a merge that overlaps another, a workbook with no
 * worksheets. The document model, not a single argument, is what is wrong.
 *
 * **Where the line falls against native errors.** A single scalar that is out of range, unparseable,
 * or the wrong type stays a native `RangeError` / `SyntaxError` / `TypeError`: those types exist for
 * exactly that, every caller already reads them, and wrapping them would make this taxonomy a
 * re-implementation of the language's. `AuthoringError` starts where a *composite* is internally
 * inconsistent, or contradicts something a workbook can express. `getColumn(0)` is a `RangeError`;
 * a table that names a column twice is an `AuthoringError`.
 *
 * It is always the calling code that is wrong, never the input file; a malformed file raises a
 * `'malformed-input'` error instead.
 */
export class AuthoringError extends XlsxError {
  override readonly name = 'AuthoringError';
  override readonly code = 'authoring';
}

/**
 * Where a report of an `'internal'` failure should go. Appended to every {@link InternalError}
 * message rather than left to the docs, because a stack trace is the only artefact of the failure
 * that reaches whoever (or whatever) is debugging it, and a pointer anywhere else is a pointer they
 * have to already be looking for.
 */
const REPORT_NOTICE =
  'This is a bug in ts-xlsx, not in your file or your code. Please report it:\n' +
  'https://github.com/shbernal/ts-xlsx/issues/new?template=agent-report.yml&labels=agent-reported';

/**
 * Thrown where an invariant the library itself maintains turns out not to hold: typically an index
 * that a preceding pass proved to be in range, re-checked because `noUncheckedIndexedAccess` makes
 * the possibility of `undefined` explicit and casting it away would be worse.
 *
 * No caller can provoke one, so it is not a failure mode to handle: seeing it means the bug is ours.
 * It exists as a distinct type so that "unreachable" is *stated* rather than implied by a bare
 * `Error`, which reads identically to a throw nobody has classified yet.
 *
 * It is the one class in the taxonomy that rewrites its own message, appending {@link REPORT_NOTICE}
 * below the invariant that broke. The constructor is where that lives so a throw site added later
 * inherits it; the alternative, a notice pasted at each of the throw sites, is one every future
 * site can forget. Every other class leaves `message` exactly as given: `'malformed-input'` is a
 * routine outcome for a library that reads untrusted files, and a "report this" banner on each
 * corrupt input would train callers to ignore the one banner that always means something.
 */
export class InternalError extends XlsxError {
  override readonly name = 'InternalError';
  override readonly code = 'internal';

  constructor(message?: string, options?: ErrorOptions) {
    super(message === undefined ? REPORT_NOTICE : `${message}\n\n${REPORT_NOTICE}`, options);
  }
}

/**
 * A name as it should appear inside an error message: quoted, and unambiguous whatever it contains.
 *
 * Almost every message this library throws names something the caller or the file chose - a sheet, a
 * table, a defined name, a module stream, a part path - and the tree had grown three ways of setting
 * that name off from the prose, split by directory rather than by intent: `"..."`, `'...'`, and
 * `JSON.stringify`. Only the third survives a name that itself contains a quote, which is exactly the
 * class of name a spreadsheet permits: a sheet called `Q1 "draft"` renders under either literal
 * spelling as a message whose reader cannot tell where the name ends. It also renders a name
 * containing a newline or a zero-width character as one that looks identical to a name that does not,
 * which for a library whose input is untrusted is the difference between a diagnostic and a decoy.
 *
 * So: one spelling, and the one that escapes. `JSON.stringify` is the whole implementation - the
 * point is not the algorithm but that every throw site reaches the same one.
 *
 * Not exported from the `/errors` entry: this is how messages are written, not part of the taxonomy
 * a caller catches.
 */
export function quoted(name: string): string {
  return JSON.stringify(name);
}

/**
 * The one message for "that token is not in the enumeration", owned here because both sides of a
 * layering boundary throw it.
 *
 * `checkedToken` in `src/xml/xml.ts` is the general form, and `core/` may not import a serialisation
 * to reach it. The response to that constraint had been to transcribe the sentence into
 * `core/table-style.ts` character for character, which is a copy that no gate can see is a copy. This
 * module is the one place below both of them, so the template lives here and neither side owns it.
 */
export function invalidToken(kind: string, value: string): AuthoringError {
  return new AuthoringError(
    `invalid ${kind} ${quoted(value)}: not a value the OOXML enumeration allows`,
  );
}

/**
 * Refuse a number OOXML cannot spell. Every numeric attribute in the format is `xsd:double`,
 * `xsd:unsignedInt` or a bounded flavour of one, and none of those lexical spaces has a form for a
 * NaN or an infinity, so a value that reaches the file as `NaN` is a package Excel reports as
 * damaged. Neither has the formula grammar, which is why the guard is here rather than in the XML
 * serialiser: the two spellings of a number this library writes -- an attribute's and a formula
 * literal's -- sit on opposite sides of the `core`/`xml` boundary and refuse the same values.
 *
 * Exported for the callers that must refuse before they write. A number can be unwritable and still
 * be read on the way to the bytes: compared, summed, walked. A comparison against `NaN` or an
 * infinity silently takes the wrong branch long before the value would have been serialised.
 *
 * @throws {AuthoringError} naming the value.
 */
export function assertWritableNumber(value: number): void {
  if (Number.isFinite(value)) return;
  throw new AuthoringError(
    `cannot write a non-finite number (${value}): it has no OOXML representation`,
  );
}

/** `U+0001`-style spelling of a code point, for an escape body or an error message. */
export function codePointHex(codePoint: number): string {
  return hex(codePoint, 4);
}

/**
 * The error for a string carrying a character the target format cannot encode, or `undefined` when
 * it carries none. Naming the code point and its offset is the whole value of the message: these
 * strings arrive from a database column or a CSV field, not from a literal the author can see.
 *
 * Two formats refuse characters for unrelated reasons -- XML 1.0's `Char` production, and UTF-8's
 * inability to encode a lone surrogate -- and each had grown its own copy of the rendering, which
 * had already drifted: the CSV one omitted the zero padding, invisibly, because a surrogate is
 * always four digits and nothing narrower had been added yet. The *patterns* stay beside their
 * consumers, which is where they are understood; only the spelling and the message skeleton are
 * shared, which is where they were wrong.
 *
 * `pattern` must not carry the `g` flag: a stateful `lastIndex` would make the same string answer
 * differently on a second call.
 *
 * @param why the clause explaining the refusal, appended after the offset.
 */
export function unrepresentable(
  text: string,
  pattern: RegExp,
  why: string,
): AuthoringError | undefined {
  const found = pattern.exec(text);
  if (found === null) return undefined;
  const codePoint = text.codePointAt(found.index) as number;
  return new AuthoringError(
    `cannot write U+${codePointHex(codePoint)} at offset ${found.index}: ${why}`,
  );
}
