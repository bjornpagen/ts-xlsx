// Modern threaded comments in the model, the review-style conversations Excel has written since 2018:
// an anchored discussion of authored messages with timestamps, replies, a resolved state, and
// @mentions. Distinct from a legacy note (`cell.note`), which is a single anonymous annotation; Excel's
// own UI draws the same line, calling these "comments" and those "notes". A cell carries one or the
// other, never both.
//
// The wire form lives in two parts (see `io/xlsx/threaded-comments.ts`): the messages per sheet, and a
// workbook-level identity registry the messages point into. The model mirrors that split: a thread's
// authors and mentioned people are {@link Person} entries resolved through
// {@link Workbook.getPerson}, not names duplicated onto every message.

import {quoted} from '../errors.ts';

/**
 * A registered identity a threaded comment can point at: an author, or someone `@mentioned` in a
 * message. One `<person>` of the workbook's `xl/persons/person.xml` registry.
 *
 * A single human legitimately has **several** entries: Excel registers a mentioned identity separately
 * from that person's authoring identity, with the same {@link displayName} and {@link userId} but a
 * different {@link id} and a different {@link providerId}. The id is therefore the only identity;
 * see {@link Workbook.getPerson}.
 */
export interface Person {
  /** Brace-wrapped GUID this identity is referenced by. The only field that identifies it. */
  readonly id: string;
  /** The name a spreadsheet app shows. Not unique, and not an identity. */
  readonly displayName: string;
  /** Identity-provider handle, `S::<email>::<tenant-guid>` for an AzureAD account. */
  readonly userId?: string;
  /** The provider that registered this entry: `AD` for a directory account, `PeoplePicker` for an
   * identity interned by being mentioned. */
  readonly providerId?: string;
}

/**
 * An `@mention` as the file spells it: who was named, and the run of {@link Comment.text} that
 * renders as the mention chip.
 *
 * The offsets are only meaningful against that exact text: shift either and a spreadsheet app
 * highlights the wrong words.
 *
 * This is the wire shape, shared with the codec that reads it. {@link Mention} is this plus the
 * identity we resolved the id to, which is the one thing the file does not carry.
 */
export interface MentionRef {
  /** The mentioned {@link Person.id} exactly as written, so a dangling mention stays diagnosable. */
  readonly personId: string;
  /** Excel's own id for this mention, preserved so re-emitting it does not invent a new one. */
  readonly mentionId?: string;
  /**
   * 0-based character offset into {@link Comment.text} where the mention starts. Verified against
   * desktop Excel by rendering: the chip covers exactly `[startIndex, startIndex + length)`.
   */
  readonly startIndex: number;
  /** Length of the mention in characters, **counting the leading `@`** (`@Grace Hopper` is 13). */
  readonly length: number;
}

/** A {@link MentionRef} with its identity resolved against the workbook's person registry. */
export interface Mention extends MentionRef {
  /**
   * The mentioned identity, resolved through the workbook registry. Absent when the file names an id
   * the registry does not hold (a mention left dangling by a foreign generator); {@link personId}
   * still says who was meant.
   */
  readonly person?: Person;
}

/** One message of a {@link CommentThread}: what a single person wrote, once. */
export interface Comment {
  /** Brace-wrapped GUID identifying this message, preserved verbatim from the file. */
  readonly id: string;
  /**
   * Who wrote it, resolved through the workbook registry. Absent when the file recorded no author, or
   * named an id the registry does not hold, which Excel treats the same way.
   */
  readonly author?: Person;
  /**
   * The author's {@link Person.id}, which must be registered ({@link Workbook.addPerson}) by the time
   * the workbook is written. Absent for a message with no recorded author, which is written under the
   * null GUID Excel uses for one and shown by Excel as "Author".
   */
  readonly personId?: string;
  /**
   * When it was written, verbatim. Excel writes local wall-clock with fractional seconds and no
   * timezone (`2026-07-24T10:56:41.72`), which is not a round-trippable instant. Keeping the string
   * spares the reader from inventing a zone the file never stated.
   */
  readonly date?: string;
  /** The message body as plain text. Mention chips are part of it; {@link mentions} spans it. */
  readonly text: string;
  /** The `@mentions` this message carries, in document order. Empty for a message that names no one. */
  readonly mentions: readonly Mention[];
}

// A GUID in any of the spellings a caller plausibly has one in: braced or bare, upper or lower case.
// `crypto.randomUUID()` produces the bare lower-case form, so accepting only the canonical spelling would
// reject the one obvious way to make an id in JavaScript.
const GUID = /^\{?([0-9a-f]{8})-([0-9a-f]{4})-([0-9a-f]{4})-([0-9a-f]{4})-([0-9a-f]{12})\}?$/i;

/**
 * The one spelling every threaded-comment identifier must take on the wire: brace-wrapped, upper-case hex.
 *
 * Verified against the OOXML schema, which pins `person/@id`, a message's `id`/`personId`/`parentId`, and a
 * mention's `mentionpersonId`/`mentionId` to exactly `\{[0-9A-F]{8}-…\}`, so a bare GUID and a lower-case one
 * are each rejected outright. So this normalises rather than merely checking: a caller passing
 * `crypto.randomUUID()` gets a valid file instead of one Excel offers to repair.
 *
 * The authoring path alone goes through here. A file's own ids are re-emitted as the file wrote them, since
 * a reader that rewrote them would break every reference pointing at them.
 *
 * @throws {SyntaxError} if the value is not a GUID in any spelling.
 */
export function commentThreadGuid(value: string, what: string): string {
  const match = GUID.exec(value.trim());
  if (match === null) {
    throw new SyntaxError(
      `${what} must be a GUID: Excel writes threaded-comment ids as ` +
        `"{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}", but got ${quoted(value)}`,
    );
  }
  return `{${match.slice(1).join('-').toUpperCase()}}`;
}

/**
 * The largest value a mention offset can take on the wire. Verified against the OOXML schema: both
 * `startIndex` and `length` are `xsd:unsignedInt`, so `4294967295` validates and `4294967296` is rejected
 * as "not a valid 'UInt32' value".
 *
 * Wildly beyond any real message, and that is the point: the ceiling exists so a value from a hostile
 * part can never reach the serialiser. JavaScript spells a large enough number in exponent form
 * (`String(1e21)` is `"1e+21"`), which is not a numeric literal any schema accepts, and one invalid
 * attribute is enough for Excel to offer to repair the whole conversation away.
 */
export const MENTION_OFFSET_MAX = 0xffff_ffff;

/**
 * A mention offset as the wire accepts it: a whole number within {@link MENTION_OFFSET_MAX}.
 *
 * The authoring path alone throws. A file's own mentions are read leniently: one carrying an unusable
 * offset is dropped, keeping the message text and losing only the chip, because a foreign generator's
 * arithmetic is not something a caller can fix, whereas their own is.
 *
 * @throws {SyntaxError} if the value is negative, fractional, or beyond the wire's ceiling.
 */
export function commentThreadOffset(value: number, what: string): number {
  if (!Number.isInteger(value) || value < 0 || value > MENTION_OFFSET_MAX) {
    throw new SyntaxError(
      `${what} must be a whole number between 0 and ${MENTION_OFFSET_MAX}: a mention's span is ` +
        `written as an unsigned 32-bit integer, but got ${value}`,
    );
  }
  return value;
}

/** A conversation anchored to one cell: what was asked, every reply, and whether it was resolved. */
export interface CommentThread {
  /**
   * A1 reference of the single cell the conversation hangs off, canonicalised with no `$` anchors and
   * always a column and a row, so two anchors compare as plain strings and a writer can resolve it without
   * re-validating it.
   */
  readonly ref: string;
  /**
   * Whether the conversation was marked resolved. A property of the *thread*: only the head carries
   * the flag on the wire, so a reply never disagrees with the thread it belongs to.
   */
  readonly resolved: boolean;
  /** The opening message first, then its replies in the order they were written. Never empty. */
  readonly comments: readonly Comment[];
}
