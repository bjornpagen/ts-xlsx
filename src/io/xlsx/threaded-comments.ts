// Modern threaded comments: the review-style conversations Excel has written since 2018, read from and
// written to their two parts: `xl/threadedComments/threadedComment{n}.xml` (per sheet) and
// `xl/persons/person.xml` (per workbook, the author registry).
//
// These are a Microsoft extension, not base ECMA-376, and they are a *separate* feature from legacy
// notes (`comments{n}.xml`, see `comments.ts`) rather than a newer spelling of them. A cell carries
// one or the other, never both: Excel refuses to add a note to a threaded cell and vice versa.
//
// A message identifies its author by `personId` into the registry and its thread by `id`/`parentId`:
// the first message of a thread has no `parentId`, every reply carries the head's `id`. Only the head
// carries `done`, so a thread's resolved state is its head's; a reply never says.
//
// `personId` is required, and it has to name a registered person: Excel 16.0 offers to repair a package
// whose message leaves it out or names an id the registry does not hold. Its own repair of the second
// rewrites the id to the null GUID, which it then opens clean with or without a registry and shows as
// "Author". That is the one spelling of "no recorded author" the wire has, so the reader reads an
// author it cannot resolve as no author, and the writer writes no author as the null GUID.
//
// These parsers describe a file that already exists, so they read leniently: an unrecognised or
// missing attribute yields a sensible default rather than a throw, and a message too incomplete to
// anchor is skipped instead of crashing the read. Optional wire attributes stay optional in the
// parsed shape rather than collapsing to `''`, so "the file did not say" never masquerades as a value
// the file contained.
//
// {@link buildCommentThreads} turns the flat message list into the model's threads, resolving each
// author and mention against the workbook's person registry. That grouping is deliberately not the
// parsers' job: they stay faithful to the part, one function per wire form.
//
// The writers are the exact inverse and carry no clock and no id generator: every guid and every
// timestamp is written from the model, verbatim, so the same workbook always serialises to the same
// bytes. A conversation Excel wrote round-trips as itself; one authored in the model carries whatever
// ids and dates the caller supplied.
//
// A message's `<text>` carries the `_xHHHH_` convention, so it is escaped on the way out and decoded
// on the way in. Nothing in the 2018 schema says so; it is a measurement. Excel Desktop reads
// `_x0041_` in a `<text>` back as `A`, leaves `_xZZZZ_`/`_x041_` alone, resolves `_x005F_x0041_` in
// one left-to-right pass to the literal text, and writes a control character back out as `_x0001_`
// when it re-saves. See `docs/knowledge/specs/spreadsheetml-xhhhh-escape-is-decoded-on-read.md`. The
// practical consequence is that the same string now behaves identically in both comment systems,
// where a threaded comment used to refuse what a legacy note accepted.

import {encodeAddress, tryDecodeCellRef} from '../../core/address.ts';
import {
  type Comment,
  type CommentThread,
  MENTION_OFFSET_MAX,
  type Mention,
  type MentionRef,
  type Person,
} from '../../core/comment-thread.ts';
import {AuthoringError, quoted} from '../../errors.ts';
import {decodeSpreadsheetText, numInteger} from '../../xml/xml-attrs.ts';
import {parseXml, TextCapture} from '../../xml/xml-read.ts';
import {boolStrict, localName, type XmlAttributes} from '../../xml/xml-scan.ts';
import {escapeAttr, escapeSpreadsheetText, textAttr, XML_DECLARATION} from '../../xml/xml.ts';
import {THREADED_COMMENTS_NS} from './namespaces.ts';

/**
 * A registered author of threaded comments: one `<person>` of `xl/persons/person.xml`, which is the
 * model's own {@link Person} with nothing added. Declared as an alias for the same reason
 * {@link ParsedMention} is: this file's write half already emits `Person` directly, so re-declaring
 * the four fields here made one feature module name one shape two ways, with nothing keeping the two
 * in step.
 */
export type ParsedPerson = Person;

/**
 * One `<mention>` of a message's `<mentions>` block: the model's own {@link MentionRef}, which is
 * exactly the four wire attributes with nothing resolved yet. Declaring it a second time here would
 * be two places for the offset convention to drift, and the offsets are what a spreadsheet app
 * highlights with.
 *
 * All four attributes are required (Excel rejects a file missing any), and note the lowercase `p`
 * in `mentionpersonId`; the capitalised spelling is not a declared attribute. The person id is that
 * attribute's value; every other field is named as the model names it.
 */
export type ParsedMention = MentionRef;

/** One message of a threaded conversation: a `<threadedComment>` of a `threadedComment{n}.xml`. */
export interface ParsedThreadedComment {
  /** A1 reference of the cell the whole thread anchors to; every message of a thread repeats it. */
  readonly ref: string;
  /** Brace-wrapped GUID identifying this message, and the `parentId` its replies carry. */
  readonly id: string;
  /** The author's {@link ParsedPerson.id}. Absent in a file that recorded no author. */
  readonly personId?: string;
  /**
   * The `dT` timestamp verbatim. Excel writes local wall-clock with fractional seconds and no
   * timezone (`2026-07-24T10:56:41.72`), which is not a round-trippable instant. Keeping the
   * string spares the reader from inventing a zone the file never stated.
   */
  readonly date?: string;
  readonly text: string;
  /** The thread head's {@link ParsedThreadedComment.id}; absent iff this message *is* the head. */
  readonly parentId?: string;
  /**
   * The `done` flag exactly as written. Excel puts it on the head alone and omits it entirely on an
   * open thread (never `done="0"`), so read a thread's resolved state off its head: a reply's
   * `false` here means "did not say", not "not resolved".
   */
  readonly done: boolean;
  /** The `<mentions>` this message carries, in document order; empty when it names no one. */
  readonly mentions: readonly ParsedMention[];
}

/**
 * Parse `xl/persons/person.xml` into its registered authors, in document order. Order carries no
 * meaning, because Excel re-sorts the list by person id when it saves, so nothing may depend on it. An
 * entry without an `id` is skipped: no message could reference it.
 */
export function parsePersons(xml: string): ParsedPerson[] {
  const persons: ParsedPerson[] = [];
  parseXml(xml, {
    onOpen(name, attrs) {
      // `personList` shares the prefix, so match the exact local name rather than a `startsWith`.
      if (localName(name) !== 'person') return;
      const {id, displayName, userId, providerId} = attrs;
      if (id === undefined) return;
      persons.push({
        id,
        displayName: displayName ?? '',
        ...(userId !== undefined ? {userId} : {}),
        ...(providerId !== undefined ? {providerId} : {}),
      });
    },
  });
  return persons;
}

// Committing a message from its close handling alone would silently drop `<threadedComment …/>`,
// which fires no close of its own; presenting the self-closing form as an empty element gives it one.
const THREADED_COMMENT_EMPTY_CLOSES: ReadonlySet<string> = new Set(['threadedComment']);

/**
 * Parse a `threadedComment{n}.xml` part into its messages, in document order: thread order, with
 * each thread's replies following its head. Grouping into threads is {@link buildCommentThreads}'s
 * job; this stays faithful to the part. A message without a `ref` or `id` cannot be anchored or
 * replied to and is skipped.
 */
export function parseThreadedComments(xml: string): ParsedThreadedComment[] {
  const messages: ParsedThreadedComment[] = [];
  let open: XmlAttributes | undefined;
  let mentions: ParsedMention[] = [];
  let text = '';
  const capture = new TextCapture('text');
  parseXml(
    xml,
    {
      onOpen(name, attrs, selfClosing) {
        const local = localName(name);
        if (local === 'threadedComment') {
          open = attrs;
          text = '';
          mentions = [];
        } else if (open !== undefined) {
          capture.open(local, selfClosing);
          if (local === 'mention') {
            const mention = mentionFrom(attrs);
            if (mention !== undefined) mentions.push(mention);
          }
        }
      },
      onText(chunk) {
        capture.text(chunk);
      },
      onClose(name) {
        const local = localName(name);
        const body = capture.close(local);
        if (body !== undefined) {
          text = body;
        } else if (local === 'threadedComment') {
          if (open !== undefined) {
            const message = threadedCommentFrom(open, text, mentions);
            if (message !== undefined) messages.push(message);
          }
          open = undefined;
          text = '';
          mentions = [];
        }
      },
    },
    {closeEmptyElements: THREADED_COMMENT_EMPTY_CLOSES},
  );
  return messages;
}

function threadedCommentFrom(
  attrs: XmlAttributes,
  text: string,
  mentions: readonly ParsedMention[],
): ParsedThreadedComment | undefined {
  const {ref, id, personId, dT, parentId} = attrs;
  if (ref === undefined || id === undefined) return undefined;
  return {
    ref,
    id,
    // Decoded once the whole body is in hand: an escape may straddle two text events, since the
    // parser reports a chunk per entity, and half of `_x0041_` decodes to nothing.
    text: decodeSpreadsheetText(text),
    done: boolStrict(attrs.done),
    mentions: [...mentions],
    ...(personId !== undefined ? {personId} : {}),
    ...(dT !== undefined ? {date: dT} : {}),
    ...(parentId !== undefined ? {parentId} : {}),
  };
}

// A mention without a target person or a usable span cannot be resolved or rendered, so it is dropped
// rather than carried as a mention over nothing: a `length` of 0 would be an invisible chip, and a
// negative or non-numeric offset would place it outside the text it is supposed to cover.
//
// The upper bound is what makes this a hostile-input guard rather than a tidiness check: an offset the
// wire cannot express would be re-emitted by the writer as an invalid attribute, and JavaScript spells a
// large enough number in exponent form (`1e+21`) which no schema accepts at all, costing the reader's
// leniency the entire conversation when Excel repairs the part. Dropping the chip costs a highlight.
function mentionFrom(attrs: XmlAttributes): ParsedMention | undefined {
  const personId = attrs.mentionpersonId;
  const startIndex = numInteger(attrs.startIndex);
  const length = numInteger(attrs.length);
  if (personId === undefined || startIndex === undefined || length === undefined) return undefined;
  // A negative offset points outside the text and a zero length spans nothing, so neither could render.
  if (startIndex < 0 || length <= 0) return undefined;
  if (startIndex > MENTION_OFFSET_MAX || length > MENTION_OFFSET_MAX) return undefined;
  return {
    personId,
    startIndex,
    length,
    ...(attrs.mentionId !== undefined ? {mentionId: attrs.mentionId} : {}),
  };
}

/**
 * Group a part's messages into the model's threads, resolving every author and mention against the
 * workbook's person registry (`personById`, typically `Workbook.getPerson`).
 *
 * Document order is thread order with each thread's replies following its head, so one pass suffices: a
 * message with no `parentId` opens a thread, and a reply joins the thread its `parentId` names. A reply
 * whose parent is unknown (a dangling `parentId` no Excel file produces) opens a thread of its own
 * rather than being dropped, so a foreign generator's damage costs structure, never content.
 */
export function buildCommentThreads(
  messages: readonly ParsedThreadedComment[],
  personById: (id: string) => Person | undefined,
): CommentThread[] {
  const threads: CommentThread[] = [];
  // The head's own message list, so a reply appends to the thread already published in `threads`.
  const commentsByHeadId = new Map<string, Comment[]>();
  for (const message of messages) {
    const comment = commentFrom(message, personById);
    const siblings =
      message.parentId === undefined ? undefined : commentsByHeadId.get(message.parentId);
    if (siblings !== undefined) {
      siblings.push(comment);
      continue;
    }
    const ref = anchorRef(message.ref);
    if (ref === undefined) continue;
    const comments = [comment];
    commentsByHeadId.set(message.id, comments);
    // Resolved is the head's flag: a reply never carries `done`, so it never contradicts its thread.
    threads.push({ref, resolved: message.done, comments});
  }
  return threads;
}

// The canonical A1 form of a thread's anchor, or undefined when the file wrote something that cannot
// anchor one: a range, a bare row or column, or outright garbage. Canonicalising here is what lets
// every later consumer compare anchors as plain strings (`$B$2` and `B2` are one cell) and keeps a
// foreign file's malformed reference out of the writer, which anchors the thread's legacy fallback by it.
function anchorRef(reference: string): string | undefined {
  const cell = tryDecodeCellRef(reference);
  return cell === undefined ? undefined : encodeAddress(cell.col, cell.row);
}

// An author id the registry does not hold is not kept: written back it would be the package Excel repairs,
// and Excel's repair is to forget who it was.
function commentFrom(
  message: ParsedThreadedComment,
  personById: (id: string) => Person | undefined,
): Comment {
  const author = message.personId === undefined ? undefined : personById(message.personId);
  return {
    id: message.id,
    text: message.text,
    mentions: message.mentions.map((mention) => mentionOf(mention, personById)),
    ...(author !== undefined ? {author, personId: author.id} : {}),
    ...(message.date !== undefined ? {date: message.date} : {}),
  };
}

function mentionOf(
  mention: ParsedMention,
  personById: (id: string) => Person | undefined,
): Mention {
  const person = personById(mention.personId);
  return {
    personId: mention.personId,
    startIndex: mention.startIndex,
    length: mention.length,
    ...(person !== undefined ? {person} : {}),
    ...(mention.mentionId !== undefined ? {mentionId: mention.mentionId} : {}),
  };
}

/**
 * Serialise one sheet's conversations into its `xl/threadedComments/threadedComment{n}.xml` part.
 *
 * Messages are written flat, in thread order with each thread's replies after its head: the shape
 * {@link parseThreadedComments} reads back. The head/reply distinction the model holds as array position
 * becomes `parentId` on every reply but the head, and `done="1"` goes on the head alone: only the head
 * carries the flag on the wire, so a reply can never contradict the thread it belongs to. An open thread
 * omits `done` entirely rather than writing `done="0"`, exactly as Excel does.
 *
 * A thread with no messages writes nothing: it has neither text to say nor a head id for its replies and
 * its legacy fallback to hang off.
 *
 * A message with no author is written under the null GUID, Excel's own spelling for one. `personById` is
 * the workbook's registry, and a message naming an id it does not hold is refused.
 *
 * @throws {AuthoringError} if a message's author is not a registered person.
 */
export function threadedCommentsXml(
  threads: readonly CommentThread[],
  personById: (id: string) => Person | undefined,
): string {
  const messages = threads.flatMap((thread) => {
    const [head, ...replies] = thread.comments;
    if (head === undefined) return [];
    return [
      threadedCommentXml(thread.ref, head, thread.resolved ? ' done="1"' : '', personById),
      ...replies.map((reply) =>
        threadedCommentXml(thread.ref, reply, ` parentId="${escapeAttr(head.id)}"`, personById),
      ),
    ];
  });
  return (
    XML_DECLARATION +
    `<ThreadedComments xmlns="${THREADED_COMMENTS_NS}">${messages.join('')}</ThreadedComments>`
  );
}

// One `<threadedComment>`. `tail` is the attribute that distinguishes the message's role: `done` for a
// resolved head, `parentId` for a reply, nothing for an open head. A `dT` the model never held is
// omitted rather than written empty, so "the file did not say" stays distinguishable from "the file said
// nothing". `personId` is required, so an absent author is written as the null GUID instead. Every value is escaped: an authored message's text and a foreign file's ids alike
// are untrusted, and an unescaped `"` would end the attribute and reshape the part.
//
// The body gets the stronger `escapeSpreadsheetText`, which the ids and the timestamp do not: it is the
// one field here that carries what a human typed, so it is the one that can hold a character XML has no
// syntax for. An id or a `dT` that held one is malformed input, and refusing it is the honest answer.
function threadedCommentXml(
  ref: string,
  comment: Comment,
  tail: string,
  personById: (id: string) => Person | undefined,
): string {
  const date = textAttr('dT', comment.date);
  const personId = comment.personId ?? UNKNOWN_AUTHOR;
  if (personId !== UNKNOWN_AUTHOR && personById(personId) === undefined) {
    throw new AuthoringError(
      `cannot write the threaded comment ${quoted(comment.id)}: its author ${quoted(personId)} is not ` +
        'a registered person. Register it with Workbook.addPerson, or leave personId out for a ' +
        'comment with no recorded author',
    );
  }
  const person = textAttr('personId', personId);
  return (
    `<threadedComment ref="${escapeAttr(ref)}"${date}${person} id="${escapeAttr(comment.id)}"${tail}>` +
    `<text>${escapeSpreadsheetText(comment.text)}</text>` +
    mentionsXml(comment.mentions) +
    '</threadedComment>'
  );
}

// The null GUID: Excel's `personId` for a message whose author it does not know.
const UNKNOWN_AUTHOR = '{00000000-0000-0000-0000-000000000000}';

// The `<mentions>` block, which follows `<text>` in the message. All four `<mention>` attributes are
// required, verified by dropping each in turn and getting `Sch_MissRequiredAttribute`, so a mention the
// model holds without a `mentionId`, or with a span the wire cannot express, cannot be written at all. It
// is dropped rather than given an invented id or a clamped span: the `@name` stays in the text and only
// the chip is lost, whereas an invalid part risks Excel repairing the whole conversation away.
//
// The bounds check is deliberately here as well as in the reader. Both the authoring verb and the parser
// already refuse an out-of-range offset, so nothing should reach this. But `restoreCommentThreads` takes
// a model wholesale, and a serialiser that *cannot* emit `length="1e+21"` beats one that merely is not
// currently handed one.
function mentionsXml(mentions: readonly Mention[]): string {
  const entries = mentions.flatMap(({personId, mentionId, startIndex, length}) => {
    if (mentionId === undefined) return [];
    // A zero-length span is in range but renders nothing, so it is no more writable than one out of range.
    if (!writableOffset(startIndex) || !writableOffset(length) || length === 0) return [];
    return [
      `<mention mentionpersonId="${escapeAttr(personId)}" mentionId="${escapeAttr(mentionId)}"` +
        ` startIndex="${startIndex}" length="${length}"/>`,
    ];
  });
  return entries.length === 0 ? '' : `<mentions>${entries.join('')}</mentions>`;
}

// An offset the wire can express: a whole number within the schema's UInt32 ceiling.
function writableOffset(value: number): boolean {
  return Number.isInteger(value) && value >= 0 && value <= MENTION_OFFSET_MAX;
}

/**
 * Serialise the workbook's identity registry into `xl/persons/person.xml`: singular and unnumbered,
 * unlike the per-sheet thread parts.
 *
 * Entries are written in registry order, which carries no meaning: Excel re-sorts the list by person id
 * whenever it saves, so this only has to be deterministic, not canonical. `userId`/`providerId` are
 * written when the model holds them; a registry read from a file holds whatever that file stated.
 */
export function personsXml(persons: readonly Person[]): string {
  const entries = persons.map((person) => {
    const userId = textAttr('userId', person.userId);
    const providerId = textAttr('providerId', person.providerId);
    return (
      `<person displayName="${escapeAttr(person.displayName)}"` +
      ` id="${escapeAttr(person.id)}"${userId}${providerId}/>`
    );
  });
  return `${XML_DECLARATION}<personList xmlns="${THREADED_COMMENTS_NS}">${entries.join('')}</personList>`;
}
