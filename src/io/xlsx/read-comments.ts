// Reading cell comments: the `xl/comments{n}.xml` part mapped back onto its cells as notes, with each
// thread's legacy fallback recognised and suppressed rather than surfaced as one.
//
// The other half of `comments.ts`, which writes the part. The two lived in one module, and the writer
// then loaded the reader's machinery: once a note's body was read through `RunAccumulator`, the
// writer-only `/node` entry went from 409.9 KB in 107 modules to 414.0 KB in 108, which is the read
// path arriving in an entry whose budget exists to say it has not.
//
// A note's body is a `CT_Rst`, the grammar of a cell's `<si>` and `<is>`, which is why it is read by the
// same run machine; see `parseComments`.

import {tryDecodeCellRef} from '../../core/address.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {numInteger} from '../../xml/xml-attrs.ts';
import {parseXml, TextCapture} from '../../xml/xml-read.ts';
import {localName} from '../../xml/xml-scan.ts';
import {RunAccumulator} from './read-rich-runs.ts';

/** One `<comment>` read back from a comments part. */
export interface ParsedComment {
  readonly text: string;
  /**
   * The thread head id this comment is the legacy fallback for, read off its synthetic `tc={headId}`
   * author; absent for a user's own note.
   */
  readonly threadId?: string;
}

// A comment names its author by index into `<authors>`, so an empty entry must still occupy its slot.
// Presenting the self-closing `<author/>` an author-less file writes as an empty element gives it the
// close that pushes it. Without this every later index would shift by one and a note could inherit a
// thread's `tc=` author. `<text/>` is here for the same reason on the other axis: an empty note is a
// note, and a self-closing one used to latch a capture nothing would close.
const COMMENT_EMPTY_CLOSES: ReadonlySet<string> = new Set(['author', 'text']);

// The author string marking a comment as a thread's legacy fallback: `tc={headThreadId}`.
const THREAD_AUTHOR_PREFIX = 'tc=';

/**
 * Parse a `comments{n}.xml` part into a map of A1 reference → comment. Text runs within one comment are
 * concatenated; an author-name run is Excel's own convention and is not stripped, so a note reads back
 * as exactly the text that was written.
 */
export function parseComments(xml: string): Map<string, ParsedComment> {
  const comments = new Map<string, ParsedComment>();
  const authors: string[] = [];
  let currentRef: string | undefined;
  let currentAuthorId: string | undefined;
  // An author name, gathered whole through the shared capture rather than a latch of its own.
  const author = new TextCapture('author');
  // A note's body is a `CT_Rst`, the grammar of a cell's `<si>` and `<is>`, so it is read by the same
  // machine: runs flattened to their text in document order, each `<t>` unescaped whole as a cell's is,
  // and a phonetic run's `<t>` left out. A capture of its own over `<text>` gathered every `<t>` beneath
  // it, furigana included, so a note written `漢字` read back as `漢字かんじ`.
  const bodyRuns = new RunAccumulator({container: 'text', readRuns: false});
  let body = '';
  parseXml(
    xml,
    {
      onOpen(name, attrs, selfClosing) {
        const local = localName(name);
        if (local === 'comment') {
          currentRef = attrs.ref;
          currentAuthorId = attrs.authorId;
          body = '';
        }
        if (!bodyRuns.open(local, attrs, selfClosing)) author.open(local, selfClosing);
      },
      onText(text) {
        bodyRuns.text(text);
        author.text(text);
      },
      onClose(name) {
        const local = localName(name);
        const bodyClose = bodyRuns.close(local);
        if (bodyClose === 'container') {
          body = bodyRuns.plainText;
          return;
        }
        if (bodyClose === 'claimed') return;
        const text = author.close(local);
        if (local === 'author') {
          authors.push(text ?? '');
        } else if (local === 'comment' && currentRef !== undefined) {
          const threadId = threadIdOf(authors[numInteger(currentAuthorId, 0) ?? -1]);
          comments.set(currentRef, {
            text: body,
            ...(threadId !== undefined ? {threadId} : {}),
          });
          currentRef = undefined;
          currentAuthorId = undefined;
        }
      },
    },
    {closeEmptyElements: COMMENT_EMPTY_CLOSES},
  );
  return comments;
}

// A missing or non-numeric `authorId` indexes nothing, so `authors[NaN]` is undefined and the comment
// reads as a plain note: the safe direction, since mistaking a note for a fallback would delete it.
function threadIdOf(author: string | undefined): string | undefined {
  if (author === undefined || !author.startsWith(THREAD_AUTHOR_PREFIX)) return undefined;
  const id = author.slice(THREAD_AUTHOR_PREFIX.length);
  return id === '' ? undefined : id;
}

/**
 * Apply a parsed comments part onto a sheet's cells as notes, addressing each by its A1 reference.
 *
 * A thread's legacy fallback is not a note and does not become one: its text is boilerplate wrapping a
 * copy of the conversation, so surfacing it as `cell.note` hands the caller garbage, and on write it
 * would be re-emitted as a plain note, destroying the `tc=`/`xr:uid` binding and leaving Excel unable to
 * see the thread at all.
 *
 * Suppressed only for a conversation the reader actually holds: a file whose thread part is missing or
 * damaged has nothing else left, so there the boilerplate is kept rather than the content lost. Call
 * after the sheet's threads are restored, since that is what this reads to decide.
 */
export function applyNotes(sheet: Worksheet, comments: ReadonlyMap<string, ParsedComment>): void {
  const headIds = new Set(
    sheet.commentThreads.flatMap((thread) => {
      const head = thread.comments[0];
      return head === undefined ? [] : [head.id];
    }),
  );
  for (const [ref, comment] of comments) {
    if (comment.threadId !== undefined && headIds.has(comment.threadId)) continue;
    // A `ref` naming no cell that can exist costs its note, not the sheet.
    if (tryDecodeCellRef(ref) === undefined) continue;
    sheet.getCell(ref).note = comment.text;
  }
}
