// Cell comments: the `xl/comments{n}.xml` part and its `xl/drawings/vmlDrawing{n}.vml` companion, as the
// writer emits them. Reading the part back onto cells is `read-comments.ts`, kept apart so the writer's
// closure does not load the reader's run machine.
//
// A comment is anchored to a cell by A1 reference and rendered by Excel as a floating box. The box's
// geometry lives in a legacy VML drawing (the pre-DrawingML shape format Excel still requires here);
// the text lives in the comments part. Both are emitted together: a comments part with no matching
// `<legacyDrawing>`/VML reads as text but renders nothing, so we never split them.
//
// Two different things share this one wire form:
//   • a user's **note** (`cell.note`), a single anonymous annotation, the whole of what the part held
//     before 2018;
//   • the legacy **fallback** Excel writes beside every modern threaded comment (see
//     `threaded-comments.ts`), so a pre-2018 reader still sees the conversation. Its text is a fixed
//     boilerplate wrapping a copy of the thread, and its author is a synthetic `tc={headId}` entry.
//
// That `tc=` author and the comment's `xr:uid` are how Excel binds a cell back to its thread, not
// decoration. Verified against desktop Excel: a package whose threadedComment part, persons registry,
// relationships and content types all survive intact still reads back as ordinary notes with zero
// threads once those two are lost. So the fallback is *derived from the thread model* on write and
// *suppressed on read* (`read-comments.ts`), rather than round-tripped as a plain note.

import {tryDecodeCellRef} from '../../core/address.ts';
import type {Cell} from '../../core/cell.ts';
import type {CommentThread} from '../../core/comment-thread.ts';
import {escapeText, textAttr, textElement, XML_DECLARATION} from '../../xml/xml.ts';
import {MARKUP_COMPATIBILITY_NS, REVISION_NS, SPREADSHEETML_NS} from './namespaces.ts';

/** A comment bound for `comments{n}.xml`, paired with the coordinates the VML anchor needs.
 */
export interface CommentCell {
  readonly ref: string;
  /** 1-based row of the commented cell. */
  readonly row: number;
  /** 1-based column of the commented cell. */
  readonly col: number;
  readonly text: string;
  /**
   * The {@link CommentThread} head id this comment is the legacy fallback for, absent for a user's own
   * note. Present means the comment is emitted with a synthetic `tc={id}` author and an `xr:uid`, the
   * pair Excel resolves the cell's thread through.
   */
  readonly threadId?: string;
}

/**
 * Gather every comment a sheet must write: its cells' notes, plus one legacy fallback per conversation
 * in `threads`. A comment anchors to its cell regardless of the cell's value, so a note (or a thread) on
 * an otherwise-empty cell is collected too.
 *
 * `threads` is the conversations the *package* will carry, not simply the ones the sheet holds. The
 * caller decides, because a fallback beside a thread whose `threadedComment` part is missing is worse
 * than no fallback at all: verified against desktop Excel, such a comment shows as neither a thread nor
 * a note, so the text disappears entirely.
 *
 * Ordered by cell, row-major, the way Excel writes the list, so a fallback lands interleaved among the
 * notes rather than appended after them, and the VML shapes follow the same order.
 */
export function collectComments(
  cells: Iterable<Cell>,
  threads: readonly CommentThread[],
  alreadyCollected: readonly CommentCell[] = [],
): CommentCell[] {
  const fallbacks = threadFallbacks(threads);
  const anchored = new Set(fallbacks.map((fallback) => fallback.ref));
  const comments = [...fallbacks, ...alreadyCollected];
  for (const cell of cells) {
    // Excel refuses to put a note and a thread on one cell, so a file carrying both (only a foreign
    // generator or a hand-edit makes one) is written back as the thread alone: two comments on one ref
    // is a shape Excel repairs by dropping both, which would lose the conversation as well as the note.
    if (cell.note === undefined || anchored.has(cell.address)) continue;
    comments.push(noteOn(cell, cell.note));
  }
  // A note gathered from a flushed row cannot know whether a thread was later anchored on the same
  // cell, so the thread's precedence is applied here over the whole merged list rather than only over
  // the live half.
  return comments
    .filter((comment) => comment.threadId !== undefined || !anchored.has(comment.ref))
    .sort((a, b) => a.row - b.row || a.col - b.col);
}

/** The notes among a run of cells, for a writer that must ask before the cells are evicted. */
export function collectNotes(cells: Iterable<Cell>): CommentCell[] {
  const notes: CommentCell[] = [];
  for (const cell of cells) {
    if (cell.note !== undefined) notes.push(noteOn(cell, cell.note));
  }
  return notes;
}

// A cell's own note as the entry both collectors gather. The text is passed in rather than read off
// the cell again, so the narrowing the caller already did is what reaches here.
function noteOn(cell: Cell, text: string): CommentCell {
  return {ref: cell.address, row: cell.row, col: cell.col, text};
}

// One legacy fallback per conversation, keyed to the thread head whose id binds it. A thread with no
// messages has nothing to write and no id to bind by, so it contributes none.
function threadFallbacks(threads: readonly CommentThread[]): CommentCell[] {
  const fallbacks: CommentCell[] = [];
  for (const thread of threads) {
    const head = thread.comments[0];
    const cell = tryDecodeCellRef(thread.ref);
    if (head === undefined || cell === undefined) continue;
    fallbacks.push({
      ref: thread.ref,
      row: cell.row,
      col: cell.col,
      text: fallbackText(thread),
      threadId: head.id,
    });
  }
  return fallbacks;
}

// The boilerplate Excel puts in front of every fallback, captured verbatim from an Excel-authored file.
// It is what a pre-2018 reader shows the user, so it is reproduced exactly rather than paraphrased.
const FALLBACK_PREAMBLE =
  '[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any ' +
  'edits to it will get removed if the file is opened in a newer version of Excel. Learn more: ' +
  'https://go.microsoft.com/fwlink/?linkid=870924\n\n';

// A whole conversation flattened into the one comment a pre-2018 reader can render: the opening message
// under `Comment:`, then each reply under its own `Reply:`, every body indented four spaces. Verified
// against desktop Excel for a thread with three replies: `Reply:` repeats per reply rather than the
// replies being joined under one heading.
function fallbackText(thread: CommentThread): string {
  const [head, ...replies] = thread.comments;
  const body = replies.map((reply) => `\nReply:\n    ${reply.text}`).join('');
  return `${FALLBACK_PREAMBLE}Comment:\n    ${head?.text ?? ''}${body}`;
}

// `xr:uid` lives in the 2014 revision namespace, declared `mc:Ignorable` exactly as Excel declares it so
// a consumer that does not know the prefix skips the attribute instead of rejecting the part.
const REVISION_NS_ATTRS = ` xmlns:mc="${MARKUP_COMPATIBILITY_NS}" mc:Ignorable="xr" xmlns:xr="${REVISION_NS}"`;

/**
 * The `xl/comments{n}.xml` part.
 *
 * Authors are laid out the way Excel lays them out: one synthetic `tc={headId}` entry per threaded
 * conversation first, then a single anonymous author shared by every note (the model carries no note
 * author). Each comment points at its own author by index, and a fallback additionally carries the
 * `xr:uid` naming its thread: the pair that keeps Excel treating the cell as threaded.
 */
export function commentsXml(comments: readonly CommentCell[]): string {
  const authorIdByThreadId = new Map<string, number>();
  for (const {threadId} of comments) {
    if (threadId !== undefined && !authorIdByThreadId.has(threadId)) {
      authorIdByThreadId.set(threadId, authorIdByThreadId.size);
    }
  }
  const noteAuthorId = authorIdByThreadId.size;
  const hasNote = comments.some((comment) => comment.threadId === undefined);
  const authors =
    [...authorIdByThreadId.keys()].map((id) => `<author>tc=${escapeText(id)}</author>`).join('') +
    (hasNote ? '<author></author>' : '');

  const list = comments
    .map((comment) => {
      const {threadId} = comment;
      const authorId = threadId === undefined ? noteAuthorId : authorIdByThreadId.get(threadId);
      const uid = textAttr('xr:uid', threadId);
      return (
        `<comment ref="${comment.ref}" authorId="${authorId}"${uid}>` +
        `<text><r>${textElement(comment.text)}</r></text>` +
        '</comment>'
      );
    })
    .join('');

  return (
    XML_DECLARATION +
    `<comments xmlns="${SPREADSHEETML_NS}"${authorIdByThreadId.size > 0 ? REVISION_NS_ATTRS : ''}>` +
    `<authors>${authors}</authors>` +
    `<commentList>${list}</commentList>` +
    '</comments>'
  );
}

// VML namespaces and the one shape type (a text box) every comment reuses.
const VML_HEADER =
  '<xml xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:x="urn:schemas-microsoft-com:office:excel">' +
  '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>' +
  '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" ' +
  'path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/>' +
  '<v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';

// The first VML shape id, and Excel's own. Shape ids live in a per-sheet id block that
// `<o:idmap data="1">` above selects, and block 1 starts at 1025 (1024 * block + 1): a shape numbered
// below its block's base belongs to a block the drawing never claimed, and Excel discards it.
const FIRST_SHAPE_ID = 1025;

/** The `xl/drawings/vmlDrawing{n}.vml` companion: one hidden text-box shape per comment, in the same
 * order as the comments part. Anchor coordinates place the box a couple of cells down-and-right of its
 * owner; Excel refines them on open, so the values are a sensible starting geometry rather than a
 * pixel-exact layout. A thread's fallback shape is `ObjectType="Note"` like any other: Excel draws the
 * threaded-comment card itself and only needs the shape to exist. */
export function vmlDrawingXml(comments: readonly CommentCell[]): string {
  const shapes = comments
    .map((comment, i) => {
      const row0 = comment.row - 1;
      const col0 = comment.col - 1;
      const anchor = `${col0 + 1}, 15, ${row0}, 2, ${col0 + 3}, 15, ${row0 + 4}, 4`;
      return (
        `<v:shape id="_x0000_s${FIRST_SHAPE_ID + i}" type="#_x0000_t202" ` +
        'style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;' +
        `z-index:${i + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">` +
        '<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/>' +
        '<v:path o:connecttype="none"/>' +
        '<v:textbox style="mso-direction-alt:auto;mso-fit-shape-to-text:t"><div style="text-align:left"></div></v:textbox>' +
        '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>' +
        `<x:Anchor>${anchor}</x:Anchor><x:AutoFill>False</x:AutoFill>` +
        `<x:Row>${row0}</x:Row><x:Column>${col0}</x:Column></x:ClientData></v:shape>`
      );
    })
    .join('');
  return `${VML_HEADER}${shapes}</xml>`;
}
