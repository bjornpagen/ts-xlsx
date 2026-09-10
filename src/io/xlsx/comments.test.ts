import assert from 'node:assert/strict';
import {test} from 'node:test';

import type {CommentThread} from '../../core/comment-thread.ts';
import {INTERNAL} from '../../core/internal.ts';
import {Workbook} from '../../core/workbook.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {
  applyNotes,
  collectComments,
  commentsXml,
  parseComments,
  vmlDrawingXml,
} from './comments.ts';
import {liveCells} from './hyperlinks.ts';
import {
  captureIn,
  partsOf,
  partsWritten,
  partText,
  roundtrip,
  sheetXml,
} from './package.test-support.ts';
import {writeXlsx} from './write.ts';

test('a cell note survives the write/read round-trip', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 'data';
  ws.getCell('A1').note = 'a helpful note';
  const back = roundtrip(wb).getWorksheet('S');
  assert.strictEqual(back?.getCell('A1').value, 'data');
  assert.strictEqual(back?.getCell('A1').note, 'a helpful note');
});

test('a note attaches to an otherwise-empty cell', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('B2').note = 'standalone';
  const back = roundtrip(wb).getWorksheet('S');
  assert.strictEqual(back?.getCell('B2').value, null);
  assert.strictEqual(back?.getCell('B2').note, 'standalone');
});

test('a note stays on its own cell and does not bleed onto neighbours', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 'noted';
  ws.getCell('A1').note = 'only here';
  ws.getCell('A2').value = 'plain';
  const back = roundtrip(wb).getWorksheet('S');
  assert.strictEqual(back?.getCell('A1').note, 'only here');
  assert.strictEqual(back?.getCell('A2').note, undefined);
});

test('note text with markup-significant characters round-trips exactly', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').note = '5 < 6 & "quoted"';
  const back = roundtrip(wb).getWorksheet('S');
  assert.strictEqual(back?.getCell('A1').note, '5 < 6 & "quoted"');
});

test('a note follows its cell when a row is inserted above it', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A2').value = 'body';
  ws.getCell('A2').note = 'travels';
  ws.insertRow(1, ['new header']);
  const back = roundtrip(wb).getWorksheet('S');
  assert.strictEqual(back?.getCell('A3').value, 'body');
  assert.strictEqual(back?.getCell('A3').note, 'travels');
  assert.strictEqual(back?.getCell('A2').note, undefined);
});

test('a noted workbook emits a comments part, a VML drawing, and a legacyDrawing reference', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').note = 'x';
  const pkg = writeXlsx(wb);
  assert.ok(partText(pkg, 'xl/comments1.xml'), 'a comments part is written');
  assert.ok(partText(pkg, 'xl/drawings/vmlDrawing1.vml'), 'a VML drawing companion is written');
  assert.match(sheetXml(pkg), /<legacyDrawing r:id="[^"]+"\/>/);
  const contentTypes = partText(pkg, '[Content_Types].xml');
  assert.match(contentTypes, /Extension="vml"/);
  assert.match(contentTypes, /PartName="\/xl\/comments1\.xml"/);
});

test('a note textbox auto-fits its text so a multi-line note is not clipped', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('B2').note = 'line one\nline two\nline three';
  const vml = partText(writeXlsx(wb), 'xl/drawings/vmlDrawing1.vml');
  const style = captureIn(vml, /<v:textbox\b[^>]*\bstyle="([^"]*)"/, 'the textbox style');
  assert.match(style, /mso-fit-shape-to-text:t/, 'the textbox grows to fit its content');
});

test('a note-free workbook writes no comment or VML parts', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').value = 'plain';
  const pkg = writeXlsx(wb);
  const names = Object.keys(partsOf(pkg));
  assert.ok(!names.some((n) => /comments\d+\.xml$/.test(n)));
  assert.ok(!names.some((n) => n.endsWith('.vml')));
  assert.ok(!partText(pkg, '[Content_Types].xml').includes('Extension="vml"'));
});

// The legacy fallback Excel writes beside every modern threaded comment. Excel binds a cell to its
// thread through that comment's synthetic `tc={headId}` author and its `xr:uid`, so the fallback is
// derived from the thread model on write and suppressed on read rather than treated as a note.

const HEAD = '{A6473DA8-237C-4E27-ADE9-48D3A7CD15A7}';
const OTHER_HEAD = '{08B3B787-8F24-49B6-8D78-A5F335591EEA}';

// The boilerplate prefix verbatim from an Excel-authored file, up to the blank line before `Comment:`.
const PREAMBLE =
  '[Threaded comment]\n\nYour version of Excel allows you to read this threaded comment; however, any ' +
  'edits to it will get removed if the file is opened in a newer version of Excel. Learn more: ' +
  'https://go.microsoft.com/fwlink/?linkid=870924';

const threadOn = (
  ref: string,
  texts: readonly string[],
  head = HEAD,
  resolved = false,
): CommentThread => ({
  ref,
  resolved,
  comments: texts.map((text, i) => ({id: i === 0 ? head : `{reply-${i}}`, text, mentions: []})),
});

// The conversations the package will carry, which is what the writer passes: always the sheet's own, since
// the writer serialises the `threadedComment` part from the very same list.
const shadowing = (sheet: Worksheet) => collectComments(liveCells(sheet), sheet.commentThreads);

test('a thread is written as a fallback comment bound to its head by a tc= author and an xr:uid', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws[INTERNAL].restoreCommentThreads([threadOn('B1', ['Is this gross or net of tax?'])]);
  const xml = commentsXml(shadowing(ws));
  assert.ok(xml.includes(`<authors><author>tc=${HEAD}</author></authors>`));
  assert.ok(xml.includes(`<comment ref="B1" authorId="0" xr:uid="${HEAD}">`));
  assert.match(
    xml,
    /xmlns:xr="http:\/\/schemas\.microsoft\.com\/office\/spreadsheetml\/2014\/revision"/,
  );
  assert.match(xml, /mc:Ignorable="xr"/, 'the prefix is declared ignorable, as Excel declares it');
});

test('the fallback text is the boilerplate, the opening message, then one Reply: per reply', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws[INTERNAL].restoreCommentThreads([
    threadOn('A1', ['Head question?', 'First reply.', 'Second reply.']),
  ]);
  const [fallback] = shadowing(ws);
  // Verified against desktop Excel for a thread with three replies: `Reply:` repeats per reply rather
  // than the replies being joined under one heading, and every body is indented four spaces.
  assert.strictEqual(
    fallback?.text,
    `${PREAMBLE}\n\nComment:\n    Head question?\nReply:\n    First reply.\nReply:\n    Second reply.`,
  );
});

test('a conversation with no note beside it still gets a comments part and a VML shape of its own', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws[INTERNAL].restoreCommentThreads([threadOn('B1', ['a conversation'])]);
  const comments = shadowing(ws);
  assert.strictEqual(comments.length, 1, 'the fallback is the whole comments part');
  assert.strictEqual(
    (vmlDrawingXml(comments).match(/<v:shape\b/g) ?? []).length,
    1,
    'and it has a box to render into: a comment with no shape reads as text but draws nothing',
  );
});

test('a conversation is never written half-emitted: the fallback and its thread part come together', () => {
  // Verified against desktop Excel, in both directions: a `tc=` fallback whose `threadedComment` part is
  // absent is shown as neither a thread NOR a note (the text simply disappears), and a thread part whose
  // fallback is absent is ignored (the cell reads blank). The two are
  // halves of one representation, so the writer derives both from the same list and a package can never
  // carry one alone.
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('B1').value = 12;
  ws.addCommentThread(threadOn('B1', ['Is this gross or net of tax?']));
  const names = Object.keys(partsWritten(wb));
  assert.ok(names.includes('xl/threadedComments/threadedComment1.xml'), 'the conversation itself');
  assert.ok(names.includes('xl/comments1.xml'), 'the fallback a pre-2018 reader renders');
  assert.ok(names.includes('xl/drawings/vmlDrawing1.vml'), 'and a box for it to render into');
});

test('each sheet numbers its own thread part, leaving a gap where a sheet has no conversation', () => {
  // Part numbers follow the sheet index, so a workbook whose first sheet has only notes writes
  // `threadedComment2.xml` with no `threadedComment1.xml` beside it. That gap is fine: nothing addresses
  // these parts by number; each sheet reaches its own through a relationship. Verified with desktop Excel,
  // which reads all three sheets back correctly from exactly this shape.
  const wb = new Workbook();
  wb.addPerson({id: HEAD, displayName: 'Ada Lovelace', providerId: 'AD'});
  wb.addWorksheet('NotesOnly').getCell('A1').note = 'just a note';
  wb.addWorksheet('Threads').addCommentThread(threadOn('C3', ['second sheet'], HEAD));
  const both = wb.addWorksheet('Both');
  both.getCell('A1').note = 'note here';
  both.addCommentThread(threadOn('B2', ['third sheet'], OTHER_HEAD));
  const names = Object.keys(partsWritten(wb));
  assert.deepStrictEqual(names.filter((n) => n.startsWith('xl/threadedComments/')).sort(), [
    'xl/threadedComments/threadedComment2.xml',
    'xl/threadedComments/threadedComment3.xml',
  ]);
  assert.deepStrictEqual(
    names.filter((n) => /^xl\/comments\d+\.xml$/.test(n)).sort(),
    ['xl/comments1.xml', 'xl/comments2.xml', 'xl/comments3.xml'],
    'the notes-only sheet still gets a comments part; the thread sheets get theirs for the fallback',
  );
  assert.ok(
    names.includes('xl/persons/person.xml'),
    'one registry for the whole workbook, however many sheets carry conversations',
  );
});

test('notes and fallbacks share one comments part, ordered by cell, each pointing at its own author', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').note = 'a real note';
  ws.getCell('D4').note = 'another real note';
  ws[INTERNAL].restoreCommentThreads([threadOn('B2', ['head'])]);
  const xml = commentsXml(shadowing(ws));
  assert.deepStrictEqual(
    [...xml.matchAll(/<comment ref="([^"]*)" authorId="(\d+)"/g)].map((m) => [m[1], m[2]]),
    [
      ['A1', '1'],
      ['B2', '0'],
      ['D4', '1'],
    ],
    'row-major order, the fallback on its tc= author and both notes on the shared anonymous one',
  );
});

test('a note-only sheet keeps the minimal comments root, declaring no namespace it never uses', () => {
  const wb = new Workbook();
  const ws = wb.addWorksheet('S');
  ws.getCell('A1').note = 'plain';
  const xml = partText(writeXlsx(wb), 'xl/comments1.xml');
  assert.ok(!xml.includes('xmlns:xr'), 'no thread means no xr:uid, so no xr declaration');
  assert.ok(!xml.includes('tc='), 'and no synthetic thread author');
});

test('a threaded cell reads back with no note, while a genuine note beside it is untouched', () => {
  const ws = new Workbook().addWorksheet('S');
  ws[INTERNAL].restoreCommentThreads([
    threadOn('B1', ['Is this gross or net of tax?', 'Gross.'], HEAD, true),
  ]);
  applyNotes(
    ws,
    new Map([
      ['B1', {text: `${PREAMBLE}\n\nComment:\n    Is this gross or net of tax?`, threadId: HEAD}],
      ['D4', {text: 'a real note'}],
    ]),
  );
  assert.strictEqual(ws.getCell('B1').note, undefined, 'the boilerplate is not an annotation');
  assert.strictEqual(ws.getCell('D4').note, 'a real note');
});

test('a fallback whose thread is missing is kept as a note rather than silently dropped', () => {
  // Only a damaged or foreign file reaches here: with the thread part gone, the boilerplate is the last
  // remaining record of the conversation, so losing structure beats losing content.
  const ws = new Workbook().addWorksheet('S');
  applyNotes(
    ws,
    new Map([['B1', {text: `${PREAMBLE}\n\nComment:\n    orphaned`, threadId: HEAD}]]),
  );
  assert.match(ws.getCell('B1').note ?? '', /orphaned$/);
});

test('a fallback is suppressed only for the thread it actually names', () => {
  const ws = new Workbook().addWorksheet('S');
  ws[INTERNAL].restoreCommentThreads([threadOn('B1', ['live thread'])]);
  applyNotes(
    ws,
    new Map([
      ['B1', {text: 'shadow of a thread we hold', threadId: HEAD}],
      ['B2', {text: 'shadow of a thread we do not', threadId: OTHER_HEAD}],
    ]),
  );
  assert.strictEqual(ws.getCell('B1').note, undefined);
  assert.strictEqual(ws.getCell('B2').note, 'shadow of a thread we do not');
});

test('an empty self-closing author still occupies its index, so a note keeps its own author', () => {
  // A comment names its author by position in `<authors>`. If a self-closing `<author/>` were skipped,
  // every later index would shift by one: the note below would inherit the thread's `tc=` author and
  // then be deleted as a fallback.
  const parsed = parseComments(
    '<comments><authors><author/><author>tc={T}</author></authors><commentList>' +
      '<comment ref="A1" authorId="0"><text><t>a real note</t></text></comment>' +
      '<comment ref="A2" authorId="1"><text><t>a fallback</t></text></comment>' +
      '</commentList></comments>',
  );
  assert.deepStrictEqual(parsed.get('A1'), {text: 'a real note'});
  assert.deepStrictEqual(parsed.get('A2'), {text: 'a fallback', threadId: '{T}'});
});

test('a comment with no resolvable author reads as a plain note, never as a fallback', () => {
  // Mistaking a note for a fallback deletes it, so an unusable authorId must fail in the safe direction.
  const parsed = parseComments(
    '<comments><authors><author>tc={T}</author></authors><commentList>' +
      '<comment ref="A1"><text><t>no authorId</t></text></comment>' +
      '<comment ref="A2" authorId="9"><text><t>out of range</t></text></comment>' +
      '<comment ref="A3" authorId="oops"><text><t>not a number</t></text></comment>' +
      '</commentList></comments>',
  );
  assert.deepStrictEqual(
    [...parsed.values()].map((comment) => comment.threadId),
    [undefined, undefined, undefined],
  );
});
