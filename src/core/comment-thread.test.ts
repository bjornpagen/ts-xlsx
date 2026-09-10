import assert from 'node:assert/strict';
import {test} from 'node:test';

import {threadAt} from './comment-thread.test-support.ts';
import type {CommentThread, Person} from './comment-thread.ts';
import {INTERNAL} from './internal.ts';
import {Workbook} from './workbook.ts';

// The two entries Excel writes for one human: an authoring identity, and the separate one it interns
// when that person is @mentioned. Same name, same userId, different id and provider.
const GRACE_AUTHOR: Person = {
  id: '{1B2C3D4E-5F60-4A71-8B92-0C1D2E3F4A5B}',
  displayName: 'Grace Hopper',
  userId: 'S::grace@example.com::00000000-0000-0000-0000-000000000000',
  providerId: 'AD',
};
const GRACE_MENTIONED: Person = {
  id: '{BA397017-DD76-4496-AA75-59ADB199950C}',
  displayName: 'Grace Hopper',
  userId: 'S::grace@example.com::00000000-0000-0000-0000-000000000000',
  providerId: 'PeoplePicker',
};

test('the person registry keeps both entries of one human, since only the id identifies one', () => {
  const wb = new Workbook();
  wb[INTERNAL].restorePersons([GRACE_AUTHOR, GRACE_MENTIONED]);
  assert.equal(wb.persons.length, 2);
  assert.equal(wb.getPerson(GRACE_AUTHOR.id)?.providerId, 'AD');
  assert.equal(wb.getPerson(GRACE_MENTIONED.id)?.providerId, 'PeoplePicker');
});

test('registering the same id twice replaces the entry, since the id is the identity', () => {
  const wb = new Workbook();
  wb.addPerson(GRACE_AUTHOR);
  wb.addPerson({...GRACE_AUTHOR, displayName: 'Rear Admiral Grace Hopper'});
  assert.equal(wb.persons.length, 1);
  assert.equal(wb.getPerson(GRACE_AUTHOR.id)?.displayName, 'Rear Admiral Grace Hopper');
});

test('registering one human under two ids keeps both, which is what Excel itself writes', () => {
  const wb = new Workbook();
  wb.addPerson(GRACE_AUTHOR);
  wb.addPerson(GRACE_MENTIONED);
  assert.deepEqual(
    wb.persons.map((person) => person.providerId),
    ['AD', 'PeoplePicker'],
  );
});

test('an id the registry does not hold resolves to nothing rather than a near match', () => {
  const wb = new Workbook();
  wb[INTERNAL].restorePersons([GRACE_AUTHOR]);
  assert.equal(wb.getPerson('{NOT-REGISTERED}'), undefined);
});

test('a workbook with no threaded comments has an empty person registry', () => {
  assert.deepEqual(new Workbook().persons, []);
});

test('restoring the registry replaces it, never accumulating a stale identity', () => {
  const wb = new Workbook();
  wb[INTERNAL].restorePersons([GRACE_AUTHOR]);
  wb[INTERNAL].restorePersons([GRACE_MENTIONED]);
  assert.deepEqual(
    wb.persons.map((person) => person.id),
    [GRACE_MENTIONED.id],
  );
  assert.equal(wb.getPerson(GRACE_AUTHOR.id), undefined);
});

test('a sheet exposes its threads in the order they were read', () => {
  const sheet = new Workbook().addWorksheet('S');
  sheet[INTERNAL].restoreCommentThreads([threadAt('B1'), threadAt('B2')]);
  assert.deepEqual(
    sheet.commentThreads.map((thread) => thread.ref),
    ['B1', 'B2'],
  );
});

test('a sheet with no threaded comments has no threads', () => {
  assert.deepEqual(new Workbook().addWorksheet('S').commentThreads, []);
});

test('a thread is found by its anchor, absolute reference or not', () => {
  const sheet = new Workbook().addWorksheet('S');
  sheet[INTERNAL].restoreCommentThreads([threadAt('B2')]);
  assert.equal(sheet.commentThreadAt('B2')?.ref, 'B2');
  assert.equal(sheet.commentThreadAt('$B$2')?.ref, 'B2');
  assert.equal(sheet.commentThreadAt('C3'), undefined);
});

test('looking for a thread does not materialise the cell it asks about', () => {
  // A read-only query that grew the grid would change what the writer emits: a phantom empty cell
  // appearing merely because something inspected the sheet.
  const sheet = new Workbook().addWorksheet('S');
  sheet.commentThreadAt('B2');
  assert.equal(sheet.hasCell(2, 2), false);
});

test('commentThreadAt rejects a reference that is not a single cell', () => {
  const sheet = new Workbook().addWorksheet('S');
  assert.throws(() => sheet.commentThreadAt('A'), /not a single-cell reference/);
  assert.throws(() => sheet.commentThreadAt('2'), /not a single-cell reference/);
});

test('an authored thread keeps the conversation but canonicalises its anchor', () => {
  // The anchor is compared as a plain string by both `commentThreadAt` and the writer's fallback comment,
  // so `$B$2` and `B2` must not become two anchors for one cell.
  const sheet = new Workbook().addWorksheet('S');
  sheet.addCommentThread({...threadAt('$B$2'), resolved: true});
  const [thread] = sheet.commentThreads;
  assert.equal(thread?.ref, 'B2');
  assert.equal(thread?.resolved, true);
  assert.deepEqual(
    thread?.comments.map((comment) => comment.text),
    ['about $B$2'],
  );
  assert.equal(sheet.commentThreadAt('B2')?.ref, 'B2');
});

test('authoring accumulates threads rather than replacing them, unlike reader restoration', () => {
  const sheet = new Workbook().addWorksheet('S');
  sheet.addCommentThread(threadAt('B1'));
  sheet.addCommentThread(threadAt('C3'));
  assert.deepEqual(
    sheet.commentThreads.map((thread) => thread.ref),
    ['B1', 'C3'],
  );
});

test('an authored id is normalised to the one spelling the format accepts', () => {
  // Verified against the OOXML schema: `person/@id`, a message's `id`/`personId`/`parentId` and a mention's
  // `mentionpersonId`/`mentionId` are each pinned to `\{[0-9A-F]{8}-…\}`. A bare GUID is rejected and so is
  // a lower-case one, which is exactly what `crypto.randomUUID()` returns, so the authoring path
  // normalises rather than refusing the one obvious way to make an id in JavaScript.
  const wb = new Workbook();
  wb.addPerson({id: 'aaaaaaaa-1111-2222-3333-444444444444', displayName: 'Ada'});
  assert.deepEqual(
    wb.persons.map((person) => person.id),
    ['{AAAAAAAA-1111-2222-3333-444444444444}'],
  );
  assert.ok(
    wb.getPerson('{AAAAAAAA-1111-2222-3333-444444444444}'),
    'found under the normalised id',
  );

  const sheet = wb.addWorksheet('S');
  sheet.addCommentThread({
    ref: 'A1',
    resolved: false,
    comments: [
      {
        id: 'bbbbbbbb-1111-2222-3333-444444444444',
        personId: 'aaaaaaaa-1111-2222-3333-444444444444',
        text: '@Ada who owns this?',
        mentions: [
          {
            personId: 'aaaaaaaa-1111-2222-3333-444444444444',
            mentionId: '{cccccccc-1111-2222-3333-444444444444}',
            startIndex: 0,
            length: 4,
          },
        ],
      },
    ],
  });
  const [comment] = sheet.commentThreads[0]?.comments ?? [];
  assert.equal(comment?.id, '{BBBBBBBB-1111-2222-3333-444444444444}');
  assert.equal(comment?.personId, '{AAAAAAAA-1111-2222-3333-444444444444}');
  assert.equal(comment?.mentions[0]?.personId, '{AAAAAAAA-1111-2222-3333-444444444444}');
  assert.equal(comment?.mentions[0]?.mentionId, '{CCCCCCCC-1111-2222-3333-444444444444}');
});

test('authoring rejects an id that is not a GUID, rather than writing a file Excel repairs', () => {
  const wb = new Workbook();
  assert.throws(() => wb.addPerson({id: 'ada@example.com', displayName: 'Ada'}), /must be a GUID/);
  const sheet = wb.addWorksheet('S');
  const thread = threadAt('A1');
  assert.throws(
    () => sheet.addCommentThread({...thread, comments: [{...thread.comments[0]!, id: 'thread-1'}]}),
    /a comment id must be a GUID/,
  );
  assert.throws(
    () =>
      sheet.addCommentThread({
        ...thread,
        comments: [{...thread.comments[0]!, personId: 'Ada Lovelace'}],
      }),
    /a comment's author id must be a GUID/,
  );
  assert.deepEqual(sheet.commentThreads, [], 'nothing half-added survives a rejection');
});

test('a message id already used on the sheet is rejected, since a reply finds its thread by it', () => {
  // Two threads sharing a head id is not a survivable shape: on re-read the second thread's replies would
  // join whichever head was seen last, and the legacy fallback comments would both claim the same
  // `tc={id}` author and `xr:uid`. Excel's own ids happen to be globally unique.
  const sheet = new Workbook().addWorksheet('S');
  const first = threadAt('B1');
  sheet.addCommentThread(first);
  assert.throws(
    () => sheet.addCommentThread({...first, ref: 'C3'}),
    /must be unique within a sheet/,
  );
  const head = first.comments[0]!;
  assert.throws(
    () =>
      sheet.addCommentThread({
        ref: 'D4',
        resolved: false,
        comments: [{...head, id: '{99999999-0000-4000-8000-000000000000}'}, {...head}],
      }),
    /must be unique within a sheet/,
    'including a duplicate inside the incoming thread itself',
  );
  assert.deepStrictEqual(
    sheet.commentThreads.map((thread) => thread.ref),
    ['B1'],
    'nothing half-added survives either rejection',
  );
});

test('two sheets may reuse one message id, since nothing resolves across a part boundary', () => {
  // The scope in which an id means anything is the sheet: a reply names its head inside the sheet's own
  // thread part, and the fallback binds its cell inside the sheet's own comments part. Rejecting this
  // would forbid a shape that cannot break.
  const wb = new Workbook();
  const thread = threadAt('B1');
  wb.addWorksheet('One').addCommentThread(thread);
  wb.addWorksheet('Two').addCommentThread(thread);
  assert.deepStrictEqual(
    wb.worksheets.map((sheet) => sheet.commentThreads.length),
    [1, 1],
  );
});

test("a mention's span must be a whole number the wire can express", () => {
  // Verified against the OOXML schema: `startIndex` and `length` are UInt32, so 4294967295 is the last
  // legal value. Authoring rejects rather than clamps, because a clamped span silently highlights the
  // wrong words, while a *file's* unusable mention is dropped on read, keeping the message text.
  const sheet = new Workbook().addWorksheet('S');
  const thread = threadAt('B1');
  const withSpan = (startIndex: number, length: number): CommentThread => ({
    ...thread,
    comments: [
      {
        ...thread.comments[0]!,
        mentions: [{personId: GRACE_MENTIONED.id, startIndex, length}],
      },
    ],
  });
  assert.throws(() => sheet.addCommentThread(withSpan(0, 1e21)), /must be a whole number/);
  assert.throws(() => sheet.addCommentThread(withSpan(-1, 4)), /startIndex must be a whole number/);
  assert.throws(
    () => sheet.addCommentThread(withSpan(1.5, 4)),
    /startIndex must be a whole number/,
  );
  assert.throws(() => sheet.addCommentThread(withSpan(0, 4294967296)), /length must be a whole/);
  assert.strictEqual(sheet.commentThreads.length, 0, 'nothing is stored from a rejected span');

  sheet.addCommentThread(withSpan(4294967295, 1));
  assert.deepStrictEqual(
    sheet.commentThreads[0]?.comments[0]?.mentions.map((m) => [m.startIndex, m.length]),
    [[4294967295, 1]],
    'the ceiling itself is accepted',
  );
});

test('a thread cannot be anchored to anything but a single cell', () => {
  // A range, a bare column, or a bare row names no one cell to hang the conversation off, and the writer
  // would have nothing to anchor its legacy fallback comment or VML shape to.
  const sheet = new Workbook().addWorksheet('S');
  for (const ref of ['A1:B2', 'A', '2', 'nonsense']) {
    assert.throws(() => sheet.addCommentThread(threadAt(ref)), SyntaxError, ref);
  }
  assert.deepEqual(sheet.commentThreads, []);
});
