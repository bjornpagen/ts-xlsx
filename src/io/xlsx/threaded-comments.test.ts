import assert from 'node:assert/strict';
import {test} from 'node:test';

import {XmlParseError} from '../../xml/errors.ts';
import {
  buildCommentThreads,
  parsePersons,
  parseThreadedComments,
  personsXml,
  threadedCommentsXml,
} from './threaded-comments.ts';

// The part bodies below are verbatim from the Excel-authored corpus fixtures under
// `test/corpus/fixtures/threaded-comment-parts-survive-roundtrip/`, so these tests read the real
// wire shape rather than a plausible reconstruction of it. `resolved-multi-author.xlsx` is the
// interesting one: a resolved thread on B1 whose reply is by a second author, and an open thread on
// B2. Anything hand-written here is labelled as such and covers a shape Excel does not produce.

const PERSONS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" ' +
  'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<person displayName="Grace Hopper" id="{1B2C3D4E-5F60-4A71-8B92-0C1D2E3F4A5B}" ' +
  'userId="S::grace@example.com::00000000-0000-0000-0000-000000000000" providerId="AD"/>' +
  '<person displayName="Ada Lovelace" id="{77EC2EBA-7DAF-4C42-8E0A-9244C8EAE97C}" ' +
  'userId="S::ada@example.com::00000000-0000-0000-0000-000000000000" providerId="AD"/>' +
  '</personList>';

const ADA = '{77EC2EBA-7DAF-4C42-8E0A-9244C8EAE97C}';
const GRACE = '{1B2C3D4E-5F60-4A71-8B92-0C1D2E3F4A5B}';
const B1_HEAD = '{A6473DA8-237C-4E27-ADE9-48D3A7CD15A7}';

const THREADED_COMMENTS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" ' +
  'xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  `<threadedComment ref="B1" dT="2026-07-26T10:54:00.01" personId="${ADA}" ` +
  `id="${B1_HEAD}" done="1"><text>Is this gross or net of tax?</text></threadedComment>` +
  `<threadedComment ref="B1" dT="2026-07-26T10:54:00.04" personId="${GRACE}" ` +
  `id="{A758ECCA-12FA-49B1-A091-F8BA683AB83C}" parentId="${B1_HEAD}">` +
  '<text>Gross. Confirmed with finance.</text></threadedComment>' +
  `<threadedComment ref="B2" dT="2026-07-26T10:54:00.04" personId="${ADA}" ` +
  'id="{08B3B787-8F24-49B6-8D78-A5F335591EEA}"><text>Where does this figure come from?</text>' +
  '</threadedComment></ThreadedComments>';

// Values verbatim from `mention-in-thread.xlsx` (markup trimmed to the one message under test), whose
// mention Excel itself re-resolved and re-emitted on save. Excel renders `@Grace Hopper` as a mention
// chip over exactly the span `startIndex="0" length="13"` names, so `startIndex` is a 0-based character
// offset into the message text and `length` includes the leading `@`.
const MENTION_PERSON = '{BA397017-DD76-4496-AA75-59ADB199950C}';
const MENTION_ID = '{3F2C1A9E-5B84-4D67-9C2E-71A0D4E8B531}';
const MENTION_MESSAGE =
  '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">' +
  `<threadedComment ref="B2" dT="2026-07-26T10:54:00.04" personId="${ADA}" ` +
  'id="{08B3B787-8F24-49B6-8D78-A5F335591EEA}">' +
  '<text>@Grace Hopper Where does this figure come from?</text>' +
  `<mentions><mention mentionpersonId="${MENTION_PERSON}" ` +
  `mentionId="${MENTION_ID}" startIndex="0" length="13"/></mentions>` +
  '</threadedComment></ThreadedComments>';

// The same file's registry. Excel interns a *mentioned* identity as its own entry: note the third
// `<person>`: same displayName and userId as the second, different id, `providerId="PeoplePicker"`.
const MENTION_PERSONS =
  '<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">' +
  `<person displayName="Ada Lovelace" id="${ADA}" ` +
  'userId="S::ada@example.com::00000000-0000-0000-0000-000000000000" providerId="AD"/>' +
  `<person displayName="Grace Hopper" id="${GRACE}" ` +
  'userId="S::grace@example.com::00000000-0000-0000-0000-000000000000" providerId="AD"/>' +
  `<person displayName="Grace Hopper" id="${MENTION_PERSON}" ` +
  'userId="S::grace@example.com::00000000-0000-0000-0000-000000000000" providerId="PeoplePicker"/>' +
  '</personList>';

test('every registered person is parsed with the identity attributes the file carried', () => {
  assert.deepStrictEqual(parsePersons(PERSONS), [
    {
      id: GRACE,
      displayName: 'Grace Hopper',
      userId: 'S::grace@example.com::00000000-0000-0000-0000-000000000000',
      providerId: 'AD',
    },
    {
      id: ADA,
      displayName: 'Ada Lovelace',
      userId: 'S::ada@example.com::00000000-0000-0000-0000-000000000000',
      providerId: 'AD',
    },
  ]);
});

test('the personList root is not mistaken for a person entry', () => {
  assert.strictEqual(parsePersons(PERSONS).length, 2);
  assert.deepStrictEqual(parsePersons('<personList/>'), []);
});

test('a person keeps optional identity attributes absent rather than empty', () => {
  const [person] = parsePersons('<personList><person displayName="Anon" id="{A}"/></personList>');
  assert.deepStrictEqual(person, {id: '{A}', displayName: 'Anon'});
});

test('two registry entries for the same human are both kept, since only the id identifies one', () => {
  // Excel registers a mentioned identity separately from the same person's authoring entry, so the
  // registry legitimately holds duplicates by name and userId. Deduplicating on either would drop an
  // entry a mention points at; the id is the only key.
  const persons = parsePersons(MENTION_PERSONS);
  const graces = persons.filter((p) => p.displayName === 'Grace Hopper');
  assert.strictEqual(graces.length, 2, 'both Grace entries survive');
  assert.deepStrictEqual(
    graces.map((p) => [p.id, p.providerId]),
    [
      [GRACE, 'AD'],
      [MENTION_PERSON, 'PeoplePicker'],
    ],
    'they differ only by id and the provider that registered them',
  );
});

test('a person with no id is skipped, since no message could reference it', () => {
  assert.deepStrictEqual(
    parsePersons('<personList><person displayName="Nobody"/></personList>'),
    [],
  );
});

test('a person with an id but no displayName is kept: the id is what messages point at', () => {
  assert.deepStrictEqual(parsePersons('<personList><person id="{A}"/></personList>'), [
    {id: '{A}', displayName: ''},
  ]);
});

test('every message is parsed in document order with its anchor, author and timestamp', () => {
  const messages = parseThreadedComments(THREADED_COMMENTS);
  assert.deepStrictEqual(
    messages.map((m) => [m.ref, m.personId, m.date, m.text]),
    [
      ['B1', ADA, '2026-07-26T10:54:00.01', 'Is this gross or net of tax?'],
      ['B1', GRACE, '2026-07-26T10:54:00.04', 'Gross. Confirmed with finance.'],
      ['B2', ADA, '2026-07-26T10:54:00.04', 'Where does this figure come from?'],
    ],
  );
});

test('the local-time dT is kept verbatim rather than reinterpreted as an instant', () => {
  const [head] = parseThreadedComments(THREADED_COMMENTS);
  assert.strictEqual(head?.date, '2026-07-26T10:54:00.01');
});

test('a reply points at its thread head and the head points at nothing', () => {
  const [head, reply, other] = parseThreadedComments(THREADED_COMMENTS);
  assert.strictEqual(head?.id, B1_HEAD);
  assert.strictEqual(head?.parentId, undefined);
  assert.strictEqual(reply?.parentId, B1_HEAD);
  assert.strictEqual(other?.parentId, undefined);
});

test('done is read off the head alone: a reply and an open thread both report false', () => {
  const [head, reply, openHead] = parseThreadedComments(THREADED_COMMENTS);
  assert.strictEqual(head?.done, true);
  assert.strictEqual(reply?.done, false);
  assert.strictEqual(openHead?.done, false);
});

test('done="true" is honoured alongside done="1"', () => {
  const messages = parseThreadedComments(
    '<ThreadedComments><threadedComment ref="A1" id="{A}" done="true"><text>x</text>' +
      '</threadedComment><threadedComment ref="A2" id="{B}" done="0"><text>y</text>' +
      '</threadedComment></ThreadedComments>',
  );
  assert.deepStrictEqual(
    messages.map((m) => m.done),
    [true, false],
  );
});

test('each message gets its own text, with no bleed from the message before it', () => {
  const messages = parseThreadedComments(
    '<ThreadedComments><threadedComment ref="A1" id="{A}"><text>first</text></threadedComment>' +
      '<threadedComment ref="A2" id="{B}"><text>second</text></threadedComment></ThreadedComments>',
  );
  assert.deepStrictEqual(
    messages.map((m) => m.text),
    ['first', 'second'],
  );
});

test('message text is entity-decoded and keeps its line breaks', () => {
  const [message] = parseThreadedComments(
    '<ThreadedComments><threadedComment ref="A1" id="{A}">' +
      '<text>5 &lt; 6 &amp; "quoted"\nsecond line</text></threadedComment></ThreadedComments>',
  );
  assert.strictEqual(message?.text, '5 < 6 & "quoted"\nsecond line');
});

// The `_xHHHH_` group. Excel Desktop was asked directly what a patched `<text>` says: the probe is
// recorded in `docs/knowledge/specs/spreadsheetml-xhhhh-escape-is-decoded-on-read.md`: the escape
// decodes, the near-misses do not, and `_x005F_` in front of one yields the literal text. These lock
// the reader to that grammar rather than to a looser one that would eat text Excel keeps.
test('a `_xHHHH_` escape in a message decodes to the character it names', () => {
  const [message] = parseThreadedComments(
    '<ThreadedComments><threadedComment ref="A1" id="{A}">' +
      '<text>a_x0001_b</text></threadedComment></ThreadedComments>',
  );
  assert.strictEqual(message?.text, 'a\u0001b');
});

test('text that only resembles the escape shape survives a message untouched', () => {
  const [message] = parseThreadedComments(
    '<ThreadedComments><threadedComment ref="A1" id="{A}">' +
      '<text>_xZZZZ_ _x041_ _x00041_</text></threadedComment></ThreadedComments>',
  );
  assert.strictEqual(message?.text, '_xZZZZ_ _x041_ _x00041_');
});

test('the escaped underscore decodes in one pass, so the message holds the literal text', () => {
  const [message] = parseThreadedComments(
    '<ThreadedComments><threadedComment ref="A1" id="{A}">' +
      '<text>_x005F_x0041_</text></threadedComment></ThreadedComments>',
  );
  assert.strictEqual(message?.text, '_x0041_', 'a second pass would collapse this to `A`');
});

test('an escape split across two text events by an entity still decodes', () => {
  // The parser reports a chunk per entity, so `_x00&#48;1_` reaches the accumulator in three pieces.
  // Decoding each chunk as it arrives would find no escape at all; decoding the assembled body does.
  const [message] = parseThreadedComments(
    '<ThreadedComments><threadedComment ref="A1" id="{A}">' +
      '<text>_x00&#48;1_</text></threadedComment></ThreadedComments>',
  );
  assert.strictEqual(message?.text, '\u0001');
});

test('a character XML cannot carry is written into a message as its escape', () => {
  const xml = threadedCommentsXml(
    [
      {
        ref: 'A1',
        resolved: false,
        comments: [{id: '{A}', text: 'a\u0001b\uD800', mentions: []}],
      },
    ],
    ANYONE,
  );
  assert.ok(xml.includes('<text>a_x0001_b_xD800_</text>'));
});

test('a message that already reads like an escape round-trips as itself, not as its character', () => {
  const text = 'a_x0041_b';
  const xml = threadedCommentsXml(
    [{ref: 'A1', resolved: false, comments: [{id: '{A}', text, mentions: []}]}],
    ANYONE,
  );
  assert.ok(xml.includes('<text>a_x005F_x0041_b</text>'), 'the underscore is escaped in front');
  assert.strictEqual(parseThreadedComments(xml)[0]?.text, text);
});

test('the two comment systems now agree on a control character in the same string', () => {
  // The asymmetry this closed: a threaded comment refused what a legacy note carried happily.
  const text = 'before\u0001after';
  const xml = threadedCommentsXml(
    [{ref: 'A1', resolved: false, comments: [{id: '{A}', text, mentions: []}]}],
    ANYONE,
  );
  assert.strictEqual(parseThreadedComments(xml)[0]?.text, text);
});

test('a message that @mentions someone keeps its text intact around the mention', () => {
  const [message] = parseThreadedComments(MENTION_MESSAGE);
  assert.strictEqual(message?.text, '@Grace Hopper Where does this figure come from?');
  assert.strictEqual(message.personId, ADA, 'the author is unaffected by the mention it contains');
});

test('a mention is parsed with the person it names and the span it highlights', () => {
  const [message] = parseThreadedComments(MENTION_MESSAGE);
  assert.deepStrictEqual(message?.mentions, [
    {personId: MENTION_PERSON, mentionId: MENTION_ID, startIndex: 0, length: 13},
  ]);
  assert.strictEqual(
    message.text.slice(0, 13),
    '@Grace Hopper',
    'the span lands on the mentioned name, leading @ included',
  );
});

test('a message that mentions nobody reports an empty list, never an absent one', () => {
  const [message] = parseThreadedComments(THREADED_COMMENTS);
  assert.deepStrictEqual(message?.mentions, []);
});

test('a mention with no target person or no usable span is dropped, not left pointing nowhere', () => {
  // Hand-written: Excel requires all four attributes and rejects a file missing any (including the
  // capitalised `mentionPersonId` below, not a declared attribute), so every entry but the last is a
  // shape only a foreign generator produces. A chip over nothing would highlight the wrong text.
  const [message] = parseThreadedComments(
    '<ThreadedComments><threadedComment ref="A1" id="{A}"><text>@Someone hi</text><mentions>' +
      `<mention mentionPersonId="${MENTION_PERSON}" startIndex="0" length="8"/>` +
      `<mention mentionpersonId="${MENTION_PERSON}" startIndex="" length="8"/>` +
      `<mention mentionpersonId="${MENTION_PERSON}" startIndex="-1" length="8"/>` +
      `<mention mentionpersonId="${MENTION_PERSON}" startIndex="0" length="0"/>` +
      `<mention mentionpersonId="${MENTION_PERSON}" startIndex="0" length="8"/>` +
      '</mentions></threadedComment></ThreadedComments>',
  );
  assert.deepStrictEqual(message?.mentions, [{personId: MENTION_PERSON, startIndex: 0, length: 8}]);
});

test('mentions do not bleed from one message into the next', () => {
  const messages = parseThreadedComments(
    '<ThreadedComments><threadedComment ref="A1" id="{A}"><text>@x</text><mentions>' +
      `<mention mentionpersonId="${MENTION_PERSON}" mentionId="{M}" startIndex="0" length="2"/>` +
      '</mentions></threadedComment>' +
      '<threadedComment ref="A2" id="{B}"><text>plain</text></threadedComment></ThreadedComments>',
  );
  assert.deepStrictEqual(
    messages.map((m) => m.mentions.length),
    [1, 0],
  );
});

// A lookup over a parsed registry, the shape `buildCommentThreads` resolves identities through (the
// reader hands it `Workbook.getPerson`).
const lookupOver = (personsXml: string) => {
  const byId = new Map(parsePersons(personsXml).map((person) => [person.id, person]));
  return (id: string) => byId.get(id);
};
const NO_PERSONS = () => undefined;
// A registry holding every id, for a serialisation test that is not about who wrote the message.
const ANYONE = (id: string) => ({id, displayName: ''});

test('messages group into one thread per head, each reply following the head it answers', () => {
  const threads = buildCommentThreads(
    parseThreadedComments(THREADED_COMMENTS),
    lookupOver(PERSONS),
  );
  assert.deepStrictEqual(
    threads.map((thread) => [thread.ref, thread.comments.map((comment) => comment.text)]),
    [
      ['B1', ['Is this gross or net of tax?', 'Gross. Confirmed with finance.']],
      ['B2', ['Where does this figure come from?']],
    ],
  );
});

test("a thread's resolved state is its head's, so its reply cannot contradict it", () => {
  const [resolved, open] = buildCommentThreads(
    parseThreadedComments(THREADED_COMMENTS),
    lookupOver(PERSONS),
  );
  assert.strictEqual(resolved?.resolved, true, 'the head carried done="1"');
  assert.strictEqual(resolved.comments.length, 2, 'including the reply, which carries no done');
  assert.strictEqual(open?.resolved, false, 'an open thread says nothing, rather than done="0"');
});

test('every message resolves its author through the registry, keeping its timestamp verbatim', () => {
  const [thread] = buildCommentThreads(
    parseThreadedComments(THREADED_COMMENTS),
    lookupOver(PERSONS),
  );
  assert.deepStrictEqual(
    thread?.comments.map((comment) => [comment.author?.displayName, comment.date]),
    [
      ['Ada Lovelace', '2026-07-26T10:54:00.01'],
      ['Grace Hopper', '2026-07-26T10:54:00.04'],
    ],
  );
});

test('an author the registry does not hold reads as no recorded author, as Excel repairs it', () => {
  // Written back, the id would make Excel offer to repair the package, and its repair forgets the id.
  const [thread] = buildCommentThreads(parseThreadedComments(THREADED_COMMENTS), NO_PERSONS);
  const head = thread?.comments[0];
  assert.strictEqual(head?.author, undefined, 'nothing is fabricated for a missing entry');
  assert.strictEqual(head?.personId, undefined, 'and the id is not kept to be written back');
  assert.strictEqual(head?.text, 'Is this gross or net of tax?', 'the message itself is kept');
});

test('a message with no author is written under the null GUID, which reads back as no author', () => {
  const thread = {ref: 'A1', resolved: false, comments: [{id: '{H}', text: 'hi', mentions: []}]};
  const xml = threadedCommentsXml([thread], NO_PERSONS);
  assert.match(xml, /personId="\{00000000-0000-0000-0000-000000000000\}" id="\{H\}"/);
  const [back] = buildCommentThreads(parseThreadedComments(xml), NO_PERSONS);
  assert.strictEqual(back?.comments[0]?.personId, undefined);
});

test('a message whose author is not a registered person is refused rather than written', () => {
  const thread = {
    ref: 'A1',
    resolved: false,
    comments: [{id: '{H}', personId: ADA, text: 'hi', mentions: []}],
  };
  assert.throws(() => threadedCommentsXml([thread], NO_PERSONS), {
    name: 'AuthoringError',
    message: new RegExp(`its author "\\${ADA.slice(0, -1)}\\}" is not a registered person`),
  });
  assert.doesNotThrow(() => threadedCommentsXml([thread], lookupOver(PERSONS)));
});

test('a mention resolves to the entry it names: the PeoplePicker one, not its author twin', () => {
  const [thread] = buildCommentThreads(
    parseThreadedComments(MENTION_MESSAGE),
    lookupOver(MENTION_PERSONS),
  );
  const [mention] = thread?.comments[0]?.mentions ?? [];
  assert.strictEqual(mention?.person?.id, MENTION_PERSON);
  assert.strictEqual(mention.person?.displayName, 'Grace Hopper');
  assert.strictEqual(
    mention.person?.providerId,
    'PeoplePicker',
    'the same human also has an AD authoring entry; resolving by name would have picked that one',
  );
  assert.deepStrictEqual(
    [mention.startIndex, mention.length],
    [0, 13],
    'the span survives resolution unchanged',
  );
});

test('an unresolvable mention keeps its span and its target id rather than vanishing', () => {
  const [thread] = buildCommentThreads(parseThreadedComments(MENTION_MESSAGE), NO_PERSONS);
  assert.deepStrictEqual(thread?.comments[0]?.mentions, [
    {personId: MENTION_PERSON, mentionId: MENTION_ID, startIndex: 0, length: 13},
  ]);
});

test('a reply whose parent is unknown opens its own thread rather than being dropped', () => {
  // Hand-written: Excel writes a reply after the head it belongs to, so a dangling parentId is foreign
  // damage. Losing the thread's shape is recoverable; losing the message is not.
  const threads = buildCommentThreads(
    parseThreadedComments(
      '<ThreadedComments><threadedComment ref="A1" id="{A}" parentId="{GONE}">' +
        '<text>orphaned reply</text></threadedComment></ThreadedComments>',
    ),
    NO_PERSONS,
  );
  assert.deepStrictEqual(
    threads.map((thread) => [thread.ref, thread.comments.map((comment) => comment.text)]),
    [['A1', ['orphaned reply']]],
  );
});

test('a part with no messages yields no threads', () => {
  assert.deepStrictEqual(buildCommentThreads([], NO_PERSONS), []);
  assert.deepStrictEqual(
    buildCommentThreads(parseThreadedComments('<ThreadedComments/>'), NO_PERSONS),
    [],
  );
});

test('parts that prefix the extension namespace instead of defaulting it still parse', () => {
  // Hand-written: Excel defaults the namespace, but a foreign generator may bind it to a prefix.
  const messages = parseThreadedComments(
    '<tc:ThreadedComments xmlns:tc="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">' +
      '<tc:threadedComment ref="A1" id="{A}" done="1"><tc:text>prefixed</tc:text>' +
      '</tc:threadedComment></tc:ThreadedComments>',
  );
  assert.deepStrictEqual(messages, [
    {ref: 'A1', id: '{A}', text: 'prefixed', done: true, mentions: []},
  ]);
});

test('a message with no text element parses as empty rather than being dropped', () => {
  // Hand-written: Excel always writes `<text>`, so this is foreign input the reader must tolerate.
  const messages = parseThreadedComments(
    '<ThreadedComments><threadedComment ref="A1" id="{A}"/>' +
      '<threadedComment ref="A2" id="{B}"></threadedComment></ThreadedComments>',
  );
  assert.deepStrictEqual(
    messages.map((m) => [m.ref, m.text]),
    [
      ['A1', ''],
      ['A2', ''],
    ],
  );
});

test('a message that cannot be anchored or replied to is skipped, not guessed at', () => {
  const messages = parseThreadedComments(
    '<ThreadedComments><threadedComment id="{A}"><text>no ref</text></threadedComment>' +
      '<threadedComment ref="A2"><text>no id</text></threadedComment>' +
      '<threadedComment ref="A3" id="{C}"><text>kept</text></threadedComment></ThreadedComments>',
  );
  assert.deepStrictEqual(
    messages.map((m) => m.ref),
    ['A3'],
  );
});

test('a message keeps an absent author and timestamp absent', () => {
  const [message] = parseThreadedComments(
    '<ThreadedComments><threadedComment ref="A1" id="{A}"><text>anon</text>' +
      '</threadedComment></ThreadedComments>',
  );
  assert.deepStrictEqual(message, {ref: 'A1', id: '{A}', text: 'anon', done: false, mentions: []});
});

test('an empty part parses as no messages and no authors', () => {
  assert.deepStrictEqual(parseThreadedComments('<ThreadedComments/>'), []);
  assert.deepStrictEqual(parsePersons('<personList></personList>'), []);
});

test('a thread anchor is canonicalised, so an absolute reference names the same cell as a relative one', () => {
  // Excel writes a plain `B2`, but a foreign generator may anchor with `$` signs. Canonicalising in the
  // builder is what lets every later consumer compare anchors as plain strings, and what lets the
  // writer place the thread's legacy fallback without re-parsing the reference.
  const threads = buildCommentThreads(
    parseThreadedComments(
      '<ThreadedComments><threadedComment ref="$B$2" id="{A}"><text>absolute</text>' +
        '</threadedComment></ThreadedComments>',
    ),
    NO_PERSONS,
  );
  assert.deepStrictEqual(
    threads.map((thread) => thread.ref),
    ['B2'],
  );
});

test('a thread whose anchor names no single cell is dropped rather than left unanchorable', () => {
  // A range, a bare column, a bare row, and outright garbage: none of them anchors a conversation, and
  // carrying one would hand the writer a reference it cannot place a fallback comment on.
  const threads = buildCommentThreads(
    parseThreadedComments(
      '<ThreadedComments>' +
        '<threadedComment ref="A1:B2" id="{A}"><text>range</text></threadedComment>' +
        '<threadedComment ref="C" id="{B}"><text>column</text></threadedComment>' +
        '<threadedComment ref="7" id="{C}"><text>row</text></threadedComment>' +
        '<threadedComment ref="not a ref" id="{D}"><text>garbage</text></threadedComment>' +
        '<threadedComment ref="D4" id="{E}"><text>kept</text></threadedComment>' +
        '</ThreadedComments>',
    ),
    NO_PERSONS,
  );
  assert.deepStrictEqual(
    threads.map((thread) => thread.ref),
    ['D4'],
  );
});

// ── Serialisation ────────────────────────────────────────────────────────────────────────────────────

test('a conversation is written flat: the head first, then its replies naming it as their parent', () => {
  const xml = threadedCommentsXml(
    [
      {
        ref: 'B1',
        resolved: false,
        comments: [
          {
            id: '{HEAD}',
            personId: ADA,
            date: '2026-07-26T10:54:00.01',
            text: 'Gross or net?',
            mentions: [],
          },
          {
            id: '{R1}',
            personId: GRACE,
            date: '2026-07-26T10:54:00.04',
            text: 'Gross.',
            mentions: [],
          },
        ],
      },
    ],
    ANYONE,
  );
  assert.ok(
    xml.includes(
      `<threadedComment ref="B1" dT="2026-07-26T10:54:00.01" personId="${ADA}" id="{HEAD}">` +
        '<text>Gross or net?</text></threadedComment>',
    ),
    'the head carries no parentId: that is what makes it the head',
  );
  assert.ok(
    xml.includes(
      `<threadedComment ref="B1" dT="2026-07-26T10:54:00.04" personId="${GRACE}" id="{R1}" ` +
        'parentId="{HEAD}"><text>Gross.</text></threadedComment>',
    ),
    'the reply repeats the anchor and names the head, so the thread is reconstructible',
  );
});

test('done marks the head alone, and an open thread says nothing rather than done="0"', () => {
  const head = {id: '{HEAD}', text: 'q', mentions: []};
  const reply = {id: '{R1}', text: 'a', mentions: []};
  const resolved = threadedCommentsXml(
    [{ref: 'A1', resolved: true, comments: [head, reply]}],
    ANYONE,
  );
  assert.strictEqual(
    (resolved.match(/\bdone="1"/g) ?? []).length,
    1,
    'exactly one done, on the head',
  );
  assert.match(resolved, /id="\{HEAD\}" done="1"/);
  const open = threadedCommentsXml([{ref: 'A1', resolved: false, comments: [head, reply]}], ANYONE);
  assert.ok(!open.includes('done='), 'Excel omits the attribute entirely on an open thread');
});

test('an authored conversation round-trips through the parser as itself', () => {
  // The two halves are inverses, so writing the model and reading it back is the strongest single check
  // that neither invents nor drops a field.
  const threads = buildCommentThreads(parseThreadedComments(THREADED_COMMENTS), NO_PERSONS);
  assert.deepStrictEqual(
    buildCommentThreads(parseThreadedComments(threadedCommentsXml(threads, ANYONE)), NO_PERSONS),
    threads,
  );
});

test('a message that names nobody writes no mentions block at all', () => {
  const xml = threadedCommentsXml(
    [{ref: 'A1', resolved: false, comments: [{id: '{H}', text: 'plain', mentions: []}]}],
    ANYONE,
  );
  assert.ok(!xml.includes('<mentions'));
});

test('a mention is written with all four attributes Excel requires, lower-case p included', () => {
  const xml = threadedCommentsXml(
    [
      {
        ref: 'B2',
        resolved: false,
        comments: [
          {
            id: '{H}',
            text: '@Grace Hopper Where does this figure come from?',
            mentions: [
              {personId: MENTION_PERSON, mentionId: MENTION_ID, startIndex: 0, length: 13},
            ],
          },
        ],
      },
    ],
    ANYONE,
  );
  assert.ok(
    xml.includes(
      `<mentions><mention mentionpersonId="${MENTION_PERSON}" mentionId="${MENTION_ID}" ` +
        'startIndex="0" length="13"/></mentions>',
    ),
  );
  assert.match(
    xml,
    /<text>[^<]*<\/text><mentions>/,
    'mentions follow text, as the schema sequences them',
  );
});

test('a mention with no id of its own is dropped, but the text it named is not', () => {
  // All four attributes are required (each one dropped in turn gives Sch_MissRequiredAttribute), and the
  // writer has no id generator, so an invalid part would risk Excel repairing the whole conversation
  // away, where dropping the chip costs only the highlight. Excel always writes the id; this needs a
  // foreign generator.
  const xml = threadedCommentsXml(
    [
      {
        ref: 'B2',
        resolved: false,
        comments: [
          {
            id: '{H}',
            text: '@Grace Hopper who owns this?',
            mentions: [{personId: MENTION_PERSON, startIndex: 0, length: 13}],
          },
        ],
      },
    ],
    ANYONE,
  );
  assert.ok(!xml.includes('<mention'));
  assert.ok(xml.includes('<text>@Grace Hopper who owns this?</text>'));
});

test('a thread with no messages writes nothing, having neither text nor a head to reply to', () => {
  const xml = threadedCommentsXml([{ref: 'A1', resolved: true, comments: []}], ANYONE);
  assert.ok(!xml.includes('<threadedComment'));
  assert.match(xml, /<ThreadedComments xmlns="[^"]*threadedcomments"><\/ThreadedComments>$/);
});

test('markup-significant characters in a message and its ids are escaped', () => {
  const xml = threadedCommentsXml(
    [
      {
        ref: 'A1',
        resolved: false,
        comments: [
          {id: '{"&<>}', personId: 'p"&', date: '"&', text: '5 < 6 & "quoted"', mentions: []},
        ],
      },
    ],
    ANYONE,
  );
  assert.ok(!/&(?!amp;|lt;|gt;|quot;|apos;|#)/.test(xml), 'no raw ampersand survives');
  assert.ok(xml.includes('<text>5 &lt; 6 &amp; "quoted"</text>'));
  assert.ok(
    xml.includes('id="{&quot;&amp;&lt;&gt;}"'),
    'an untrusted id cannot break out of its attribute',
  );
});

test('the person registry round-trips through the parser as itself, order and all', () => {
  const persons = parsePersons(MENTION_PERSONS);
  assert.deepStrictEqual(parsePersons(personsXml(persons)), persons);
  assert.deepStrictEqual(
    parsePersons(personsXml(persons)).map((person) => person.providerId),
    ['AD', 'AD', 'PeoplePicker'],
    'the mentioned identity stays its own entry rather than merging into its author twin',
  );
});

test('a person writes only the identity attributes the model holds', () => {
  const xml = personsXml([{id: '{A}', displayName: 'Anon'}]);
  assert.ok(xml.includes('<person displayName="Anon" id="{A}"/>'));
  assert.ok(!xml.includes('userId='), 'an absent handle is omitted, not written empty');
  assert.ok(!xml.includes('providerId='));
});

test('a display name carrying markup cannot reshape the registry', () => {
  const xml = personsXml([{id: '{A}', displayName: 'Ada "&" <Lovelace>', userId: 'S::a&b'}]);
  assert.ok(!/&(?!amp;|lt;|gt;|quot;|apos;|#)/.test(xml));
  assert.ok(xml.includes('displayName="Ada &quot;&amp;&quot; &lt;Lovelace&gt;"'));
});

test('an empty registry still writes a well-formed part rather than a stub', () => {
  assert.match(personsXml([]), /<personList xmlns="[^"]*threadedcomments"><\/personList>$/);
});

// ── Hostile input ────────────────────────────────────────────────────────────────────────────────────
// Both parts come from an untrusted file. The SAX layer already bounds the shapes that would otherwise
// need guarding here: it is a single non-recursive O(n) pass, and it decodes entities without ever
// expanding them, so neither nesting depth nor a billion-laughs payload can reach these parsers, and the
// package's inflate ceiling bounds the bytes. What is left is this reader's own arithmetic and state
// machine, and the fact that what it accepts, the writer re-emits.

test('a mention offset beyond what the wire can express is dropped, chip lost and text kept', () => {
  // Verified against the OOXML schema: `startIndex` and `length` are both UInt32, so 4294967295 is the
  // last valid value and 4294967296 fails as "not a valid 'UInt32' value". This is the guard that matters
  // most, because accepting such a value would put it in OUR output: `String(1e21)` is `"1e+21"`, which is
  // not a numeric literal any schema accepts, and one invalid attribute is enough for Excel to offer to
  // repair the conversation away entirely.
  const mention = (startIndex: string, length: string) =>
    `<mention mentionpersonId="${MENTION_PERSON}" mentionId="${MENTION_ID}"` +
    ` startIndex="${startIndex}" length="${length}"/>`;
  const [message] = parseThreadedComments(
    '<ThreadedComments><threadedComment ref="A1" id="{A}"><text>@x over the ceiling</text><mentions>' +
      mention('4294967296', '2') +
      mention('0', '4294967296') +
      mention('1e21', '2') +
      mention('4294967295', '2') +
      '</mentions></threadedComment></ThreadedComments>',
  );
  assert.deepStrictEqual(
    message?.mentions.map((m) => [m.startIndex, m.length]),
    [[4294967295, 2]],
    'only the offset the wire can express survives',
  );
  assert.strictEqual(message.text, '@x over the ceiling', 'the message itself is untouched');
});

test('the writer cannot emit an out-of-range span even from a model that was handed one', () => {
  // `restoreCommentThreads` takes a model wholesale, so the serialiser refuses the span itself rather
  // than trusting that every path into the model already checked it.
  const xml = threadedCommentsXml(
    [
      {
        ref: 'A1',
        resolved: false,
        comments: [
          {
            id: '{H}',
            text: '@x hi',
            mentions: [
              {personId: MENTION_PERSON, mentionId: MENTION_ID, startIndex: 0, length: 1e21},
              {personId: MENTION_PERSON, mentionId: MENTION_ID, startIndex: 1.5, length: 2},
            ],
          },
        ],
      },
    ],
    ANYONE,
  );
  assert.ok(!xml.includes('<mention'), 'neither span is written');
  assert.ok(!xml.includes('e+'), 'so no exponent-form number reaches the part');
  assert.ok(xml.includes('<text>@x hi</text>'), 'and the message survives without its chips');
});

test('an entity a part declares itself is not expanded, so a nested-entity payload stays inert', () => {
  // Not merely bounded but structurally impossible: the scanner skips markup declarations without reading
  // them and decodes only the five predefined entities, so `&lol;` resolves to nothing to expand.
  const xml =
    '<?xml version="1.0"?><!DOCTYPE ThreadedComments [<!ENTITY lol "haha">' +
    '<!ENTITY lol2 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">]>' +
    '<ThreadedComments><threadedComment ref="A1" id="{A}"><text>&lol2;</text></threadedComment>' +
    '</ThreadedComments>';
  const [message] = parseThreadedComments(xml);
  assert.strictEqual(message?.text, '&lol2;', 'the reference is carried verbatim, never expanded');
});

test('a message nested inside another is not fabricated into a thread of its own', () => {
  // A shape no producer emits and the schema forbids. The parser holds one open message at a time, so the
  // inner close commits and the outer one finds nothing left to commit: the damage costs a message, not
  // a crash and not a duplicate.
  const messages = parseThreadedComments(
    '<ThreadedComments><threadedComment ref="A1" id="{A}"><text>outer</text>' +
      '<threadedComment ref="A2" id="{B}"><text>inner</text></threadedComment>' +
      '</threadedComment></ThreadedComments>',
  );
  assert.deepStrictEqual(
    messages.map((m) => [m.id, m.text]),
    [['{B}', 'inner']],
  );
});

test('a mention outside any message is discarded rather than attaching to the next one', () => {
  const messages = parseThreadedComments(
    '<ThreadedComments>' +
      `<mention mentionpersonId="${MENTION_PERSON}" mentionId="${MENTION_ID}" startIndex="0" length="2"/>` +
      '<threadedComment ref="A1" id="{A}"><text>plain</text></threadedComment></ThreadedComments>',
  );
  assert.deepStrictEqual(
    messages.map((m) => m.mentions.length),
    [0],
  );
});

test('a stray text element outside a message cannot leak into the next message', () => {
  const messages = parseThreadedComments(
    '<ThreadedComments><text>orphaned</text>' +
      '<threadedComment ref="A1" id="{A}"><text>mine</text></threadedComment></ThreadedComments>',
  );
  assert.deepStrictEqual(
    messages.map((m) => m.text),
    ['mine'],
  );
});

test('a truncated part fails loudly rather than yielding a half-read conversation', () => {
  // A truncated tag is unrecoverable: anything after it is unparsed bytes, and guessing where the element
  // ended would invent structure. The throw propagates out of the whole read, as it does for any malformed
  // part (verified): a corrupt package is a hard error here, never a silently halved conversation. That is
  // a different failure from an *unreachable* part, which is tolerated because nothing is ambiguous about
  // it. See the `threaded-comment-rel-empty-target-tolerated` corpus case.
  assert.throws(
    () => parseThreadedComments('<ThreadedComments><threadedComment ref="A1" id="{A"'),
    XmlParseError,
  );
});

test('a person entry repeated in the registry is parsed as written, leaving precedence to the caller', () => {
  // The parser is faithful to the part; the workbook registry is what collapses ids (last wins). A parser
  // that de-duplicated here would hide a foreign generator's damage from anything trying to diagnose it.
  const persons = parsePersons(
    `<personList><person id="${MENTION_PERSON}" displayName="First"/>` +
      `<person id="${MENTION_PERSON}" displayName="Second"/></personList>`,
  );
  assert.deepStrictEqual(
    persons.map((person) => person.displayName),
    ['First', 'Second'],
  );
});
