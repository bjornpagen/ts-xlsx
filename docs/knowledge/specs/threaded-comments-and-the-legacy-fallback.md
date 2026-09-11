# Modern threaded comments, and why we own their legacy fallback

Cluster: comments

## Scenario

A reviewer leaves a modern threaded comment in Excel: a conversation with an author identity, a
timestamp, replies, an optional resolved state, and possibly an `@mention`. This is not the same
feature as a legacy note (`cell.note`); the two coexist in real files, on different cells of the same
sheet. Round-tripping such a file must keep the conversation, and callers must be able to read it as
structure and author a new one.

The trap that made this urgent: Excel writes a **legacy fallback `<comment>`** beside every thread,
carrying the `[Threaded comment]\n\nYour version of Excel allows you to read this…` boilerplate with
a copy of the messages, authored by a synthetic `tc={headId}` entry. Read naïvely, that boilerplate
becomes `cell.note`, so a load and save produced a garbage note _and_ lost the conversation.

> Implemented and corpus-locked (2026-07-26). Read, write, and authoring all ship; see
> `threaded-comment-conversations-read-into-model`,
> `threaded-comment-authored-in-model-round-trips`, `threaded-comment-parts-survive-roundtrip`, and
> `threaded-comment-rel-empty-target-tolerated`. Recorded here because the wire facts below cost real
> probe work to establish and several are not derivable from the schema.

## Wire format (verified against Excel-authored bytes, not inferred)

An MS-XLSX extension, not base ECMA-376. Both parts' roots share the namespace
`http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments` (plural
`threadedcomments`).

- **Per sheet**, `xl/threadedComments/threadedComment{n}.xml`, root `<ThreadedComments>`:
  ```xml
  <threadedComment ref="A1" dT="2026-07-24T10:56:41.72" personId="{GUID}" id="{GUID}" done="1">
    <text>the message body</text>
    <mentions><mention mentionpersonId="{GUID}" mentionId="{GUID}" startIndex="0" length="13"/></mentions>
  </threadedComment>
  ```
  A reply carries `parentId` equal to the head's `id` and follows its head in document order. `done="1"`
  sits on the **head only** and is absent, not `done="0"`, when the thread is open, so a message's
  resolved-ness is a property of its thread. `dT` is **local wall-clock with fractional seconds and
  no zone or `Z`**, not a round-trippable instant, which is why we carry it as a verbatim string
  rather than a `Date`.
- **Per workbook**, `xl/persons/person.xml`, **singular and unnumbered**, root `<personList>`:
  ```xml
  <person displayName="Jane Doe" id="{GUID}" userId="S::jane@example.com::{tenant}" providerId="AD"/>
  ```
- **Relationship types**, the _only_ wiring, since no worksheet element names these parts the way none
  names a pivot table: `…/office/2017/10/relationships/threadedComment` from the sheet and
  `…/office/2017/10/relationships/person` from the workbook.
- **Content types**, both `+xml` overrides: `application/vnd.ms-excel.threadedcomments+xml` and
  `application/vnd.ms-excel.person+xml`.

Facts that are easy to get subtly wrong, each established by a probe:

- **Every one of the six id attributes is pinned to a brace-wrapped, UPPER-CASE hex GUID**:
  `person/@id`, a message's `id`, `personId` and `parentId`, and a mention's `mentionpersonId` and
  `mentionId`. Both halves bite: a bare GUID is rejected and so is a lower-case one, so
  **`crypto.randomUUID()` output is invalid as written**, which is the one obvious way to make an id in
  JS. The authoring path therefore _normalises_, adding braces and upper-casing, rather than merely
  validating, and the reader never rewrites an id, because every reference points at it. See
  `commentThreadGuid`.
- **Note the lowercase `p` in `mentionpersonId`.** `mentionPersonId` is rejected as undeclared, a
  useful negative control, since it proves the validator really inspects this part rather than
  skipping it as unknown.
- **All four mention attributes are required**, since dropping each in turn gives
  `Sch_MissRequiredAttribute`. `startIndex` is a 0-based character offset into the message text and
  `length` **includes the leading `@`**, so `@Grace Hopper` is 13, settled by _rendering_, the only
  oracle that could: Excel draws the chip over exactly that span.
- **`startIndex` and `length` are `xsd:unsignedInt`**: `4294967295` validates, `4294967296` fails as
  "not a valid 'UInt32' value".
- **A message's `personId` is required and must name a registered person.** Leaving it out is
  `Sch_MissRequiredAttribute`, and Excel offers to repair the package; an id the registry does not
  hold validates but gets the same repair prompt (open verdict, Excel 16.0 build 20326, 2026-09-11).
  Excel's repair of the second rewrites the id to the null GUID
  `{00000000-0000-0000-0000-000000000000}` and shows the message by "Author", and a package carrying
  the null GUID opens clean with a registry or without one. That is the wire's only spelling of an
  unknown author, so the writer writes a message with no author under it and refuses an authored id
  nobody registered, and the reader reads an unregistered id as no author.
- **A mentioned identity gets its own `<person>` entry with `providerId="PeoplePicker"`**, distinct
  from the same human's `providerId="AD"` authoring entry: same `displayName` and `userId`, different
  `id`. So `providerId` is not always `"AD"`, the registry legitimately holds several entries per
  human, and **the registry must be keyed by person `id` alone**, since interning by name or `userId`
  merges these two and silently breaks the mention on write.
- **Excel's own `<personList>` order is not stable**, because its save re-sorts by person id, so
  nothing may depend on entry order.
- **Thread part numbers follow the sheet index and a gap is fine.** A workbook whose first sheet has
  only notes writes `threadedComment2.xml` with no `threadedComment1.xml`, and nothing addresses these
  parts by number.
- **Three things Excel writes that are unnecessary**, all omitted by our output with every thread,
  reply, resolved flag and mention still read back: `xmlns:x` on both part roots, declared and never
  used; `shapeId="0"` on a fallback `<comment>`; and an `xr:uid` on a _genuine_ note, its own
  revision id, which is revision metadata and `mc:Ignorable`.
- **Excel writes the fallback text with CRLF; we write LF.** XML 1.0 §2.11 line-end normalisation
  makes these identical to every reader, so the corpus compares after normalising. Do not "fix" the
  writer to emit CR.
- **Mentions cannot be authored locally at all.** Excel's `@` picker never appears for a workbook on
  a local disk; it needs OneDrive or SharePoint so it can resolve and notify people. Ground truth came
  from injecting a schema-correct shape and letting Excel re-canonicalize it, so the committed fixture
  is Excel's own output rather than a reconstruction.

## The load-bearing finding: `tc=` plus `xr:uid` bind a cell to its thread

The single most consequential fact, and the reason this feature is not "preserve the parts and move
on". Preserving `threadedComment{n}.xml`, `person.xml`, both relationships and both content-type
overrides yields a package that **validates clean and in which Excel reads back zero threaded
comments and three ordinary notes.** Cause: re-serialising `comments{n}.xml` from the note model
collapses `<authors>` to one empty `<author/>` and drops each `<comment>`'s `xr:uid`, and Excel
matches a cell to its thread through exactly those two. The thread part is intact and orphaned.

The converse holds too, and is worse: a `tc=`-authored `<comment>` with its `xr:uid` but **no** thread
part opens clean, validates clean, and reports `CommentsThreaded.Count = 0` _and_ `Comments.Count = 0`,
so the text is shown as neither a thread nor a note. It simply disappears.

So **the fallback comment and the thread part are two halves of one representation and must be
emitted together or not at all.** In `write.ts` both halves derive from a single `threads` variable,
which makes the invariant structural rather than a rule someone has to remember.

## Behavior as shipped

**Read.** `parsePersons` and `parseThreadedComments` (`src/io/xlsx/threaded-comments.ts`) stay faithful
to document order; `buildCommentThreads` groups messages into threads, resolves authors and mentions
against the workbook registry, and canonicalises each thread's `ref`. Persons load before the sheet
loop, since the sheets resolve through the registry.

**The model** (`src/core/comment-thread.ts`, on the public barrel): `Person`, `Comment`,
`CommentThread`, `Mention`. Two shape decisions worth not re-litigating:

- **There is no `cell.commentThread`.** A `Cell` has no back-reference to its sheet, let alone to the
  workbook the persons live on, and a thread is not cell _content_ the way a note is, so putting it in
  `CellModel` would advertise authoring the writer ignores. It is a sheet-level read view instead
  (`commentThreads`, `commentThreadAt(ref)`), mirroring `loadedPivotTables`. `commentThreadAt`
  canonicalises the ref **without** going through `getCell`, which would materialise a cell and
  change written output on a read-only query.
- **A mention carries both the resolved object and the raw id** (`Mention.person` plus `personId`), so
  an id the registry does not hold stays diagnosable instead of reading as blank. `mentionId` is carried
  for the same reason: the writer cannot invent one. A message's author does not keep an unresolved id:
  written back it would be the package Excel repairs, and Excel's repair forgets it.

**Write.** `threadedCommentsXml` and `personsXml`, exact inverses of the parsers, carrying **no clock
and no id generator**: every guid and timestamp comes from the model, so one workbook always
serialises byte-identically. Attribute order matches Excel's. The registry is emitted only when some
sheet emits a thread part, since with no conversation nothing can reference a `<person>`.

**Authoring.** `Worksheet.addCommentThread` and `Workbook.addPerson`. These had to land with the
writer: once emission came from the model, `restoreCommentThreads` was the only way in, which is a
reader's name on the write path. Authoring is "pass a fully-formed conversation", so an **injected
id/clock factory turned out to be unnecessary** for determinism, because the serialiser cannot reach a
clock even in principle. A generator would be pure convenience.

**The fallback is owned, not tolerated** (the D4/D5 rationale):

- On read, a `tc={guid}` author marks its comment as a fallback, and it is dropped **only when we hold
  the thread it names**, keyed on the thread _id_ rather than on the cell having been threaded. A file
  whose thread part is missing or unreachable therefore _keeps_ the boilerplate as `cell.note`, because
  it is then the last remaining record of what was said.
- On write it is rebuilt from the thread model with its `tc={headId}` author, its `xr:uid`, and Excel's
  own author layout, thread authors first then the shared anonymous note author. `Reply:` **repeats
  per reply** (`Comment:\n    <head>\nReply:\n    <r1>\nReply:\n    <r2>`), every body indented four
  spaces, one `<comment>` per thread rather than per message.
- `<authors>` indices are consequently load-bearing: a self-closing `<author/>` must be given an
  explicit close tag or every later index shifts and an ordinary note inherits a `tc=` author, at
  which point suppression deletes it. The `xmlns:mc`, `mc:Ignorable="xr"` and `xmlns:xr` declarations
  are emitted only when a fallback exists, so a note-only part stays byte-unchanged.
- **A note and a thread cannot share a cell.** Excel refuses both orders, since `AddComment` on a
  threaded cell and `AddCommentThreaded` on a noted cell each fail with `0x800A03EC`, so coexistence is
  _per sheet, not per cell_. Only a foreign generator or a hand-edit produces that shape: tolerate it,
  do not model it.

## Hostile input: the exposure runs toward the writer

"Bound the new parser paths" had no allocation left to guard: the package's inflate ceiling bounds
the bytes, and the SAX layer bounds the shapes, in one non-recursive O(n) pass with entities decoded
but never expanded, so billion-laughs is structurally impossible rather than mitigated. The real
exposure was the other direction: **what the reader accepts, the writer re-emits.** A mention offset
from a foreign part flows straight into our output, and `String(1e21)` is `"1e+21"`, a numeric literal
no schema accepts, so one hostile offset makes _our_ file invalid and invites Excel to repair the whole
conversation away.

Bounded at the schema's UInt32 ceiling in three places by design (`MENTION_OFFSET_MAX`,
`commentThreadOffset`): the parser drops such a mention while keeping the message text, the authoring
verb throws, and `mentionsXml` refuses it regardless, so the serialiser _cannot_ emit an
out-of-range span however the model was populated. Generalisable: on a feature that round-trips, "is
this input safe to hold?" is the wrong question. Ask **"is this input safe to write back?"**

Two more deliberate asymmetries:

- **A message id is only meaningful within its sheet.** A reply names its head inside the sheet's own
  thread part, and the fallback binds its cell by the same id inside the sheet's own comments part, so
  nothing resolves across a part boundary. A **cross-sheet** id collision is therefore harmless and is
  _not_ rejected, while a **within-sheet** one is, since two heads sharing an id would make a reply
  join the wrong thread on re-read and hand two fallbacks one `tc=` author and one `xr:uid`.
- **An unreachable part is tolerated; a malformed one is not.** A sheet rel with `Target=""` yields no
  threads, keeps the boilerplate as a note, and re-writes neither the thread part nor an orphaned `tc=`
  author, and notably does **not** adopt the part by filename, since guessing is how a conversation
  gets attached to the wrong sheet. A _truncated_ part instead throws out of the whole read, like any
  corrupt part: a broken package is a hard error, never a silently halved conversation.
- Tolerated without ceremony: an orphan `personId`, where the id stays on the model unresolved; a
  dangling `parentId`, where the message opens its own thread rather than being lost; a thread on an
  empty cell, since a comment anchors regardless of the cell's value; and a duplicated `<person>`,
  parsed as written, because precedence belongs to the workbook registry.

## Deliberately not modeled

A genuine note's own `xr:uid` and `shapeId`, which is revision metadata Excel does not need back from
us; notification and mention delivery, which is a service concern rather than a document one; and any
per-message `resolved` flag, since resolution is a thread property because only the head carries
`done`.

Related: `reading-comments-with-vml-drawing-must-terminate`, `comment-explicit-size`,
`worksheet-comment-rel-empty-target-tolerated` (the legacy-note sibling of the tolerance case).
No ADR: ADR-0014 §3 already frames picking up a preserved-only part family as a normal feature slice
rather than a decision needing its own record.
