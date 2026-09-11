# Comment Thread

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `Comment`

<sub>interface</sub>

One message of a [`CommentThread`](./comment-thread.md#commentthread): what a single person wrote, once.

```ts
interface Comment {
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
```

---

### `CommentThread`

<sub>interface</sub>

A conversation anchored to one cell: what was asked, every reply, and whether it was resolved.

```ts
interface CommentThread {
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
```

---

### `Mention`

<sub>interface</sub>

A [`MentionRef`](./comment-thread.md#mentionref) with its identity resolved against the workbook's person registry.

```ts
interface Mention extends MentionRef {
  /**
   * The mentioned identity, resolved through the workbook registry. Absent when the file names an id
   * the registry does not hold (a mention left dangling by a foreign generator); {@link personId}
   * still says who was meant.
   */
  readonly person?: Person;
}
```

---

### `MentionRef`

<sub>interface</sub>

An `@mention` as the file spells it: who was named, and the run of [`Comment.text`](./comment-thread.md#comment) that
renders as the mention chip.

The offsets are only meaningful against that exact text: shift either and a spreadsheet app
highlights the wrong words.

This is the wire shape, shared with the codec that reads it. [`Mention`](./comment-thread.md#mention) is this plus the
identity we resolved the id to, which is the one thing the file does not carry.

```ts
interface MentionRef {
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
```

---

### `Person`

<sub>interface</sub>

A registered identity a threaded comment can point at: an author, or someone `@mentioned` in a
message. One `<person>` of the workbook's `xl/persons/person.xml` registry.

A single human legitimately has **several** entries: Excel registers a mentioned identity separately
from that person's authoring identity, with the same `displayName` and `userId` but a
different `id` and a different `providerId`. The id is therefore the only identity;
see [`Workbook.getPerson`](./workbook.md#workbookgetperson).

```ts
interface Person {
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
```
