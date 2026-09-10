// Structural facts read back out of a written package.
//
// These answer "what is actually in the zip": part inventory, comment threads, notes,
// without going through our own reader, so a case can assert on the bytes rather than on a
// round-trip through the code that produced them.

import assert from 'node:assert/strict';

import {strFromU8, strToU8, unzipSync, zipSync} from 'fflate';

import {readXlsx, type WorkbookInstance, type WorksheetInstance, writeXlsx} from './runtime.ts';

/** A package flattened to part name → part text: what every XML-level probe here reads from. */
export type PartMap = Record<string, string>;

/**
 * A written workbook read back: the write→read cycle that fifty capability bodies across seven adapter
 * modules spell out by hand.
 *
 * The src side has had this in its own support module for as long as there has been one, and named it
 * for the same reason: what a capability is asking is "does this survive a round-trip", and spelling
 * that as a nested call each time hides the question inside the plumbing.
 */
export function roundtrip(workbook: WorkbookInstance): WorkbookInstance {
  return readXlsx(writeXlsx(workbook));
}

/**
 * One named part's text, failing the case (naming the part) when it is absent.
 *
 * {@link partMapOf} hands back a bare record, so every caller reaching one part off it writes
 * `?? ''` or a non-null assertion. The empty-string spelling is the dangerous one: a negative
 * assertion built on it goes on passing if the writer ever renames the part it reads.
 */
export function partOf(buffer: Uint8Array, name: string): string {
  return strFromU8(partBytesOf(buffer, name));
}

/** One named part's bytes, for a binary part, failing the case (naming the part) when it is absent. */
export function partBytesOf(buffer: Uint8Array, name: string): Uint8Array {
  const bytes = unzipSync(buffer, {filter: (file) => file.name === name})[name];
  assert.ok(bytes !== undefined, `expected part ${name}`);
  return bytes;
}

/** Every part path in a package, without inflating any of them. */
export function partNamesOf(buffer: Uint8Array): string[] {
  const names: string[] = [];
  unzipSync(buffer, {
    filter: (file) => {
      names.push(file.name);
      return false;
    },
  });
  return names;
}

/** What {@link patchedPackage} does to a package, each map keyed by part path. */
export interface PackagePatch {
  /** Rewrite a part's text. The part must exist and the rewrite must change it. */
  readonly edit?: Readonly<Record<string, (xml: string) => string>>;
  /** Add a part, or replace one outright, with these bytes or this text. */
  readonly put?: Readonly<Record<string, Uint8Array | string>>;
  /** Remove parts. Each must exist. */
  readonly drop?: readonly string[];
}

/**
 * A written package with some parts rewritten, added or removed: how a case hands the reader markup
 * a real producer emits and this writer never does.
 *
 * Nineteen sites spelled this out by hand, reading the part through a non-null assertion, whose
 * `TypeError` names nothing, or through `|| new Uint8Array()`, which patches an empty string and
 * lets every negative assertion after it pass. An edit that changes nothing fails as well: a
 * `replace` whose search string the writer stopped emitting leaves the case reading the unpatched
 * package while it claims to read a hostile one.
 */
export function patchedPackage(buffer: Uint8Array, patch: PackagePatch): Uint8Array {
  const files: Record<string, Uint8Array> = unzipSync(buffer);
  for (const [name, transform] of Object.entries(patch.edit ?? {})) {
    const bytes = files[name];
    assert.ok(bytes !== undefined, `expected part ${name}`);
    const before = strFromU8(bytes);
    const after = transform(before);
    assert.notEqual(after, before, `edit to ${name} changed nothing`);
    files[name] = strToU8(after);
  }
  for (const name of patch.drop ?? []) {
    assert.ok(files[name] !== undefined, `expected part ${name}`);
    delete files[name];
  }
  for (const [name, data] of Object.entries(patch.put ?? {})) {
    files[name] = typeof data === 'string' ? strToU8(data) : data;
  }
  return zipSync(files);
}

export function partMapOf(buffer: Uint8Array): PartMap {
  const unzipped = unzipSync(buffer);
  const out: Record<string, string> = {};
  for (const name of Object.keys(unzipped)) out[name] = strFromU8(unzipped[name]!);
  return out;
}

// Total matches of `pattern` across every part whose name matches `inParts`, for facts that live in
// the part's content rather than in the package's shape.
export const countIn = (parts: Record<string, string>, inParts: RegExp, pattern: RegExp) =>
  Object.keys(parts)
    .filter((p) => inParts.test(p))
    .reduce((n, p) => n + [...(parts[p] ?? '').matchAll(pattern)].length, 0);

// A sheet's legacy notes as `{<ref>: <text>}`, in row-major order.
export const notesOf = (sheet: WorksheetInstance) => {
  const notes: Record<string, string> = {};
  for (const {cells} of sheet.rows()) {
    for (const cell of cells) if (cell.note !== undefined) notes[cell.address] = cell.note;
  }
  return notes;
};

// The modern threaded conversations a workbook holds → the person registry plus, per sheet, each thread's
// anchor/resolved state and its messages in order (author resolved through the registry, the raw author id
// alongside so an unresolved one is visible, and each @mention's resolved identity + text span). `refs`
// additionally probes the per-cell lookup, reporting the anchor of the thread found at each reference (null
// for none), so a case can assert a noted cell is not mistaken for a threaded one.
//
// Persons are sorted by id because the registry's order is meaningless: Excel re-sorts the part by person
// id whenever it saves, so only membership is a fact.
export const commentThreadFacts = (wb: WorkbookInstance, refs: string[] = []) => {
  type Person = WorkbookInstance['persons'][number];
  const identity = (person: Person) => ({
    id: person.id,
    displayName: person.displayName,
    userId: person.userId ?? null,
    providerId: person.providerId ?? null,
  });
  // A comment's author is a *lookup* through the registry and can miss; a registry entry cannot. Two
  // functions rather than one nullable one, so `persons` is not reported as possibly-null purely
  // because the author path shares its shape: a case reading a registry entry should not have to
  // narrow past a `null` that cannot occur.
  const identityOrNull = (person: Person | undefined) => (person == null ? null : identity(person));
  return {
    persons: wb.persons
      .map((person) => identity(person))
      .sort((a, b) => String(a.id).localeCompare(String(b.id))),
    sheets: wb.worksheets.map((sheet) => ({
      name: sheet.name,
      threads: sheet.commentThreads.map((thread) => ({
        ref: thread.ref,
        resolved: thread.resolved,
        comments: thread.comments.map((comment) => ({
          id: comment.id,
          author: identityOrNull(comment.author),
          authorId: comment.personId ?? null,
          date: comment.date ?? null,
          text: comment.text,
          mentions: comment.mentions.map((mention) => ({
            person: identityOrNull(mention.person),
            personId: mention.personId,
            startIndex: mention.startIndex,
            length: mention.length,
            // The exact run of text the mention chip covers: the only check that proves the span was not
            // shifted, since the offsets alone are just numbers.
            span: comment.text.slice(mention.startIndex, mention.startIndex + mention.length),
          })),
        })),
      })),
      at: Object.fromEntries(refs.map((ref) => [ref, sheet.commentThreadAt(ref)?.ref ?? null])),
      // Every legacy note the sheet reads back as `{<ref>: <text>}`. Reported alongside the threads because
      // the interesting fact is what is NOT here: Excel writes a boilerplate fallback comment beside every
      // thread, and surfacing that as a note would hand the caller garbage.
      notes: notesOf(sheet),
    })),
  };
};

// Every legacy fallback `<comment>` a comments part carries, paired with the thread it binds to: the
// comments whose `xr:uid` also appears as a `tc={guid}` author. Excel writes exactly one per threaded
// conversation and fills it with the boilerplate a pre-2018 reader shows in place of the conversation,
// so both the binding and that rendered text are observable facts. Sorted by uid, since order in the
// part follows cell position and carries no meaning here.
export const threadFallbackComments = (parts: Record<string, string>) => {
  const found: Array<{uid: string; text: string}> = [];
  for (const name of Object.keys(parts).filter((p) => /comments\d+\.xml$/.test(p))) {
    const xml = parts[name] ?? '';
    const threadAuthors = new Set(
      [...xml.matchAll(/<author>tc=([^<]*)<\/author>/g)].map((m) => m[1]),
    );
    for (const block of xml.matchAll(/<comment\b([^>]*)>([\s\S]*?)<\/comment>/g)) {
      const uid = (/\bxr:uid="([^"]*)"/.exec(block[1] ?? '') ?? [])[1];
      if (uid === undefined || !threadAuthors.has(uid)) continue;
      const text = [...(block[2] ?? '').matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)]
        .map((m) => m[1] ?? '')
        .join('')
        // `&amp;` last, so an escaped entity (`&amp;lt;`) decodes to the entity and not to `<`.
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&amp;/g, '&')
        // Line-end normalisation, which XML 1.0 §2.11 requires of every processor before the
        // application sees the text: a literal CRLF in element content IS a lone LF as far as any
        // reader is concerned. Excel writes the boilerplate with CRLF, we write LF, and no consumer can
        // tell, so doing it here keeps the comparison about wording instead of about line endings.
        .replace(/\r\n?/g, '\n');
      found.push({uid, text});
    }
  }
  return found.sort((a, b) => a.uid.localeCompare(b.uid));
};

// Package-part facts a passthrough round-trip must preserve: the mirror of the oracle's
// `packageFactsFromZip`: counts of part families the reader does not fully model (drawings, VML,
// media, pivot tables/caches, comments) plus the worksheet/drawing reference flags that wire
// unmodeled features (a vector-shape drawing, a header/footer image) into the sheet.
export const packagePartFacts = (parts: Record<string, string>) => {
  const names = Object.keys(parts);
  const at = (rx: RegExp) => {
    const found = names.find((p) => rx.test(p));
    return found !== undefined ? (parts[found] ?? '') : '';
  };
  const ws1 = at(/worksheets\/sheet1\.xml$/);
  const drawing1 = at(/drawings\/drawing1\.xml$/);
  const wb = at(/^xl\/workbook\.xml$/);
  return {
    drawings: names.filter((p) => /xl\/drawings\/drawing\d+\.xml$/.test(p)).length,
    vml: names.filter((p) => /vmlDrawing\d+\.vml$/.test(p)).length,
    media: names.filter((p) => /xl\/media\//.test(p)).length,
    pivotTables: names.filter((p) => /pivotTables\/pivotTable\d+\.xml$/.test(p)).length,
    pivotCache: names.filter((p) => /pivotCache\/.+\.xml$/.test(p)).length,
    slicers: names.filter((p) => /slicer/i.test(p)).length,
    comments: names.filter((p) => /comments\d+\.xml$/.test(p)).length,
    // Modern threaded comments (2018): the per-sheet thread parts and the workbook-level author registry
    // (`xl/persons/person.xml`, singular and unnumbered). Both are serialised from the model, so a missing
    // one means the conversation or its authors were dropped, and a doubled one means something is emitting
    // a part twice.
    threadedComments: names.filter((p) => /threadedComments\/threadedComment\d+\.xml$/.test(p))
      .length,
    persons: names.filter((p) => /xl\/persons\/person\d*\.xml$/.test(p)).length,
    // The conversation *inside* those parts, so preservation is asserted on content and not merely on
    // a part existing: one `<threadedComment>` per message (replies carry `parentId`), a thread head
    // marked resolved by `done`, and the `<person>` registry entries the `personId`s resolve through.
    // The distinct personId set is what proves a multi-author thread keeps each message's author,
    // reported sorted, since only membership is meaningful.
    threadedCommentMessages: countIn(parts, /threadedComments\//, /<threadedComment\b/g),
    threadedCommentReplies: countIn(
      parts,
      /threadedComments\//,
      /<threadedComment\b[^>]*\bparentId=/g,
    ),
    resolvedThreadHeads: countIn(
      parts,
      /threadedComments\//,
      /<threadedComment\b[^>]*\bdone="(?:1|true)"/g,
    ),
    threadedCommentAuthorIds: [
      ...new Set(
        names
          .filter((p) => /threadedComments\//.test(p))
          .flatMap((p) => [...(parts[p] ?? '').matchAll(/personId="([^"]*)"/g)])
          .map((m) => m[1]!),
      ),
    ].sort(),
    personEntries: countIn(parts, /xl\/persons\/person\d*\.xml$/, /<person\b/g),
    // An @mention inside a message: `<mention>` names the mentioned person and pins the span of text
    // that renders as the mention chip. Verified against desktop Excel (2026-07-26): `startIndex` is a
    // 0-based character offset and `length` COVERS the leading `@`, so the offsets are only meaningful
    // against the exact message text: drop or shift either and Excel highlights the wrong words.
    threadedCommentMentions: countIn(parts, /threadedComments\//, /<mention\b/g),
    threadedCommentMentionSpans: names
      .filter((p) => /threadedComments\//.test(p))
      .flatMap((p) => [
        ...(parts[p] ?? '').matchAll(
          /<mention\b[^>]*\bstartIndex="([^"]*)"[^>]*\blength="([^"]*)"/g,
        ),
      ])
      .map((m) => `${m[1]}:${m[2]}`)
      .sort(),
    threadedCommentMentionPersonIds: [
      ...new Set(
        names
          .filter((p) => /threadedComments\//.test(p))
          .flatMap((p) => [...(parts[p] ?? '').matchAll(/\bmentionpersonId="([^"]*)"/g)])
          .map((m) => m[1]!),
      ),
    ].sort(),
    // Excel interns a mentioned identity as its OWN `<person>` entry with `providerId="PeoplePicker"`,
    // separate from the same human's `providerId="AD"` authoring entry (verified: Excel rewrote an
    // injected mention to point at a new person id it added on save). So the registry legitimately holds
    // several entries per human, distinguished only by id; collapsing them by name or userId corrupts it.
    personProviderIds: [
      ...new Set(
        names
          .filter((p) => /xl\/persons\/person\d*\.xml$/.test(p))
          .flatMap((p) => [...(parts[p] ?? '').matchAll(/\bproviderId="([^"]*)"/g)])
          .map((m) => m[1]!),
      ),
    ].sort(),
    // How Excel binds a thread to the legacy fallback `<comment>` it writes beside it: the fallback's
    // author is a synthetic `tc={headThreadId}` entry in the comments part's `<authors>`, and the
    // `<comment>` itself carries `xr:uid="{headThreadId}"`. Lose either and Excel stops recognising the
    // cell as threaded: it renders an ordinary note and ignores the thread part, however intact that
    // part still is. So these are the load-bearing halves of "the conversation survived", not trivia.
    commentFallbackThreadAuthors: countIn(
      parts,
      /comments\d+\.xml$/,
      /<author>tc=\{[^}]*\}<\/author>/g,
    ),
    commentFallbackUids: threadFallbackComments(parts).map((fallback) => fallback.uid),
    // The text a pre-2018 reader renders for a thread: fixed boilerplate, then the opening message and
    // each reply. It is regenerated from the thread model rather than carried through, so comparing it
    // is what proves the regeneration matches Excel's own wording and reply layout instead of merely
    // producing something.
    commentFallbackTexts: threadFallbackComments(parts).map((fallback) => fallback.text),
    // Every `<comment>` of the comments part, and every VML shape backing one. A comment with no shape
    // reads as text but renders nothing, so the two counts must agree, and a thread's fallback must
    // not turn into a *second* comment beside a note on the same cell.
    commentEntries: countIn(parts, /comments\d+\.xml$/, /<comment\b/g),
    commentVmlShapes: countIn(parts, /vmlDrawing\d+\.vml$/, /<v:shape\b/g),
    externalLinks: names.filter((p) => /xl\/externalLinks\/externalLink\d+\.xml$/.test(p)).length,
    // The `<externalReference>` registrations in workbook.xml: one per `[n]` a formula resolves an
    // external cell through. Reported as a count (the rel ids are renumbered on write, the ordering and
    // arity are what must survive).
    externalReferenceCount: [...wb.matchAll(/<externalReference\b/g)].length,
    // The `TargetMode="External"` source-workbook pointers carried by every externalLink's own rels:
    // dropping these orphans the link. Sorted so the comparison is order-independent.
    externalTargets: names
      .filter((p) => /xl\/externalLinks\/_rels\/.+\.rels$/.test(p))
      .flatMap((p) => [...(parts[p] ?? '').matchAll(/Target="([^"]*)"\s+TargetMode="External"/g)])
      .map((m) => m[1]!)
      .sort(),
    hasLegacyDrawingHF: /<legacyDrawingHF\b/.test(ws1),
    hasDrawingRef: /<drawing\b/.test(ws1),
    hasHeaderFooterImageToken: /&amp;G|&G/.test(ws1),
    drawingHasShape: /<xdr:sp\b/.test(drawing1),
    drawingHasPicture: /<xdr:pic\b/.test(drawing1),
  };
};
