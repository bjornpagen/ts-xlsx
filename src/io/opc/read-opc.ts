// The OPC (Open Packaging Conventions) layer of the reader: reading a part's `.rels`, resolving a
// part's declared content type, and walking the transitive closure of parts a preserved reference
// reaches. Every helper here is pure over the inflated package: it takes part text/bytes accessors
// and returns paths or records, touching no Workbook model. The path arithmetic those answers are
// expressed in is `part-paths.ts`, which holds both directions of it.

import {strFromU8} from 'fflate';

import type {PreservedPart, PreservedRelationship} from '../../core/preserved.ts';
import {isRelType} from '../../rel-type.ts';
import {openElements} from '../../xml/xml-read.ts';
import {extensionOf, relsPathFor, resolveRelativePart} from './part-paths.ts';
import {DEFAULT_MAX_UNCOMPRESSED} from './read-options.ts';
import {inflateSpreadsheetPackage} from './sniff-format.ts';

// The two ways a reader reaches into an inflated package: a part's UTF-8-decoded text, or its raw
// bytes. Built once per read (see {@link packageAccessors}) so the buffered and streaming readers
// share one implementation of the lookup rather than each re-declaring the same closures.
export interface PackageAccessors {
  /** A part's decoded text, or undefined when the package holds no such part. Decodes lazily, so a
   * part the reader never asks for is never stringified. */
  partText: (path: string) => string | undefined;
  /** A part's raw bytes, or undefined when the package holds no such part. */
  partBytes: (path: string) => Uint8Array | undefined;
  /** The package's own spelling of the part the other two would answer for `path`, or undefined when
   * there is none: the key an editor replaces or deletes, so it acts on the entry a read found rather
   * than adding a second one beside it. */
  partKey: (path: string) => string | undefined;
}

/** A spreadsheet package opened for reading: inflated under the read bound, its parts bound to
 * accessors, its office document located through the package's own relationship graph, and that
 * document read if it is XML. */
export interface OpenedSpreadsheet {
  /** The inflated parts, for a reader that hands the whole package on to another codec. */
  readonly files: Record<string, Uint8Array>;
  readonly pkg: PackageAccessors;
  /**
   * Where this package keeps its office document, resolved through `_rels/.rels`. Conventionally
   * `xl/workbook.xml` (or `xl/workbook.bin` for a `.xlsb`), but the convention is Excel's habit
   * rather than the format's rule, and a package is free to name the part anything its root
   * relationship points at.
   */
  readonly documentPath: string;
  /** The office document's text, or undefined when the package's document is not XML (a `.xlsb`
   * carries a BIFF12 part instead). Each entry point answers that case for itself: they genuinely
   * want different answers, and the difference is documented where they diverge. */
  readonly workbookXml: string | undefined;
}

/**
 * Open a spreadsheet package: the six-step preamble every reader shares, and in particular the three
 * decisions worth having exactly one of. The inflate bound is a security decision (an unbounded
 * inflate is a zip bomb), *where the office document lives* is a question only the package can
 * answer, and "is that document XML" is the dispatch decision between the two codecs. None should be
 * restated once per entry point.
 *
 * It deliberately stops short of the shared strings and the style table: those are codec-specific
 * (`xl/sharedStrings.xml` against BIFF12's own record stream), and lifting them here would put
 * SpreadsheetML knowledge into the container layer.
 */
export function openSpreadsheetPackage(
  data: Uint8Array,
  maxUncompressedBytes: number | undefined,
): OpenedSpreadsheet {
  const files = inflateSpreadsheetPackage(data, maxUncompressedBytes ?? DEFAULT_MAX_UNCOMPRESSED);
  const pkg = packageAccessors(files);
  // The graph first, the convention only as a fallback: a package whose rels are damaged should still
  // open at the path Excel would have written, and one that simply named its part something else must
  // open at all.
  const documentPath =
    readPartRelationships('', pkg.partText).targetPath(OFFICE_DOCUMENT_REL) ??
    conventionalDocument(pkg);
  // Read as text only when the part is XML. `.xlsb`'s office document is BIFF12, and decoding those
  // bytes as UTF-8 yields a string rather than the `undefined` that dispatches to the binary codec.
  const workbookXml = extensionOf(documentPath) === 'xml' ? pkg.partText(documentPath) : undefined;
  return {files, pkg, documentPath, workbookXml};
}

const OFFICE_DOCUMENT_REL = 'officeDocument';

// Where Excel puts an office document, XML first: not a rule of the format (see
// `OpenedSpreadsheet.documentPath`), but what a package with an unreadable relationship graph is
// searched for, in the order that keeps an `.xlsx` an `.xlsx`.
const CONVENTIONAL_DOCUMENTS = ['xl/workbook.xml', 'xl/workbook.bin'] as const;

// The first conventional path the package actually holds, falling back to the XML one so a package
// holding neither still reports the absence in the terms a reader expects.
function conventionalDocument(pkg: PackageAccessors): string {
  return (
    CONVENTIONAL_DOCUMENTS.find((path) => pkg.partBytes(path) !== undefined) ??
    CONVENTIONAL_DOCUMENTS[0]
  );
}

// Bind the part-lookup accessors over an inflated package (a part-path → bytes map).
//
// Every member is declared as a function-typed property rather than with method syntax, and each
// is written as an arrow here, because every reader destructures them off the returned object:
// `const {partText, partBytes} = packageAccessors(files)`. Method syntax would say these values
// may read `this`, which they never do (they close over `files`), and would make each of those
// fifteen destructurings report as an unbound method. Property syntax is also the stricter
// declaration: method-syntax parameters are checked bivariantly even under `strictFunctionTypes`.
export function packageAccessors(files: Record<string, Uint8Array>): PackageAccessors {
  // OPC compares part names case-insensitively (ASCII), so `XL/Workbook.xml` and `xl/workbook.xml`
  // name one part, and a package written by a producer that cased them differently from Excel used to
  // read as a package with no workbook at all. The fold is for *lookup* only: `files` keeps the
  // package's own spelling, which is what a preserved part is re-emitted under. Exact-first keeps the
  // ordinary path a single map read and keeps two entries differing only in case resolving to the one
  // actually asked for; the folded map holds the first of them in package order, which is a choice
  // only a package no OPC writer can produce ever notices.
  const folded = new Map<string, string>();
  for (const path of Object.keys(files)) {
    const key = path.toLowerCase();
    if (!folded.has(key)) folded.set(key, path);
  }
  const partKey = (path: string): string | undefined =>
    files[path] === undefined ? folded.get(path.toLowerCase()) : path;
  const lookup = (path: string): Uint8Array | undefined => {
    const key = partKey(path);
    return key === undefined ? undefined : files[key];
  };
  return {
    partText: (path: string): string | undefined => {
      const bytes = lookup(path);
      return bytes === undefined ? undefined : strFromU8(bytes);
    },
    partBytes: lookup,
    partKey,
  };
}

// A relationship as declared, with the fields a preserved-part closure needs: its id, Type URI,
// Target, and whether the target lies outside the package (`TargetMode="External"`). This is the one
// shape a `.rels` part is read into; the narrower views above and below are projections of it, so the
// element is scanned once and there is one answer to what counts as a relationship.
export interface RelationshipRecord {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly external: boolean;
}

// OPC makes Id, Type and Target all mandatory, so an element missing any of the three is not a
// relationship and is skipped rather than half-read.
export function parseRelationshipRecords(xml: string): RelationshipRecord[] {
  const records: RelationshipRecord[] = [];
  for (const {attrs} of openElements(xml, 'Relationship')) {
    if (attrs.Id !== undefined && attrs.Type !== undefined && attrs.Target !== undefined) {
      records.push({
        id: attrs.Id,
        type: attrs.Type,
        target: attrs.Target,
        external: attrs.TargetMode === 'External',
      });
    }
  }
  return records;
}

// One part's `.rels`, parsed once and then queried many times, with the owning part's path bound in so
// every answer comes back as a package path rather than a target still needing resolution.
//
// A worksheet is the reason this exists. Its rels part is the index to nearly everything hanging off the
// sheet (notes, threads, printer settings, the drawing, the background image, tables, pivots, and the
// preserved-reference closure) and each of those lookups used to re-read and re-parse the same XML,
// eight times per sheet on a workbook of any size.
export interface PartRelationships {
  /** Every relationship the part declares, in declaration order. */
  readonly records: readonly RelationshipRecord[];
  /** The relationship with this id, or undefined. Its `target` is raw: an external one is a URL, not a
   * package path, so a caller that may see `TargetMode="External"` must read it before resolving. */
  byId(id: string): RelationshipRecord | undefined;
  /** Resolve one of this part's targets against the part's own directory. */
  pathOf(target: string): string;
  /** The package part reached through the first relationship of this class ({@link isRelType}), or
   * undefined when the part declares none: the single-part lookup (notes, printer settings, drawing,
   * background) in one call. */
  targetPath(name: string): string | undefined;
  /** Every package part reached through a relationship of this class, in declaration order. For a part
   * class one sheet may reference more than once (tables, pivot tables). */
  targetPaths(name: string): string[];
  /**
   * The text of the single part reached through a relationship of this type, or `undefined` when the
   * part declares no such relationship *or* names one the package does not contain.
   *
   * Those two are deliberately one answer. A relationship pointing at an absent part is a damaged
   * package, and every reader here answers a damaged package the same way Excel does: the feature is
   * simply not there, rather than the load failing over it. Collapsing `targetPath` then `partText`
   * into one call is what lets that contract be stated once instead of at each of the seven places
   * that used to spell it out.
   */
  relatedText(name: string): string | undefined;
  /** As {@link relatedText}, for a part whose content is opaque bytes (a printer-settings blob, an
   * image) rather than XML. */
  relatedBytes(name: string): Uint8Array | undefined;
  /**
   * The text of a part reached through the relationship that names it, falling back to
   * `conventionalPath` only when the part declares no such relationship, and to `''` when neither
   * resolves.
   *
   * The order is the whole point, and it decides more than where a lookup lands. A package is free
   * to name any part anything -- the workbook part included -- so the part's own relationships are
   * what say where its pool and its stylesheet live. Resolved conventional-path-first, a renamed
   * pool reads as no pooled strings at all, and a renamed stylesheet silently changes cell *types*,
   * because the date test reads `numFmt` off the resolved style to tell `45000` from a date. The
   * conventional path stays as the fallback for a package whose rels are damaged.
   *
   * The `''` is a real package state, not a lookup that failed quietly: a workbook may carry no pool
   * and no stylesheet, and every reader of these parses the empty text into the empty table it would
   * have built anyway. Stated here so the two readers that resolve these parts cannot answer the
   * same package differently -- a streamed read must not decode a cell differently from a buffered
   * one.
   */
  relatedTextOrPath(name: string, conventionalPath: string): string;
}

// Read and parse a part's `.rels`. A part with no rels part yields an empty set rather than undefined,
// so a caller never has to distinguish "no relationships" from "no rels part": nothing downstream
// treats those two differently.
export function readPartRelationships(
  partPath: string,
  partText: (path: string) => string | undefined,
  partBytes?: (path: string) => Uint8Array | undefined,
): PartRelationships {
  const records = parseRelationshipRecords(partText(relsPathFor(partPath)) ?? '');
  const byId = new Map(records.map((record) => [record.id, record]));
  const pathOf = (target: string): string => resolveRelativePart(partPath, target);
  const targetsOf = (name: string): string[] =>
    records.filter((record) => isRelType(record.type, name)).map((record) => record.target);
  const targetPath = (name: string): string | undefined => targetsOf(name).map(pathOf)[0];
  return {
    records,
    byId: (id) => byId.get(id),
    pathOf,
    targetPath,
    targetPaths: (name) => targetsOf(name).map(pathOf),
    relatedText: (name) => {
      const path = targetPath(name);
      return path === undefined ? undefined : partText(path);
    },
    // `partBytes` is optional because the two readers that only ever ask for XML (the streaming row
    // reader, the xlsb workbook part) have no bytes accessor to hand over; asking one of them for a
    // binary part is a bug in that reader rather than a package that lacks it.
    relatedBytes: (name) => {
      const path = targetPath(name);
      return path === undefined ? undefined : partBytes?.(path);
    },
    relatedTextOrPath: (name, conventionalPath) => {
      const path = targetPath(name);
      return (path === undefined ? undefined : partText(path)) ?? partText(conventionalPath) ?? '';
    },
  };
}

// Resolve a package part path to its declared content type the way OPC does: an `<Override>` naming
// the exact part wins, else the `<Default>` registered for the part's extension. An unknown part
// falls back to the generic binary type so re-declaring it never emits an empty content type.
export function contentTypeResolver(contentTypesXml: string): (path: string) => string {
  const overrides = new Map<string, string>();
  const defaults = new Map<string, string>();
  for (const {local, attrs} of openElements(contentTypesXml, 'Override', 'Default')) {
    if (local === 'Override' && attrs.PartName !== undefined && attrs.ContentType !== undefined) {
      // Folded on both sides, as the `<Default Extension>` key below already was: the two halves of
      // one resolver disagreeing about case meant an override whose spelling differed from the zip
      // entry's fell through to the generic binary type, and a preserved part is re-declared under
      // whatever this returns.
      overrides.set(attrs.PartName.toLowerCase(), attrs.ContentType);
    } else if (
      local === 'Default' &&
      attrs.Extension !== undefined &&
      attrs.ContentType !== undefined
    ) {
      defaults.set(attrs.Extension.toLowerCase(), attrs.ContentType);
    }
  }
  return (path: string): string =>
    overrides.get(`/${path}`.toLowerCase()) ??
    defaults.get(extensionOf(path)) ??
    'application/octet-stream';
}

// Gather the transitive closure of package parts reachable from an entry part: the part itself, then
// every internal part its relationships target, breadth-first, each with its raw bytes, content type,
// and (internal) relationships. Returns undefined when the entry part is absent (a dangling reference
// preserves nothing). A `visited` set dedupes shared parts and bounds the walk to the (finite,
// inflate-capped) package, so a maliciously self-referential rels graph cannot loop.
export function capturePartClosure(
  entryPath: string,
  partText: (path: string) => string | undefined,
  partBytes: (path: string) => Uint8Array | undefined,
  contentTypeOf: (path: string) => string,
): readonly PreservedPart[] | undefined {
  const parts: PreservedPart[] = [];
  const visited = new Set<string>();
  // Walked with a cursor rather than `shift()`, which is O(n) per dequeue and made the walk quadratic
  // in reachable parts. Bounded by the inflate cap either way, so it never mattered for an ordinary
  // file, but this walks an untrusted package's relationship graph and a hostile one chooses the n.
  const queue: string[] = [entryPath];
  for (let next = 0; next < queue.length; next++) {
    const path = queue[next];
    if (path === undefined || visited.has(path)) continue;
    visited.add(path);
    const bytes = partBytes(path);
    if (bytes === undefined) continue;
    const relsXml = partText(relsPathFor(path));
    const rels: PreservedRelationship[] = [];
    if (relsXml !== undefined) {
      for (const rel of parseRelationshipRecords(relsXml)) {
        if (rel.external) {
          // A linked workbook lives outside the package: keep the wiring verbatim (an externalLink
          // part's pointer to its source), but do not walk into it: there is no package part to visit.
          rels.push({id: rel.id, type: rel.type, targetPath: rel.target, external: true});
          continue;
        }
        const targetPath = resolveRelativePart(path, rel.target);
        rels.push({id: rel.id, type: rel.type, targetPath});
        queue.push(targetPath);
      }
    }
    parts.push({path, contentType: contentTypeOf(path), bytes, rels});
  }
  return parts.some((part) => part.path === entryPath) ? parts : undefined;
}
