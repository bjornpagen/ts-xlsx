// The package plumbing every xlsx test needs: unzip a written workbook, reach one part, or feed the
// reader hand-authored markup.
//
// This existed nineteen times across fifteen files, in four families, with three different answers
// to "the part is not there": a `TypeError` from `strFromU8` on undefined, a silent empty string, or
// an assertion naming the part. The silent one was the common spelling in the two files with the
// most negative assertions, where a dozen `assert.doesNotMatch` calls would have gone on passing
// against `''` had the writer ever renamed the part they read. So every accessor here asserts,
// except the one whose whole purpose is absence.
//
// Not a `.test.ts` file: `node --test` would try to run it and find no tests. `tsconfig.build.json`
// excludes the `.test-support.ts` suffix, so it ships nowhere.

import assert from 'node:assert/strict';

import {strFromU8, strToU8, unzipSync, zipSync} from 'fflate';

import {Workbook} from '../../core/workbook.ts';
import {AuthoringError} from '../../errors.ts';
import {relTypeSegment} from '../../rel-type.ts';
import {openElements} from '../../xml/xml-read.ts';
import {PKG_RELS_NS, RELATIONSHIPS_NS} from '../opc/namespaces.ts';
import {relsPathFor, resolveRelativePart} from '../opc/part-paths.ts';
import {parseRelationshipRecords} from '../opc/read-opc.ts';
import {readXlsx} from './read.ts';
import {type WriteOptions, writeXlsx} from './write.ts';

/** The first worksheet's part path, which most writer tests are reaching for. */
export const SHEET1 = 'xl/worksheets/sheet1.xml';

// One unzip per package, not one per accessor. Every accessor below used to inflate the whole
// package to answer for a single part, and a test that checks four parts of one workbook paid for
// four; `read.test.ts` does exactly that repeatedly. Keyed on the package's own bytes, so a test
// that writes twice gets two entries and a test that never re-reads pays nothing, and weakly so
// that a suite building hundreds of workbooks does not hold them all.
//
// The record `unzipSync` returns is the memo itself, so nothing may write to it. `patchParts` is
// the one caller that wants to, and copies first.
const inflated = new WeakMap<Uint8Array, Record<string, Uint8Array>>();

function filesOf(pkg: Uint8Array): Record<string, Uint8Array> {
  const hit = inflated.get(pkg);
  if (hit !== undefined) return hit;
  const files = unzipSync(pkg);
  inflated.set(pkg, files);
  return files;
}

/** Every part of a package, decoded to text, keyed by part path. */
export function partsOf(pkg: Uint8Array): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [name, bytes] of Object.entries(filesOf(pkg))) out[name] = strFromU8(bytes);
  return out;
}

/**
 * Every part of the package a workbook writes to, decoded to text: {@link partsOf} for the many tests
 * that start from a `Workbook` rather than from bytes.
 *
 * Two test files declared this locally, byte-identical down to the comment, each shadowing the
 * imported `partsOf` it wrapped, which is the drift this module exists to stop. The package is checked by
 * {@link assertRelationshipsWired} on the way, so every test that writes through here checks the graph.
 */
export function partsWritten(workbook: Workbook): Record<string, string> {
  const pkg = writeXlsx(workbook);
  assertRelationshipsWired(pkg);
  return partsOf(pkg);
}

/** One named part's text. Fails the test, naming the part, when it is absent. */
export function partText(pkg: Uint8Array, name: string): string {
  return strFromU8(partBytes(pkg, name));
}

/** One named part's bytes, for a binary part. Fails the test, naming the part, when it is absent. */
export function partBytes(pkg: Uint8Array, name: string): Uint8Array {
  const bytes = filesOf(pkg)[name];
  assert.ok(bytes, `expected part ${name}`);
  return bytes;
}

/**
 * The text of the one part whose path matches `pattern`, for a part the writer numbers. Fails the test
 * when no part matches, and when several do: a lookup that silently took the first of two would be
 * asserting on whichever one the zip happened to list first.
 */
export function partMatching(pkg: Uint8Array, pattern: RegExp): string {
  const names = Object.keys(filesOf(pkg)).filter((name) => pattern.test(name));
  const [name] = names;
  assert.ok(
    name !== undefined && names.length === 1,
    `expected exactly one part matching ${String(pattern)}; got ${names.join(', ')}`,
  );
  return partText(pkg, name);
}

/**
 * The first match of `pattern` in some markup. Fails the test, naming the pattern, when there is none.
 *
 * Parts reached this rule and elements did not. The spelling it replaces, `xml.match(re)?.[0] ?? ''`,
 * turns a missing element into an empty string, and every `doesNotMatch` asserted on that string then
 * passes for the wrong reason.
 */
export function elementIn(xml: string, pattern: RegExp, label?: string): string {
  const match = xml.match(pattern);
  assert.ok(match, `expected ${label ?? String(pattern)} in the markup`);
  return match[0];
}

/** {@link elementIn} for the first capture group, when the test wants an attribute list or a value. */
export function captureIn(xml: string, pattern: RegExp, label?: string): string {
  const captured = xml.match(pattern)?.[1];
  assert.ok(captured !== undefined, `expected ${label ?? String(pattern)} in the markup`);
  return captured;
}

/**
 * One named part's text out of an already-decoded {@link partsOf}/{@link partsWritten} record. Fails
 * the test naming the part when it is absent, which {@link partText} does for a package's bytes: a
 * test that already has the whole record should not have to unzip again to get the assertion.
 */
export function partIn(parts: Record<string, string>, name: string): string {
  const xml = parts[name];
  assert.ok(xml !== undefined, `expected part ${name}`);
  return xml;
}

/**
 * One named part's text from an already-unzipped map, or undefined.
 *
 * {@link partIn}'s counterpart, for the test whose claim is about absence. The shape it replaces is
 * `const x = parts[name]; if (x !== undefined) assert.doesNotMatch(x, …)`, which reads as care and
 * is the vacuity this module exists to remove: the assertion it guards passes when the part is gone.
 */
export function optionalPartIn(parts: Record<string, string>, name: string): string | undefined {
  return parts[name];
}

/** One named part's text, or undefined: for a test asserting a part was *not* written. */
export function optionalPartText(pkg: Uint8Array, name: string): string | undefined {
  const bytes = filesOf(pkg)[name];
  return bytes === undefined ? undefined : strFromU8(bytes);
}

/** The first worksheet's XML. */
export function sheetXml(pkg: Uint8Array): string {
  return partText(pkg, SHEET1);
}

/**
 * Write a workbook and read it straight back: the round-trip under test.
 *
 * The helper existed and was imported by two files while ninety-eight call sites spelled it out.
 * That is worse than its being missing: the next author reads the ninety-eight and copies one. The
 * package is checked by {@link assertRelationshipsWired} before it is read back.
 */
export function roundtrip(workbook: Workbook, options?: WriteOptions): Workbook {
  const pkg = writeXlsx(workbook, options);
  assertRelationshipsWired(pkg);
  return readXlsx(pkg);
}

/**
 * Rewrite named parts of a package and return the new bytes, asserting each named part exists first.
 *
 * The other half of what {@link readPatched} does, for a test that patches a package it built itself
 * rather than the stock one, or that wants the bytes rather than the model. Every site doing this by
 * hand read the part it was about to patch through a cast or a `?? new Uint8Array()`, which is the
 * failure this module exists to remove: the first spelling throws a `TypeError` naming nothing when
 * the writer renames a part, and the second silently patches an empty string, after which every
 * negative assertion built on the result passes for the wrong reason.
 *
 * An edit that changes nothing fails too. A `replace` whose search string the writer stopped emitting
 * is a no-op, and the test goes on reading the unpatched package while claiming to read a hostile one.
 */
export function patchParts(
  pkg: Uint8Array,
  edits: Record<string, (xml: string) => string>,
): Uint8Array {
  const files = {...filesOf(pkg)};
  for (const [name, edit] of Object.entries(edits)) {
    const before = strFromU8(partBytes(pkg, name));
    const after = edit(before);
    assert.notEqual(after, before, `edit to ${name} changed nothing`);
    files[name] = strToU8(after);
  }
  return zipSync(files);
}

/**
 * Read an xlsx built from hand-authored parts patched into a written package, keyed by part path, so
 * a case can feed the reader markup the writer itself only produces on round-trip: an Excel-authored
 * x14 extLst block, a foreign dxf table, a docProps element left empty.
 *
 * The base package is a one-sheet workbook with `A1` set, so a test can assert the read succeeded at
 * all rather than passing by reading nothing.
 */
export function readPatched(parts: Record<string, string>): Workbook {
  const files = unzipSync(writeXlsx(sheeted()));
  for (const [name, xml] of Object.entries(parts)) files[name] = strToU8(xml);
  return readXlsx(zipSync(files));
}

// Relationship types a consumer finds by scanning the owner's `.rels` part for the type, so no element
// of the owner cites their id. Named by the Type's last segment; anything else a package declares and
// nothing cites is a relationship the writer recorded for a part it then forgot to reference.
const FOUND_BY_TYPE: ReadonlySet<string> = new Set([
  'officeDocument',
  'core-properties',
  'extended-properties',
  'styles',
  'theme',
  'sharedStrings',
  'sheetMetadata',
  'person',
  'comments',
  'threadedComment',
  'pivotTable',
  'pivotCacheDefinition',
  'vbaProject',
  'chartStyle',
  'chartColorStyle',
]);

// The attributes a part cites a relationship through: the relationships namespace's `id`, `embed`,
// `link` and `pict`, and VML's `o:relid`.
const CITING_ATTRIBUTES: readonly (readonly [namespace: string, local: string])[] = [
  ...['id', 'embed', 'link', 'pict'].map((local) => [RELATIONSHIPS_NS, local] as const),
  ['urn:schemas-microsoft-com:office:office', 'relid'],
];

/**
 * Assert a package's relationship graph is wired both ways: every relationship id a part cites is
 * declared in that part's `.rels`, every internal relationship targets a part the package holds, and
 * every internal relationship is either cited or of a type a consumer finds by scanning for it.
 *
 * The writer used to list a sheet's relationships in one place, hand out their ids in another, and
 * decide whether the `.rels` part existed in a third, so a part could cite an id nothing declared, or
 * declare one nothing used, with every package-level test still green. This checks the result rather
 * than any one of those places.
 */
export function assertRelationshipsWired(pkg: Uint8Array): void {
  const files = filesOf(pkg);
  const problems: string[] = [];
  for (const path of Object.keys(files)) {
    if (!/\.(xml|vml)$/i.test(path) || path === '[Content_Types].xml') continue;
    const relsXml = optionalPartText(pkg, relsPathFor(path));
    const declared = relsXml === undefined ? [] : parseRelationshipRecords(relsXml);
    const ids = new Set(declared.map((rel) => rel.id));

    const cited = new Set<string>();
    for (const {attrs, scope} of openElements(partText(pkg, path))) {
      for (const [namespace, local] of CITING_ATTRIBUTES) {
        const id = scope.attr(attrs, namespace, local);
        if (id !== undefined) cited.add(id);
      }
    }
    for (const id of cited) {
      if (!ids.has(id)) problems.push(`${path} cites ${id}, which its .rels does not declare`);
    }
    for (const rel of declared) {
      if (rel.external) continue;
      const target = resolveRelativePart(path, rel.target);
      // A `#` target is a place in the document (a picture's link to `#Sheet1!C3`), not a part.
      if (!rel.target.startsWith('#') && files[target] === undefined) {
        problems.push(`${path} declares ${rel.id} to ${target}, which the package does not hold`);
      }
      if (!cited.has(rel.id) && !FOUND_BY_TYPE.has(relTypeSegment(rel.type))) {
        problems.push(
          `${path} declares ${rel.id} (${relTypeSegment(rel.type)}), which nothing cites`,
        );
      }
    }
  }
  const rootRels = optionalPartText(pkg, relsPathFor(''));
  for (const rel of rootRels === undefined ? [] : parseRelationshipRecords(rootRels)) {
    if (!rel.external && files[resolveRelativePart('', rel.target)] === undefined) {
      problems.push(`the package root declares ${rel.id} to ${rel.target}, which it does not hold`);
    }
  }
  assert.deepEqual(problems, [], 'every relationship is declared, targets a part, and is used');
}

/** A one-sheet workbook with `A1` set: the least a workbook needs before the writer reaches anything. */
export function sheeted(): Workbook {
  const workbook = new Workbook();
  workbook.addWorksheet('S').getCell('A1').value = 1;
  return workbook;
}

/** Assert that writing {@link sheeted} after `mutate` is refused as an authoring mistake. */
export function refuses(mutate: (workbook: Workbook) => void): void {
  const workbook = sheeted();
  mutate(workbook);
  assert.throws(() => writeXlsx(workbook), AuthoringError);
}

function zipParts(parts: Record<string, string | Uint8Array | undefined>): Uint8Array {
  const files: Record<string, Uint8Array> = {};
  for (const [path, data] of Object.entries(parts)) {
    if (data !== undefined) files[path] = typeof data === 'string' ? strToU8(data) : data;
  }
  return zipSync(files);
}

// The three parts every hand-authored foreign package needs, plus the content-type declaration a
// real one always carries. Not what this library writes -- what a *foreign* generator minimally
// emits: one sheet named `S`, reached by relationship, with an empty `<sheetData>`.
//
// The relationship `Type` is the placeholder `x` on purpose. The reader resolves a worksheet by
// following the `r:id` the sheet list cites, not by matching the relationship's type, and fourteen
// hand-built packages relied on that without any of them saying so.
const FOREIGN_PACKAGE_DEFAULTS: Record<string, string> = {
  '[Content_Types].xml':
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="xml" ContentType="application/xml"/></Types>',
  'xl/workbook.xml':
    '<?xml version="1.0"?><workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>',
  'xl/_rels/workbook.xml.rels':
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>',
  [SHEET1]: '<?xml version="1.0"?><worksheet><sheetData/></worksheet>',
};

/**
 * A package assembled from hand-authored parts: what a *foreign* generator might emit, as distinct
 * from {@link readPatched}, which starts from a package this library wrote. This is how a reader case
 * says "given this markup and nothing else".
 *
 * Fourteen cases in one file built this by hand, four or five entries each, with the two long
 * namespace URIs spelled out every time. The duplication was the smaller cost: each copy also made
 * an undeclared claim about what the reader tolerates -- most omitted `[Content_Types].xml`, none
 * declared a content type for the worksheet -- and those claims were spread across fourteen literals
 * where nobody could see them together.
 *
 * Override any part by path, add any part by path, and pass `undefined` to drop a default one, which
 * is how a case says "and no content types at all". Returns the zipped bytes, since that is what
 * every caller wants: a package to hand to the reader.
 */
export function foreignPackage(
  parts: Record<string, string | Uint8Array | undefined> = {},
): Uint8Array {
  return zipParts({...FOREIGN_PACKAGE_DEFAULTS, ...parts});
}

/** A worksheet part around hand-authored `<row>` markup: the override a reader case makes most. */
export function foreignSheet(rows: string): string {
  return `<?xml version="1.0"?><worksheet><sheetData>${rows}</sheetData></worksheet>`;
}

const OFFICE_RELATIONSHIP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** A `.rels` part around {@link relationship} entries. */
export function relationshipsPart(entries: string): string {
  return `<?xml version="1.0"?><Relationships xmlns="${PKG_RELS_NS}">${entries}</Relationships>`;
}

/** One relationship whose type is an ECMA-376 one, named by its last segment (`drawing`, `image`). */
export function relationship(id: string, type: string, target: string): string {
  return `<Relationship Id="${id}" Type="${OFFICE_RELATIONSHIP}/${type}" Target="${target}"/>`;
}

/**
 * One relationship whose type is Microsoft-namespaced, named by the path after the namespace root
 * (`2007/relationships/slicer`). Only the suffix marks a slicer or a VBA project for preservation, so
 * the year is the caller's to spell.
 */
export function msRelationship(id: string, type: string, target: string): string {
  return `<Relationship Id="${id}" Type="http://schemas.microsoft.com/office/${type}" Target="${target}"/>`;
}

/** A content-types part declaring the workbook and first worksheet, with `extra` entries between. */
export function typedContentTypes(extra = ''): string {
  return (
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    extra +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    '</Types>'
  );
}

/** A namespaced worksheet part with an empty `<sheetData>` and `tail` after it. */
export function typedWorksheet(tail = ''): string {
  return (
    '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    `xmlns:r="${OFFICE_RELATIONSHIP}"><sheetData/>${tail}</worksheet>`
  );
}

/**
 * {@link foreignPackage} with every relationship typed as a real one, namespaces declared, and the
 * package root's relationships present.
 *
 * The placeholder `Type="x"` is fine for a reader that follows `r:id`s. It is not fine for anything
 * that decides by type, which is what preservation does: an unmodelled part rides through because its
 * relationship says drawing, chart or slicer. Overrides work the same way, by path, with `undefined`
 * dropping a default, and a binary part may be given as bytes.
 */
export function typedForeignPackage(
  parts: Record<string, string | Uint8Array | undefined> = {},
): Uint8Array {
  return zipParts({
    '[Content_Types].xml': typedContentTypes(),
    '_rels/.rels': relationshipsPart(relationship('rId1', 'officeDocument', 'xl/workbook.xml')),
    'xl/workbook.xml':
      '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      `xmlns:r="${OFFICE_RELATIONSHIP}"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': relationshipsPart(
      relationship('rId1', 'worksheet', 'worksheets/sheet1.xml'),
    ),
    [SHEET1]: typedWorksheet(),
    ...parts,
  });
}
