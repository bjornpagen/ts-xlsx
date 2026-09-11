// Structurally edit an existing `.xlsm` at the *package* level: remove a module, or add a reference,
// returning new package bytes with every other part preserved byte-for-byte. Only `xl/vbaProject.bin` is
// rewritten (plus dropping a now-stale signature); worksheets, styles, drawings, and every other part
// ride through untouched.
//
// This is the highest-fidelity way to tweak an existing macro project. The alternative (`readXlsx` →
// the matching `Workbook` method → `writeXlsx`) rebuilds the whole package from the parsed model, which
// re-serialises every part and so only preserves what the model captures. For a rich, real-world
// workbook that round-trip can perturb parts Excel is strict about; splicing the original bytes cannot,
// because it never re-authors anything but the macro project. Use this when the input is a real file
// whose non-macro content must be preserved exactly.
//
// Authoring or editing module SOURCE is not here: Excel runs a module's compiled p-code, not its source,
// so that needs the offline `tools/vba-compiler` (VBIDE), which can produce a whole edited `.xlsm`
// directly (ADR 0019).

import {strToU8, zipSync} from 'fflate';

import {VbaAuthorError} from '../../vba/errors.ts';
import {
  addVbaReference,
  removeVbaModule,
  type VbaLibraryReference,
} from '../../vba/project-editor.ts';
import {elementRange} from '../../xml/xml-read.ts';
import type {XmlAttributes} from '../../xml/xml-scan.ts';
import {relsPathFor, resolveRelativePart} from '../opc/part-paths.ts';
import {
  type PackageAccessors,
  packageAccessors,
  parseRelationshipRecords,
  readPartRelationships,
} from '../opc/read-opc.ts';
import {DEFAULT_MAX_UNCOMPRESSED, type ReadPackageOptions} from '../opc/read-options.ts';
import {inflateSpreadsheetPackage} from '../opc/sniff-format.ts';
import {FIXED_ENTRY_MTIME} from '../opc/zip-mtime.ts';

const OFFICE_DOCUMENT_REL = 'officeDocument';
const VBA_PROJECT_REL = 'vbaProject';
// Every signature flavour Excel writes over a VBA project (legacy, agile, V3) shares this local-name
// prefix; all become stale the instant the project's bytes change and must be dropped with it.
const VBA_SIGNATURE_REL_INFIX = 'vbaProjectSignature';

/**
 * Remove a standard module from an existing macro-enabled package's VBA project, returning new package
 * bytes. Every part but `xl/vbaProject.bin` is preserved byte-for-byte (see {@link removeVbaModule} for
 * what changes within it), and any digital signature over the old project is dropped because it cannot
 * validate the new bytes.
 *
 * @throws {VbaAuthorError} if the package carries no VBA project, `name` is not in the project, or names
 *   a `document`/`designer` module.
 * @throws {VbaParseError} if the attached `vbaProject.bin` is malformed.
 * @throws {PackageReadError} if the input is not a readable ZIP, or exceeds the inflate bound
 *   ({@link ReadPackageOptions.maxUncompressedBytes}, defaulting as `readXlsx` does).
 */
export function editXlsxVbaRemoveModule(
  xlsx: Uint8Array,
  name: string,
  options: ReadPackageOptions = {},
): Uint8Array {
  return applyToVbaProjectPart(xlsx, options, (bin) => removeVbaModule(bin, name));
}

/**
 * Add a registered (COM type-library) reference to an existing macro-enabled package's VBA project,
 * returning new package bytes. Every part but `xl/vbaProject.bin` is preserved byte-for-byte (see
 * {@link addVbaReference} for what changes within it), and any digital signature over the old project is
 * dropped because it cannot validate the new bytes.
 *
 * @throws {VbaAuthorError} if the package carries no VBA project, or any field of `ref` is invalid (see
 *   {@link VbaLibraryReference}).
 * @throws {VbaParseError} if the attached `vbaProject.bin` is malformed.
 * @throws {PackageReadError} if the input is not a readable ZIP, or exceeds the inflate bound
 *   ({@link ReadPackageOptions.maxUncompressedBytes}, defaulting as `readXlsx` does).
 */
export function editXlsxVbaAddReference(
  xlsx: Uint8Array,
  ref: VbaLibraryReference,
  options: ReadPackageOptions = {},
): Uint8Array {
  return applyToVbaProjectPart(xlsx, options, (bin) => addVbaReference(bin, ref));
}

// Shared plumbing for every package-level VBA edit: unzip, locate `xl/vbaProject.bin`, replace it with
// whatever `apply` produces, drop a now-stale signature, and re-zip. `apply` is expected to validate
// fail-closed itself (every project-editor primitive does), so a bad edit throws before `files` is
// touched.
function applyToVbaProjectPart(
  xlsx: Uint8Array,
  options: ReadPackageOptions,
  apply: (bin: Uint8Array) => Uint8Array,
): Uint8Array {
  // Through the shared inflater, not `unzipSync`. These two functions take raw caller-supplied bytes
  // and are exported from the package entry, so they are readers, and `inflatePackage`'s header calls
  // decompression "the reader's first hostile-input surface" for the reason that applies here too: an
  // uncapped `unzipSync` believes whatever the archive expands to. It also hands back a null-prototype
  // map already, which is why nothing re-copies one here.
  const files = inflateSpreadsheetPackage(
    xlsx,
    options.maxUncompressedBytes ?? DEFAULT_MAX_UNCOMPRESSED,
  );

  // Every part is found the way `readXlsx` finds it, case folded, and then replaced or deleted under
  // the package's own spelling of it. An exact-key lookup refused a package the reader opens (an entry
  // `xl/VbaProject.bin` under a target `vbaProject.bin`) as having no project, and a delete by the
  // resolved spelling left the real entry beside a new one.
  const pkg = packageAccessors(files);
  const binPath = locateVbaProjectPart(pkg);
  const binKey = binPath === undefined ? undefined : pkg.partKey(binPath);
  const bin = binKey === undefined ? undefined : files[binKey];
  if (binPath === undefined || binKey === undefined || bin === undefined) {
    throw new VbaAuthorError('package has no VBA project to edit');
  }

  files[binKey] = apply(bin);
  dropStaleSignature(files, pkg, binPath);

  // Re-stamped rather than preserved: `unzipSync` hands back bytes and drops each entry's original
  // timestamp, so there is nothing to carry through: the choice is a pinned stamp or the clock, and
  // the clock would make editing the same file twice produce two different packages.
  return zipSync(files, {mtime: FIXED_ENTRY_MTIME});
}

// Resolve the package's `xl/vbaProject.bin` part the way the reader does: `_rels/.rels` → the
// officeDocument (workbook) part → its `.rels` → the `vbaProject` relationship, each target resolved
// relative to its referrer. undefined when the package declares no such relationship (a macro-free book).
function locateVbaProjectPart(pkg: PackageAccessors): string | undefined {
  const workbookPath = readPartRelationships('', pkg.partText).targetPath(OFFICE_DOCUMENT_REL);
  if (workbookPath === undefined) return undefined;
  return readPartRelationships(workbookPath, pkg.partText).targetPath(VBA_PROJECT_REL);
}

// Editing the project invalidates any signature over it, so remove every signature part the project's
// `.rels` reaches, the relationships that point at them, and their content-type overrides, leaving a
// package that advertises no signature rather than a broken one (mirrors Workbook.vbaProjectBytes).
//
// `pkg` was bound before the project was replaced and still answers for the edited package: it reads
// bytes out of `files` when asked, and replacing the project kept the entry's key.
function dropStaleSignature(
  files: Record<string, Uint8Array>,
  pkg: PackageAccessors,
  binPath: string,
): void {
  const binRelsKey = pkg.partKey(relsPathFor(binPath));
  const binRels = binRelsKey === undefined ? undefined : pkg.partText(binRelsKey);
  if (binRelsKey === undefined || binRels === undefined) return;

  const signatureRels = parseRelationshipRecords(binRels).filter(
    (rel) => !rel.external && rel.type.includes(VBA_SIGNATURE_REL_INFIX),
  );
  if (signatureRels.length === 0) return;

  const contentTypesKey = pkg.partKey(CONTENT_TYPES_PART);
  let contentTypes = contentTypesKey === undefined ? undefined : pkg.partText(contentTypesKey);
  let rels = binRels;
  for (const rel of signatureRels) {
    const partPath = resolveRelativePart(binPath, rel.target);
    const partKey = pkg.partKey(partPath);
    if (partKey !== undefined) delete files[partKey];
    const partName = `/${partPath}`.toLowerCase();
    rels = removeElements(rels, 'Relationship', (attrs) => attrs.Id === rel.id);
    if (contentTypes !== undefined) {
      contentTypes = removeElements(
        contentTypes,
        'Override',
        (attrs) => attrs.PartName?.toLowerCase() === partName,
      );
    }
  }

  if (parseRelationshipRecords(rels).length === 0) delete files[binRelsKey];
  else files[binRelsKey] = strToU8(rels);
  if (contentTypesKey !== undefined && contentTypes !== undefined) {
    files[contentTypesKey] = strToU8(contentTypes);
  }
}

const CONTENT_TYPES_PART = '[Content_Types].xml';

// Drop every `local` element whose decoded attributes satisfy `drop`, splicing at the offsets the
// scanner reports so everything else passes through byte for byte. A pattern over the raw text could
// only see one quoting style and one spelling of an attribute, and left a single-quoted `Id` or
// `PartName` pointing at a part the edit had just deleted. Each scan starts where the previous element
// ended, so the whole pass is one walk of the part however many elements it drops.
function removeElements(
  xml: string,
  local: string,
  drop: (attrs: XmlAttributes) => boolean,
): string {
  let kept = '';
  let rest = xml;
  let range = elementRange(rest, [local]);
  while (range !== undefined) {
    kept += rest.slice(0, drop(range.attrs) ? range.start : range.end);
    rest = rest.slice(range.end);
    range = elementRange(rest, [local]);
  }
  return kept + rest;
}
