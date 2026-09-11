// Structural edits to an existing `vbaProject.bin` that do NOT touch any module's compiled p-code:
// remove a standard module, or add a registered library reference. Each is a surgical splice: parse the
// original container, rebuild its whole storage/stream tree, and change only the `dir` records (and, for
// a removal, the `PROJECT`/`PROJECTwm` text) the edit demands. Every module stream, `_VBA_PROJECT`, and
// every untouched record rides through byte-for-byte; preservation is by *not touching* them.
//
// Authoring or editing module SOURCE is deliberately NOT here. Excel does not recompile from source on
// open. A module runs the p-code it ships, and only a real Excel can produce genuinely source-matched
// p-code, so source authoring/editing lives in the offline `tools/vba-compiler` (VBIDE), not in this
// pure-TS path (ADR 0019). These splices are safe precisely because they leave every module's p-code
// exactly as its own compiler wrote it.

import {utf16leBytes} from '../bytes.ts';
import {InternalError, quoted} from '../errors.ts';
import {hex} from '../hex.ts';
import {concat, readU16, spliceBytes, writeU16, writeU32} from './bytes.ts';
import {sameEntryName} from './cfb-format.ts';
import {type CfbNode, isStream, writeCompoundFile} from './cfb-writer.ts';
import {CompoundFile} from './cfb.ts';
import {type Decoder, decoderForCodePage, type Encoder, encoderForCodePage} from './codepage.ts';
import {
  dirRecord,
  dirRecords,
  REC_MODULE_NAME,
  REC_MODULE_STREAMNAME,
  REC_MODULE_TERMINATOR,
  REC_MODULES_COUNT,
  REC_REFERENCE_NAME,
  REC_REFERENCE_NAME_UNICODE,
  REC_REFERENCE_REGISTERED,
} from './dir-records.ts';
import {VbaAuthorError, VbaParseError} from './errors.ts';
import {compressContainer, decompressContainer} from './ms-ovba.ts';
import {DIR_PATH, parseVbaProjectIn, PROJECT_PATH, PROJECTWM_PATH, VBA_STORAGE} from './project.ts';
import {validateReferenceName} from './vba-encoding.ts';

/**
 * Remove a standard module from an existing `vbaProject.bin`, returning new bytes that carry every
 * remaining module, reference, and host-info record unchanged. It drops the module's `VBA/<name>`
 * stream, its MODULE record block in `dir` (decrementing `MODULES_COUNT`), and its `Module=`/`Class=` +
 * workspace lines in `PROJECT`/`PROJECTwm`.
 *
 * Only `procedural` and `class` modules can be removed this way. Removing a `document` module (e.g.
 * `ThisWorkbook`) or a `designer` module (a UserForm) would leave the host referencing code that no
 * longer exists, since their names are tied to a worksheet/workbook `codeName` or a designer storage
 * this project-level primitive has no visibility into. Editing such a module's code-behind is a job for
 * the offline `tools/vba-compiler` (in-place mode), which drives the real host.
 *
 * @throws {VbaParseError} if `bin` is not a parseable VBA project (validated before any edit).
 * @throws {VbaAuthorError} if `name` is not in the project, or names a `document`/`designer` module.
 */
export function removeVbaModule(bin: Uint8Array, name: string): Uint8Array {
  // Parse fail-closed first: validates the container and resolves the module's kind/stream name, so
  // nothing is mutated on a bad input or an unsupported module kind. The container is opened once and
  // shared with the parse, rather than built again below over the same bytes.
  const cfb = new CompoundFile(bin);
  const project = parseVbaProjectIn(cfb);
  const nameKey = name.toUpperCase(); // VBA names are case-insensitive
  const module = project.modules.find((m) => m.name.toUpperCase() === nameKey);
  if (!module) throw new VbaAuthorError(`module ${quoted(name)} is not in the VBA project`);
  if (module.kind !== 'procedural' && module.kind !== 'class') {
    throw new VbaAuthorError(
      `cannot remove module ${quoted(name)}: its kind ${quoted(module.kind)} is tied to host linkage this ` +
        'primitive cannot verify',
    );
  }

  // One decoder for the whole removal, as `removeProjectwmRecord` already took. Two calls built two
  // from the same code page, which is a `TextDecoder` allocated to read a handful of record names.
  const decoder = decoderForCodePage(project.codePage);
  const dirCompressed = cfb.readStream(DIR_PATH);
  if (!dirCompressed) throw new VbaParseError("VBA project has no 'dir' stream");
  const patchedDir = removeModuleDirRecord(
    decompressContainer(dirCompressed),
    module.streamName,
    decoder,
  );

  // Leave _VBA_PROJECT untouched. Resetting it to an "unmatchable version" cookie does NOT force Excel
  // to recompile from source (Excel runs the p-code as-is); on a project that carries real p-code the
  // reset actively crashes the VBA load (verified 2026-07-24, ADR 0019). The surviving modules keep
  // their own compiled p-code; the `dir` stream, authoritative for the module list, no longer names
  // the removed module, which is what makes the removal take.
  const replacements: StreamReplacement[] = [{path: DIR_PATH, data: compressContainer(patchedDir)}];

  const projectText = cfb.readStream(PROJECT_PATH);
  if (projectText) {
    replacements.push({
      path: PROJECT_PATH,
      data: removeProjectStreamLines(projectText, module.name, module.kind, decoder),
    });
  }
  const projectwm = cfb.readStream(PROJECTWM_PATH);
  if (projectwm) {
    replacements.push({
      path: PROJECTWM_PATH,
      data: removeProjectwmRecord(projectwm, project.modules.length, module.name, decoder),
    });
  }

  const newTree = withoutStream(replaceStreams(cfb.tree(), replacements), [
    VBA_STORAGE,
    module.streamName,
  ]);
  return writeCompoundFile(newTree);
}

/**
 * A registered (COM Automation type-library) reference to add to an existing VBA project: the shape of
 * a real "add a reference to Microsoft Scripting Runtime" call. Project references (to another VBA
 * project) and control references (to an ActiveX control library) are out of scope. See
 * {@link addVbaReference}.
 */
export interface VbaLibraryReference {
  /**
   * The reference's namespace name in the VBA editor: what a qualified reference like
   * `Scripting.Dictionary` resolves through. Must be a valid VBA identifier, at most 31 characters, as
   * real type libraries use (e.g. `Scripting`, `Office`, `stdole`).
   */
  readonly name: string;
  /**
   * The friendly name shown in the References dialog, e.g. `Microsoft Scripting Runtime`. Real projects
   * usually keep this distinct from {@link name}; defaults to {@link name} if omitted.
   */
  readonly displayName?: string;
  /** The type library's GUID, e.g. `{420B2830-E718-11CF-893D-00A0C9054228}` (braces optional). */
  readonly guid: string;
  /** The type library's major version, an integer in `[0, 0xFFFF]` ([MS-OVBA] `LibidMajorVersion`). */
  readonly majorVersion: number;
  /** The type library's minor version, an integer in `[0, 0xFFFF]` ([MS-OVBA] `LibidMinorVersion`). */
  readonly minorVersion: number;
  /**
   * The type library's LCID, an integer in `[0, 0xFFFFFFFF]`. Defaults to `0` (locale-neutral), the
   * overwhelming common case (every reference in a real project observed while building this had `0`).
   */
  readonly lcid?: number;
  /** Absolute Windows path to the type library file, e.g. `C:\Windows\System32\scrrun.dll`. */
  readonly path: string;
}

// LibidMajorVersion/LibidMinorVersion ([MS-OVBA] 2.1.1.8): 1*4HEXDIG, so at most 0xFFFF.
const MAX_LIBID_VERSION = 0xffff;
// LibidLcid: 1*8HEXDIG, so at most 0xFFFFFFFF (practically always 0, locale-neutral).
const MAX_LIBID_LCID = 0xffffffff;
// LibidRegName: *255(%x01-FF), so at most 255 bytes, never NUL.
const MAX_DISPLAY_NAME_CHARS = 255;
// LibidPath is bounded by nothing in [MS-OVBA] but SizeOfLibid, a u32, which is no bound at all on a
// string built in memory. 32,767 characters is the longest path Windows can open (the extended-length
// `\\?\` limit), so a longer one names no type library a host could ever load.
const MAX_LIBID_PATH_CHARS = 32_767;
const GUID_PATTERN =
  /^\{?([0-9A-Fa-f]{8})-([0-9A-Fa-f]{4})-([0-9A-Fa-f]{4})-([0-9A-Fa-f]{4})-([0-9A-Fa-f]{12})\}?$/;

interface NormalizedReference {
  readonly name: string;
  readonly libid: string;
}

// Validate every field fail-closed and assemble the Libid string ([MS-OVBA] 2.1.1.8 LibidReference ABNF:
// `*\G{GUID}#Major.Minor#LCID#Path#RegName`, hex digit strings with no `0x` prefix), confirmed
// byte-for-byte against a real Excel-authored reference (2026-07-23):
// `*\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\Windows\System32\scrrun.dll#Microsoft Scripting Runtime`.
function normalizeReference(ref: VbaLibraryReference): NormalizedReference {
  validateReferenceName(ref.name);

  const guidMatch = GUID_PATTERN.exec(ref.guid.trim());
  if (!guidMatch) throw new VbaAuthorError(`invalid reference GUID ${quoted(ref.guid)}`);
  const guid = `{${guidMatch.slice(1, 6).join('-').toUpperCase()}}`;

  for (const [field, value] of [
    ['majorVersion', ref.majorVersion],
    ['minorVersion', ref.minorVersion],
  ] as const) {
    if (!Number.isInteger(value) || value < 0 || value > MAX_LIBID_VERSION) {
      throw new VbaAuthorError(
        `reference ${field} must be an integer in [0, 0xFFFF], got ${value}`,
      );
    }
  }
  const lcid = ref.lcid ?? 0;
  if (!Number.isInteger(lcid) || lcid < 0 || lcid > MAX_LIBID_LCID) {
    throw new VbaAuthorError(`reference lcid must be an integer in [0, 0xFFFFFFFF], got ${lcid}`);
  }

  // Measured before the path is quoted into any message: a refusal should not carry the megabyte it refuses.
  if (ref.path.length > MAX_LIBID_PATH_CHARS) {
    throw new VbaAuthorError(
      `reference path is ${ref.path.length} characters long, past the ${MAX_LIBID_PATH_CHARS} a Windows path can hold`,
    );
  }
  if (ref.path.length === 0 || ref.path.includes('\0') || ref.path.includes('#')) {
    throw new VbaAuthorError(
      `invalid reference path ${quoted(ref.path)} (must be non-empty and contain no NUL or '#')`,
    );
  }
  const displayName = ref.displayName ?? ref.name;
  if (
    displayName.length === 0 ||
    displayName.length > MAX_DISPLAY_NAME_CHARS ||
    displayName.includes('\0')
  ) {
    throw new VbaAuthorError(`invalid reference display name ${quoted(displayName)}`);
  }

  const libid =
    `*\\G${guid}#${hex(ref.majorVersion, 1)}.${hex(ref.minorVersion, 1)}#${hex(lcid, 1)}` +
    `#${ref.path}#${displayName}`;
  return {name: ref.name, libid};
}

/**
 * Add a registered (COM type-library) reference to an existing `vbaProject.bin`, returning new bytes
 * that carry every existing module, reference, and host-info record unchanged. It grows the project's
 * `dir` stream by one `REFERENCENAME` + `REFERENCEREGISTERED` record pair, positioned immediately before
 * `MODULES_COUNT` (references have no count field of their own; `MODULES_COUNT` simply marks where the
 * reference array ends). It needs no change to `PROJECT`/`PROJECTwm`: a real Excel-authored `PROJECT`
 * stream carries no `Reference=` line at all: references live only in `dir` (confirmed against a genuine
 * Excel-authored project).
 *
 * @throws {VbaParseError} if `bin` is not a parseable VBA project (validated before any edit).
 * @throws {VbaAuthorError} if any field of `ref` is invalid (see {@link VbaLibraryReference}), or the
 *   assembled reference text has a character the project's code page cannot represent.
 */
export function addVbaReference(bin: Uint8Array, ref: VbaLibraryReference): Uint8Array {
  const normalized = normalizeReference(ref);

  // Parse fail-closed first: validates the container before any mutation, over the one container this
  // function opens rather than a second built over the same bytes.
  const cfb = new CompoundFile(bin);
  const encode = encoderForCodePage(parseVbaProjectIn(cfb).codePage);

  const dirCompressed = cfb.readStream(DIR_PATH);
  if (!dirCompressed) throw new VbaParseError("VBA project has no 'dir' stream");
  const records = buildReferenceDirRecords(normalized, encode);
  const patchedDir = insertReferenceDirRecords(decompressContainer(dirCompressed), records);

  // Leave _VBA_PROJECT untouched; see the note in removeVbaModule. The new reference is unused by the
  // existing modules' p-code, so they load and run unchanged; only the `dir` reference array grows.
  return writeCompoundFile(
    replaceStreams(cfb.tree(), [{path: DIR_PATH, data: compressContainer(patchedDir)}]),
  );
}

// Build the REFERENCENAME + REFERENCEREGISTERED record bytes ([MS-OVBA] 2.3.4.2.2.2 / .2.2.5) for one
// reference. REFERENCENAME's MBCS/Unicode name pair mirrors MODULE_NAME/MODULE_NAME_UNICODE's shape;
// REFERENCEREGISTERED is one record carrying SizeOfLibid + Libid + two zero Reserved fields.
function buildReferenceDirRecords(ref: NormalizedReference, encode: Encoder): Uint8Array {
  const libid = encode(ref.libid);
  const sizeOfLibid = new Uint8Array(4);
  writeU32(sizeOfLibid, 0, libid.length);
  const reserved = new Uint8Array(6); // Reserved1 (u32) and Reserved2 (u16), both zero
  return concat([
    dirRecord(REC_REFERENCE_NAME, encode(ref.name)),
    dirRecord(REC_REFERENCE_NAME_UNICODE, utf16leBytes(ref.name)),
    dirRecord(REC_REFERENCE_REGISTERED, concat([sizeOfLibid, libid, reserved])),
  ]);
}

// Insert new reference dir records right before MODULES_COUNT (0x000f). The reference array has no
// explicit count field; MODULES_COUNT is simply the next record once the last reference ends (confirmed
// against a real Excel-authored dir stream). Every other record, other references and all modules, rides
// through unchanged.
function insertReferenceDirRecords(dir: Uint8Array, records: Uint8Array): Uint8Array {
  let insertAt = -1;
  for (const {id, recordStart} of dirRecords(dir, 'overruns while adding a reference')) {
    if (id === REC_MODULES_COUNT) {
      insertAt = recordStart;
      break;
    }
  }
  if (insertAt < 0) throw new VbaParseError('dir stream is missing MODULES_COUNT');

  return spliceBytes(dir, insertAt, insertAt, records);
}

// Remove one module's MODULE record block from a decompressed `dir` stream, and decrement MODULES_COUNT.
// A block runs from its MODULE_NAME record, which [MS-OVBA] puts first in every MODULE record, through
// its own MODULE_TERMINATOR, identified by matching MODULE_STREAMNAME against `streamName`. Every other
// record (PROJECTREFERENCES, other modules, project-level fields) is carried through untouched.
function removeModuleDirRecord(dir: Uint8Array, streamName: string, decoder: Decoder): Uint8Array {
  let countAt = -1;
  let blockStart = -1;
  let removeStart = -1;
  let removeEnd = -1;
  let currentStream: string | undefined;
  for (const {id, recordStart, dataStart, size, end} of dirRecords(
    dir,
    'overruns while removing a module',
  )) {
    if (id === REC_MODULES_COUNT) {
      if (size < 2) throw new VbaParseError('PROJECTMODULES MODULES_COUNT record is malformed');
      // The first one wins. A `dir` carrying a second is not a file Excel wrote, and taking the last
      // would let it choose which two bytes of the stream this function overwrites.
      if (countAt < 0) countAt = dataStart;
    } else if (id === REC_MODULE_NAME) {
      blockStart = recordStart;
    } else if (id === REC_MODULE_STREAMNAME) {
      currentStream = decoder.decode(dir.subarray(dataStart, dataStart + size));
    } else if (id === REC_MODULE_TERMINATOR) {
      if (currentStream === streamName) {
        removeStart = blockStart;
        removeEnd = end;
      }
      currentStream = undefined;
      blockStart = -1;
    }
  }
  if (countAt < 0) throw new VbaParseError('dir stream is missing MODULES_COUNT');
  if (removeStart < 0 || removeEnd < 0) {
    throw new VbaParseError(`module stream ${quoted(streamName)} not found in the dir stream`);
  }

  // In a `dir` Excel wrote, MODULES_COUNT precedes every module block, so the splice below moves no
  // byte of it. That is an invariant of the input rather than of this code, and the input came out of
  // a file the caller did not write: were it false, both writes would land inside an unrelated
  // record's payload and quietly corrupt the macro project. Checked, not assumed.
  if (countAt >= removeStart) {
    throw new VbaParseError('dir stream declares MODULES_COUNT after a module block');
  }
  const count = readU16(dir, countAt);
  if (count === 0) {
    throw new VbaParseError('dir stream declares zero modules but carries a module block');
  }
  const out = spliceBytes(dir, removeStart, removeEnd);
  writeU16(out, countAt, count - 1);
  return out;
}

// Remove a module's declaration line (`Module=`/`Class=`) and its `[Workspace]` line from the `PROJECT`
// stream by splicing its bytes, so every other line is left exactly as it was, down to its line ending.
// The stream is MBCS in the project code page, and re-encoding it after an edit could only invert what
// the encoder knows, which under a multi-byte page is ASCII alone: a `Description="日本"` made removing
// an unrelated module throw. Each line is decoded only to be compared.
function removeProjectStreamLines(
  stream: Uint8Array,
  name: string,
  kind: 'procedural' | 'class',
  decoder: Decoder,
): Uint8Array {
  const lines = projectLines(stream, decoder);
  const declLine = `${kind === 'procedural' ? 'Module' : 'Class'}=${name}`;
  const removals: ProjectLine[] = [];
  const decl = lines.find((line) => line.text === declLine);
  if (decl !== undefined) removals.push(decl);

  const wsIndex = lines.findIndex((line) => line.text.trim() === '[Workspace]');
  for (const line of wsIndex < 0 ? [] : lines.slice(wsIndex + 1)) {
    if (line.text.trim() === '' || line.text.startsWith('[')) break;
    if (line.text.startsWith(`${name}=`)) {
      removals.push(line);
      break;
    }
  }

  // The later line first, so the earlier cut does not move the bytes the later one names.
  return removals
    .sort((a, b) => b.start - a.start)
    .reduce((bytes, line) => spliceBytes(bytes, line.start, line.end), stream);
}

/** One line of the `PROJECT` stream: its text, and the bytes the whole line occupies, ending included. */
interface ProjectLine {
  readonly text: string;
  readonly start: number;
  readonly end: number;
}

// Lines end at LF, a CRLF line keeping its CR out of its text. CR, LF and `=` all sit below 0x40, and no
// multi-byte code page a project declares (Shift_JIS, GBK, Big5, EUC-KR) uses a byte that low as a
// trail byte, so a line boundary can never fall inside a character.
function projectLines(stream: Uint8Array, decoder: Decoder): ProjectLine[] {
  const lines: ProjectLine[] = [];
  for (let start = 0; start < stream.length;) {
    const lf = stream.indexOf(0x0a, start);
    const end = lf < 0 ? stream.length : lf + 1;
    const contentEnd = lf > start && stream[lf - 1] === 0x0d ? lf - 1 : lf < 0 ? stream.length : lf;
    lines.push({text: decoder.decode(stream.subarray(start, contentEnd)), start, end});
    start = end;
  }
  return lines;
}

// Remove a module's (MBCS name, UTF-16 name) pair from the binary PROJECTwm stream.
// `existingModuleCount` (from the already fail-closed-parsed project, before
// removal) bounds the walk to the module records, so it never mistakes the terminator for a record.
function removeProjectwmRecord(
  wm: Uint8Array,
  existingModuleCount: number,
  name: string,
  decoder: Decoder,
): Uint8Array {
  let pos = 0;
  let removeStart = -1;
  let removeEnd = -1;
  for (let i = 0; i < existingModuleCount; i++) {
    const recordStart = pos;
    const mbcsEnd = wm.indexOf(0x00, pos);
    if (mbcsEnd < 0)
      throw new VbaParseError('PROJECTwm record is missing its MBCS name terminator');
    const mbcsName = decoder.decode(wm.subarray(pos, mbcsEnd));
    pos = mbcsEnd + 1;
    let utf16End = pos;
    while (utf16End + 1 < wm.length && (wm[utf16End] !== 0 || wm[utf16End + 1] !== 0))
      utf16End += 2;
    if (utf16End + 1 >= wm.length) {
      throw new VbaParseError('PROJECTwm record is missing its Unicode name terminator');
    }
    pos = utf16End + 2;
    if (mbcsName === name) {
      removeStart = recordStart;
      removeEnd = pos;
    }
  }
  if (removeStart < 0 || removeEnd < 0) {
    throw new VbaParseError(`module ${quoted(name)} not found in the PROJECTwm stream`);
  }

  return spliceBytes(wm, removeStart, removeEnd);
}

/** A stream's new bytes, addressed by its storage path from the root. */
interface StreamReplacement {
  readonly path: readonly string[];
  readonly data: Uint8Array;
}

// The tree with each replacement's stream swapped for its new bytes, every path segment matched as
// [MS-CFB] compares names. A same-named stream in another storage is carried through byte-for-byte, as
// is everything else. Each path was just read through the sibling tree `tree()` walks, so one that
// finds nothing here is a fault in this module rather than in the file.
function replaceStreams(
  nodes: readonly CfbNode[],
  replacements: readonly StreamReplacement[],
): CfbNode[] {
  let applied = 0;
  const rebuild = (
    level: readonly CfbNode[],
    pending: readonly StreamReplacement[],
    depth: number,
  ): CfbNode[] =>
    level.map((node) => {
      const here = pending.filter((replacement) => {
        const segment = replacement.path[depth];
        return segment !== undefined && sameEntryName(segment, node.name);
      });
      if (here.length === 0) return node;
      if (isStream(node)) {
        const replacement = here.find((candidate) => candidate.path.length === depth + 1);
        if (replacement === undefined) return node;
        applied++;
        return {name: node.name, data: replacement.data};
      }
      return {name: node.name, children: rebuild(node.children, here, depth + 1)};
    });
  const rebuilt = rebuild(nodes, replacements, 0);
  if (applied !== replacements.length) {
    throw new InternalError('a VBA stream read by its path is missing from the container tree');
  }
  return rebuilt;
}

// The tree without the stream at `path`. The module's stream was read through this same path by the
// parse that validated the removal, so a path that finds nothing is a fault here, as above.
function withoutStream(nodes: readonly CfbNode[], path: readonly string[]): CfbNode[] {
  const [name, ...rest] = path;
  const index = nodes.findIndex(
    (node) =>
      name !== undefined &&
      sameEntryName(node.name, name) &&
      isStream(node) === (rest.length === 0),
  );
  const target = nodes[index];
  if (target === undefined) {
    throw new InternalError(
      'a VBA module stream read by its path is missing from the container tree',
    );
  }
  const kept = isStream(target)
    ? []
    : [{name: target.name, children: withoutStream(target.children, rest)}];
  return [...nodes.slice(0, index), ...kept, ...nodes.slice(index + 1)];
}
