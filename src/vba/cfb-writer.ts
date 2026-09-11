// Writer for the OLE2 / Compound File Binary format ([MS-CFB]): the encode counterpart to cfb.ts.
//
// Produces a v3 (512-byte sector) container from a hierarchy of storages and streams, the substrate a
// synthesized vbaProject.bin is built on (its modules live inside a `VBA` storage, with `PROJECT` and
// `PROJECTwm` at the root). Streams below the 4096-byte mini cutoff are packed into the mini stream and
// chained through the mini-FAT; larger streams take whole regular sectors. Each storage's children are
// emitted as a name-ordered balanced binary tree ([MS-CFB] 2.6.4), so a host that *navigates* the tree
// (Excel) reaches every entry, not only a linear scanner like this library's own reader.
//
// Unlike the reader, this is not a hostile-input path: we are the producer. It still validates its
// contract (name length, sibling-name uniqueness, size bound) and fails closed with VbaAuthorError,
// because a silently malformed container would surface far downstream as an unopenable workbook.
//
// `writeCompoundFile` is the longest function in the tree and stays one: it is a single algorithm over
// shared mutable state, sectioned by the banner comments below (directory build, sector layout,
// serialize), and the state it shares is precisely what a split would have to pass along. If it is
// ever split, the seam is an explicit layout record (`{dirStart, miniFatStart, miniStreamStart,
// fatStart, difatStart, totalSectors}`) computed by one function and consumed by the serializer, with
// the fixpoint loop that sizes the FAT and DIFAT named and tested on its own. That loop is now tested
// through the public entry point instead (`cfb-writer.test.ts` sweeps the sector-boundary crossings),
// which is the coverage the split would have bought, without the seam.

import {quoted} from '../errors.ts';
import {
  CFB_SIGNATURE,
  DIFSECT,
  DIR_ENTRY_FIELD,
  DIR_ENTRY_SIZE,
  ENDOFCHAIN,
  FATSECT,
  FREESECT,
  HEADER_DIFAT_SLOTS,
  HEADER_FIELD,
  MAX_NAME_CHARS,
  MINI_SECTOR_SHIFT,
  MINI_STREAM_CUTOFF,
  NOSTREAM,
  TYPE_ROOT,
  TYPE_STORAGE,
  TYPE_STREAM,
} from './cfb-format.ts';
import {VbaAuthorError} from './errors.ts';

export interface CfbStream {
  /** The exact directory-entry name. At most 31 UTF-16 code units ([MS-CFB] name limit). */
  readonly name: string;
  readonly data: Uint8Array;
}

export interface CfbStorage {
  readonly name: string;
  readonly children: readonly CfbNode[];
}

export type CfbNode = CfbStream | CfbStorage;

/**
 * Which arm of {@link CfbNode} a node is.
 *
 * Exported beside the union rather than kept private, because the discriminant is `'data' in node`
 * and that is the sort of test callers re-spell inline: `project-editor.ts` had it three times, once
 * negated. The name says what the test means, and a node it answers `false` for is a storage.
 */
export function isStream(node: CfbNode): node is CfbStream {
  return 'data' in node;
}

// The v3 layout this writer chooses to emit. Not shared with the reader, which takes the sector size
// off the header it was handed because a file may legally say otherwise.
const SECTOR_SHIFT = 9;
const SECTOR = 1 << SECTOR_SHIFT;
const MINI_SECTOR = 1 << MINI_SECTOR_SHIFT;
const ENTRIES_PER_DIR_SECTOR = SECTOR / DIR_ENTRY_SIZE; // 4
const FAT_ENTRIES_PER_SECTOR = SECTOR / 4; // 128

// The red-black colour byte ([MS-CFB] 2.6.1). Every entry this writer emits is black, which is
// legal for any tree; the reader ignores the byte entirely, so it is not a shared constant.
const COLOR_BLACK = 1;
// Mutable, and carrying an `index` the reader's has no use for: this tree is under construction and
// its sibling links are these indices. `cfb.ts` declares its own, fully readonly, for that reason.
interface DirEntry {
  /**
   * Position in the directory stream. Sibling and child links are stored as these indices, so an
   * entry carries its own; the alternative is looking the object back up in `entries`, which under
   * `noUncheckedIndexedAccess` yields `DirEntry | undefined` at every use site.
   */
  readonly index: number;
  name: string;
  type: number;
  startSector: number;
  size: number;
  left: number;
  right: number;
  child: number;
}

/** A stream too large for the mini stream: it owns whole sectors, placed during layout. */
interface BigStream {
  entry: DirEntry;
  data: Uint8Array;
  sectors: number;
}

/**
 * Encode a hierarchy of storages and streams into a v3 compound file. The Root Entry is synthesized
 * automatically; `root` is its top-level children. Every stream becomes a directory entry reachable both
 * by linear scan and by tree navigation.
 *
 * @throws {VbaAuthorError} if any name is empty or exceeds 31 characters, sibling names collide, or the
 *   project is so large it would need more than 109 FAT sectors (~7 MB, far beyond any real project).
 */
export function writeCompoundFile(root: readonly CfbNode[]): Uint8Array {
  const entries: DirEntry[] = [];
  const addEntry = (name: string, type: number, startSector: number): DirEntry => {
    const entry: DirEntry = {
      index: entries.length,
      name,
      type,
      startSector,
      size: 0,
      left: NOSTREAM,
      right: NOSTREAM,
      child: NOSTREAM,
    };
    entries.push(entry);
    return entry;
  };

  const rootEntry = addEntry('Root Entry', TYPE_ROOT, ENDOFCHAIN);

  // Sub-cutoff stream bytes accumulate into the mini stream (chained in the mini-FAT); larger streams
  // are laid out later directly in the regular FAT. Depth-first walk fixes a deterministic layout.
  //
  // Kept as its own chunks and copied into the buffer one by one, rather than accumulated in a
  // `number[]`. Every sub-cutoff stream a project has goes through here (`dir`, `PROJECT`,
  // `PROJECTwm`, and each module's compressed source), so a JS array would box hundreds of thousands
  // of bytes the caller already handed over as bytes. Each chunk is padded to a whole number of mini
  // sectors, so the running length is both the next chunk's start offset and the stream's final size.
  const miniChunks: Uint8Array[] = [];
  let miniLength = 0;
  const miniFat: number[] = [];
  const bigStreams: BigStream[] = [];

  // `siblings` is the parent's child list itself rather than its index, so a child is appended to an
  // array we hold, with no lookup that could come back empty.
  const addNode = (node: CfbNode, siblings: DirEntry[]): void => {
    if (isStream(node)) {
      const entry = addEntry(node.name, TYPE_STREAM, ENDOFCHAIN);
      entry.size = node.data.length;
      siblings.push(entry);
      if (node.data.length === 0) {
        // an empty stream owns no sectors
      } else if (node.data.length >= MINI_STREAM_CUTOFF) {
        bigStreams.push({entry, data: node.data, sectors: Math.ceil(node.data.length / SECTOR)});
      } else {
        const startMini = miniLength / MINI_SECTOR;
        const numMini = Math.ceil(node.data.length / MINI_SECTOR);
        for (let k = 0; k < numMini; k++)
          miniFat.push(k < numMini - 1 ? startMini + k + 1 : ENDOFCHAIN);
        const padded = new Uint8Array(numMini * MINI_SECTOR);
        padded.set(node.data);
        miniChunks.push(padded);
        miniLength += padded.length;
        entry.startSector = startMini;
      }
    } else {
      const entry = addEntry(node.name, TYPE_STORAGE, 0);
      siblings.push(entry);
      const kids: DirEntry[] = [];
      for (const c of node.children) addNode(c, kids);
      // Each storage (Root included) links its children as a balanced search tree the host navigates.
      linkChildren(entry, kids);
    }
  };

  validateSiblingNames('Root Entry', root);
  const rootKids: DirEntry[] = [];
  for (const n of root) addNode(n, rootKids);
  linkChildren(rootEntry, rootKids);

  // ── Sector layout ─────────────────────────────────────────────────────────────────────────────────
  // Physical order: directory, mini-FAT, mini stream, each big stream, FAT, DIFAT. Region starts are
  // assigned first so chains can reference them; the FAT is filled last, once every sector is placed.
  const dirSectors = Math.ceil(entries.length / ENTRIES_PER_DIR_SECTOR);
  const miniFatSectors =
    miniFat.length > 0 ? Math.ceil((miniFat.length * 4) / FAT_ENTRIES_PER_SECTOR) : 0;
  const miniStreamSectors = Math.ceil(miniLength / SECTOR);
  const baseSectors =
    dirSectors +
    miniFatSectors +
    miniStreamSectors +
    bigStreams.reduce((total, big) => total + big.sectors, 0);

  let fatSectors = 0;
  let difatSectors = 0;
  for (;;) {
    const total = baseSectors + fatSectors + difatSectors;
    const needFat = Math.ceil(total / FAT_ENTRIES_PER_SECTOR);
    // Each DIFAT sector holds 127 FAT pointers + a next-DIFAT pointer; the first 109 live in the header.
    const needDifat =
      needFat > HEADER_DIFAT_SLOTS
        ? Math.ceil((needFat - HEADER_DIFAT_SLOTS) / (FAT_ENTRIES_PER_SECTOR - 1))
        : 0;
    if (needFat === fatSectors && needDifat === difatSectors) break;
    fatSectors = needFat;
    difatSectors = needDifat;
  }
  if (fatSectors > HEADER_DIFAT_SLOTS) {
    throw new VbaAuthorError(
      `project needs ${fatSectors} FAT sectors, exceeding the ${HEADER_DIFAT_SLOTS}-sector single-header bound`,
    );
  }

  let cursor = 0;
  const dirStart = cursor;
  cursor += dirSectors;
  const miniFatStart = miniFatSectors > 0 ? cursor : ENDOFCHAIN;
  cursor += miniFatSectors;
  const miniStreamStart = miniStreamSectors > 0 ? cursor : ENDOFCHAIN;
  cursor += miniStreamSectors;
  for (const big of bigStreams) {
    big.entry.startSector = cursor;
    cursor += big.sectors;
  }
  const fatStart = cursor;
  cursor += fatSectors;
  const difatStart = difatSectors > 0 ? cursor : ENDOFCHAIN;
  cursor += difatSectors;
  const totalSectors = cursor;

  rootEntry.startSector = miniStreamStart;
  rootEntry.size = miniLength;

  // ── FAT ─────────────────────────────────────────────────────────────────────────────────────────
  const fat = new Array<number>(fatSectors * FAT_ENTRIES_PER_SECTOR).fill(FREESECT);
  const chainRegion = (start: number, count: number): void => {
    for (let k = 0; k < count; k++) fat[start + k] = k < count - 1 ? start + k + 1 : ENDOFCHAIN;
  };
  chainRegion(dirStart, dirSectors);
  if (miniFatSectors > 0) chainRegion(miniFatStart, miniFatSectors);
  if (miniStreamSectors > 0) chainRegion(miniStreamStart, miniStreamSectors);
  for (const big of bigStreams) chainRegion(big.entry.startSector, big.sectors);
  for (let k = 0; k < fatSectors; k++) fat[fatStart + k] = FATSECT;
  for (let k = 0; k < difatSectors; k++) fat[difatStart + k] = DIFSECT;

  // ── Serialize ──────────────────────────────────────────────────────────────────────────────────
  const buf = new Uint8Array((totalSectors + 1) * SECTOR); // +1 for the header sector
  const dv = new DataView(buf.buffer);
  const at = (sector: number): number => (sector + 1) * SECTOR;

  writeHeader(dv, {
    fatSectors,
    dirStart,
    miniFatStart,
    miniFatSectors,
    difatStart,
    difatSectors,
    fatStart,
  });

  for (const e of entries) writeDirEntry(dv, at(dirStart) + e.index * DIR_ENTRY_SIZE, e);

  for (let i = 0; i < miniFatSectors * FAT_ENTRIES_PER_SECTOR; i++) {
    dv.setUint32(at(miniFatStart) + i * 4, miniFat[i] ?? FREESECT, true);
  }
  // Guarded rather than looped-and-skipped: with no mini stream, miniStreamStart is ENDOFCHAIN and
  // `at()` of it is far outside the buffer, which a zero-length `set` would still reject.
  if (miniLength > 0) {
    let offset = at(miniStreamStart);
    for (const chunk of miniChunks) {
      buf.set(chunk, offset);
      offset += chunk.length;
    }
  }
  for (const big of bigStreams) buf.set(big.data, at(big.entry.startSector));
  for (const [i, sector] of fat.entries()) dv.setUint32(at(fatStart) + i * 4, sector, true);

  // DIFAT sectors carry FAT-sector pointers 110.. (unreachable under the enforced bound, but the layout
  // is honoured: each DIFAT sector's tail points to the next, the last to ENDOFCHAIN).
  for (let d = 0; d < difatSectors; d++) {
    const base = at(difatStart + d);
    for (let i = 0; i < FAT_ENTRIES_PER_SECTOR - 1; i++) {
      const fatIdx = HEADER_DIFAT_SLOTS + d * (FAT_ENTRIES_PER_SECTOR - 1) + i;
      dv.setUint32(base + i * 4, fatIdx < fatSectors ? fatStart + fatIdx : FREESECT, true);
    }
    dv.setUint32(
      base + (FAT_ENTRIES_PER_SECTOR - 1) * 4,
      d < difatSectors - 1 ? difatStart + d + 1 : ENDOFCHAIN,
      true,
    );
  }

  return buf;
}

function validateSiblingNames(storageName: string, siblings: readonly CfbNode[]): void {
  const seen = new Set<string>();
  for (const node of siblings) {
    if (node.name.length === 0) throw new VbaAuthorError('entry name must not be empty');
    if (node.name.length > MAX_NAME_CHARS) {
      throw new VbaAuthorError(
        `entry name ${quoted(node.name)} exceeds the ${MAX_NAME_CHARS}-character CFB limit`,
      );
    }
    if (seen.has(node.name)) {
      throw new VbaAuthorError(
        `duplicate entry name ${quoted(node.name)} under storage ${quoted(storageName)}`,
      );
    }
    seen.add(node.name);
    if (!isStream(node)) validateSiblingNames(node.name, node.children);
  }
}

// Order a storage's children by the [MS-CFB] 2.6.4 comparison (shorter names sort first; ties broken by
// uppercased UTF-16 code units) and link them as a balanced binary tree. A host locates a child by
// walking this tree from the storage's `child` pointer, so the ordering and links must form a valid
// search tree.
//
// The recursion halves a slice rather than a lo/hi index pair, which makes "the slice is empty" and
// "there is no midpoint entry" the same observable fact: the `undefined` check *is* the base case.
function linkChildren(storage: DirEntry, kids: readonly DirEntry[]): void {
  const build = (nodes: readonly DirEntry[]): number => {
    const mid = nodes.length >> 1;
    const node = nodes[mid];
    if (node === undefined) return NOSTREAM;
    node.left = build(nodes.slice(0, mid));
    node.right = build(nodes.slice(mid + 1));
    return node.index;
  };
  storage.child = build([...kids].sort((a, b) => compareNames(a.name, b.name)));
}

function compareNames(a: string, b: string): number {
  if (a.length !== b.length) return a.length - b.length;
  const ua = a.toUpperCase();
  const ub = b.toUpperCase();
  return ua < ub ? -1 : ua > ub ? 1 : 0;
}

function writeHeader(
  dv: DataView,
  p: {
    fatSectors: number;
    dirStart: number;
    miniFatStart: number;
    miniFatSectors: number;
    difatStart: number;
    difatSectors: number;
    fatStart: number;
  },
): void {
  dv.setUint32(HEADER_FIELD.signatureLow, CFB_SIGNATURE.low, true);
  dv.setUint32(HEADER_FIELD.signatureHigh, CFB_SIGNATURE.high, true);
  dv.setUint16(HEADER_FIELD.minorVersion, 0x003e, true);
  dv.setUint16(HEADER_FIELD.majorVersion, 0x0003, true); // v3: 512-byte sectors
  dv.setUint16(HEADER_FIELD.byteOrder, 0xfffe, true); // little-endian
  dv.setUint16(HEADER_FIELD.sectorShift, SECTOR_SHIFT, true);
  dv.setUint16(HEADER_FIELD.miniSectorShift, MINI_SECTOR_SHIFT, true);
  dv.setUint32(HEADER_FIELD.fatSectors, p.fatSectors, true);
  dv.setUint32(HEADER_FIELD.firstDirectorySector, p.dirStart, true);
  dv.setUint32(HEADER_FIELD.miniStreamCutoff, MINI_STREAM_CUTOFF, true);
  dv.setUint32(HEADER_FIELD.firstMiniFatSector, p.miniFatStart, true);
  dv.setUint32(HEADER_FIELD.miniFatSectors, p.miniFatSectors, true);
  dv.setUint32(HEADER_FIELD.firstDifatSector, p.difatSectors > 0 ? p.difatStart : ENDOFCHAIN, true);
  dv.setUint32(HEADER_FIELD.difatSectors, p.difatSectors, true);
  // The header's FAT-sector pointers, contiguous from fatStart under the enforced bound.
  for (let i = 0; i < HEADER_DIFAT_SLOTS; i++) {
    dv.setUint32(HEADER_FIELD.difat + i * 4, i < p.fatSectors ? p.fatStart + i : FREESECT, true);
  }
}

function writeDirEntry(dv: DataView, off: number, e: DirEntry): void {
  for (let i = 0; i < e.name.length; i++) dv.setUint16(off + i * 2, e.name.charCodeAt(i), true);
  // The name's byte length counts its NUL terminator.
  dv.setUint16(off + DIR_ENTRY_FIELD.nameLength, (e.name.length + 1) * 2, true);
  dv.setUint8(off + DIR_ENTRY_FIELD.objectType, e.type);
  dv.setUint8(off + DIR_ENTRY_FIELD.color, COLOR_BLACK);
  dv.setUint32(off + DIR_ENTRY_FIELD.leftSibling, e.left, true);
  dv.setUint32(off + DIR_ENTRY_FIELD.rightSibling, e.right, true);
  dv.setUint32(off + DIR_ENTRY_FIELD.child, e.child, true);
  dv.setUint32(off + DIR_ENTRY_FIELD.startSector, e.startSector, true);
  dv.setUint32(off + DIR_ENTRY_FIELD.sizeLow, e.size, true);
  // The size's high half, the CLSID, and the state and time fields stay zero: valid for a v3 entry.
}
