// Reader for the OLE2 / Compound File Binary format ([MS-CFB]).
//
// `vbaProject.bin` is a CFB container (the same "structured storage" behind legacy .doc/.xls). We read
// streams by storage path and, for the edit-in-place path, reconstruct the whole storage/stream hierarchy so it
// can be re-emitted through the writer with one stream swapped, so this is a deliberate subset of
// [MS-CFB]: header → FAT → directory, plus the mini-FAT for sub-cutoff streams, and the red-black
// sibling tree each storage navigates.
//
// It parses an untrusted blob, so every sector index, chain, and stream size is bounds-checked against
// the file and every chain walk is cycle-guarded. A malformed container fails closed with a
// VbaParseError instead of reading out of bounds, looping forever, or over-allocating.

import {quoted} from '../errors.ts';
import {concat, decodeUtf16le, readU16, readU32} from './bytes.ts';
import {
  DIR_ENTRY_SIZE,
  MAX_REGULAR_SECTOR,
  MINI_STREAM_CUTOFF,
  NOSTREAM,
  sameEntryName,
  TYPE_EMPTY,
  TYPE_ROOT,
  TYPE_STORAGE,
  TYPE_STREAM,
} from './cfb-format.ts';
import type {CfbNode} from './cfb-writer.ts';
import {VbaParseError} from './errors.ts';

// Fully readonly, and without the writer's `index`: this is a view of a file we were handed, not a
// tree under construction. `cfb-writer.ts` declares its own for that reason.
interface DirEntry {
  readonly name: string;
  readonly type: number; // 0=empty 1=storage 2=stream 5=root
  readonly startSector: number;
  readonly size: number;
  // Red-black sibling-tree links (entry indices, or a terminal marker for "none"). The reader keeps
  // them so it can walk the tree a host navigates and rebuild the hierarchy, not merely scan names.
  readonly left: number;
  readonly right: number;
  readonly child: number;
}

// The two things a sector chain's kind decides, and the only two: the regular FAT addresses sectors
// in the file itself and finds the next one in `#fat`, the mini-FAT addresses them in the mini stream
// and finds the next one in `#miniFat`. Everything else about following a chain -- the cycle guard,
// the byte budget, the refusal of a chain that ends short -- is the same question either way, which
// is what lets `#walkChain` be one walk. `what` names the chain in both of its refusals, so a failure
// still says which chain it was on.
interface SectorChain {
  readonly what: string;
  readonly sectorSize: number;
  readonly sliceAt: (sector: number, take: number) => Uint8Array;
  readonly next: (sector: number) => number;
}

const CFB_SIGNATURE_LO = 0xe011cfd0;
const CFB_SIGNATURE_HI = 0xe11ab1a1;

// The sibling tree a storage navigates is red-black, so a container written by anything that follows
// [MS-CFB] has depth logarithmic in its entry count: a project with a thousand streams nests around
// ten. A hostile file is under no such obligation, and linking every entry as the left child of the one
// before it makes depth equal to entry count, which the recursive walk turns into an uncatchable
// `RangeError` instead of a `VbaParseError`. Same reasoning and same remedy as `customui/ribbon.ts`.
const MAX_TREE_DEPTH = 256;

export class CompoundFile {
  readonly #buf: Uint8Array;
  readonly #sectorSize: number;
  readonly #miniSectorSize: number;
  readonly #miniCutoff: number;
  readonly #maxSector: number;
  readonly #fat: number[];
  readonly #miniFat: number[];
  readonly #dir: DirEntry[];
  readonly #miniStream: Uint8Array;

  constructor(buf: Uint8Array) {
    this.#buf = buf;

    if (buf.length < 512) throw new VbaParseError('compound file shorter than its 512-byte header');
    if (readU32(this.#buf, 0) !== CFB_SIGNATURE_LO || readU32(this.#buf, 4) !== CFB_SIGNATURE_HI) {
      throw new VbaParseError('not a compound file (bad OLE2 signature)');
    }

    const sectorShift = readU16(this.#buf, 30);
    const miniSectorShift = readU16(this.#buf, 32);
    // [MS-CFB] fixes these: 512-byte sectors (shift 9) for v3, 4096 (shift 12) for v4; mini shift 6.
    if (sectorShift !== 9 && sectorShift !== 12) {
      throw new VbaParseError(`unsupported sector shift ${sectorShift}`);
    }
    if (miniSectorShift !== 6)
      throw new VbaParseError(`unsupported mini-sector shift ${miniSectorShift}`);
    this.#sectorSize = 1 << sectorShift;
    this.#miniSectorSize = 1 << miniSectorShift;
    this.#maxSector = Math.floor(buf.length / this.#sectorSize);

    const numFatSectors = readU32(this.#buf, 44);
    const firstDirSector = readU32(this.#buf, 48);
    this.#miniCutoff = readU32(this.#buf, 56);
    // Checked on the same terms as the two shifts above, and for a sharper reason: a wrong shift is
    // caught downstream by arithmetic that stops making sense, while a wrong cutoff routes every
    // stream through the *other* allocator, which is bounds-checked and will happily return bytes.
    if (this.#miniCutoff !== MINI_STREAM_CUTOFF) {
      throw new VbaParseError(`unsupported mini-stream cutoff ${this.#miniCutoff}`);
    }
    const firstMiniFatSector = readU32(this.#buf, 60);
    const firstDifatSector = readU32(this.#buf, 68);
    const numDifatSectors = readU32(this.#buf, 72);

    const fatSectorIds = this.#readDifat(numFatSectors, firstDifatSector, numDifatSectors);
    this.#fat = this.#readFat(fatSectorIds);
    this.#miniFat = this.#readChainValues(firstMiniFatSector);
    this.#dir = this.#readDirectory(firstDirSector);

    const root = this.#dir.find((e) => e.type === TYPE_ROOT);
    if (!root) throw new VbaParseError('compound file has no root storage entry');
    // The mini-stream lives in the regular FAT, addressed from the root entry's start sector.
    this.#miniStream = this.#readViaFat(root.startSector, root.size);
  }

  /**
   * Read a stream's raw bytes by its storage path from the root (`['VBA', 'dir']`), or `undefined` when
   * no stream sits there. Each segment is matched among one storage's children, as [MS-CFB] compares
   * names, by walking the sibling tree that storage links.
   *
   * A path because a name is not an address. [MS-OVBA] puts `PROJECT` at the root, `dir` and every
   * module stream in `VBA`, and each UserForm's `f` and `o` in a storage of its own, so a bare name
   * repeats. Looked up by name, the first entry in directory order won: a module named `f` read a
   * UserForm's form data, and an edit wrote a module named `PROJECT` over the root `PROJECT`.
   *
   * @returns bytes that may be a **view onto the caller's buffer** rather than a copy. A stream small
   *   enough to fit in one sector is assembled by `concat` from a single chunk, and `concat` hands a
   *   lone chunk straight back; anything larger is copied into a fresh buffer. So whether the result
   *   aliases the source depends on the stream's size, which is a rule no caller should have to know
   *   and none here relies on: nothing mutates what this returns. Stated because it is the sort of
   *   thing that is true until someone writes the first mutation, and then true only for large
   *   streams. Copy before mutating.
   */
  readStream(path: readonly string[]): Uint8Array | undefined {
    const rootIdx = this.#rootIndex();
    const seen = new Set([rootIdx]);
    let firstChild = (this.#dir[rootIdx] as DirEntry).child;
    for (const [depth, segment] of path.entries()) {
      const stream = depth === path.length - 1;
      const entry = this.#childIndices(firstChild, seen, depth)
        .map((idx) => this.#dir[idx] as DirEntry)
        .find(
          (e) =>
            (stream ? e.type === TYPE_STREAM : e.type === TYPE_STORAGE) &&
            sameEntryName(e.name, segment),
        );
      if (entry === undefined) return undefined;
      if (stream) return this.#readEntryData(entry);
      firstChild = entry.child;
    }
    return undefined;
  }

  /**
   * Reconstruct the container's top-level children as the writer's node shape, recursing into every
   * storage, so a caller can swap one stream and re-emit the whole hierarchy with {@link writeCompoundFile}.
   * Walks the red-black sibling tree each storage navigates (not the linear directory scan), so any part
   * a host reaches is carried through. Cycle- and bounds-guarded like every other chain walk here.
   *
   * @returns nodes whose stream bytes carry {@link readStream}'s aliasing caveat.
   */
  tree(): CfbNode[] {
    const rootIdx = this.#rootIndex();
    return this.#buildSiblings((this.#dir[rootIdx] as DirEntry).child, new Set([rootIdx]));
  }

  #rootIndex(): number {
    const rootIdx = this.#dir.findIndex((e) => e.type === TYPE_ROOT);
    if (rootIdx < 0) throw new VbaParseError('compound file has no root storage entry');
    return rootIdx;
  }

  #buildSiblings(firstChild: number, seen: Set<number>, depth = 0): CfbNode[] {
    if (depth > MAX_TREE_DEPTH) throw new VbaParseError('directory tree nests too deep');
    return this.#childIndices(firstChild, seen, depth).map((idx) => {
      const e = this.#dir[idx] as DirEntry;
      return e.type === TYPE_STORAGE
        ? {name: e.name, children: this.#buildSiblings(e.child, seen, depth + 1)}
        : {name: e.name, data: this.#readEntryData(e)};
    });
  }

  // The entries one storage links as its children, in the sibling tree's in-order sequence: the one
  // walk both a path lookup and the whole-tree rebuild go through. `seen` spans the whole descent, so a
  // link back to any entry already visited is a cycle, and `depth` carries the storage nesting in, so
  // a tree deep in siblings and storages together is refused on the one bound.
  #childIndices(firstChild: number, seen: Set<number>, depth: number): number[] {
    const indices: number[] = [];
    const walk = (idx: number, siblingDepth: number): void => {
      if (idx >= MAX_REGULAR_SECTOR) return; // NOSTREAM / terminal marker → no such sibling
      if (siblingDepth > MAX_TREE_DEPTH) throw new VbaParseError('directory tree nests too deep');
      if (idx >= this.#dir.length) throw new VbaParseError('directory sibling index out of range');
      if (seen.has(idx)) throw new VbaParseError('cycle in directory sibling tree');
      seen.add(idx);
      const e = this.#dir[idx] as DirEntry;
      if (e.type === TYPE_EMPTY) throw new VbaParseError('directory tree links an empty entry');
      walk(e.left, siblingDepth + 1);
      indices.push(idx);
      walk(e.right, siblingDepth + 1);
    };
    walk(firstChild, depth);
    return indices;
  }

  #readEntryData(entry: DirEntry): Uint8Array {
    if (entry.size >= this.#miniCutoff) return this.#readViaFat(entry.startSector, entry.size);
    return this.#readViaMiniFat(entry.startSector, entry.size);
  }

  /**
   * The FAT-sector ids, bounded by the file rather than by the header's claim about it.
   *
   * `numFatSectors` and `numDifat` are unvalidated `u32`s at header offsets 44 and 72, and each id
   * costs a sector's worth of FAT entries downstream. Believing them let 1 MB of crafted container
   * amplify into 272 MB of heap, linear in input size, so a 10 MB project was a fatal OOM rather than
   * a typed failure. The file cannot hold more FAT sectors than it holds sectors, so `#maxSector` is
   * the honest ceiling for both loops, and ids repeating freely is why the count must bound the
   * DIFAT-sector loop too, not only the header one.
   */
  #readDifat(numFatSectors: number, firstDifat: number, numDifat: number): number[] {
    const wanted = Math.min(numFatSectors, this.#maxSector);
    const ids: number[] = [];
    // The first 109 FAT-sector pointers live in the header; the rest chain through DIFAT sectors.
    for (let i = 0; i < 109 && ids.length < wanted; i++) {
      const v = readU32(this.#buf, 76 + i * 4);
      if (v >= MAX_REGULAR_SECTOR) break;
      ids.push(v);
    }
    const seen = new Set<number>();
    let sector = firstDifat;
    const perSector = this.#sectorSize / 4 - 1;
    for (let s = 0; s < numDifat && sector < MAX_REGULAR_SECTOR && ids.length < wanted; s++) {
      if (seen.has(sector)) throw new VbaParseError('cycle in DIFAT sector chain');
      seen.add(sector);
      const base = this.#dataSectorOffset(sector);
      for (let i = 0; i < perSector && ids.length < wanted; i++) {
        const v = readU32(this.#buf, base + i * 4);
        if (v < MAX_REGULAR_SECTOR) ids.push(v);
      }
      sector = readU32(this.#buf, base + perSector * 4);
    }
    return ids;
  }

  // A FAT entry past the sector count addresses a sector that is not in the file, so it can never be
  // followed: reading more than `#maxSector` of them is work with no possible use.
  #readFat(fatSectorIds: number[]): number[] {
    const fat: number[] = [];
    const perSector = this.#sectorSize / 4;
    for (const sid of fatSectorIds) {
      const base = this.#dataSectorOffset(sid);
      for (let i = 0; i < perSector && fat.length < this.#maxSector; i++) {
        fat.push(readU32(this.#buf, base + i * 4));
      }
      if (fat.length >= this.#maxSector) break;
    }
    return fat;
  }

  #readChainValues(firstSector: number): number[] {
    // Every uint32 in a sector chain (used for the mini-FAT), cycle-guarded and capped on the same
    // terms as {@link #readFat}: a mini-FAT entry beyond the mini stream's sector count is unusable.
    const values: number[] = [];
    const perSector = this.#sectorSize / 4;
    const cap = this.#maxSector * (this.#sectorSize / this.#miniSectorSize);
    const seen = new Set<number>();
    let sector = firstSector;
    while (sector < MAX_REGULAR_SECTOR && values.length < cap) {
      if (seen.has(sector)) throw new VbaParseError('cycle in mini-FAT sector chain');
      seen.add(sector);
      const base = this.#dataSectorOffset(sector);
      for (let i = 0; i < perSector && values.length < cap; i++) {
        values.push(readU32(this.#buf, base + i * 4));
      }
      sector = this.#nextInFat(sector);
    }
    return values;
  }

  #readDirectory(firstDirSector: number): DirEntry[] {
    const raw = this.#readChainFull(firstDirSector);
    const entries: DirEntry[] = [];
    // Empty slots are kept as placeholders (not skipped) so array indices stay equal to the on-disk
    // directory-entry ids the sibling-tree links reference: the tree walk in #buildSiblings needs them.
    for (let off = 0; off + DIR_ENTRY_SIZE <= raw.length; off += DIR_ENTRY_SIZE) {
      const type = raw[off + 66] as number;
      if (
        type !== TYPE_EMPTY &&
        type !== TYPE_STORAGE &&
        type !== TYPE_STREAM &&
        type !== TYPE_ROOT
      ) {
        throw new VbaParseError(`directory entry has invalid object type ${type}`);
      }
      const left = readU32(raw, off + 68);
      const right = readU32(raw, off + 72);
      const child = readU32(raw, off + 76);
      if (type === TYPE_EMPTY) {
        entries.push({
          name: '',
          type,
          startSector: 0,
          size: 0,
          left: NOSTREAM,
          right: NOSTREAM,
          child: NOSTREAM,
        });
        continue;
      }
      const nameLen = readU16(raw, off + 64);
      if (nameLen > 64 || nameLen % 2 !== 0) {
        throw new VbaParseError(`directory entry has invalid name length ${nameLen}`);
      }
      const name = decodeUtf16le(raw.subarray(off, off + Math.max(0, nameLen - 2)));
      const startSector = readU32(raw, off + 116);
      const size = readU32(raw, off + 120); // low 32 bits, ample for a VBA project
      // The high half is not ample for anything this library reads, but ignoring it is not the same
      // as refusing it: a stream declaring 4 GiB + 10 bytes would otherwise read back as its first
      // ten, silently, which is a module's source truncated rather than a file rejected.
      if (readU32(raw, off + 124) !== 0) {
        throw new VbaParseError(
          `directory entry ${quoted(name)} declares a stream larger than 4 GiB`,
        );
      }
      entries.push({name, type, startSector, size, left, right, child});
    }
    return entries;
  }

  // Follow a FAT chain to its end, collecting whole sectors: the directory's read, since a directory
  // declares no byte length of its own and so has no size to fall short of.
  #readChainFull(startSector: number): Uint8Array {
    return this.#walkChain(startSector, undefined, this.#fatChain('FAT sector chain'));
  }

  #readViaFat(startSector: number, size: number): Uint8Array {
    return this.#walkChain(startSector, size, this.#fatChain('stream FAT chain'));
  }

  #readViaMiniFat(startSector: number, size: number): Uint8Array {
    return this.#walkChain(startSector, size, this.#miniFatChain());
  }

  /**
   * Follow a sector chain, collecting `size` bytes, or whole sectors to the end of the chain when
   * there is no declared size.
   *
   * One walk where there were three. The FAT and the mini-FAT differ in where a sector's bytes live
   * and in how the next sector is found, and in nothing else -- and the rest of what these loops do
   * is refuse a cycle in a chain that came out of a file. That guard is the only thing between a
   * crafted FAT and a loop that never ends, and a guard held in three copies is three places to fix
   * the next thing found wrong with it.
   *
   * A chain that ends before the declared size is **refused**, not returned short. That is the same
   * call the 4-GiB check on a directory entry makes a few lines up, for the same reason: a stream
   * that reads back as its first sector is a module's source truncated, and it is truncated
   * silently, with every parser downstream seeing a well-formed prefix rather than a damaged file.
   */
  #walkChain(startSector: number, size: number | undefined, chain: SectorChain): Uint8Array {
    const chunks: Uint8Array[] = [];
    const seen = new Set<number>();
    let sector = startSector;
    // No declared size means "to the end of the chain", which is a budget nothing can exhaust.
    let remaining = size ?? Infinity;
    while (sector < MAX_REGULAR_SECTOR && remaining > 0) {
      if (seen.has(sector)) throw new VbaParseError(`cycle in ${chain.what}`);
      seen.add(sector);
      const take = Math.min(chain.sectorSize, remaining);
      chunks.push(chain.sliceAt(sector, take));
      remaining -= take;
      sector = chain.next(sector);
    }
    if (size !== undefined && remaining > 0) {
      throw new VbaParseError(
        `${chain.what} ends after ${size - remaining} bytes, but ${size} were declared`,
      );
    }
    return concat(chunks);
  }

  // A sector of the regular FAT: its bytes sit in the file itself, past the header.
  #fatChain(what: string): SectorChain {
    return {
      what,
      sectorSize: this.#sectorSize,
      sliceAt: (sector, take) => {
        const base = this.#dataSectorOffset(sector);
        return this.#buf.subarray(base, base + take);
      },
      next: (sector) => this.#nextInFat(sector),
    };
  }

  // A sector of the mini-FAT: its bytes sit in the mini stream, which is itself a stream read
  // through the regular FAT, so both bounds are checked against that buffer rather than the file.
  #miniFatChain(): SectorChain {
    return {
      what: 'stream mini-FAT chain',
      sectorSize: this.#miniSectorSize,
      sliceAt: (sector, take) => {
        const base = sector * this.#miniSectorSize;
        if (base + take > this.#miniStream.length) {
          throw new VbaParseError('mini-stream sector runs past end of the mini stream');
        }
        return this.#miniStream.subarray(base, base + take);
      },
      next: (sector) => {
        if (sector >= this.#miniFat.length) throw new VbaParseError('mini-FAT index out of range');
        return this.#miniFat[sector] as number;
      },
    };
  }

  #nextInFat(sector: number): number {
    if (sector >= this.#fat.length) throw new VbaParseError('FAT index out of range');
    return this.#fat[sector] as number;
  }

  #dataSectorOffset(sector: number): number {
    if (sector >= this.#maxSector) {
      throw new VbaParseError(`sector ${sector} is out of range (file has ${this.#maxSector})`);
    }
    const base = (sector + 1) * this.#sectorSize;
    if (base + this.#sectorSize > this.#buf.length) {
      throw new VbaParseError(`sector ${sector} runs past end of file`);
    }
    return base;
  }
}
