// The [MS-CFB] numbers the reader and the writer both answer to.
//
// `vbaProject.bin` is the one place in this tree where a hand-written binary reader and a
// hand-written binary writer are expected to round-trip each other's output byte for byte, and the
// whole edit-in-place path rests on that. Two independent transcriptions of the same numbers is the
// cheapest way for that to break, and it would break quietly: a wrong `TYPE_*` or field offset in one
// file yields a container this library still reads and Excel rejects.
//
// What lives here is what the format fixes for every file: the signature, the chain markers and object
// types, where each header and directory-entry field sits, and the values the reader refuses to see
// otherwise (the mini-sector shift, the mini-stream cutoff, the header's 109 DIFAT slots). What does
// not is the v3 layout the writer chooses to emit (512-byte sectors, version 3), because the reader
// takes those off the header it was handed: a file may legally say otherwise, and a shared constant
// would invite a reader to trust the layout over the header field.

/** The header's first eight bytes, as the two little-endian `u32`s both directions handle them as. */
export const CFB_SIGNATURE = {low: 0xe011cfd0, high: 0xe11ab1a1} as const;

/** Byte offsets of the header fields either direction reads or writes ([MS-CFB] 2.2). */
export const HEADER_FIELD = {
  signatureLow: 0,
  signatureHigh: 4,
  minorVersion: 24,
  majorVersion: 26,
  byteOrder: 28,
  sectorShift: 30,
  miniSectorShift: 32,
  fatSectors: 44,
  firstDirectorySector: 48,
  miniStreamCutoff: 56,
  firstMiniFatSector: 60,
  miniFatSectors: 64,
  firstDifatSector: 68,
  difatSectors: 72,
  difat: 76,
} as const;

/** Byte offsets within a directory entry, from the entry's start ([MS-CFB] 2.6.1). The name's UTF-16
 * code units begin at offset 0. */
export const DIR_ENTRY_FIELD = {
  nameLength: 64,
  objectType: 66,
  color: 67,
  leftSibling: 68,
  rightSibling: 72,
  child: 76,
  startSector: 116,
  sizeLow: 120,
  sizeHigh: 124,
} as const;

/** FAT-sector pointers the header holds before a DIFAT sector is needed ([MS-CFB] 2.2). */
export const HEADER_DIFAT_SLOTS = 109;

/** 64-byte mini sectors: [MS-CFB] 2.2 fixes the shift at 6 for both versions. */
export const MINI_SECTOR_SHIFT = 6;

/** Sector chain markers ([MS-CFB] 2.2). */
export const FREESECT = 0xffffffff;
export const ENDOFCHAIN = 0xfffffffe;
export const FATSECT = 0xfffffffd;
export const DIFSECT = 0xfffffffc;

/**
 * Sector values 0xFFFFFFFA..0xFFFFFFFF are the reserved markers above, not data-sector indices, so
 * any value at or above this ceiling is chain-terminal. A reader walking a hostile chain tests the
 * ceiling rather than the four markers by name: an unassigned reserved value is terminal too.
 */
export const MAX_REGULAR_SECTOR = 0xfffffffa;

/** Directory-tree links reuse the sector convention: NOSTREAM, and any value at or above
 * {@link MAX_REGULAR_SECTOR}, mean "no such sibling or child". */
export const NOSTREAM = 0xffffffff;

/** Object types ([MS-CFB] 2.6.1). */
export const TYPE_EMPTY = 0;
export const TYPE_STORAGE = 1;
export const TYPE_STREAM = 2;
export const TYPE_ROOT = 5;

/**
 * The stream size at or above which a stream lives in the regular FAT rather than the mini stream
 * ([MS-CFB] 2.2, header offset 56). The spec fixes it at 4096, so the reader checks the header's value
 * against this rather than believing it: a crafted 0 or 0xFFFFFFFF routes every stream through the
 * wrong allocator, and since both destinations are bounds-checked the result is a module's source read
 * back as something nobody wrote rather than a clean rejection.
 */
export const MINI_STREAM_CUTOFF = 4096;

/** A directory entry is a fixed 128 bytes ([MS-CFB] 2.6.1), whatever the sector size. */
export const DIR_ENTRY_SIZE = 128;

/** 32 UTF-16 code units including the NUL terminator. This is also why a VBA module name is capped
 * at 31: the module is a stream, and the stream's name is the module's. */
export const MAX_NAME_CHARS = 31;

/**
 * Whether two directory-entry names are the same entry. [MS-CFB] 2.6.4 orders siblings by length and
 * then by their uppercase forms, and forbids two siblings that compare equal, so a host looking a name
 * up treats `dir` and `DIR` as one entry. A lookup here does the same, or a file whose producer cased a
 * name differently would read as missing a stream it has.
 */
export function sameEntryName(a: string, b: string): boolean {
  return a.length === b.length && a.toUpperCase() === b.toUpperCase();
}
