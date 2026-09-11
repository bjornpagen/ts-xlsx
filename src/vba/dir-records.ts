// The `dir` stream's record grammar: the ids [MS-OVBA] 2.3.4.2 names, the one walk over them, and the
// one way a record is written.
//
// The stream is a flat sequence of TLV records, `Id(u16) Size(u32) data[Size]`, which is simple
// enough that three call sites each wrote their own walk. They were right to agree on the shape and
// wrong to have three chances to disagree about the exception:
//
// **PROJECTVERSION carries VersionMajor (u32) + VersionMinor (u16) after its fixed `Size=4` field,
// but `Size` only accounts for the 4-byte major.** The extra 2-byte minor is uncounted, so a uniform
// walk misaligns by 2 bytes from that record onward unless it skips them explicitly.
//
// Misaligning is not a thrown error. It is a walk that reads a length out of the middle of somebody's
// payload and then either overruns, which is caught, or lands on a plausible-looking boundary, which
// is not. The second is how a structural edit corrupts a macro project without saying anything. One
// walk, stated once, is what stops the fourth copy from being written without the correction.

import {hex} from '../hex.ts';
import {readU16, readU32, writeU16, writeU32} from './bytes.ts';
import {VbaParseError} from './errors.ts';

/** PROJECTCODEPAGE: the code page every MBCS name in the project decodes through. */
export const REC_PROJECT_CODEPAGE = 0x0003;

/** PROJECTVERSION: the record whose `Size` under-counts its payload by the trailing VersionMinor. */
export const REC_PROJECT_VERSION = 0x0009;

/** REFERENCEREGISTERED: a reference to a registered type library, by libid. */
export const REC_REFERENCE_REGISTERED = 0x000d;

/** PROJECTMODULES MODULES_COUNT: how many module blocks follow, and the first record after the
 * reference array, which has no count of its own. */
export const REC_MODULES_COUNT = 0x000f;

/** REFERENCENAME: a reference's MBCS name. */
export const REC_REFERENCE_NAME = 0x0016;

/** MODULENAME: opens a module's record block. */
export const REC_MODULE_NAME = 0x0019;

/** MODULESTREAMNAME: the storage name the module's p-code lives under. */
export const REC_MODULE_STREAMNAME = 0x001a;

/** MODULETYPE for a standard module. */
export const REC_MODULE_TYPE_PROCEDURAL = 0x0021;

/** MODULETYPE for a document module (`ThisWorkbook`, a sheet's code-behind). */
export const REC_MODULE_TYPE_DOCUMENT = 0x0022;

/** MODULEENDOFBLOCK: closes a module's record block. */
export const REC_MODULE_TERMINATOR = 0x002b;

/** MODULEOFFSET: where the module's source begins inside its stream. */
export const REC_MODULE_OFFSET = 0x0031;

/**
 * REFERENCENAME's Unicode half. A *literal* `0x003E` marker rather than a nested record id, but it is
 * laid out as its own Id+Size+data TLV, so a generic walk sees REFERENCENAME as two chained records,
 * exactly like MODULENAME/MODULENAME_UNICODE. Verified against a real Excel-authored dir stream
 * (2026-07-23).
 */
export const REC_REFERENCE_NAME_UNICODE = 0x003e;

/** One record as the walk sees it. */
export interface DirRecord {
  /** The record id; one of the `REC_*` constants above, or a record this library does not consume. */
  readonly id: number;
  /** Offset of the record's `Id` field: where a splice that drops the record cuts from. */
  readonly recordStart: number;
  /** Offset of the record's payload. */
  readonly dataStart: number;
  /** The `Size` field, which is the payload length for every record but PROJECTVERSION. */
  readonly size: number;
  /** Offset of the next record, PROJECTVERSION's uncounted VersionMinor already skipped. Never
   * recompute this from `dataStart + size`: that is the misalignment this module exists to prevent. */
  readonly end: number;
}

/**
 * Walk a decompressed `dir` stream, yielding each record in order. Consumers may stop early.
 *
 * @param context the phrase a truncated record's error carries, naming what the caller was doing:
 *   the message reaches a user who has no idea what a TLV is but does know they were removing a
 *   module.
 * @throws {VbaParseError} when a record's payload runs past the end of the stream.
 */
export function* dirRecords(dir: Uint8Array, context: string): Generator<DirRecord> {
  let pos = 0;
  while (pos + 6 <= dir.length) {
    const id = readU16(dir, pos);
    const size = readU32(dir, pos + 2);
    const dataStart = pos + 6;
    if (dataStart + size > dir.length) {
      throw new VbaParseError(`dir record 0x${hex(id, 4)} ${context}`);
    }
    // The whole reason this walk is not inline at three call sites.
    const end = dataStart + size + (id === REC_PROJECT_VERSION ? 2 : 0);
    yield {id, recordStart: pos, dataStart, size, end};
    pos = end;
  }
}

/**
 * One record's bytes, `Id(u16) Size(u32) data[Size]`: the grammar {@link dirRecords} walks, written.
 * Built as bytes rather than spread into a `number[]`, which overflowed the call stack on a payload of a
 * few hundred thousand bytes. Not for PROJECTVERSION, whose uncounted VersionMinor no edit here writes.
 */
export function dirRecord(id: number, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(6 + data.length);
  writeU16(out, 0, id);
  writeU32(out, 2, data.length);
  out.set(data, 6);
  return out;
}
