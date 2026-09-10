// Byte primitives for the VBA subsystem. Every structure under `src/vba/` is a little-endian binary record
// ([MS-CFB] sectors and directory entries, [MS-OVBA] `dir` TLVs and compressed chunks), so these are
// what its parsers read with and what its splices write with.
//
// The bound is checked here, once, rather than at each caller. A plain `buf[at] | (buf[at + 1] << 8)`
// reads `undefined | (undefined << 8)` past the end, which is `0`, so a truncated `vbaProject.bin` out of
// an untrusted `.xlsm` would parse as a file full of zeros instead of failing. `DataView.getUint16` is
// the obvious fix but the wrong one here: it throws `RangeError`, which is outside this package's
// failure taxonomy, and it is slow. Measured at 20M reads, a `DataView` constructed per call runs ~90x
// slower than the raw index and one cached per buffer ~8x, while the explicit `undefined` check below
// is indistinguishable from the unchecked read: V8 already bounds-checks the load, so the branch is
// free. `ms-ovba.ts` calls `readU16` once per copy token, so that difference is not academic.

// Neither joining chunks nor decoding UTF-16LE is a VBA concern; this subsystem is simply one of
// several callers of each, and `src/bytes.ts` is where both live. Re-exported rather than imported
// through, so a parser here reaches for one module and gets the whole vocabulary.
export {concat, decodeUtf16le} from '../bytes.ts';

import {VbaParseError} from './errors.ts';

function truncated(at: number, need: number, length: number): VbaParseError {
  return new VbaParseError(
    `read of ${need} bytes at offset ${at} runs past the end of a ${length}-byte buffer`,
  );
}

/** Read a little-endian `uint16`. @throws {VbaParseError} if the two bytes are not both in `buf`. */
export function readU16(buf: Uint8Array, at: number): number {
  const b0 = buf[at];
  const b1 = buf[at + 1];
  if (b0 === undefined || b1 === undefined) throw truncated(at, 2, buf.length);
  return b0 | (b1 << 8);
}

/** Read a little-endian `uint32`. @throws {VbaParseError} if the four bytes are not all in `buf`. */
export function readU32(buf: Uint8Array, at: number): number {
  const b0 = buf[at];
  const b1 = buf[at + 1];
  const b2 = buf[at + 2];
  const b3 = buf[at + 3];
  if (b0 === undefined || b1 === undefined || b2 === undefined || b3 === undefined) {
    throw truncated(at, 4, buf.length);
  }
  return (b0 | (b1 << 8) | (b2 << 16) | (b3 << 24)) >>> 0;
}

/**
 * Write a little-endian `uint16` over bytes already in `buf`. The value is asserted rather than
 * masked: a caller that arrived at `-1` or `0x1_0000` has miscounted something, and silently storing
 * `0xffff` turns that into a `dir` stream declaring 65,535 modules.
 */
export function writeU16(buf: Uint8Array, at: number, value: number): void {
  if (!Number.isInteger(value) || value < 0 || value > 0xffff) {
    throw new VbaParseError(`${value} does not fit the uint16 at offset ${at}`);
  }
  if (at + 1 >= buf.length) throw truncated(at, 2, buf.length);
  buf[at] = value & 0xff;
  buf[at + 1] = (value >> 8) & 0xff;
}

/** Write a little-endian `uint32` over bytes already in `buf`, asserted rather than masked as {@link writeU16} is. */
export function writeU32(buf: Uint8Array, at: number, value: number): void {
  if (!Number.isInteger(value) || value < 0 || value > 0xffff_ffff) {
    throw new VbaParseError(`${value} does not fit the uint32 at offset ${at}`);
  }
  if (at + 3 >= buf.length) throw truncated(at, 4, buf.length);
  buf[at] = value & 0xff;
  buf[at + 1] = (value >>> 8) & 0xff;
  buf[at + 2] = (value >>> 16) & 0xff;
  buf[at + 3] = (value >>> 24) & 0xff;
}

/**
 * Replace `src[start..end)` with `insert`, returning a new array.
 *
 * The three record edits the project editor makes are all this shape: splice a MODULE record block in,
 * cut one out, cut a PROJECTwm entry out. Each was an allocate-and-copy-around written out longhand,
 * and each had to get the same three offsets right; saying it once means the arithmetic is checked
 * once. Passing no `insert` is a pure deletion, and `start === end` a pure insertion.
 */
export function spliceBytes(
  src: Uint8Array,
  start: number,
  end: number,
  insert: Uint8Array = EMPTY,
): Uint8Array {
  const out = new Uint8Array(src.length - (end - start) + insert.length);
  out.set(src.subarray(0, start), 0);
  out.set(insert, start);
  out.set(src.subarray(end), start + insert.length);
  return out;
}

const EMPTY = new Uint8Array(0);
