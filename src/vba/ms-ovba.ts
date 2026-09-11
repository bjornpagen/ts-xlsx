// MS-OVBA §2.4.1: compression and decompression of a "CompressedContainer".
//
// VBA module source and the project `dir` stream are stored in Office's own run-length compression,
// NOT deflate. A container is a 0x01 signature byte followed by one or more chunks; each chunk
// decompresses to at most 4096 bytes and is either a raw 4096-byte copy or a stream of literal/copy
// tokens. Reference: [MS-OVBA] 2.4.1.3.6 (decompressing a CompressedContainer), 2.4.1.3.19.3 (the
// CopyToken bit-packing), and 2.4.1.3.7 (compressing a chunk).
//
// The decompressor is a hostile-input parser: the container comes from an untrusted file, so every
// length and back-reference is bounds-checked and the total output is capped. A malformed container
// fails closed with a VbaParseError rather than over-allocating, looping, or reading out of bounds. The
// compressor is the authoring inverse: it is fed our own bytes, and its output re-expands to the input
// byte-for-byte (the round-trip is the correctness contract).

import {hex} from '../hex.ts';
import {readU16} from './bytes.ts';
import {VbaParseError} from './errors.ts';

// A decompressed chunk covers at most 4096 bytes; both directions honour this window ([MS-OVBA]
// 2.4.1.3.6). The chunk header's bits 12-14 carry a fixed 0b011 signature, bit 15 the compressed flag.
const MAX_CHUNK_DECOMPRESSED = 4096;
const CHUNK_SIGNATURE = 0b011 << 12;
const CHUNK_SIGNATURE_MASK = 0b111 << 12;
const CHUNK_COMPRESSED_FLAG = 0x8000;

// The encoder's match search: how far down a prefix's chain it will walk, and how wide the table of
// chains is. The cap is what turns "how compressible is this data" from a time bound into a size one,
// and 64 is well past the point where a further candidate can pay for the token it might shorten.
const MAX_MATCH_CANDIDATES = 64;
const HASH_BITS = 12;
const NO_CANDIDATE = -1;

// A single VBA project is well under a megabyte; 64 MiB is far above any legitimate container yet
// bounds a decompression bomb (a small container that expands without limit) to a survivable size.
const DEFAULT_MAX_OUTPUT = 64 * 1024 * 1024;

/**
 * Decompress an MS-OVBA CompressedContainer beginning at `start` in `buf`.
 * @param maxOutput hard ceiling on decompressed bytes; exceeding it throws (bomb guard).
 */
export function decompressContainer(
  buf: Uint8Array,
  start = 0,
  maxOutput = DEFAULT_MAX_OUTPUT,
): Uint8Array {
  if (start >= buf.length) {
    throw new VbaParseError(
      `compressed container starts past end of stream (${start} >= ${buf.length})`,
    );
  }
  if (buf[start] !== 0x01) {
    throw new VbaParseError(
      `compressed container must begin with a 0x01 signature byte, found 0x${hex(buf[start] ?? 0, 2)}`,
    );
  }

  const out = new ByteSink(maxOutput);
  let pos = start + 1;

  while (pos + 2 <= buf.length) {
    const header = readU16(buf, pos);
    pos += 2;

    // Bits 0-11: (chunk data size - 1). Bit 15: compressed flag. Bits 12-14: the fixed 0b011 signature.
    const chunkDataSize = (header & 0x0fff) + 1;
    const compressed = (header & CHUNK_COMPRESSED_FLAG) !== 0;
    if ((header & CHUNK_SIGNATURE_MASK) !== CHUNK_SIGNATURE) {
      throw new VbaParseError(`chunk header has a bad 0b011 signature (0x${hex(header, 4)})`);
    }
    const chunkEnd = pos + chunkDataSize;
    if (chunkEnd > buf.length) {
      throw new VbaParseError(`chunk data (${chunkDataSize} bytes) runs past end of stream`);
    }

    if (!compressed) {
      // A raw chunk carries its bytes verbatim. [MS-OVBA] fixes its length at 4096, but the length is
      // read from the header all the same: the bound that matters here is the stream's, checked above,
      // and a shorter raw chunk is what earlier versions of this encoder wrote.
      for (let i = pos; i < chunkEnd; i++) out.push(buf[i] as number);
      pos = chunkEnd;
      continue;
    }

    const chunkStart = out.length; // decompressed offset where this chunk began
    while (pos < chunkEnd) {
      const flagByte = buf[pos++] as number;
      for (let bit = 0; bit < 8 && pos < chunkEnd; bit++) {
        const isCopy = (flagByte >> bit) & 1;
        if (!isCopy) {
          out.push(buf[pos++] as number);
          continue;
        }
        if (pos + 2 > chunkEnd) {
          throw new VbaParseError('copy token truncated at chunk end');
        }
        const token = readU16(buf, pos);
        pos += 2;
        const {lengthMask, bitCount} = copyTokenHelp(out.length - chunkStart);
        const length = (token & lengthMask) + 3;
        const copyOffset = (token >> (16 - bitCount)) + 1;
        const src = out.length - copyOffset;
        if (src < chunkStart) {
          // A back-reference may only reach data emitted since this chunk began ([MS-OVBA] 2.4.1.3.19).
          throw new VbaParseError('copy token references before the start of its chunk');
        }
        // Byte-by-byte so overlapping runs (run-length expansion) grow correctly.
        for (let i = 0; i < length; i++) out.push(out.at(src + i));
      }
      if (out.length - chunkStart > MAX_CHUNK_DECOMPRESSED) {
        // [MS-OVBA] 2.4.1.3.6 caps a chunk at 4096 decompressed bytes, and the cap is load-bearing
        // rather than advisory: past it `copyTokenHelp` returns a bit split no producer emits (13+
        // offset bits, a length mask collapsed to 7), so every later token in the chunk decodes into
        // bytes nobody wrote. Excel rejects such a container; accepting it would mean reading a
        // module's source as something its author never compiled.
        throw new VbaParseError(
          `chunk decompressed to more than ${MAX_CHUNK_DECOMPRESSED} bytes, which [MS-OVBA] forbids`,
        );
      }
    }
    pos = chunkEnd;
  }

  return out.bytes();
}

/**
 * Bytes under construction: a growable `Uint8Array` behind a length cursor rather than a `number[]`.
 * Both directions of this codec accumulate into one, because the reason is the same on both sides -
 * the decoder's ceiling admits 64 MiB and the encoder rewrites a whole `dir` stream, and that many
 * boxed slots cost several times their byte count in real memory before the conversion to bytes ever
 * happens. The cursor is what makes the representation work for a run-length decoder: a
 * back-reference reads a byte this same sink already wrote, and {@link at} stays valid across a grow
 * because the buffer is copied whole. {@link set} is what makes it work for the encoder, whose flag
 * byte is written before the eight tokens it describes are known.
 *
 * The optional ceiling is the decode path's bomb guard, enforced on every write rather than after
 * each token so a hostile container can never provoke an allocation past it: growth is capped at the
 * ceiling too. The encoder is fed our own bytes and passes none.
 */
class ByteSink {
  #buf: Uint8Array;
  #length = 0;
  readonly #limit: number;

  constructor(limit = Infinity) {
    this.#limit = limit;
    // One chunk's worth to start, but never more than the whole container is allowed to produce.
    this.#buf = new Uint8Array(Math.min(MAX_CHUNK_DECOMPRESSED, limit));
  }

  get length(): number {
    return this.#length;
  }

  at(index: number): number {
    return this.#buf[index] as number;
  }

  /** Overwrite an already-written byte, the encoder's deferred flag byte. */
  set(index: number, byte: number): void {
    this.#buf[index] = byte;
  }

  push(byte: number): void {
    if (this.#length >= this.#limit) {
      throw new VbaParseError(
        `decompressed output exceeds the ${this.#limit}-byte ceiling (possible bomb)`,
      );
    }
    if (this.#length === this.#buf.length) {
      const grown = new Uint8Array(
        Math.min(Math.max(this.#buf.length * 2, MAX_CHUNK_DECOMPRESSED), this.#limit),
      );
      grown.set(this.#buf);
      this.#buf = grown;
    }
    this.#buf[this.#length++] = byte;
  }

  pushAll(bytes: Uint8Array): void {
    for (const byte of bytes) this.push(byte);
  }

  /** The bytes written so far, as an exactly-sized array that does not alias the growth buffer. */
  bytes(): Uint8Array {
    return this.#buf.slice(0, this.#length);
  }
}

/**
 * Compress `data` into an MS-OVBA CompressedContainer: the inverse of {@link decompressContainer}.
 * Every 4096-decompressed-byte window is emitted as a compressed chunk of literal and copy tokens, or
 * stored raw when those tokens would not fit the 4096 bytes a chunk's data may hold.
 *
 * The result re-expands to `data` byte-for-byte, with one exception the format itself prescribes: a
 * final window that is shorter than 4096 bytes and does not compress (only possible past about 3640
 * bytes, where a literal's flag-bit overhead first overflows the chunk) is stored raw and padded with
 * `0x00` to 4096, so it re-expands with those zeros after it ([MS-OVBA] 2.4.1.3.10).
 */
export function compressContainer(data: Uint8Array): Uint8Array {
  const out = new ByteSink();
  out.push(0x01); // container signature; an empty input yields just this byte
  for (let start = 0; start < data.length; start += MAX_CHUNK_DECOMPRESSED) {
    const chunk = data.subarray(start, Math.min(start + MAX_CHUNK_DECOMPRESSED, data.length));
    const tokens = compressChunk(chunk);
    // [MS-OVBA] 2.4.1.3.7 goes raw only when the token stream overflows the chunk, and a raw chunk is
    // always 4096 bytes: its size field MUST be 4095 (2.4.1.1.5), and a conforming decompressor reads
    // exactly 4096 bytes from one (2.4.1.3.3). Storing a short chunk raw whenever its tokens were no
    // smaller, as this did, wrote 1-byte and 4095-byte raw chunks a conforming reader overruns.
    const compressed = tokens.length <= MAX_CHUNK_DECOMPRESSED;
    const body = compressed ? tokens : padToChunk(chunk);
    const header =
      (compressed ? CHUNK_COMPRESSED_FLAG : 0) | CHUNK_SIGNATURE | ((body.length - 1) & 0x0fff);
    out.push(header & 0xff);
    out.push((header >> 8) & 0xff);
    out.pushAll(body);
  }
  return out.bytes();
}

// A raw chunk's data at the one length the format allows it, zero-filled past the window's own bytes.
function padToChunk(chunk: Uint8Array): Uint8Array {
  if (chunk.length === MAX_CHUNK_DECOMPRESSED) return chunk;
  const padded = new Uint8Array(MAX_CHUNK_DECOMPRESSED);
  padded.set(chunk);
  return padded;
}

// Encode one decompressed chunk (≤ 4096 bytes) as a sequence of MS-OVBA token groups: a flag byte whose
// bits mark the next up-to-8 tokens as literal (0) or copy (1). A copy token replaces a run of 3+ bytes
// that recurs earlier in the *same* chunk; matches may overlap the current position (run-length growth),
// which the decompressor reproduces byte-by-byte. The bit split between the offset and length fields
// widens as the chunk fills, exactly as the decoder computes it, so both agree on every token's shape.
//
// Candidates come from a hash chain over three-byte prefixes rather than a rescan of the whole back
// window. That rescan made the encoder quadratic in the chunk size, and the size is a cost an
// untrusted file gets to choose: `removeVbaModule` and `addVbaReference` recompress a `dir` stream
// that arrived inside an `.xlsm` the caller did not write, and data with no matches cost 95 ms per
// 4 KB chunk. The chain gives up no token the rescan would have emitted: a run under three bytes is
// never worth a copy token, so a prefix that collides with nothing could not have been encoded anyway.
function compressChunk(chunk: Uint8Array): Uint8Array {
  const tokens = new ByteSink();
  // Walking a chain from its head yields candidates newest-first, so equal-length matches keep the
  // smallest offset (a marginally cheaper token), the tie-break the nearest-first rescan had.
  const heads = new Int32Array(1 << HASH_BITS).fill(NO_CANDIDATE);
  const prev = new Int32Array(MAX_CHUNK_DECOMPRESSED).fill(NO_CANDIDATE);
  const remember = (at: number): void => {
    if (at + 2 >= chunk.length) return; // no three-byte prefix starts here, so nothing can match it
    const bucket = hash3(chunk, at);
    prev[at] = heads[bucket] as number;
    heads[bucket] = at;
  };

  let pos = 0;
  while (pos < chunk.length) {
    const flagIndex = tokens.length;
    tokens.push(0);
    let flags = 0;
    for (let bit = 0; bit < 8 && pos < chunk.length; bit++) {
      const {lengthMask, bitCount} = copyTokenHelp(pos);
      const maxLength = lengthMask + 3;
      const windowStart = Math.max(0, pos - (1 << bitCount));

      let bestLength = 0;
      let bestOffset = 0;
      let candidate = pos + 2 < chunk.length ? (heads[hash3(chunk, pos)] as number) : NO_CANDIDATE;
      for (let tried = 0; candidate >= windowStart && tried < MAX_MATCH_CANDIDATES; tried++) {
        let len = 0;
        while (
          len < maxLength &&
          pos + len < chunk.length &&
          chunk[candidate + len] === chunk[pos + len]
        ) {
          len++;
        }
        if (len > bestLength) {
          bestLength = len;
          bestOffset = pos - candidate;
          if (bestLength === maxLength) break; // cannot improve
        }
        candidate = prev[candidate] as number;
      }

      if (bestLength >= 3) {
        const token = ((bestOffset - 1) << (16 - bitCount)) | (bestLength - 3);
        tokens.push(token & 0xff);
        tokens.push((token >> 8) & 0xff);
        flags |= 1 << bit;
        // Every position a copy token covers still enters the chain: a later match may begin inside
        // the run this one emitted.
        for (let i = 0; i < bestLength; i++) remember(pos + i);
        pos += bestLength;
      } else {
        tokens.push(chunk[pos] as number);
        remember(pos);
        pos++;
      }
    }
    tokens.set(flagIndex, flags);
  }
  return tokens.bytes();
}

// The three-byte prefix a copy token is built on, folded into one bucket per byte of the largest
// window a chunk can offer.
function hash3(chunk: Uint8Array, pos: number): number {
  const first = chunk[pos] as number;
  const second = chunk[pos + 1] as number;
  const third = chunk[pos + 2] as number;
  return ((first << 8) ^ (second << 4) ^ third) & ((1 << HASH_BITS) - 1);
}

/**
 * Bit split for a CopyToken given how many bytes have been emitted since the current chunk began
 * ([MS-OVBA] 2.4.1.3.19.3): the offset field grows and the length field shrinks as the chunk fills.
 */
function copyTokenHelp(decompressedSoFar: number): {lengthMask: number; bitCount: number} {
  // Bounded at both ends: the floor of 4 is the spec's, and the ceiling of 12 follows from the 4096-byte
  // chunk window, since `log2(4096)` is 12 and no legal chunk can have emitted more. Without it a
  // decoder that has somehow run past the window keeps widening the offset field until the length mask
  // is 7 bits wide, which decodes later tokens under a split the format never defines.
  const bitCount = Math.min(
    Math.max(Math.ceil(Math.log2(Math.max(decompressedSoFar, 1))), 4),
    Math.log2(MAX_CHUNK_DECOMPRESSED),
  );
  const lengthMask = 0xffff >> bitCount;
  return {lengthMask, bitCount};
}
