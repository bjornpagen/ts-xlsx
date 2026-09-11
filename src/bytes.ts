// The byte primitive every layer needs. It sits beside `errors.ts`, below all of them: the ZIP
// container, the XLSX write stream and the CFB reader all end up holding a list of chunks and
// wanting one buffer, and none of them is the right owner of that for the others.

/**
 * Join byte chunks into one buffer.
 *
 * Both parameters exist to avoid work the caller has already done. `size` is the total length when
 * the caller knows it, as a reader draining a length-prefixed stream does, which skips a pass over
 * the chunks. A lone chunk is handed straight back rather than copied, which is the common case on
 * the inflate path, where a part that fits in one chunk would otherwise be duplicated in memory the
 * moment it is read.
 *
 * The returned buffer therefore aliases the caller's chunk when there is exactly one. Every caller
 * here is a reader assembling bytes it then only reads, so this is sound; a caller that means to
 * mutate the result must copy it.
 */
export function concat(chunks: readonly Uint8Array[], size?: number): Uint8Array {
  const first = chunks[0];
  if (chunks.length === 1 && first !== undefined) return first;
  let total = size;
  if (total === undefined) {
    total = 0;
    for (const chunk of chunks) total += chunk.length;
  }
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}

// The base64 alphabet, in index order (RFC 4648 §4).
const BASE64_ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';

/**
 * Standard base64 of `bytes`, padded with `=`.
 *
 * Spelled out rather than delegated because both platform routes are unavailable here: `Buffer` is
 * a Node global, which is exactly what a browser-safe module may not reach for, and `btoa` is
 * declared deprecated in Node's types (so the `no-deprecated` lint rule rejects it) and takes a
 * binary string rather than bytes anyway. The callers encode a 16-byte salt and a 64-byte hash, a
 * sheet password's when one is set and a binary sheet's credential when one is read, so the loop below
 * is not on any path where its cost is measurable.
 */
export function toBase64(bytes: Uint8Array): string {
  let out = '';
  for (let i = 0; i < bytes.length; i += 3) {
    const a = bytes[i] ?? 0;
    const b = bytes[i + 1] ?? 0;
    const c = bytes[i + 2] ?? 0;
    const triple = (a << 16) | (b << 8) | c;
    const remaining = bytes.length - i;
    // `charAt` rather than an index: `noUncheckedIndexedAccess` types `s[i]` as possibly
    // undefined, and the alphabet is exactly 64 characters wide by construction.
    out += BASE64_ALPHABET.charAt((triple >>> 18) & 63);
    out += BASE64_ALPHABET.charAt((triple >>> 12) & 63);
    out += remaining > 1 ? BASE64_ALPHABET.charAt((triple >>> 6) & 63) : '=';
    out += remaining > 2 ? BASE64_ALPHABET.charAt(triple & 63) : '=';
  }
  return out;
}

// UTF-16LE, both directions, because three directories each had an encoder and two had a decoder.
//
// The semantics are identical everywhere it is used and were stated three times in three different
// ways: little-endian code units, and a lone surrogate carried through as the code unit it is rather
// than substituted. That last part is load-bearing in every one of the callers and for a different
// reason each time -- a sheet-protection password is a credential, not text to be displayed, so
// U+FFFD would silently change it; a `.csv` written as UTF-16 loses nothing by passing the unit
// through; an [MS-OVBA] name field is bytes the format defines. One rule, three reasons, one
// implementation.
//
// The little-endian *integer* readers stay where they are: `vba/bytes.ts`'s per-call `DataView` and
// `xlsb/primitives.ts`'s per-record one are each right for their own access pattern, and both say so.

/** A string's UTF-16LE bytes, code unit by code unit, with lone surrogates carried through. */
export function utf16leBytes(text: string): Uint8Array {
  const bytes = new Uint8Array(text.length * 2);
  for (let i = 0; i < text.length; i++) {
    const unit = text.charCodeAt(i);
    bytes[i * 2] = unit & 0xff;
    bytes[i * 2 + 1] = unit >>> 8;
  }
  return bytes;
}

// A decode runs in code-unit batches rather than one `String.fromCharCode` call per character
// (quadratic concatenation) or one spread of every unit (which blows the argument limit on a long
// string). 4096 is comfortably under every engine's limit and makes the batching invisible.
//
// It is not a micro-optimisation here: the [MS-OVBA] copy used the per-character technique this
// comment names as wrong, and was bounded only by the fields it happened to be pointed at (a
// [MS-CFB] directory-entry name is at most 31 code units). The bound is the caller's accident, and a
// field of file-chosen length is one edit away.
const CHARS_PER_BATCH = 4096;

/**
 * Decode UTF-16LE code units. A trailing odd byte is dropped: these fields are length-prefixed by
 * their producer, and half a code unit carries nothing to decode.
 */
export function decodeUtf16le(bytes: Uint8Array): string {
  let text = '';
  let batch: number[] = [];
  for (let i = 0; i + 1 < bytes.length; i += 2) {
    batch.push((bytes[i] as number) | ((bytes[i + 1] as number) << 8));
    if (batch.length === CHARS_PER_BATCH) {
      text += String.fromCharCode(...batch);
      batch = [];
    }
  }
  return batch.length > 0 ? text + String.fromCharCode(...batch) : text;
}
