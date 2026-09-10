// BIFF12 bytes for the tests that need a record the fixture does not carry: the record framing, the
// primitive encodings a payload is made of, and a way to change one record of a real part.
//
// Not a `.test.ts` file: `node --test` would try to run it and find no tests. `tsconfig.build.json`
// excludes the `.test-support.ts` suffix, so it ships nowhere.

import {readRecords} from './record-stream.ts';

/**
 * Frame a record the way [MS-XLSB] 2.1.4 says a writer must, so a test states the *encoding* rather
 * than a hand-copied byte soup: 7 bits per prefix byte, high bit meaning "one more".
 */
export function frame(type: number, payload: Uint8Array): Uint8Array {
  const header: number[] = [];
  if (type < 0x80) header.push(type);
  else header.push((type & 0x7f) | 0x80, (type >> 7) & 0x7f);
  let size = payload.length;
  do {
    const piece = size & 0x7f;
    size >>>= 7;
    header.push(size > 0 ? piece | 0x80 : piece);
  } while (size > 0);
  return Uint8Array.from([...header, ...payload]);
}

export function concat(...parts: Uint8Array[]): Uint8Array {
  return Uint8Array.from(parts.flatMap((part) => [...part]));
}

/** A little-endian 32-bit word, the way every fixed-width field sits in a record. */
export function word(value: number): Uint8Array {
  const out = new Uint8Array(4);
  new DataView(out.buffer).setUint32(0, value >>> 0, true);
  return out;
}

/** An `XLWideString`: a 4-byte character count then the UTF-16LE code units. */
export function wide(text: string): Uint8Array {
  const out = new Uint8Array(4 + text.length * 2);
  const view = new DataView(out.buffer);
  view.setUint32(0, text.length, true);
  for (let index = 0; index < text.length; index++) {
    view.setUint16(4 + index * 2, text.charCodeAt(index), true);
  }
  return out;
}

/**
 * A formula token stream citing defined name `index` (1-based) `count` times, joined by `&`: 3,599
 * bytes for 600 citations. Over a long name, the stream whose decoded text is that name's length
 * times `count`, while its own length stays five bytes a citation.
 */
export function nameCitations(count: number, index = 1): Uint8Array {
  const cite = concat(Uint8Array.of(0x43), word(index));
  return concat(cite, ...Array.from({length: count - 1}, () => concat(cite, Uint8Array.of(0x08))));
}

/**
 * A part re-framed record by record, each replaced by the bytes `edit` returns for it. Returning
 * `frame(type, data)` keeps a record as it was; returning more than one framed record inserts some.
 * How a test changes one field of an Excel-authored part without hand-copying the rest of it.
 */
export function reframe(
  part: Uint8Array,
  edit: (type: number, data: Uint8Array) => Uint8Array,
): Uint8Array {
  return concat(...[...readRecords(part)].map(({type, data}) => edit(type, data)));
}
