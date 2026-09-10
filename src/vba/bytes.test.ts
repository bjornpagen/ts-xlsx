// The byte primitives. The bound belongs to the read, not to each caller remembering to guard: a
// `vbaProject.bin` comes out of an untrusted `.xlsm`, and a truncated one must fail closed rather than
// parse as zeros.

import {strict as assert} from 'node:assert';
import {test} from 'node:test';

import {decodeUtf16le, readU16, readU32} from './bytes.ts';
import {VbaParseError} from './errors.ts';

test('readU16/readU32 read the last valid offset and refuse the one past it', () => {
  const buf = Uint8Array.of(0x01, 0x02, 0x03, 0x04, 0x05, 0x06);

  assert.equal(readU16(buf, 0), 0x0201);
  assert.equal(readU16(buf, 4), 0x0605, 'the last offset holding two whole bytes');
  assert.equal(readU32(buf, 0), 0x04030201);
  assert.equal(readU32(buf, 2), 0x06050403, 'the last offset holding four whole bytes');

  for (const [read, at] of [
    [readU16, 5],
    [readU16, 6],
    [readU32, 3],
    [readU32, 100],
  ] as const) {
    assert.throws(
      () => read(buf, at),
      (err: unknown) => err instanceof VbaParseError && /past the end/.test((err as Error).message),
      `a read at ${at} is truncated, and says so`,
    );
  }
});

test('readU32 returns an unsigned value with the high bit set', () => {
  assert.equal(readU32(Uint8Array.of(0xff, 0xff, 0xff, 0xff), 0), 0xffffffff);
});

test('decodeUtf16le drops a trailing half code unit rather than reading past it', () => {
  assert.equal(decodeUtf16le(Uint8Array.of(0x56, 0x00, 0x42, 0x00, 0x41, 0x00)), 'VBA');
  assert.equal(decodeUtf16le(Uint8Array.of(0x56, 0x00, 0x42)), 'V');
  assert.equal(decodeUtf16le(new Uint8Array()), '');
});
