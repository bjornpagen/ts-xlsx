import assert from 'node:assert/strict';
import {test} from 'node:test';

import {XlsbParseError} from './errors.ts';
import {readRecords} from './record-stream.ts';
import {concat, frame} from './records.test-support.ts';

test('a one-byte type and one-byte size frame a short record', () => {
  const records = [...readRecords(frame(2, Uint8Array.of(1, 2, 3)))];
  assert.deepEqual(records, [{type: 2, data: Uint8Array.of(1, 2, 3)}]);
});

test("the spec's own worked example decodes to BrtCommentText with 200 payload bytes", () => {
  // [MS-XLSB] 2.1.4 walks these four header bytes through by hand: type 125 + 4*128 = 637, size
  // 72 + 128 = 200. Decoding the spec's example is the tightest possible check on both prefixes.
  const part = concat(
    Uint8Array.of(0b11111101, 0b00000100, 0b11001000, 0b00000001),
    new Uint8Array(200),
  );
  const records = [...readRecords(part)];
  assert.equal(records.length, 1);
  assert.equal(records[0]?.type, 637);
  assert.equal(records[0]?.data.length, 200);
});

test('a type of 128 or above takes a second byte and still round-trips', () => {
  for (const type of [128, 156, 615, 3072, 16383]) {
    const records = [...readRecords(frame(type, Uint8Array.of(9)))];
    assert.equal(records[0]?.type, type, `type ${type}`);
  }
});

test('a size crossing each 7-bit boundary takes another byte and still round-trips', () => {
  for (const size of [0, 1, 127, 128, 16383, 16384, 2097151, 2097152]) {
    const records = [...readRecords(frame(1, new Uint8Array(size)))];
    assert.equal(records[0]?.data.length, size, `size ${size}`);
  }
});

test('records are yielded in order and the whole part is consumed', () => {
  const part = concat(
    frame(129, new Uint8Array(0)),
    frame(0, Uint8Array.of(7)),
    frame(2, Uint8Array.of(1, 2, 3, 4)),
    frame(130, new Uint8Array(0)),
  );
  assert.deepEqual(
    [...readRecords(part)].map((record) => record.type),
    [129, 0, 2, 130],
  );
});

test('an empty part yields no records', () => {
  assert.deepEqual([...readRecords(new Uint8Array(0))], []);
});

test('a payload is a view onto the part, not a copy', () => {
  const part = frame(2, Uint8Array.of(1, 2, 3));
  const [record] = [...readRecords(part)];
  assert.ok(record);
  assert.equal(record.data.buffer, part.buffer);
});

test('a record whose declared size runs past the end of the part is rejected', () => {
  // A hostile file's whole lever is the declared length. Framing says 4000 bytes follow; three do.
  const part = concat(Uint8Array.of(2, 0xa0, 0x1f), Uint8Array.of(1, 2, 3));
  assert.throws(() => [...readRecords(part)], XlsbParseError);
});

test('a part ending mid-header is rejected rather than read as undefined', () => {
  for (const truncated of [
    Uint8Array.of(0x80), // two-byte type, second byte missing
    Uint8Array.of(2), // type only, no size
    Uint8Array.of(2, 0x80), // size continues but the stream does not
  ]) {
    assert.throws(() => [...readRecords(truncated)], XlsbParseError);
  }
});

test('the fourth size byte terminates the size regardless of its continuation bit', () => {
  // [MS-XLSB]: "The high bit of the fourth byte MUST be ignored." A file that sets it anyway must
  // not drag a fifth byte into the size: that would shift the payload and desync the whole part.
  const part = concat(Uint8Array.of(1, 0x83, 0x80, 0x80, 0x80), new Uint8Array(3));
  const [record] = [...readRecords(part)];
  assert.equal(record?.data.length, 3);
});

test('framing is lazy: a caller that stops early does not frame the rest of the part', () => {
  // The tail is deliberately malformed: reaching it would throw, so completing without one proves
  // the generator never looked past the record the caller asked for.
  const part = concat(frame(2, Uint8Array.of(1)), Uint8Array.of(3, 0xff, 0xff, 0xff, 0x7f));
  for (const record of readRecords(part)) {
    assert.equal(record.type, 2);
    break;
  }
});
