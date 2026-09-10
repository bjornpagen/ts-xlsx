// The TLV walk, on its own, away from the container that normally carries it.
//
// The round-trip tests over a real `vbaProject.bin` are the proof that the walk is right; these are
// the proof of *why* it is written the way it is, which a passing round-trip does not show.

import {strict as assert} from 'node:assert';
import {test} from 'node:test';

import {dirRecords, REC_MODULES_COUNT, REC_PROJECT_VERSION} from './dir-records.ts';
import {VbaParseError} from './errors.ts';
import {rec, u16le, u32le} from './vba.test-support.ts';

test('dirRecords walks a flat TLV stream, reporting each record and where the next begins', () => {
  const stream = Uint8Array.from([
    ...rec(0x0003, u16le(1252)),
    ...rec(REC_MODULES_COUNT, u16le(2)),
  ]);
  assert.deepEqual(
    [...dirRecords(stream, 'overruns stream')],
    [
      {id: 0x0003, recordStart: 0, dataStart: 6, size: 2, end: 8},
      {id: REC_MODULES_COUNT, recordStart: 8, dataStart: 14, size: 2, end: 16},
    ],
  );
});

test("dirRecords skips PROJECTVERSION's uncounted VersionMinor, so the record after it is aligned", () => {
  // Size=4 counts only VersionMajor; the trailing 2-byte VersionMinor is not counted. A walk that
  // trusted Size alone would start the next record two bytes early, in the middle of this payload.
  const stream = Uint8Array.from([
    ...rec(REC_PROJECT_VERSION, u32le(4)),
    ...u16le(0x000a),
    ...rec(REC_MODULES_COUNT, u16le(2)),
  ]);
  const records = [...dirRecords(stream, 'overruns stream')];
  assert.equal(records[0]?.size, 4, 'the Size field is reported verbatim');
  assert.equal(records[0]?.end, 12, 'but the next record is two bytes further on');
  assert.equal(records[1]?.id, REC_MODULES_COUNT);
  assert.equal(records[1]?.size, 2);
});

test('dirRecords refuses a record whose payload runs past the stream, naming the operation', () => {
  const truncated = Uint8Array.from([...u16le(0x0019), ...u32le(64), 0x41, 0x42]);
  assert.throws(
    () => [...dirRecords(truncated, 'overruns while removing a module')],
    (error: unknown) => {
      assert.ok(error instanceof VbaParseError);
      assert.match(error.message, /0x19/);
      assert.match(error.message, /overruns while removing a module/);
      return true;
    },
  );
});

test('dirRecords stops at a trailing fragment too short to be a record', () => {
  const stream = Uint8Array.from([...rec(0x0010, []), 0x00, 0x00, 0x00]);
  assert.deepEqual(
    [...dirRecords(stream, 'overruns stream')].map((r) => r.id),
    [0x0010],
  );
});
