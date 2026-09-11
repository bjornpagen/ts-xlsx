import assert from 'node:assert/strict';
import {test} from 'node:test';

import {XlsbParseError} from './errors.ts';
import {RecordReader} from './primitives.ts';
import {concat, wide, word} from './records.test-support.ts';
import {SheetProtectionRecords} from './sheet-protection.ts';

// `fLocked` and the fifteen operation flags, each a Bool32, with `allowed` naming the operations set.
function flags(locked: boolean, allowed: readonly number[] = []): Uint8Array {
  return concat(
    word(locked ? 1 : 0),
    ...Array.from({length: 15}, (_, i) => word(allowed.includes(i) ? 1 : 0)),
  );
}

function legacy(protpwd: number, body: Uint8Array): Uint8Array {
  const pwd = new Uint8Array(2);
  new DataView(pwd.buffer).setUint16(0, protpwd, true);
  return concat(pwd, body);
}

test('a sheet whose fLocked is clear is not protected, whatever its flags say', () => {
  const records = new SheetProtectionRecords();
  records.legacy(new RecordReader(legacy(0, flags(false, [0, 1, 2]))));
  assert.equal(records.result(), undefined);
});

test('a flag is recorded only where it departs from what the XML twin would leave unstated', () => {
  // Excel's plain Protect(): objects and scenarios not allowed, both selections allowed.
  const records = new SheetProtectionRecords();
  records.legacy(new RecordReader(legacy(0, flags(true, [10, 14]))));
  assert.deepEqual(records.result(), {flags: {objects: false, scenarios: false}});
});

test('a legacy password hash is kept as the four hex digits the XML twin writes', () => {
  const records = new SheetProtectionRecords();
  records.legacy(new RecordReader(legacy(0xcc3d, flags(true, [0, 1, 10, 14]))));
  assert.deepEqual(records.result(), {flags: {}, legacyPasswordHash: 'CC3D'});
});

test('the agile record carries the credential in the spelling the XML twin uses', () => {
  const records = new SheetProtectionRecords();
  records.legacy(new RecordReader(legacy(0, flags(true, [0, 1, 10, 14]))));
  records.iso(
    new RecordReader(
      concat(
        word(100000),
        flags(true, [0, 1, 10, 14]),
        word(2),
        Uint8Array.of(0xcb, 0xf1),
        word(1),
        Uint8Array.of(0xa0),
        wide('SHA-512'),
      ),
    ),
  );
  assert.deepEqual(records.result(), {
    flags: {},
    credential: {algorithmName: 'SHA-512', hashValue: 'y/E=', saltValue: 'oA==', spinCount: 100000},
  });
});

test('a truncated protection record fails as malformed input', () => {
  const records = new SheetProtectionRecords();
  assert.throws(() => records.legacy(new RecordReader(legacy(0, word(1)))), XlsbParseError);
  assert.throws(
    () => records.iso(new RecordReader(concat(word(1), flags(true), word(64)))),
    XlsbParseError,
  );
});
