// The CFB reader: malformed inputs fail closed.

import {strict as assert} from 'node:assert';
import {test} from 'node:test';

import {CompoundFile} from './cfb.ts';
import {VbaParseError} from './errors.ts';
import {buildVbaProjectBin, CODE_PAGE, MODULES} from './vba.test-support.ts';

test('CompoundFile rejects a non-CFB signature', () => {
  assert.throws(() => new CompoundFile(new Uint8Array(512)), VbaParseError);
});

test('CompoundFile rejects a truncated header', () => {
  assert.throws(() => new CompoundFile(new Uint8Array(16)), VbaParseError);
});

test('CompoundFile rejects an out-of-range directory sector', () => {
  const bin = buildVbaProjectBin(CODE_PAGE, MODULES);
  const dv = new DataView(bin.buffer);
  dv.setUint32(48, 9999, true); // first directory sector points past the file
  assert.throws(() => new CompoundFile(bin), VbaParseError);
});
