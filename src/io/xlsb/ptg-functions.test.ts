// The `Ftab` transcription, checked against what the specification says about its shape.
//
// About six hundred function names typed out of a normative table, with exactly one failure mode: a
// slip in the typing. It is invisible at runtime, because a name in the wrong slot still decodes, just
// to a different function, for every formula that cites that index. Nothing verified the
// transcription, so a slip made when it was written would still be there.
//
// These assert facts stated by [MS-XLS] 2.5.198.17 and by this module's own header, deliberately not
// re-derived from the runs the table is written as: a test that recomputed the table from the same
// literals would only prove the reduce works.

import assert from 'node:assert/strict';
import {test} from 'node:test';

import {FTAB, FTAB_USER_DEFINED, functionNameFor} from './ptg-functions.ts';

test('the function table runs from COUNT past RTD, where [MS-XLS] stops, to AVERAGEIFS', () => {
  assert.equal(functionNameFor(0x0000), 'COUNT', 'index 0 is COUNT');
  assert.equal(functionNameFor(0x017b), 'RTD', 'RTD ends the [MS-XLS] table');
  // As Excel 16.0 (build 20326) cited them in an `.xlsb` it saved.
  assert.equal(functionNameFor(0x017c), 'CUBEVALUE', 'and BIFF12 carries on at CUBEVALUE');
  assert.equal(functionNameFor(0x01d0), 'RANDBETWEEN');
  assert.equal(functionNameFor(0x01e0), 'IFERROR');
  assert.equal(functionNameFor(0x01e4), 'AVERAGEIFS', 'to AVERAGEIFS');
  assert.equal(FTAB.length, 0x01e5, 'so the table is exactly that many slots long');
  assert.equal(functionNameFor(0x01e5), undefined, 'nothing past it decodes');
});

test('the table has exactly the seven undefined indices the specification leaves undefined', () => {
  // Four gaps, as the module header describes them, spanning seven indices. An index here names no
  // function, so a formula citing one is not decodable rather than decodable as something adjacent.
  // An index loop, not `flatMap`: the table is built by assigning into a bare array, so its holes are
  // genuinely absent rather than present-and-undefined, and every iteration method skips them.
  const holes: number[] = [];
  for (let index = 0; index < FTAB.length; index++) {
    if (FTAB[index] === undefined) holes.push(index);
  }
  assert.deepEqual(holes, [0x00ca, 0x00cb, 0x00d9, 0x00da, 0x00f9, 0x00fa, 0x014d]);
});

test('every recorded arity is a count a call can actually pop', () => {
  // An arity is spent as `stack.splice(stack.length - arity, arity)`. A negative or fractional one
  // does not throw there: it silently takes a different run of operands, and the formula decodes to
  // something the file never said. That an arity *exists* is now the type's problem; that it is a
  // count is this one's.
  const wrong = FTAB.flatMap((entry) =>
    entry === undefined || entry[1] === 'variadic' || (Number.isInteger(entry[1]) && entry[1] >= 0)
      ? []
      : [`${entry[0]} takes ${entry[1]}`],
  );
  assert.deepEqual(wrong, []);
});

test('the user-defined index is a real slot, and the one the decoder reaches for', () => {
  assert.equal(FTAB_USER_DEFINED, 0x00ff);
  assert.ok(
    functionNameFor(FTAB_USER_DEFINED) !== undefined,
    'it names a function rather than being one of the gaps',
  );
});

test('no function name is transcribed into two slots', () => {
  // A duplicate is what a copy-paste slip during transcription looks like: the second occurrence sits
  // where some other function should be, and that function then decodes as this one.
  const seen = new Map<string, number>();
  const duplicates: string[] = [];
  for (let index = 0; index < FTAB.length; index++) {
    const name = FTAB[index]?.[0];
    if (name === undefined) continue;
    const first = seen.get(name);
    if (first === undefined) seen.set(name, index);
    else duplicates.push(`${name} at 0x${first.toString(16)} and 0x${index.toString(16)}`);
  }
  assert.deepEqual(duplicates, []);
});
