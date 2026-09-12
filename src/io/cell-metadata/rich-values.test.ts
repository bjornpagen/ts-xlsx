import assert from 'node:assert/strict';
import {test} from 'node:test';

import {parseRichValueErrors, RICH_VALUE_ERROR_TYPES} from './rich-values.ts';

const NS = 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata';

// The two parts as Excel 16.0 (build 20326) saved them for a sheet holding a spill blocked by a value, a
// FILTER matching nothing, a FIELDVALUE of a number, a bare LAMBDA, a spill blocked by a merge and a
// spill off the sheet's edge.
const STRUCTURES =
  `<rvStructures xmlns="${NS}" count="3">` +
  '<s t="_error"><k n="colOffset" t="i"/><k n="errorType" t="i"/><k n="rwOffset" t="i"/><k n="subType" t="i"/></s>' +
  '<s t="_error"><k n="errorType" t="i"/><k n="subType" t="i"/></s>' +
  '<s t="_error"><k n="errorType" t="i"/><k n="field" t="s"/><k n="subType" t="i"/></s>' +
  '</rvStructures>';
const VALUES =
  `<rvData xmlns="${NS}" count="6">` +
  '<rv s="0"><v>0</v><v>8</v><v>2</v><v>1</v></rv><rv s="1"><v>13</v><v>3</v></rv>' +
  '<rv s="2"><v>12</v><v>x</v><v>0</v></rv><rv s="1"><v>13</v><v>24</v></rv>' +
  '<rv s="0"><v>0</v><v>8</v><v>2</v><v>6</v></rv><rv s="0"><v>0</v><v>8</v><v>0</v><v>3</v></rv>' +
  '</rvData>';

test('each rich value Excel saves names the error it showed, whatever its structure', () => {
  assert.deepEqual(
    [...parseRichValueErrors(STRUCTURES, VALUES)],
    [
      [0, '#SPILL!'],
      [1, '#CALC!'],
      [2, '#FIELD!'],
      [3, '#CALC!'],
      [4, '#SPILL!'],
      [5, '#SPILL!'],
    ],
  );
});

test('the errors Excel stores only as rich values are the six it read back from their errorType', () => {
  assert.deepEqual(
    [...RICH_VALUE_ERROR_TYPES],
    [
      ['#SPILL!', 8],
      ['#CONNECT!', 9],
      ['#BLOCKED!', 10],
      ['#UNKNOWN!', 11],
      ['#FIELD!', 12],
      ['#CALC!', 13],
    ],
  );
});

function one(structure: string, value: string): unknown {
  return parseRichValueErrors(
    `<rvStructures xmlns="${NS}">${structure}</rvStructures>`,
    `<rvData xmlns="${NS}">${value}</rvData>`,
  ).get(0);
}

test('a classic error Excel reads back from a rich value reads as that error', () => {
  const structure = '<s t="_error"><k n="errorType" t="i"/></s>';
  assert.equal(one(structure, '<rv s="0"><v>3</v></rv>'), '#REF!');
  assert.equal(one(structure, '<rv s="0"><v>4</v></rv>'), '#NAME?');
  assert.equal(one(structure, '<rv s="0"><v>6</v></rv>'), '#N/A');
});

test('a rich value names no error unless an error structure gives it a known errorType', () => {
  const error = '<s t="_error"><k n="errorType" t="i"/></s>';
  // Excel 16.0 showed #CALC! for 14 but read it back as a different CVErr code, so it is no answer.
  assert.equal(one(error, '<rv s="0"><v>14</v></rv>'), undefined, 'an unconfirmed errorType');
  assert.equal(one(error, '<rv s="0"><v>eight</v></rv>'), undefined, 'an unreadable errorType');
  assert.equal(one(error, '<rv s="0"><v/></rv>'), undefined, 'an empty errorType');
  assert.equal(one(error, '<rv s="1"><v>8</v></rv>'), undefined, 'a structure the part lacks');
  assert.equal(
    one('<s t="_localImage"><k n="errorType" t="i"/></s>', '<rv s="0"><v>8</v></rv>'),
    undefined,
    'a structure that is not an error',
  );
  assert.equal(
    one('<s t="_error"><k n="subType" t="i"/></s>', '<rv s="0"><v>8</v></rv>'),
    undefined,
    'an error structure with no errorType key',
  );
});

test('an absent part names no error', () => {
  assert.equal(parseRichValueErrors('', VALUES).size, 0);
  assert.equal(parseRichValueErrors(STRUCTURES, '').size, 0);
});
