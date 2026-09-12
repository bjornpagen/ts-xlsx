import assert from 'node:assert/strict';
import {test} from 'node:test';

import {NO_METADATA} from '../cell-metadata/metadata.ts';
import {parseMetadataPart} from './read-metadata.ts';
import {BRT} from './record-types.ts';
import {concat, frame, reframe, wide, word} from './records.test-support.ts';

const hex = (text: string) => Uint8Array.from(Buffer.from(text, 'hex'));
const empty = new Uint8Array(0);

// The part as Excel 16.0 (build 20326) saved it for a sheet holding dynamic-array formulas and six
// rich-value errors. Records the model does not read (the part's and the lists' own begin and end, the
// future-record framing, the empty record opening each block's properties) are kept, as the reader has
// to walk past them.
const RICH_VALUE_BLOCK = (index: number) =>
  concat(
    frame(BRT.BeginFmd, empty),
    frame(35, hex('01000200')),
    frame(5002, empty),
    frame(BRT.RichValueBlock, concat(word(0), word(index))),
    frame(36, empty),
    frame(53, empty),
  );
const EXCEL_SAVED = concat(
  frame(332, empty),
  frame(334, word(2)),
  frame(BRT.Mdtinfo, concat(hex('b0c06ad8c0d40100'), wide('XLDAPR'))),
  frame(BRT.Mdtinfo, concat(hex('b0c06a98c0d40100'), wide('XLRICHVALUE'))),
  frame(336, empty),
  frame(BRT.BeginEsfmd, concat(word(1), wide('XLDAPR'))),
  frame(BRT.BeginFmd, empty),
  frame(35, hex('02000200')),
  frame(4096, empty),
  frame(BRT.DynamicArrayProperties, hex('0100')),
  frame(36, empty),
  frame(53, empty),
  frame(BRT.EndEsfmd, empty),
  frame(BRT.BeginEsfmd, concat(word(6), wide('XLRICHVALUE'))),
  ...[0, 1, 2, 3, 4, 5].map(RICH_VALUE_BLOCK),
  frame(BRT.EndEsfmd, empty),
  frame(BRT.BeginEsmdb, concat(word(1), word(1))),
  frame(BRT.Mdb, concat(word(1), word(1), word(0))),
  frame(BRT.EndEsmdb, empty),
  frame(BRT.BeginEsmdb, concat(word(6), word(0))),
  ...[0, 1, 2, 3, 4, 5].map((index) => frame(BRT.Mdb, concat(word(1), word(2), word(index)))),
  frame(BRT.EndEsmdb, empty),
  frame(333, empty),
);

test('the metadata part Excel saves reads as the XML form of the same part does', () => {
  assert.deepEqual(parseMetadataPart(EXCEL_SAVED), {
    typeNames: ['XLDAPR', 'XLRICHVALUE'],
    dynamicArrayBlocks: [true],
    richValueBlocks: [0, 1, 2, 3, 4, 5],
    cellBlocks: [[{type: 1, value: 0}]],
    valueBlocks: [0, 1, 2, 3, 4, 5].map((index) => [{type: 2, value: index}]),
  });
});

test('a dynamic-array block with fDynamic clear marks nothing, whatever else its flags say', () => {
  // Excel 16.0 read every formula the block marked as a legacy array formula with bit 0 clear, and as a
  // spilling one with bit 1 set beside bit 0.
  const withFlags = (flags: string) =>
    reframe(EXCEL_SAVED, (type, data) =>
      frame(type, type === BRT.DynamicArrayProperties ? hex(flags) : data),
    );
  assert.deepEqual(parseMetadataPart(withFlags('0000')).dynamicArrayBlocks, [false]);
  assert.deepEqual(parseMetadataPart(withFlags('0300')).dynamicArrayBlocks, [true]);
});

test('a metadata block reads the records its payload holds, not the count it claims', () => {
  const part = concat(
    frame(BRT.BeginEsmdb, concat(word(1), word(0))),
    frame(BRT.Mdb, concat(word(1_000_000), word(2), word(7))),
    frame(BRT.EndEsmdb, empty),
  );
  assert.deepEqual(parseMetadataPart(part).valueBlocks, [[{type: 2, value: 7}]]);
});

test('a block record outside the list of its type is not read as one', () => {
  const part = concat(
    frame(BRT.BeginEsfmd, concat(word(1), wide('XLSOMETHINGELSE'))),
    frame(BRT.BeginFmd, empty),
    frame(BRT.DynamicArrayProperties, hex('0100')),
    frame(BRT.RichValueBlock, concat(word(0), word(3))),
    frame(BRT.EndEsfmd, empty),
    frame(BRT.Mdb, concat(word(1), word(1), word(0))),
  );
  assert.deepEqual(parseMetadataPart(part), {
    typeNames: [],
    dynamicArrayBlocks: [],
    richValueBlocks: [],
    cellBlocks: [],
    valueBlocks: [],
  });
});

test('a workbook with no metadata part has no metadata', () => {
  assert.equal(parseMetadataPart(undefined), NO_METADATA);
});
