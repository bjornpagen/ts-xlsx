import assert from 'node:assert/strict';
import {test} from 'node:test';

import type {ErrorCode} from '../../core/value.ts';
import {indexCellMetadata, NO_METADATA, type WorkbookMetadata} from './metadata.ts';

// A part declaring the rich-value type ahead of the dynamic-array one, with a block of each.
const METADATA: WorkbookMetadata = {
  typeNames: ['XLRICHVALUE', 'XLDAPR'],
  dynamicArrayBlocks: [true],
  richValueBlocks: [1, 0],
  cellBlocks: [[{type: 2, value: 0}], [{type: 1, value: 0}]],
  valueBlocks: [[{type: 1, value: 0}], [{type: 1, value: 1}], [{type: 2, value: 0}]],
};
const ERRORS: ReadonlyMap<number, ErrorCode> = new Map([
  [0, '#CALC!'],
  [1, '#SPILL!'],
]);

test('a cm marks a dynamic array only through a dynamic-array block with fDynamic set', () => {
  assert.deepEqual([...indexCellMetadata(METADATA, ERRORS).dynamicArrayCells], [1]);
  const off = {...METADATA, dynamicArrayBlocks: [false]};
  assert.deepEqual([...indexCellMetadata(off, ERRORS).dynamicArrayCells], []);
});

test('a vm names the error of the rich value its block names, through the rich-value type only', () => {
  assert.deepEqual(
    [...indexCellMetadata(METADATA, ERRORS).valueErrors],
    [
      [1, '#SPILL!'],
      [2, '#CALC!'],
    ],
  );
});

test('a vm resolving to no block, no rich value or no error names nothing', () => {
  const metadata: WorkbookMetadata = {
    ...METADATA,
    richValueBlocks: [-1, 7, 1],
    valueBlocks: [
      [{type: 1, value: 0}],
      [{type: 1, value: 1}],
      [{type: 1, value: 9}],
      [{type: 3, value: 2}],
    ],
  };
  assert.deepEqual([...indexCellMetadata(metadata, ERRORS).valueErrors], []);
});

test('no metadata part resolves nothing', () => {
  const index = indexCellMetadata(NO_METADATA, ERRORS);
  assert.equal(index.dynamicArrayCells.size + index.valueErrors.size, 0);
});
