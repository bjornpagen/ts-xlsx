import assert from 'node:assert/strict';
import {test} from 'node:test';

import type {CellValue} from './value.ts';
import {Workbook} from './workbook.ts';
import type {Worksheet} from './worksheet.ts';

// A Ctrl+Shift+Enter formula over `ref`, anchored at its first cell, beside a column of numbers. Excel 16.0
// refused every edit below that cuts through such a range, and allowed every one that moves or removes it
// whole (`test/corpus/fixtures/excel-oracle/array-formula-ranges.json`).
function sheetWithArray(ref: string, dynamic = false): Worksheet {
  const sheet = new Workbook().addWorksheet('S');
  for (let row = 1; row <= 6; row++) sheet.getCell(`A${row}`).value = row;
  const anchor = ref.split(':')[0] ?? ref;
  sheet.getCell(anchor).value = {
    shareType: 'array',
    formula: 'A1:A3*2',
    ref,
    ...(dynamic ? {dynamic: true} : {}),
  };
  return sheet;
}

const refAt = (sheet: Worksheet, address: string): unknown =>
  (sheet.getCell(address).value as {ref?: string} | null)?.ref ?? null;

function refuses(edit: (sheet: Worksheet) => void, ref = 'B1:B3'): void {
  const sheet = sheetWithArray(ref);
  const before = sheet.model;
  assert.throws(
    () => edit(sheet),
    (error: Error) => {
      assert.equal(error.name, 'AuthoringError');
      assert.match(error.message, new RegExp(`array formula in B1, which fills ${ref}`));
      return true;
    },
  );
  assert.deepEqual(sheet.model, before, 'the sheet is left untouched');
}

test('a row edit cutting through a Ctrl+Shift+Enter range is refused', () => {
  refuses((sheet) => sheet.insertRow(2, []));
  refuses((sheet) => sheet.spliceRows(3, 0, [], []));
  refuses((sheet) => sheet.spliceRows(2, 1));
  refuses((sheet) => sheet.spliceRows(3, 2), 'B1:B3');
  refuses((sheet) => sheet.duplicateRow(1));
  refuses((sheet) => sheet.duplicateRow(1, {insert: false}));
});

test('a row edit moving or removing a Ctrl+Shift+Enter range whole goes ahead', () => {
  const above = sheetWithArray('B1:B3');
  above.insertRow(1, []);
  assert.equal(refAt(above, 'B2'), 'B2:B4', 'an insert at its first row moves it down');

  const below = sheetWithArray('B1:B3');
  below.insertRow(4, []);
  assert.equal(refAt(below, 'B1'), 'B1:B3', 'an insert after its last row leaves it');

  const whole = sheetWithArray('B1:B3');
  whole.spliceRows(1, 4);
  assert.equal(whole.getCell('B1').value, null, 'a delete taking every row removes it');

  const copiedBelow = sheetWithArray('B1:B3');
  copiedBelow.duplicateRow(3);
  assert.equal(refAt(copiedBelow, 'B1'), 'B1:B3', 'a copy landing below leaves it');
});

test('a column edit cutting through a Ctrl+Shift+Enter range is refused, and one moving it is not', () => {
  refuses((sheet) => sheet.insertColumn(3, []), 'B1:C3');
  refuses((sheet) => sheet.spliceColumns(3, 1), 'B1:C3');

  const moved = sheetWithArray('B1:C3');
  moved.insertColumn(2, []);
  assert.equal(refAt(moved, 'C1'), 'C1:D3');

  const removed = sheetWithArray('B1:C3');
  removed.spliceColumns(2, 2);
  assert.equal(removed.getCell('B1').value, null);
});

test('an edit through a dynamic array goes ahead, as Excel lets it', () => {
  const sheet = sheetWithArray('B1:B3', true);
  sheet.insertRow(2, []);
  assert.deepEqual(sheet.getCell('B1').value, {
    shareType: 'array',
    formula: 'A1:A4*2',
    ref: 'B1:B4',
    dynamic: true,
  } satisfies CellValue);
});
