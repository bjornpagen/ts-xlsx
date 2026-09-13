import assert from 'node:assert/strict';
import {test} from 'node:test';

import {decodeCellRef} from '../../core/address.ts';
import {arrayRangeConflicts, type FormulaPlacement} from '../../core/array-formula-ranges.ts';
import {AuthoringError} from '../../errors.ts';
import {admitting, arrayRangeRepair, repairedSheetNames} from './read-repair.ts';

function at(address: string, arrayRef?: string): FormulaPlacement {
  const {col, row} = decodeCellRef(address);
  return {address, col, row, ...(arrayRef === undefined ? {} : {arrayRef})};
}

const repairOf = (placements: FormulaPlacement[]) => {
  const {cellsRemoved, formulasRemoved} = arrayRangeRepair(placements);
  return {
    cells: cellsRemoved.map(({address}) => address),
    formulas: formulasRemoved.map(({address}) => address),
  };
};

test('a repair keeps the array formula and takes the cell of any formula inside its range', () => {
  // Excel 16.0 kept B1:B3 and removed B2's cell information, B2 holding a plain or an array formula.
  assert.deepEqual(repairOf([at('B1', 'B1:B3'), at('B2')]), {cells: ['B2'], formulas: []});
  assert.deepEqual(
    repairOf([at('B2', 'B2:B4'), at('B1', 'B1:B3')]),
    {cells: ['B2'], formulas: []},
    'in reading order, whatever order the cells are given in',
  );
});

test('a repair takes the formula of an array whose range shares a cell with an earlier one', () => {
  // Excel 16.0 kept B1:C2 and removed A2's formula, leaving its value.
  assert.deepEqual(repairOf([at('A2', 'A2:B3'), at('B1', 'B1:C2')]), {cells: [], formulas: ['A2']});
});

test('a range whose formula a repair took claims no cell', () => {
  // A3 lies in A2:B3 alone, and A2 lost its formula to B1:C2.
  assert.deepEqual(repairOf([at('B1', 'B1:C2'), at('A2', 'A2:B3'), at('A3')]), {
    cells: [],
    formulas: ['A2'],
  });
});

test('what a repair leaves stands', () => {
  const placements = [
    at('B1', 'B1:B5'),
    at('B3', 'B3:C4'),
    at('C2', 'C2:D3'),
    at('A1', 'A1:A2'),
    at('A2'),
    at('D5'),
  ];
  const {cellsRemoved, formulasRemoved} = arrayRangeRepair(placements);
  const taken = new Set([...cellsRemoved, ...formulasRemoved].map(({address}) => address));
  assert.deepEqual(taken, new Set(['A2', 'B3']));
  assert.deepEqual(arrayRangeConflicts(placements.filter(({address}) => !taken.has(address))), []);
});

test('a workbook of sheet names is repaired in order, each against the names before it', () => {
  const repaired = repairedSheetNames([
    {name: 'S', part: 1},
    {name: 's', part: 2},
    {name: 'x'.repeat(40), part: 3},
    {name: '', part: 4},
  ]);
  assert.deepEqual(repaired, [
    {name: 'S', part: 1},
    {name: 's (2)', part: 2},
    {name: 'x'.repeat(31), part: 3},
    {name: 'Sheet1', part: 4},
  ]);
});

test('admitting swallows a refusal about the input and nothing else', () => {
  assert.equal(
    admitting((): number => {
      throw new AuthoringError('refused');
    }),
    undefined,
  );
  assert.equal(
    admitting(() => 7),
    7,
  );
  assert.throws(
    () =>
      admitting(() => {
        throw new TypeError('a bug');
      }),
    TypeError,
  );
});
