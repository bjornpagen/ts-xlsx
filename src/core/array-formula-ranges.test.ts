import assert from 'node:assert/strict';
import {test} from 'node:test';

import {decodeCellRef} from './address.ts';
import {
  type ArrayRangeConflict,
  arrayRangeConflicts,
  arrayRangeRepair,
  type FormulaPlacement,
  formulaPlacement,
} from './array-formula-ranges.ts';

function at(address: string, arrayRef?: string): FormulaPlacement {
  const {col, row} = decodeCellRef(address);
  return {address, col, row, ...(arrayRef === undefined ? {} : {arrayRef})};
}

const described = (conflicts: ArrayRangeConflict[]) =>
  conflicts.map(({array, other, kind}) => `${array.address} ${kind} ${other}`);

test('ranges holding only their own formula and sharing no cell stand', () => {
  assert.deepEqual(
    arrayRangeConflicts([at('B1', 'B1:B3'), at('C1', 'C1:D2'), at('B4', 'B4'), at('E1')]),
    [],
  );
});

test('another formula in a range, of any kind, is a conflict of the array formula around it', () => {
  // Excel 16.0 offered to repair a legacy B1:B3 holding a plain formula in B2, or a one-cell array.
  assert.deepEqual(described(arrayRangeConflicts([at('B1', 'B1:B3'), at('B2')])), [
    'B1 formula-inside B2',
  ]);
  assert.deepEqual(
    described(arrayRangeConflicts([at('B1', 'B1:C3'), at('C1', 'C1')])),
    ['B1 formula-inside C1'],
    'a cell right of the formula, on its own row',
  );
});

test('an array formula whose own cell is in an earlier range conflicts, and so does the earlier one', () => {
  // B2's own cell lies in B1:B3, so B1 holds a formula; B2's range then shares no cell with a range kept.
  assert.deepEqual(described(arrayRangeConflicts([at('B2', 'B2:B4'), at('B1', 'B1:B3')])), [
    'B1 formula-inside B2',
  ]);
});

test('ranges sharing a cell conflict even when neither holds the other formula', () => {
  // Excel 16.0 offered to repair B1:C2 beside A2:B3, which share B2 alone.
  assert.deepEqual(described(arrayRangeConflicts([at('B1', 'B1:C2'), at('A2', 'A2:B3')])), [
    'A2 ranges-overlap B1',
  ]);
});

test('what is left once every conflict is taken as a plain formula stands', () => {
  const placements = [
    at('B1', 'B1:B5'),
    at('B3', 'B3:C4'),
    at('C2', 'C2:D3'),
    at('A1', 'A1:A2'),
    at('D5'),
  ];
  const conflicts = arrayRangeConflicts(placements);
  const conflicting = new Set(conflicts.map(({array}) => array.address));
  const left = placements.map((placement) =>
    conflicting.has(placement.address) ? {...placement, arrayRef: undefined} : placement,
  );
  assert.deepEqual(arrayRangeConflicts(left), []);
});

const repaired = (placements: FormulaPlacement[]) => {
  const {cellsRemoved, formulasRemoved} = arrayRangeRepair(placements);
  return {
    cells: cellsRemoved.map(({address}) => address),
    formulas: formulasRemoved.map(({address}) => address),
  };
};

test('a repair keeps the array formula and takes the cell of any formula inside its range', () => {
  // Excel 16.0 kept B1:B3 and removed B2's cell information, B2 holding a plain or an array formula.
  assert.deepEqual(repaired([at('B1', 'B1:B3'), at('B2')]), {cells: ['B2'], formulas: []});
  assert.deepEqual(
    repaired([at('B2', 'B2:B4'), at('B1', 'B1:B3')]),
    {cells: ['B2'], formulas: []},
    'in reading order, whatever order the cells are given in',
  );
});

test('a repair takes the formula of an array whose range shares a cell with an earlier one', () => {
  // Excel 16.0 kept B1:C2 and removed A2's formula, leaving its value.
  assert.deepEqual(repaired([at('A2', 'A2:B3'), at('B1', 'B1:C2')]), {cells: [], formulas: ['A2']});
});

test('a range whose formula a repair took claims no cell', () => {
  // A3 lies in A2:B3 alone, and A2 lost its formula to B1:C2.
  assert.deepEqual(repaired([at('B1', 'B1:C2'), at('A2', 'A2:B3'), at('A3')]), {
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

test('a ref naming no range starting at its own cell fills nothing', () => {
  assert.deepEqual(arrayRangeConflicts([at('B2', 'B1:B3'), at('B2', 'junk'), at('B3')]), []);
});

test('a cell places a formula for every formula kind, and nothing for a value', () => {
  assert.deepEqual(
    formulaPlacement('B1', 2, 1, {shareType: 'array', formula: 'A1:A3', ref: 'B1:B3'}),
    {address: 'B1', col: 2, row: 1, arrayRef: 'B1:B3'},
  );
  assert.deepEqual(formulaPlacement('B1', 2, 1, {formula: 'A1'}), {address: 'B1', col: 2, row: 1});
  assert.deepEqual(formulaPlacement('B2', 2, 2, {sharedFormula: 'B1'}), {
    address: 'B2',
    col: 2,
    row: 2,
  });
  assert.deepEqual(formulaPlacement('B3', 2, 3, {shareType: 'dataTable', ref: 'B3:B4'}), {
    address: 'B3',
    col: 2,
    row: 3,
  });
  assert.equal(formulaPlacement('B4', 2, 4, 6), undefined);
  assert.equal(formulaPlacement('B5', 2, 5, {error: '#SPILL!'}), undefined);
});
