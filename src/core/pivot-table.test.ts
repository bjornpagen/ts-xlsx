import assert from 'node:assert/strict';
import {test} from 'node:test';

import {AuthoringError} from '../errors.ts';
import {PivotTable, type PivotTableOptions} from './pivot-table.ts';
import {Workbook} from './workbook.ts';
import type {Worksheet} from './worksheet.ts';

function sourceSheet(headers: readonly string[] = ['Name', 'Region', 'Amount']): Worksheet {
  const src = new Workbook().addWorksheet('Data');
  headers.forEach((header, col) => {
    src.getCell(`${String.fromCharCode(65 + col)}1`).value = header;
    src.getCell(`${String.fromCharCode(65 + col)}2`).value = col;
  });
  return src;
}

function pivot(overrides: Partial<PivotTableOptions>): PivotTable {
  return new PivotTable({
    source: sourceSheet(),
    rows: ['Name'],
    columns: ['Region'],
    values: ['Amount'],
    ...overrides,
  });
}

function refusal(message: RegExp): (error: unknown) => boolean {
  return (error) => error instanceof AuthoringError && message.test(error.message);
}

test('authoring rejects unsupported shapes at add time', () => {
  assert.throws(() => pivot({metric: 'avg' as never}), refusal(/unsupported pivot metric "avg"/));
  assert.throws(() => pivot({rows: ['Nope']}), refusal(/"Nope" is not a column header/));
  assert.throws(() => pivot({values: ['Amount', 'Name']}), refusal(/exactly one value field/));
  assert.throws(() => pivot({rows: []}), refusal(/at least one row field/));
  assert.throws(() => pivot({columns: []}), refusal(/at least one column field/));
});

test('a field named as both a row and a column field is refused', () => {
  assert.throws(
    () => pivot({rows: ['Name'], columns: ['Name']}),
    refusal(/pivot field "Name" is named as both a row and a column field/),
  );
  assert.throws(
    () => pivot({rows: ['Name', 'Region'], columns: ['Amount', 'Region']}),
    refusal(/"Region" is named as both a row and a column field/),
  );
});

test('a field named twice in one axis is refused', () => {
  assert.throws(
    () => pivot({rows: ['Name', 'Name']}),
    refusal(/pivot row field "Name" is named more than once/),
  );
  assert.throws(
    () => pivot({columns: ['Region', 'Region']}),
    refusal(/pivot column field "Region" is named more than once/),
  );
});

test('a source header row that repeats a name, in any case, is refused', () => {
  for (const repeat of ['Name', 'name', 'NAME']) {
    assert.throws(
      () => pivot({source: sourceSheet(['Name', 'Region', 'Amount', repeat])}),
      refusal(new RegExp(`header "${repeat}" repeats "Name"`)),
    );
  }
});

test('a field name binds its header in any case, and naming one field in two cases is a repeat', () => {
  const folded = pivot({rows: ['name'], columns: ['REGION'], values: ['aMOUNT']});
  assert.deepEqual([folded.rowFields, folded.columnFields, folded.valueField], [[0], [1], 2]);
  assert.throws(
    () => pivot({rows: ['Name', 'NAME']}),
    refusal(/pivot row field "NAME" is named more than once/),
  );
});

test('valid role sets construct and bind each field where it was named', () => {
  const plain = pivot({});
  assert.deepEqual(plain.rowFields, [0]);
  assert.deepEqual(plain.columnFields, [1]);
  assert.equal(plain.valueField, 2);

  const nested = pivot({
    source: sourceSheet(['Name', 'Region', 'Amount', 'Quarter']),
    rows: ['Name', 'Quarter'],
    columns: ['Region'],
  });
  assert.deepEqual(nested.rowFields, [0, 3]);
  assert.deepEqual(nested.columnFields, [1]);

  // Aggregating a field that is also an axis is an ordinary pivot, not a field used twice.
  const countByName = pivot({values: ['Name'], metric: 'count'});
  assert.deepEqual(countByName.rowFields, [0]);
  assert.equal(countByName.valueField, 0);
});
