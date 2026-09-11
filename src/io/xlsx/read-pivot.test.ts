import assert from 'node:assert/strict';
import {test} from 'node:test';

import {PivotTable, type PivotTableOptions} from '../../core/pivot-table.ts';
import {Workbook} from '../../core/workbook.ts';
import {pivotCacheDefinitionXml, pivotTableXml} from './pivot.ts';
import {parsePivotTable} from './read-pivot.ts';

// Build a pivot the way the writer does, then render its two definition parts. The reader is the
// inverse of the writer, so round-tripping our own output is the sharpest check that the two agree.
function renderedPivot(
  options: Omit<PivotTableOptions, 'source'>,
  data: readonly (readonly (string | number)[])[],
  name = 'PivotTable1',
  cacheId = '1',
): {table: string; cache: string} {
  const wb = new Workbook();
  const src = wb.addWorksheet('Data');
  data.forEach((row, r) => {
    row.forEach((value, c) => {
      src.getCell(encode(c + 1, r + 1)).value = value;
    });
  });
  const pivot = new PivotTable({source: src, ...options});
  return {
    table: pivotTableXml(pivot, name, cacheId),
    cache: pivotCacheDefinitionXml(pivot, 'rId1'),
  };
}

function encode(col: number, row: number): string {
  let letters = '';
  for (let n = col; n > 0; n = Math.floor((n - 1) / 26)) {
    letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  }
  return `${letters}${row}`;
}

const SALES: readonly (readonly (string | number)[])[] = [
  ['Name', 'Region', 'Amount'],
  ['Alice', 'West', 10],
  ['Bob', 'East', 20],
  ['Cara', 'West', 30],
];

test('a written pivot round-trips its field roles, source, and value field back into a model', () => {
  const {table, cache} = renderedPivot(
    {rows: ['Name'], columns: ['Region'], values: ['Amount']},
    SALES,
  );
  const parsed = parsePivotTable(table, cache);

  assert.equal(parsed.name, 'PivotTable1');
  assert.equal(parsed.cacheId, '1');
  assert.deepEqual(
    parsed.fields.map((field) => field.name),
    ['Name', 'Region', 'Amount'],
  );
  assert.deepEqual(parsed.rowFields, [0]);
  assert.deepEqual(parsed.columnFields, [1]);
  assert.equal(parsed.valueField, 2);
  assert.equal(parsed.valueFieldName, 'Amount');
});

test('the worksheet source reference and sheet name survive the round-trip', () => {
  const {table, cache} = renderedPivot(
    {rows: ['Name'], columns: ['Region'], values: ['Amount']},
    SALES,
  );
  const parsed = parsePivotTable(table, cache);
  assert.equal(parsed.source.kind, 'worksheet');
  assert.equal(parsed.source.sheet, 'Data');
  assert.equal(parsed.source.ref, 'A1:C4');
  assert.equal(parsed.source.inAnotherWorkbook, false);
});

test('a worksheet source reached through a relationship is a range in another workbook', () => {
  const cache =
    '<pivotCacheDefinition xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<cacheSource type="worksheet"><worksheetSource ref="A1:B2" sheet="Data" q:id="rId1"/></cacheSource>' +
    '<cacheFields count="1"><cacheField name="Amount"/></cacheFields></pivotCacheDefinition>';
  const parsed = parsePivotTable('<pivotTableDefinition name="P" cacheId="1"/>', cache);
  assert.deepEqual(parsed.source, {
    kind: 'worksheet',
    sheet: 'Data',
    ref: 'A1:B2',
    inAnotherWorkbook: true,
  });
});

test('sum, the metric whose subtotal attribute is omitted, reads back as sum', () => {
  const {table, cache} = renderedPivot(
    {rows: ['Name'], columns: ['Region'], values: ['Amount']},
    SALES,
  );
  const parsed = parsePivotTable(table, cache);
  assert.equal(parsed.metric, 'sum');
  assert.equal(parsed.valueCaption, 'Sum of Amount');
});

test('a non-sum metric reads back from its subtotal attribute with its caption', () => {
  const {table, cache} = renderedPivot(
    {rows: ['Name'], columns: ['Region'], values: ['Amount'], metric: 'average'},
    SALES,
  );
  const parsed = parsePivotTable(table, cache);
  assert.equal(parsed.metric, 'average');
  assert.equal(parsed.valueCaption, 'Average of Amount');
});

test('field names carrying XML specials decode back to their original text', () => {
  const data: readonly (readonly (string | number)[])[] = [
    ['Smith & Co', '<Region>', 'Am"t'],
    ['a', 'x', 1],
    ['b', 'y', 2],
  ];
  const {table, cache} = renderedPivot(
    {rows: ['Smith & Co'], columns: ['<Region>'], values: ['Am"t']},
    data,
  );
  const parsed = parsePivotTable(table, cache);
  assert.deepEqual(
    parsed.fields.map((field) => field.name),
    ['Smith & Co', '<Region>', 'Am"t'],
  );
  assert.equal(parsed.valueFieldName, 'Am"t');
  assert.equal(parsed.valueCaption, 'Sum of Am"t');
});

test('multiple row fields keep their order and axis', () => {
  const data: readonly (readonly (string | number)[])[] = [
    ['Year', 'Name', 'Region', 'Amount'],
    [2024, 'a', 'x', 1],
    [2025, 'b', 'y', 2],
  ];
  const {table, cache} = renderedPivot(
    {rows: ['Year', 'Name'], columns: ['Region'], values: ['Amount']},
    data,
  );
  const parsed = parsePivotTable(table, cache);
  assert.deepEqual(parsed.rowFields, [0, 1]);
  assert.deepEqual(parsed.columnFields, [2]);
});

// The read path is lenient: a file that already exists is reconstructed as best it can be, never
// rejected. Each case below is malformed in a way a hostile or foreign producer might emit.

test('an unrecognised subtotal is read leniently as sum, not rejected', () => {
  const table =
    `<pivotTableDefinition name="P" cacheId="1">` +
    `<rowFields count="1"><field x="0"/></rowFields>` +
    `<colFields count="1"><field x="1"/></colFields>` +
    `<dataFields count="1"><dataField name="Bogus of Amount" fld="2" subtotal="bogus"/></dataFields>` +
    `</pivotTableDefinition>`;
  const cache =
    `<pivotCacheDefinition><cacheSource type="worksheet"><worksheetSource ref="A1:C4" sheet="Data"/></cacheSource>` +
    `<cacheFields count="3"><cacheField name="Name"/><cacheField name="Region"/><cacheField name="Amount"/></cacheFields>` +
    `</pivotCacheDefinition>`;
  const parsed = parsePivotTable(table, cache);
  assert.equal(parsed.metric, 'sum');
  assert.equal(parsed.valueField, 2);
});

test('a cache with no cacheSource at all defaults to a worksheet source with empty coordinates', () => {
  const table = `<pivotTableDefinition name="P" cacheId="1"><dataFields count="1"><dataField fld="0"/></dataFields></pivotTableDefinition>`;
  const cache = `<pivotCacheDefinition><cacheFields count="1"><cacheField name="Amount"/></cacheFields></pivotCacheDefinition>`;
  const parsed = parsePivotTable(table, cache);
  assert.deepEqual(parsed.source, {
    kind: 'worksheet',
    sheet: '',
    ref: '',
    inAnotherWorkbook: false,
  });
  assert.equal(parsed.valueFieldName, 'Amount');
});

test('an external cache source reports its kind and carries no worksheet coordinates', () => {
  const table = `<pivotTableDefinition name="P" cacheId="1"><dataFields count="1"><dataField fld="0"/></dataFields></pivotTableDefinition>`;
  const cache =
    `<pivotCacheDefinition><cacheSource type="external"><connection id="1"/></cacheSource>` +
    `<cacheFields count="1"><cacheField name="Amount"/></cacheFields></pivotCacheDefinition>`;
  const parsed = parsePivotTable(table, cache);
  assert.equal(parsed.source.kind, 'external');
  assert.equal(parsed.source.sheet, '');
  assert.equal(parsed.source.ref, '');
});

test('a consolidation cache source is recognised by kind', () => {
  const cache =
    `<pivotCacheDefinition><cacheSource type="consolidation"><rangeSets/></cacheSource>` +
    `<cacheFields count="1"><cacheField name="Amount"/></cacheFields></pivotCacheDefinition>`;
  const parsed = parsePivotTable(`<pivotTableDefinition name="P" cacheId="1"/>`, cache);
  assert.equal(parsed.source.kind, 'consolidation');
});

test('an unrecognised cache source type degrades to unknown, not a throw', () => {
  const cache =
    `<pivotCacheDefinition><cacheSource type="wormhole"><worksheetSource ref="A1:B2" sheet="Data"/></cacheSource>` +
    `<cacheFields count="1"><cacheField name="Amount"/></cacheFields></pivotCacheDefinition>`;
  const parsed = parsePivotTable(`<pivotTableDefinition name="P" cacheId="1"/>`, cache);
  // The declared type is not one we model, so the kind reports `unknown`, but a `<worksheetSource>`
  // that rides along is still read, since a foreign producer may pair either with the other.
  assert.equal(parsed.source.kind, 'unknown');
  assert.equal(parsed.source.sheet, 'Data');
  assert.equal(parsed.source.ref, 'A1:B2');
});

test('a cacheSource with no type attribute defaults to worksheet', () => {
  const cache =
    `<pivotCacheDefinition><cacheSource><worksheetSource ref="A1:B2" sheet="Data"/></cacheSource>` +
    `<cacheFields count="1"><cacheField name="Amount"/></cacheFields></pivotCacheDefinition>`;
  const parsed = parsePivotTable(`<pivotTableDefinition name="P" cacheId="1"/>`, cache);
  assert.equal(parsed.source.kind, 'worksheet');
  assert.equal(parsed.source.sheet, 'Data');
});

test('a table with no dataField reports no value field instead of a wild index', () => {
  const table = `<pivotTableDefinition name="P" cacheId="1"><rowFields count="1"><field x="0"/></rowFields></pivotTableDefinition>`;
  const cache = `<pivotCacheDefinition><cacheFields count="1"><cacheField name="Name"/></cacheFields></pivotCacheDefinition>`;
  const parsed = parsePivotTable(table, cache);
  assert.equal(parsed.valueField, -1);
  assert.equal(parsed.valueFieldName, '');
});

test('a non-integer field index is dropped rather than admitted as NaN', () => {
  const table =
    `<pivotTableDefinition name="P" cacheId="1">` +
    `<rowFields count="2"><field x="0"/><field x="../etc"/></rowFields>` +
    `<dataFields count="1"><dataField fld="1"/></dataFields>` +
    `</pivotTableDefinition>`;
  const cache = `<pivotCacheDefinition><cacheFields count="2"><cacheField name="Name"/><cacheField name="Amount"/></cacheFields></pivotCacheDefinition>`;
  const parsed = parsePivotTable(table, cache);
  assert.deepEqual(parsed.rowFields, [0]);
});

test('only the first dataField is modeled; a stray second one is ignored', () => {
  const table =
    `<pivotTableDefinition name="P" cacheId="1">` +
    `<dataFields count="2">` +
    `<dataField name="Sum of Amount" fld="1"/>` +
    `<dataField name="Count of Name" fld="0" subtotal="count"/>` +
    `</dataFields>` +
    `</pivotTableDefinition>`;
  const cache = `<pivotCacheDefinition><cacheFields count="2"><cacheField name="Name"/><cacheField name="Amount"/></cacheFields></pivotCacheDefinition>`;
  const parsed = parsePivotTable(table, cache);
  assert.equal(parsed.valueField, 1);
  assert.equal(parsed.metric, 'sum');
});
