import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Workbook} from '../../core/workbook.ts';
import {partIn, partsWritten, roundtrip} from './package.test-support.ts';
import {writeXlsx} from './write.ts';

// A source sheet whose data carries every XML-special character plus a missing value in an axis
// column: the shape that corrupts a naive pivot writer that fails to entity-escape shared items.
function specialCharsWorkbook(): Workbook {
  const wb = new Workbook();
  const src = wb.addWorksheet('Data');
  src.getCell('A1').value = 'Name';
  src.getCell('B1').value = 'Region';
  src.getCell('C1').value = 'Amount';
  src.getCell('A2').value = 'Smith & Co';
  src.getCell('B2').value = '<West>';
  src.getCell('C2').value = 10;
  // A3 (Name) left empty: a missing axis value.
  src.getCell('B3').value = 'East';
  src.getCell('C3').value = 20;
  src.getCell('A4').value = 'It\'s "best"';
  src.getCell('B4').value = 'West';
  src.getCell('C4').value = 30;
  const dst = wb.addWorksheet('Pivot');
  dst.addPivotTable({source: src, rows: ['Name'], columns: ['Region'], values: ['Amount']});
  return wb;
}

// A raw `&` that is not the start of a valid entity: the escaping bug this feature exists to prevent.
const RAW_AMP = /&(?!amp;|lt;|gt;|quot;|apos;|#\d+;|#x[0-9a-fA-F]+;)/;

test('a pivot over source data with XML-special characters and a null value writes without throwing', () => {
  assert.doesNotThrow(() => writeXlsx(specialCharsWorkbook()));
});

test('the pivot cache serialises special characters entity-escaped into well-formed XML', () => {
  const cache = partIn(
    partsWritten(specialCharsWorkbook()),
    'xl/pivotCache/pivotCacheDefinition1.xml',
  );
  assert.doesNotMatch(cache, RAW_AMP, 'no raw unescaped "&" may leak into the cache');
  assert.match(cache, /<s v="Smith &amp; Co"\/>/);
  assert.match(cache, /<s v="&lt;West&gt;"\/>/);
  assert.match(cache, /<s v="It&apos;s &quot;best&quot;"\/>/);
});

test('a missing axis value becomes a blank shared item, not an empty string', () => {
  const cache = partIn(
    partsWritten(specialCharsWorkbook()),
    'xl/pivotCache/pivotCacheDefinition1.xml',
  );
  // The Name field carries a blank shared item and is flagged as containing one.
  assert.match(cache, /name="Name"[^>]*>\s*<sharedItems containsBlank="1"[^>]*>.*<m\/>/s);
  assert.doesNotMatch(cache, /<s v=""\/>/, 'a missing value must be <m/>, never an empty <s>');
});

test('a numeric value field is described as numeric and stored inline in the records', () => {
  const parts = partsWritten(specialCharsWorkbook());
  const cache = partIn(parts, 'xl/pivotCache/pivotCacheDefinition1.xml');
  assert.match(
    cache,
    /name="Amount"[^>]*>\s*<sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" containsInteger="1" minValue="10" maxValue="30"\/>/,
  );
  const records = partIn(parts, 'xl/pivotCache/pivotCacheRecords1.xml');
  // Each record: an index into each axis field's catalogue, then the inline numeric amount.
  assert.match(records, /<r><x v="0"\/><x v="0"\/><n v="10"\/><\/r>/);
  assert.match(records, /<r><x v="2"\/><x v="2"\/><n v="30"\/><\/r>/);
});

test('an inline string field escapes its values in the records', () => {
  const wb = new Workbook();
  const src = wb.addWorksheet('Data');
  src.getCell('A1').value = 'Name';
  src.getCell('B1').value = 'Region';
  src.getCell('C1').value = 'Amount';
  src.getCell('D1').value = 'Note';
  src.getCell('A2').value = 'a';
  src.getCell('B2').value = 'x';
  src.getCell('C2').value = 1;
  src.getCell('D2').value = 'see <this> & that';
  wb.addWorksheet('P').addPivotTable({
    source: src,
    rows: ['Name'],
    columns: ['Region'],
    values: ['Amount'],
  });

  const parts = partsWritten(wb);
  const records = partIn(parts, 'xl/pivotCache/pivotCacheRecords1.xml');
  // Note is neither an axis nor the value field, so it rides inline as an escaped <s>.
  assert.match(records, /<s v="see &lt;this&gt; &amp; that"\/>/);
  assert.doesNotMatch(records, RAW_AMP);
});

test('the pivot is wired end to end: content types, workbook cache, sheet link, and the rel chain', () => {
  const parts = partsWritten(specialCharsWorkbook());

  // Content types declare all three generated parts.
  const types = partIn(parts, '[Content_Types].xml');
  assert.match(types, /PartName="\/xl\/pivotTables\/pivotTable1\.xml"/);
  assert.match(types, /PartName="\/xl\/pivotCache\/pivotCacheDefinition1\.xml"/);
  assert.match(types, /PartName="\/xl\/pivotCache\/pivotCacheRecords1\.xml"/);

  // The workbook registers the cache and relates to its definition by the same id.
  const workbook = partIn(parts, 'xl/workbook.xml');
  const cacheRelId = workbook.match(/<pivotCache cacheId="1" r:id="(rId\d+)"\/>/)?.[1];
  assert.ok(cacheRelId, 'the cache must be registered in <pivotCaches>');
  const workbookRels = partIn(parts, 'xl/_rels/workbook.xml.rels');
  assert.match(
    workbookRels,
    new RegExp(
      `Id="${cacheRelId}"[^>]*\\/pivotCacheDefinition"[^>]*Target="pivotCache\\/pivotCacheDefinition1\\.xml"`,
    ),
  );

  // The host sheet reaches the pivot table part (no reference in the sheet body itself).
  const sheetRels = partIn(parts, 'xl/worksheets/_rels/sheet2.xml.rels');
  assert.match(sheetRels, /\/pivotTable"[^>]*Target="\.\.\/pivotTables\/pivotTable1\.xml"/);

  // The chain: pivot table → cache definition → cache records.
  assert.match(
    partIn(parts, 'xl/pivotTables/_rels/pivotTable1.xml.rels'),
    /\/pivotCacheDefinition"[^>]*Target="\.\.\/pivotCache\/pivotCacheDefinition1\.xml"/,
  );
  assert.match(
    partIn(parts, 'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels'),
    /\/pivotCacheRecords"[^>]*Target="pivotCacheRecords1\.xml"/,
  );
});

test('the pivot table binds the cache and sums the value field', () => {
  const table = partIn(partsWritten(specialCharsWorkbook()), 'xl/pivotTables/pivotTable1.xml');
  assert.match(table, /cacheId="1"/);
  assert.match(table, /<pivotField axis="axisRow"/);
  assert.match(table, /<pivotField axis="axisCol"/);
  assert.match(table, /<pivotField dataField="1"/);
  assert.match(table, /<dataField name="Sum of Amount" fld="2"/);
});

test('a non-sum metric carries its subtotal function and an Excel-style caption', () => {
  const wb = new Workbook();
  const src = wb.addWorksheet('Data');
  src.getCell('A1').value = 'Name';
  src.getCell('B1').value = 'Region';
  src.getCell('C1').value = 'Amount';
  src.getCell('A2').value = 'a';
  src.getCell('B2').value = 'x';
  src.getCell('C2').value = 5;
  wb.addWorksheet('P').addPivotTable({
    source: src,
    rows: ['Name'],
    columns: ['Region'],
    values: ['Amount'],
    metric: 'average',
  });

  const table = partIn(partsWritten(wb), 'xl/pivotTables/pivotTable1.xml');
  assert.match(table, /<dataField name="Average of Amount" fld="2" subtotal="average"/);
});

test("sum omits the subtotal attribute: it is Excel's implicit default", () => {
  const table = partIn(partsWritten(specialCharsWorkbook()), 'xl/pivotTables/pivotTable1.xml');
  assert.doesNotMatch(
    table,
    /subtotal=/,
    'the default sum aggregation writes no subtotal attribute',
  );
});

test('a count aggregates a non-numeric value field, describing it as a plain shared-items field', () => {
  const wb = new Workbook();
  const src = wb.addWorksheet('Data');
  src.getCell('A1').value = 'Name';
  src.getCell('B1').value = 'Region';
  src.getCell('C1').value = 'Status';
  src.getCell('A2').value = 'a';
  src.getCell('B2').value = 'x';
  src.getCell('C2').value = 'open';
  wb.addWorksheet('P').addPivotTable({
    source: src,
    rows: ['Name'],
    columns: ['Region'],
    values: ['Status'],
    metric: 'count',
  });

  const parts = partsWritten(wb);
  const table = partIn(parts, 'xl/pivotTables/pivotTable1.xml');
  assert.match(table, /<dataField name="Count of Status" fld="2" subtotal="count"/);
  // A text value field is not summarised as numeric: it carries a bare <sharedItems/> and rides
  // inline in the records, where the count aggregation tallies its non-blank cells.
  const cache = partIn(parts, 'xl/pivotCache/pivotCacheDefinition1.xml');
  assert.match(cache, /name="Status" numFmtId="0"><sharedItems\/>/);
  assert.match(partIn(parts, 'xl/pivotCache/pivotCacheRecords1.xml'), /<s v="open"\/>/);
});

test('a package carrying a pivot still reads back its sheets', () => {
  const back = roundtrip(specialCharsWorkbook());
  assert.deepEqual(
    back.worksheets.map((sheet) => sheet.name),
    ['Data', 'Pivot'],
  );
  assert.equal(back.getWorksheet('Data')?.getCell('A2').value, 'Smith & Co');
});

test('a loaded pivot is reconstructed as an inspectable model on its host sheet', () => {
  const back = roundtrip(specialCharsWorkbook());
  const loaded = back.getWorksheet('Pivot')?.loadedPivotTables ?? [];
  assert.equal(loaded.length, 1);
  const pivot = loaded[0];
  assert.ok(pivot);
  assert.deepEqual(
    pivot.fields.map((field) => field.name),
    ['Name', 'Region', 'Amount'],
  );
  assert.deepEqual(pivot.rowFields, [0]);
  assert.deepEqual(pivot.columnFields, [1]);
  assert.equal(pivot.valueField, 2);
  assert.equal(pivot.valueFieldName, 'Amount');
  assert.equal(pivot.metric, 'sum');
  assert.equal(pivot.source.kind, 'worksheet');
  assert.equal(pivot.source.sheet, 'Data');
  assert.equal(pivot.source.ref, 'A1:C4');
});

test('the pivot model is exposed only on its host sheet, not the source sheet', () => {
  const back = roundtrip(specialCharsWorkbook());
  assert.equal(back.getWorksheet('Data')?.loadedPivotTables.length, 0);
  assert.equal(back.getWorksheet('Pivot')?.loadedPivotTables.length, 1);
});

test('a loaded pivot decodes XML-special field names back to their original text', () => {
  const wb = new Workbook();
  const src = wb.addWorksheet('Data');
  src.getCell('A1').value = 'Smith & Co';
  src.getCell('B1').value = '<Region>';
  src.getCell('C1').value = 'Am"t';
  src.getCell('A2').value = 'a';
  src.getCell('B2').value = 'x';
  src.getCell('C2').value = 1;
  wb.addWorksheet('Pivot').addPivotTable({
    source: src,
    rows: ['Smith & Co'],
    columns: ['<Region>'],
    values: ['Am"t'],
  });
  const pivot = roundtrip(wb).getWorksheet('Pivot')?.loadedPivotTables[0];
  assert.ok(pivot);
  assert.deepEqual(
    pivot.fields.map((field) => field.name),
    ['Smith & Co', '<Region>', 'Am"t'],
  );
  assert.equal(pivot.valueFieldName, 'Am"t');
});

test('an authored pivot is not surfaced as a loaded one before a round-trip', () => {
  assert.equal(specialCharsWorkbook().getWorksheet('Pivot')?.loadedPivotTables.length, 0);
});

test('two pivot tables number their parts and caches independently', () => {
  const wb = new Workbook();
  const src = wb.addWorksheet('Data');
  src.getCell('A1').value = 'Name';
  src.getCell('B1').value = 'Region';
  src.getCell('C1').value = 'Amount';
  src.getCell('A2').value = 'a';
  src.getCell('B2').value = 'x';
  src.getCell('C2').value = 1;
  wb.addWorksheet('P1').addPivotTable({
    source: src,
    rows: ['Name'],
    columns: ['Region'],
    values: ['Amount'],
  });
  wb.addWorksheet('P2').addPivotTable({
    source: src,
    rows: ['Name'],
    columns: ['Region'],
    values: ['Amount'],
  });

  const parts = partsWritten(wb);
  assert.ok(parts['xl/pivotCache/pivotCacheDefinition1.xml']);
  assert.ok(parts['xl/pivotCache/pivotCacheDefinition2.xml']);
  const workbook = partIn(parts, 'xl/workbook.xml');
  assert.match(workbook, /<pivotCache cacheId="1"/);
  assert.match(workbook, /<pivotCache cacheId="2"/);
});

test('a value field that is also an axis field is flagged as a data field on its axis', () => {
  for (const [values, axis] of [
    [['Name'], 'axisRow'],
    [['Region'], 'axisCol'],
  ] as const) {
    const wb = new Workbook();
    const src = wb.addWorksheet('Data');
    src.getCell('A1').value = 'Name';
    src.getCell('B1').value = 'Region';
    src.getCell('A2').value = 'a';
    src.getCell('B2').value = 'x';
    wb.addWorksheet('P').addPivotTable({
      source: src,
      rows: ['Name'],
      columns: ['Region'],
      values,
      metric: 'count',
    });

    const table = partIn(partsWritten(wb), 'xl/pivotTables/pivotTable1.xml');
    assert.match(table, new RegExp(`<pivotField axis="${axis}" dataField="1" showAll="0">`));
    assert.equal(table.match(/dataField="1"/g)?.length, 1, 'only the value field is flagged');
  }
});
