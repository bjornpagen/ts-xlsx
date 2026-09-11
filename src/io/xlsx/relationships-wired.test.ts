// The writer's relationship graph, checked from its output: every id a part cites is declared, every
// declared relationship reaches a part and is used. `assertRelationshipsWired` also runs inside
// `roundtrip` and `partsWritten`, so every test writing through them checks it too; this file proves it
// against a workbook carrying every relationship kind the writer generates, and proves it can fail.

import assert from 'node:assert/strict';
import {test} from 'node:test';

import {threadAt} from '../../core/comment-thread.test-support.ts';
import {Workbook} from '../../core/workbook.ts';
import {relTypeSegment} from '../../rel-type.ts';
import {buildVbaProjectBin, CODE_PAGE, MODULES} from '../../vba/vba.test-support.ts';
import {parseRelationshipRecords} from '../opc/read-opc.ts';
import {RelationshipLedger} from './package-plan.ts';
import {
  assertRelationshipsWired,
  optionalPartText,
  partText,
  patchParts,
  SHEET1,
} from './package.test-support.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

const ONE_PX_PNG = Uint8Array.from(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
  ),
  (c) => c.charCodeAt(0),
);

function everyRelationshipKind(): Workbook {
  const workbook = new Workbook();
  workbook.vbaProjectBytes = buildVbaProjectBin(CODE_PAGE, MODULES);
  workbook.addPerson({id: '{AAAAAAAA-1111-2222-3333-444444444444}', displayName: 'Ada'});

  const sheet = workbook.addWorksheet('Data');
  sheet.addTable({name: 'Sales', ref: 'E1', columns: [{name: 'h1'}, {name: 'h2'}], rowCount: 1});
  const image = workbook.addImage({buffer: ONE_PX_PNG, extension: 'png'});
  sheet.addImage(image, {tl: {col: 0, row: 5}, br: {col: 2, row: 8}});
  sheet.addImage(image, {tl: {col: 3, row: 5}, br: {col: 5, row: 8}});
  sheet.addBackgroundImage(image);
  sheet.getCell('A10').note = 'a note';
  sheet.addCommentThread(threadAt('B10'));
  sheet.pageSetup.printerSettings = Uint8Array.of(1, 2, 3);
  sheet.getCell('A12').value = {text: 'site', hyperlink: 'https://example.invalid/'};
  sheet.getCell('A13').value = {text: 'here', hyperlink: '#Data!A1'};

  const source = workbook.addWorksheet('Source');
  source.addRow(['Name', 'Region', 'Amount']);
  source.addRow(['a', 'north', 1]);
  source.addRow(['b', 'south', 2]);
  workbook.addWorksheet('Pivot').addPivotTable({
    source,
    rows: ['Name'],
    columns: ['Region'],
    values: ['Amount'],
  });
  return workbook;
}

function relationshipTypes(pkg: Uint8Array, relsPath: string): string[] {
  return parseRelationshipRecords(partText(pkg, relsPath)).map((rel) => relTypeSegment(rel.type));
}

test('a workbook carrying every generated relationship kind is wired both ways', () => {
  const pkg = writeXlsx(everyRelationshipKind(), {useSharedStrings: true});

  // Not vacuous: each kind is really in the package the check passes over.
  assert.deepEqual(relationshipTypes(pkg, 'xl/worksheets/_rels/sheet1.xml.rels').sort(), [
    'comments',
    'drawing',
    'hyperlink',
    'image',
    'printerSettings',
    'table',
    'threadedComment',
    'vmlDrawing',
  ]);
  assert.deepEqual(relationshipTypes(pkg, 'xl/worksheets/_rels/sheet3.xml.rels'), ['pivotTable']);
  assert.deepEqual(relationshipTypes(pkg, 'xl/drawings/_rels/drawing1.xml.rels'), [
    'image',
    'image',
  ]);
  assert.deepEqual(relationshipTypes(pkg, 'xl/pivotTables/_rels/pivotTable1.xml.rels'), [
    'pivotCacheDefinition',
  ]);
  assert.deepEqual(relationshipTypes(pkg, 'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels'), [
    'pivotCacheRecords',
  ]);
  for (const kind of ['sharedStrings', 'person', 'vbaProject', 'pivotCacheDefinition']) {
    assert.ok(relationshipTypes(pkg, 'xl/_rels/workbook.xml.rels').includes(kind), kind);
  }

  assertRelationshipsWired(pkg);
  assertRelationshipsWired(writeXlsx(readXlsx(pkg)));
});

test('ids in a generated .rels part run in order from rId1, whatever kinds a sheet mixes', () => {
  const pkg = writeXlsx(everyRelationshipKind());
  const ids = parseRelationshipRecords(partText(pkg, 'xl/worksheets/_rels/sheet1.xml.rels')).map(
    (rel) => rel.id,
  );
  assert.deepEqual(
    ids,
    ids.map((_, i) => `rId${i + 1}`),
  );
});

test('a sheet that relates to nothing gets no .rels part', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').getCell('A1').value = {text: 'here', hyperlink: '#S!B2'};
  assert.equal(
    optionalPartText(writeXlsx(workbook), 'xl/worksheets/_rels/sheet1.xml.rels'),
    undefined,
  );
});

test('the check refuses a cited id nothing declares', () => {
  const pkg = patchParts(writeXlsx(everyRelationshipKind()), {
    [SHEET1]: (xml) => xml.replace(/<tablePart r:id="rId\d+"\/>/, '<tablePart r:id="rId99"/>'),
  });
  assert.throws(() => assertRelationshipsWired(pkg), /sheet1\.xml cites rId99/);
});

test('the check refuses a declared relationship nothing cites', () => {
  const pkg = patchParts(writeXlsx(everyRelationshipKind()), {
    [SHEET1]: (xml) => xml.replace(/<tableParts\b.*?<\/tableParts>/, ''),
  });
  assert.throws(() => assertRelationshipsWired(pkg), /\(table\), which nothing cites/);
});

test('the check refuses a relationship to a part the package does not hold', () => {
  const pkg = patchParts(writeXlsx(everyRelationshipKind()), {
    'xl/worksheets/_rels/sheet1.xml.rels': (xml) =>
      xml.replace('../tables/table1.xml', '../tables/table9.xml'),
  });
  assert.throws(() => assertRelationshipsWired(pkg), /xl\/tables\/table9\.xml, which the package/);
});

test('a ledger numbers its relationships in order and names targets from its owner', () => {
  const rels = new RelationshipLedger('xl/worksheets/sheet3.xml');
  assert.equal(rels.add('t/table', 'xl/tables/table2.xml'), 'rId1');
  assert.equal(rels.addExternal('t/hyperlink', 'https://example.invalid/'), 'rId2');
  assert.equal(rels.add('t/image', 'xl/media/image1.png'), 'rId3');
  assert.deepEqual(rels.relationships, [
    {id: 'rId1', type: 't/table', target: '../tables/table2.xml'},
    {id: 'rId2', type: 't/hyperlink', target: 'https://example.invalid/', external: true},
    {id: 'rId3', type: 't/image', target: '../media/image1.png'},
  ]);

  const root = new RelationshipLedger('');
  root.add('t/officeDocument', 'xl/workbook.xml');
  assert.equal(
    root.relationships[0]?.target,
    'xl/workbook.xml',
    'the package root names paths as they are',
  );
});
