import assert from 'node:assert/strict';
import {test} from 'node:test';

import {XML_DECLARATION} from '../../xml/xml.ts';
import {CellMetadataTable, parseDynamicArrayCellMetadata} from './cell-metadata.ts';

// The part as Excel 16.0 (build 20326) saved it for a workbook holding three dynamic-array formulas.
const EXCEL_SAVED =
  '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
  'xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">' +
  '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" ' +
  'pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" ' +
  'clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>' +
  '<futureMetadata name="XLDAPR" count="1"><bk><extLst>' +
  '<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">' +
  '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>' +
  '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>';

// A part declaring a rich-value type ahead of the dynamic-array one, with a cell-metadata block for each.
function twoTypes(dynamicProperties: string): string {
  return (
    '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<metadataTypes count="2"><metadataType name="XLRICHVALUE"/><metadataType name="XLDAPR"/></metadataTypes>' +
    `<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext>${dynamicProperties}</ext></extLst></bk></futureMetadata>` +
    '<cellMetadata count="2"><bk><rc t="1" v="0"/></bk><bk><rc t="2" v="0"/></bk></cellMetadata>' +
    '</metadata>'
  );
}

test('the cell metadata Excel saves marks its one block as a dynamic array', () => {
  assert.deepEqual([...parseDynamicArrayCellMetadata(EXCEL_SAVED)], [1]);
});

test('the writer emits the part Excel saves, byte for byte past the declaration', () => {
  const table = new CellMetadataTable();
  assert.equal(table.isEmpty, true, 'a table no cell pointed into is empty');
  assert.equal(table.markDynamicArray(), 1);
  assert.equal(table.isEmpty, false);
  assert.equal(table.toXml(), XML_DECLARATION + EXCEL_SAVED);
});

test('a block marks a dynamic array only through the dynamic-array type, with fDynamic set', () => {
  const dynamic = '<xda:dynamicArrayProperties xmlns:xda="urn:x" fDynamic="1"/>';
  assert.deepEqual(
    [...parseDynamicArrayCellMetadata(twoTypes(dynamic))],
    [2],
    'the block pointing at the rich-value type marks nothing',
  );
  const off = '<xda:dynamicArrayProperties xmlns:xda="urn:x" fDynamic="0"/>';
  assert.deepEqual(
    [...parseDynamicArrayCellMetadata(twoTypes(off))],
    [],
    'fDynamic off marks none',
  );
});

test('a record naming a block the part does not hold marks nothing', () => {
  const beyond = EXCEL_SAVED.replace('<rc t="1" v="0"/>', '<rc t="1" v="3"/>');
  assert.deepEqual([...parseDynamicArrayCellMetadata(beyond)], []);
  const unreadable = EXCEL_SAVED.replace('<rc t="1" v="0"/>', '<rc t="one" v="0"/>');
  assert.deepEqual([...parseDynamicArrayCellMetadata(unreadable)], []);
});

test('value metadata is not cell metadata, even after an empty cell-metadata list', () => {
  const xml = EXCEL_SAVED.replace(
    '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>',
    '<cellMetadata count="0"/><valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata>',
  );
  assert.deepEqual([...parseDynamicArrayCellMetadata(xml)], []);
});

test('an absent part marks nothing', () => {
  assert.deepEqual([...parseDynamicArrayCellMetadata('')], []);
});
