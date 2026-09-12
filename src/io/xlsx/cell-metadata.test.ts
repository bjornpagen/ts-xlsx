import assert from 'node:assert/strict';
import {test} from 'node:test';

import {XML_DECLARATION} from '../../xml/xml.ts';
import {indexCellMetadata} from '../cell-metadata/metadata.ts';
import {parseRichValueErrors} from '../cell-metadata/rich-values.ts';
import {
  parseMetadataPart,
  RICH_VALUE_STRUCTURES_XML,
  WorkbookMetadataTable,
} from './cell-metadata.ts';

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const XLRD = 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata';
const XDA = 'http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray';
const FLAGS =
  'minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" ' +
  'rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"';
const DYNAMIC_TYPE = `<metadataType name="XLDAPR" ${FLAGS} cellMeta="1"/>`;
const DYNAMIC_BLOCKS =
  '<futureMetadata name="XLDAPR" count="1"><bk><extLst>' +
  '<ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}">' +
  '<xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>';
const rvb = (i: number) =>
  `<bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="${i}"/></ext></extLst></bk>`;

// The part as Excel 16.0 (build 20326) saved it for a workbook holding three dynamic-array formulas.
const EXCEL_SAVED_DYNAMIC =
  `<metadata xmlns="${MAIN}" xmlns:xda="${XDA}">` +
  `<metadataTypes count="1">${DYNAMIC_TYPE}</metadataTypes>${DYNAMIC_BLOCKS}` +
  '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>';

// The part in the form Excel saved it for dynamic-array formulas beside two rich-value errors.
const EXCEL_SAVED_BOTH =
  `<metadata xmlns="${MAIN}" xmlns:xlrd="${XLRD}" xmlns:xda="${XDA}">` +
  `<metadataTypes count="2">${DYNAMIC_TYPE}<metadataType name="XLRICHVALUE" ${FLAGS}/></metadataTypes>` +
  `${DYNAMIC_BLOCKS}<futureMetadata name="XLRICHVALUE" count="2">${rvb(0)}${rvb(1)}</futureMetadata>` +
  '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>' +
  '<valueMetadata count="2"><bk><rc t="2" v="0"/></bk><bk><rc t="2" v="1"/></bk></valueMetadata>' +
  '</metadata>';

const dynamicCells = (xml: string) => [
  ...indexCellMetadata(parseMetadataPart(xml), new Map()).dynamicArrayCells,
];

test('the cell metadata Excel saves marks its one block as a dynamic array', () => {
  assert.deepEqual(dynamicCells(EXCEL_SAVED_DYNAMIC), [1]);
});

test('the writer emits the part Excel saves, byte for byte past the declaration', () => {
  const table = new WorkbookMetadataTable();
  assert.equal(table.isEmpty, true, 'a table no cell pointed into is empty');
  assert.equal(table.markDynamicArray(), 1);
  assert.equal(table.isEmpty, false);
  assert.equal(table.hasRichValues, false);
  assert.equal(table.toXml(), XML_DECLARATION + EXCEL_SAVED_DYNAMIC);
});

test('the value metadata Excel saves names one rich value per block', () => {
  const metadata = parseMetadataPart(EXCEL_SAVED_BOTH);
  assert.deepEqual(metadata.typeNames, ['XLDAPR', 'XLRICHVALUE']);
  assert.deepEqual(metadata.richValueBlocks, [0, 1]);
  assert.deepEqual(metadata.valueBlocks, [[{type: 2, value: 0}], [{type: 2, value: 1}]]);
  const errors = new Map([
    [0, '#SPILL!'],
    [1, '#CALC!'],
  ] as const);
  assert.deepEqual(
    [...indexCellMetadata(metadata, errors).valueErrors],
    [
      [1, '#SPILL!'],
      [2, '#CALC!'],
    ],
  );
});

test('the writer shares one rich value per error, in the form Excel saves the part', () => {
  const table = new WorkbookMetadataTable();
  table.markDynamicArray();
  assert.equal(table.markRichValueError('#SPILL!'), 1);
  assert.equal(table.markRichValueError('#CALC!'), 2);
  assert.equal(table.markRichValueError('#SPILL!'), 1, 'a second #SPILL! shares the first');
  assert.equal(table.markRichValueError('#N/A'), undefined, 'a literal needs no metadata');
  assert.equal(table.hasRichValues, true);
  assert.equal(table.toXml(), XML_DECLARATION + EXCEL_SAVED_BOTH);
  assert.equal(
    table.richValuesXml(),
    `${XML_DECLARATION}<rvData xmlns="${XLRD}" count="2">` +
      '<rv s="0"><v>8</v><v>0</v></rv><rv s="0"><v>13</v><v>0</v></rv></rvData>',
  );
  const index = indexCellMetadata(
    parseMetadataPart(table.toXml()),
    parseRichValueErrors(RICH_VALUE_STRUCTURES_XML, table.richValuesXml()),
  );
  assert.deepEqual(
    [...index.valueErrors],
    [
      [1, '#SPILL!'],
      [2, '#CALC!'],
    ],
  );
});

test('rich-value errors alone declare only their own type', () => {
  const table = new WorkbookMetadataTable();
  table.markRichValueError('#FIELD!');
  assert.equal(
    table.toXml(),
    `${XML_DECLARATION}<metadata xmlns="${MAIN}" xmlns:xlrd="${XLRD}">` +
      `<metadataTypes count="1"><metadataType name="XLRICHVALUE" ${FLAGS}/></metadataTypes>` +
      `<futureMetadata name="XLRICHVALUE" count="1">${rvb(0)}</futureMetadata>` +
      '<valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>',
  );
});

// A part declaring a rich-value type ahead of the dynamic-array one, with a cell-metadata block for each.
function twoTypes(dynamicProperties: string): string {
  return (
    `<metadata xmlns="${MAIN}">` +
    '<metadataTypes count="2"><metadataType name="XLRICHVALUE"/><metadataType name="XLDAPR"/></metadataTypes>' +
    `<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext>${dynamicProperties}</ext></extLst></bk></futureMetadata>` +
    '<cellMetadata count="2"><bk><rc t="1" v="0"/></bk><bk><rc t="2" v="0"/></bk></cellMetadata>' +
    '</metadata>'
  );
}

test('a block marks a dynamic array only through the dynamic-array type, with fDynamic set', () => {
  const dynamic = '<xda:dynamicArrayProperties xmlns:xda="urn:x" fDynamic="1"/>';
  assert.deepEqual(dynamicCells(twoTypes(dynamic)), [2], 'the rich-value block marks nothing');
  const off = '<xda:dynamicArrayProperties xmlns:xda="urn:x" fDynamic="0"/>';
  assert.deepEqual(dynamicCells(twoTypes(off)), [], 'fDynamic off marks none');
});

test('a record naming a block the part does not hold marks nothing', () => {
  const beyond = EXCEL_SAVED_DYNAMIC.replace('<rc t="1" v="0"/>', '<rc t="1" v="3"/>');
  assert.deepEqual(dynamicCells(beyond), []);
  const unreadable = EXCEL_SAVED_DYNAMIC.replace('<rc t="1" v="0"/>', '<rc t="one" v="0"/>');
  assert.deepEqual(dynamicCells(unreadable), []);
});

test('value metadata is not cell metadata, even after an empty cell-metadata list', () => {
  const xml = EXCEL_SAVED_DYNAMIC.replace(
    '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>',
    '<cellMetadata count="0"/><valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata>',
  );
  assert.deepEqual(dynamicCells(xml), []);
  assert.deepEqual(parseMetadataPart(xml).valueBlocks, [[{type: 1, value: 0}]]);
});

test('an absent part marks nothing', () => {
  assert.deepEqual(dynamicCells(''), []);
});
