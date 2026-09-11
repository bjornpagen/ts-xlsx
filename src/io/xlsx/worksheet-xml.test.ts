// Where worksheet-xml.ts places each element in a worksheet part: CT_Worksheet is a sequence, and a
// consumer repairs a part whose children are out of schema order.

import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Workbook} from '../../core/workbook.ts';
import {partIn, partsWritten, sheetXml} from './package.test-support.ts';
import {writeXlsx} from './write.ts';

test('<pageMargins> is placed after <sheetData>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.pageMargins.top = 1;
  const xml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  assert.ok(
    xml.indexOf('<sheetData') < xml.indexOf('<pageMargins'),
    'pageMargins must follow sheetData',
  );
});

test('<printOptions> precedes <pageMargins> in schema order', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.printOptions.horizontalCentered = true;
  s.pageMargins.top = 1;
  const xml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  assert.ok(
    xml.indexOf('<printOptions') < xml.indexOf('<pageMargins'),
    'printOptions must precede pageMargins',
  );
});

test('<cols> is placed after <sheetFormatPr> and before <sheetData>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.getColumn(1).width = 8;
  const xml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  const fmt = xml.indexOf('<sheetFormatPr');
  const cols = xml.indexOf('<cols>');
  const data = xml.indexOf('<sheetData>');
  assert.ok(
    fmt < cols && cols < data,
    `expected sheetFormatPr < cols < sheetData, got ${fmt},${cols},${data}`,
  );
});

test('<tableParts> follows <headerFooter> in the worksheet element order', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.headerFooter.oddHeader = 'H';
  s.addTable({name: 'T', ref: 'A1', columns: [{name: 'A'}], rowCount: 1});
  const xml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  assert.ok(
    xml.indexOf('<headerFooter') < xml.indexOf('<tableParts'),
    'tableParts must follow headerFooter per CT_Worksheet',
  );
});

test('<pageSetup> sits between <pageMargins> and <headerFooter>', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.pageMargins.top = 1;
  sheet.pageSetup.scale = 90;
  sheet.headerFooter.oddHeader = 'H';

  const xml = sheetXml(writeXlsx(wb));
  assert.ok(
    xml.indexOf('<pageMargins') < xml.indexOf('<pageSetup'),
    'pageSetup follows pageMargins',
  );
  assert.ok(
    xml.indexOf('<pageSetup') < xml.indexOf('<headerFooter'),
    'pageSetup precedes headerFooter',
  );
});
