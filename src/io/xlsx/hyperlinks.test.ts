import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Workbook} from '../../core/workbook.ts';
import {
  elementIn,
  foreignPackage,
  optionalPartIn,
  partIn,
  partsWritten,
  patchParts,
  roundtrip,
  SHEET1,
} from './package.test-support.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

test('an external hyperlink round-trips with its target, and the cell keeps its own value', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'Example';
  sheet.addHyperlink({ref: 'A1', target: 'https://example.com'});

  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual(back?.hyperlinks, [{ref: 'A1', target: 'https://example.com'}]);
  assert.equal(back?.getCell('A1').value, 'Example');
});

test('an external URL keeps its "#" fragment through a round-trip', () => {
  const url = 'http://host/ui/#/case/2007720723';
  const wb = new Workbook();
  wb.addWorksheet('S').addHyperlink({ref: 'A1', target: url});

  assert.equal(roundtrip(wb).getWorksheet('S')?.hyperlinkAt('A1')?.target, url);
});

test('a tooltip survives the round-trip', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addHyperlink({
    ref: 'A1',
    target: 'https://example.com',
    tooltip: 'go to example',
  });

  assert.equal(roundtrip(wb).getWorksheet('S')?.hyperlinkAt('A1')?.tooltip, 'go to example');
});

test('an internal "#"-target is written as a location with no external relationship', () => {
  const wb = new Workbook();
  wb.addWorksheet('Main').addHyperlink({ref: 'A1', target: "#'Target'!A1"});
  wb.addWorksheet('Target');

  const parts = partsWritten(wb);
  const sheetXml = partIn(parts, 'xl/worksheets/sheet1.xml');
  const link = elementIn(sheetXml, /<hyperlink\b[^>]*\/?>/);
  assert.match(link, /location="[^"]*Target[^"]*A1[^"]*"/, 'the internal target rides in location');
  assert.doesNotMatch(link, /r:id=/, 'an internal link uses no relationship id');
  // An internal link must not produce a sheet rels part carrying an External relationship. Which is
  // two claims, and the `if (rels !== undefined)` this used to be spelled with made the second one
  // vacuous: a writer that stopped emitting the rels part entirely passed it.
  const rels = optionalPartIn(parts, 'xl/worksheets/_rels/sheet1.xml.rels');
  assert.ok(rels === undefined || !/TargetMode="External"/.test(rels));
});

test('an internal "#"-target round-trips verbatim', () => {
  const wb = new Workbook();
  wb.addWorksheet('Main').addHyperlink({ref: 'A1', target: '#Sheet2!A1'});
  wb.addWorksheet('Sheet2');

  assert.equal(roundtrip(wb).getWorksheet('Main')?.hyperlinkAt('A1')?.target, '#Sheet2!A1');
});

test('an external link produces exactly one External relationship of hyperlink type', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addHyperlink({ref: 'A1', target: 'https://example.com'});

  const parts = partsWritten(wb);
  const rels = partIn(parts, 'xl/worksheets/_rels/sheet1.xml.rels');
  const external = /<Relationship\b[^>]*TargetMode="External"[^>]*\/>/g;
  assert.equal([...rels.matchAll(external)].length, 1);
  const link = elementIn(rels, external);
  assert.match(link, /Type="[^"]*\/hyperlink"/);
  assert.match(link, /Target="https:\/\/example\.com"/);
});

test('the <hyperlinks> element sits after <mergeCells> and before <pageMargins>', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.addHyperlink({ref: 'A1', target: 'https://example.com'});
  sheet.mergeCells('B1:C1');
  sheet.pageMargins.left = 0.5;

  const sheetXml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  const merge = sheetXml.indexOf('<mergeCells');
  const links = sheetXml.indexOf('<hyperlinks>');
  const margins = sheetXml.indexOf('<pageMargins');
  assert.ok(merge >= 0 && links >= 0 && margins >= 0);
  assert.ok(
    merge < links && links < margins,
    `order was mergeCells@${merge} hyperlinks@${links} pageMargins@${margins}`,
  );
});

test('the reader rejoins a foreign file’s location fragment onto the relationship target', () => {
  // A foreign producer stores an external URL's fragment in the hyperlink's `location`, apart from
  // the bare relationship Target: the reader must rejoin them into the whole URL.
  const archive = foreignPackage({
    [SHEET1]:
      '<?xml version="1.0"?><worksheet xmlns:r="x"><sheetData>' +
      '<row r="1"><c r="A1" t="inlineStr"><is><t>link</t></is></c></row>' +
      '</sheetData><hyperlinks><hyperlink ref="A1" r:id="rId1" location="myhash"/></hyperlinks></worksheet>',
    'xl/worksheets/_rels/sheet1.xml.rels':
      '<Relationships><Relationship Id="rId1" Type="x/hyperlink" Target="http://localhost/" TargetMode="External"/></Relationships>',
  });

  const back = readXlsx(archive).getWorksheet('S');
  assert.equal(back?.hyperlinkAt('A1')?.target, 'http://localhost/#myhash');
  assert.equal(back?.getCell('A1').value, 'link');
});

test('a hyperlink relationship id does not collide with a table on the same sheet', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'h';
  sheet.addTable({name: 'T', ref: 'A3', columns: [{name: 'c'}], rowCount: 1});
  sheet.addHyperlink({ref: 'A1', target: 'https://example.com'});

  const parts = partsWritten(wb);
  const rels = partIn(parts, 'xl/worksheets/_rels/sheet1.xml.rels');
  const ids = [...rels.matchAll(/Id="(rId\d+)"/g)].map((m) => m[1]);
  // Uniqueness is vacuously true of nothing, and the point of this test is that the hyperlink and
  // the table both claim an id in the same part: fewer than two and it is not testing that at all.
  assert.ok(ids.length >= 2, `expected the link and the table to claim ids; got ${ids.join(', ')}`);
  assert.equal(
    new Set(ids).size,
    ids.length,
    `relationship ids must be unique; got ${ids.join(', ')}`,
  );
  // The link reads back intact despite sharing the rels part with the table.
  assert.equal(roundtrip(wb).getWorksheet('S')?.hyperlinkAt('A1')?.target, 'https://example.com');
});

test('a link over a number, boolean, formula, date, label or empty cell is read, and the value kept', () => {
  // OOXML stores a link beside the cell, so Excel puts one on any value. The model used to fold a link
  // into the value as a label, and had to choose between a non-text value and its link.
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  const when = new Date(Date.UTC(2024, 0, 15));
  sheet.getCell('A1').value = 42;
  sheet.getCell('B1').value = true;
  sheet.getCell('C1').value = {formula: '1+1', result: 2};
  sheet.getCell('D1').value = when;
  sheet.getCell('E1').value = 'label';
  const refs = ['A1', 'B1', 'C1', 'D1', 'E1', 'F1'];
  const links = refs.map((ref) => `<hyperlink ref="${ref}" location="S!H1"/>`).join('');
  const read = readXlsx(
    patchParts(writeXlsx(wb), {
      [SHEET1]: (xml) =>
        xml.replace('</sheetData>', `</sheetData><hyperlinks>${links}</hyperlinks>`),
    }),
  );
  const back = read.getWorksheet('S');
  assert.ok(back);

  assert.deepEqual(
    back.hyperlinks,
    refs.map((ref) => ({ref, target: '#S!H1'})),
  );
  assert.equal(back.getCell('A1').value, 42);
  assert.equal(back.getCell('B1').value, true);
  assert.deepEqual(back.getCell('C1').value, {formula: '1+1', result: 2});
  const date = back.getCell('D1').value;
  assert.ok(date instanceof Date && date.getTime() === when.getTime(), 'the date is still a date');
  assert.equal(back.getCell('E1').value, 'label');
  assert.equal(back.getCell('F1').value, null);
});

test('a hyperlink spanning a range is read with its whole range', () => {
  const archive = foreignPackage({
    [SHEET1]:
      '<?xml version="1.0"?><worksheet xmlns:r="x"><sheetData>' +
      '<row r="1"><c r="D1" t="inlineStr"><is><t>go</t></is></c></row>' +
      '</sheetData><hyperlinks><hyperlink ref="D1:H1" location="Sheet1!A1"/></hyperlinks></worksheet>',
  });

  const back = readXlsx(archive).getWorksheet('S');
  assert.deepEqual(back?.hyperlinks, [{ref: 'D1:H1', target: '#Sheet1!A1'}]);
  assert.equal(back?.hyperlinkAt('F1')?.target, '#Sheet1!A1', 'every covered cell opens it');
});

test('a link naming a whole column, or nothing at all, is dropped on read and the rest are kept', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const links =
    '<hyperlink ref="A:A" location="S!H1"/><hyperlink ref="junk!!" location="S!H1"/>' +
    '<hyperlink ref="B2" location="S!H1"/>';
  const back = readXlsx(
    patchParts(writeXlsx(wb), {
      [SHEET1]: (xml) =>
        xml.replace('</sheetData>', `</sheetData><hyperlinks>${links}</hyperlinks>`),
    }),
  ).getWorksheet('S');

  assert.deepEqual(back?.hyperlinks, [{ref: 'B2', target: '#S!H1'}]);
});

test('two links over one cell keep their order through a save, so the same one wins', () => {
  // Excel keeps a link added over a cell inside a wider link beside it.
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.addHyperlink({ref: 'A1:C1', target: 'https://one.example/'});
  sheet.addHyperlink({ref: 'B1', target: 'https://two.example/'});

  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual(
    back?.hyperlinks.map((link) => link.ref),
    ['A1:C1', 'B1'],
  );
  assert.equal(back?.hyperlinkAt('B1')?.target, 'https://two.example/');
  assert.equal(back?.hyperlinkAt('C1')?.target, 'https://one.example/');
});
