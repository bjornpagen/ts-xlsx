import assert from 'node:assert/strict';
import {test} from 'node:test';

import {isHyperlinkValue} from '../../core/value.ts';
import {Workbook} from '../../core/workbook.ts';
import {
  elementIn,
  foreignPackage,
  optionalPartIn,
  partIn,
  partsWritten,
  roundtrip,
  SHEET1,
} from './package.test-support.ts';
import {readXlsx} from './read.ts';

function hyperlinkOf(workbook: Workbook, sheet: string, ref: string) {
  const value = workbook.getWorksheet(sheet)?.getCell(ref).value;
  assert.ok(
    value !== undefined && value !== null && isHyperlinkValue(value),
    'expected a hyperlink value',
  );
  return value;
}

test('an external hyperlink round-trips with its target and visible label', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {hyperlink: 'https://example.com', text: 'Example'};

  const back = hyperlinkOf(roundtrip(wb), 'S', 'A1');
  assert.equal(back.hyperlink, 'https://example.com');
  assert.equal(back.text, 'Example');
});

test('an external URL keeps its "#" fragment through a round-trip', () => {
  const url = 'http://host/ui/#/case/2007720723';
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {hyperlink: url, text: 'open case'};

  const back = hyperlinkOf(roundtrip(wb), 'S', 'A1');
  assert.equal(back.hyperlink, url, 'the fragment tail must not be dropped');
  assert.equal(back.text, 'open case');
});

test('a tooltip survives the round-trip', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {
    hyperlink: 'https://example.com',
    text: 'Example',
    tooltip: 'go to example',
  };

  const back = hyperlinkOf(roundtrip(wb), 'S', 'A1');
  assert.equal(back.tooltip, 'go to example');
});

test('an internal "#"-target is written as a location with no external relationship', () => {
  const wb = new Workbook();
  wb.addWorksheet('Main').getCell('A1').value = {hyperlink: "#'Target'!A1", text: 'go'};
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
  wb.addWorksheet('Main').getCell('A1').value = {hyperlink: '#Sheet2!A1', text: 'go'};
  wb.addWorksheet('Sheet2');

  const back = hyperlinkOf(roundtrip(wb), 'Main', 'A1');
  assert.equal(back.hyperlink, '#Sheet2!A1');
});

test('an external link produces exactly one External relationship of hyperlink type', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {hyperlink: 'https://example.com', text: 'x'};

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
  sheet.getCell('A1').value = {hyperlink: 'https://example.com', text: 'x'};
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

  const back = hyperlinkOf(readXlsx(archive), 'S', 'A1');
  assert.equal(back.hyperlink, 'http://localhost/#myhash');
  assert.equal(back.text, 'link');
});

test('a hyperlink relationship id does not collide with a table on the same sheet', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'h';
  sheet.addTable({name: 'T', ref: 'A3', columns: [{name: 'c'}], rowCount: 1});
  sheet.getCell('A1').value = {hyperlink: 'https://example.com', text: 'h'};

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
  const back = hyperlinkOf(roundtrip(wb), 'S', 'A1');
  assert.equal(back.hyperlink, 'https://example.com');
});

test('a hyperlink spanning a range anchors on its top-left cell instead of crashing', () => {
  // Excel writes a multi-cell hyperlink as `ref="D1:H1"`; the reader must fold it onto the range's
  // top-left cell (D1) rather than asking the sheet for a range address it cannot resolve.
  const archive = foreignPackage({
    [SHEET1]:
      '<?xml version="1.0"?><worksheet xmlns:r="x"><sheetData>' +
      '<row r="1"><c r="D1" t="inlineStr"><is><t>go</t></is></c></row>' +
      '</sheetData><hyperlinks><hyperlink ref="D1:H1" location="Sheet1!A1"/></hyperlinks></worksheet>',
  });

  const back = hyperlinkOf(readXlsx(archive), 'S', 'D1');
  assert.equal(back.hyperlink, '#Sheet1!A1');
  assert.equal(back.text, 'go');
  // The extent is not just anchored: it is recorded, so the clickable area is not silently shrunk.
  assert.equal(back.range, 'D1:H1');
});

test('a range hyperlink keeps its extent through a write→read round-trip', () => {
  // A single-cell link must stay plain; only a genuine multi-cell link carries `range`, and that
  // extent has to survive serialisation rather than collapsing back to the top-left cell.
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('D1').value = {hyperlink: 'https://example.com', text: 'go', range: 'D1:H1'};
  sheet.getCell('A1').value = {hyperlink: 'https://example.com/plain', text: 'one'};

  const back = roundtrip(wb);
  const ranged = hyperlinkOf(back, 'S', 'D1');
  assert.equal(ranged.range, 'D1:H1');
  assert.equal(ranged.hyperlink, 'https://example.com');
  const single = hyperlinkOf(back, 'S', 'A1');
  assert.equal(single.range, undefined);
});
