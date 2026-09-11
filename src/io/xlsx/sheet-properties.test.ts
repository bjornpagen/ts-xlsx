// The worksheet print layout sheet-properties.ts reads and writes: page setup and the printer-settings
// blob behind it, margins, print options, header/footer, and manual page breaks.

import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Workbook} from '../../core/workbook.ts';
import {
  captureIn,
  optionalPartText,
  partBytes,
  partIn,
  partsWritten,
  partText,
  patchParts,
  roundtrip,
  SHEET1,
  sheetXml,
} from './package.test-support.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

test('setting a subset of margins emits all six pageMargins attributes', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.pageMargins.left = 0.1;
  s.pageMargins.right = 0.1;
  const xml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  const tag = captureIn(xml, /<pageMargins ([^/]*)\/>/);
  for (const side of ['left', 'right', 'top', 'bottom', 'header', 'footer']) {
    assert.match(tag, new RegExp(`\\b${side}="[0-9.]+"`), `missing ${side}`);
  }
  // The explicitly-set sides keep their values; the untouched ones fall back to defaults.
  assert.match(tag, /left="0.1"/);
  assert.match(tag, /right="0.1"/);
  assert.match(tag, /top="0.75"/);
});

test('a sheet with no margins set emits no <pageMargins>', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const xml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  assert.doesNotMatch(xml, /<pageMargins/);
});

test('print-option toggles are emitted and survive a write→read round-trip', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.printOptions.horizontalCentered = true;
  s.printOptions.gridLines = true;
  s.printOptions.headings = true;

  const xml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  const tag = captureIn(xml, /<printOptions ([^/]*)\/>/);
  assert.match(tag, /horizontalCentered="1"/);
  assert.match(tag, /gridLines="1"/);
  assert.match(tag, /headings="1"/);
  assert.doesNotMatch(tag, /verticalCentered/, 'an untouched flag is not fabricated');

  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual(back?.printOptions, {horizontalCentered: true, gridLines: true, headings: true});
});

test('a print-option flag forced off round-trips as an explicit "0", not a dropped default', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.printOptions.gridLinesSet = false;

  const xml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<printOptions gridLinesSet="0"\/>/);

  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual(
    back?.printOptions,
    {gridLinesSet: false},
    'the explicit false survives, not swallowed by the default',
  );
});

test('a sheet with no print options set emits no <printOptions> element', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const xml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  assert.doesNotMatch(xml, /<printOptions/, 'an empty print-option set fabricates nothing');
});

test('header/footer variants emit their children and gate them with different* flags', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  Object.assign(s.headerFooter, {
    oddHeader: 'ODD-H',
    evenHeader: 'EVEN-H',
    firstFooter: 'FIRST-F',
  });
  const xml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<headerFooter[^>]* differentOddEven="1"/);
  assert.match(xml, /<headerFooter[^>]* differentFirst="1"/);
  assert.match(xml, /<oddHeader>ODD-H<\/oddHeader>/);
  assert.match(xml, /<evenHeader>EVEN-H<\/evenHeader>/);
  assert.match(xml, /<firstFooter>FIRST-F<\/firstFooter>/);
});

test('an odd-only header/footer sets no different* flags', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.headerFooter.oddHeader = 'H';
  const xml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<headerFooter><oddHeader>H<\/oddHeader><\/headerFooter>/);
  assert.doesNotMatch(xml, /different/);
});

test('a sheet with no header/footer emits no <headerFooter>', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const xml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  assert.doesNotMatch(xml, /<headerFooter/);
});

test('header/footer text is XML-escaped and placed after <pageMargins>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.pageMargins.top = 1;
  s.headerFooter.oddHeader = 'a & b < c';
  const xml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<oddHeader>a &amp; b &lt; c<\/oddHeader>/);
  assert.ok(
    xml.indexOf('<pageMargins') < xml.indexOf('<headerFooter'),
    'headerFooter follows pageMargins',
  );
});

test('header text carries the _xHHHH_ convention, same as a cell value', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  // Excel decodes `_xHHHH_` in a `<headerFooter>` child and writes one back on save (measured over
  // COM), so a header both may carry a character XML itself cannot and must have a literal that
  // *looks* like an escape protected; otherwise `_x0041_` would come back as `A`.
  s.headerFooter.oddHeader = '&C[\u0001][_x0041_]';
  const xml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<oddHeader>&amp;C\[_x0001_\]\[_x005F_x0041_\]<\/oddHeader>/);
});

test('an astral character in a header needs no escape, a control one does', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.headerFooter.oddFooter = '\u{1F600}\u0001';
  const xml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<oddFooter>\u{1F600}_x0001_<\/oddFooter>/u);
});

test('header text round-trips through the escape convention', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  const authored = '&L[_x0041_]&C[\u0001]&R[\u{1F600}]';
  s.headerFooter.oddHeader = authored;
  const reread = roundtrip(wb);
  assert.equal(reread.worksheets[0]?.headerFooter.oddHeader, authored);
});

test('manual row breaks are emitted as <rowBreaks> and round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.rowBreaks.push({id: 3, max: 16383, man: true}, {id: 6, max: 16383, man: true});

  const xml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  assert.match(
    xml,
    /<rowBreaks count="2" manualBreakCount="2">/,
    'both breaks are counted as manual',
  );
  assert.match(
    xml,
    /<brk id="3" max="16383" man="1"\/>/,
    'the first break carries its column span',
  );

  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual(
    back?.rowBreaks.map((brk) => brk.id),
    [3, 6],
    'the break rows survive a write→read round-trip',
  );
});

// Every break was written `man="1"` and counted as manual, so one save hardened a producer's automatic
// breaks into author-set ones; `min` was dropped on read.
test('an automatic break stays automatic, and manualBreakCount counts manual breaks only', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.rowBreaks.push(
    {id: 3, max: 16383, man: true},
    {id: 6, min: 2, max: 16383, man: false},
    {id: 9},
  );

  const xml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<rowBreaks count="3" manualBreakCount="2">/);
  assert.match(xml, /<brk id="6" min="2" max="16383"\/>/, 'the automatic break carries no man');
  assert.match(xml, /<brk id="9" man="1"\/>/, 'a break authored without man is a manual one');

  assert.deepEqual(roundtrip(wb).getWorksheet('S')?.rowBreaks, [
    {id: 3, max: 16383, man: true},
    {id: 6, min: 2, max: 16383, man: false},
    {id: 9, man: true},
  ]);
});

test('a sheet with no manual row breaks emits no <rowBreaks> element', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const xml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  assert.doesNotMatch(xml, /<rowBreaks/, 'an empty break list fabricates nothing');
});

test('column-break <brk> elements land on the column-break model, not the row-break one', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  const patched = patchParts(writeXlsx(wb), {
    [SHEET1]: (xml) =>
      xml.replace(
        '</worksheet>',
        '<colBreaks count="1" manualBreakCount="1"><brk id="2" max="1048575" man="1"/></colBreaks></worksheet>',
      ),
  });
  const back = readXlsx(patched).getWorksheet('S');
  assert.deepEqual(back?.rowBreaks, [], 'a column break must not land on the row-break model');
  assert.deepEqual(
    back?.columnBreaks,
    [{id: 2, max: 1048575, man: true}],
    'the column break is surfaced on the column-break model with its span',
  );
});

test('manual column breaks are emitted as <colBreaks> and round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.columnBreaks.push({id: 2, max: 1048575, man: true}, {id: 5, max: 1048575, man: true});

  const xml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  assert.match(
    xml,
    /<colBreaks count="2" manualBreakCount="2">/,
    'both breaks are counted as manual',
  );
  assert.match(xml, /<brk id="2" max="1048575" man="1"\/>/, 'the first break carries its row span');

  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual(
    back?.columnBreaks.map((brk) => brk.id),
    [2, 5],
    'the break columns survive a write→read round-trip',
  );
});

test('row and column breaks coexist on one sheet without cross-contaminating', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.rowBreaks.push({id: 3, max: 16383, man: true});
  sheet.columnBreaks.push({id: 4, max: 1048575, man: true});

  const xml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  assert.match(
    xml,
    /<rowBreaks[^>]*>.*<\/rowBreaks><colBreaks/,
    '<colBreaks> follows <rowBreaks> in schema order',
  );

  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual(
    back?.rowBreaks.map((brk) => brk.id),
    [3],
    'the row break stays a row break',
  );
  assert.deepEqual(
    back?.columnBreaks.map((brk) => brk.id),
    [4],
    'the column break stays a column break',
  );
});

test('a sheet with no manual column breaks emits no <colBreaks> element', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const xml = partIn(partsWritten(wb), 'xl/worksheets/sheet1.xml');
  assert.doesNotMatch(xml, /<colBreaks/, 'an empty column-break list fabricates nothing');
});

test('page margins round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.pageMargins.left = 0.5;
  sheet.pageMargins.top = 1.25;
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.pageMargins.left, 0.5);
  assert.equal(back?.pageMargins.top, 1.25);
});

test('a fit-to-page setup round-trips its flag, counts, and scale', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.pageSetup.fitToPage = true;
  sheet.pageSetup.fitToWidth = 1;
  sheet.pageSetup.fitToHeight = 0;
  sheet.pageSetup.scale = 80;
  sheet.getCell('A1').value = 'x';

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.pageSetup.fitToPage, true);
  assert.equal(back?.pageSetup.fitToWidth, 1);
  assert.equal(back?.pageSetup.fitToHeight, 0);
  assert.equal(back?.pageSetup.scale, 80);
});

test('the fit-to-page flag rides <pageSetUpPr> under <sheetPr>, the counts ride <pageSetup>', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.pageSetup.fitToPage = true;
  sheet.pageSetup.fitToWidth = 1;

  const xml = sheetXml(writeXlsx(wb));
  assert.match(xml, /<sheetPr><pageSetUpPr fitToPage="1"\/><\/sheetPr>/);
  assert.match(xml, /<pageSetup fitToWidth="1"\/>/);
});

test('<pageSetUpPr> follows <outlinePr> under <sheetPr> in CT_SheetPr order', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.outline.summaryBelow = false;
  sheet.pageSetup.fitToPage = true;

  const xml = sheetXml(writeXlsx(wb));
  assert.match(
    xml,
    /<sheetPr><outlinePr summaryBelow="0"\/><pageSetUpPr fitToPage="1"\/><\/sheetPr>/,
  );
});

test('orientation and pageOrder round-trip and emit only when set', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.pageSetup.orientation = 'landscape';
  sheet.pageSetup.pageOrder = 'overThenDown';

  const xml = sheetXml(writeXlsx(wb));
  assert.match(xml, /<pageSetup pageOrder="overThenDown" orientation="landscape"\/>/);
  assert.doesNotMatch(xml, /scale=|fitToWidth=|fitToHeight=/);

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.pageSetup.orientation, 'landscape');
  assert.equal(back?.pageSetup.pageOrder, 'overThenDown');
});

test('paperSize round-trips and leads the <pageSetup> attributes', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.pageSetup.paperSize = 9;
  sheet.pageSetup.scale = 96;

  const xml = sheetXml(writeXlsx(wb));
  assert.match(xml, /<pageSetup paperSize="9" scale="96"\/>/);

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.pageSetup.paperSize, 9);
});

test('an opaque printer-settings blob round-trips its exact bytes', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  const blob = new Uint8Array([0x00, 0x01, 0xff, 0x7f, 0x80, 0x42]);
  sheet.pageSetup.printerSettings = blob;

  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual(back?.pageSetup.printerSettings, blob, 'the DEVMODE bytes survive verbatim');
});

test('a printer-settings blob wires up the r:id, the .bin part, its rel, and a content type', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').pageSetup.printerSettings = new Uint8Array([1, 2, 3]);

  const pkg = writeXlsx(wb);
  // The blob is the only reason the element exists, so <pageSetup> emits carrying just the r:id.
  assert.match(sheetXml(pkg), /<pageSetup r:id="rId1"\/>/);

  assert.deepEqual(
    partBytes(pkg, 'xl/printerSettings/printerSettings1.bin'),
    new Uint8Array([1, 2, 3]),
  );

  const rels = partText(pkg, 'xl/worksheets/_rels/sheet1.xml.rels');
  assert.match(rels, /Id="rId1"[^>]*Target="\.\.\/printerSettings\/printerSettings1\.bin"/);
  assert.match(rels, /Type="[^"]*\/printerSettings"/);

  const contentTypes = partText(pkg, '[Content_Types].xml');
  assert.match(contentTypes, /<Default Extension="bin" ContentType="[^"]*printerSettings"\/>/);
});

test('a printer-settings blob rides alongside a table without stealing its rel id', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'h';
  sheet.getCell('A2').value = 'v';
  sheet.addTable({name: 'T', ref: 'A1', columns: [{name: 'h'}], rowCount: 1});
  sheet.pageSetup.printerSettings = new Uint8Array([9]);

  const pkg = writeXlsx(wb);
  const rels = partText(pkg, 'xl/worksheets/_rels/sheet1.xml.rels');
  // The table keeps rId1; the printer-settings blob follows it at rId2, so neither reference collides.
  assert.match(rels, /Id="rId1"[^>]*Target="\.\.\/tables\/table1\.xml"/);
  assert.match(rels, /Id="rId2"[^>]*Target="\.\.\/printerSettings\/printerSettings1\.bin"/);
  assert.match(sheetXml(pkg), /<pageSetup r:id="rId2"\/>/);
});

test('a sheet with no printer settings writes no .bin part and no r:id', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').pageSetup.scale = 90;

  const pkg = writeXlsx(wb);
  assert.doesNotMatch(sheetXml(pkg), /r:id=/);
  assert.equal(optionalPartText(pkg, 'xl/printerSettings/printerSettings1.bin'), undefined);
  assert.doesNotMatch(partText(pkg, '[Content_Types].xml'), /Extension="bin"/);
});

test('a non-numeric paperSize is dropped on read, not stored as NaN', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').pageSetup.scale = 96;
  const patched = patchParts(writeXlsx(wb), {
    [SHEET1]: (xml) => xml.replace('<pageSetup ', '<pageSetup paperSize="A4" '),
  });
  const back = readXlsx(patched).getWorksheet('S');
  assert.equal(back?.pageSetup.paperSize, undefined);
  assert.equal(back?.pageSetup.scale, 96);
});

test('a sheet with no page setup emits neither <pageSetUpPr> nor <pageSetup>', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'y';

  const xml = sheetXml(writeXlsx(wb));
  assert.doesNotMatch(xml, /<pageSetUpPr/);
  assert.doesNotMatch(xml, /<pageSetup/);
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(back?.pageSetup.fitToPage, undefined);
  assert.equal(back?.pageSetup.scale, undefined);
});

test('a <pageSetUpPr> present only for other reasons leaves fitToPage unset', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'y';
  const patched = patchParts(writeXlsx(wb), {
    [SHEET1]: (xml) =>
      xml.replace('<dimension', '<sheetPr><pageSetUpPr autoPageBreaks="0"/></sheetPr><dimension'),
  });
  const back = readXlsx(patched).getWorksheet('S');
  assert.equal(back?.pageSetup.fitToPage, undefined);
});

test('page setup survives a worksheet model export/import', () => {
  const wb = new Workbook();
  const src = wb.addWorksheet('Src');
  src.pageSetup.fitToPage = true;
  src.pageSetup.scale = 75;
  const dst = wb.addWorksheet('Dst');
  dst.model = src.model;
  assert.equal(dst.pageSetup.fitToPage, true);
  assert.equal(dst.pageSetup.scale, 75);
});
