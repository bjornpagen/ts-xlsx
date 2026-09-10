import assert from 'node:assert/strict';
import {test} from 'node:test';

import {strFromU8, strToU8, unzipSync} from 'fflate';

import {INTERNAL, NAMED_STYLE_ID} from '../../core/internal.ts';
import {Workbook} from '../../core/workbook.ts';
import {
  captureIn,
  partIn,
  partsWritten as partsOf,
  patchParts,
  roundtrip,
  SHEET1,
} from './package.test-support.ts';
import {STYLES_PART} from './part-names.ts';
import {readXlsx} from './read.ts';
import {buildPackageParts, writeXlsx, writeXlsxAsync} from './write.ts';

test('a one-sheet workbook writes the full set of OPC parts', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const names = Object.keys(partsOf(wb)).sort();
  assert.deepEqual(names, [
    '[Content_Types].xml',
    '_rels/.rels',
    'docProps/app.xml',
    'docProps/core.xml',
    'xl/_rels/workbook.xml.rels',
    'xl/styles.xml',
    'xl/theme/theme1.xml',
    'xl/workbook.xml',
    'xl/worksheets/sheet1.xml',
  ]);
});

test('a workbook with no worksheets is refused rather than written corrupt', () => {
  assert.throws(() => writeXlsx(new Workbook()), /no worksheets/);
});

test('the content types and rels declare each worksheet consistently', () => {
  const wb = new Workbook();
  wb.addWorksheet('One').getCell('A1').value = 1;
  wb.addWorksheet('Two').getCell('A1').value = 2;
  const parts = partsOf(wb);
  for (const i of [1, 2]) {
    assert.match(
      partIn(parts, '[Content_Types].xml'),
      new RegExp(`/xl/worksheets/sheet${i}\\.xml`),
    );
    assert.match(
      partIn(parts, 'xl/_rels/workbook.xml.rels'),
      new RegExp(`worksheets/sheet${i}\\.xml`),
    );
    assert.ok(parts[`xl/worksheets/sheet${i}.xml`], `sheet${i}.xml part exists`);
  }
});

test('a default sheet is visible (no state attribute); explicit states are written', () => {
  const wb = new Workbook();
  wb.addWorksheet('Visible').getCell('A1').value = 'x';
  wb.addWorksheet('Hidden', {state: 'hidden'}).getCell('A1').value = 'x';
  const xml = partIn(partsOf(wb), 'xl/workbook.xml');
  assert.match(xml, /<sheet name="Visible" sheetId="1" r:id="rId1"\/>/);
  assert.match(xml, /<sheet name="Hidden" sheetId="2" state="hidden" r:id="rId2"\/>/);
});

test('cell values serialise by type with a computed dimension', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('B2').value = 42;
  s.getCell('C2').value = true;
  s.getCell('B3').value = 'hi';
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<dimension ref="B2:C3"\/>/);
  assert.match(xml, /<c r="B2"><v>42<\/v><\/c>/);
  assert.match(xml, /<c r="C2" t="b"><v>1<\/v><\/c>/);
  assert.match(xml, /<c r="B3" t="inlineStr"><is><t>hi<\/t><\/is><\/c>/);
});

test('XML-special characters in text and formulas are escaped', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'a < b & c > d';
  s.getCell('A2').value = {formula: 'IF(A1<B1,"x"&"y","")'};
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<t>a &lt; b &amp; c &gt; d<\/t>/);
  assert.match(xml, /<f>IF\(A1&lt;B1,"x"&amp;"y",""\)<\/f>/);
  // No raw ampersand survives except as the head of an entity: the check the corpus's
  // xmlWellFormed applies to reject unescaped specials.
  assert.doesNotMatch(xml, /&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/);
});

test('a formula supplied with a leading = is stored without it', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: '=1+2', result: 3};
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<f>1\+2<\/f>/);
  assert.doesNotMatch(xml, /<f>=/);
});

test('a string with edge whitespace carries xml:space="preserve"', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = '  padded  ';
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<t xml:space="preserve"> {2}padded {2}<\/t>/);
});

test('a non-finite number is written as a valueless cell, never a bare NaN/Infinity token', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = Number.POSITIVE_INFINITY;
  sheet.getCell('A2').value = Number.NaN;
  sheet.getCell('A3').value = 5; // a sibling finite cell must survive
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.doesNotMatch(
    xml,
    /<v>[^<]*(NaN|Infinity)[^<]*<\/v>/,
    'no bare non-finite token is emitted',
  );
  assert.match(xml, /<c r="A3"><v>5<\/v><\/c>/, 'a finite sibling cell is unaffected');
});

test('a formula whose cached result is non-finite keeps its formula but caches no value', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: '1/0', result: Number.POSITIVE_INFINITY};
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<c r="A1"><f>1\/0<\/f><\/c>/, 'the formula survives with no cached <v>');
  assert.doesNotMatch(xml, /Infinity/, 'no Infinity token leaks into the sheet');
});

test('a formula cell with a string result is typed t="str"', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = {formula: 'A2&A3', result: 'joined'};
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<c r="A1" t="str"><f>A2&amp;A3<\/f><v>joined<\/v><\/c>/);
});

test('every written package ships a default theme part', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const parts = partsOf(wb);
  assert.ok(parts['xl/theme/theme1.xml'], 'theme part present');
  assert.match(partIn(parts, '[Content_Types].xml'), /theme\+xml/);
});

test('a column width emits a <col> with customWidth', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.getColumn(2).width = 12;
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<cols><col min="2" max="2" width="12" customWidth="1"\/><\/cols>/);
});

test('adjacent equivalent columns coalesce into a single <col> span', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  for (let i = 1; i <= 4; i++) {
    const c = s.getColumn(i);
    c.width = 12;
    c.outlineLevel = 1;
  }
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<col min="1" max="4" width="12" customWidth="1" outlineLevel="1"\/>/);
  assert.equal(
    (xml.match(/<col\b/g) ?? []).length,
    1,
    'four equivalent columns collapse to one span',
  );
});

test('columns that differ are not coalesced', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getColumn(1).width = 12;
  s.getColumn(2).width = 20;
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<col min="1" max="1" width="12"/);
  assert.match(xml, /<col min="2" max="2" width="20"/);
});

test('a hidden column emits hidden="1" and needs no width', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.getColumn(3).hidden = true;
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<col min="3" max="3" hidden="1"\/>/);
});

test('column outline grouping serializes onto the <col>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  const c = s.getColumn(2);
  c.outlineLevel = 1;
  c.collapsed = true;
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<col min="2" max="2" outlineLevel="1" collapsed="1"\/>/);
});

test('an ungrouped column emits no outline attributes', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.getColumn(2).width = 10;
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.doesNotMatch(xml, /outlineLevel|collapsed/);
});

test('the last column serializes, and one past the limit never reaches the model to be written', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.getColumn(16384).width = 10;
  assert.throws(() => s.getColumn(16385), RangeError, 'the writer has no out-of-grid case to drop');
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /min="16384" max="16384"/);
  assert.doesNotMatch(xml, /16385/);
});

test('a sheet with no column definitions emits no <cols>', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.doesNotMatch(xml, /<cols>/);
});

test('row height and outline flags serialize onto the <row>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A2').value = 'd';
  const r = s.getRow(2);
  r.height = 30;
  r.hidden = true;
  r.outlineLevel = 1;
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<row r="2" ht="30" customHeight="1" hidden="1" outlineLevel="1">/);
});

test('a row carrying only metadata is emitted with no cells', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.getRow(5).hidden = true;
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<row r="5" hidden="1"><\/row>/);
});

test('a formatted-but-empty cell is emitted as a styled <c> with no value, not dropped', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  // B2 carries a fill but no value: a real formatted blank Excel keeps, not a cell to discard.
  s.getCell('B2').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF00FF00'}};
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(
    xml,
    /<c r="B2" s="\d+"\/>/,
    'the styled empty cell is written with its style and no <v>',
  );
  assert.doesNotMatch(
    xml,
    /<c r="B2"[^>]*>.*<\/c>/,
    'a valueless styled cell carries no child content',
  );
});

test('an empty cell with no style of its own is not serialised', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  // Touching a cell without giving it a value or style must not fabricate a <c> for it.
  s.getCell('C3');
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.doesNotMatch(xml, /r="C3"/, 'a value-less, style-less cell contributes nothing');
});

test('a collapsed flag is emitted only where set, not on sibling rows', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getRow(2).outlineLevel = 1;
  s.getRow(3).collapsed = true;
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<row r="2" outlineLevel="1">/);
  assert.match(xml, /<row r="3" collapsed="1">/);
});

test('a fully-hidden outline group derives the collapsed toggle onto its summary row', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'top';
  for (const r of [2, 3, 4]) {
    s.getCell(`A${r}`).value = `d${r}`;
    s.getRow(r).outlineLevel = 1;
    s.getRow(r).hidden = true;
  }
  s.getCell('A5').value = 'summary';
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  // Summary-below is Excel's default: the summary row terminates the group from beneath and carries
  // the collapse toggle, so the outline expands in a single click.
  assert.match(
    xml,
    /<row r="5"[^>]*\bcollapsed="1"/,
    'the summary row carries the collapse toggle',
  );
  // The hidden detail rows themselves never carry it.
  assert.doesNotMatch(xml, /<row r="2"[^>]*collapsed/);
  assert.doesNotMatch(xml, /<row r="4"[^>]*collapsed/);
});

test('a partially-visible outline group derives no collapsed summary', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A2').value = 'd2';
  s.getCell('A3').value = 'd3';
  s.getCell('A4').value = 'summary';
  s.getRow(2).outlineLevel = 1;
  s.getRow(2).hidden = true;
  // Row 3 is grouped but visible: the group is expanded, so nothing is collapsed.
  s.getRow(3).outlineLevel = 1;
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.doesNotMatch(xml, /collapsed/);
});

test('with summary-above outlines, the collapse toggle derives onto the row above the group', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.outline.summaryBelow = false;
  s.getCell('A1').value = 'summary';
  for (const r of [2, 3, 4]) {
    s.getCell(`A${r}`).value = `d${r}`;
    s.getRow(r).outlineLevel = 1;
    s.getRow(r).hidden = true;
  }
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<row r="1"[^>]*\bcollapsed="1"/, 'the summary sits above its detail group');
});

test('sheet default row height and column width land on <sheetFormatPr>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.properties.defaultRowHeight = 30;
  s.properties.defaultColWidth = 20;
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(
    xml,
    /<sheetFormatPr defaultRowHeight="30" defaultColWidth="20" customHeight="1"\/>/,
  );
});

test('an unset default row height falls back to 15 with no customHeight', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<sheetFormatPr defaultRowHeight="15"\/>/);
});

test('grouped columns report their depth as outlineLevelCol', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.getColumn(2).outlineLevel = 1;
  s.getColumn(3).outlineLevel = 2;
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<sheetFormatPr [^>]*\boutlineLevelCol="2"/);
  assert.doesNotMatch(xml, /outlineLevelRow/);
});

test('grouped rows report their depth as outlineLevelRow', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'summary';
  s.getCell('A2').value = 'detail';
  s.getRow(2).outlineLevel = 1;
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<sheetFormatPr [^>]*\boutlineLevelRow="1"/);
  assert.doesNotMatch(xml, /outlineLevelCol/);
});

test('an ungrouped sheet emits neither outline level', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.doesNotMatch(xml, /outlineLevel/);
});

test('the workbook always declares a window view, defaulting to the first tab', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const xml = partIn(partsOf(wb), 'xl/workbook.xml');
  assert.match(
    xml,
    /<bookViews><workbookView xWindow="-110" yWindow="-110" windowWidth="19420" windowHeight="12220"\/><\/bookViews>/,
  );
  // CT_Workbook order: <bookViews> precedes <sheets>.
  assert.ok(xml.indexOf('<bookViews>') < xml.indexOf('<sheets>'), 'bookViews must precede sheets');
});

test('a customised window view is written verbatim and survives a round-trip', () => {
  const wb = new Workbook();
  wb.addWorksheet('One').getCell('A1').value = 1;
  wb.addWorksheet('Two').getCell('A1').value = 2;
  wb.view.x = 240;
  wb.view.y = 120;
  wb.view.width = 15000;
  wb.view.height = 9000;
  wb.view.activeTab = 1;
  const xml = partIn(partsOf(wb), 'xl/workbook.xml');
  assert.match(
    xml,
    /<workbookView xWindow="240" yWindow="120" windowWidth="15000" windowHeight="9000" activeTab="1"\/>/,
  );
  const reread = roundtrip(wb);
  assert.deepEqual(reread.view, {x: 240, y: 120, width: 15000, height: 9000, activeTab: 1});
});

test('exactly one sheet is marked selected, and it is the active tab', () => {
  const wb = new Workbook();
  for (const name of ['One', 'Two', 'Three']) wb.addWorksheet(name).getCell('A1').value = name;
  wb.view.activeTab = 2;
  const parts = partsOf(wb);
  const selected = [1, 2, 3].filter((n) =>
    partIn(parts, `xl/worksheets/sheet${n}.xml`).includes('tabSelected="1"'),
  );
  assert.deepEqual(selected, [3], 'only the active sheet carries tabSelected');
});

test('an out-of-range active tab still selects a sheet rather than none', () => {
  const wb = new Workbook();
  wb.addWorksheet('Only').getCell('A1').value = 'x';
  wb.view.activeTab = 7;
  const parts = partsOf(wb);
  assert.match(partIn(parts, 'xl/worksheets/sheet1.xml'), /<sheetView tabSelected="1"/);
  // The written activeTab is clamped with it, so the two can never disagree.
  assert.doesNotMatch(partIn(parts, 'xl/workbook.xml'), /activeTab/);
});

test('a frozen sheet carries tabSelected alongside its pane', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'header';
  s.freeze(1);
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1"/);
});

test('title and company are written to their two different parts, and round-trip', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  wb.properties.title = 'Planning Ateliers';
  wb.properties.creator = 'A. Author';
  wb.properties.company = 'Acme & Co';
  const parts = partsOf(wb);

  const core = partIn(parts, 'docProps/core.xml');
  // cp:coreProperties is a sequence, so the order is load-bearing, not incidental.
  assert.match(core, /<dc:title>Planning Ateliers<\/dc:title><dc:creator>A\. Author<\/dc:creator>/);
  assert.match(
    partIn(parts, 'docProps/app.xml'),
    /<Application>ts-xlsx<\/Application><Company>Acme &amp; Co<\/Company>/,
  );

  const reopened = roundtrip(wb);
  assert.equal(reopened.properties.title, 'Planning Ateliers');
  assert.equal(reopened.properties.company, 'Acme & Co');
});

test('a workbook naming neither writes neither element', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const parts = partsOf(wb);
  assert.doesNotMatch(partIn(parts, 'docProps/core.xml'), /dc:title/);
  assert.doesNotMatch(partIn(parts, 'docProps/app.xml'), /Company/);
});

test('hiding the grid emits showGridLines="0", and only when asked', () => {
  const wb = new Workbook();
  const plain = wb.addWorksheet('Plain');
  plain.getCell('A1').value = 'x';
  const hidden = wb.addWorksheet('Hidden');
  hidden.getCell('A1').value = 'x';
  hidden.view.showGridLines = false;
  const parts = partsOf(wb);

  assert.doesNotMatch(
    partIn(parts, 'xl/worksheets/sheet1.xml'),
    /showGridLines/,
    'a sheet that never mentioned gridlines writes no attribute',
  );
  assert.match(partIn(parts, 'xl/worksheets/sheet2.xml'), /<sheetView showGridLines="0"/);
});

test('a hidden grid survives a frozen pane, and round-trips', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'header';
  s.freeze(1);
  s.view.showGridLines = false;
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  // The frozen arm is a separate return in sheetViewsXml, so it gets its own assertion.
  assert.match(
    xml,
    /<sheetView showGridLines="0" tabSelected="1" workbookViewId="0"><pane ySplit="1"/,
  );

  const reopened = roundtrip(wb);
  assert.equal(reopened.requireWorksheet('S').view.showGridLines, false);
  assert.equal(reopened.requireWorksheet('S').view.state, 'frozen');
});

test('a visible grid reads back unset rather than true', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const reopened = roundtrip(wb);
  // Excel's default is on, so "unset" and "on" are the same state: recording `true` would make a
  // re-write fabricate an attribute the source never carried.
  assert.equal(reopened.requireWorksheet('S').view.showGridLines, undefined);
});

test('setting a subset of margins emits all six pageMargins attributes', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.pageMargins.left = 0.1;
  s.pageMargins.right = 0.1;
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
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
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.doesNotMatch(xml, /<pageMargins/);
});

test('<pageMargins> is placed after <sheetData>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.pageMargins.top = 1;
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.ok(
    xml.indexOf('<sheetData') < xml.indexOf('<pageMargins'),
    'pageMargins must follow sheetData',
  );
});

test('print-option toggles are emitted and survive a write→read round-trip', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.printOptions.horizontalCentered = true;
  s.printOptions.gridLines = true;
  s.printOptions.headings = true;

  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
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

  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<printOptions gridLinesSet="0"\/>/);

  const back = roundtrip(wb).getWorksheet('S');
  assert.deepEqual(
    back?.printOptions,
    {gridLinesSet: false},
    'the explicit false survives, not swallowed by the default',
  );
});

test('<printOptions> precedes <pageMargins> in schema order', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.printOptions.horizontalCentered = true;
  s.pageMargins.top = 1;
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.ok(
    xml.indexOf('<printOptions') < xml.indexOf('<pageMargins'),
    'printOptions must precede pageMargins',
  );
});

test('a sheet with no print options set emits no <printOptions> element', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
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
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
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
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<headerFooter><oddHeader>H<\/oddHeader><\/headerFooter>/);
  assert.doesNotMatch(xml, /different/);
});

test('a sheet with no header/footer emits no <headerFooter>', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.doesNotMatch(xml, /<headerFooter/);
});

test('header/footer text is XML-escaped and placed after <pageMargins>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.pageMargins.top = 1;
  s.headerFooter.oddHeader = 'a & b < c';
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
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
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<oddHeader>&amp;C\[_x0001_\]\[_x005F_x0041_\]<\/oddHeader>/);
});

test('an astral character in a header needs no escape, a control one does', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.headerFooter.oddFooter = '\u{1F600}\u0001';
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
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

test('<cols> is placed after <sheetFormatPr> and before <sheetData>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.getColumn(1).width = 8;
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  const fmt = xml.indexOf('<sheetFormatPr');
  const cols = xml.indexOf('<cols>');
  const data = xml.indexOf('<sheetData>');
  assert.ok(
    fmt < cols && cols < data,
    `expected sheetFormatPr < cols < sheetData, got ${fmt},${cols},${data}`,
  );
});

test('an empty-body table refs the full header row and writes a table part', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({
    name: 'T1',
    ref: 'A1',
    columns: [{name: 'Alpha'}, {name: 'Beta'}],
    rowCount: 0,
  });
  const parts = partsOf(wb);
  const table = partIn(parts, 'xl/tables/table1.xml');
  assert.match(table, /ref="A1:B1"/);
  assert.match(table, /<tableColumns count="2">/);
  assert.match(table, /<autoFilter ref="A1:B1"\/>/);
  assert.match(partIn(parts, '[Content_Types].xml'), /\/xl\/tables\/table1\.xml/);
  assert.match(
    partIn(parts, 'xl/worksheets/_rels/sheet1.xml.rels'),
    /Target="\.\.\/tables\/table1\.xml"/,
  );
  assert.match(
    partIn(parts, 'xl/worksheets/sheet1.xml'),
    /<tableParts count="1"><tablePart r:id="rId1"\/><\/tableParts>/,
  );
});

test('a data row extends the table ref by one row', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({
    name: 'T',
    ref: 'A1',
    columns: [{name: 'A'}, {name: 'B'}],
    rowCount: 1,
  });
  const table = partIn(partsOf(wb), 'xl/tables/table1.xml');
  assert.match(table, /ref="A1:B2"/);
});

test('a headerless table sets headerRowCount="0" and emits no autoFilter', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({
    name: 'H',
    ref: 'A1',
    columns: [{name: 'A'}, {name: 'B'}],
    rowCount: 2,
    headerRow: false,
  });
  const table = partIn(partsOf(wb), 'xl/tables/table1.xml');
  assert.match(table, /headerRowCount="0"/);
  assert.doesNotMatch(table, /<autoFilter/);
});

test('a totals-row column serialises its function and keeps every column', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({
    name: 'T',
    ref: 'A1',
    columns: [
      {name: 'Item', totalsRowLabel: 'Total'},
      {name: 'Amount', totalsRowFunction: 'sum'},
    ],
    rowCount: 2,
    totalsRow: true,
  });
  const table = partIn(partsOf(wb), 'xl/tables/table1.xml');
  assert.match(table, /ref="A1:B4"/);
  assert.match(table, /totalsRowCount="1"/);
  assert.match(table, /<tableColumn id="1" name="Item" totalsRowLabel="Total"\/>/);
  assert.match(table, /<tableColumn id="2" name="Amount" totalsRowFunction="sum"\/>/);
});

test('a no-totals table omits totalsRowShown unless the flag is set explicitly', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({name: 'T', ref: 'A1', columns: [{name: 'A'}], rowCount: 1});
  const table = partIn(partsOf(wb), 'xl/tables/table1.xml');
  assert.doesNotMatch(
    table,
    /totalsRowShown/,
    'an unset flag emits no attribute: Excel must not see a spurious one',
  );
});

test('an explicit totalsRowShown flag round-trips as "0" or "1"', () => {
  const off = new Workbook();
  off
    .addWorksheet('S')
    .addTable({name: 'T', ref: 'A1', columns: [{name: 'A'}], rowCount: 1, totalsRowShown: false});
  assert.match(partIn(partsOf(off), 'xl/tables/table1.xml'), /totalsRowShown="0"/);

  const on = new Workbook();
  on.addWorksheet('S').addTable({
    name: 'T',
    ref: 'A1',
    columns: [{name: 'A'}],
    rowCount: 1,
    totalsRowShown: true,
  });
  assert.match(partIn(partsOf(on), 'xl/tables/table1.xml'), /totalsRowShown="1"/);
});

test("a table with no explicit style is written with Excel's default TableStyleMedium2", () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({name: 'T', ref: 'A1', columns: [{name: 'A'}], rowCount: 1});
  const table = partIn(partsOf(wb), 'xl/tables/table1.xml');
  assert.match(table, /<tableStyleInfo name="TableStyleMedium2"[^>]*showRowStripes="1"[^>]*\/>/);
});

test('an explicit table style is emitted verbatim, omitting the attributes it leaves unset', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({
    name: 'T',
    ref: 'A1',
    columns: [{name: 'A'}],
    rowCount: 1,
    style: {name: 'Assignment schedule', showRowStripes: false},
  });
  const table = partIn(partsOf(wb), 'xl/tables/table1.xml');
  assert.match(
    table,
    /name="Assignment schedule"/,
    'a custom style name survives instead of being rewritten',
  );
  assert.match(
    table,
    /showRowStripes="0"/,
    'the source\'s stripe choice is preserved, not forced to "1"',
  );
  assert.doesNotMatch(table, /showFirstColumn/, 'an unset banding flag emits no attribute');
});

test('an illegal table name is rejected at definition time', () => {
  const s = new Workbook().addWorksheet('S');
  assert.throws(
    () => s.addTable({name: "Bob's Accounts", ref: 'A1', columns: [{name: 'A'}], rowCount: 1}),
    /identifier/,
  );
  assert.throws(
    () => s.addTable({name: '1Digit', ref: 'A1', columns: [{name: 'A'}], rowCount: 1}),
    /identifier/,
  );
  assert.throws(
    () => s.addTable({name: 'test-name', ref: 'A1', columns: [{name: 'A'}], rowCount: 1}),
    /identifier/,
  );
});

test('a valid identifier table name is written verbatim', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({
    name: 'Valid_Name',
    ref: 'A1',
    columns: [{name: 'A'}],
    rowCount: 1,
  });
  const table = partIn(partsOf(wb), 'xl/tables/table1.xml');
  assert.match(table, /name="Valid_Name"/);
  assert.match(table, /displayName="Valid_Name"/);
});

test('tables are numbered globally across sheets with sheet-local rel ids', () => {
  const wb = new Workbook();
  wb.addWorksheet('One').addTable({name: 'Ta', ref: 'A1', columns: [{name: 'A'}], rowCount: 1});
  wb.addWorksheet('Two').addTable({name: 'Tb', ref: 'A1', columns: [{name: 'A'}], rowCount: 1});
  const parts = partsOf(wb);
  assert.ok(parts['xl/tables/table1.xml'], 'first table part');
  assert.ok(parts['xl/tables/table2.xml'], 'second table part (globally numbered)');
  assert.match(
    partIn(parts, 'xl/worksheets/_rels/sheet2.xml.rels'),
    /Target="\.\.\/tables\/table2\.xml"/,
  );
});

test('a merge overlapping a table is rejected; a disjoint merge is written', () => {
  const overlap = new Workbook();
  const s1 = overlap.addWorksheet('S');
  s1.addTable({name: 'T', ref: 'A1', columns: [{name: 'A'}, {name: 'B'}], rowCount: 2});
  s1.mergeCells('A2:B2');
  assert.throws(() => writeXlsx(overlap), /overlaps table/);

  const disjoint = new Workbook();
  const s2 = disjoint.addWorksheet('S');
  s2.addTable({name: 'T', ref: 'A1', columns: [{name: 'A'}, {name: 'B'}], rowCount: 2});
  s2.mergeCells('D5:E5');
  const xml = partIn(partsOf(disjoint), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<mergeCells count="1"><mergeCell ref="D5:E5"\/><\/mergeCells>/);
});

test('<tableParts> follows <headerFooter> in the worksheet element order', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.headerFooter.oddHeader = 'H';
  s.addTable({name: 'T', ref: 'A1', columns: [{name: 'A'}], rowCount: 1});
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.ok(
    xml.indexOf('<headerFooter') < xml.indexOf('<tableParts'),
    'tableParts must follow headerFooter per CT_Worksheet',
  );
});

test('an unprotected sheet emits no <sheetProtection> element', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.doesNotMatch(xml, /<sheetProtection/);
});

test('protecting a sheet emits a self-closing <sheetProtection sheet="1"> after <sheetData>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.protect();
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<sheetProtection sheet="1"\/>/);
  assert.ok(
    xml.indexOf('</sheetData>') < xml.indexOf('<sheetProtection'),
    'sheetProtection must follow sheetData per CT_Worksheet',
  );
});

test('an unprotected password derives an OOXML-agile credential onto <sheetProtection>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.protect('secret');
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /algorithmName="SHA-512"/);
  assert.match(xml, /hashValue="[^"]+"/);
  assert.match(xml, /saltValue="[^"]+"/);
  assert.match(xml, /spinCount="100000"/);
  assert.match(xml, /sheet="1"/);
});

test('protection flags invert to OOXML "forbidden" booleans, writing only non-default values', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  // sort defaults to forbidden under protection, so allowing it must emit sort="0".
  // selectLockedCells defaults to permitted, so forbidding it must emit selectLockedCells="1".
  s.protect(undefined, {sort: true, autoFilter: true, selectLockedCells: false});
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /sort="0"/);
  assert.match(xml, /autoFilter="0"/);
  assert.match(xml, /selectLockedCells="1"/);
});

test('a flag left at its OOXML default is omitted from <sheetProtection>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  // Allowing selection (its default) and forbidding sort (its default) are both no-ops.
  s.protect(undefined, {selectLockedCells: true, sort: false});
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  const tag = (xml.match(/<sheetProtection[^>]*\/>/) as RegExpMatchArray)[0];
  assert.doesNotMatch(tag, /selectLockedCells=/);
  assert.doesNotMatch(tag, /sort=/);
  assert.match(tag, /sheet="1"/);
});

test('unprotect() removes a sheet-protection element previously set', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.protect('pw');
  s.unprotect();
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.doesNotMatch(xml, /<sheetProtection/);
});

test('a sheet autofilter emits an <autoFilter> element after <sheetProtection>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.protect('pw');
  s.autoFilter = 'A1:C10';
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<autoFilter ref="A1:C10"\/>/);
  assert.ok(
    xml.indexOf('<sheetProtection') < xml.indexOf('<autoFilter'),
    'autoFilter follows sheetProtection in CT_Worksheet order',
  );
});

test('a sheet with no autofilter emits no <autoFilter> and no _FilterDatabase', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const parts = partsOf(wb);
  assert.doesNotMatch(partIn(parts, 'xl/worksheets/sheet1.xml'), /<autoFilter/);
  assert.doesNotMatch(partIn(parts, 'xl/workbook.xml'), /_FilterDatabase/);
});

test('a sheet autofilter generates a hidden, sheet-scoped _FilterDatabase built-in', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.autoFilter = 'A1:C10';
  const xml = partIn(partsOf(wb), 'xl/workbook.xml');
  assert.match(
    xml,
    /<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">S!\$A\$1:\$C\$10<\/definedName>/,
  );
});

test('a _FilterDatabase quotes a sheet name that needs it and uses the sheet 0-based index', () => {
  const wb = new Workbook();
  wb.addWorksheet('First').getCell('A1').value = 'x';
  const second = wb.addWorksheet('Sales 2024');
  second.getCell('A1').value = 'y';
  second.autoFilter = 'A1:B5';
  const xml = partIn(partsOf(wb), 'xl/workbook.xml');
  assert.match(
    xml,
    /<definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">'Sales 2024'!\$A\$1:\$B\$5<\/definedName>/,
  );
});

test('a sheet autofilter round-trips through write then read, and drops no user defined name', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.autoFilter = 'A1:C10';
  wb.defineName({name: 'TaxRate', refersTo: 'S!$A$1'});

  const back = roundtrip(wb);
  const sheet = back.getWorksheet('S');
  assert.ok(sheet !== undefined);
  assert.deepEqual(
    sheet.autoFilter,
    {ref: 'A1:C10', columns: []},
    'the filter range survives the round-trip',
  );
  // The system-generated _FilterDatabase is reconstructed from the sheet, never surfaced as a name…
  assert.deepEqual(
    back.definedNames.map((n) => n.name),
    ['TaxRate'],
    'only the user name is exposed; _FilterDatabase is filtered out',
  );
  // …and re-writing does not accumulate a duplicate.
  const rewritten = partIn(partsOf(back), 'xl/workbook.xml');
  assert.equal(
    (rewritten.match(/_FilterDatabase/g) ?? []).length,
    1,
    'exactly one _FilterDatabase after a re-write',
  );
});

test('a criteria-bearing autofilter nests <filterColumn> children under the <autoFilter>', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.autoFilter = {
    ref: 'A1:B4',
    columns: [
      {colId: 0, criteria: {kind: 'values', values: ['apple', 'pear'], blank: false}},
      {
        colId: 1,
        criteria: {kind: 'custom', and: false, predicates: [{operator: 'greaterThan', val: '6'}]},
      },
    ],
  };
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.match(
    xml,
    /<autoFilter ref="A1:B4"><filterColumn colId="0"><filters><filter val="apple"\/><filter val="pear"\/><\/filters><\/filterColumn><filterColumn colId="1"><customFilters><customFilter operator="greaterThan" val="6"\/><\/customFilters><\/filterColumn><\/autoFilter>/,
  );
});

test('a values filter and a custom filter both survive a write→read round-trip', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.getCell('A1').value = 'x';
  s.autoFilter = {
    ref: 'A1:C10',
    columns: [
      {colId: 0, criteria: {kind: 'values', values: ['red', 'blue'], blank: true}},
      {
        colId: 2,
        criteria: {
          kind: 'custom',
          and: true,
          predicates: [
            {operator: 'greaterThanOrEqual', val: '1'},
            {operator: 'lessThan', val: '9'},
          ],
        },
      },
    ],
  };

  const sheet = roundtrip(wb).getWorksheet('S');
  assert.ok(sheet !== undefined);
  assert.deepEqual(sheet.autoFilter, {
    ref: 'A1:C10',
    columns: [
      {colId: 0, criteria: {kind: 'values', values: ['red', 'blue'], blank: true}},
      {
        colId: 2,
        criteria: {
          kind: 'custom',
          and: true,
          predicates: [
            {operator: 'greaterThanOrEqual', val: '1'},
            {operator: 'lessThan', val: '9'},
          ],
        },
      },
    ],
  });
});

test('a filter column addressing a column outside the range is dropped on read, not thrown', () => {
  const wb = new Workbook();
  const s = wb.addWorksheet('S');
  s.autoFilter = 'A1:B4';

  // Hand-forge an out-of-range <filterColumn colId="5"> onto the worksheet part (the kind of thing
  // a corrupt producer emits) and re-zip. Load-repair must keep the range and drop the bad column,
  // never throwing (the strict setter that authors go through would reject the same colId).
  const patched = patchParts(writeXlsx(wb), {
    [SHEET1]: (xml) =>
      xml.replace(
        '<autoFilter ref="A1:B4"/>',
        '<autoFilter ref="A1:B4"><filterColumn colId="5"><filters><filter val="z"/></filters></filterColumn></autoFilter>',
      ),
  });

  const sheet = readXlsx(patched).getWorksheet('S');
  assert.ok(sheet !== undefined);
  assert.deepEqual(
    sheet.autoFilter,
    {ref: 'A1:B4', columns: []},
    'the range survives, the bad column drops',
  );
});

test('a quote-prefixed cell emits quotePrefix on its xf and survives a round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  const cell = sheet.getCell('A1');
  cell.value = '=not-a-formula';
  cell.quotePrefix = true;

  const parts = partsOf(wb);
  assert.match(
    partIn(parts, 'xl/styles.xml'),
    /<xf [^>]*quotePrefix="1"/,
    'the cell-format record carries quotePrefix="1"',
  );

  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(
    back?.getCell('A1').quotePrefix,
    true,
    'the quote-prefix flag survives read/modify/write',
  );
  assert.equal(back?.getCell('A1').value, '=not-a-formula', 'the literal content is preserved');
});

test('a cell with no quote-prefix flag does not gain one on read', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'plain';
  const back = roundtrip(wb).getWorksheet('S');
  assert.equal(
    back?.getCell('A1').quotePrefix,
    undefined,
    'an ordinary cell reports no quote-prefix flag',
  );
});

test('a cell linking to a named cell style keeps its fill and xfId link across a round-trip', () => {
  const wb = new Workbook();
  // A yellow fill supplied only through the "Accent" named style; the cell's direct format is empty.
  wb[INTERNAL].restoreNamedStyles([
    {name: 'Normal', builtinId: 0},
    {name: 'Accent', fill: {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFFFF00'}}},
  ]);
  const sheet = wb.addWorksheet('S');
  const cell = sheet.getCell('A1');
  cell.value = 'x';
  cell[NAMED_STYLE_ID] = 1;

  const styles = partIn(partsOf(wb), 'xl/styles.xml');
  assert.match(styles, /<cellStyleXfs count="2"/, 'the named-style layer is emitted');

  const back = roundtrip(wb).getWorksheet('S');
  const a1 = back?.getCell('A1');
  const fill = a1?.fill;
  assert.ok(
    fill?.type === 'pattern',
    'the named-style fill resolves onto the cell as a pattern fill',
  );
  assert.equal(fill.pattern, 'solid', 'the named-style fill resolves onto the cell on read');
  assert.equal(fill.fgColor?.argb, 'FFFFFF00', 'the resolved fill is the named-style yellow');
  assert.equal(a1?.[NAMED_STYLE_ID], 1, 'the cell keeps its link to the named style');
});

test('manual row breaks are emitted as <rowBreaks> and round-trip', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  sheet.rowBreaks.push({id: 3, max: 16383, man: true}, {id: 6, max: 16383, man: true});

  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
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

test('a sheet with no manual row breaks emits no <rowBreaks> element', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
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

  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
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

  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
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
  const xml = partIn(partsOf(wb), 'xl/worksheets/sheet1.xml');
  assert.doesNotMatch(xml, /<colBreaks/, 'an empty column-break list fabricates nothing');
});

test('the async writer produces the same package the sync one does', async () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  // Big enough that the deflater's level actually changes the output size; below roughly 200 rows
  // every level compresses this identically, and the size assertion below could not fail.
  for (let row = 1; row <= 300; row++) {
    sheet.getCell(`A${row}`).value = `label ${row % 37} value`;
    sheet.getCell(`B${row}`).value = row * 7;
  }
  sheet.getCell('C1').value = {formula: 'SUM(B1:B300)', result: 315_000};
  sheet.getRow(1).height = 20;
  wb.addWorksheet('T').getCell('A1').value = true;

  const [sync, async] = [writeXlsx(wb), await writeXlsxAsync(wb)];
  const [inflatedSync, inflatedAsync] = [unzipSync(sync), unzipSync(async)];

  assert.deepEqual(
    Object.keys(inflatedAsync).sort(),
    Object.keys(inflatedSync).sort(),
    'both paths emit the same part set',
  );
  for (const [name, bytes] of Object.entries(inflatedSync)) {
    assert.deepEqual(
      strFromU8(inflatedAsync[name] as Uint8Array),
      strFromU8(bytes),
      `${name} is identical across the two writers`,
    );
  }
  // The archives themselves, not just their contents: with entry timestamps pinned there is nothing
  // left in a package that the two paths could legitimately disagree about, so this also pins them
  // to the same compression settings, which comparing inflated content alone would never notice.
  assert.deepEqual(async, sync, 'the two writers produce the same archive byte for byte');
});

test('the async writer honours WriteOptions the same way', async () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'repeated';
  sheet.getCell('A2').value = 'repeated';

  const pooled = unzipSync(await writeXlsxAsync(wb, {useSharedStrings: true}));
  const inline = unzipSync(await writeXlsxAsync(wb));
  assert.ok(pooled['xl/sharedStrings.xml'], 'pooling emits the shared-strings part');
  assert.equal(inline['xl/sharedStrings.xml'], undefined, 'the default keeps strings inline');
});

test('a package written asynchronously reads back through readXlsx', async () => {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'round trip';

  const back = readXlsx(await writeXlsxAsync(wb));
  assert.equal(back.getWorksheet('S')?.getCell('A1').value, 'round trip');
});

test('the async writer rejects rather than throwing where the sync one throws', async () => {
  const call = writeXlsxAsync(new Workbook());
  assert.ok(call instanceof Promise, 'part-building failures surface as a rejection, not a throw');
  await assert.rejects(call, /no worksheets/);
});

test('the package part map has no prototype, so a part named __proto__ is a part', () => {
  // A preserved part can keep its original zip entry name, which comes from an untrusted package.
  // Assigned onto a plain object, `__proto__` silently drops the part and re-points the map's
  // prototype at its bytes; `inflate.ts` and `edit-vba.ts` both defend against this on the read side.
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 'x';
  const parts = buildPackageParts(wb);

  assert.equal(Object.getPrototypeOf(parts), null, 'the map is null-prototype');
  assert.ok(Object.keys(parts).length > 0);
});

test('two package parts claiming one path is refused rather than silently overwritten', () => {
  // The paths twenty write sites choose come from several independent numberings, and when one was
  // wrong the later write simply won: the earlier part vanished from the zip while the content types
  // still declared both, which is a package Excel repairs. The refusal is what checks every one of
  // those numberings at once instead of trusting each.
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.getCell('A1').value = 'x';
  // A preserved part planted on a path the writer generates for itself: the shape the pivot bug took.
  wb[INTERNAL].addPreservedReference({
    relType:
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition',
    entryPath: STYLES_PART,
    parts: [
      {
        path: STYLES_PART,
        contentType: 'application/xml',
        bytes: strToU8('<styleSheet/>'),
        rels: [],
      },
    ],
  });

  assert.throws(() => buildPackageParts(wb), {
    name: 'InternalError',
    message: /two package parts claim the path "xl\/styles\.xml"/,
  });
});

// `<workbookPr>` and `<sheetPr>` say two things the model preserves rather than derives, and both
// have a default worth staying silent about. What these pin is the silence: a workbook that never
// mentioned a date system must not come back carrying `date1904="0"`, since that is an attribute in
// every package Excel writes without one.
test('the default date system and absent code names write no element at all', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').getCell('A1').value = 1;
  const parts = partsOf(workbook);
  assert.doesNotMatch(partIn(parts, 'xl/workbook.xml'), /<workbookPr/);
  assert.doesNotMatch(partIn(parts, 'xl/worksheets/sheet1.xml'), /<sheetPr/);
});

test('the 1904 date system and the code names are written as the attributes that carry them', () => {
  const workbook = new Workbook();
  workbook.dateEpoch = 1904;
  workbook.codeName = 'ThisWorkbook';
  const sheet = workbook.addWorksheet('S');
  sheet.codeName = 'Sheet1';
  sheet.getCell('A1').value = new Date(Date.UTC(2023, 2, 15));
  const parts = partsOf(workbook);

  assert.match(
    partIn(parts, 'xl/workbook.xml'),
    /<workbookPr date1904="1" codeName="ThisWorkbook"\/>/,
  );
  // A code name with nothing else to say still needs its element, and Excel writes that one
  // self-closing; a sheet carrying children keeps them.
  assert.match(partIn(parts, 'xl/worksheets/sheet1.xml'), /<sheetPr codeName="Sheet1"\/>/);
  // 43538 is Excel's own serial for 2023-03-15 under the 1904 system, 1462 below the 1900 one's.
  assert.match(partIn(parts, 'xl/worksheets/sheet1.xml'), /<v>43538<\/v>/);
});

test('a code name rides alongside what else <sheetPr> carries, rather than replacing it', () => {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('S');
  sheet.codeName = 'Sheet1';
  sheet.tabColor = {argb: 'FFFF0000'};
  sheet.getCell('A1').value = 1;
  assert.match(
    partIn(partsOf(workbook), 'xl/worksheets/sheet1.xml'),
    /<sheetPr codeName="Sheet1"><tabColor rgb="FFFF0000"\/><\/sheetPr>/,
  );
});
