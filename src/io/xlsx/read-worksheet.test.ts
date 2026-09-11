import assert from 'node:assert/strict';
import {test} from 'node:test';

import {MAX_COLUMN, MAX_ROW} from '../../core/address.ts';
import {Workbook} from '../../core/workbook.ts';
import {parseXmlPasses} from '../../xml/xml-read.ts';
import {worksheetPass} from './read-worksheet.ts';
import {sheetViewsXml} from './sheet-properties.ts';

// `customWidth`/`customHeight` are xsd:booleans, so a foreign producer may spell false either way.
// Excel writes the digit, which is why the long spelling went unnoticed for so long.
//
// The body pass alone, which is what every assertion below is about: none of them involves a second
// reader of the part, so running the four the package read composes it with would add nothing to
// check. `self-closing-capture.test.ts` is where the composition itself is under test.
function sheet(body: string) {
  const worksheet = new Workbook().addWorksheet('S');
  parseXmlPasses(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      body +
      '</worksheet>',
    [worksheetPass(worksheet, [], [], 1900, new Set())],
  );
  return worksheet;
}

// ── Custom views ─────────────────────────────────────────────────────────────────────────────────────
// `CT_CustomSheetView` repeats the sheet's own pane, breaks, margins, print options, page setup,
// header/footer and autofilter under the same local names, and `<customSheetViews>` follows
// `<sheetViews>` and `<autoFilter>` in the part. A saved view is Excel's, and it is not the sheet.

const CUSTOM_VIEWS =
  '<customSheetViews><customSheetView guid="{C3D4E5F6-0000-4000-8000-000000000001}">' +
  '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>' +
  '<selection pane="bottomRight" activeCell="C3" sqref="C3"/>' +
  '<rowBreaks count="1" manualBreakCount="1"><brk id="9" max="16383" man="1"/></rowBreaks>' +
  '<colBreaks count="1" manualBreakCount="1"><brk id="4" max="1048575" man="1"/></colBreaks>' +
  '<pageMargins left="2" right="2" top="2" bottom="2" header="1" footer="1"/>' +
  '<printOptions gridLines="1"/>' +
  '<pageSetup orientation="landscape"/>' +
  '<headerFooter><oddHeader>view header</oddHeader></headerFooter>' +
  '<autoFilter ref="A1:B9"/>' +
  '</customSheetView></customSheetViews>';

// Everything a custom view could have written onto the sheet, as plain data.
function viewSettings(worksheet: ReturnType<typeof sheet>) {
  return {
    view: {...worksheet.view},
    rowBreaks: [...worksheet.rowBreaks],
    columnBreaks: [...worksheet.columnBreaks],
    pageMargins: {...worksheet.pageMargins},
    printOptions: {...worksheet.printOptions},
    pageSetup: {...worksheet.pageSetup},
    headerFooter: {...worksheet.headerFooter},
    autoFilter: worksheet.autoFilter,
  };
}

test('a custom view on a sheet with no settings of its own leaves the sheet without any', () => {
  const plain = sheet('<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData/>');
  const viewed = sheet(
    `<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetData/>${CUSTOM_VIEWS}`,
  );
  assert.deepEqual(viewSettings(viewed), viewSettings(plain));
});

test("a custom view neither overwrites the sheet's own settings nor adds to them", () => {
  // The sheet's pane and filter come before the saved view and its breaks, margins and header after,
  // so this covers both an overwrite and an append.
  const own =
    '<sheetViews><sheetView workbookViewId="0">' +
    '<pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
    '<sheetData/><autoFilter ref="C1:D5"/>';
  const tail =
    '<pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>' +
    '<headerFooter><oddFooter>sheet footer</oddFooter></headerFooter>' +
    '<rowBreaks count="1" manualBreakCount="1"><brk id="3" max="16383" man="1"/></rowBreaks>';
  assert.deepEqual(viewSettings(sheet(own + CUSTOM_VIEWS + tail)), viewSettings(sheet(own + tail)));
});

test('cells after a custom view still read', () => {
  const worksheet = sheet(
    `<sheetData/>${CUSTOM_VIEWS}<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>`,
  );
  assert.deepEqual([...worksheet.merges], ['A1:B1'], 'the skip ends where the views do');
});

// ── Whole-line filters ───────────────────────────────────────────────────────────────────────────────

test('an autofilter over whole columns or whole rows reads as no filter, not as an abort', () => {
  for (const ref of ['A:C', '1:5']) {
    const worksheet = sheet(
      `<sheetData><row r="1"><c r="A1"><v>7</v></c></row></sheetData><autoFilter ref="${ref}"/>`,
    );
    assert.equal(worksheet.autoFilter, undefined, ref);
    assert.equal(worksheet.getCell('A1').value, 7, `${ref}: the cells are intact`);
  }
});

// ── Header and footer text ───────────────────────────────────────────────────────────────────────────

test('an empty header element sets nothing, and the element after it still reads its text', () => {
  const worksheet = sheet(
    '<sheetData/><headerFooter><oddHeader/><oddFooter>page &amp;P</oddFooter></headerFooter>',
  );
  assert.equal(worksheet.headerFooter.oddHeader, undefined);
  assert.equal(worksheet.headerFooter.oddFooter, 'page &P');
});

test('customWidth="false" suppresses the width exactly as customWidth="0" does', () => {
  for (const spelling of ['0', 'false']) {
    const width = sheet(
      `<cols><col min="2" max="2" width="12" customWidth="${spelling}"/></cols>`,
    ).getColumn(2).width;
    assert.equal(width, undefined, `customWidth="${spelling}" is not a custom width`);
  }
  assert.equal(
    sheet('<cols><col min="2" max="2" width="12" customWidth="1"/></cols>').getColumn(2).width,
    12,
    'and the flag set still carries the width through',
  );
});

test('customHeight="false" suppresses the row height exactly as customHeight="0" does', () => {
  for (const spelling of ['0', 'false']) {
    const height = sheet(
      `<sheetData><row r="3" ht="30" customHeight="${spelling}"/></sheetData>`,
    ).getRow(3).height;
    assert.equal(height, undefined, `customHeight="${spelling}" is not a custom height`);
  }
  assert.equal(
    sheet('<sheetData><row r="3" ht="30" customHeight="1"/></sheetData>').getRow(3).height,
    30,
    'and the flag set still carries the height through',
  );
});

test('a `<col>` span past XFD is clamped, not walked', () => {
  // Unclamped, `max="99999999"` materialised 16.7 million column records before dying on the map's
  // size limit. Assert the record count rather than the elapsed time: the count is what bounds it.
  const worksheet = sheet('<cols><col min="1" max="99999999" width="12" customWidth="1"/></cols>');
  assert.equal([...worksheet.columns()].length, MAX_COLUMN);
  assert.equal(worksheet.getColumn(MAX_COLUMN).width, 12, 'the span reaches the last real column');
});

test('a `<col>` element wholly outside the grid is dropped', () => {
  const worksheet = sheet(
    `<cols><col min="${MAX_COLUMN + 1}" max="99999" width="12" customWidth="1"/></cols>`,
  );
  assert.equal([...worksheet.columns()].length, 0);
});

test('a `<row>` past the last row is dropped rather than clamped onto it', () => {
  const worksheet = sheet(
    `<sheetData><row r="${MAX_ROW + 1}" ht="30" customHeight="1"/></sheetData>`,
  );
  assert.equal([...worksheet.rows()].length, 0);
  assert.equal(worksheet.getRow(MAX_ROW).height, undefined, 'and nothing landed on the last row');
});

// The streamer checked the row and dropped its cells, while this reader placed a cell whose own `r`
// was in the grid, so the same markup gave `A5 = 9` here and nothing through the streamer.
test('a cell in a `<row>` past the last row is dropped even when its own `r` is in the grid', () => {
  const worksheet = sheet(
    `<sheetData><row r="${MAX_ROW + 1}"><c r="A5"><v>9</v></c></row></sheetData>`,
  );
  assert.equal([...worksheet.rows()].length, 0, 'no row 5 was created to hold the cell');
  assert.equal(worksheet.getCell('A5').value, null);
});

test('a `<pane>` split the writer could not write back is dropped, not carried to it', () => {
  // Each of these used to land in `view` verbatim and surface later out of the serializer, as a
  // RangeError naming a column the file never mentioned.
  for (const attrs of [
    'xSplit="abc"',
    'xSplit="1.5" ySplit="2"',
    'ySplit="-3"',
    'xSplit="16384"',
    'ySplit="1048576"',
  ]) {
    const worksheet = sheet(
      `<sheetViews><sheetView workbookViewId="0"><pane ${attrs} topLeftCell="B2" state="frozen"/></sheetView></sheetViews>`,
    );
    assert.equal(worksheet.view.state, 'frozen', `${attrs}: the pane itself still reads`);
    for (const axis of ['xSplit', 'ySplit'] as const) {
      const split = worksheet.view[axis];
      assert.ok(
        split === undefined || (Number.isInteger(split) && split >= 0),
        `${attrs}: ${axis} is ${String(split)}, which freeze() would refuse`,
      );
    }
    assert.doesNotThrow(() => sheetViewsXml(worksheet.view, true), `${attrs}: and re-writes clean`);
  }
});

test('a well-formed `<pane>` split still reads', () => {
  const worksheet = sheet(
    '<sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="1" topLeftCell="C2" state="frozen"/></sheetView></sheetViews>',
  );
  assert.equal(worksheet.view.xSplit, 2);
  assert.equal(worksheet.view.ySplit, 1);
  assert.equal(worksheet.view.topLeftCell, 'C2');
});
