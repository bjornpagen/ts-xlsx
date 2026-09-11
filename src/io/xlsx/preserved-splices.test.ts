import assert from 'node:assert/strict';
import {test} from 'node:test';

import type {SheetSplice} from '../../core/formula-references.ts';
import {splicePreservedPart} from './preserved-splices.ts';

const CHART = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const PIVOT_CACHE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml';

const insertRows = (sheet: string, start: number, count = 1): SheetSplice => ({
  sheet,
  splice: {axis: 'row', start, count: 0, delta: count},
});
const deleteRows = (sheet: string, start: number, count: number): SheetSplice => ({
  sheet,
  splice: {axis: 'row', start, count, delta: -count},
});

const spliced = (xml: string, contentType: string, splices: readonly SheetSplice[]): string =>
  new TextDecoder().decode(
    splicePreservedPart(new TextEncoder().encode(xml), contentType, splices),
  );

test("a chart's series references move, in order, and nothing else in the part changes", () => {
  const chart =
    '<c:chartSpace xmlns:c="c"><c:ser><c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx>' +
    '<c:val><c:numRef><c:f>Data!$B$2:$B$5</c:f><c:numCache><c:pt idx="0"><c:v>2</c:v></c:pt>' +
    '</c:numCache></c:numRef></c:val></c:ser><!-- <c:f>Data!$B$2</c:f> --></c:chartSpace>';
  assert.equal(
    spliced(chart, CHART, [
      insertRows('Data', 3),
      insertRows('Other', 1),
      deleteRows('Data', 2, 1),
    ]),
    chart,
    'an insert at row 3 and a delete of row 2 cancel out for the range, and leave the name',
  );
  assert.equal(
    spliced(chart, CHART, [insertRows('Data', 1, 2)]),
    chart
      .replace('<c:f>Data!$B$1</c:f>', '<c:f>Data!$B$3</c:f>')
      .replace('<c:f>Data!$B$2:$B$5</c:f>', '<c:f>Data!$B$4:$B$7</c:f>'),
    'the comment is not a reference, and the cached value is not moved',
  );
});

test('a chart part nothing moved in is handed back as the same bytes', () => {
  const bytes = new TextEncoder().encode('<c:chartSpace><c:f>Data!$B$1</c:f></c:chartSpace>');
  assert.equal(splicePreservedPart(bytes, CHART, [insertRows('Other', 1)]), bytes);
  assert.equal(splicePreservedPart(bytes, CHART, []), bytes);
  assert.equal(
    splicePreservedPart(bytes, 'application/vnd.openxmlformats-officedocument.drawing+xml', [
      insertRows('Data', 1),
    ]),
    bytes,
    'a part of a kind that spells no references is not read',
  );
});

test('a reference is read decoded and written back escaped, and one holding markup is left', () => {
  const chart =
    "<c:chartSpace><c:f>'R&amp;D'!$A$1</c:f><c:f><![CDATA['R&D'!$A$1]]></c:f></c:chartSpace>";
  assert.equal(
    spliced(chart, CHART, [insertRows('R&D', 1)]),
    "<c:chartSpace><c:f>'R&amp;D'!$A$2</c:f><c:f><![CDATA['R&D'!$A$1]]></c:f></c:chartSpace>",
  );
});

test('a media type is matched ignoring case', () => {
  assert.equal(
    spliced('<c:f>Data!$A$1</c:f>', CHART.toUpperCase(), [insertRows('Data', 1)]),
    '<c:f>Data!$A$2</c:f>',
  );
});

test("a pivot cache's worksheet source moves through the splices of its sheet", () => {
  const cache = (source: string) =>
    '<pivotCacheDefinition xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1">' +
    `<cacheSource type="worksheet">${source}</cacheSource><cacheFields count="0"/></pivotCacheDefinition>`;
  assert.equal(
    spliced(cache('<worksheetSource ref="A1:B5" sheet="Data"/>'), PIVOT_CACHE, [
      insertRows('Data', 3),
      insertRows('Other', 1),
    ]),
    cache('<worksheetSource ref="A1:B6" sheet="Data"/>'),
  );
  assert.equal(
    spliced(cache('<worksheetSource ref="A1:B5" sheet="Data"></worksheetSource>'), PIVOT_CACHE, [
      deleteRows('Data', 5, 1),
    ]),
    cache('<worksheetSource ref="A1:B4" sheet="Data"/>'),
  );
  const whole = cache('<worksheetSource ref="A1:B5" sheet="Data"/>');
  assert.equal(
    spliced(whole, PIVOT_CACHE, [deleteRows('Data', 1, 5)]),
    whole,
    'a delete of the whole source leaves it',
  );
  const table = cache('<worksheetSource name="Sales"/>');
  assert.equal(
    spliced(table, PIVOT_CACHE, [insertRows('Data', 1)]),
    table,
    'a table source has no ref',
  );
  const elsewhere = cache('<worksheetSource ref="A1:B5" sheet="Data" r:id="rId2"/>');
  assert.equal(
    spliced(elsewhere, PIVOT_CACHE, [insertRows('Data', 1)]),
    elsewhere,
    "another workbook's range is not this workbook's to move",
  );
});

test('a part the scanner cannot read is handed back as it came', () => {
  const bytes = new TextEncoder().encode('<c:chartSpace><c:f>Data!$A$1</c:chartSpace>');
  assert.equal(splicePreservedPart(bytes, CHART, [insertRows('Data', 1)]), bytes);
});
