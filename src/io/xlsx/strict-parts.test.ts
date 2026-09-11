import assert from 'node:assert/strict';
import {test} from 'node:test';

import {transitionalPart} from './strict-parts.ts';

const DRAWING = 'application/vnd.openxmlformats-officedocument.drawing+xml';

const translated = (xml: string, contentType = DRAWING): string =>
  new TextDecoder().decode(transitionalPart(new TextEncoder().encode(xml), contentType));

test("a Strict part's namespace declarations and graphic frame are translated, and nothing else moves", () => {
  const drawing =
    '<?xml version="1.0"?>\n<xdr:wsDr xmlns:xdr="http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing" ' +
    'xmlns:a="http://purl.oclc.org/ooxml/drawingml/main">\n  <!-- keep -->\n' +
    "<a:graphic><a:graphicData uri='http://purl.oclc.org/ooxml/drawingml/chart'>" +
    '<c:chart xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" r:id="rId1"/>' +
    '</a:graphicData></a:graphic><a:t>http://purl.oclc.org/ooxml/drawingml/main</a:t></xdr:wsDr>';
  assert.equal(
    translated(drawing),
    '<?xml version="1.0"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">\n  <!-- keep -->\n' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
      '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>' +
      '</a:graphicData></a:graphic><a:t>http://purl.oclc.org/ooxml/drawingml/main</a:t></xdr:wsDr>',
    'text that happens to spell a namespace is text',
  );
});

test('a DrawingML percentage becomes thousandths and a chart percentage a plain integer', () => {
  const part =
    '<q:theme xmlns:q="http://purl.oclc.org/ooxml/drawingml/main">' +
    '<q:gs pos="33.5%"><q:srgbClr val="FF0000"><q:alpha val="60%"/><q:lumMod val="-5%"/></q:srgbClr></q:gs>' +
    '<q:clrScheme name="50%"/><q:buSzPct val="80%"/>' +
    '<c:holeSize xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart" val="50%"/>' +
    '<c:gapWidth xmlns:c="http://purl.oclc.org/ooxml/drawingml/chart" val="150"/></q:theme>';
  assert.equal(
    translated(part),
    '<q:theme xmlns:q="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<q:gs pos="33500"><q:srgbClr val="FF0000"><q:alpha val="60000"/><q:lumMod val="-5000"/></q:srgbClr></q:gs>' +
      '<q:clrScheme name="50%"/><q:buSzPct val="80%"/>' +
      '<c:holeSize xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" val="50"/>' +
      '<c:gapWidth xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" val="150"/></q:theme>',
    'a name is not a percentage, a bullet size is `N%` in Transitional too, and an integer stays',
  );
});

test('a part with nothing Strict in it, or not XML, or unreadable, is handed back as the same bytes', () => {
  for (const [xml, contentType] of [
    ['<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>', DRAWING],
    ['http://purl.oclc.org/ooxml/drawingml/main', 'image/png'],
    ['<a:theme xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"><!-- unterminated', DRAWING],
  ] as const) {
    const bytes = new TextEncoder().encode(xml);
    assert.equal(transitionalPart(bytes, contentType), bytes, xml);
  }
});
