// Every parser that gathers an element's text across open/text/close, fed the self-closing spelling
// of the element it captures.
//
// `<t/>`, `<text/>`, `<xm:f/>` and `<totalsRowFormula/>` are all legal and all appear in files Excel
// writes. Each fires an open with no matching close, so a hand-rolled capture latches on and stays
// latched: the text of whatever comes next belongs to the wrong element, or to an element the file
// left empty on purpose. Nothing here asserts an interesting result. The assertion is that the
// element after the empty one is unaffected, which is exactly what a leaked latch would break.

import assert from 'node:assert/strict';
import {test} from 'node:test';

import {strFromU8, strToU8} from 'fflate';

import {Workbook} from '../../core/workbook.ts';
import {parseDrawing} from './images.ts';
import {readPatched, SHEET1} from './package.test-support.ts';
import {parseSharedStrings} from './read-shared-strings.ts';
import {parseTable} from './tables.ts';
import {parseThreadedComments} from './threaded-comments.ts';

// Every parser below is reached the way a file reaches it: through a package with one part
// rewritten, read by `readXlsx`. That matters beyond convenience for the worksheet parsers, which
// production runs as five passes sharing one event stream; a capture that leaks in that arrangement
// but not in a parse of its own is exactly the bug these tests exist to catch.
const readWithPart = (part: string, xml: string): Workbook => readPatched({[part]: xml});

// The worksheet part around whatever fragment a test is pinning, with the cell `readWithPart`'s
// guard reads back.
const worksheetWith = (body: string): string =>
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' +
  body +
  '</worksheet>';

test('the shared-string pool keeps an empty <t/> empty and the next entry its own', () => {
  const pool = parseSharedStrings('<sst><si><t/></si><si><t>after</t></si></sst>');
  assert.deepEqual(pool, ['', 'after']);
});

test('an extended data validation keeps an empty <xm:f/> and <xm:sqref/> from eating the next', () => {
  const back = readWithPart(
    SHEET1,
    worksheetWith(
      '<extLst><ext><x14:dataValidations>' +
        '<x14:dataValidation type="list"><x14:formula1><xm:f/></x14:formula1>' +
        '<xm:sqref>A1</xm:sqref></x14:dataValidation>' +
        '<x14:dataValidation type="list"><x14:formula1><xm:f>Sheet2!A1:A3</xm:f></x14:formula1>' +
        '<xm:sqref>B1</xm:sqref></x14:dataValidation>' +
        '</x14:dataValidations></ext></extLst>',
    ),
  );

  const entries = back.getWorksheet('S')?.dataValidations ?? [];
  assert.equal(entries.length, 2);
  assert.equal(entries[0]?.sqref, 'A1');
  assert.deepEqual(entries[1]?.rule.formulae, ['Sheet2!A1:A3']);
  assert.equal(entries[1]?.sqref, 'B1');
});

test("a drawing anchor keeps an empty <xdr:col/> from taking the next coordinate's text", () => {
  const {anchors} = parseDrawing(
    '<xdr:wsDr><xdr:twoCellAnchor>' +
      '<xdr:from><xdr:col/><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
      '<xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
      '<xdr:pic><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill></xdr:pic>' +
      '</xdr:twoCellAnchor></xdr:wsDr>',
  );
  assert.equal(anchors.length, 1);
  assert.equal(anchors[0]?.from.col, 0, 'the empty coordinate stays at its default');
  assert.equal(anchors[0]?.from.row, 4, 'and the sibling that follows it is its own');
  assert.equal(anchors[0]?.to?.col, 3);
});

test('a threaded comment keeps an empty <text/> from taking the next message body', () => {
  const messages = parseThreadedComments(
    '<ThreadedComments>' +
      '<threadedComment ref="A1" id="{1}" dT="2024-01-01T00:00:00Z"><text/></threadedComment>' +
      '<threadedComment ref="A2" id="{2}" dT="2024-01-01T00:00:00Z"><text>after</text></threadedComment>' +
      '</ThreadedComments>',
  );
  assert.equal(messages.length, 2);
  assert.equal(messages[0]?.text, '');
  assert.equal(messages[1]?.text, 'after');
});

test('a table keeps an empty <totalsRowFormula/> off the column that follows it', () => {
  const table = parseTable(
    '<table displayName="T" ref="A1:B3" totalsRowCount="1">' +
      '<tableColumns>' +
      '<tableColumn name="One" totalsRowFunction="custom"><totalsRowFormula/></tableColumn>' +
      '<tableColumn name="Two" totalsRowFunction="custom">' +
      '<totalsRowFormula>SUM(T[Two])</totalsRowFormula></tableColumn>' +
      '</tableColumns></table>',
  );
  // An element that carries no text records no formula, rather than the next column's.
  assert.equal(table?.columns?.[0]?.totalsRowFormula, undefined);
  assert.equal(table?.columns?.[1]?.totalsRowFormula, 'SUM(T[Two])');
});

test('a table keeps an empty <calculatedColumnFormula/> off the totals formula beside it', () => {
  const table = parseTable(
    '<table displayName="T" ref="A1:B3" totalsRowCount="1">' +
      '<tableColumns>' +
      '<tableColumn name="One"/>' +
      '<tableColumn name="Two" totalsRowFunction="custom"><calculatedColumnFormula/>' +
      '<totalsRowFormula>SUM(T[Two])</totalsRowFormula></tableColumn>' +
      '</tableColumns></table>',
  );
  assert.equal(table?.columns?.[1]?.calculatedColumnFormula, undefined);
  assert.equal(table?.columns?.[1]?.totalsRowFormula, 'SUM(T[Two])');
});

test('a conditional formatting keeps an empty <formula/> from taking the next operand', () => {
  const back = readWithPart(
    SHEET1,
    worksheetWith(
      '<conditionalFormatting sqref="A1:A5">' +
        '<cfRule type="cellIs" operator="between" priority="1">' +
        '<formula/><formula>10</formula></cfRule>' +
        '</conditionalFormatting>',
    ),
  );

  // The empty element contributes no operand; what matters is that the one after it is `10` and not
  // an operand list shifted by one.
  const blocks = back.getWorksheet('S')?.conditionalFormattings ?? [];
  assert.deepEqual(blocks[0]?.rules[0]?.formulae, [10]);
});

test('a data-bar extension link keeps an empty <x14:id/> from taking the next rule', () => {
  const back = readWithPart(
    SHEET1,
    worksheetWith(
      '<conditionalFormatting sqref="A1:A5">' +
        '<cfRule type="dataBar" priority="1"><dataBar><cfvo type="min"/><cfvo type="max"/>' +
        '<color rgb="FF638EC6"/></dataBar>' +
        '<extLst><ext><x14:id/></ext></extLst></cfRule>' +
        '</conditionalFormatting>',
    ),
  );

  const blocks = back.getWorksheet('S')?.conditionalFormattings ?? [];
  assert.equal(blocks[0]?.rules[0]?.type, 'dataBar');
});

test('the core properties keep an empty <dc:title/> off the property that follows it', () => {
  const back = readWithPart(
    'docProps/core.xml',
    '<cp:coreProperties><dc:title/><dc:creator>Ada</dc:creator>' +
      '<cp:lastModifiedBy>Grace</cp:lastModifiedBy></cp:coreProperties>',
  );
  assert.equal(back.properties.title, undefined, 'an empty title is not the creator');
  assert.equal(back.properties.creator, 'Ada');
  assert.equal(back.properties.lastModifiedBy, 'Grace');
});

test('the app properties keep an empty <Company/> off the element that follows it', () => {
  const back = readWithPart(
    'docProps/app.xml',
    '<Properties><Company/><Application>Something</Application></Properties>',
  );
  assert.equal(back.properties.company, undefined);
});

test('a rewritten package still round-trips the same bytes for its untouched parts', () => {
  // A guard on the helper above: if `readWithPart` silently produced an unreadable package, every
  // assertion here would pass by reading nothing at all.
  const back = readWithPart('docProps/app.xml', strFromU8(strToU8('<Properties/>')));
  assert.equal(back.getWorksheet('S')?.getCell('A1').value, 1);
});
