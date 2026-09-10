import assert from 'node:assert/strict';
import {test} from 'node:test';

import type {Fill, Font} from '../../core/style.ts';
import {Workbook} from '../../core/workbook.ts';
import {XlsxError} from '../../errors.ts';
import {partText, patchParts, roundtrip} from './package.test-support.ts';
import {parseIndexedColors} from './read-styles.ts';
import {readXlsx} from './read.ts';
import {StyleRegistry} from './styles.ts';
import {writeXlsx} from './write.ts';

// An authored table style replaces the preserved definition it shares a name with, and the writer found
// that definition by pulling `name="..."` back out of its bytes. A file quoting the attribute the other
// legal way, `name='Harbour'`, slipped past, and the stylesheet carried two definitions of one name.
test('an authored table style replaces a preserved one of the same name, however the file quoted it', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').getCell('A1').value = 1;
  const preserved = `<tableStyles count="1"><tableStyle name='Harbour' pivot='0'></tableStyle></tableStyles>`;
  const read = readXlsx(
    patchParts(writeXlsx(workbook), {
      'xl/styles.xml': (xml) => xml.replace('</styleSheet>', `${preserved}</styleSheet>`),
    }),
  );
  assert.deepEqual(
    read.tableStyles.styles.map((style) => style.name),
    ['Harbour'],
    'the preserved definition is named from the scanner, not from a double-quoted pattern',
  );

  read.addTableStyle({name: 'Harbour', elements: {wholeTable: {border: {top: {style: 'thin'}}}}});
  const block = /<tableStyles\b[^>]*>[\s\S]*?<\/tableStyles>/.exec(
    partText(writeXlsx(read), 'xl/styles.xml'),
  )?.[0];
  assert.ok(block, 'the stylesheet carries a <tableStyles> block');
  const names = [...block.matchAll(/<tableStyle\s[^>]*?\bname=(["'])(.*?)\1/g)].map((m) => m[2]);
  assert.deepEqual(names, ['Harbour'], 'one definition of the name, not two');
  assert.match(block, /^<tableStyles count="1"/);
  assert.doesNotMatch(block, /name='Harbour'/, 'and it is the authored one');
});

const solid = (argb: string): Fill => ({type: 'pattern', pattern: 'solid', fgColor: {argb}});

test('parseIndexedColors extracts each rgbColor verbatim, or nothing when the palette is default', () => {
  const styles =
    '<styleSheet><colors><indexedColors>' +
    '<rgbColor rgb="ff000000"/><rgbColor rgb="ff3f6797"/><rgbColor rgb="ffaaaaaa"/>' +
    '</indexedColors></colors></styleSheet>';
  assert.deepEqual(parseIndexedColors(styles), [
    '<rgbColor rgb="ff000000"/>',
    '<rgbColor rgb="ff3f6797"/>',
    '<rgbColor rgb="ffaaaaaa"/>',
  ]);
  assert.deepEqual(parseIndexedColors('<styleSheet><dxfs count="0"/></styleSheet>'), []);
});

test('a seeded indexed-color palette re-emits as a <colors> block; an unseeded one emits none', () => {
  const seeded = new StyleRegistry();
  seeded.seedIndexedColors(['<rgbColor rgb="ff3f6797"/>', '<rgbColor rgb="ffaaaaaa"/>']);
  const xml = seeded.toXml();
  assert.match(
    xml,
    /<colors><indexedColors><rgbColor rgb="ff3f6797"\/><rgbColor rgb="ffaaaaaa"\/><\/indexedColors><\/colors>/,
  );
  // The palette is a late child of <styleSheet>: after <dxfs>, before the close.
  assert.ok(xml.indexOf('<dxfs') < xml.indexOf('<colors>'), '<colors> follows <dxfs>');
  assert.doesNotMatch(
    new StyleRegistry().toXml(),
    /<colors>/,
    'the default palette writes no <colors>',
  );
});

test('an absent or "none" fill with no format resolves to the default xf 0: no style entry', () => {
  const styles = new StyleRegistry();
  assert.equal(styles.styleId({}), 0);
  assert.equal(styles.styleId({fill: {type: 'pattern', pattern: 'none'}}), 0);
  assert.equal(styles.styleId({numFmt: ''}), 0);
});

test('an identical fill interns to one shared xf index however many times it is seen', () => {
  const styles = new StyleRegistry();
  const first = styles.styleId({fill: solid('FFFF0000')});
  for (let i = 0; i < 40; i++) {
    assert.equal(
      styles.styleId({fill: solid('FFFF0000')}),
      first,
      'every identical fill returns the same index',
    );
  }
  assert.notEqual(first, 0, 'a real fill gets a non-default index');
});

test('a gradient fill emits a <gradientFill> with its stops, and interns like any other fill', () => {
  const styles = new StyleRegistry();
  const grad: Fill = {
    type: 'gradient',
    gradient: 'linear',
    degree: 90,
    stops: [
      {position: 0, color: {argb: 'FFFFFFFF'}},
      {position: 1, color: {argb: 'FF4472C4'}},
    ],
  };
  const first = styles.styleId({fill: grad});
  assert.notEqual(first, 0, 'a gradient is a real fill, never the default xf 0');
  assert.equal(styles.styleId({fill: grad}), first, 'an identical gradient interns to the same xf');

  const xml = styles.toXml();
  assert.match(
    xml,
    /<fill><gradientFill degree="90"><stop position="0"><color rgb="FFFFFFFF"\/><\/stop><stop position="1"><color rgb="FF4472C4"\/><\/stop><\/gradientFill><\/fill>/,
  );
});

test('a path gradient emits its type and insets and omits a zero degree', () => {
  const styles = new StyleRegistry();
  styles.styleId({
    fill: {
      type: 'gradient',
      gradient: 'path',
      left: 0.5,
      right: 0.5,
      top: 0.25,
      bottom: 0.25,
      stops: [{position: 0, color: {argb: 'FF000000'}}],
    },
  });
  const xml = styles.toXml();
  assert.match(xml, /<gradientFill type="path" left="0.5" right="0.5" top="0.25" bottom="0.25">/);
  assert.doesNotMatch(xml, /degree=/, 'a path gradient with no degree emits none');
});

test('an identical number format interns to one shared xf index', () => {
  const styles = new StyleRegistry();
  const first = styles.styleId({numFmt: '0.00%'});
  for (let i = 0; i < 40; i++) {
    assert.equal(
      styles.styleId({numFmt: '0.00%'}),
      first,
      'every identical numFmt returns the same index',
    );
  }
  assert.notEqual(first, 0, 'a real numFmt gets a non-default index');
});

test('fill and number format are independent facets: the same fill under two formats is two xfs', () => {
  const styles = new StyleRegistry();
  const plain = styles.styleId({fill: solid('FFFF0000')});
  const formatted = styles.styleId({fill: solid('FFFF0000'), numFmt: '0.00'});
  assert.notEqual(plain, formatted, 'adding a number format to a fill is a distinct style');
});

test('genuinely different styles get distinct xf indices: dedup does not over-collapse', () => {
  const styles = new StyleRegistry();
  const red = styles.styleId({fill: solid('FFFF0000')});
  const blue = styles.styleId({fill: solid('FF0000FF')});
  const pct = styles.styleId({numFmt: '0.00%'});
  const cur = styles.styleId({numFmt: '"$"#,##0.00'});
  assert.equal(new Set([red, blue, pct, cur]).size, 4, 'four distinct styles, four indices');
});

test('the emitted stylesheet reflects the interned fills and cell formats', () => {
  const styles = new StyleRegistry();
  styles.styleId({fill: solid('FFFF0000')});
  styles.styleId({fill: solid('FFFF0000')}); // deduped
  styles.styleId({fill: solid('FF00FF00')});
  const xml = styles.toXml();

  // Two reserved fills (none, gray125) + two custom = 4; the default xf + two custom = 3.
  assert.match(xml, /<fills count="4">/);
  assert.match(xml, /<cellXfs count="3">/);
  // The visible colour is the pattern foreground with an automatic indexed background.
  assert.match(
    xml,
    /<patternFill patternType="solid"><fgColor rgb="FFFF0000"\/><bgColor indexed="64"\/><\/patternFill>/,
  );
  // A styled xf references its fill and flags applyFill; the default xf does neither.
  assert.match(xml, /<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"\/>/);
  assert.match(
    xml,
    /<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"\/>/,
  );
});

test('a "#"-prefixed fill colour is normalized to a bare 8-hex ARGB, never emitted as a 9-char value', () => {
  const styles = new StyleRegistry();
  styles.styleId({fill: solid('#FFBFBFBF')});
  const xml = styles.toXml();

  assert.match(
    xml,
    /<fgColor rgb="FFBFBFBF"\/>/,
    'the leading "#" is stripped so the value is a valid 8 hex digits',
  );
  assert.doesNotMatch(xml, /rgb="#/, 'no colour is ever written with a leading "#"');
});

test('a bare 8-hex fill colour is preserved verbatim, including its original casing', () => {
  const styles = new StyleRegistry();
  styles.styleId({fill: solid('ff00ff00')});
  // Foreign files store rgb in lowercase; round-tripping must not silently re-case it.
  assert.match(styles.toXml(), /<fgColor rgb="ff00ff00"\/>/);
});

test('a 6-hex RGB fill colour is promoted to ARGB with a fully-opaque alpha', () => {
  const styles = new StyleRegistry();
  styles.styleId({fill: solid('00FF00')});
  // A colour written without its alpha channel is the common case; it must not reach Excel as a
  // 6-char rgb (which renders black), but as a valid opaque 8-hex ARGB.
  assert.match(styles.toXml(), /<fgColor rgb="FF00FF00"\/>/);
});

test('a "#"-prefixed 6-hex RGB is both stripped and promoted to opaque ARGB', () => {
  const styles = new StyleRegistry();
  styles.styleId({fill: solid('#00FF00')});
  assert.match(styles.toXml(), /<fgColor rgb="FF00FF00"\/>/);
});

test('a lowercase ARGB is written back exactly as it arrived', () => {
  const styles = new StyleRegistry();
  styles.styleId({fill: solid('ff00ff00')});
  // The writer shares its ARGB grammar with the resolver, which uppercases; this one must not, or a
  // foreign file's own casing stops surviving a read-write round trip.
  assert.match(styles.toXml(), /<fgColor rgb="ff00ff00"\/>/);
});

test('an ARGB that is neither 6 nor 8 hex digits is rejected at the API surface', () => {
  const styles = new StyleRegistry();
  // A malformed colour silently renders as flat black in Excel; fail loud instead of writing it.
  // Native `SyntaxError`, not the library's taxonomy: one string that does not parse, the same kind
  // of failure a comment id that is not a GUID raises. Asserted as *not* an XlsxError too, so a
  // later re-wrap reddens the suite rather than quietly changing what a caller catches.
  assert.throws(() => styles.styleId({fill: solid('12345')}), {
    name: 'SyntaxError',
    message: /Invalid ARGB colour "12345"/,
  });
  assert.throws(() => styles.styleId({fill: solid('GGGGGGGG')}), {name: 'SyntaxError'});
  assert.throws(() => styles.styleId({fill: solid('red')}), {name: 'SyntaxError'});
  assert.throws(
    () => styles.styleId({fill: solid('red')}),
    (error: unknown) => !(error instanceof XlsxError),
  );
});

test('a custom number format is defined in <numFmts> from id 164 and referenced by its xf', () => {
  const styles = new StyleRegistry();
  styles.styleId({numFmt: '0.00%'});
  styles.styleId({numFmt: '"$"#,##0.00'});
  const xml = styles.toXml();

  assert.match(xml, /<numFmts count="2">/);
  assert.match(xml, /<numFmt numFmtId="164" formatCode="0.00%"\/>/);
  // A quoted currency literal survives with its markup-significant characters escaped.
  assert.match(xml, /<numFmt numFmtId="165" formatCode="&quot;\$&quot;#,##0.00"\/>/);
  // The referencing xf names the custom id and flags applyNumberFormat.
  assert.match(
    xml,
    /<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"\/>/,
  );
});

test('an empty registry still emits a valid minimal stylesheet with no <numFmts>', () => {
  const xml = new StyleRegistry().toXml();
  assert.doesNotMatch(xml, /<numFmts/); // omitted entirely when all-built-in
  assert.match(xml, /<fonts count="1">/); // just the default font
  assert.match(xml, /<fills count="2">/); // just the two reserved fills
  assert.match(xml, /<borders count="1">/); // just the empty default border
  assert.match(xml, /<cellXfs count="1">/); // just the default xf
});

test('an identical font interns to one shared xf index', () => {
  const styles = new StyleRegistry();
  const first = styles.styleId({font: {bold: true}});
  for (let i = 0; i < 40; i++) {
    assert.equal(
      styles.styleId({font: {bold: true}}),
      first,
      'every identical font returns the same index',
    );
  }
  assert.notEqual(first, 0, 'a real font gets a non-default index');
});

test('a font that overrides nothing resolves to the default xf 0: no font entry', () => {
  const styles = new StyleRegistry();
  assert.equal(styles.styleId({font: {}}), 0);
  // A boolean flag that is explicitly false is the default and adds no <font>.
  assert.equal(styles.styleId({font: {bold: false, italic: false}}), 0);
  assert.doesNotMatch(styles.toXml(), /<b\/>/);
  assert.match(styles.toXml(), /<fonts count="1">/);
});

test('font, fill, and number format are independent facets composed into one xf', () => {
  const styles = new StyleRegistry();
  const bold = styles.styleId({font: {bold: true}});
  const boldRed = styles.styleId({font: {bold: true}, fill: solid('FFFF0000')});
  const boldRedPct = styles.styleId({font: {bold: true}, fill: solid('FFFF0000'), numFmt: '0.00%'});
  assert.equal(
    new Set([bold, boldRed, boldRedPct]).size,
    3,
    'each added facet is a distinct composed style',
  );
});

test('a custom font is defined in <fonts> after the default and referenced by its xf', () => {
  const styles = new StyleRegistry();
  styles.styleId({
    font: {bold: true, italic: true, size: 14, color: {argb: 'FF3A80D5'}, name: 'Arial'},
  });
  const xml = styles.toXml();

  // Default font (id 0) + one custom (id 1).
  assert.match(xml, /<fonts count="2">/);
  // The facets serialise in ECMA-376 child order, with the typeface attribute-escaped.
  assert.match(
    xml,
    /<font><b\/><i\/><sz val="14"\/><color rgb="FF3A80D5"\/><name val="Arial"\/><\/font>/,
  );
  // The referencing xf names font id 1 and flags applyFont; the default xf keeps font id 0.
  assert.match(
    xml,
    /<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"\/>/,
  );
  assert.match(xml, /<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"\/>/);
});

test('an underline flag serialises single as a bare tag and a named variant with its val', () => {
  const styles = new StyleRegistry();
  styles.styleId({font: {underline: true}});
  styles.styleId({font: {underline: 'double'}});
  const xml = styles.toXml();
  assert.match(xml, /<font><u\/><\/font>/); // true → single → bare <u/>
  assert.match(xml, /<font><u val="double"\/><\/font>/);
});

test('an identical border interns to one shared xf index', () => {
  const styles = new StyleRegistry();
  const first = styles.styleId({border: {top: {style: 'thin'}}});
  for (let i = 0; i < 40; i++) {
    assert.equal(
      styles.styleId({border: {top: {style: 'thin'}}}),
      first,
      'every identical border returns the same index',
    );
  }
  assert.notEqual(first, 0, 'a real border gets a non-default index');
});

test('a border that styles no edge resolves to the default xf 0: no border entry', () => {
  const styles = new StyleRegistry();
  assert.equal(styles.styleId({border: {}}), 0);
  assert.match(styles.toXml(), /<borders count="1">/); // just the empty default border
});

test('font, fill, number format, and border are independent facets composed into one xf', () => {
  const styles = new StyleRegistry();
  const b = styles.styleId({border: {top: {style: 'thin'}}});
  const bFill = styles.styleId({border: {top: {style: 'thin'}}, fill: solid('FFFF0000')});
  const bFillFont = styles.styleId({
    border: {top: {style: 'thin'}},
    fill: solid('FFFF0000'),
    font: {bold: true},
  });
  assert.equal(
    new Set([b, bFill, bFillFont]).size,
    3,
    'each added facet is a distinct composed style',
  );
});

test('a custom border is defined in <borders> after the default and referenced by its xf', () => {
  const styles = new StyleRegistry();
  styles.styleId({
    border: {top: {style: 'thin'}, bottom: {style: 'medium', color: {argb: 'FF3A80D5'}}},
  });
  const xml = styles.toXml();

  // Empty default border (id 0) + one custom (id 1).
  assert.match(xml, /<borders count="2">/);
  // Edges serialise in schema order (left, right, top, bottom, diagonal); a styleless edge is
  // a bare self-closing tag, a styled one carries its style and any colour child.
  assert.match(
    xml,
    /<border><left\/><right\/><top style="thin"\/><bottom style="medium"><color rgb="FF3A80D5"\/><\/bottom><diagonal\/><\/border>/,
  );
  // The referencing xf names border id 1 and flags applyBorder; the default xf keeps border id 0.
  assert.match(
    xml,
    /<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"\/>/,
  );
  assert.match(xml, /<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"\/>/);
});

test('an identical alignment interns to one shared xf index', () => {
  const styles = new StyleRegistry();
  const first = styles.styleId({alignment: {horizontal: 'center'}});
  for (let i = 0; i < 40; i++) {
    assert.equal(
      styles.styleId({alignment: {horizontal: 'center'}}),
      first,
      'every identical alignment returns the same index',
    );
  }
  assert.notEqual(first, 0, 'a real alignment gets a non-default index');
});

test('an all-default alignment resolves to the default xf 0: no <alignment>', () => {
  const styles = new StyleRegistry();
  // `general` horizontal is the default, and boolean flags left off contribute nothing.
  assert.equal(
    styles.styleId({alignment: {horizontal: 'general', wrapText: false, shrinkToFit: false}}),
    0,
  );
  assert.doesNotMatch(styles.toXml(), /<alignment/);
});

test('alignment composes into the xf as a child element, not a shared sub-table', () => {
  const styles = new StyleRegistry();
  styles.styleId({
    alignment: {horizontal: 'center', vertical: 'top', wrapText: true, indent: 2, textRotation: 45},
  });
  const xml = styles.toXml();

  // The aligned xf carries an <alignment> child in ECMA-376 attribute order and flags applyAlignment;
  // it is no longer self-closing. The default xf stays self-closing with no alignment.
  assert.match(
    xml,
    /<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="top" textRotation="45" wrapText="1" indent="2"\/><\/xf>/,
  );
  assert.match(xml, /<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"\/>/);
});

test('alignment is an independent facet composed alongside fill, font, and border', () => {
  const styles = new StyleRegistry();
  const a = styles.styleId({alignment: {wrapText: true}});
  const aFill = styles.styleId({alignment: {wrapText: true}, fill: solid('FFFF0000')});
  const aFillFont = styles.styleId({
    alignment: {wrapText: true},
    fill: solid('FFFF0000'),
    font: {bold: true},
  });
  assert.equal(
    new Set([a, aFill, aFillFont]).size,
    3,
    'each added facet is a distinct composed style',
  );
});

test('an identical protection interns to one shared xf index', () => {
  const styles = new StyleRegistry();
  const first = styles.styleId({protection: {locked: false}});
  for (let i = 0; i < 40; i++) {
    assert.equal(
      styles.styleId({protection: {locked: false}}),
      first,
      'every identical protection returns the same index',
    );
  }
  assert.notEqual(first, 0, 'a real protection gets a non-default index');
});

test('an all-default protection resolves to the default xf 0: no <protection>', () => {
  const styles = new StyleRegistry();
  // locked defaults to TRUE and hidden to false in OOXML, so a locked, non-hidden cell restates
  // the default and carries no information: it must not spend an xf entry.
  assert.equal(styles.styleId({protection: {locked: true, hidden: false}}), 0);
  assert.doesNotMatch(styles.toXml(), /<protection/);
});

test('protection composes into the xf as a child element, flagging applyProtection', () => {
  const styles = new StyleRegistry();
  styles.styleId({protection: {locked: false, hidden: true}});
  const xml = styles.toXml();

  // Only the meaningful flags serialise: the unlocked cell writes locked="0", the hidden one hidden="1".
  assert.match(
    xml,
    /<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyProtection="1"><protection locked="0" hidden="1"\/><\/xf>/,
  );
  assert.match(xml, /<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"\/>/);
});

test('alignment and protection compose as two xf children in schema order', () => {
  const styles = new StyleRegistry();
  styles.styleId({alignment: {horizontal: 'center'}, protection: {locked: false}});
  const xml = styles.toXml();

  // <alignment> precedes <protection> in the xf body, and both apply flags are set.
  assert.match(
    xml,
    /<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1" applyProtection="1"><alignment horizontal="center"\/><protection locked="0"\/><\/xf>/,
  );
});

test('protection is an independent facet composed alongside fill, font, and border', () => {
  const styles = new StyleRegistry();
  const p = styles.styleId({protection: {locked: false}});
  const pFill = styles.styleId({protection: {locked: false}, fill: solid('FFFF0000')});
  const pFillFont = styles.styleId({
    protection: {locked: false},
    fill: solid('FFFF0000'),
    font: {bold: true},
  });
  assert.equal(
    new Set([p, pFill, pFillFont]).size,
    3,
    'each added facet is a distinct composed style',
  );
});

test('a non-string number format is dropped rather than corrupting the styles part', () => {
  const styles = new StyleRegistry();
  // A caller wrongly assigns a structured `{id, formatCode}` object where a format-code string belongs.
  const id = styles.styleId({numFmt: {id: 164, formatCode: '0.00'} as unknown as string});
  assert.equal(id, 0, 'a non-string numFmt contributes no style: it resolves to the default xf 0');
  const xml = styles.toXml();
  assert.doesNotMatch(
    xml,
    /\[object Object\]/,
    'the styles part must never carry a stringified object',
  );
  assert.doesNotMatch(xml, /<numFmts/, 'no custom numFmt entry is emitted for the dropped object');
});

test('a valid format-code string still interns alongside other facets', () => {
  const styles = new StyleRegistry();
  const id = styles.styleId({
    numFmt: 'yyyy-mmm-dd',
    font: {bold: true},
    protection: {locked: false},
  });
  assert.notEqual(id, 0, 'the string format composes into a real style');
  assert.match(styles.toXml(), /formatCode="yyyy-mmm-dd"/);
});

test('the quote-prefix flag emits quotePrefix="1" on the xf and forms a distinct style', () => {
  const styles = new StyleRegistry();
  assert.equal(
    styles.styleId({quotePrefix: false}),
    0,
    'an unset quote-prefix contributes no style',
  );
  const q = styles.styleId({quotePrefix: true});
  assert.notEqual(q, 0, 'a quote-prefixed style is a real entry, not the default xf 0');
  assert.equal(
    styles.styleId({quotePrefix: true}),
    q,
    'identical quote-prefixed styles intern to one entry',
  );
  assert.match(
    styles.toXml(),
    /<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" quotePrefix="1"\/>/,
  );
});

test('seeded named cell styles emit a cellStyleXfs/cellStyles layer and cells link to it via xfId', () => {
  const styles = new StyleRegistry();
  styles.seedNamedStyles([
    {name: 'Normal', builtinId: 0},
    {name: 'Accent', fill: solid('FFFFFF00')},
  ]);
  // A cell that links to the Accent named style carries its xfId even with no direct facet of its own.
  const linked = styles.styleId({xfId: 1});
  assert.notEqual(linked, 0, 'a cell linking to a named style needs a real cellXfs entry');
  const xml = styles.toXml();
  assert.match(
    xml,
    /<cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"\/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" applyFill="1"\/><\/cellStyleXfs>/,
  );
  assert.match(
    xml,
    /<cellStyles count="2"><cellStyle name="Normal" xfId="0" builtinId="0"\/><cellStyle name="Accent" xfId="1"\/><\/cellStyles>/,
  );
  assert.match(
    xml,
    /<cellXfs count="2">.*<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="1"\/><\/cellXfs>/,
  );
});

test('with no named styles the default cellStyleXfs and cellStyles layer is emitted unchanged', () => {
  const xml = new StyleRegistry().toXml();
  assert.match(
    xml,
    /<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"\/><\/cellStyleXfs>/,
  );
  assert.match(
    xml,
    /<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"\/><\/cellStyles>/,
  );
});

test('font 0 is the workbook default the registry was built with, not an assumed Calibri', () => {
  const styles = new StyleRegistry({
    defaultFont: {size: 11, color: {theme: 1}, name: 'Aptos Narrow', family: 2, scheme: 'minor'},
  });
  assert.match(
    styles.toXml(),
    /<fonts count="1"><font><sz val="11"\/><color theme="1"\/><name val="Aptos Narrow"\/><family val="2"\/><scheme val="minor"\/><\/font><\/fonts>/,
  );
});

test('a registry built with no workbook still emits the Office default font 0', () => {
  // ~40 call sites in this file construct a registry bare; the fallback is what keeps them meaning
  // what they did, and what keeps a plain `new Workbook()` emitting the bytes it always has.
  assert.match(
    new StyleRegistry().toXml(),
    /<fonts count="1"><font><sz val="11"\/><color theme="1"\/><name val="Calibri"\/><family val="2"\/><scheme val="minor"\/><\/font><\/fonts>/,
  );
});

test("the Office fallback is exactly what a plain workbook's resolved default font serialises to", () => {
  // The constant and the model's resolution are two statements of one fact, and only this holds them
  // together: drift would make a bare registry and a workbook-built one disagree on font 0.
  assert.equal(
    new StyleRegistry({defaultFont: new Workbook().defaultFont}).toXml(),
    new StyleRegistry().toXml(),
  );
});

test('a cell carrying the workbook default font interns to id 0, adding no redundant entry', () => {
  const defaultFont: Font = {size: 11, color: {theme: 1}, name: 'Aptos Narrow', scheme: 'minor'};
  const styles = new StyleRegistry({defaultFont});
  assert.equal(styles.styleId({font: defaultFont}), 0, 'it needs no cellXfs entry either');
  assert.match(styles.toXml(), /<fonts count="1">/);
});

test('a cell carrying the *source* default font also interns to id 0, so it follows a new default', () => {
  // The reader flattens font 0 onto every xf that names it, so a cell that merely inherited the
  // file's default arrives carrying it as a concrete face. Treating that as an authored intent would
  // pin the cell to the old face through a custom entry while its unstyled neighbours moved.
  const declaredDefaultFont: Font = {size: 11, name: 'Aptos Narrow'};
  const styles = new StyleRegistry({
    defaultFont: {size: 11, name: 'Georgia'},
    declaredDefaultFont,
  });
  assert.equal(styles.styleId({font: declaredDefaultFont}), 0);
  assert.match(styles.toXml(), /<fonts count="1"><font><sz val="11"\/><name val="Georgia"\/>/);
});

test('a font that is neither default is still interned as a custom entry', () => {
  const styles = new StyleRegistry({
    defaultFont: {size: 11, name: 'Georgia'},
    declaredDefaultFont: {size: 11, name: 'Aptos Narrow'},
  });
  assert.equal(styles.styleId({font: {size: 11, name: 'Courier New'}}), 1);
  assert.match(styles.toXml(), /<fonts count="2">/);
});

test('a number format code keeps its bare apostrophe but not a character XML cannot carry', () => {
  const styles = new StyleRegistry({defaultFont: {size: 11, name: 'Calibri'}});
  styles.styleId({numFmt: "d mmm yy's"});
  assert.ok(
    styles.toXml().includes('formatCode="d mmm yy\'s"'),
    'the apostrophe is not markup-significant inside a double-quoted attribute, so it stays bare',
  );

  const hostile = new StyleRegistry({defaultFont: {size: 11, name: 'Calibri'}});
  hostile.styleId({numFmt: 'BAD\u0001'});
  assert.throws(
    () => hostile.toXml(),
    {name: 'AuthoringError'},
    'and the unrepresentable one is refused',
  );
});

test('every alignment facet set to a non-default value survives a round-trip', () => {
  // The assertion the shared facet table makes cheap and nobody wrote while it was seven hand-paired
  // clauses: a facet written but not read would survive the write and vanish here, and one read but
  // not written would never reach the file at all.
  const alignment = {
    horizontal: 'centerContinuous',
    vertical: 'justify',
    textRotation: 45,
    wrapText: true,
    indent: 3,
    shrinkToFit: true,
    readingOrder: 2,
  } as const;
  const workbook = new Workbook();
  workbook.addWorksheet('S').getCell('A1').alignment = alignment;
  const back = roundtrip(workbook).getWorksheet('S')?.getCell('A1').alignment;
  assert.deepEqual(back, alignment, 'all seven facets come back exactly as they were set');
});

test('alignment attributes are emitted in CT_CellAlignment order', () => {
  // ECMA-376 orders the attributes, and the facet table is the declaration that fixes that order for
  // both directions; a table reordered for readability would quietly reorder the output.
  const styles = new StyleRegistry();
  styles.styleId({
    alignment: {
      horizontal: 'right',
      vertical: 'top',
      textRotation: 90,
      wrapText: true,
      indent: 1,
      shrinkToFit: true,
      readingOrder: 1,
    },
  });
  const tag = /<alignment ([^/>]*)\/>/.exec(styles.toXml())?.[1];
  assert.equal(
    tag,
    'horizontal="right" vertical="top" textRotation="90" wrapText="1" indent="1" shrinkToFit="1" readingOrder="1"',
  );
});
