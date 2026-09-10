import assert from 'node:assert/strict';
import {test} from 'node:test';

import {
  applyThemeOverrides,
  DEFAULT_THEME_XML,
  parseThemeColorScheme,
  parseThemeFontScheme,
} from './theme-xml.ts';

test('parseThemeColorScheme reads srgbClr and sysClr slots', () => {
  const xml =
    '<a:theme><a:themeElements><a:clrScheme name="X">' +
    '<a:dk1><a:sysClr val="windowText" lastClr="1A1A1A"/></a:dk1>' +
    '<a:lt1><a:sysClr val="window" lastClr="FAFAFA"/></a:lt1>' +
    '<a:accent1><a:srgbClr val="BB2649"/></a:accent1>' +
    '</a:clrScheme>' +
    '<a:fmtScheme><a:gs><a:srgbClr val="DEADBE"/></a:gs></a:fmtScheme>' +
    '</a:themeElements></a:theme>';
  const scheme = parseThemeColorScheme(xml);
  // A sysClr's `val` names an operating-system colour; only its `lastClr` is a usable value.
  assert.equal(scheme.dk1, '1A1A1A');
  assert.equal(scheme.lt1, 'FAFAFA');
  assert.equal(scheme.accent1, 'BB2649');
  // Slots the scheme does not declare stay absent, and nothing outside <clrScheme> is picked up.
  assert.equal(scheme.accent2, undefined);
  assert.equal(Object.values(scheme).includes('DEADBE'), false);
});

test('parseThemeColorScheme drops a slot it cannot decode rather than guessing', () => {
  const xml =
    '<a:clrScheme><a:accent1><a:hslClr hue="0" sat="0" lum="0"/></a:accent1>' +
    '<a:accent2><a:srgbClr val="zzzzzz"/></a:accent2></a:clrScheme>';
  assert.deepEqual(parseThemeColorScheme(xml), {});
});

test('a theme part with no colour scheme yields nothing', () => {
  assert.deepEqual(parseThemeColorScheme('<a:theme/>'), {});
});

test('parseThemeFontScheme reads the major and minor latin faces', () => {
  const scheme = parseThemeFontScheme(DEFAULT_THEME_XML);
  assert.equal(scheme.major, 'Calibri Light');
  assert.equal(scheme.minor, 'Calibri');
  assert.deepEqual(parseThemeFontScheme('<a:theme/>'), {}, 'and nothing out of a part without one');
});

test('a typeface carrying an entity reads back decoded, matching what the writer escaped', () => {
  // The writer escapes the face through `escapeAttr`, so the reader is the half that has to agree
  // with it: without decoding, authoring "A&B" reads back as "A&amp;B" and a second write doubles
  // the escape. The colour slots need none of this, being six hex digits.
  const authored = applyThemeOverrides(DEFAULT_THEME_XML, {
    fonts: {major: 'Ampersand & Co', minor: 'Angle <Bracket>'},
  });
  assert.ok(authored.includes('typeface="Ampersand &amp; Co"'), 'the part holds the escaped form');
  assert.deepEqual(parseThemeFontScheme(authored), {
    major: 'Ampersand & Co',
    minor: 'Angle <Bracket>',
  });
});

test('applyThemeOverrides replaces only the slots the caller named', () => {
  const xml = applyThemeOverrides(DEFAULT_THEME_XML, {colors: {accent1: 'BB2649'}});
  assert.ok(xml.includes('<a:accent1><a:srgbClr val="BB2649"/></a:accent1>'));
  assert.ok(xml.includes('<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>'), 'the rest stand');
  // dk1/lt1 are sysClr in the source; re-serialising them as srgbClr would pin them to one machine.
  assert.ok(xml.includes('<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'));
  assert.ok(xml.includes('<a:fmtScheme'), 'and the format scheme is untouched');
});

test('an authored typeface is escaped into the attribute, every character of it', () => {
  // The escape this replaced handled `& < > "` and stopped there, so a face carrying an apostrophe
  // or a newline went into the attribute raw.
  const face = 'A&B <C> "D" \'E\'\nF';
  const xml = applyThemeOverrides(DEFAULT_THEME_XML, {fonts: {minor: face}});
  const typeface = /<a:minorFont>[\s\S]*?<a:latin typeface="([^"]*)"/.exec(xml)?.[1];
  assert.equal(typeface, 'A&amp;B &lt;C&gt; &quot;D&quot; &apos;E&apos;&#10;F');
});

test('an authored typeface XML cannot represent is refused rather than written raw', () => {
  assert.throws(() => applyThemeOverrides(DEFAULT_THEME_XML, {fonts: {major: 'Bad\u0001Face'}}), {
    name: 'AuthoringError',
  });
});

test('a theme with nothing overridden comes back byte for byte', () => {
  // The edit is a splice at offsets the scanner found, so everything outside the ranges it names is
  // the source's own bytes: its whitespace, its comments, its attribute order, its prefix.
  const source =
    '<?xml version="1.0"?>\n' +
    "<!-- a designer's theme, with prose in it -->\n" +
    '<x:theme xmlns:x="http://schemas.openxmlformats.org/drawingml/2006/main" name="Brand">\n' +
    '  <x:themeElements>\n' +
    '    <x:clrScheme name="Brand">\n' +
    '      <x:dk1><x:sysClr val="windowText" lastClr="000000"/></x:dk1>\n' +
    '    </x:clrScheme>\n' +
    '  </x:themeElements>\n' +
    '</x:theme>';
  assert.equal(applyThemeOverrides(source, {}), source);
  assert.equal(applyThemeOverrides(source, {colors: {}, fonts: {}}), source);
});

test('a </clrScheme> inside a comment does not end the block the override replaces', () => {
  // `<clrScheme>[\s\S]*?</clrScheme>` terminated here, so the override was spliced into the middle of
  // the scheme and the real slots survived after it.
  const source =
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:clrScheme name="X">' +
    '<!-- was </a:clrScheme> before the redesign -->' +
    '<a:dk1><a:srgbClr val="111111"/></a:dk1>' +
    '</a:clrScheme>' +
    '</a:theme>';
  const xml = applyThemeOverrides(source, {colors: {accent1: 'BB2649'}});
  assert.equal(xml.match(/<a:dk1>/g)?.length, 1, 'the real dk1 was replaced, not duplicated');
  assert.ok(!xml.includes('was </a:clrScheme> before'), 'the comment went with the block body');
  assert.ok(xml.includes('<a:accent1><a:srgbClr val="BB2649"/></a:accent1>'));
});

test('an override reaches a <latin> written as an element pair, not only as an empty tag', () => {
  // `<a:latin\b)[^>]*(/>)` could not match `<a:latin ...></a:latin>`, so the override was dropped
  // with no error and the file kept the face it had.
  const source =
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:fontScheme><a:majorFont><a:latin typeface="Old" panose="020F0302"></a:latin>' +
    '</a:majorFont></a:fontScheme></a:theme>';
  const xml = applyThemeOverrides(source, {fonts: {major: 'New'}});
  assert.ok(xml.includes('typeface="New"'), 'the override landed');
  assert.ok(!xml.includes('typeface="Old"'), 'and replaced the face that was there');
});

test('an overridden typeface keeps the panose metric beside it', () => {
  // The doc comment claimed this for as long as it was untrue: the pattern captured everything after
  // the element name and re-emitted the typeface alone.
  const source =
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:fontScheme><a:minorFont><a:latin typeface="Old" panose="020F0502020204030204"/>' +
    '<a:ea typeface=""/></a:minorFont></a:fontScheme></a:theme>';
  const xml = applyThemeOverrides(source, {fonts: {minor: 'Aptos Narrow'}});
  assert.ok(xml.includes('panose="020F0502020204030204"'), 'panose survived the override');
  assert.ok(xml.includes('typeface="Aptos Narrow"'));
  assert.ok(xml.includes('<a:ea typeface=""/>'), 'and so did the east-asian face beside it');
});

test('a DrawingML prefix carrying regex metacharacters is matched literally', () => {
  // The prefix comes out of the part's own root element, and `.` and `-` are legal in an NCName. It
  // used to be interpolated into a `RegExp` unescaped, so `a.b:` matched `axb:` as well.
  const source =
    '<a.b:theme xmlns:a.b="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a.b:clrScheme name="X"><a.b:dk1><a.b:srgbClr val="111111"/></a.b:dk1></a.b:clrScheme>' +
    '</a.b:theme>';
  const xml = applyThemeOverrides(source, {colors: {accent1: 'BB2649'}});
  assert.ok(xml.includes('<a.b:accent1><a.b:srgbClr val="BB2649"/></a.b:accent1>'));
});
