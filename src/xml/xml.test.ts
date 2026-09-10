import assert from 'node:assert/strict';
import {test} from 'node:test';

import {AuthoringError} from '../errors.ts';
import {
  escapeAttr,
  escapeSpreadsheetText,
  escapeText,
  intAttr,
  numAttr,
  textElement,
} from './xml.ts';

// The three classes of character XML 1.0 cannot carry, one representative each.
const CONTROL = '\u0001';
const NONCHARACTER = '\uFFFE';
const LONE_SURROGATE = '\uD800';

test('the five XML-syntactic characters are escaped as before', () => {
  assert.equal(escapeText('a & b < c > d'), 'a &amp; b &lt; c &gt; d');
  assert.equal(escapeAttr(`a & b < "c" 'd'`), 'a &amp; b &lt; &quot;c&quot; &apos;d&apos;');
});

test('tab, line feed and carriage return survive both paths', () => {
  assert.equal(escapeText('a\tb\nc\rd'), 'a\tb\nc\rd');
  assert.equal(escapeAttr('a\tb\nc\rd'), 'a&#9;b&#10;c&#13;d');
  assert.equal(escapeSpreadsheetText('a\tb\nc\rd'), 'a\tb\nc\rd');
});

test('U+007F and the C1 controls are legal XML 1.0 and pass through untouched', () => {
  assert.equal(escapeSpreadsheetText('a\u007Fb\u0085c\u009Fd'), 'a\u007Fb\u0085c\u009Fd');
  assert.equal(escapeText('a\u007Fb'), 'a\u007Fb');
  assert.equal(escapeAttr('a\u007Fb'), 'a\u007Fb');
});

test('a cell value escapes each unrepresentable class as _xHHHH_', () => {
  assert.equal(escapeSpreadsheetText(`a${CONTROL}b`), 'a_x0001_b');
  assert.equal(escapeSpreadsheetText(`a${NONCHARACTER}b`), 'a_xFFFE_b');
  assert.equal(escapeSpreadsheetText(`a${LONE_SURROGATE}b`), 'a_xD800_b');
});

test('a paired surrogate is one astral character, not two escapes', () => {
  assert.equal(escapeSpreadsheetText('a\u{1F600}b'), 'a\u{1F600}b');
});

test('a cell value that already reads as an escape has its underscore escaped', () => {
  assert.equal(escapeSpreadsheetText('a_x0041_b'), 'a_x005F_x0041_b');
});

test('escaping the underscore is idempotent under a second write', () => {
  const once = escapeSpreadsheetText('_x0041_');
  assert.equal(once, '_x005F_x0041_');
  assert.equal(escapeSpreadsheetText(once), '_x005F_x005F_x0041_');
});

test('text that only resembles an escape is left alone', () => {
  assert.equal(escapeSpreadsheetText('a_b'), 'a_b');
  assert.equal(escapeSpreadsheetText('a_xb'), 'a_xb');
  assert.equal(escapeSpreadsheetText('a_xZZZZ_b'), 'a_xZZZZ_b');
  assert.equal(escapeSpreadsheetText('a_x041_b'), 'a_x041_b');
});

test('the _xHHHH_ escape runs before the XML escape, not after', () => {
  assert.equal(escapeSpreadsheetText(`&${CONTROL}<`), '&amp;_x0001_&lt;');
});

test('a <t> element carries the cell-value escape', () => {
  assert.equal(textElement(`a${CONTROL}b`), '<t>a_x0001_b</t>');
});

test('escapeText refuses each unrepresentable class', () => {
  for (const char of [CONTROL, NONCHARACTER, LONE_SURROGATE]) {
    assert.throws(() => escapeText(`a${char}b`), AuthoringError);
  }
});

test('escapeAttr refuses each unrepresentable class', () => {
  for (const char of [CONTROL, NONCHARACTER, LONE_SURROGATE]) {
    assert.throws(() => escapeAttr(`a${char}b`), AuthoringError);
  }
});

test('the refusal names the code point and where it is', () => {
  assert.throws(
    () => escapeAttr(`Sheet${CONTROL}A`),
    (error: unknown) => {
      assert.ok(error instanceof AuthoringError);
      assert.match(error.message, /U\+0001/);
      assert.match(error.message, /offset 5/);
      return true;
    },
  );
});

test('numAttr omits an unset value and writes a finite one', () => {
  assert.equal(numAttr('scale', undefined), '');
  assert.equal(numAttr('scale', 0), ' scale="0"');
  assert.equal(numAttr('tint', -0.25), ' tint="-0.25"');
});

test('numAttr refuses a number OOXML cannot spell', () => {
  for (const value of [Number.NaN, Infinity, -Infinity]) {
    assert.throws(() => numAttr('scale', value), AuthoringError);
  }
});

test('intAttr omits an unset value and writes an integer at or above its floor', () => {
  assert.equal(intAttr('priority', undefined), '');
  assert.equal(intAttr('stdDev', -2), ' stdDev="-2"');
  assert.equal(intAttr('rank', 0, 0), ' rank="0"');
});

test('intAttr refuses a value an xsd:int or xsd:unsignedInt cannot hold', () => {
  for (const value of [Number.NaN, Infinity, 1.5, 2 ** 60]) {
    assert.throws(() => intAttr('priority', value), AuthoringError, String(value));
  }
  assert.throws(() => intAttr('rank', -1, 0), AuthoringError, 'below the floor');
});
