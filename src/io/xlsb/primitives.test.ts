import assert from 'node:assert/strict';
import {test} from 'node:test';

import {XlsbParseError} from './errors.ts';
import {errorCodeFor, RecordReader} from './primitives.ts';
import {wide, word} from './records.test-support.ts';

function bytes(...values: number[]): Uint8Array {
  return Uint8Array.from(values);
}

test('fixed-width integers and doubles read little-endian and advance the cursor', () => {
  const reader = new RecordReader(bytes(0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0xff, 0xff));
  assert.equal(reader.u8(), 0x01);
  assert.equal(reader.u16(), 0x0302);
  assert.equal(reader.u32(), 0x07060504);
  assert.equal(reader.i16(), -1);
  assert.equal(reader.done, true);
});

test('reading past the end of a record throws rather than returning garbage', () => {
  const reader = new RecordReader(bytes(1, 2, 3));
  assert.throws(() => reader.u32(), XlsbParseError);
});

test('an RkNumber in float form reproduces the double Excel wrote', () => {
  // Both words are lifted from a workbook Excel itself saved: B2 = 10 and B8 = 1.23. The second
  // carries the fX100 flag, which is the whole reason the encoding exists.
  assert.equal(new RecordReader(word(0x40240000)).rk(), 10);
  assert.equal(new RecordReader(word(0x405ec001)).rk(), 1.23);
});

test('an RkNumber in integer form sign-extends its 30-bit payload', () => {
  const asInt = (value: number): number => (value << 2) | 0b10;
  assert.equal(new RecordReader(word(asInt(0))).rk(), 0);
  assert.equal(new RecordReader(word(asInt(100))).rk(), 100);
  assert.equal(new RecordReader(word(asInt(-5))).rk(), -5);
  assert.equal(new RecordReader(word(asInt(-536870912))).rk(), -536870912);
  assert.equal(new RecordReader(word(asInt(536870911))).rk(), 536870911);
});

test('an RkNumber marked fX100 divides by 100 in both the integer and float forms', () => {
  assert.equal(new RecordReader(word(((123 << 2) | 0b11) >>> 0)).rk(), 1.23);
  assert.equal(new RecordReader(word(0x405ec001)).rk(), 1.23);
});

test('an XLWideString decodes its UTF-16LE units, including non-Latin and surrogate pairs', () => {
  for (const text of ['', 'Amount', 'Ünïcodé', '日本語', '𝄞 clef']) {
    assert.equal(new RecordReader(wide(text)).wideString(), text, text);
  }
});

test('a string longer than one decoding batch reassembles exactly', () => {
  const text = 'x'.repeat(10000);
  assert.equal(new RecordReader(wide(text)).wideString(), text);
});

test('an XLNullableWideString distinguishes an absent string from an empty one', () => {
  assert.equal(new RecordReader(word(0xffffffff)).nullableWideString(), undefined);
  assert.equal(new RecordReader(wide('')).nullableWideString(), '');
  assert.equal(new RecordReader(wide('rId1')).nullableWideString(), 'rId1');
});

test('a forged character count is rejected before any string is materialised', () => {
  // The attack this guard exists for: four bytes of header claiming a billion characters.
  const reader = new RecordReader(bytes(0xff, 0xff, 0xff, 0x3f, 0x41, 0x00));
  assert.throws(() => reader.wideString(), XlsbParseError);
});

test('a RichStr yields its text and leaves the run tail unread', () => {
  const reader = new RecordReader(Uint8Array.from([0x00, ...wide('widget')]));
  assert.equal(reader.richString(), 'widget');
});

test('a Cell header separates the column from the style index and drops the phonetic flag', () => {
  // column 1, iStyleRef 3, with the high byte (fPhShow + reserved) set to prove it is masked off.
  const reader = new RecordReader(Uint8Array.from([...word(1), ...word(0xff000003)]));
  assert.deepEqual(reader.cell(), {column: 1, styleIndex: 3});
});

test('an UncheckedRfX reads four zero-based inclusive bounds', () => {
  const reader = new RecordReader(
    Uint8Array.from([...word(0), ...word(7), ...word(0), ...word(1)]),
  );
  assert.deepEqual(reader.range(), {rowFirst: 0, rowLast: 7, colFirst: 0, colLast: 1});
});

test('a BrtColor decodes each of its four encodings onto the model colour', () => {
  const color = (...values: number[]): ReturnType<RecordReader['color']> =>
    new RecordReader(bytes(...values)).color();
  // Automatic: the record names no colour at all, so the cell gains none.
  assert.equal(color(0x00, 0x00, 0, 0, 0, 0, 0, 0), undefined);
  // Indexed 64: the automatic-background placeholder Excel writes on every solid fill.
  assert.deepEqual(color(0x03, 0x40, 0, 0, 0, 0, 0, 0), {indexed: 64});
  // ARGB, straight from a fill Excel wrote for interior colour 0xEEDDCC (BGR) → FFCCDDEE.
  assert.deepEqual(color(0x05, 0xff, 0, 0, 0xcc, 0xdd, 0xee, 0xff), {argb: 'FFCCDDEE'});
  // Theme index 1, the default font colour.
  assert.deepEqual(color(0x07, 0x01, 0, 0, 0, 0, 0, 0xff), {theme: 1});
});

test('a BrtColor tint is carried as the OOXML fraction and omitted when zero', () => {
  const reader = new RecordReader(bytes(0x07, 0x01, 0xff, 0x7f, 0, 0, 0, 0xff));
  assert.deepEqual(reader.color(), {theme: 1, tint: 1});
  const darkened = new RecordReader(bytes(0x07, 0x01, 0x00, 0x80, 0, 0, 0, 0xff));
  assert.deepEqual(darkened.color(), {theme: 1, tint: -1});
});

test('every BErr code maps to its error string and an unknown code maps to none', () => {
  assert.equal(errorCodeFor(0x00), '#NULL!');
  assert.equal(errorCodeFor(0x07), '#DIV/0!');
  assert.equal(errorCodeFor(0x0f), '#VALUE!');
  assert.equal(errorCodeFor(0x17), '#REF!');
  assert.equal(errorCodeFor(0x1d), '#NAME?');
  assert.equal(errorCodeFor(0x24), '#NUM!');
  assert.equal(errorCodeFor(0x2a), '#N/A');
  assert.equal(errorCodeFor(0x99), undefined);
});
