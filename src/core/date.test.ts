import assert from 'node:assert/strict';
import {test} from 'node:test';

import {DEFAULT_DATE_NUMFMT, dateToSerial, isDateFormat, serialToDate} from './date.ts';

const iso = (serial: number): string => serialToDate(serial, 1900).toISOString();

test('serial 1 reads as 1900-01-01, not 1899-12-31', () => {
  assert.equal(iso(1), '1900-01-01T00:00:00.000Z');
});

test('consecutive serials map to consecutive days', () => {
  assert.equal(iso(2), '1900-01-02T00:00:00.000Z');
  assert.equal(iso(3), '1900-01-03T00:00:00.000Z');
});

test('serial 59 is the real 1900-02-28 just below the phantom leap day', () => {
  assert.equal(iso(59), '1900-02-28T00:00:00.000Z');
});

test('serial 61 is 1900-03-01, the day after the phantom 1900-02-29', () => {
  assert.equal(iso(61), '1900-03-01T00:00:00.000Z');
});

test('a modern date round-trips through serial and back exactly', () => {
  const date = new Date('2020-03-04T00:00:00.000Z');
  assert.equal(serialToDate(dateToSerial(date, 1900), 1900).toISOString(), date.toISOString());
});

test('a fractional serial carries the time of day', () => {
  const noon = new Date('2020-03-04T12:00:00.000Z');
  const serial = dateToSerial(noon, 1900);
  assert.equal(serial % 1, 0.5, 'noon is half a day past midnight');
  assert.equal(serialToDate(serial, 1900).toISOString(), noon.toISOString());
});

test('dateToSerial reproduces the leap-year boundary in reverse', () => {
  assert.equal(dateToSerial(new Date('1900-01-01T00:00:00.000Z'), 1900), 1);
  assert.equal(dateToSerial(new Date('1900-02-28T00:00:00.000Z'), 1900), 59);
  assert.equal(dateToSerial(new Date('1900-03-01T00:00:00.000Z'), 1900), 61);
});

// The shift compared the instant with the phantom day, so any time after midnight on 1900-02-28 was
// 60 and a fraction, escaped it, and landed on the non-existent 1900-02-29. Reading 59.5 and writing
// it back changed the stored value to 60.5.
test('a time of day on 1900-02-28 converts to that day, and its serial round-trips', () => {
  assert.equal(dateToSerial(new Date('1900-02-28T12:00:00.000Z'), 1900), 59.5);
  assert.equal(dateToSerial(new Date('1900-02-28T23:59:59.000Z'), 1900) < 60, true);
  assert.equal(dateToSerial(serialToDate(59.5, 1900), 1900), 59.5);
  assert.equal(dateToSerial(new Date('1900-03-01T06:00:00.000Z'), 1900), 61.25);
});

// The numbers here are Excel Desktop's own, read back over COM from a workbook saved with the 1904
// system on and one with it off: 2023-03-15 is serial 45000 in the first and 43538 in the second.
// That difference is why the epoch is a parameter rather than a constant, and 1462 is the gap a
// workbook read under the wrong system lands by.
test('the 1904 system counts from 1904-01-01, with no phantom leap day', () => {
  assert.equal(serialToDate(0, 1904).toISOString(), '1904-01-01T00:00:00.000Z');
  assert.equal(serialToDate(1, 1904).toISOString(), '1904-01-02T00:00:00.000Z');
  assert.equal(serialToDate(59, 1904).toISOString(), '1904-02-29T00:00:00.000Z');
  assert.equal(serialToDate(60, 1904).toISOString(), '1904-03-01T00:00:00.000Z');
});

test('the same calendar date is 1462 serials apart between the two systems', () => {
  const date = new Date('2023-03-15T00:00:00.000Z');
  assert.equal(dateToSerial(date, 1900), 45000, 'the serial Excel stores in a 1900 workbook');
  assert.equal(dateToSerial(date, 1904), 43538, 'and in a 1904 one');
});

test('a date round-trips through the 1904 system exactly, times included', () => {
  for (const iso of ['1904-01-01T00:00:00.000Z', '2020-03-04T12:00:00.000Z']) {
    const date = new Date(iso);
    assert.equal(serialToDate(dateToSerial(date, 1904), 1904).toISOString(), iso);
  }
});

test('isDateFormat recognises date and time codes', () => {
  for (const code of [
    'yyyy-mm-dd',
    'DD/MM/YYYY',
    'mm-dd-yy',
    'h:mm:ss',
    '[$-409]mmmm d, yyyy',
    DEFAULT_DATE_NUMFMT,
  ]) {
    assert.equal(isDateFormat(code), true, `${code} should be a date format`);
  }
});

test('isDateFormat rejects number, currency, percent, and text codes', () => {
  for (const code of [
    'General',
    '0.00',
    '#,##0',
    '0.00%',
    '_("$"* #,##0.00_)',
    '@',
    '#,##0" days"',
  ]) {
    assert.equal(isDateFormat(code), false, `${code} should not be a date format`);
  }
});
