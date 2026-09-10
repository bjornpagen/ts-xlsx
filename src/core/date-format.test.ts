import assert from 'node:assert/strict';
import {test} from 'node:test';

import {formatSerialDate} from './date-format.ts';

test('formatSerialDate speaks the Excel format-code vocabulary, case-insensitively', () => {
  const moment = new Date(Date.UTC(2024, 2, 5, 14, 7, 9, 456));
  const render = (code: string) => formatSerialDate(moment, code, true);
  assert.equal(render('yyyy-mm-dd'), '2024-03-05');
  assert.equal(render('YYYY-MM-DD'), '2024-03-05', 'a code is case-insensitive, as Excel reads it');
  assert.equal(render('dd/mm/yyyy'), '05/03/2024');
  assert.equal(render('d/m/yy'), '5/3/24');
  assert.equal(render('d mmm yy'), '5 Mar 24');
  assert.equal(render('dddd, mmmm d, yyyy'), 'Tuesday, March 5, 2024');
  assert.equal(render('mmmmm'), 'M', 'five m is the initial Excel uses on a chart axis');
});

test('m is minutes beside an hour or a second, and a month anywhere else', () => {
  // The rule a token table cannot express, and the reason the CSV writer used to need a second,
  // incompatible date vocabulary of its own.
  const moment = new Date(Date.UTC(2024, 2, 5, 14, 7, 9));
  const render = (code: string) => formatSerialDate(moment, code, true);
  assert.equal(render('yyyy-mm-dd hh:mm:ss'), '2024-03-05 14:07:09');
  assert.equal(render('mm'), '03', 'alone, it is the month');
  assert.equal(render('hh:mm'), '14:07', 'after an hour, minutes');
  assert.equal(render('mm:ss'), '07:09', 'before seconds, minutes');
  assert.equal(render('h" hours "mm'), '14 hours 07', 'a literal between them does not break it');
});

test('an hour counts to twelve only when the code asks for a meridiem', () => {
  const afternoon = new Date(Date.UTC(2024, 2, 5, 14, 7, 0));
  const midnight = new Date(Date.UTC(2024, 2, 5, 0, 30, 0));
  assert.equal(formatSerialDate(afternoon, 'h:mm', true), '14:07');
  assert.equal(formatSerialDate(afternoon, 'h:mm AM/PM', true), '2:07 PM');
  assert.equal(formatSerialDate(midnight, 'h:mm AM/PM', true), '12:30 AM', 'midnight is 12, not 0');
  assert.equal(formatSerialDate(midnight, 'h:mm A/P', true), '12:30 A');
});

// Every meridiem rendered `AM`/`PM`, so `h:mm A/P` gave `9:30 AM` where Excel shows `9:30 A`. The
// expected strings are the `.Text` Excel Desktop displayed for each code at 09:30 and 21:30.
test('a meridiem renders as Excel renders it: the long form in capitals, the short form as spelled', () => {
  const morning = new Date(Date.UTC(2024, 2, 5, 9, 30, 0));
  const evening = new Date(Date.UTC(2024, 2, 5, 21, 30, 0));
  const cases: [string, string, string][] = [
    ['h:mm AM/PM', '9:30 AM', '9:30 PM'],
    ['h:mm am/pm', '9:30 AM', '9:30 PM'],
    ['h:mm Am/Pm', '9:30 AM', '9:30 PM'],
    ['h:mm A/P', '9:30 A', '9:30 P'],
    ['h:mm a/p', '9:30 a', '9:30 p'],
    ['h:mm A/p', '9:30 A', '9:30 p'],
  ];
  for (const [code, am, pm] of cases) {
    assert.equal(formatSerialDate(morning, code, true), am, code);
    assert.equal(formatSerialDate(evening, code, true), pm, code);
  }
});

test('quoted literals, escapes and bracketed directives are honoured', () => {
  const moment = new Date(Date.UTC(2024, 2, 5));
  const render = (code: string) => formatSerialDate(moment, code, true);
  assert.equal(render('"Q"m yyyy'), 'Q3 2024', 'a quoted literal is text, not placeholders');
  assert.equal(render('\\dyyyy'), 'd2024', 'a backslash escapes the letter after it');
  assert.equal(render('[$-409]mmmm d, yyyy'), 'March 5, 2024', 'a locale directive is dropped');
  assert.equal(render('[Red]yyyy'), '2024', 'and so is a colour');
  assert.equal(render('yyyy;"negative"'), '2024', 'only the first section applies to a date');
});

test('a fractional-seconds run renders the fraction rather than the zeros', () => {
  // `.0` after a seconds run is a placeholder. Emitting it verbatim, which a renderer that did not
  // know the form would do, reads as a plausible time and is silently wrong.
  const moment = new Date(Date.UTC(2024, 2, 5, 14, 7, 9, 456));
  assert.equal(formatSerialDate(moment, 'ss.0', true), '09.4');
  assert.equal(formatSerialDate(moment, 'hh:mm:ss.000', true), '14:07:09.456');
  assert.equal(
    formatSerialDate(moment, '0.0', true),
    '0.0',
    'and it is a placeholder only after s',
  );
  // The check read only the previous run's text, so an escaped or quoted `s` opened a fraction too.
  assert.equal(formatSerialDate(moment, 'ss\\s.0', true), '09s.0', 'an escaped s is a literal');
  assert.equal(formatSerialDate(moment, 'ss"s".0', true), '09s.0', 'and so is a quoted one');
});

test('an Invalid Date renders as nothing rather than as the words', () => {
  assert.equal(formatSerialDate(new Date(Number.NaN), 'yyyy-mm-dd', true), '');
});
