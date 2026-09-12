import assert from 'node:assert/strict';
import {test} from 'node:test';

import {
  type CellValue,
  cellValueToText,
  coerceCellValue,
  detectValueType,
  isDataTableFormulaValue,
  isErrorCode,
  isErrorValue,
  isFormulaValue,
  isRichTextValue,
  isSharedFormulaValue,
  richTextToPlain,
  ValueType,
} from './value.ts';

test('detectValueType classifies each primitive', () => {
  assert.equal(detectValueType(null), ValueType.Null);
  assert.equal(detectValueType(42), ValueType.Number);
  assert.equal(detectValueType('hi'), ValueType.String);
  assert.equal(detectValueType(true), ValueType.Boolean);
  assert.equal(detectValueType(new Date(0)), ValueType.Date);
});

test('an invalid Date is still a Date-typed value', () => {
  assert.equal(detectValueType(new Date(Number.NaN)), ValueType.Date);
});

test('NaN and Infinity remain Number-typed: the model does not judge finiteness', () => {
  assert.equal(detectValueType(Number.NaN), ValueType.Number);
  assert.equal(detectValueType(Number.POSITIVE_INFINITY), ValueType.Number);
});

test('detectValueType classifies structural values', () => {
  assert.equal(detectValueType({error: '#REF!'}), ValueType.Error);
  assert.equal(detectValueType({formula: 'A1+B1'}), ValueType.Formula);
  assert.equal(detectValueType({formula: 'A1', result: 3}), ValueType.Formula);
  assert.equal(detectValueType({sharedFormula: 'A1', result: 3}), ValueType.Formula);
  assert.equal(detectValueType({richText: [{text: 'a'}]}), ValueType.RichText);
});

// A hyperlink lives on the sheet, beside the cell it covers, so the old `{hyperlink, text}` value is
// refused like any other shape the model does not know rather than stored as a label.
test('a hyperlink-shaped object is not a cell value', () => {
  const link = {hyperlink: 'https://x', text: 'x'} as unknown as CellValue;
  assert.throws(() => detectValueType(link), TypeError);
});

test('detectValueType throws on an unrecognised object shape', () => {
  assert.throws(() => detectValueType({nonsense: true} as unknown as CellValue), TypeError);
});

test('type guards discriminate the structural shapes', () => {
  assert.ok(isFormulaValue({formula: 'A1'}));
  assert.ok(!isFormulaValue({sharedFormula: 'A1'}));
  assert.ok(isSharedFormulaValue({sharedFormula: 'A1'}));
  assert.ok(isRichTextValue({richText: []}));
  assert.ok(isErrorValue({error: '#REF!'}));
  assert.ok(isDataTableFormulaValue({shareType: 'dataTable', ref: 'B2:B5'}));
});

test('type guards reject the primitive leaves and each other', () => {
  const primitives: readonly CellValue[] = [null, 42, 'text', true, new Date(0)];
  for (const guard of [
    isErrorValue,
    isFormulaValue,
    isSharedFormulaValue,
    isDataTableFormulaValue,
    isRichTextValue,
  ]) {
    for (const value of primitives)
      assert.ok(!guard(value), `${guard.name} accepted ${JSON.stringify(value)}`);
  }
  assert.ok(!isErrorValue({richText: []}));
  assert.ok(!isRichTextValue({error: '#REF!'}));
  assert.ok(!isDataTableFormulaValue({sharedFormula: 'A1'}));
});

test('cellValueToText renders every primitive leaf', () => {
  assert.equal(cellValueToText(null), '');
  assert.equal(cellValueToText(42), '42');
  assert.equal(cellValueToText(-0.5), '-0.5');
  assert.equal(cellValueToText('text'), 'text');
  assert.equal(cellValueToText(true), 'TRUE', "Excel's literal, not JavaScript's");
  assert.equal(cellValueToText(false), 'FALSE');
  assert.equal(cellValueToText(new Date(Date.UTC(2026, 0, 2))), '2026-01-02T00:00:00.000Z');
});

test('cellValueToText gives an invalid Date no text rather than throwing', () => {
  assert.equal(cellValueToText(new Date(Number.NaN)), '');
});

test('cellValueToText renders the structural kinds', () => {
  assert.equal(cellValueToText({error: '#REF!'}), '#REF!');
  assert.equal(cellValueToText({richText: [{text: 'foo'}, {text: 'bar'}]}), 'foobar');
});

test('cellValueToText renders a formula as its cached result, or nothing', () => {
  assert.equal(cellValueToText({formula: 'A1+B1', result: 3}), '3');
  assert.equal(cellValueToText({formula: 'A1+B1'}), '', 'no cached result is no text');
  assert.equal(cellValueToText({sharedFormula: 'A1', result: 'ok'}), 'ok');
  assert.equal(cellValueToText({shareType: 'dataTable', ref: 'B2:B5', result: 7}), '7');
  assert.equal(
    cellValueToText({formula: 'A1/0', result: {error: '#DIV/0!'}}),
    '#DIV/0!',
    'a cached error is the text the grid shows',
  );
});

test('cellValueToText applies no number format: the style is not the value', () => {
  assert.equal(cellValueToText(0.1 + 0.2), '0.30000000000000004');
});

test('cellValueToText rejects a value outside the union, as detectValueType does', () => {
  assert.throws(() => cellValueToText({nonsense: true} as unknown as CellValue), TypeError);
});

test('richTextToPlain concatenates every run in order, ignoring per-run formatting', () => {
  assert.equal(
    richTextToPlain({richText: [{text: 'foo'}, {text: 'bar', font: {bold: true}}, {text: 'baz'}]}),
    'foobarbaz',
  );
  assert.equal(richTextToPlain({richText: []}), '', 'no runs flattens to the empty string');
});

test('isErrorCode recognises the canonical literals only', () => {
  assert.ok(isErrorCode('#DIV/0!'));
  assert.ok(isErrorCode('#N/A'));
  assert.ok(!isErrorCode('#WHATEVER!'));
  assert.ok(!isErrorCode('42'));
});

test('coerceCellValue maps undefined to the empty cell', () => {
  assert.equal(coerceCellValue(undefined), null);
});

test('coerceCellValue preserves a numeric-looking string as a string', () => {
  const coerced = coerceCellValue('1000.80');
  assert.equal(coerced, '1000.80');
  assert.equal(detectValueType(coerced), ValueType.String);
});

test('coerceCellValue passes valid values through unchanged', () => {
  const date = new Date(0);
  assert.equal(coerceCellValue(date), date);
  const formula = {formula: 'A1'};
  assert.equal(coerceCellValue(formula), formula);
});

test('coerceCellValue rejects an unrecognised shape at the assignment site', () => {
  assert.throws(() => coerceCellValue({bogus: 1} as unknown as CellValue), TypeError);
});
