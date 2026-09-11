import assert from 'node:assert/strict';
import {test} from 'node:test';

import {spliceFormula, translateFormula} from './formula-references.ts';
import type {AxisSplice} from './grid-shift.ts';

test('translateFormula shifts a relative reference by the row and column delta', () => {
  assert.equal(translateFormula('A1*2', 0, 1), 'A2*2', 'one row down');
  assert.equal(translateFormula('A1*2', 0, 2), 'A3*2', 'two rows down');
  assert.equal(translateFormula('A1', 1, 0), 'B1', 'one column across');
  assert.equal(translateFormula('B2+C3', 2, 3), 'D5+E6', 'both axes, several references');
});

test('translateFormula answers #REF! on either axis when the shift leaves the grid', () => {
  // One question, two wrong answers before this: the column axis threw a bare `RangeError` out of
  // `numberToColumn`'s bounds assert, and the row axis emitted `A0` or `A-4`, which is not a
  // reference. The throw is the worse of the two, because this runs on the READ path with deltas
  // taken from a file's own shared-formula geometry, so an odd file aborted the whole sheet read with
  // an error outside the library's taxonomy. `#REF!` is what Excel writes for the same shift.
  assert.equal(translateFormula('XFD1', 1, 0), '#REF!', 'one column past the last');
  assert.equal(translateFormula('A1', -1, 0), '#REF!', 'one column before the first');
  assert.equal(translateFormula('A1*2', 0, -1), '#REF!*2', 'row 0 is not a row');
  assert.equal(translateFormula('A1*2', 0, -5), '#REF!*2', 'and neither is a negative one');
  assert.equal(translateFormula('A1048576', 0, 5), '#REF!', 'past the last row');
});

test('translateFormula answers #REF! for a reference the grid never had a column for', () => {
  // The guard above could not see these: `CELL_REFERENCE` matches three letters, so `ZZZ1` (column
  // 18278) went into the strict decoder and threw before the shift was ever computed. A reference
  // past XFD and a shift past XFD are the same answer, and one of them used to abort the read.
  assert.equal(translateFormula('ZZZ1*2', 0, 1), '#REF!*2');
  assert.equal(translateFormula('$ZZZ$1*2', 0, 1), '#REF!*2', 'an anchor does not exempt it');
  assert.equal(translateFormula('SUM(A1,ZZZ1)', 1, 0), 'SUM(B1,#REF!)', 'one operand at a time');
});

test('translateFormula leaves the references that stay on the grid alone', () => {
  assert.equal(translateFormula('XFC1', 1, 0), 'XFD1', 'the last column is reachable');
  assert.equal(translateFormula('A1048571', 0, 5), 'A1048576', 'and so is the last row');
  assert.equal(translateFormula('SUM(A1,XFD1)', 1, 0), 'SUM(B1,#REF!)', 'one operand at a time');
});

test('translateFormula leaves an absolute axis fixed and shifts only the relative one', () => {
  assert.equal(translateFormula('$A$1', 3, 4), '$A$1', 'fully absolute never moves');
  assert.equal(translateFormula('$A1', 5, 1), '$A2', 'absolute column, relative row');
  assert.equal(translateFormula('A$1', 1, 5), 'B$1', 'relative column, absolute row');
  assert.equal(translateFormula('$A$1+B1', 1, 1), '$A$1+C2', 'mixed within one formula');
});

test('translateFormula is the identity for a zero delta', () => {
  assert.equal(translateFormula('SUM($A$1:B7)*C8', 0, 0), 'SUM($A$1:B7)*C8');
});

test('translateFormula shifts both endpoints of a range independently', () => {
  assert.equal(translateFormula('SUM(A1:B2)', 1, 10), 'SUM(B11:C12)');
  assert.equal(translateFormula('A1:$B$2', 0, 5), 'A6:$B$2', 'the absolute endpoint stays');
});

// Filling `COUNTIF(A:A,A1)` right is exactly what Excel stores as a shared formula, and the clone used
// to read back as `COUNTIF(A:A,B1)`: only `A1`-shaped references moved, so a whole-line range never did.
test('translateFormula shifts a whole-column range, each end only where it is relative', () => {
  assert.equal(
    translateFormula('COUNTIF(A:A,A1)', 1, 1),
    'COUNTIF(B:B,B2)',
    'the range moves with the cell beside it',
  );
  assert.equal(translateFormula('SUM(A:A)', 1, 0), 'SUM(B:B)');
  assert.equal(translateFormula('SUM($A:$A)', 1, 0), 'SUM($A:$A)', 'anchored at both ends');
  assert.equal(translateFormula('SUM(A:$C)', 1, 0), 'SUM(B:$C)', 'anchored at one end');
  assert.equal(translateFormula('SUM(A:A)', 0, 5), 'SUM(A:A)', 'a row shift cannot move a column');
});

test('translateFormula shifts a whole-row range, each end only where it is relative', () => {
  assert.equal(translateFormula('SUM(1:1)', 0, 1), 'SUM(2:2)');
  assert.equal(translateFormula('SUM($1:2)', 0, 1), 'SUM($1:3)', 'anchored at one end');
  assert.equal(translateFormula('SUM(1:1)', 3, 0), 'SUM(1:1)', 'a column shift cannot move a row');
});

test('translateFormula shifts a sheet-qualified whole-line range but not the sheet name', () => {
  assert.equal(translateFormula('SUM(Sheet1!A:A)', 1, 0), 'SUM(Sheet1!B:B)');
  assert.equal(translateFormula("SUM('My Sheet'!1:1)", 0, 1), "SUM('My Sheet'!2:2)");
});

test('translateFormula answers #REF! for a whole-line range pushed off the grid', () => {
  assert.equal(translateFormula('SUM(XFD:XFD)', 1, 0), 'SUM(#REF!)', 'past the last column');
  assert.equal(translateFormula('SUM(1:1)', 0, -1), 'SUM(#REF!)', 'before the first row');
  assert.equal(translateFormula('SUM(1048576:1048576)', 0, 1), 'SUM(#REF!)', 'past the last row');
});

test('translateFormula leaves a time in a string alone and still reads a cell range as two cells', () => {
  assert.equal(translateFormula('IF(A1="10:30",1,0)', 0, 1), 'IF(A2="10:30",1,0)');
  assert.equal(translateFormula('SUM(A1:B2)', 1, 1), 'SUM(B2:C3)');
});

test('translateFormula never touches a function name or a defined name', () => {
  assert.equal(translateFormula('SUM(A1:A3)', 0, 1), 'SUM(A2:A4)', 'SUM has no row digits');
  assert.equal(translateFormula('TaxRate*A1', 2, 2), 'TaxRate*C3', 'a defined name is left alone');
  assert.equal(
    translateFormula('LOG10(A1)', 0, 1),
    'LOG10(A2)',
    'a call ending in digits is not a reference',
  );
});

test('translateFormula shifts a sheet-qualified cell but not the sheet name', () => {
  assert.equal(translateFormula('Sheet1!A1', 0, 1), 'Sheet1!A2', 'the cell after ! moves');
  assert.equal(
    translateFormula('Q1!A1', 0, 1),
    'Q1!A2',
    'a sheet name that looks like a reference is untouched',
  );
  assert.equal(
    translateFormula("'My Sheet'!A1+B2", 1, 1),
    "'My Sheet'!B2+C3",
    'a quoted sheet name is copied verbatim',
  );
});

test('translateFormula reads a bare cell-shaped name before a colon as a cell, as Excel does', () => {
  // `Q1:Q4!B2` looks like a 3-D span over quarter-named sheets, but a span over sheets named like
  // cells must be quoted, `'Q1:Q4'!B2`. Excel reads the bare form as the range from cell Q1 to Q4!B2
  // and re-spells it `Q1:'Q4'!B2` (Excel 16.0 build 20326), so Q1 is a relative cell and moves.
  assert.equal(translateFormula('SUM(Q1:Q4!B2)', 0, 1), 'SUM(Q2:Q4!B3)');
  assert.equal(translateFormula("SUM('Q1:Q4'!B2)", 0, 1), "SUM('Q1:Q4'!B3)", 'the quoted span');
});

test('translateFormula copies a string literal verbatim, references outside it still move', () => {
  assert.equal(translateFormula('IF(A1>0,"A1 is B2",B2)', 0, 1), 'IF(A2>0,"A1 is B2",B3)');
});

// Every splice expectation below is what Excel 16.0 (build 20326) wrote back through `Range.Formula`
// after the same edit made over COM, on a workbook of sheets S1, S2 and S3, unless a comment says
// otherwise.

const insertRows = (start: number, count: number): AxisSplice => ({
  axis: 'row',
  start,
  count: 0,
  delta: count,
});
const deleteRows = (start: number, count: number): AxisSplice => ({
  axis: 'row',
  start,
  count,
  delta: -count,
});
const deleteColumns = (start: number, count: number): AxisSplice => ({
  axis: 'col',
  start,
  count,
  delta: -count,
});

const onS1 = (formula: string, splice: AxisSplice) =>
  spliceFormula(formula, 'S1', {sheet: 'S1', splice});

test('an insert shifts every reference at or after it, absolute or not, and grows a range it lands in', () => {
  const insert = insertRows(5, 2);
  assert.equal(onS1('SUM(A1:A10)', insert), 'SUM(A1:A12)');
  assert.equal(onS1('A5+$A$5+A$5+$A5', insert), 'A7+$A$7+A$7+$A7');
  assert.equal(onS1('SUM(5:6)+SUM(A:A)+SUM(4:4)', insert), 'SUM(7:8)+SUM(A:A)+SUM(4:4)');
  assert.equal(onS1('SUM(A11:A12)', insert), 'SUM(A13:A14)');
});

test('an insert at a range top moves the range, and one below its bottom leaves it', () => {
  assert.equal(onS1('SUM(A5:A10)', insertRows(5, 1)), 'SUM(A6:A11)');
  assert.equal(onS1('SUM(A1:A4)', insertRows(5, 1)), 'SUM(A1:A4)');
});

test('a reference the delete takes is #REF!, and a range it cuts into shrinks', () => {
  const remove = deleteRows(5, 2);
  assert.equal(onS1('SUM(A3:A10)', remove), 'SUM(A3:A8)');
  assert.equal(onS1('A5+A7+A12', remove), '#REF!+A5+A10');
  assert.equal(onS1('SUM(A5:A6)', remove), 'SUM(#REF!)');
  assert.equal(onS1('SUM(A6:A12)', remove), 'SUM(A5:A10)');
  assert.equal(onS1('SUM(5:6)+SUM(4:7)', remove), 'SUM(#REF!)+SUM(4:5)');
  // Not in the probe: the same rule from the other side, a range whose bottom the delete takes.
  assert.equal(onS1('SUM(A3:A5)', remove), 'SUM(A3:A4)');
});

test('a splice that deletes and inserts at once reads as the delete, then the insert', () => {
  // Excel has no single edit that does both, so this is two of its edits composed: rows 5 and 6 out,
  // then three rows in at 5. `A6:A12` is `A5:A10` after the delete, and the insert at its top then
  // moves it whole.
  const splice: AxisSplice = {axis: 'row', start: 5, count: 2, delta: 1};
  assert.equal(onS1('SUM(A6:A12)', splice), 'SUM(A8:A13)');
  assert.equal(onS1('SUM(A3:A10)+A5+SUM(A1:A4)', splice), 'SUM(A3:A11)+#REF!+SUM(A1:A4)');
});

test('a column delete follows the same rules on the other axis', () => {
  assert.equal(
    onS1('SUM(B1:E1)+C1+F1+SUM(C:D)+SUM(C1:D1)', deleteColumns(3, 2)),
    'SUM(B1:C1)+#REF!+D1+SUM(#REF!)+SUM(#REF!)',
  );
  assert.equal(onS1('SUM(5:6)', deleteColumns(1, 1)), 'SUM(5:6)', 'a whole row has no column');
});

test('a reference pushed past the last row is #REF!, and a range bottom already there stays', () => {
  const insert = insertRows(3, 1);
  assert.equal(onS1('A1048576', insert), '#REF!');
  assert.equal(onS1('SUM(A1048570:A1048576)', insert), 'SUM(A1048571:A1048576)');
});

test("another sheet's formula moves a reference qualified with the spliced sheet, and only that one", () => {
  const edit = {sheet: 'S1', splice: insertRows(5, 2)};
  assert.equal(
    spliceFormula("'S1'!A5+SUM('S1'!A1:A10)+A5", 'S2', edit),
    "'S1'!A7+SUM('S1'!A1:A12)+A5",
  );
  assert.equal(
    spliceFormula('Data!B9+Other!B9', 'S2', {sheet: 'data', splice: insertRows(1, 1)}),
    'Data!B10+Other!B9',
    'a sheet name matches in any case',
  );
});

test('a deleted qualified reference keeps its sheet', () => {
  assert.equal(
    spliceFormula("'S1'!A5+SUM('S1'!A5:A6)+'S1'!A8", 'S2', {sheet: 'S1', splice: deleteRows(5, 2)}),
    "'S1'!#REF!+SUM('S1'!#REF!)+'S1'!A6",
  );
});

test('a 3-D span, a string, a structured reference and an external reference do not move', () => {
  const edit = {sheet: 'S1', splice: insertRows(5, 2)};
  assert.equal(spliceFormula("SUM(S1:'S3'!A5)", 'S2', edit), "SUM(S1:'S3'!A5)");
  assert.equal(spliceFormula('INDIRECT("S1!A5")&"A5"', 'S1', edit), 'INDIRECT("S1!A5")&"A5"');
  assert.equal(spliceFormula('SUM(Table1[A5])+A5', 'S1', edit), 'SUM(Table1[A5])+A7');
  assert.equal(spliceFormula("[1]S1!A5+'[2]S1'!A5", 'S1', edit), "[1]S1!A5+'[2]S1'!A5");
});

test('a function name, a defined name and a number are not references', () => {
  const insert = insertRows(1, 1);
  assert.equal(
    onS1('LOG10(A1)+TaxRate+A1B+1.5E3+R1C1', insert),
    'LOG10(A2)+TaxRate+A1B+1.5E3+R1C1',
  );
  assert.equal(onS1('RC1', insert), 'RC2', 'RC1 is column RC, row 1');
});

test('a formula on a sheet the splice did not touch, naming no sheet, is left as it is', () => {
  const formula = 'SUM(A1:A10)*$B$5';
  assert.equal(spliceFormula(formula, 'S2', {sheet: 'S1', splice: insertRows(1, 3)}), formula);
});

test('a formula with no home sheet moves only what names the spliced sheet', () => {
  // A workbook-scoped defined name is written fully qualified, as `'S1'!$A$5` is.
  const edit = {sheet: 'S1', splice: insertRows(5, 2)};
  assert.equal(spliceFormula("'S1'!$A$5", undefined, edit), "'S1'!$A$7");
  assert.equal(spliceFormula('$A$5', undefined, edit), '$A$5');
});
