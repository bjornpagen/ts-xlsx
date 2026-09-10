// A token from a closed OOXML enumeration is checked on the way out and dropped on the way in.
//
// Every attribute here is typed as a union on the public surface, so an out-of-union value reaches
// the writer only through untyped JavaScript, a `JSON.parse`, or a cast. That is not a reason to
// interpolate it: the compiler having been bypassed is exactly the case the runtime is for, and the
// document a bogus token produces is not merely invalid but structurally broken, since nothing
// escapes the quote a caller can put in it.

import assert from 'node:assert/strict';
import {test} from 'node:test';

import {zipSync} from 'fflate';

import {ERROR_CODES} from '../../core/value.ts';
import {
  partIn,
  partsWritten as partsOf,
  patchParts,
  refuses,
  roundtrip,
  SHEET1,
  sheeted,
  sheetXml,
} from './package.test-support.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

// The shape of the attack: a value that closes the attribute and the element behind it.
const ESCAPE = 'x"/><x/><y a="';

test('<pageSetup> refuses a foreign orientation or page order', () => {
  refuses((wb) => {
    Object.assign(wb.getWorksheet('S')!.pageSetup, {orientation: ESCAPE});
  });
  refuses((wb) => {
    Object.assign(wb.getWorksheet('S')!.pageSetup, {pageOrder: 'sideways'});
  });
});

test('<dataValidation> refuses a foreign type, operator, error style or IME mode', () => {
  for (const facet of [
    {type: ESCAPE},
    {operator: 'sortOf'},
    {errorStyle: 'boom'},
    {imeMode: ESCAPE},
  ]) {
    refuses((wb) => {
      wb.getWorksheet('S')!.addDataValidation('A1', {type: 'whole', ...facet} as never);
    });
  }
});

test('<cfRule> refuses a foreign type, operator, time period, icon set or anchor type', () => {
  const rules = [
    {type: ESCAPE, priority: 1},
    {type: 'cellIs', operator: 'sortOf', priority: 1},
    {type: 'timePeriod', timePeriod: 'lastFortnight', priority: 1},
    {type: 'iconSet', iconSet: '9Arrows', priority: 1},
    {type: 'colorScale', cfvo: [{type: 'middling'}], priority: 1},
  ];
  for (const rule of rules) {
    refuses((wb) => {
      wb.getWorksheet('S')!.addConditionalFormatting({ref: 'A1:A5', rules: [rule]} as never);
    });
  }
});

test('a sheet tab and the document window refuse a foreign visibility', () => {
  refuses((wb) => {
    Object.assign(wb.getWorksheet('S')!, {state: ESCAPE});
  });
  refuses((wb) => {
    Object.assign(wb.view, {visibility: 'translucent'});
  });
});

test('a two-cell image anchor refuses a foreign edit mode', () => {
  refuses((wb) => {
    const id = wb.addImage({
      buffer: Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
      extension: 'png',
    });
    wb.getWorksheet('S')!.addImageAnchor(id, {
      from: {col: 0, row: 0},
      to: {col: 2, row: 2},
      editAs: ESCAPE as never,
    });
  });
});

test('<tableColumn> refuses a foreign totals-row function', () => {
  refuses((wb) => {
    const sheet = wb.getWorksheet('S')!;
    sheet.getCell('A1').value = 'h';
    sheet.getCell('A2').value = 1;
    sheet.addTable({
      name: 'T',
      ref: 'A1',
      columns: [{name: 'h', totalsRowFunction: ESCAPE}],
      rowCount: 1,
      totalsRow: true,
    } as never);
  });
});

test('<customFilter> refuses a foreign comparison operator', () => {
  refuses((wb) => {
    wb.getWorksheet('S')!.autoFilter = {
      ref: 'A1:A5',
      columns: [
        {
          colId: 0,
          criteria: {kind: 'custom', predicates: [{operator: ESCAPE, val: 'x'}]},
        },
      ],
    } as never;
  });
});

// A cell error's code goes into `<v>` unescaped, and `isErrorValue` asks only for an `error` key, so a
// code smuggled past the type wrote markup into the sheet: `{error: '</v></c><evil/>'}` closed the
// cell and opened an element of its own. A cached formula error takes the same path.
test('a cell error value and a cached formula error refuse a code outside ERROR_CODES', () => {
  refuses((wb) => {
    wb.getWorksheet('S')!.getCell('B1').value = {error: ESCAPE} as never;
  });
  refuses((wb) => {
    wb.getWorksheet('S')!.getCell('B1').value = {
      formula: '1/0',
      result: {error: '#WHATEVER!'},
    } as never;
  });
});

test('every error code still writes, as a value and as a cached formula result', () => {
  const wb = sheeted();
  const sheet = wb.getWorksheet('S')!;
  for (const [index, error] of ERROR_CODES.entries()) {
    sheet.getCell(`B${index + 1}`).value = {error};
    sheet.getCell(`C${index + 1}`).value = {formula: 'NA()', result: {error}};
  }

  const back = roundtrip(wb).getWorksheet('S');
  for (const [index, error] of ERROR_CODES.entries()) {
    assert.deepEqual(back?.getCell(`B${index + 1}`).value, {error});
    assert.deepEqual(back?.getCell(`C${index + 1}`).value, {formula: 'NA()', result: {error}});
  }
});

test('the tokens the enumerations do allow still round-trip', () => {
  const wb = sheeted();
  const sheet = wb.getWorksheet('S')!;
  sheet.pageSetup.orientation = 'landscape';
  sheet.pageSetup.pageOrder = 'overThenDown';
  sheet.addDataValidation('A1', {
    type: 'whole',
    operator: 'greaterThan',
    errorStyle: 'warning',
    formulae: [1],
  });
  sheet.addConditionalFormatting({
    ref: 'A1:A5',
    rules: [{type: 'timePeriod', timePeriod: 'lastWeek', priority: 1}],
  });
  wb.addWorksheet('Hidden', {state: 'veryHidden'});
  wb.view.visibility = 'hidden';

  const back = roundtrip(wb);
  const sheetBack = back.getWorksheet('S');
  assert.equal(sheetBack?.pageSetup.orientation, 'landscape');
  assert.equal(sheetBack?.pageSetup.pageOrder, 'overThenDown');
  assert.equal(sheetBack?.dataValidations[0]?.rule.operator, 'greaterThan');
  assert.equal(sheetBack?.conditionalFormattings[0]?.rules[0]?.timePeriod, 'lastWeek');
  assert.equal(back.getWorksheet('Hidden')?.state, 'veryHidden');
  assert.equal(back.view.visibility, 'hidden');
});

test('a foreign IME mode in a file is dropped on read, and the validation it sat on is kept', () => {
  const wb = sheeted();
  wb.getWorksheet('S')!.addDataValidation('A1', {type: 'whole', formulae: [1], imeMode: 'off'});
  const doctored = patchParts(writeXlsx(wb), {
    [SHEET1]: (xml) => xml.replace('imeMode="off"', 'imeMode="sideways"'),
  });

  const back = readXlsx(doctored);
  const rule = back.getWorksheet('S')?.dataValidationAt('A1');
  assert.equal(rule?.type, 'whole', 'the validation is kept');
  assert.equal(rule?.imeMode, undefined, 'the foreign IME mode is dropped');
  assert.doesNotThrow(() => writeXlsx(back), 'and what was read writes back');
});

test('a foreign token in a file is dropped on read rather than carried into a write that refuses it', () => {
  const wb = sheeted();
  wb.getWorksheet('S')!.pageSetup.orientation = 'landscape';
  const parts = partsOf(wb);
  const doctored = partIn(parts, 'xl/worksheets/sheet1.xml')
    .replace('orientation="landscape"', 'orientation="sideways"')
    .replace(
      '<pageSetup',
      '<conditionalFormatting sqref="A1:A5"><cfRule type="frobnicate" priority="1"/>' +
        '</conditionalFormatting><pageSetup',
    );
  const files: Record<string, Uint8Array> = {};
  for (const [name, text] of Object.entries(parts)) {
    files[name] = new TextEncoder().encode(name === 'xl/worksheets/sheet1.xml' ? doctored : text);
  }

  const back = readXlsx(zipSync(files));
  const sheet = back.getWorksheet('S');
  assert.equal(sheet?.pageSetup.orientation, undefined, 'the foreign orientation is dropped');
  assert.deepEqual(
    sheet?.conditionalFormattings[0]?.rules,
    [],
    'the foreign rule is dropped whole',
  );
  // Why dropped and not preserved: what the reader accepts, the writer must be able to write. The
  // block it leaves behind carries no rule, and `CT_ConditionalFormatting` requires one, so the
  // writer omits the element rather than emitting it empty.
  const rewritten = sheetXml(writeXlsx(back));
  assert.ok(!rewritten.includes('<conditionalFormatting'));
});
