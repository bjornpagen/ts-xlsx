// A number that cannot be spelled in OOXML is refused at the moment it would become bytes.
//
// Every attribute exercised here is `xsd:double` or `xsd:unsignedInt` in ECMA-376, and each sits on
// a plain data shape (`PageSetup`, `PageBreak`, `Color`, `ConditionalFormattingRule`, `Extent`,
// `WorkbookView`, row/column properties) that the model stores verbatim with no accessor to validate
// through. So the write path is the one place the mistake can be caught, and it is caught per
// family: the helper being correct is not evidence a call site adopted it.

import assert from 'node:assert/strict';
import {test} from 'node:test';

import type {DefinedName} from '../../core/workbook.ts';
import {partIn, partsWritten, refuses, sheeted} from './package.test-support.ts';
import {writeXlsx} from './write.ts';

const UNWRITABLE = [Number.NaN, Infinity, -Infinity] as const;

test('<pageSetup> refuses a non-finite scaling attribute', () => {
  for (const value of UNWRITABLE) {
    refuses((wb) => {
      wb.getWorksheet('S')!.pageSetup.scale = value;
    });
    refuses((wb) => {
      wb.getWorksheet('S')!.pageSetup.fitToWidth = value;
    });
  }
});

test('<brk> refuses a non-finite row or column index', () => {
  refuses((wb) => {
    wb.getWorksheet('S')!.rowBreaks.push({id: Number.NaN});
  });
  refuses((wb) => {
    wb.getWorksheet('S')!.columnBreaks.push({id: 1, max: Infinity});
  });
});

test('a colour refuses a non-finite theme, tint or index', () => {
  for (const color of [{theme: Number.NaN}, {tint: Infinity}, {indexed: -Infinity}]) {
    refuses((wb) => {
      wb.getWorksheet('S')!.getCell('A1').fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: color,
      };
    });
  }
});

test('<cfRule> refuses a non-finite rank or standard deviation', () => {
  refuses((wb) => {
    wb.getWorksheet('S')!.addConditionalFormatting({
      ref: 'A1:A5',
      rules: [{type: 'top10', rank: Number.NaN, priority: 1}],
    });
  });
  refuses((wb) => {
    wb.getWorksheet('S')!.addConditionalFormatting({
      ref: 'A1:A5',
      rules: [{type: 'aboveAverage', stdDev: Infinity, priority: 1}],
    });
  });
});

test('a picture refuses a non-finite extent or rotation', () => {
  const png = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  refuses((wb) => {
    const id = wb.addImage({buffer: png, extension: 'png'});
    // The EMU-level anchor, because the pixel-taking `addImage` would convert the bad value first
    // and the attribute is what is under test.
    wb.getWorksheet('S')!.addImageAnchor(id, {
      from: {col: 0, row: 0},
      ext: {cx: Number.NaN, cy: 100},
    });
  });
  refuses((wb) => {
    const id = wb.addImage({buffer: png, extension: 'png'});
    wb.getWorksheet('S')!.addImageAnchor(id, {
      from: {col: 0, row: 0},
      to: {col: 2, row: 2},
      rotation: Infinity,
    });
  });
});

test('an anchor grid point is refused on the same terms as the extent', () => {
  const png = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  refuses((wb) => {
    const id = wb.addImage({buffer: png, extension: 'png'});
    wb.getWorksheet('S')!.addImageAnchor(id, {
      from: {col: 0, row: 0, rowOff: Number.NaN},
      ext: {cx: 100, cy: 100},
    });
  });
});

test('<workbookView> refuses a non-finite window rect', () => {
  refuses((wb) => {
    wb.view.width = Number.NaN;
  });
});

test('an outline level is refused rather than dropped when it is not a number', () => {
  for (const value of UNWRITABLE) {
    refuses((wb) => {
      wb.getWorksheet('S')!.getColumn(2).outlineLevel = value;
    });
    refuses((wb) => {
      wb.getWorksheet('S')!.getRow(2).outlineLevel = value;
    });
  }
});

// A finite level slips past the non-finite gate and is the worse hazard of the two: the writer finds
// a group's end by walking outwards while the neighbouring level exceeds the summary's, every
// unmapped row answers `0`, and `0 > -1` holds for every row on the sheet.
test('an outline level is refused when it is negative or fractional, not only when it is not a number', () => {
  for (const value of [-1, -0.5, 2.5]) {
    refuses((wb) => {
      wb.getWorksheet('S')!.getRow(2).outlineLevel = value;
    });
    refuses((wb) => {
      wb.getWorksheet('S')!.getColumn(2).outlineLevel = value;
    });
  }
});

test('a pivot cache blanks a non-finite source value rather than spelling it', () => {
  // `scalarOf` blanks a non-finite cell before it can reach the cache, so this asserts that guard
  // holds from the outside: the write succeeds, and neither the shared item nor the field's
  // min/max carries the token `NaN` into an attribute OOXML has no lexical form for.
  const workbook = sheeted();
  const source = workbook.getWorksheet('S')!;
  source.getCell('A1').value = 'Region';
  source.getCell('B1').value = 'Quarter';
  source.getCell('C1').value = 'Amount';
  source.getCell('A2').value = 'West';
  source.getCell('B2').value = 'Q1';
  source.getCell('C2').value = Number.NaN;
  source.getCell('A3').value = 'East';
  source.getCell('B3').value = 'Q2';
  source.getCell('C3').value = 4;
  workbook
    .addWorksheet('P')
    .addPivotTable({source, rows: ['Region'], columns: ['Quarter'], values: ['Amount']});

  const parts = partsWritten(workbook);
  for (const [path, xml] of Object.entries(parts)) {
    assert.ok(!/NaN|Infinity/.test(xml), `${path} carries an unwritable number`);
  }
});

test('a defined name scoped to a sheet the workbook does not have is refused', () => {
  const workbook = sheeted();
  // `definedNames` is live, so this is the route a name reaches the writer without passing
  // `defineName`'s scope check. `localSheetId` is an `xsd:unsignedInt`; the `-1` an unguarded lookup
  // miss produces is a package Excel offers to repair.
  (workbook.definedNames as DefinedName[]).push({
    name: 'Local',
    refersTo: 'Gone!$A$1',
    scope: 'Gone',
  });

  assert.throws(() => writeXlsx(workbook), {
    name: 'AuthoringError',
    message: /"Local".*"Gone"/,
  });
});

test('a scope matching its sheet only in case still resolves to that sheet', () => {
  // `defineName` accepts the scope case-insensitively, the way sheet names are identified
  // everywhere else, so the writer must resolve it the same way rather than refuse it.
  const workbook = sheeted();
  workbook.addWorksheet('Data').getCell('A1').value = 1;
  workbook.defineName({name: 'Local', refersTo: 'Data!$A$1', scope: 'dAtA'});

  const workbookPart = partIn(partsWritten(workbook), 'xl/workbook.xml');
  assert.match(workbookPart, /<definedName name="Local" localSheetId="1">/);
});
