import assert from 'node:assert/strict';
import {test} from 'node:test';

import {INTERNAL} from './internal.ts';
import {ValueType} from './value.ts';
import {type PreservedTheme, Workbook} from './workbook.ts';

test('addWorksheet appends sheets with stable, distinct ids', () => {
  const wb = new Workbook();
  const a = wb.addWorksheet('Alpha');
  const b = wb.addWorksheet('Beta');
  assert.equal(a.id, 1);
  assert.equal(b.id, 2);
  assert.deepEqual(
    wb.worksheets.map((s) => s.name),
    ['Alpha', 'Beta'],
  );
});

test('a new worksheet defaults to visible', () => {
  const wb = new Workbook();
  assert.equal(wb.addWorksheet('S').state, 'visible');
  assert.equal(wb.addWorksheet('H', {state: 'hidden'}).state, 'hidden');
});

test('getWorksheet finds sheets by case-insensitive name and by id', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('Data');
  assert.equal(wb.getWorksheet('data'), sheet);
  assert.equal(wb.getWorksheet('DATA'), sheet);
  assert.equal(wb.getWorksheet(1), sheet);
  assert.equal(wb.getWorksheet('missing'), undefined);
  assert.equal(wb.getWorksheet(99), undefined);
});

test('duplicate sheet names are rejected case-insensitively', () => {
  const wb = new Workbook();
  wb.addWorksheet('Sheet1');
  assert.throws(() => wb.addWorksheet('sheet1'), /already exists/);
});

test('invalid sheet names are rejected up front', () => {
  const wb = new Workbook();
  assert.throws(() => wb.addWorksheet(''), /cannot be empty/);
  assert.throws(() => wb.addWorksheet('a'.repeat(32)), /31-character limit/);
  assert.throws(() => wb.addWorksheet('a/b'), /forbids/);
  assert.throws(() => wb.addWorksheet('a:b'), /forbids/);
  assert.throws(() => wb.addWorksheet("'quoted'"), /apostrophe/);
});

test("a cell's col and row are 1-based numbers matching its position", () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  const cell = sheet.getCell('B3');
  cell.value = 'x';
  assert.equal(typeof cell.col, 'number');
  assert.equal(typeof cell.row, 'number');
  assert.equal(cell.col, 2);
  assert.equal(cell.row, 3);
  assert.equal(cell.address, 'B3');
  assert.equal(cell.type, ValueType.String);
});

test('getCell returns the same cell instance on repeat access', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  const first = sheet.getCell('A1');
  first.value = 7;
  assert.equal(sheet.getCell('A1'), first);
  assert.equal(sheet.getCell('$A$1').value, 7);
});

test('getCell rejects a whole-row or whole-column reference', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  assert.throws(() => sheet.getCell('1'), /not a single-cell reference/);
  assert.throws(() => sheet.getCell('A'), /not a single-cell reference/);
});

test('cells materialise lazily: only touched positions exist', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  assert.equal(sheet.hasCell(3, 2), false);
  sheet.getCell('B3');
  assert.equal(sheet.hasCell(3, 2), true);
  assert.equal(sheet.hasCell(1, 1), false);
});

test('assigning undefined clears a cell back to null/empty', () => {
  const wb = new Workbook();
  const cell = wb.addWorksheet('S').getCell('A1');
  cell.value = 42;
  cell.value = undefined;
  assert.equal(cell.value, null);
  assert.equal(cell.type, ValueType.Null);
});

test('requireWorksheet returns the sheet a partial lookup would', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('Data');
  assert.strictEqual(wb.requireWorksheet('Data'), sheet);
  assert.strictEqual(wb.requireWorksheet('DATA'), sheet, 'case-insensitive, as getWorksheet is');
  assert.strictEqual(wb.requireWorksheet(sheet.id), sheet, 'and by numeric id');
});

test('requireWorksheet names every sheet the workbook does have', () => {
  const wb = new Workbook();
  wb.addWorksheet('Summary');
  wb.addWorksheet('Raw data');
  assert.throws(
    () => wb.requireWorksheet('Sheet1'),
    /no worksheet "Sheet1"; this workbook has "Summary", "Raw data"/,
  );
  assert.throws(() => wb.requireWorksheet(99), /no worksheet id 99; this workbook has "Summary"/);
});

test('requireWorksheet says so when there are no sheets at all', () => {
  assert.throws(
    () => new Workbook().requireWorksheet('Any'),
    /no worksheet "Any": this workbook has no worksheets/,
  );
});

// Carrying pictures between workbooks: the attached-part half of a sheet copy, which the semantic
// `model` contract deliberately leaves alone (ADR-0005).

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]);
const GIF = new Uint8Array([0x47, 0x49, 0x46, 9, 9]);

function sheetShowingPictures() {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('Src');
  sheet.addImageAnchor(workbook.addImage({buffer: PNG}), {
    from: {col: 0, row: 0},
    ext: {cx: 100, cy: 50},
  });
  sheet.addBackgroundImage(workbook.addImage({buffer: GIF}));
  return {workbook, sheet};
}

test('exportImages resolves a sheet-anchored image id into the picture itself', () => {
  const {workbook, sheet} = sheetShowingPictures();

  const images = workbook.exportImages(sheet);

  assert.equal(images.anchored.length, 1);
  assert.deepEqual(images.anchored[0]?.image, {extension: 'png', data: PNG});
  assert.deepEqual(images.anchored[0]?.anchor, {from: {col: 0, row: 0}, ext: {cx: 100, cy: 50}});
  assert.deepEqual(images.background, {extension: 'gif', data: GIF});
});

test('exportImages refuses a sheet whose image ids this workbook never registered', () => {
  const {sheet} = sheetShowingPictures();

  assert.throws(() => new Workbook().exportImages(sheet), {
    name: 'AuthoringError',
    message: /worksheet "Src" shows image id 0, which is not registered/,
  });
});

test('importImages carries a picture into another workbook and re-anchors it there', () => {
  const {workbook: source, sheet: src} = sheetShowingPictures();
  const destination = new Workbook();
  // The destination already holds an unrelated picture, so a carried image cannot be assumed to
  // land on the id it had in the source.
  destination.addImage({buffer: new Uint8Array([0x42, 0x4d, 7])});
  const dst = destination.addWorksheet('Dst');

  destination.importImages(dst, source.exportImages(src));

  assert.deepEqual(destination.exportImages(dst), source.exportImages(src));
  assert.notEqual(dst.images[0]?.imageId, src.images[0]?.imageId, 'the id is rebound, not copied');
  assert.deepEqual(destination.getImage(dst.images[0]?.imageId ?? -1), {
    extension: 'png',
    data: PNG,
  });
});

test('importing the same pictures twice registers them once', () => {
  const {workbook: source, sheet: src} = sheetShowingPictures();
  const destination = new Workbook();
  const carried = source.exportImages(src);

  destination.importImages(destination.addWorksheet('One'), carried);
  destination.importImages(destination.addWorksheet('Two'), carried);

  assert.equal(destination.media.length, 2, 'one png and one gif, not two of each');
});

test('a picture that differs only in extension spelling is not registered twice', () => {
  const destination = new Workbook();
  const sheet = destination.addWorksheet('S');
  const anchor = {from: {col: 0, row: 0}, ext: {cx: 10, cy: 10}} as const;

  destination.importImages(sheet, {
    anchored: [
      {image: {extension: 'png', data: PNG}, anchor},
      {image: {extension: '.PNG', data: PNG}, anchor},
    ],
    background: undefined,
  });

  assert.equal(destination.media.length, 1);
  assert.deepEqual(
    sheet.images.map((image) => image.imageId),
    [0, 0],
  );
});

test("importImages replaces the destination sheet's pictures rather than adding to them", () => {
  const {workbook: source, sheet: src} = sheetShowingPictures();
  const destination = new Workbook();
  const dst = destination.addWorksheet('Dst');
  dst.addImageAnchor(destination.addImage({buffer: new Uint8Array([0x42, 0x4d, 7])}), {
    from: {col: 5, row: 5},
    ext: {cx: 1, cy: 1},
  });
  dst.addBackgroundImage(destination.addImage({buffer: new Uint8Array([0x42, 0x4d, 8])}));

  destination.importImages(dst, source.exportImages(src));

  assert.equal(dst.images.length, 1, 'the picture it held is gone, not kept alongside');
  assert.deepEqual(destination.getImage(dst.backgroundImageId ?? -1), {
    extension: 'gif',
    data: GIF,
  });
});

test('an import carrying no background clears the one the destination held', () => {
  const destination = new Workbook();
  const dst = destination.addWorksheet('Dst');
  dst.addBackgroundImage(destination.addImage({buffer: GIF}));

  destination.importImages(dst, {anchored: [], background: undefined});

  assert.equal(dst.backgroundImageId, undefined);
});

// ── Theme scheme resolution ─────────────────────────────────────────────────
// Two sources feed the scheme a `theme="n"` cell resolves against: the part the reader restored, and
// what a caller authored over it. Neither can be silently lost, because a wrong scheme fails
// nothing: every themed cell simply renders the old palette.

function themePart(): PreservedTheme {
  return {
    entryPath: 'xl/theme/theme1.xml',
    parts: [
      {
        path: 'xl/theme/theme1.xml',
        contentType: 'application/vnd.openxmlformats-officedocument.theme+xml',
        bytes: new TextEncoder().encode('<a:theme/>'),
        rels: [],
      },
    ],
  };
}

test('authoring a slot wins over the workbook scheme and leaves the rest alone', () => {
  const wb = new Workbook();
  assert.equal(wb.themeColors.accent1, '4472C4', 'precondition: the Office default');

  wb.setTheme({colors: {accent1: 'FF0000'}});
  assert.equal(wb.themeColors.accent1, 'FF0000');
  assert.equal(wb.themeColors.accent2, 'ED7D31', 'and the unauthored slots stand');
});

test("a restored part's scheme is what a theme reference resolves against", () => {
  const wb = new Workbook();
  assert.equal(wb.themeColors.accent1, '4472C4', 'precondition: the Office default');

  wb[INTERNAL].restoreThemePart(themePart(), {colors: {accent1: '00FF00'}, fonts: {}});
  assert.equal(wb.themeColors.accent1, '00FF00');
});

test('a part declaring no scheme falls back to the Office default rather than resolving nothing', () => {
  const wb = new Workbook();
  wb[INTERNAL].restoreThemePart(themePart(), {colors: {}, fonts: {}});
  assert.equal(wb.themeColors.accent1, '4472C4');
  assert.equal(wb.themeFonts.minor, 'Calibri');
});

test('a colour resolved through the theme follows a later setTheme', () => {
  const wb = new Workbook();
  assert.equal(wb.resolveColor({theme: 4}), 'FF4472C4');
  wb.setTheme({colors: {accent1: '112233'}});
  assert.equal(wb.resolveColor({theme: 4}), 'FF112233', 'theme="4" is accent1');
});

// setTheme validated a colour and then stored it as spelled, so `#FF0000` read back raw from
// themeColors and resolved to `FF#FF0000`, and `80FF0000` resolved to a ten-digit colour.
test('an authored theme colour is stored in the one spelling a slot holds', () => {
  const wb = new Workbook();
  wb.setTheme({colors: {accent1: '#FF0000', accent2: '80336699', accent3: 'aabbcc'}});
  assert.equal(wb.themeColors.accent1, 'FF0000');
  assert.equal(wb.themeColors.accent2, '336699');
  assert.equal(wb.themeColors.accent3, 'AABBCC');
  assert.deepEqual(wb.themeOverrides?.colors, {
    accent1: 'FF0000',
    accent2: '336699',
    accent3: 'AABBCC',
  });
  assert.equal(wb.resolveColor({theme: 4}), 'FFFF0000');
  assert.equal(wb.resolveColor({theme: 5}), 'FF336699');
  assert.match(wb.resolveColor({theme: 4, tint: 0.4}) ?? '', /^FF[0-9A-F]{6}$/);
});

test('setTheme refuses a colour it cannot read, natively and without storing any of the call', () => {
  const wb = new Workbook();
  // Native, not the library's taxonomy: one string that does not parse. An `XlsxError` subclass is
  // not a `SyntaxError`, so a later re-wrap reddens this rather than changing what a caller catches.
  assert.throws(
    () => wb.setTheme({colors: {accent1: '112233', accent2: 'rebeccapurple'}}),
    (error: unknown) => error instanceof SyntaxError,
  );
  assert.equal(wb.themeColors.accent1, '4472C4', 'the readable colour beside it was not stored');
  assert.equal(wb.themeOverrides, undefined);
});

test('activeTabIndex resolves a tab against the sheets that exist', () => {
  const wb = new Workbook();
  for (const name of ['One', 'Two', 'Three']) wb.addWorksheet(name);

  wb.view.activeTab = 2;
  assert.equal(wb.activeTabIndex, 2);
  // The first sheet is a legitimate answer arrived at two ways: the tab names it, or the tab names
  // nothing that exists and the fallback does. The guard used to spell the first as out of range,
  // which read as an off-by-one and was not one: both branches yield 0.
  wb.view.activeTab = 0;
  assert.equal(wb.activeTabIndex, 0);
  for (const tab of [-1, 3, 7, 1.5, Number.NaN]) {
    wb.view.activeTab = tab;
    assert.equal(wb.activeTabIndex, 0, `${tab} names no sheet`);
  }
});
