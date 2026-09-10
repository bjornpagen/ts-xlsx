import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Workbook} from '../../core/workbook.ts';
import {
  optionalPartText,
  partText,
  patchParts,
  roundtrip,
  sheetXml,
} from './package.test-support.ts';
import {parseSharedStrings} from './read-shared-strings.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

function bookWithStrings(...values: string[]): Workbook {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('S');
  values.forEach((value, i) => {
    sheet.getCell(`A${i + 1}`).value = value;
  });
  return workbook;
}

test('useSharedStrings writes a sharedStrings part and stores the cell as a t="s" reference', () => {
  const pkg = writeXlsx(bookWithStrings('hello'), {useSharedStrings: true});

  assert.match(partText(pkg, 'xl/sharedStrings.xml'), /<si><t>hello<\/t><\/si>/);
  const sheet = partText(pkg, 'xl/worksheets/sheet1.xml');
  assert.match(sheet, /t="s"><v>0<\/v>/);
  assert.doesNotMatch(sheet, /inlineStr/);
});

test('without the option, strings stay inline and no sharedStrings part is written', () => {
  const pkg = writeXlsx(bookWithStrings('hello'));

  assert.equal(optionalPartText(pkg, 'xl/sharedStrings.xml'), undefined);
  assert.match(partText(pkg, 'xl/worksheets/sheet1.xml'), /t="inlineStr"><is><t>hello<\/t>/);
});

test('an enabled workbook with no string cells never fabricates an empty sharedStrings part', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').getCell('A1').value = 42;
  const pkg = writeXlsx(workbook, {useSharedStrings: true});

  assert.equal(optionalPartText(pkg, 'xl/sharedStrings.xml'), undefined);
  // The workbook rels must not dangle a reference to a part that was omitted.
  assert.doesNotMatch(partText(pkg, 'xl/_rels/workbook.xml.rels'), /sharedStrings/);
});

test('an identical string is pooled once: count counts references, uniqueCount counts entries', () => {
  const pkg = writeXlsx(bookWithStrings('dup', 'dup', 'other'), {useSharedStrings: true});
  const sst = partText(pkg, 'xl/sharedStrings.xml');

  assert.match(sst, /count="3"/);
  assert.match(sst, /uniqueCount="2"/);
  // Both duplicate cells reference index 0; the distinct one gets index 1.
  const sheet = partText(pkg, 'xl/worksheets/sheet1.xml');
  assert.match(sheet, /r="A1"[^>]* t="s"><v>0<\/v>/);
  assert.match(sheet, /r="A2"[^>]* t="s"><v>0<\/v>/);
  assert.match(sheet, /r="A3"[^>]* t="s"><v>1<\/v>/);
});

test('shared and inline storage both read back to the same values', () => {
  const shared = roundtrip(bookWithStrings('a', 'b', 'a'), {useSharedStrings: true});
  const inline = roundtrip(bookWithStrings('a', 'b', 'a'));
  for (const wb of [shared, inline]) {
    const sheet = wb.getWorksheet('S');
    assert.ok(sheet);
    assert.equal(sheet.getCell('A1').value, 'a');
    assert.equal(sheet.getCell('A2').value, 'b');
    assert.equal(sheet.getCell('A3').value, 'a');
  }
});

test('under the option a rich-text cell is pooled as a rich <si> of runs and reads back formatted', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').getCell('A1').value = {
    richText: [{text: 'bold', font: {bold: true}}, {text: 'plain'}],
  };
  const pkg = writeXlsx(workbook, {useSharedStrings: true});

  // The runs become a rich <si> entry the cell references by index: the shape Excel itself writes.
  assert.match(
    partText(pkg, 'xl/sharedStrings.xml'),
    /<si><r><rPr><b\/><\/rPr><t>bold<\/t><\/r><r><t>plain<\/t><\/r><\/si>/,
  );
  const sheet = partText(pkg, 'xl/worksheets/sheet1.xml');
  assert.match(sheet, /r="A1"[^>]* t="s"><v>0<\/v>/);
  assert.doesNotMatch(sheet, /inlineStr/);

  const value = readXlsx(pkg).getWorksheet('S')?.getCell('A1').value;
  assert.ok(value && typeof value === 'object' && 'richText' in value);
  assert.deepEqual(value.richText, [{text: 'bold', font: {bold: true}}, {text: 'plain'}]);
});

test('a plain string and rich runs of the same text stay distinct entries in the pool', () => {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('S');
  sheet.getCell('A1').value = 'text';
  sheet.getCell('A2').value = {richText: [{text: 'text', font: {italic: true}}]};
  const pkg = writeXlsx(workbook, {useSharedStrings: true});
  const sst = partText(pkg, 'xl/sharedStrings.xml');

  // The plain <t> entry and the <r>-run entry render to different markup, so neither collapses
  // into the other: two references, two distinct entries.
  assert.match(sst, /uniqueCount="2"/);
  const cells = partText(pkg, 'xl/worksheets/sheet1.xml');
  assert.match(cells, /r="A1"[^>]* t="s"><v>0<\/v>/);
  assert.match(cells, /r="A2"[^>]* t="s"><v>1<\/v>/);
});

test('identical rich runs are pooled once, like plain strings', () => {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('S');
  const runs = {richText: [{text: 'x', font: {bold: true}}]};
  sheet.getCell('A1').value = runs;
  sheet.getCell('A2').value = {richText: [{text: 'x', font: {bold: true}}]};
  const sst = partText(writeXlsx(workbook, {useSharedStrings: true}), 'xl/sharedStrings.xml');

  assert.match(sst, /count="2"/);
  assert.match(sst, /uniqueCount="1"/);
});

test('parseSharedStrings reconstructs a foreign rich <si> into runs, not flattened text', () => {
  const sst =
    '<?xml version="1.0"?>' +
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">' +
    '<si><t>plain</t></si>' +
    '<si><r><rPr><b/><color rgb="FFFF0000"/></rPr><t>red</t></r><r><t> tail</t></r></si>' +
    '</sst>';
  const entries = parseSharedStrings(sst);

  assert.equal(entries[0], 'plain');
  assert.deepEqual(entries[1], {
    richText: [{text: 'red', font: {bold: true, color: {argb: 'FFFF0000'}}}, {text: ' tail'}],
  });
});

// What `parseSharedStrings` uniquely owns once the run grammar moved into the shared machine: what
// committing one `<si>` means. One `<r>` anywhere makes the entry rich, and a bare `<t>` beside the
// runs is not silently promoted into one.
test('an <si> holding both a bare <t> and runs commits as rich text', () => {
  const [entry] = parseSharedStrings('<sst><si><t>bare</t><r><t>run</t></r></si></sst>');
  assert.deepEqual(entry, {richText: [{text: 'run'}]});
});

test('a run boundary is not a text boundary: two runs stay two entries of one string', () => {
  const [entry] = parseSharedStrings('<sst><si><r><t>a</t><t>b</t></r><r><t>c</t></r></si></sst>');
  assert.deepEqual(entry, {richText: [{text: 'ab'}, {text: 'c'}]});
});

test('each <si> starts clean, so a rich entry does not leak into the plain one after it', () => {
  const entries = parseSharedStrings(
    '<sst><si><r><rPr><b/></rPr><t>rich</t></r></si><si><t>plain</t></si></sst>',
  );
  assert.deepEqual(entries, [{richText: [{text: 'rich', font: {bold: true}}]}, 'plain']);
});

test('a t="s" cell pointing at a foreign rich <si> reads back as rich text', () => {
  // Author a package whose A1 is a t="s" cell naming entry 0, then swap in a rich pool: the markup
  // Excel writes but our writer only produces on round-trip.
  const patched = patchParts(writeXlsx(bookWithStrings('plain'), {useSharedStrings: true}), {
    'xl/sharedStrings.xml': () =>
      '<?xml version="1.0"?>' +
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">' +
      '<si><r><rPr><i/></rPr><t>emph</t></r><r><t>rest</t></r></si></sst>',
  });
  assert.match(sheetXml(patched), /<c r="A1" t="s"><v>0<\/v><\/c>/, 'A1 names entry 0');

  const value = readXlsx(patched).getWorksheet('S')?.getCell('A1').value;
  assert.ok(value && typeof value === 'object' && 'richText' in value);
  assert.deepEqual(value.richText, [{text: 'emph', font: {italic: true}}, {text: 'rest'}]);
});
