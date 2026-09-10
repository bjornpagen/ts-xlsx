import assert from 'node:assert/strict';
import {test} from 'node:test';

import type {CfValueObjectType} from '../../core/conditional-formatting.ts';
import {Workbook} from '../../core/workbook.ts';
import {AuthoringError} from '../../errors.ts';
import {
  elementIn,
  partText,
  readPatched,
  roundtrip,
  SHEET1,
  sheetXml,
} from './package.test-support.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

const stylesXml = (pkg: Uint8Array): string => partText(pkg, 'xl/styles.xml');

test('a dataBar rule emits a well-formed block with two cfvo anchors and a bar colour', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addConditionalFormatting({
    ref: 'A1:A3',
    rules: [
      {
        type: 'dataBar',
        color: {argb: 'FF638EC6'},
        cfvo: [
          {type: 'num', value: 0},
          {type: 'num', value: 1},
        ],
      },
    ],
  });
  const xml = sheetXml(writeXlsx(workbook));
  const block = elementIn(xml, /<conditionalFormatting[\s\S]*?<\/conditionalFormatting>/);

  assert.equal([...xml.matchAll(/<conditionalFormatting\b/g)].length, 1, 'one block');
  assert.match(block, /<dataBar>/);
  assert.equal([...block.matchAll(/<cfvo\b/g)].length, 2, 'both anchors present');
  assert.match(block, /<color rgb="FF638EC6"\/>/);
});

test('a dataBar rule reads back on the same range with its type, colour, and both anchors', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addConditionalFormatting({
    ref: 'A1:A3',
    rules: [
      {
        type: 'dataBar',
        color: {argb: 'FF638EC6'},
        cfvo: [
          {type: 'num', value: 0},
          {type: 'num', value: 1},
        ],
      },
    ],
  });
  const rule = roundtrip(workbook).getWorksheet('S')?.conditionalFormattings[0]?.rules[0];
  assert.equal(rule?.type, 'dataBar');
  assert.equal(rule?.color?.argb, 'FF638EC6');
  assert.deepEqual(
    rule?.cfvo?.map((v) => v.value),
    [0, 1],
    'both numeric anchors survive in order',
  );
});

test('a minimal dataBar (no cfvo, no colour) gains Excel default min/max anchors and a bar colour', () => {
  const workbook = new Workbook();
  workbook
    .addWorksheet('S')
    .addConditionalFormatting({ref: 'A1:A3', rules: [{type: 'dataBar', priority: 1}]});
  const block = elementIn(sheetXml(writeXlsx(workbook)), /<dataBar[\s\S]*?<\/dataBar>/);

  assert.equal(
    [...block.matchAll(/<cfvo\b/g)].length,
    2,
    'a default data bar carries a min and a max cfvo',
  );
  assert.match(block, /<cfvo type="min"\/>/);
  assert.match(block, /<cfvo type="max"\/>/);
  assert.match(block, /<color\b/, 'a default bar colour is emitted');
});

test('a rule over a multi-area ref emits one block whose sqref lists every area', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addConditionalFormatting({
    ref: 'A1:C1 A3:C3 A5:C5',
    rules: [
      {
        type: 'colorScale',
        cfvo: [{type: 'min'}, {type: 'max'}],
        colors: [{argb: 'FFFF0000'}, {argb: 'FF00FF00'}],
      },
    ],
  });
  const xml = sheetXml(writeXlsx(workbook));

  assert.equal(
    [...xml.matchAll(/<conditionalFormatting\b/g)].length,
    1,
    'one block, not one per area',
  );
  assert.equal(xml.match(/<conditionalFormatting\b[^>]*sqref="([^"]*)"/)?.[1], 'A1:C1 A3:C3 A5:C5');
});

test('a stopIfTrue rule serialises the flag, references a dxf for its style, and both survive a read', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addConditionalFormatting({
    ref: 'A1:A10',
    rules: [
      {
        type: 'cellIs',
        operator: 'greaterThan',
        formulae: [3],
        stopIfTrue: true,
        style: {fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FFFF0000'}}},
      },
    ],
  });
  const pkg = writeXlsx(workbook);
  const xml = sheetXml(pkg);

  assert.match(
    xml,
    /<cfRule type="cellIs" dxfId="0" priority="1" stopIfTrue="1" operator="greaterThan">/,
  );
  assert.match(
    stylesXml(pkg),
    /<dxfs count="1"><dxf><fill><patternFill patternType="solid"><bgColor rgb="FFFF0000"\/>/,
  );

  const rule = readXlsx(pkg).getWorksheet('S')?.conditionalFormattings[0]?.rules[0];
  assert.equal(rule?.stopIfTrue, true, 'the flag reads back');
  assert.equal(rule?.dxfId, '0', 'the differential-style reference reads back');
});

test('an expression rule with no formula serialises without throwing', () => {
  const workbook = new Workbook();
  workbook
    .addWorksheet('S')
    .addConditionalFormatting({ref: 'A1:A10', rules: [{type: 'expression', style: {}}]});
  const xml = sheetXml(writeXlsx(workbook));
  assert.match(
    xml,
    /<cfRule type="expression"[^>]*\/>/,
    'a formula-less rule is a self-closing cfRule',
  );
});

test('a self-closing rule read from a file (a duplicateValues) round-trips its type, dxfId, and priority', () => {
  const part =
    '<?xml version="1.0"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>' +
    '<conditionalFormatting sqref="A1:A1048576">' +
    '<cfRule type="duplicateValues" dxfId="0" priority="1"/></conditionalFormatting></worksheet>';
  const xml = sheetXml(writeXlsx(readPatched({[SHEET1]: part})));

  assert.match(
    xml,
    /<cfRule type="duplicateValues" dxfId="0" priority="1"\/>/,
    'the rule survives, not dropped',
  );
  assert.equal([...xml.matchAll(/<conditionalFormatting\b/g)].length, 1, 'no empty block shell');
});

test('a foreign differential style with a custom number format round-trips verbatim, never "[object Object]"', () => {
  const code = '_(* #,##0_);_(* \\(#,##0\\);_(* &quot;-&quot;_);_(@_)';
  const styles =
    '<?xml version="1.0"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
    '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
    '<borders count="1"><border/></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
    `<dxfs count="1"><dxf><numFmt numFmtId="164" formatCode="${code}"/></dxf></dxfs>` +
    '</styleSheet>';
  const sheet1 =
    '<?xml version="1.0"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>' +
    '<conditionalFormatting sqref="A1:A5">' +
    '<cfRule type="expression" dxfId="0" priority="1"><formula>A1&gt;2</formula></cfRule>' +
    '</conditionalFormatting></worksheet>';
  const out = stylesXml(writeXlsx(readPatched({[SHEET1]: sheet1, 'xl/styles.xml': styles})));

  assert.doesNotMatch(
    out,
    /\[object Object\]/,
    'the dxf numFmt is a real format code, not a coerced object',
  );
  assert.match(out, /formatCode="_\(\* #,##0_\)/, 'the exact custom format code is preserved');
});

test('an x14 extLst conditional formatting is left untouched and writing the sheet does not crash', () => {
  const sheet1 =
    '<?xml version="1.0"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>' +
    '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" ' +
    'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
    '<x14:conditionalFormattings><x14:conditionalFormatting ' +
    'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' +
    '<x14:cfRule type="expression" priority="1" id="{GUID}"><xm:f>A1&gt;2</xm:f></x14:cfRule>' +
    '<xm:sqref>A1:A5</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst></worksheet>';
  const workbook = readPatched({[SHEET1]: sheet1});

  // The x14 rule is namespace-prefixed and is not read into the classic model, so it is neither
  // half-parsed into a broken rule nor does it make the writer throw.
  assert.equal(
    workbook.getWorksheet('S')?.conditionalFormattings.length,
    0,
    'no classic rule was fabricated',
  );
  assert.doesNotThrow(
    () => writeXlsx(workbook),
    'writing a sheet whose CF lived only in x14 does not crash',
  );
});

test('a colorScale colour with a malformed theme attribute drops it rather than round-tripping NaN', () => {
  const sheet1 =
    '<?xml version="1.0"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>' +
    '<conditionalFormatting sqref="A1:A3"><cfRule type="colorScale" priority="1"><colorScale>' +
    '<cfvo type="min"/><cfvo type="max"/>' +
    '<color theme="oops"/><color rgb="FF00FF00"/>' +
    '</colorScale></cfRule></conditionalFormatting></worksheet>';
  const out = sheetXml(writeXlsx(readPatched({[SHEET1]: sheet1})));

  // The shared validated parseColor drops the non-integer theme instead of coercing it to NaN, so the
  // writer never emits `theme="NaN"`, which Excel would reject.
  assert.doesNotMatch(out, /theme="NaN"/, 'the malformed theme is not written as NaN');
  assert.match(out, /<color rgb="FF00FF00"\/>/, 'the well-formed sibling colour survives');
});

test('a malformed priority is dropped on read rather than poisoning every later rule with NaN', () => {
  const sheet1 =
    '<?xml version="1.0"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>' +
    '<conditionalFormatting sqref="A1:A3">' +
    '<cfRule type="duplicateValues" priority="oops"/>' +
    '<cfRule type="duplicateValues" priority="2"/>' +
    '</conditionalFormatting></worksheet>';
  const xml = sheetXml(writeXlsx(readPatched({[SHEET1]: sheet1})));

  assert.doesNotMatch(xml, /priority="NaN"/, 'no rule re-writes with a NaN priority');
  assert.equal(
    [...xml.matchAll(/priority="(\d+)"/g)].length,
    2,
    'both rules keep a real numeric priority',
  );
});

test('a malformed dxfId is dropped on read rather than round-tripping as NaN', () => {
  const sheet1 =
    '<?xml version="1.0"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>' +
    '<conditionalFormatting sqref="A1:A3">' +
    '<cfRule type="duplicateValues" dxfId="oops" priority="1"/>' +
    '</conditionalFormatting></worksheet>';
  const xml = sheetXml(writeXlsx(readPatched({[SHEET1]: sheet1})));

  assert.doesNotMatch(xml, /dxfId="NaN"/, 'the malformed dxfId is not written as NaN');
  assert.doesNotMatch(
    xml,
    /dxfId="oops"/,
    'the malformed dxfId is not echoed back verbatim either',
  );
});

test('a malformed rank and stdDev are dropped on read rather than round-tripping as NaN', () => {
  const sheet1 =
    '<?xml version="1.0"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>' +
    '<conditionalFormatting sqref="A1:A3">' +
    '<cfRule type="top10" rank="oops" priority="1"/>' +
    '<cfRule type="aboveAverage" stdDev="oops" priority="2"/>' +
    '</conditionalFormatting></worksheet>';
  const xml = sheetXml(writeXlsx(readPatched({[SHEET1]: sheet1})));

  assert.doesNotMatch(xml, /rank="NaN"/, 'the malformed rank is not written as NaN');
  assert.doesNotMatch(xml, /stdDev="NaN"/, 'the malformed stdDev is not written as NaN');
});

// `priority` and `stdDev` are xsd:int and `rank` is xsd:unsignedInt, but they were read as any finite
// number, so `priority="1.5"`, `rank="2.5"` and `stdDev="-1.5"` came back out verbatim.
test('a fractional priority, rank or stdDev, or a negative rank, is dropped on read', () => {
  const sheet1 =
    '<?xml version="1.0"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>' +
    '<conditionalFormatting sqref="A1:A3">' +
    '<cfRule type="duplicateValues" priority="1.5"/>' +
    '<cfRule type="top10" rank="2.5" priority="2"/>' +
    '<cfRule type="top10" rank="-3" priority="3"/>' +
    '<cfRule type="aboveAverage" stdDev="-1.5" priority="4"/>' +
    '<cfRule type="aboveAverage" stdDev="-2" priority="5"/>' +
    '</conditionalFormatting></worksheet>';
  const book = readPatched({[SHEET1]: sheet1});
  const rules = book.getWorksheet('S')?.conditionalFormattings[0]?.rules ?? [];
  assert.deepEqual(
    rules.map((rule) => [rule.priority, rule.rank, rule.stdDev]),
    [
      [undefined, undefined, undefined],
      [2, undefined, undefined],
      [3, undefined, undefined],
      [4, undefined, undefined],
      [5, undefined, -2],
    ],
    'only integers the attributes can hold are kept, a negative stdDev included',
  );
  const block = elementIn(
    sheetXml(writeXlsx(book)),
    /<conditionalFormatting[\s\S]*?<\/conditionalFormatting>/,
  );
  assert.doesNotMatch(block, /="-?\d+\.\d+"|rank="-/, 'and none is written back');
});

test('an authored priority, rank or stdDev the schema cannot hold is refused at write', () => {
  const refused = (rule: Record<string, unknown>) => {
    const workbook = new Workbook();
    workbook.addWorksheet('S').addConditionalFormatting({
      ref: 'A1:A3',
      rules: [{type: 'top10', ...rule}],
    });
    assert.throws(() => writeXlsx(workbook), AuthoringError, JSON.stringify(rule));
  };
  refused({priority: Number.NaN});
  refused({priority: 1.5});
  refused({rank: -1});
  refused({rank: 2.5});
  refused({stdDev: 0.5});
});

// An authored NaN priority used to spread through `Math.max` into the counter, so every rule after it
// that took an automatic priority was written `priority="NaN"` too.
test('a NaN priority is refused before it can reach the automatic priorities after it', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addConditionalFormatting({
    ref: 'A1:A3',
    rules: [{type: 'duplicateValues', priority: Number.NaN}, {type: 'uniqueValues'}],
  });
  assert.throws(() => writeXlsx(workbook), AuthoringError);
});

test('a NaN cfvo value is refused rather than written as val="NaN"', () => {
  for (const type of ['colorScale', 'dataBar'] as const) {
    const workbook = new Workbook();
    workbook.addWorksheet('S').addConditionalFormatting({
      ref: 'A1:A3',
      rules: [{type, priority: 1, cfvo: [{type: 'num', value: Number.NaN}, {type: 'max'}]}],
    });
    assert.throws(() => writeXlsx(workbook), AuthoringError, type);
  }
});

// A block with no `sqref` read as `ref: ''` and was written back as `sqref=""`, and one whose `sqref`
// did not decode was written back verbatim. Neither names a cell to format.
test('a conditional format whose sqref is missing or does not decode is dropped on read', () => {
  const sheet1 =
    '<?xml version="1.0"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>' +
    '<conditionalFormatting><cfRule type="duplicateValues" priority="1"/></conditionalFormatting>' +
    '<conditionalFormatting sqref="not a range"><cfRule type="duplicateValues" priority="2"/>' +
    '</conditionalFormatting>' +
    '<conditionalFormatting sqref="B1:B3"><cfRule type="duplicateValues" priority="3"/>' +
    '</conditionalFormatting></worksheet>';
  const book = readPatched({[SHEET1]: sheet1});
  assert.deepEqual(
    book.getWorksheet('S')?.conditionalFormattings.map((cf) => cf.ref),
    ['B1:B3'],
  );
  assert.doesNotMatch(sheetXml(writeXlsx(book)), /sqref=""|sqref="not a range"/);
});

test('authoring a conditional format on a range that names no cells is refused', () => {
  const sheet = new Workbook().addWorksheet('S');
  for (const ref of ['', 'not a range']) {
    assert.throws(
      () => sheet.addConditionalFormatting({ref, rules: [{type: 'duplicateValues'}]}),
      AuthoringError,
      JSON.stringify(ref),
    );
  }
  assert.deepEqual(sheet.conditionalFormattings, []);
});

function dataBarBook(rule: Record<string, unknown>): Workbook {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addConditionalFormatting({
    ref: 'A1:A3',
    rules: [
      {
        type: 'dataBar',
        color: {argb: 'FF638EC6'},
        cfvo: [
          {type: 'num', value: 0},
          {type: 'num', value: 1},
        ],
        ...rule,
      },
    ],
  });
  return workbook;
}

test('a gradient dataBar writes the classic element plus a linked x14 extension carrying the flag', () => {
  const xml = sheetXml(writeXlsx(dataBarBook({gradient: true})));

  // The classic element is unchanged: every consumer still understands the bar's anchors and colour.
  assert.match(
    xml,
    /<dataBar><cfvo type="num" val="0"\/><cfvo type="num" val="1"\/><color rgb="FF638EC6"\/><\/dataBar>/,
  );
  // The cfRule links its extension by a shared id, which the worksheet x14 block echoes on its cfRule.
  const id = xml.match(/<x14:id>(\{[^}]+\})<\/x14:id>/)?.[1];
  assert.ok(id, 'the classic cfRule carries an x14:id link');
  assert.match(
    xml,
    new RegExp(`<x14:cfRule type="dataBar" id="${id.replace(/[{}]/g, '\\$&')}">`),
    'the extension echoes the same id',
  );
  assert.match(
    xml,
    /<x14:dataBar gradient="1"><x14:cfvo type="num"><xm:f>0<\/xm:f><\/x14:cfvo><x14:cfvo type="num"><xm:f>1<\/xm:f><\/x14:cfvo><\/x14:dataBar>/,
  );
  assert.match(xml, /<xm:sqref>A1:A3<\/xm:sqref>/);
});

test('the gradient flag survives a write→read round-trip through the x14 extension', () => {
  for (const gradient of [true, false]) {
    const rule = roundtrip(dataBarBook({gradient})).getWorksheet('S')?.conditionalFormattings[0]
      ?.rules[0];
    assert.equal(rule?.type, 'dataBar');
    assert.equal(rule?.gradient, gradient, `gradient=${gradient} reads back`);
    // The classic facets still survive alongside the enriched ones.
    assert.equal(rule?.color?.argb, 'FF638EC6');
    assert.deepEqual(
      rule?.cfvo?.map((v) => v.value),
      [0, 1],
    );
  }
});

test('a dataBar negative-fill and axis colour round-trip through the x14 extension', () => {
  const book = dataBarBook({negativeFillColor: {argb: 'FFFF0000'}, axisColor: {argb: 'FF000000'}});
  const xml = sheetXml(writeXlsx(book));
  assert.match(xml, /<x14:negativeFillColor rgb="FFFF0000"\/>/);
  assert.match(xml, /<x14:axisColor rgb="FF000000"\/>/);

  const rule = roundtrip(book).getWorksheet('S')?.conditionalFormattings[0]?.rules[0];
  assert.equal(rule?.negativeFillColor?.argb, 'FFFF0000');
  assert.equal(rule?.axisColor?.argb, 'FF000000');
});

test('a plain dataBar carrying no x14 facet stays classic-only, fabricating no extension', () => {
  const xml = sheetXml(writeXlsx(dataBarBook({})));
  assert.match(xml, /<dataBar>/, 'the classic element is present');
  assert.doesNotMatch(xml, /x14:dataBar/, 'no x14 data-bar extension is written');
  assert.doesNotMatch(xml, /<extLst>/, 'no extLst is fabricated for a plain data bar');
});

test('a sheet with both an extended validation and a gradient dataBar emits one worksheet extLst', () => {
  const workbook = dataBarBook({gradient: true});
  workbook
    .getWorksheet('S')
    ?.addDataValidation('B1', {type: 'list', formulae: ['Other!$A$1:$A$3']}, {extended: true});
  const xml = sheetXml(writeXlsx(workbook));

  // Both extensions ride as sibling <ext> blocks inside a single worksheet <extLst>, never two.
  assert.match(
    xml,
    /uri="\{78C0D931-6437-407d-A8EE-F0AAD7539E65\}"/,
    'the CF extension is present',
  );
  assert.match(
    xml,
    /uri="\{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF\}"/,
    'the DV extension is present',
  );
  assert.doesNotMatch(
    xml,
    /<\/extLst><extLst>/,
    'the two extensions are not split across two worksheet extLst',
  );

  const sheet = roundtrip(workbook).getWorksheet('S');
  assert.equal(sheet?.conditionalFormattings[0]?.rules[0]?.gradient, true, 'the gradient survives');
  assert.equal(sheet?.dataValidations.length, 1, 'the extended validation survives');
});

// The third scale type. `dataBar` and `colorScale` were exercised above from the moment they
// landed; `iconSet` reached the model, the reader and the writer without a single test naming it,
// so `iconSetXml` was serialising a documented rule type that nothing had ever read back.

test('an iconSet rule emits its icon family and one cfvo per threshold', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addConditionalFormatting({
    ref: 'A1:A9',
    rules: [
      {
        type: 'iconSet',
        iconSet: '3TrafficLights1',
        cfvo: [
          {type: 'percent', value: 0},
          {type: 'percent', value: 33},
          {type: 'percent', value: 67},
        ],
      },
    ],
  });
  const block = elementIn(sheetXml(writeXlsx(workbook)), /<iconSet[\s\S]*?<\/iconSet>/);

  assert.match(block, /<iconSet iconSet="3TrafficLights1">/);
  assert.equal([...block.matchAll(/<cfvo\b/g)].length, 3, 'all three thresholds present');
  assert.match(block, /<cfvo type="percent" val="33"\/>/);
});

test('an iconSet rule reads back with its type, icon family, and every threshold in order', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addConditionalFormatting({
    ref: 'A1:A9',
    rules: [
      {
        type: 'iconSet',
        iconSet: '5Rating',
        cfvo: [
          {type: 'min'},
          {type: 'percentile', value: 25},
          {type: 'percentile', value: 50},
          {type: 'percentile', value: 75},
          {type: 'max'},
        ],
      },
    ],
  });
  const rule = roundtrip(workbook).getWorksheet('S')?.conditionalFormattings[0]?.rules[0];
  assert.equal(rule?.type, 'iconSet');
  assert.equal(rule?.iconSet, '5Rating');
  assert.deepEqual(
    rule?.cfvo?.map((v) => v.type),
    ['min', 'percentile', 'percentile', 'percentile', 'max'],
    'every anchor survives in order, min/max carrying no value',
  );
});

test('an iconSet with no named family emits the element without an empty iconSet attribute', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addConditionalFormatting({
    ref: 'A1:A3',
    rules: [{type: 'iconSet', priority: 1, cfvo: [{type: 'min'}, {type: 'max'}]}],
  });
  const block = elementIn(sheetXml(writeXlsx(workbook)), /<iconSet[\s\S]*?<\/iconSet>/);

  assert.match(block, /^<iconSet>/, 'an absent family names no attribute at all');
  assert.doesNotMatch(block, /iconSet=""/);
});

// `gradient` and `aboveAverage` are both default-true xsd:booleans, so a producer that spells false
// the long way must turn them off exactly as the digit does, and an unrecognised token, on a
// default-true attribute, still reads as true rather than as absent.

test('x14 gradient="false" turns the bar flat exactly as gradient="0" does', () => {
  const written = sheetXml(writeXlsx(dataBarBook({gradient: true})));
  for (const spelling of ['0', 'false']) {
    const sheet1 = written.replace('gradient="1"', `gradient="${spelling}"`);
    const rule = readPatched({[SHEET1]: sheet1}).getWorksheet('S')?.conditionalFormattings[0]
      ?.rules[0];
    assert.equal(rule?.gradient, false, `gradient="${spelling}" reads as flat`);
  }
});

test('aboveAverage="false" reads as below-average; garbage stays above, absence stays absent', () => {
  const read = (attr: string): boolean | undefined => {
    const sheet1 =
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/>' +
      `<conditionalFormatting sqref="A1:A3"><cfRule type="aboveAverage" priority="1"${attr}/>` +
      '</conditionalFormatting></worksheet>';
    return readPatched({[SHEET1]: sheet1}).getWorksheet('S')?.conditionalFormattings[0]?.rules[0]
      ?.aboveAverage;
  };
  assert.equal(read(' aboveAverage="0"'), false);
  assert.equal(read(' aboveAverage="false"'), false, 'the long spelling turns it off too');
  assert.equal(read(' aboveAverage="1"'), true);
  assert.equal(read(' aboveAverage="yes"'), true, 'a default-true attribute reads garbage as on');
  assert.equal(read(''), undefined, 'an absent attribute stays absent, not present-and-true');
});

test('a cfvo whose type is not a schema token reads back as num, the schema default', () => {
  const part =
    '<?xml version="1.0"?>' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<sheetData/>' +
    '<conditionalFormatting sqref="A1:A3"><cfRule type="colorScale" priority="1"><colorScale>' +
    '<cfvo type="nonsense" val="0"/><cfvo type="max"/>' +
    '<color rgb="FFFFFFFF"/><color rgb="FF638EC6"/>' +
    '</colorScale></cfRule></conditionalFormatting></worksheet>';
  const rule = readPatched({[SHEET1]: part}).getWorksheet('S')?.conditionalFormattings[0]?.rules[0];

  assert.deepEqual(
    rule?.cfvo?.map((v) => v.type),
    ['num', 'max'],
    'the unrecognised anchor falls back rather than leaving the scale a member short',
  );
  assert.equal(rule?.cfvo?.[0]?.value, 0, 'and still carries its value');
});

test('every schema token of the cfvo type union round-trips unchanged', () => {
  const types: CfValueObjectType[] = ['num', 'percent', 'max', 'min', 'percentile', 'formula'];
  const workbook = new Workbook();
  workbook.addWorksheet('S').addConditionalFormatting({
    ref: 'A1:A9',
    rules: types.map((type, index) => ({
      type: 'dataBar',
      priority: index + 1,
      color: {argb: 'FF638EC6'},
      cfvo: [{type, value: type === 'formula' ? '$A$1' : 1}, {type: 'max'}],
    })),
  });

  const rules = roundtrip(workbook).getWorksheet('S')?.conditionalFormattings[0]?.rules;
  assert.deepEqual(
    rules?.map((rule) => rule.cfvo?.[0]?.type),
    types,
    'no legal anchor type is narrowed away',
  );
});

test('an authored dxfId that is not an index is refused rather than written as dxfId="NaN"', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addConditionalFormatting({
    ref: 'A1:A3',
    rules: [
      {type: 'cellIs', operator: 'greaterThan', priority: 1, formulae: [1], dxfId: 'x'},
      {type: 'cellIs', operator: 'greaterThan', priority: 2, formulae: [2], dxfId: '-1'},
      {type: 'cellIs', operator: 'greaterThan', priority: 3, formulae: [3], dxfId: '0'},
    ],
  });
  const xml = sheetXml(writeXlsx(workbook));

  assert.doesNotMatch(
    xml,
    /dxfId="(NaN|x|-1)"/,
    'the same predicate the reader uses applies on the way out',
  );
  assert.equal([...xml.matchAll(/dxfId="/g)].length, 1, 'only the real index survives');
  assert.match(xml, /dxfId="0"/);
});

test('a rule whose dxfId is unusable falls through to the style authored beside it', () => {
  const workbook = new Workbook();
  workbook.addWorksheet('S').addConditionalFormatting({
    ref: 'A1:A3',
    rules: [
      {
        type: 'cellIs',
        operator: 'greaterThan',
        priority: 1,
        formulae: [1],
        dxfId: 'x',
        style: {font: {bold: true}},
      },
    ],
  });
  const pkg = writeXlsx(workbook);

  assert.match(sheetXml(pkg), /dxfId="0"/, 'the interned style takes the slot');
  assert.match(stylesXml(pkg), /<dxfs count="1">/);
});
