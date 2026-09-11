import assert from 'node:assert/strict';
import {test} from 'node:test';

import {parseStyleTable} from './read-styles.ts';

// The style reader narrows every enumerated attribute through a guard: a valid token passes verbatim,
// an unrecognised one is dropped rather than trusted into the model as a bogus union member.

// `fillId` is an index, and a slot was pushed per fill *body* rather than per `<fill>`, so a malformed
// fill table shifted every later cell onto the fill before its own.
const RED = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF0000'}};
const GREEN = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF00FF00'}};
const solidBody = (argb: string) =>
  `<patternFill patternType="solid"><fgColor rgb="${argb}"/></patternFill>`;
const fillTable = (fills: string, xfs: string) =>
  `<styleSheet><fills>${fills}</fills><cellXfs>${xfs}</cellXfs></styleSheet>`;

test('a <fill> holding two bodies takes one slot, keeping its first body', () => {
  const table = parseStyleTable(
    fillTable(
      `<fill>${solidBody('FFFF0000')}${solidBody('FF0000FF')}</fill><fill>${solidBody('FF00FF00')}</fill>`,
      '<xf fillId="0"/><xf fillId="1"/>',
    ),
  );
  assert.deepEqual(table.cellXfs[0]?.fill, RED, 'the first body wins');
  assert.deepEqual(table.cellXfs[1]?.fill, GREEN, 'and fillId 1 is still the second <fill>');
});

test('a body directly under <fills> belongs to no fill and shifts no index', () => {
  const table = parseStyleTable(
    fillTable(`${solidBody('FFFF0000')}<fill>${solidBody('FF00FF00')}</fill>`, '<xf fillId="0"/>'),
  );
  assert.deepEqual(table.cellXfs[0]?.fill, GREEN);
});

test('a self-closing <fill/> still occupies its slot', () => {
  const table = parseStyleTable(
    fillTable(`<fill/><fill>${solidBody('FF00FF00')}</fill>`, '<xf fillId="0"/><xf fillId="1"/>'),
  );
  assert.equal(table.cellXfs[0]?.fill, undefined);
  assert.deepEqual(table.cellXfs[1]?.fill, GREEN);
});

test('a valid border-edge style is kept; an unrecognised one is dropped', () => {
  const good = parseStyleTable(
    '<styleSheet>' +
      '<borders count="2"><border/><border><left style="thin"/></border></borders>' +
      '<cellXfs count="1"><xf borderId="1"/></cellXfs>' +
      '</styleSheet>',
  );
  assert.deepEqual(good.cellXfs[0]?.border, {left: {style: 'thin'}});

  const bad = parseStyleTable(
    '<styleSheet>' +
      '<borders count="2"><border/><border><left style="frobnicate"/></border></borders>' +
      '<cellXfs count="1"><xf borderId="1"/></cellXfs>' +
      '</styleSheet>',
  );
  assert.equal(bad.cellXfs[0]?.border, undefined, 'the invalid edge is dropped, leaving no border');
});

test('a named underline is kept, "none" reads false, and an unknown token stays a plain underline', () => {
  const read = (val: string): unknown => {
    const table = parseStyleTable(
      `<styleSheet><fonts count="1"><font><u val="${val}"/></font></fonts>` +
        '<cellXfs count="1"><xf fontId="0"/></cellXfs></styleSheet>',
    );
    return table.cellXfs[0]?.font?.underline;
  };
  assert.equal(read('double'), 'double');
  assert.equal(read('none'), false);
  assert.equal(
    read('squiggly'),
    true,
    'an unrecognised style keeps the underline, drops the token',
  );
});

test('an unrecognised vertAlign or scheme token is dropped from the font', () => {
  const table = parseStyleTable(
    '<styleSheet><fonts count="1">' +
      '<font><b/><vertAlign val="diagonal"/><scheme val="fancy"/></font>' +
      '</fonts><cellXfs count="1"><xf fontId="0"/></cellXfs></styleSheet>',
  );
  assert.deepEqual(table.cellXfs[0]?.font, {bold: true}, 'only the valid facet survives');
});

test('a valid vertAlign and scheme pass through verbatim', () => {
  const table = parseStyleTable(
    '<styleSheet><fonts count="1">' +
      '<font><vertAlign val="superscript"/><scheme val="minor"/></font>' +
      '</fonts><cellXfs count="1"><xf fontId="0"/></cellXfs></styleSheet>',
  );
  assert.equal(table.cellXfs[0]?.font?.vertAlign, 'superscript');
  assert.equal(table.cellXfs[0]?.font?.scheme, 'minor');
});

test('font 0 is surfaced as the declared default font, whatever face it names', () => {
  const table = parseStyleTable(
    '<styleSheet><fonts count="2">' +
      '<font><sz val="11"/><color theme="1"/><name val="Aptos Narrow"/><family val="2"/><scheme val="minor"/></font>' +
      '<font><b/><sz val="18"/></font>' +
      '</fonts><cellXfs count="1"><xf fontId="1"/></cellXfs></styleSheet>',
  );
  assert.deepEqual(table.defaultFont, {
    size: 11,
    color: {theme: 1},
    name: 'Aptos Narrow',
    family: 2,
    scheme: 'minor',
  });
});

test('a file declaring no font table declares no default font', () => {
  // Distinct from declaring Calibri: the library must not fabricate a declaration a file never made,
  // because doing so would let a re-write state a default the source never had.
  const table = parseStyleTable('<styleSheet><cellXfs count="1"><xf/></cellXfs></styleSheet>');
  assert.equal(table.defaultFont, undefined);
  assert.equal(parseStyleTable('').defaultFont, undefined);
});

test('an unrecognised alignment token is dropped; a valid one is kept', () => {
  const bad = parseStyleTable(
    '<styleSheet><cellXfs count="1">' +
      '<xf><alignment horizontal="sideways" vertical="floating"/></xf>' +
      '</cellXfs></styleSheet>',
  );
  assert.equal(
    bad.cellXfs[0]?.alignment,
    undefined,
    'both bogus tokens drop, leaving no alignment',
  );

  const good = parseStyleTable(
    '<styleSheet><cellXfs count="1">' +
      '<xf><alignment horizontal="center" vertical="top"/></xf>' +
      '</cellXfs></styleSheet>',
  );
  assert.deepEqual(good.cellXfs[0]?.alignment, {horizontal: 'center', vertical: 'top'});
});

// A named style's label, the `name`/`builtinId` pair on <cellStyle>, is invisible to a round trip:
// a reader that drops it and a writer that never emits it agree with each other, and the corpus is a
// fixed point of both (ADR 0012). So it is asserted here, against the parse, in the only place that
// can see the difference.

test("a named style's label reaches the model, name and builtinId both", () => {
  const table = parseStyleTable(
    '<styleSheet>' +
      '<cellStyleXfs count="2"><xf/><xf/></cellStyleXfs>' +
      '<cellXfs count="1"><xf xfId="1"/></cellXfs>' +
      '<cellStyles count="2">' +
      '<cellStyle name="Normal" xfId="0" builtinId="0"/>' +
      '<cellStyle name="Heading 1" xfId="1" builtinId="16"/>' +
      '</cellStyles>' +
      '</styleSheet>',
  );
  assert.deepEqual(
    table.namedStyles.map((style) => [style.name, style.builtinId]),
    [
      ['Normal', 0],
      ['Heading 1', 16],
    ],
  );
});

test('a duplicated xfId is titled by the first label that claims it', () => {
  // A foreign file may name the same base twice. Neither label is more right than the other, so the
  // rule is simply stated and locked here: first declaration wins, and the second is dropped rather
  // than layered over it.
  const table = parseStyleTable(
    '<styleSheet>' +
      '<cellStyleXfs count="1"><xf/></cellStyleXfs>' +
      '<cellStyles count="2">' +
      '<cellStyle name="First" xfId="0" builtinId="0"/>' +
      '<cellStyle name="Second" xfId="0" builtinId="16"/>' +
      '</cellStyles>' +
      '</styleSheet>',
  );
  assert.deepEqual(
    table.namedStyles.map((style) => [style.name, style.builtinId]),
    [['First', 0]],
  );
});

test('a label out of order titles the entry its xfId names, not the one it was declared at', () => {
  // The labels zip against cellStyleXfs by xfId, not by document position, so a file listing them in
  // any other order must still title the right base, and an unlabelled base must stay unlabelled
  // rather than inherit its neighbour's name.
  const table = parseStyleTable(
    '<styleSheet>' +
      '<cellStyleXfs count="3"><xf/><xf/><xf/></cellStyleXfs>' +
      '<cellStyles count="2">' +
      '<cellStyle name="Title" xfId="2" builtinId="15"/>' +
      '<cellStyle name="Custom" xfId="0"/>' +
      '</cellStyles>' +
      '</styleSheet>',
  );
  assert.deepEqual(
    table.namedStyles.map((style) => [style.name, style.builtinId]),
    [
      ['Custom', undefined],
      [undefined, undefined],
      ['Title', 15],
    ],
  );
});
