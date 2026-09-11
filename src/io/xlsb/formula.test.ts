import assert from 'node:assert/strict';
import {test} from 'node:test';

import {decodeFormula, type FormulaScope, formulaAnchor} from './formula.ts';
import {builtinFunctionAt, functionNameFor} from './ptg-functions.ts';
import {nameCitations} from './records.test-support.ts';

function bytes(...values: number[]): Uint8Array {
  return Uint8Array.from(values);
}

const NONE = new Uint8Array(0);

// A workbook of three sheets whose extern-sheet table names each of them and, at index 3, an
// external workbook: the shape the resolver has to tell apart.
const SCOPE: FormulaScope = {
  sheetNames: ['Calc', 'Data', 'More'],
  externSheets: [
    {supBook: 0, firstSheet: 0, lastSheet: 0},
    {supBook: 0, firstSheet: 1, lastSheet: 2},
    {supBook: 0, firstSheet: 9, lastSheet: 9},
    {supBook: 1, firstSheet: 0, lastSheet: 0},
  ],
  selfSupBook: 0,
  names: ['Rate'],
};

// `PtgRef` (value class): the ptg, a 4-byte row, then the packed column word.
function ref(row: number, column: number, flags = 0xc000): Uint8Array {
  const out = new Uint8Array(7);
  const view = new DataView(out.buffer);
  view.setUint8(0, 0x44);
  view.setUint32(1, row, true);
  view.setUint16(5, column | flags, true);
  return out;
}

test('a postfix stream folds back into infix text', () => {
  // 1, 2, 3, multiply, add: the tokens behind `1+2*3`.
  const rgce = bytes(0x1e, 1, 0, 0x1e, 2, 0, 0x1e, 3, 0, 0x05, 0x03);
  assert.equal(decodeFormula(rgce, NONE, SCOPE), '1+2*3');
});

test('a stored parenthesis is emitted where the author typed it', () => {
  const rgce = bytes(0x1e, 1, 0, 0x1e, 2, 0, 0x03, 0x15, 0x1e, 3, 0, 0x05);
  assert.equal(decodeFormula(rgce, NONE, SCOPE), '(1+2)*3');
});

test('the relative-axis flags decide where the dollar signs go', () => {
  assert.equal(decodeFormula(ref(0, 0, 0x0000), NONE, SCOPE), '$A$1');
  assert.equal(decodeFormula(ref(1, 2, 0x4000), NONE, SCOPE), 'C$2');
  assert.equal(decodeFormula(ref(1, 2, 0x8000), NONE, SCOPE), '$C2');
  assert.equal(decodeFormula(ref(1, 2, 0xc000), NONE, SCOPE), 'C2');
});

test('a 3-D reference resolves its sheet span through the extern-sheet table', () => {
  // PtgRef3d (value class): ptg, ixti, row, packed column.
  const at = (ixti: number) => bytes(0x5a, ixti, 0, 0, 0, 0, 0, 0x00, 0xc0);
  assert.equal(decodeFormula(at(0), NONE, SCOPE), 'Calc!A1');
  assert.equal(decodeFormula(at(1), NONE, SCOPE), 'Data:More!A1');
  // A span naming a sheet the workbook does not have resolves to nothing rather than to a guess.
  assert.equal(decodeFormula(at(2), NONE, SCOPE), undefined);
  // A supporting book that is not this workbook: its sheets cannot be named from here.
  assert.equal(decodeFormula(at(3), NONE, SCOPE), undefined);
});

test('a reference through an unrecognised externals table does not resolve', () => {
  // With a supporting book this reader could not account for, the whole table's indices are
  // untrustworthy: naming a sheet from them could name the wrong one.
  const opaque: FormulaScope = {...SCOPE, selfSupBook: undefined};
  assert.equal(decodeFormula(bytes(0x5a, 0, 0, 0, 0, 0, 0, 0x00, 0xc0), NONE, opaque), undefined);
});

test('a name is cited by its position in the file, counting the ones the model drops', () => {
  assert.equal(decodeFormula(bytes(0x43, 1, 0, 0, 0), NONE, SCOPE), 'Rate');
  assert.equal(decodeFormula(bytes(0x43, 2, 0, 0, 0), NONE, SCOPE), undefined);
});

test('an unrecognised token abandons the formula rather than guessing past it', () => {
  // 0x18 opens the extended-token family (structured references), which this reader does not decode.
  // Its length is not knowable from the ptg alone, so continuing would desynchronise the walk.
  assert.equal(decodeFormula(bytes(0x1e, 1, 0, 0x18, 0x19, 0, 0), NONE, SCOPE), undefined);
});

test('a short stream citing a long name costs the formula, not the read', () => {
  // A five-byte `PtgName` stands for a name of any length, so decoded text is not bounded by the bytes
  // it came from. This stream used to throw a native `RangeError: Invalid string length`.
  const rgce = nameCitations(600);
  assert.equal(rgce.length, 3_599);
  assert.equal(decodeFormula(rgce, NONE, {...SCOPE, names: ['x'.repeat(1 << 20)]}), undefined);
});

test('a formula as long as Excel allows still decodes', () => {
  const name = 'x'.repeat(8_192);
  assert.equal(decodeFormula(nameCitations(1), NONE, {...SCOPE, names: [name]}), name);
});

test('a byte above 0x7f names no token, even where masking it would name an operand', () => {
  // A classed token is a five-bit base under two bits of result class, so the family is 0x20..0x7f.
  // Each payload here is one the value-class token of the same base decodes, which is asserted
  // alongside, so the refusal is of the byte and not of a malformed stream.
  const payloads: ReadonlyMap<number, readonly number[]> = new Map([
    [0x23, [1, 0, 0, 0]],
    [0x24, [0, 0, 0, 0, 0x00, 0xc0]],
    [0x25, [0, 0, 0, 0, 3, 0, 0, 0, 0x00, 0xc0, 0x01, 0xc0]],
    [0x2a, new Array<number>(6).fill(0)],
    [0x2b, new Array<number>(12).fill(0)],
    [0x3a, [0, 0, 0, 0, 0, 0, 0x00, 0xc0]],
    [0x3b, [0, 0, 0, 0, 0, 0, 3, 0, 0, 0, 0x00, 0xc0, 0x01, 0xc0]],
    [0x3c, new Array<number>(8).fill(0)],
    [0x3d, new Array<number>(14).fill(0)],
  ]);
  for (const [base, payload] of payloads) {
    const valueClass = base + 0x20;
    assert.notEqual(
      decodeFormula(bytes(valueClass, ...payload), NONE, SCOPE),
      undefined,
      `0x${valueClass.toString(16)} decodes`,
    );
  }
  for (let ptg = 0x80; ptg <= 0xff; ptg++) {
    const payload = payloads.get((ptg & 0x1f) | 0x20) ?? [];
    assert.equal(
      decodeFormula(bytes(ptg, ...payload), NONE, SCOPE),
      undefined,
      `0x${ptg.toString(16)}`,
    );
  }
});

test('a stream that does not reduce to one value is rejected', () => {
  // Two operands and no operator: a well-formed formula leaves exactly one string behind.
  assert.equal(decodeFormula(bytes(0x1e, 1, 0, 0x1e, 2, 0), NONE, SCOPE), undefined);
  // An operator with nothing to consume.
  assert.equal(decodeFormula(bytes(0x03), NONE, SCOPE), undefined);
});

test('a numeric literal the format cannot spell makes the formula undecodable', () => {
  // `PtgNum` whose eight bytes are IEEE-754 +Infinity. `String(value)` renders it `Infinity`, which
  // used to reach the model as the formula text `INFINITY` and be written back out as a formula Excel
  // reports as damaged. There is no spelling for it, so there is no formula.
  const infinity = bytes(0x1f, 0, 0, 0, 0, 0, 0, 0xf0, 0x7f);
  assert.equal(decodeFormula(infinity, NONE, SCOPE), undefined);
  const notANumber = bytes(0x1f, 0, 0, 0, 0, 0, 0, 0xf8, 0x7f);
  assert.equal(decodeFormula(notANumber, NONE, SCOPE), undefined);
});

test('a numeric literal keeps the exponent case Excel writes', () => {
  // 1e+21, where JavaScript spells the exponent lowercase and Excel spells it `1E+21`. The two codecs
  // used to disagree, so the same literal read back differently depending on which file it came from.
  const large = new Uint8Array(9);
  large[0] = 0x1f;
  new DataView(large.buffer).setFloat64(1, 1e21, true);
  assert.equal(decodeFormula(large, NONE, SCOPE), '1E+21');
});

test('a token running past the end of its own stream costs the formula, not the workbook', () => {
  // `PtgInt` declaring a 16-bit operand with one byte behind it. This used to escape as an
  // `XlsbParseError` and abort the entire read, while the *same* damage spelled as an unrecognised
  // opcode degraded to the cached result. The record-level bound is what fails a malformed file
  // closed; a formula that runs off its own end is one undecodable formula.
  assert.equal(decodeFormula(bytes(0x1e, 1), NONE, SCOPE), undefined);
});

test('an array constant declaring more elements than any workbook holds is refused', () => {
  // The elements are read from the trailing block, which cannot outrun its record, but the row and
  // column counts are multiplied before any of it is read, so a forged pair must not buy a loop.
  const rgce = bytes(0x60, ...new Array<number>(14).fill(0));
  const rgcb = bytes(0xff, 0xff, 0xff, 0x0f, 0xff, 0xff, 0xff, 0x0f);
  assert.equal(decodeFormula(rgce, rgcb, SCOPE), undefined);
});

test('a deferred formula names the cell it defers to, whose column lives in the extra block', () => {
  assert.deepEqual(formulaAnchor(bytes(0x01, 3, 0, 0, 0), bytes(5, 0, 0, 0)), {row: 3, column: 5});
  // An ordinary formula defers to nothing.
  assert.equal(formulaAnchor(ref(0, 0), NONE), undefined);
  // A deferral whose column is missing is not one.
  assert.equal(formulaAnchor(bytes(0x01, 3, 0, 0, 0), NONE), undefined);
});

test('the function table matches the specification at every run boundary', () => {
  // The table is transcribed as contiguous runs, so a single miscounted entry would shift every name
  // after it. These anchors are the first and last of each run, plus the gaps between them.
  assert.equal(functionNameFor(0x0000), 'COUNT');
  assert.equal(functionNameFor(0x0004), 'SUM');
  assert.equal(functionNameFor(0x0013), 'PI');
  assert.equal(functionNameFor(0x00a9), 'COUNTA');
  assert.equal(functionNameFor(0x00c9), 'UNREGISTER');
  assert.equal(functionNameFor(0x00cc), 'USDOLLAR');
  assert.equal(functionNameFor(0x00d8), 'RANK');
  assert.equal(functionNameFor(0x00db), 'ADDRESS');
  assert.equal(functionNameFor(0x00f8), 'PAUSE');
  assert.equal(functionNameFor(0x00fb), 'RESUME');
  assert.equal(functionNameFor(0x00ff), 'User Defined Function');
  assert.equal(functionNameFor(0x014c), 'TINV');
  assert.equal(functionNameFor(0x014e), 'MOVIE.COMMAND');
  assert.equal(functionNameFor(0x017b), 'RTD');
  assert.equal(functionNameFor(0x017c), 'CUBEVALUE');
  assert.equal(functionNameFor(0x01e4), 'AVERAGEIFS');

  for (const gap of [0x00ca, 0x00cb, 0x00d9, 0x00da, 0x00f9, 0x00fa, 0x014d, 0x01e5]) {
    assert.equal(functionNameFor(gap), undefined);
  }
});

test('a fixed-arity function knows how many operands belong to its call', () => {
  const arityOf = (index: number): number | 'variadic' | undefined => builtinFunctionAt(index)?.[1];
  assert.equal(arityOf(0x0013), 0, 'PI');
  assert.equal(arityOf(0x0018), 1, 'ABS');
  assert.equal(arityOf(0x001b), 2, 'ROUND');
  assert.equal(arityOf(0x001f), 3, 'MID');
  assert.equal(arityOf(0x0077), 4, 'REPLACE');
  // A variadic function has no fixed arity: its call token carries its own count instead.
  assert.equal(arityOf(0x0004), 'variadic', 'SUM');
  assert.equal(arityOf(0x0001), 'variadic', 'IF');
});

test('a fixed-arity call with no arity to go on is not decoded', () => {
  // PtgFunc citing SUM: variadic, so the stream states no argument count and the call cannot be
  // reconstructed. Excel only ever writes fixed-arity functions through this token.
  assert.equal(decodeFormula(bytes(...ref(0, 0), 0x41, 0x04, 0x00), NONE, SCOPE), undefined);
  // PtgFunc citing ABS, which takes exactly one.
  assert.equal(decodeFormula(bytes(...ref(0, 0), 0x41, 0x18, 0x00), NONE, SCOPE), 'ABS(A1)');
});

// Both arrived with Excel 2007 and sit past RTD, where the [MS-XLS] table ends. Excel 16.0 (build
// 20326) saves `=IFERROR(1,0)` as 1, an ifError attribute, 0, a goto attribute and a PtgFunc citing
// 0x01E0, and a SUMIFS as a PtgFuncVar citing 0x01E2 with its count.
test('a function the BIFF12 table adds past RTD decodes by its index', () => {
  assert.equal(
    decodeFormula(
      bytes(0x1e, 1, 0, 0x19, 0x80, 3, 0, 0x1e, 0, 0, 0x19, 0x08, 2, 0, 0x41, 0xe0, 0x01),
      NONE,
      SCOPE,
    ),
    'IFERROR(1,0)',
  );
  assert.equal(
    decodeFormula(bytes(...ref(0, 0), ...ref(1, 0), 0x1e, 1, 0, 0x42, 3, 0xe2, 0x01), NONE, SCOPE),
    'SUMIFS(A1,A2,1)',
  );
});
