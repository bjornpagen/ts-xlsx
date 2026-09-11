import assert from 'node:assert/strict';
import {test} from 'node:test';

import {
  definedNameKeys,
  formulaNamesInScope,
  mangleFormula,
  mangleFunctions,
  mangleFunctionValues,
  mangleParams,
  quoteSheetName,
  unmangleFunctions,
} from './formula.ts';

// A workbook that defines no names, where every function passed as a value takes its prefix.
const NO_NAMES: ReadonlySet<string> = new Set();

// Counted as a bracket, an escaped `[` left the structured reference open to the end of the formula,
// so a modern function after it went unprefixed and Excel could not evaluate it.
test('an escaped bracket inside a structured reference neither closes nor extends its opaque region', () => {
  assert.equal(mangleFunctions("T[[a']b]]"), "T[[a']b]]");
  assert.equal(
    mangleFunctions("SUBTOTAL(109,T[[a'[b]])+FILTER(x)"),
    "SUBTOTAL(109,T[[a'[b]])+_xlfn._xlws.FILTER(x)",
  );
});

test('a modern function called by its plain name gains the _xlfn. prefix', () => {
  assert.equal(mangleFunctions('SORTBY(B1:D1,B2:D2)'), '_xlfn.SORTBY(B1:D1,B2:D2)');
  assert.equal(mangleFunctions('XLOOKUP(1,B:B,C:C)'), '_xlfn.XLOOKUP(1,B:B,C:C)');
});

test('a legacy function is left untouched', () => {
  assert.equal(mangleFunctions('SUM(A1:A9)'), 'SUM(A1:A9)');
  assert.equal(mangleFunctions('IF(A1>0,"y","n")'), 'IF(A1>0,"y","n")');
});

test('nested modern functions each get the prefix, legacy ones do not', () => {
  assert.equal(mangleFunctions('SUM(FILTER(A:A,B:B=1))'), 'SUM(_xlfn._xlws.FILTER(A:A,B:B=1))');
  assert.equal(
    mangleFunctions('COUNTA(UNIQUE(FILTER(a,b=1)))'),
    'COUNTA(_xlfn.UNIQUE(_xlfn._xlws.FILTER(a,b=1)))',
  );
});

test('an already-prefixed function is not double-prefixed', () => {
  assert.equal(mangleFunctions('_xlfn.XLOOKUP(1,B:B,C:C)'), '_xlfn.XLOOKUP(1,B:B,C:C)');
  assert.ok(!mangleFunctions('_xlfn.XLOOKUP(1,B:B,C:C)').includes('_xlfn._xlfn'));
});

test('a function name inside a string literal is never mangled', () => {
  assert.equal(mangleFunctions('IF(A1="FILTER(",1,2)'), 'IF(A1="FILTER(",1,2)');
  assert.equal(mangleFunctions('CONCAT("SORT()",A1)'), '_xlfn.CONCAT("SORT()",A1)');
});

test('mangling introduces no @ implicit-intersection operator', () => {
  const out = mangleFunctions('IFS(B1>0,"pos",B1<0,"neg")');
  assert.equal(out, '_xlfn.IFS(B1>0,"pos",B1<0,"neg")');
  assert.ok(!/(^|[^A-Za-z0-9_])@/.test(out));
});

test('matching is case-insensitive on the function name but preserves its casing', () => {
  assert.equal(mangleFunctions('filter(A:A,B:B=1)'), '_xlfn._xlws.filter(A:A,B:B=1)');
  assert.equal(mangleFunctions('Xlookup(1,A:A,B:B)'), '_xlfn.Xlookup(1,A:A,B:B)');
});

test('a LET/LAMBDA formula gets the _xlfn. prefix on every modern function', () => {
  const out = mangleFunctions(
    'LET(a,B2:B9,b,BYROW(a,LAMBDA(r,SUM(r))),COUNTA(UNIQUE(FILTER(a,b=1))))',
  );
  assert.ok(out.includes('_xlfn.LET'));
  assert.ok(out.includes('_xlfn.BYROW'));
  assert.ok(out.includes('_xlfn.LAMBDA'));
  assert.ok(out.includes('_xlfn.UNIQUE'));
  assert.ok(out.includes('_xlfn._xlws.FILTER'));
  assert.ok(out.includes('SUM(r)'));
  assert.ok(out.includes('COUNTA('));
});

test('a bare-name post-2007 function (trig / bitwise / engineering) gains the prefix', () => {
  assert.equal(mangleFunctions('SEC(A1)'), '_xlfn.SEC(A1)');
  assert.equal(mangleFunctions('BITAND(5,3)'), '_xlfn.BITAND(5,3)');
  assert.equal(mangleFunctions('IMCOSH("2+i")'), '_xlfn.IMCOSH("2+i")');
  assert.equal(mangleFunctions('AGGREGATE(9,4,A1:A9)'), '_xlfn.AGGREGATE(9,4,A1:A9)');
  assert.equal(mangleFunctions('XOR(A1,B1)'), '_xlfn.XOR(A1,B1)');
  assert.equal(mangleFunctions('ISOWEEKNUM(A1)'), '_xlfn.ISOWEEKNUM(A1)');
});

test('a pre-2007 function whose name resembles a modern one is left untouched', () => {
  // SIN/COS/TAN and GAMMALN predate the frozen grammar and must NOT be prefixed, even though
  // SEC/CSC/COT and GAMMA (their newer cousins) are.
  assert.equal(mangleFunctions('SIN(A1)'), 'SIN(A1)');
  assert.equal(mangleFunctions('COS(A1)'), 'COS(A1)');
  assert.equal(mangleFunctions('GAMMALN(A1)'), 'GAMMALN(A1)');
  assert.equal(mangleFunctions('WEEKNUM(A1)'), 'WEEKNUM(A1)');
});

// Excel stores FILTER and SORT as `_xlfn._xlws.FILTER` and `_xlfn._xlws.SORT`. Knowing only `_xlfn.`,
// the reader left `_xlws.FILTER` in the model, the writer then wrote it back without `_xlfn.`, and an
// authored FILTER went out as `_xlfn.FILTER`, a spelling Excel does not write.
test('the worksheet-only functions take _xlws. after _xlfn., and unmangling strips both', () => {
  assert.equal(mangleFunctions('FILTER(A1:A3,A1:A3>0)'), '_xlfn._xlws.FILTER(A1:A3,A1:A3>0)');
  assert.equal(mangleFunctions('SORT(A1:A3)'), '_xlfn._xlws.SORT(A1:A3)');
  assert.equal(mangleFunctions('PY(0,0)'), '_xlfn._xlws.PY(0,0)');
  assert.equal(
    unmangleFunctions('_xlfn._xlws.FILTER(A1:A3,A1:A3>0)', NO_NAMES),
    'FILTER(A1:A3,A1:A3>0)',
  );
  assert.equal(unmangleFunctions('SUM(_xlfn._xlws.SORT(A1:A3))', NO_NAMES), 'SUM(SORT(A1:A3))');
  assert.equal(
    mangleFunctions('_xlfn._xlws.SORT(A1:A3)'),
    '_xlfn._xlws.SORT(A1:A3)',
    'not doubled',
  );
});

// Each of these was written bare, and current Excel reads a future function without its prefix as
// `#NAME?`. The spellings are what Excel Desktop itself saved.
test('every family Excel writes with _xlfn. that the registry lacked is prefixed', () => {
  for (const call of [
    'CEILING.MATH(4.3)',
    'FLOOR.MATH(4.3)',
    'FORECAST.LINEAR(1,A1:A3,A1:A3)',
    'FORECAST.ETS(4,A1:A3,A1:A3)',
    'FORECAST.ETS.CONFINT(4,A1:A3,A1:A3)',
    'FORECAST.ETS.SEASONALITY(A1:A3,A1:A3)',
    'FORECAST.ETS.STAT(A1:A3,A1:A3,1)',
    'GROUPBY(A1:A3,A1:A3,x)',
    'PIVOTBY(A1:A3,A1:A3,A1:A3,x)',
    'PERCENTOF(A1:A3,A1:A3)',
    'REGEXTEST(A1,"a")',
    'REGEXEXTRACT(A1,"a")',
    'REGEXREPLACE(A1,"a","b")',
    'TRIMRANGE(A1:A3)',
    'ANCHORARRAY(E1)',
    'SINGLE(A1:A3)',
    'FIELDVALUE(A1,"x")',
    'STOCKHISTORY("MSFT",1)',
    'IMAGE("https://example.com/a.png")',
    'TRANSLATE("hola","es","en")',
    'DETECTLANGUAGE("hola")',
    'PYTHON_STR(A1)',
  ]) {
    assert.equal(mangleFunctions(call), `_xlfn.${call}`, call);
  }
});

// [MS-XLSX] lists these among its future functions with no prefix, and Excel writes them bare. ISO.CEILING
// used to be prefixed, which is a spelling Excel does not read.
test('a post-2007 function Excel writes bare is left bare', () => {
  for (const call of [
    'NETWORKDAYS.INTL(1,30)',
    'WORKDAY.INTL(1,5)',
    'ISO.CEILING(4.3)',
    'ECMA.CEILING(4.3,1)',
  ]) {
    assert.equal(mangleFunctions(call), call);
  }
});

test('a dotted 2010 statistical function is matched whole and prefixed', () => {
  assert.equal(mangleFunctions('NORM.DIST(A1,0,1,TRUE)'), '_xlfn.NORM.DIST(A1,0,1,TRUE)');
  assert.equal(mangleFunctions('BETA.INV(0.5,2,3)'), '_xlfn.BETA.INV(0.5,2,3)');
  assert.equal(mangleFunctions('T.DIST.2T(2,10)'), '_xlfn.T.DIST.2T(2,10)');
  assert.equal(mangleFunctions('NORM.S.INV(0.9)'), '_xlfn.NORM.S.INV(0.9)');
});

test('a dotted function only gets one prefix, on the whole name, not per segment', () => {
  const out = mangleFunctions('CHISQ.DIST.RT(3,2)');
  assert.equal(out, '_xlfn.CHISQ.DIST.RT(3,2)');
  assert.ok(!out.includes('_xlfn.DIST'), 'the tail segment must not be prefixed on its own');
});

test('an already-prefixed dotted function is not double-prefixed', () => {
  assert.equal(mangleFunctions('_xlfn.NORM.DIST(A1,0,1,TRUE)'), '_xlfn.NORM.DIST(A1,0,1,TRUE)');
  assert.ok(!mangleFunctions('_xlfn.NORM.DIST(A1,0,1,TRUE)').includes('_xlfn._xlfn'));
});

test('a dotted function nested beside plain and legacy functions is prefixed correctly', () => {
  assert.equal(
    mangleFunctions('SUM(NORM.DIST(A1,0,1,TRUE),STDEV.S(B:B),AVERAGE(C:C))'),
    'SUM(_xlfn.NORM.DIST(A1,0,1,TRUE),_xlfn.STDEV.S(B:B),AVERAGE(C:C))',
  );
});

test('a decimal literal adjacent to a dotted call is not mistaken for a function', () => {
  assert.equal(mangleFunctions('NORM.DIST(A1,0,1,TRUE)*1.5'), '_xlfn.NORM.DIST(A1,0,1,TRUE)*1.5');
});

test('unmangle strips _xlfn. and _xlpm. back to the plain names', () => {
  assert.equal(unmangleFunctions('_xlfn.XLOOKUP(1,B:B,C:C)', NO_NAMES), 'XLOOKUP(1,B:B,C:C)');
  assert.equal(unmangleFunctions('_xlfn.LET(_xlpm.a,B2:B9,_xlpm.a)', NO_NAMES), 'LET(a,B2:B9,a)');
  assert.equal(
    unmangleFunctions('_xlfn.NORM.DIST(A1,0,1,TRUE)', NO_NAMES),
    'NORM.DIST(A1,0,1,TRUE)',
  );
});

test('mangle then unmangle round-trips a plain formula', () => {
  for (const f of [
    'FILTER(A:A,B:B=1)',
    'SUM(A1:A9)',
    'IFS(B1>0,"pos")',
    'XLOOKUP(1,B:B,C:C)',
    'NORM.DIST(A1,0,1,TRUE)',
    'T.DIST.2T(2,10)',
  ]) {
    assert.equal(unmangleFunctions(mangleFunctions(f), NO_NAMES), f);
  }
});

test('a LET parameter is _xlpm.-prefixed at its declaration and every reference', () => {
  assert.equal(mangleParams('LET(x,1,x+1)'), 'LET(_xlpm.x,1,_xlpm.x+1)');
  assert.equal(
    mangleParams('LET(a,B2:B9,b,2,a+b)'),
    'LET(_xlpm.a,B2:B9,_xlpm.b,2,_xlpm.a+_xlpm.b)',
  );
});

test('a LAMBDA prefixes every parameter but leaves its body cells and legacy calls alone', () => {
  assert.equal(mangleParams('LAMBDA(a,b,a+b)'), 'LAMBDA(_xlpm.a,_xlpm.b,_xlpm.a+_xlpm.b)');
  assert.equal(mangleParams('LAMBDA(r,SUM(r)+A1)'), 'LAMBDA(_xlpm.r,SUM(_xlpm.r)+A1)');
});

test('the _xlpm. prefix is scoped: a same-named reference outside the binding is untouched', () => {
  // The leading `x` is a defined-name reference, not the LET parameter: it must survive verbatim.
  assert.equal(mangleParams('x+LET(x,1,x)'), 'x+LET(_xlpm.x,1,_xlpm.x)');
  assert.equal(mangleParams('LET(x,1,x)+x'), 'LET(_xlpm.x,1,_xlpm.x)+x');
});

test('a parameter name inside a string literal is never mistaken for a reference', () => {
  assert.equal(mangleParams('LET(x,1,"x is one"&x)'), 'LET(_xlpm.x,1,"x is one"&_xlpm.x)');
});

test('a parameter name inside a single-quoted sheet name is never prefixed', () => {
  // The sheet name `x` collides with the LET parameter; skipping the quoted region keeps it verbatim.
  assert.equal(mangleParams("LET(x,1,'x sheet'!A1&x)"), "LET(_xlpm.x,1,'x sheet'!A1&_xlpm.x)");
});

test('a parameter name inside a structured reference is never prefixed', () => {
  assert.equal(mangleParams('LET(x,1,Table[x]&x)'), 'LET(_xlpm.x,1,Table[x]&_xlpm.x)');
});

// Excel's names are case-insensitive, and it writes every reference in the declared spelling: it saves
// `LET(x,1,X+1)` as `_xlfn.LET(_xlpm.x,1,_xlpm.x+1)`. Matched case-sensitively, the `X` stayed bare and
// read back as a defined name.
test('a parameter reference matches its declaration whatever its case, in the declared spelling', () => {
  assert.equal(mangleParams('LET(x,1,X+1)'), 'LET(_xlpm.x,1,_xlpm.x+1)');
  assert.equal(
    mangleFormula('LAMBDA(Val,val*2)(3)', NO_NAMES),
    '_xlfn.LAMBDA(_xlpm.Val,_xlpm.Val*2)(3)',
  );
  assert.equal(
    mangleParams('let(total,SUM(A1:A3),TOTAL*2)'),
    'let(_xlpm.total,SUM(A1:A3),_xlpm.total*2)',
  );
});

test('a shadowing binding decides the spelling of the references inside it', () => {
  assert.equal(
    mangleParams('LET(x,1,LET(X,2,x)+x)'),
    'LET(_xlpm.x,1,LET(_xlpm.X,2,_xlpm.X)+_xlpm.x)',
  );
});

test('a lambda-valued parameter called as a function is prefixed', () => {
  assert.equal(
    mangleParams('LET(f,LAMBDA(v,v+1),f(5))'),
    'LET(_xlpm.f,LAMBDA(_xlpm.v,_xlpm.v+1),_xlpm.f(5))',
  );
});

test('nested LET/LAMBDA scopes each bind their own parameters', () => {
  assert.equal(
    mangleParams('LET(a,B2:B9,BYROW(a,LAMBDA(r,SUM(r))))'),
    'LET(_xlpm.a,B2:B9,BYROW(_xlpm.a,LAMBDA(_xlpm.r,SUM(_xlpm.r))))',
  );
});

test('a formula with no LET/LAMBDA passes through parameter mangling unchanged', () => {
  for (const f of [
    'SUM(A1:A9)',
    'IF(A1>0,"y","n")',
    'NORM.DIST(A1,0,1,TRUE)',
    'Table1[[#Data],[Col]]',
  ]) {
    assert.equal(mangleParams(f), f);
  }
});

test('mangleFormula applies both prefixes in the correct order', () => {
  assert.equal(mangleFormula('LET(x,1,x+1)', NO_NAMES), '_xlfn.LET(_xlpm.x,1,_xlpm.x+1)');
  assert.equal(
    mangleFormula(
      'LET(a,B2:B9,b,BYROW(a,LAMBDA(r,SUM(r))),COUNTA(UNIQUE(FILTER(a,b=1))))',
      NO_NAMES,
    ),
    '_xlfn.LET(_xlpm.a,B2:B9,_xlpm.b,_xlfn.BYROW(_xlpm.a,_xlfn.LAMBDA(_xlpm.r,SUM(_xlpm.r))),COUNTA(_xlfn.UNIQUE(_xlfn._xlws.FILTER(_xlpm.a,_xlpm.b=1))))',
  );
});

test('mangleFormula then unmangle round-trips a LET/LAMBDA formula', () => {
  for (const f of ['LET(x,1,x+1)', 'LAMBDA(a,b,a+b)', 'LET(f,LAMBDA(v,v+1),f(5))', 'SUM(A1:A9)']) {
    assert.equal(unmangleFunctions(mangleFormula(f, NO_NAMES), NO_NAMES), f);
  }
});

test('quoteSheetName leaves a plain identifier bare and quotes anything else', () => {
  assert.equal(quoteSheetName('Data'), 'Data');
  assert.equal(quoteSheetName('_totals.2024'), '_totals.2024');
  assert.equal(quoteSheetName('Odd Name'), "'Odd Name'");
  assert.equal(quoteSheetName('2024'), "'2024'");
  // A name that would read as a cell address has to be quoted, or `A1!A1` is ambiguous.
  assert.equal(quoteSheetName('A1'), "'A1'");
  assert.equal(quoteSheetName("Bob's"), "'Bob''s'");
});

test('quoteSheetName quotes an R1C1 reference and a boolean, and leaves a column past XFD bare', () => {
  // As Excel 16.0 build 20326 spells `='<name>'!B2` back.
  for (const name of ['R1C1', 'RC', 'R', 'C', 'R1C', 'R1X', 'TRUE', 'false']) {
    assert.equal(quoteSheetName(name), `'${name}'`, name);
  }
  for (const name of ['XFE1', 'RR', 'CC', 'RCX', 'R1.5', 'TRUE1']) {
    assert.equal(quoteSheetName(name), name, name);
  }
  assert.equal(quoteSheetName('Data', 'RC'), "'Data:RC'", 'either end of a span');
});

test('quoteSheetName quotes a 3-D span as a whole or not at all', () => {
  // The quotes delimit the sheet reference, not either endpoint, so one awkward name puts both
  // inside them, which is how Excel spells `'Odd Name:More'!A1`.
  assert.equal(quoteSheetName('Data', 'More'), 'Data:More');
  assert.equal(quoteSheetName('Odd Name', 'More'), "'Odd Name:More'");
  assert.equal(quoteSheetName('Data', 'Odd Name'), "'Data:Odd Name'");
});

// Excel 16.0 (build 20326) stores a built-in function a formula passes as a value under `_xleta.`,
// uppercased, in every position it was given one. Written bare, the name is a reference to a defined
// name spelled like the function, and Excel shows `#NAME?` when the workbook defines none.
test('a function passed as a value is written under _xleta., uppercased, wherever it sits', () => {
  assert.equal(mangleFormula('BYROW(A1:A3,sum)', NO_NAMES), '_xlfn.BYROW(A1:A3,_xleta.SUM)');
  assert.equal(
    mangleFormula('LET(f,MAX,f(A1:A3))', NO_NAMES),
    '_xlfn.LET(_xlpm.f,_xleta.MAX,_xlpm.f(A1:A3))',
  );
  assert.equal(mangleFormula('ISERROR(-SUM+1)', NO_NAMES), 'ISERROR(-_xleta.SUM+1)');
  assert.equal(
    mangleFormula('HSTACK(SUM,(MAX))', NO_NAMES),
    '_xlfn.HSTACK(_xleta.SUM,(_xleta.MAX))',
  );
  assert.equal(mangleFormula('BYROW(A1:A3, ABS )', NO_NAMES), '_xlfn.BYROW(A1:A3, _xleta.ABS )');
  assert.equal(mangleFormula('MAP(A1:A3,T.TEST)', NO_NAMES), '_xlfn.MAP(A1:A3,_xleta.T.TEST)');
  // A future function passed as a value is not being called, so it takes no `_xlfn.` beside it.
  assert.equal(
    mangleFormula('BYROW(A1:A3,CONCAT)+BYROW(A1:A3,FILTER)', NO_NAMES),
    '_xlfn.BYROW(A1:A3,_xleta.CONCAT)+_xlfn.BYROW(A1:A3,_xleta.FILTER)',
  );
});

test('a function name called, naming a container, qualified or inside a literal is not a value', () => {
  for (const formula of [
    'SUM(A1:A3)',
    'SUM (A1:A3)',
    'SUM!A1+SUM:Other!A1',
    "S1!SUM+'My Sheet'!MAX+[1]!MIN",
    'ISNA(#N/A)',
    'Rate[Amount]+T[[#Data],[SUM]]',
    '"SUM"&T1',
    'LOG10+TRUE',
    'LAMBDA+GET.CELL',
    '_xlpm.SUM+_xleta.MAX',
  ]) {
    assert.equal(mangleFunctionValues(formula, NO_NAMES), formula, formula);
  }
});

test('a LET or LAMBDA parameter spelled like a function is a parameter, not a function value', () => {
  // As Excel saved `=LET(SUM,A1:A3,SUM)` and `=LAMBDA(SUM,SUM)(1)`.
  assert.equal(
    mangleFormula('LET(SUM,A1:A3,SUM)', NO_NAMES),
    '_xlfn.LET(_xlpm.SUM,A1:A3,_xlpm.SUM)',
  );
  assert.equal(
    mangleFormula('LAMBDA(SUM,SUM)(1)', NO_NAMES),
    '_xlfn.LAMBDA(_xlpm.SUM,_xlpm.SUM)(1)',
  );
});

// With `LEN` defined on S2, Excel wrote `MAP(A1:A3,LEN)` bare on S2, where the name is visible and
// meant, and as `_xleta.LEN` on S1, where it is not.
test('a defined name visible from the formula keeps a function name bare, since there it means the name', () => {
  const names = [{name: 'len', scope: 'S2'}, {name: 'Max'}];
  const formula = 'MAP(A1:A3,LEN)+MAP(A1:A3,MAX)';
  assert.equal(
    mangleFormula(formula, formulaNamesInScope(names, 's2')),
    '_xlfn.MAP(A1:A3,LEN)+_xlfn.MAP(A1:A3,MAX)',
  );
  assert.equal(
    mangleFormula(formula, formulaNamesInScope(names, 'S1')),
    '_xlfn.MAP(A1:A3,_xleta.LEN)+_xlfn.MAP(A1:A3,MAX)',
  );
  // A workbook-level name's own formula sees the workbook-level names alone.
  assert.equal(
    mangleFormula(formula, formulaNamesInScope(names, undefined)),
    '_xlfn.MAP(A1:A3,_xleta.LEN)+_xlfn.MAP(A1:A3,MAX)',
  );
});

test('reading sheds _xleta. unless the workbook defines a name spelled like the function', () => {
  assert.equal(unmangleFunctions('_xlfn.BYROW(A1:A3,_xleta.SUM)', NO_NAMES), 'BYROW(A1:A3,SUM)');
  assert.equal(
    unmangleFunctions('_xlfn.MAP(A1:A3,_xleta.LEN)', definedNameKeys([{name: 'Len'}])),
    'MAP(A1:A3,_xleta.LEN)',
  );
  assert.equal(unmangleFunctions('"_xleta.SUM"&_xleta.T.TEST', NO_NAMES), '"_xleta.SUM"&T.TEST');
});

test('a function passed as a value survives the write and the read whatever names the workbook defines', () => {
  for (const names of [NO_NAMES, definedNameKeys([{name: 'SUM'}, {name: 'MAX'}])]) {
    for (const formula of ['BYROW(A1:A3,SUM)', 'LET(f,MAX,f(A1:A3))', 'ISERROR(SUM+1)']) {
      assert.equal(unmangleFunctions(mangleFormula(formula, names), names), formula, formula);
    }
  }
  const captured = definedNameKeys([{name: 'SUM'}]);
  assert.equal(
    unmangleFunctions(mangleFormula('MAP(A1:A3,_xleta.SUM)', captured), captured),
    'MAP(A1:A3,_xleta.SUM)',
  );
});
