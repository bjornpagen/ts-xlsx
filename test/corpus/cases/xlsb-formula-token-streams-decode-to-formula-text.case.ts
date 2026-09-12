import type {Assert, Case, CorpusApi} from '../case.ts';

// An `.xlsx` stores a formula as the text the author typed. An `.xlsb` stores the same formula as a
// `Ptg` token stream: postfix, with functions cited by number, sheets by index into a table, and
// defined names by position. So "the two forms produce one model" is a much stronger claim about
// formulas than about anything else in the file: it means a stack machine reconstructed the exact
// characters Excel would have written, parentheses and `$` anchors and quoting included.
//
// The fixtures are a pair Excel itself saved from one in-memory workbook, `source.xlsb` and
// `source.xlsx`, so the XML twin is an independent oracle rather than something this library
// produced. The workbook is a grammar tour: one formula per token class. See `author.ps1` beside
// them, and the spec `docs/knowledge/specs/xlsb-binary-format-output.md`.
export default {
  id: 'xlsb-formula-token-streams-decode-to-formula-text',
  cluster: 'xlsx-io',
  description:
    'A binary .xlsb formula, stored as a Ptg token stream, reads back as the same formula text its ' +
    '.xlsx twin states: operators and parentheses, every operand kind, references including 3-D ' +
    'and whole-column, built-in and post-2007 function calls, array constants, and defined names.',
  provenance: {source: 'upstream-issue'},
  behavior: [
    {
      name: 'every formula reads back as the text its XML twin states',
      expect(api: CorpusApi, assert: Assert) {
        const result = api.xlsbFormulaTextMatchesXlsxTwin();
        // The message carries each disagreement, so a regression names the token class that broke.
        assert.deepEqual(result.differences, []);
        // A guard on the oracle itself: if the fixture ever stopped carrying formulas, an empty
        // comparison would otherwise pass silently.
        assert.equal(result.compared > 30, true);
      },
    },
    {
      name: 'operator precedence survives because parentheses are stored, not inferred',
      expect(api: CorpusApi, assert: Assert) {
        // A postfix stream carries no parentheses of its own: `1+2*3` and `(1+2)*3` differ only in
        // token order. Excel records the author's parentheses explicitly, so a decoder that instead
        // re-derived them from precedence would turn the second into the first.
        assert.equal(api.xlsbFormula('Calc', 'C1').formula, '1+2*3');
        assert.equal(api.xlsbFormula('Calc', 'C2').formula, '(1+2)*3');
      },
    },
    {
      name: 'unary and postfix operators keep their operand',
      expect(api: CorpusApi, assert: Assert) {
        assert.equal(api.xlsbFormula('Calc', 'C4').formula, '-A1');
        assert.equal(api.xlsbFormula('Calc', 'C5').formula, 'A1%');
        assert.equal(api.xlsbFormula('Calc', 'C6').formula, 'A2^3');
      },
    },
    {
      name: 'literal operands decode to their own spelling',
      expect(api: CorpusApi, assert: Assert) {
        assert.equal(api.xlsbFormula('Calc', 'C3').formula, '1.5+2');
        // A quote inside a string literal is doubled, as the formula grammar (not the stored text)
        // requires: the token holds `say "hi"` and the formula must read `"say ""hi"""`.
        assert.equal(api.xlsbFormula('Calc', 'C8').formula, '"say ""hi"""');
        assert.equal(api.xlsbFormula('Calc', 'C10').formula, 'A1<>A2');
        assert.equal(api.xlsbFormula('Calc', 'C11').formula, 'TRUE');
        assert.equal(api.xlsbFormula('Calc', 'C12').formula, '#N/A');
      },
    },
    {
      name: 'a reference keeps the axes its author anchored',
      expect(api: CorpusApi, assert: Assert) {
        // Row and column relativity are two flag bits on one packed word; getting either backwards
        // silently moves every `$`.
        assert.equal(api.xlsbFormula('Calc', 'C13').formula, '$A$1+A$2+$A3');
      },
    },
    {
      name: 'a whole-column or whole-row reference reads in its abbreviated form',
      expect(api: CorpusApi, assert: Assert) {
        // Stored as an ordinary range spanning every row (or every column) of the grid. `A:A` is not
        // shorthand a reader may choose: it is the only spelling Excel writes.
        assert.equal(api.xlsbFormula('Calc', 'C23').formula, 'SUM(A:A)');
        assert.equal(api.xlsbFormula('Calc', 'C24').formula, 'SUM(Data!2:2)');
      },
    },
    {
      name: 'a sheet-qualified reference resolves its sheet through the extern-sheet table',
      expect(api: CorpusApi, assert: Assert) {
        // The token names no sheet: it carries an index into a table of sheet *spans*, which is why
        // a single-sheet reference and a three-dimensional one decode through the same indirection.
        assert.equal(api.xlsbFormula('Calc', 'C19').formula, 'SUM(Data!A1:B2)');
        assert.equal(api.xlsbFormula('Calc', 'C20').formula, 'Data!$A$1');
        assert.equal(api.xlsbFormula('Calc', 'C22').formula, 'SUM(Data:More!A1)');
        // A name that is not a plain identifier is quoted, a decision the reader must make itself,
        // since the stored form has no quoting at all.
        assert.equal(api.xlsbFormula('Calc', 'C21').formula, "'Odd Name'!A1");
      },
    },
    {
      name: 'a reference to nothing reads as the reference error, not as a cell',
      expect(api: CorpusApi, assert: Assert) {
        assert.equal(api.xlsbFormula('Calc', 'C35').formula, '#REF!+1');
      },
    },
    {
      name: 'a function call recovers its name and its arguments',
      expect(api: CorpusApi, assert: Assert) {
        // A call token cites its function by number. A fixed-arity call carries no argument count at
        // all: how many operands belong to it is a property of the function, so `PI()` and
        // `SUM(A1:A5)` are distinguished by the table, not the stream.
        assert.equal(api.xlsbFormula('Calc', 'C18').formula, 'PI()');
        assert.equal(api.xlsbFormula('Calc', 'C14').formula, 'SUM(A1:A5)');
        assert.equal(api.xlsbFormula('Calc', 'C15').formula, 'SUM(A1:A5)/COUNT(A1:A5)');
        assert.equal(api.xlsbFormula('Calc', 'C16').formula, 'IF(A1>0,"pos","neg")');
        assert.equal(api.xlsbFormula('Calc', 'C17').formula, 'SUM(A1:A2,A4)');
      },
    },
    {
      name: 'an omitted argument stays omitted',
      expect(api: CorpusApi, assert: Assert) {
        // The gap in `IF(A1>0,,1)` is a real operand whose text is nothing; dropping it would shift
        // every later argument one place left.
        assert.equal(api.xlsbFormula('Calc', 'C34').formula, 'IF(A1>0,,1)');
      },
    },
    {
      name: 'a post-2007 function is called through the placeholder name Excel registers for it',
      expect(api: CorpusApi, assert: Assert) {
        // The built-in function table froze around Excel 2007, so a newer function is stored as a
        // call to "user defined" whose name comes from a hidden defined name. The model holds the
        // readable name either way, exactly as it does when reading the XML form's `_xlfn.` prefix.
        assert.equal(api.xlsbFormula('Calc', 'C29').formula, 'TEXTJOIN(",",TRUE,A1:A3)');
      },
    },
    {
      name: 'reference-set operators keep their punctuation',
      expect(api: CorpusApi, assert: Assert) {
        // Intersection is a space and union is a comma, operators that look like formatting, and
        // that a reader dropping them would turn into a different formula that still parses.
        assert.equal(api.xlsbFormula('Calc', 'C28').formula, 'SUM(A1:A3 A2:A5)');
        assert.equal(api.xlsbFormula('Calc', 'C30').formula, 'SUM((A1:A2,A4:A5))');
      },
    },
    {
      name: 'an array constant decodes with its shape and its element types',
      expect(api: CorpusApi, assert: Assert) {
        // The elements live outside the token stream, in a trailing block whose row and column counts
        // are the only thing saying where each row ends, so the non-square constant is the case that
        // catches a decoder reading those two counts the wrong way round.
        assert.equal(api.xlsbFormula('Calc', 'C31').formula, 'SUM({1,2;3,4})');
        assert.equal(api.xlsbFormula('Calc', 'C32').formula, 'SUM({1,2,3;4,5,6})');
        // Number, string, boolean and error elements are each encoded differently *and to different
        // widths*, so one misread width desynchronises everything after it.
        assert.equal(api.xlsbFormula('Calc', 'C33').formula, 'COUNTA({"a",TRUE;#N/A,5})');
      },
    },
    {
      name: 'an array formula states its formula on the cell that owns it',
      expect(api: CorpusApi, assert: Assert) {
        // An array formula's cells do not carry the formula: they point at the group's top-left, and
        // the record naming the group comes *after* them in the stream.
        const cell = api.xlsbFormula('Calc', 'F1');
        assert.equal(cell.formula, 'SUM(A1:A5*2)');
        assert.equal(cell.result, 30);
        // And it reads as the array formula it is, over the range the group record names.
        assert.equal(cell.arrayRef, 'F1');
      },
    },
    {
      name: 'a cached result survives beside the formula that produced it',
      expect(api: CorpusApi, assert: Assert) {
        assert.equal(api.xlsbFormula('Calc', 'C14').result, 15);
        assert.equal(api.xlsbFormula('Calc', 'C16').result, 'pos');
        assert.equal(api.xlsbFormula('Calc', 'C9').result, true);
        assert.deepEqual(api.xlsbFormula('Calc', 'C35').result, {error: '#REF!'});
      },
    },
    {
      name: 'a filled-down formula reads its own translated text on every cell',
      expect(api: CorpusApi, assert: Assert) {
        // Where the XML form stores one master and marks the rest as clones, Excel's binary form
        // writes each cell's formula out in full, so a clone reads back with the same text either
        // way, and only the pointer back to the master is absent. The grouping is a storage
        // optimisation, not a fact about the sheet.
        assert.deepEqual(api.xlsbFilledFormulaColumn(), [
          {address: 'D1', formula: 'A1*2', shared: false},
          {address: 'D2', formula: 'A2*2', shared: false},
          {address: 'D3', formula: 'A3*2', shared: false},
          {address: 'D4', formula: 'A4*2', shared: false},
          {address: 'D5', formula: 'A5*2', shared: false},
        ]);
      },
    },
    {
      name: 'defined names read back with their targets and their scope',
      expect(api: CorpusApi, assert: Assert) {
        // A name's target is a token stream too, so names are unreadable until formulas are, and a
        // name may hold a constant rather than a reference.
        assert.deepEqual(api.xlsbDefinedNames(), [
          {name: 'Factor', refersTo: '2'},
          {name: 'Local', scope: 'Calc', refersTo: 'Calc!$A$1'},
          {name: 'Rate', refersTo: 'Calc!$E$1'},
        ]);
      },
    },
    {
      name: 'a formula referring to a defined name cites it by name, not by index',
      expect(api: CorpusApi, assert: Assert) {
        // The token holds a position in the file's name list, a list that also contains the hidden
        // placeholders Excel registers for post-2007 functions. Filtering those out of the model
        // while still counting them for lookup is the whole difference between `Rate*A1` and a
        // formula naming the wrong thing.
        assert.equal(api.xlsbFormula('Calc', 'C25').formula, 'Rate*A1');
        assert.equal(api.xlsbFormula('Calc', 'C26').formula, 'Factor*A2');
        assert.equal(api.xlsbFormula('Calc', 'C27').formula, 'Local+1');
      },
    },
  ],
} satisfies Case;
