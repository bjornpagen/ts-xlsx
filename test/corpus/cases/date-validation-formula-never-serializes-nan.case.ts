// Cluster: data-validation
//
// Real-world scenario: a user adds a date-type data validation to a range, expecting the validation
// bound to be written as a real date. When the operand is a genuine Date, the worksheet XML must
// carry a valid date serial in the formula. When the operand is not a coercible date, the writer must
// never emit the literal token "NaN" into the formula element: Excel then treats the bound as broken
// and the validation silently fails. Dropping the bound instead is no better, since it silently
// changes the rule, so the write refuses the number the format cannot spell.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'date-validation-formula-never-serializes-nan',
  provenance: {source: 'upstream-issue'},
  cluster: 'data-validation',
  description:
    'A date-type data validation writes a real date serial for a genuine Date operand, and never ' +
    'emits the literal "NaN" into the validation formula for a non-coercible operand (which would ' +
    'silently break the bound in Excel): the write refuses it instead.',

  behavior: [
    {
      name: 'a genuine Date operand writes a valid date serial, not NaN',
      async expect(api: CorpusApi, assert: Assert) {
        const {formula1, hasNaN} = await api.authorDateValidation('2020-01-01T00:00:00.000Z');
        assert.strictEqual(hasNaN, false, 'no NaN token is emitted for a real date');
        assert.ok(
          formula1 && /^\d+(\.\d+)?$/.test(formula1),
          `the bound is a numeric serial; got ${JSON.stringify(formula1)}`,
        );
      },
    },
    {
      name: 'a non-coercible operand never emits the literal NaN into the formula: the write is refused',
      async expect(api: CorpusApi, assert: Assert) {
        const {formula1, hasNaN, refused} = await api.authorDateValidation('invalid');
        assert.strictEqual(
          hasNaN,
          false,
          `a non-coercible operand must not serialize "NaN"; got ${JSON.stringify(formula1)}`,
        );
        assert.strictEqual(refused, true, 'the unwritable bound is refused, not silently dropped');
      },
    },
  ],
} satisfies Case;
