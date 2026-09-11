// Cluster: types
//
// Real-world scenario: a generator copies computed results into a workbook, errors included, so a
// cell's value is `#N/A`, `#DIV/0!`, or one of the errors newer Excel shows: `#SPILL!` from a dynamic
// array with no room, `#CALC!` from a FILTER that matched nothing, `#BUSY!` while a function is still
// computing. A typed cell stores the error as its `<v>` text, and the question is which spellings Excel
// reads there. The answer came from Excel Desktop (build 20326), recorded in
// `test/corpus/fixtures/excel-oracle/error-literals.json`: `#GETTING_DATA` and `#BUSY!` open clean and
// read back as themselves, like the classic seven, while `#SPILL!`, `#CALC!` and the other newer errors
// make Excel offer to repair the package. Excel saves those as `#VALUE!` with a rich value beside the
// cell naming the real error, which is why the literal is not a spelling it accepts.
//
// So an error the format can hold literally is written and reads back as itself, and one it cannot is
// refused where the workbook is written rather than emitted as a package Excel repairs.

import type {Assert, Case, CorpusApi} from '../case.ts';

const LITERALS = ['#N/A', '#DIV/0!', '#GETTING_DATA', '#BUSY!'];
const RICH_VALUE_ERRORS = ['#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#UNKNOWN!'];

export default {
  id: 'an-error-value-is-written-only-as-a-literal-excel-reads',
  provenance: {
    source: 'excel-desktop-verification',
    ref: 'test/corpus/fixtures/excel-oracle/error-literals.json',
  },
  cluster: 'types',
  description:
    'An error value whose spelling Excel reads from a typed cell is written under t="e" and reads ' +
    'back as itself; one Excel stores only as #VALUE! beside a rich value, such as #SPILL! or #CALC!, ' +
    'is refused at write rather than written as a literal Excel repairs.',

  behavior: [
    {
      name: 'every error Excel reads literally is written as its literal and reads back as itself',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.errorLiteralReport(LITERALS);
        for (const code of LITERALS) {
          assert.strictEqual(report[code]?.refused, null, `${code} is written`);
          assert.strictEqual(report[code]?.written, code, `${code} is the cell's <v>`);
          assert.deepStrictEqual(report[code]?.readBack, {error: code}, `${code} reads back`);
        }
      },
    },
    {
      name: 'an error Excel stores only beside a rich value is refused rather than written literally',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.errorLiteralReport(RICH_VALUE_ERRORS);
        for (const code of RICH_VALUE_ERRORS) {
          assert.strictEqual(report[code]?.written, null, `${code} reaches no <v>`);
          assert.match(report[code]?.refused ?? '', /error/, `${code} is refused`);
        }
      },
    },
  ],
} satisfies Case;
