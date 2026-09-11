// Cluster: security
//
// Real-world scenario: a `.xlsx` this library did not write names a sheet twice, names one nothing
// at all, puts a `/` in one, runs one past 31 characters, or calls a table `1 bad` or `T1`. Excel
// repairs every one of those on load. The reader used to hand each straight to the model's authoring
// guards and let the refusal out: an `AuthoringError`, which the taxonomy defines as "the calling code
// is wrong, never the input file", or -- worse for the table names -- a native `SyntaxError` or
// `RangeError`, which `catch (e) { if (e instanceof XlsxError) … }` does not see at all, so the
// caller cannot tell "your file is broken" from "my own code threw".
//
// The rule this locks: a name the model refuses costs that name or the feature carrying it, never
// the read. A sheet keeps its cells and its place in the order under a repaired name -- dropping it
// would take every `localSheetId` indexing past it with it -- while a table or a defined name, whose
// name is its identity, is dropped whole. And the model that comes back is one this library can
// still write.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'a-hostile-name-costs-that-name-not-the-read',
  provenance: {source: 'hostile-input-probe'},
  cluster: 'security',
  description:
    'A sheet, table or defined name a foreign file gives that the model would refuse is repaired ' +
    'or dropped at the reader boundary: the read succeeds, the rest of the workbook is intact, ' +
    'the result re-writes, and no AuthoringError or native RangeError/SyntaxError escapes.',

  behavior: [
    {
      name: 'no refused name aborts the read',
      expect(api: CorpusApi, assert: Assert) {
        for (const row of api.hostileNameReport()) {
          assert.equal(
            row.threw,
            false,
            `${row.mutation}: the read survives it (${row.errorName})`,
          );
          assert.equal(row.keptSiblingCell, 'keep', `${row.mutation}: A1 is still readable`);
        }
      },
    },
    {
      name: 'what the reader produced is something the writer can still serialise',
      expect(api: CorpusApi, assert: Assert) {
        for (const row of api.hostileNameReport()) {
          assert.equal(row.rewrote, true, `${row.mutation}: the model re-writes`);
        }
      },
    },
    {
      name: 'a sheet whose name the model refuses keeps its cells and its place in the order',
      expect(api: CorpusApi, assert: Assert) {
        const byMutation = new Map(api.hostileNameReport().map((row) => [row.mutation, row]));
        const sheetCases = [
          'duplicate sheet name',
          'empty sheet name',
          'sheet name with a forbidden character',
          'sheet name over the length limit',
          'sheet name edged with an apostrophe',
        ];
        for (const mutation of sheetCases) {
          const row = byMutation.get(mutation);
          assert.ok(row !== undefined, `${mutation}: reported`);
          assert.equal(row.sheetNames.length, 2, `${mutation}: no sheet vanished`);
          assert.equal(row.sheetNames[0], 'Alpha', `${mutation}: the good sheet is untouched`);
        }
      },
    },
    {
      name: 'each refused sheet name is repaired the way Excel repairs it',
      expect(api: CorpusApi, assert: Assert) {
        const named = (mutation: string) =>
          api.hostileNameReport().find((row) => row.mutation === mutation)?.sheetNames[1];
        assert.equal(named('duplicate sheet name'), 'Alpha (2)', 'a collision takes a suffix');
        assert.equal(named('empty sheet name'), 'Sheet1', 'an empty name gets a generated one');
        assert.equal(named('sheet name with a forbidden character'), 'ab', 'the character goes');
        assert.equal(
          named('sheet name over the length limit'),
          'B'.repeat(31),
          'an over-long name is cut to the 31-character limit',
        );
        assert.equal(
          named('sheet name edged with an apostrophe'),
          'Beta',
          'an edge apostrophe cannot be told from the quoting of a qualified reference',
        );
      },
    },
    {
      name: 'a table or defined name whose name is its identity is dropped whole, alone',
      expect(api: CorpusApi, assert: Assert) {
        const byMutation = new Map(api.hostileNameReport().map((row) => [row.mutation, row]));

        for (const mutation of [
          'table name that is not an identifier',
          'table name over the length limit',
          'table name that reads as a cell',
        ]) {
          const row = byMutation.get(mutation);
          assert.equal(row?.tables, 0, `${mutation}: the table is dropped`);
          assert.equal(row?.definedNames, 1, `${mutation}: the defined name beside it is not`);
          assert.deepEqual(row?.sheetNames, ['Alpha', 'Beta'], `${mutation}: both sheets remain`);
        }

        const emptyName = byMutation.get('defined name with no name at all');
        assert.equal(emptyName?.definedNames, 0, 'the nameless defined name is dropped');
        assert.equal(emptyName?.tables, 1, 'and the table beside it is not');
      },
    },
  ],
} satisfies Case;
