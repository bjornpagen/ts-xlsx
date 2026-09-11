// Cluster: tables
//
// Real-world scenario: a report generator builds a pivot from configuration, so the field names for
// each role come from data rather than from a person looking at the result. Three slips come out of
// that. The same field is listed as both a row and a column ("Region by Region"), or twice in the
// row list, or the source sheet's header row names two columns the same, in any letter case ("Name"
// and "name"), as exported data often does. A pivot table can place a field on one axis, once, and a
// cache names each field once: a package that says otherwise opens with Excel's "We found a problem
// with some content" repair prompt, while the same pivot with each field in one place opens clean.
// These are authoring mistakes, so they are refused when the pivot is added, not written.
//
// One shape looks like a fourth slip and is not: counting a field that is also a row ("Count of
// Name" by Name) is an ordinary pivot. Excel writes that field with its axis *and* flagged as a data
// field; leaving the flag off an axis field is the same repair prompt, and adding it opens clean.

import type {Assert, Case, CorpusApi} from '../case.ts';

const HEADERS = ['Name', 'Region', 'Amount'];

export default {
  id: 'pivot-field-plays-one-axis-role',
  provenance: {source: 'excel-desktop-verification'},
  cluster: 'tables',
  description:
    'A pivot places each source field on at most one axis, at most once, over a header row whose ' +
    'names are unique regardless of case: a field named as both row and column, a field named twice ' +
    'in one axis, and a duplicated header are each refused as authoring errors. A value field that ' +
    'is also an axis field is written as both, which is what Excel needs to open it clean.',

  behavior: [
    {
      name: 'a field named as both a row and a column field is refused when the pivot is added',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.pivotFieldWiring({
          headers: HEADERS,
          rows: ['Name'],
          columns: ['Name'],
          values: ['Amount'],
        });
        assert.ok(report.refusal, JSON.stringify(report));
        assert.strictEqual(report.refusal.code, 'authoring');
        assert.match(report.refusal.message, /"Name"/);
      },
    },
    {
      name: 'a field named twice in the row fields is refused when the pivot is added',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.pivotFieldWiring({
          headers: HEADERS,
          rows: ['Name', 'Name'],
          columns: ['Region'],
          values: ['Amount'],
        });
        assert.ok(report.refusal, JSON.stringify(report));
        assert.strictEqual(report.refusal.code, 'authoring');
        assert.match(report.refusal.message, /"Name"/);
      },
    },
    {
      name: 'a source header row naming two columns the same, ignoring case, is refused',
      expect(api: CorpusApi, assert: Assert) {
        for (const duplicate of ['Name', 'NAME']) {
          const report = api.pivotFieldWiring({
            headers: [...HEADERS, duplicate],
            rows: ['Name'],
            columns: ['Region'],
            values: ['Amount'],
          });
          assert.ok(report.refusal, JSON.stringify(report));
          assert.strictEqual(report.refusal.code, 'authoring');
          assert.match(report.refusal.message, new RegExp(`"${duplicate}"`));
        }
      },
    },
    {
      // Header names are unique ignoring case, so a name in any case binds exactly one column. Matched
      // exactly, `name` over a `Name` header was "not a column header" of a sheet that shows one.
      name: 'a field is named by its header in any case',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.pivotFieldWiring({
          headers: HEADERS,
          rows: ['name'],
          columns: ['REGION'],
          values: ['amount'],
        });
        assert.strictEqual(report.refusal, null, JSON.stringify(report));
        assert.deepStrictEqual(report.rowFields, [0]);
        assert.deepStrictEqual(report.columnFields, [1]);
        assert.deepStrictEqual(report.dataFields, [2]);
      },
    },
    {
      name: 'one field named twice in two cases is still named twice',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.pivotFieldWiring({
          headers: HEADERS,
          rows: ['Name', 'NAME'],
          columns: ['Region'],
          values: ['Amount'],
        });
        assert.ok(report.refusal, JSON.stringify(report));
        assert.strictEqual(report.refusal.code, 'authoring');
      },
    },
    {
      name: 'a value field that is also a row field is written on its axis and flagged as data',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.pivotFieldWiring({
          headers: HEADERS,
          rows: ['Name'],
          columns: ['Region'],
          values: ['Name'],
        });
        assert.strictEqual(report.refusal, null, JSON.stringify(report));
        assert.deepStrictEqual(report.rowFields, [0]);
        assert.deepStrictEqual(report.columnFields, [1]);
        assert.deepStrictEqual(report.dataFields, [0]);
        assert.deepStrictEqual(report.pivotFields?.[0], {axis: 'axisRow', dataField: true});
      },
    },
    {
      name: 'a pivot with every field in one place still places each where it was named',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.pivotFieldWiring({
          headers: HEADERS,
          rows: ['Name'],
          columns: ['Region'],
          values: ['Amount'],
        });
        assert.strictEqual(report.refusal, null, JSON.stringify(report));
        assert.deepStrictEqual(report.pivotFields, [
          {axis: 'axisRow', dataField: false},
          {axis: 'axisCol', dataField: false},
          {axis: null, dataField: true},
        ]);
      },
    },
  ],
} satisfies Case;
