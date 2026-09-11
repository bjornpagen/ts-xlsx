// Cluster: pivot
//
// Real-world scenario: a dashboard workbook saved by Excel charts a data sheet and summarises it in a
// pivot, and the code refreshing it inserts rows into the data before saving. Neither the chart nor
// the pivot cache is modelled, so both ride through a read and a save as the bytes Excel wrote, and
// both name the data range: left as written, the chart stops plotting the rows that moved and the
// pivot refreshes over rows that are no longer the data, in a file that opens without complaint.
//
// Excel moves both on the same edit. The fixture is Excel's own save of the workbook before the edits,
// and every expectation is what Excel stored after them, recorded in
// `test/corpus/fixtures/excel-oracle/references-outside-cells-through-a-splice.json`. Excel also
// re-derives a series whose name or categories a delete takes, which nothing here asserts.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'references-outside-cells-through-a-splice/before.xlsx';

export default {
  id: 'a-loaded-chart-and-pivot-cache-move-with-a-splice',
  provenance: {
    source: 'excel-desktop-verification',
    ref: 'test/corpus/fixtures/excel-oracle/references-outside-cells-through-a-splice.json',
  },
  cluster: 'pivot',
  description:
    "A row or column splice of a loaded workbook moves the references in a preserved chart's series " +
    "and in a preserved pivot cache's worksheet source, as Excel does, and the loaded pivot's model " +
    'view says what is written. A splice of another sheet moves neither, and a delete that takes the ' +
    "whole pivot source leaves the cache's source as written.",

  behavior: [
    {
      name: 'an insert, a delete and a column insert on the data sheet move the series and the cache source',
      expect(api: CorpusApi, assert: Assert) {
        const [otherSheet, insertRow, deleteRow, insertColumn] =
          api.preservedReferencesThroughSplices(FIXTURE, [
            {sheet: 'S1', op: 'insert', axis: 'row', start: 5, count: 2},
            {sheet: 'S3', op: 'insert', axis: 'row', start: 3, count: 1},
            {sheet: 'S3', op: 'delete', axis: 'row', start: 4, count: 1},
            {sheet: 'S3', op: 'insert', axis: 'column', start: 2, count: 1},
          ]);
        assert.deepEqual(otherSheet, {
          cacheSource: {ref: 'A1:B5', sheet: 'S3'},
          loadedSource: {ref: 'A1:B5', sheet: 'S3'},
          chartReferences: ["'S3'!$B$1", "'S3'!$A$2:$A$5", "'S3'!$B$2:$B$5"],
        });
        assert.deepEqual(insertRow, {
          cacheSource: {ref: 'A1:B6', sheet: 'S3'},
          loadedSource: {ref: 'A1:B6', sheet: 'S3'},
          chartReferences: ["'S3'!$B$1", "'S3'!$A$2:$A$6", "'S3'!$B$2:$B$6"],
        });
        assert.deepEqual(deleteRow, {
          cacheSource: {ref: 'A1:B5', sheet: 'S3'},
          loadedSource: {ref: 'A1:B5', sheet: 'S3'},
          chartReferences: ["'S3'!$B$1", "'S3'!$A$2:$A$5", "'S3'!$B$2:$B$5"],
        });
        assert.deepEqual(insertColumn, {
          cacheSource: {ref: 'A1:C5', sheet: 'S3'},
          loadedSource: {ref: 'A1:C5', sheet: 'S3'},
          chartReferences: ["'S3'!$C$1", "'S3'!$A$2:$A$5", "'S3'!$C$2:$C$5"],
        });
      },
    },
    {
      name: "a delete of the pivot source's every row leaves the cache source as written",
      expect(api: CorpusApi, assert: Assert) {
        const [deleted] = api.preservedReferencesThroughSplices(FIXTURE, [
          {sheet: 'S3', op: 'delete', axis: 'row', start: 1, count: 5},
        ]);
        assert.deepEqual(deleted?.cacheSource, {ref: 'A1:B5', sheet: 'S3'});
        assert.deepEqual(deleted?.loadedSource, {ref: 'A1:B5', sheet: 'S3'});
      },
    },
  ],
} satisfies Case;
