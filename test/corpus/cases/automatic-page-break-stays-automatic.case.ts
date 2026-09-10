// Cluster: xlsx-io
//
// Real-world scenario: a workbook from a producer that records automatic page breaks (soft breaks,
// stored without `man`) beside the author's manual ones, some with a `min` extent. The reader dropped
// `min` and never recorded that a break was automatic, and the writer then wrote every break as manual
// and counted them all in `manualBreakCount`, so a single save hardened the soft breaks into
// author-set ones that stay put when the content above them changes.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'automatic-page-break-stays-automatic',
  cluster: 'xlsx-io',
  provenance: {source: 'audit'},
  description:
    'A row break a file declares automatic reads as automatic with its extent, and a save writes it ' +
    'back without `man`, counting only the manual break in `manualBreakCount`.',

  behavior: [
    {
      name: 'a read reports which break is manual and keeps the extent of the automatic one',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(api.pageBreakKindsReport().read, [
          {id: 3, max: 16383, man: true},
          {id: 7, min: 2, max: 16383, man: false},
        ]);
      },
    },
    {
      name: 'a save writes the automatic break without man and counts only the manual one',
      expect(api: CorpusApi, assert: Assert) {
        const {rewritten} = api.pageBreakKindsReport();
        assert.deepEqual(rewritten.counts, {count: '2', manualBreakCount: '1'});
        assert.deepEqual(rewritten.breaks, [
          {id: '3', min: null, max: '16383', man: '1'},
          {id: '7', min: '2', max: '16383', man: null},
        ]);
      },
    },
  ],
} satisfies Case;
