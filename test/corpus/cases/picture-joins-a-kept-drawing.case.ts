// Cluster: images
//
// Real-world scenario: a report template carries a chart and a callout shape, and a fill-and-save job
// adds a logo to the same sheet. The library keeps a drawing holding content it does not model byte for
// byte, and a worksheet references only one drawing, so the logo has to join the kept drawing: written
// as a drawing of its own, it would leave the chart and the shape unreferenced, and the saved file
// would lose them.
//
// The fixture is a workbook Excel 16.0 saved (see `author.ps1` beside it) whose sheet holds a clustered
// column chart and a rectangle in one drawing.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'picture-joins-a-kept-drawing/excel-saved.xlsx';

export default {
  id: 'picture-joins-a-kept-drawing',
  cluster: 'images',
  provenance: {source: 'excel-desktop', excel: '16.0 build 20326'},
  description:
    'A picture added to a sheet whose drawing holds a chart and a shape is written into that drawing, ' +
    'beside them: the saved sheet references one drawing holding all three, the picture reaches its ' +
    'bytes and the chart its part, and all three survive a second read and save.',
  behavior: [
    {
      name: 'the drawing is kept whole on read, and a picture can still be anchored on its sheet',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.pictureJoinsKeptDrawingReport(FIXTURE);
        assert.strictEqual(report.imagesModeledOnRead, 0);
        assert.strictEqual(report.addError, null);
      },
    },
    {
      name: 'the saved sheet references one drawing holding the chart, the shape and the picture',
      expect(api: CorpusApi, assert: Assert) {
        const {saved} = api.pictureJoinsKeptDrawingReport(FIXTURE);
        assert.deepEqual(saved, {
          drawingsReferenced: 1,
          anchors: ['chart', 'shape', 'picture'],
          pictureReachesMedia: true,
          chartReached: true,
        });
      },
    },
    {
      name: 'read back and saved again, the drawing still holds all three',
      expect(api: CorpusApi, assert: Assert) {
        const {resaved} = api.pictureJoinsKeptDrawingReport(FIXTURE);
        assert.deepEqual(resaved, {
          drawingsReferenced: 1,
          anchors: ['chart', 'shape', 'picture'],
          pictureReachesMedia: true,
          chartReached: true,
        });
      },
    },
  ],
} satisfies Case;
