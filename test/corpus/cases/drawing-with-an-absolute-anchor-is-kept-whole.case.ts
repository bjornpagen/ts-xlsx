// Cluster: images
//
// Real-world scenario: a dashboard carries a logo anchored between cells and a watermark anchored to the
// page (`<xdr:absoluteAnchor>`), both in the sheet's one drawing part. The reader modeled the cell-anchored
// picture and did not recognise the absolute anchor as content it cannot write back, so the drawing
// counted as fully modeled and a save rewrote it from the logo alone: the watermark was gone.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'drawing-with-an-absolute-anchor-is-kept-whole',
  cluster: 'images',
  provenance: {source: 'audit'},
  description:
    'A drawing holding an absolutely anchored picture beside a cell-anchored one models no picture ' +
    'from the part and writes the whole drawing back as it was read, both anchors included.',

  behavior: [
    {
      name: 'a drawing holding an absolutely anchored picture models no picture from it',
      expect(api: CorpusApi, assert: Assert) {
        assert.equal(api.absoluteAnchorDrawingReport().imagesModeled, 0);
      },
    },
    {
      name: 'a save writes the drawing back whole, the absolute anchor included',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.absoluteAnchorDrawingReport();
        assert.deepEqual(report.anchorKinds, [['twoCellAnchor', 'absoluteAnchor']]);
        assert.equal(report.drawingKept, true);
      },
    },
  ],
} satisfies Case;
