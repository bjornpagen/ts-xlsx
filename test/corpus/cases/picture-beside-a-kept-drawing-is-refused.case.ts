// Cluster: images
//
// Real-world scenario: a report template carries a chart or a callout shape, and a fill-and-save job
// adds a logo to the same sheet. The library keeps a drawing holding content it does not model byte
// for byte, but a worksheet references only one drawing, so writing the logo as a new drawing left the
// chart or shape unreferenced: the saved file silently lost it. Until pictures can be merged into a
// kept drawing, the add must be refused and the kept content must survive the attempt.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'picture-beside-a-kept-drawing-is-refused',
  cluster: 'images',
  provenance: {source: 'audit'},
  description:
    'A sheet whose drawing holds a shape the library does not model keeps that drawing across a ' +
    'save, refuses a picture added beside it with an error naming the sheet, and still writes the ' +
    'shape after the refused attempt.',

  behavior: [
    {
      name: 'a drawing holding a shape is kept whole on read, with no picture modelled from it',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.pictureBesideKeptDrawingReport();
        assert.equal(report.keptShape, true);
        assert.equal(report.imagesModeled, 0);
      },
    },
    {
      name: 'anchoring a picture on that sheet is refused with an error naming the sheet',
      expect(api: CorpusApi, assert: Assert) {
        const {addError} = api.pictureBesideKeptDrawingReport();
        assert.ok(addError !== null, 'the add was refused');
        assert.match(addError, /"S"/);
      },
    },
    {
      name: 'the refused attempt leaves the shape in the drawing the written sheet references',
      expect(api: CorpusApi, assert: Assert) {
        assert.equal(api.pictureBesideKeptDrawingReport().shapeAfterAttempt, true);
      },
    },
  ],
} satisfies Case;
