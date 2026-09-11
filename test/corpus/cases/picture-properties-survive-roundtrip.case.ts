// Cluster: images
//
// Real-world scenario: a report template's logo has alternative text for screen readers and a title, is
// cropped to the mark, and links to the company site; a navigation picture beside it jumps to a cell in
// the workbook. A library fills the template in and saves it, and every one of those was gone: the
// drawing was rewritten from the anchors alone, so the logo lost its accessibility text, uncropped
// itself, and stopped being a link.
//
// The fixture is Excel's own file, saved from those two pictures (see `author.ps1` beside it): alternative
// text and a title on `cNvPr`, a crop as `a:srcRect`, and each link as an `a:hlinkClick` naming a hyperlink
// relationship, external for the URL and `#Sheet1!C3` for the place in the workbook. Read back over COM,
// Excel reports the same text, crop and link addresses for this library's rewrite as for its own file, and
// opens the rewrite clean (Excel 16.0 build 20326).

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'picture-properties-survive-roundtrip/pictures.xlsx';

const LOGO = {
  description: 'Company logo',
  title: 'Logo',
  // CropLeft 3pt, CropTop 1.5pt and CropRight 6pt of a 30pt by 15pt picture, as Excel stored them.
  crop: {left: 0.09999, top: 0.09999, right: 0.19997},
  hyperlink: {target: 'https://example.com/about', tooltip: 'About us'},
};

const JUMP = {
  description: null,
  title: null,
  crop: null,
  hyperlink: {target: '#Sheet1!C3', tooltip: 'Jump'},
};

export default {
  id: 'picture-properties-survive-roundtrip',
  provenance: {source: 'excel-desktop-verification'},
  cluster: 'images',
  description:
    "A picture's alternative text, title, crop and link, to a URL or to a place in the workbook, read " +
    'from an Excel-authored drawing and survive a write and a second read.',

  behavior: [
    {
      name: 'the alternative text, title, crop and link Excel wrote are read',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepStrictEqual(api.readFixturePictureProperties(FIXTURE).read, [LOGO, JUMP]);
      },
    },
    {
      name: 'they survive a write and a second read unchanged',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepStrictEqual(api.readFixturePictureProperties(FIXTURE).rewritten, [LOGO, JUMP]);
      },
    },
  ],
} satisfies Case;
