// Cluster: types
//
// Real-world scenario: a workbook saved from Excel as "Strict Open XML Spreadsheet" (ISO/IEC 29500
// Strict) holds two charts and a shape, is read, and is saved again. The library writes a Transitional
// package, and the parts it models come out Transitional because they are written from the model. The
// theme, the drawing and the charts are not modelled: they ride through as the bytes Strict wrote,
// with `purl.oclc.org` namespaces, Strict relationship types, and DrawingML percentages spelled `60%`
// where Transitional stores thousandths. In a package that says it is Transitional, the Open XML SDK
// cannot load such a theme at all, and refuses every `60%` it meets once the namespace is fixed.
//
// The fixture is Excel 16.0's own Strict save: a doughnut chart, a column chart and a filled, 40%
// transparent rectangle. Excel read the rewritten package back over COM with the same hole size, gap
// width, overlap, series formula and transparency as the original, and OpenXmlValidator reports on it
// only the chart extension findings it reports on Excel's own Transitional charts.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'strict-workbook-writes-as-transitional/charts-and-shape.xlsx';

export default {
  id: 'a-strict-workbook-writes-as-transitional',
  provenance: {source: 'excel-desktop-verification'},
  cluster: 'types',
  description:
    'A Strict workbook read and written again comes out wholly Transitional: its preserved theme, ' +
    'drawing and charts carry Transitional namespaces and relationship types, and their DrawingML ' +
    'percentages are the thousandths Transitional stores rather than the percent strings Strict writes.',

  behavior: [
    {
      name: 'no written part mentions a Strict namespace, and every relationship type is Transitional',
      expect(api: CorpusApi, assert: Assert) {
        const {strictParts, relationshipTypes} = api.strictSpellingsAfterWrite(FIXTURE);
        assert.deepEqual(strictParts, []);
        assert.deepEqual(
          relationshipTypes.filter(
            (type) => !type.startsWith('http://schemas.openxmlformats.org/'),
          ),
          [],
        );
        assert.ok(
          relationshipTypes.includes(
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
          ),
          'the drawing still reaches its charts',
        );
      },
    },
    {
      name: 'a DrawingML percentage is written in thousandths, and a chart percentage as it was',
      expect(api: CorpusApi, assert: Assert) {
        const {attributes} = api.strictSpellingsAfterWrite(FIXTURE, [
          'xl/drawings/drawing1.xml#a:alpha val',
          'xl/drawings/drawing1.xml#a:shade val',
          'xl/theme/theme1.xml#a:gs pos',
          'xl/charts/chart1.xml#c:holeSize val',
          'xl/charts/chart2.xml#c:gapWidth val',
        ]);
        assert.deepEqual(attributes['xl/drawings/drawing1.xml#a:alpha val'], ['60000']);
        assert.deepEqual(attributes['xl/drawings/drawing1.xml#a:shade val'], ['15000']);
        assert.ok(
          (attributes['xl/theme/theme1.xml#a:gs pos'] ?? []).every((pos) => /^\d+$/.test(pos)),
          `every gradient stop position is an integer: ${JSON.stringify(attributes['xl/theme/theme1.xml#a:gs pos'])}`,
        );
        assert.deepEqual(attributes['xl/charts/chart1.xml#c:holeSize val'], ['50']);
        assert.deepEqual(attributes['xl/charts/chart2.xml#c:gapWidth val'], ['150']);
      },
    },
  ],
} satisfies Case;
