// Cluster: conditional-formatting
//
// Real-world scenario: a data bar made in Excel 2010 or later keeps most of its look in the worksheet's
// 2009 extension, an `<x14:dataBar>` the classic `<dataBar>` links to by id: the automatic anchors the
// ribbon gives a new bar, a border, the direction the bar grows, where the axis sits, and the fill and
// border colours of a negative bar. The classic element carries only the anchors, the bar colour, the
// lengths and whether the value shows. Reading and saving the workbook must keep every one of those, or
// each bar in a dashboard falls back to Excel 2007's look on the first save: bordered bars lose their
// border, negative bars turn the positive colour and grow the wrong way from a vanished axis.
//
// The fixture is a workbook Excel 16.0 saved (see `author.ps1` beside it) with five bars: automatic
// anchors; a solid black border, right-to-left, a solid fill, no axis and hidden values; red negative
// bars with a blue border about a green axis at the middle; a green bar from 20% to 80% of the cell,
// which Excel keeps classic; and a number and a percentile for anchors.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'data-bar-facets/excel-saved.xlsx';

// A bar stating nothing beyond its anchors and colour; each entry below says what it adds.
const PLAIN = {
  color: 'FF638EC6',
  minLength: null,
  maxLength: null,
  showValue: null,
  gradient: null,
  border: null,
  borderColor: null,
  direction: null,
  negativeFillColor: null,
  negativeBorderColor: null,
  negativeBarColorSameAsPositive: null,
  negativeBarBorderColorSameAsPositive: null,
  axisPosition: null,
  axisColor: null,
};

// What every bar Excel saved with an extension states there, whatever else it sets.
const EXCEL_2010 = {minLength: 0, maxLength: 100};

export default {
  id: 'data-bar-facets-round-trip',
  cluster: 'conditional-formatting',
  provenance: {source: 'excel-desktop', excel: '16.0 build 20326'},
  description:
    'A data bar Excel saved reads with the anchors, lengths, border, direction, axis and negative-bar ' +
    'colours it keeps in the worksheet extension, and a save writes every one back as Excel saved it.',
  behavior: [
    {
      name: 'each bar reads with the look Excel gave it, the facets only the extension holds included',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(api.dataBarsAsRead(FIXTURE), [
          {
            ...PLAIN,
            ...EXCEL_2010,
            ref: 'A1:A6',
            anchors: ['autoMin', 'autoMax'],
            negativeBarColorSameAsPositive: true,
            axisPosition: 'none',
          },
          {
            ...PLAIN,
            ...EXCEL_2010,
            ref: 'B1:B6',
            anchors: ['min', 'max'],
            showValue: false,
            gradient: false,
            border: true,
            borderColor: 'FF000000',
            direction: 'rightToLeft',
            negativeBarColorSameAsPositive: true,
            axisPosition: 'none',
          },
          {
            ...PLAIN,
            ...EXCEL_2010,
            ref: 'C1:C6',
            anchors: ['min', 'max'],
            border: true,
            borderColor: 'FF000000',
            negativeFillColor: 'FFFF0000',
            negativeBorderColor: 'FF0000FF',
            negativeBarBorderColorSameAsPositive: false,
            axisPosition: 'middle',
            axisColor: 'FF00FF00',
          },
          {
            ...PLAIN,
            ref: 'D1:D6',
            anchors: ['min', 'max'],
            color: 'FF92D050',
            minLength: 20,
            maxLength: 80,
          },
          {
            ...PLAIN,
            ...EXCEL_2010,
            ref: 'E1:E6',
            anchors: ['num:-2', 'percentile:90'],
            negativeBarColorSameAsPositive: true,
            axisPosition: 'none',
          },
        ]);
      },
    },
    {
      name: 'a save writes each bar back in both forms as Excel saved it, but for the id linking the two',
      expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = api.conditionalFormatRulesAsStored(FIXTURE);
        assert.strictEqual(source.extension.length, 4, 'four of the five bars have an extension');
        assert.deepEqual(rewritten, source);
      },
    },
  ],
} satisfies Case;
