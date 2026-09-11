import type {Assert, Case, CorpusApi} from '../case.ts';
import type {Untyped} from '../untyped.ts';

// `.xlsb` (BIFF12) is the binary serialisation of the same spreadsheet model `.xlsx` spells in XML:
// the same OPC/ZIP package and relationship graph, with the workbook, worksheet, shared-string and
// style parts stored as binary record streams. A reader that supports both must therefore produce
// *one* model, not two similar ones: a caller converting between the forms, or simply handed a file
// it did not choose the format of, must not have to care which it got.
//
// The fixtures are a pair Excel itself saved from one in-memory workbook, `source.xlsb` and
// `source.xlsx`, which makes the XML twin an independent oracle rather than something this library
// produced. See `author.ps1` beside them, and the spec
// `docs/knowledge/specs/xlsb-binary-format-output.md`.
export default {
  id: 'xlsb-binary-workbook-reads-like-its-xlsx-twin',
  cluster: 'xlsx-io',
  description:
    'A binary .xlsb workbook reads into the same model its .xlsx twin does: values, number ' +
    'formats, fonts, fills, borders, alignment, cell and sheet protection, sheet order and ' +
    'visibility, row and column geometry, and merges, and a malformed binary part fails with a ' +
    'typed parse error.',
  provenance: {source: 'upstream-issue'},
  behavior: [
    {
      name: 'the binary and XML readings of one workbook produce the same model',
      expect(api: CorpusApi, assert: Assert) {
        const result = api.xlsbModelMatchesXlsxTwin();
        // The message carries the first differing field, so a regression names itself.
        assert.equal(result.firstDifference, null);
        assert.equal(result.identical, true);
      },
    },
    {
      // Excel writes a password-protected sheet's credential in a second record beside the flags, and
      // stores each flag as "allowed", the inverse of the XML attribute. A reading that got either wrong
      // in both codecs at once would still pass the comparison above, so the values are stated too.
      name: 'sheet protection, its allowed operations, and its password credential survive',
      expect(api: CorpusApi, assert: Assert) {
        const protection = api.xlsbSheetProtection();
        assert.equal(protection.Values, null, 'an unprotected sheet reads as unprotected');
        assert.deepEqual(protection.Quiet, {
          flags: {
            objects: false,
            scenarios: false,
            selectLockedCells: false,
            selectUnlockedCells: false,
          },
        });
        const grid = protection.Grid;
        assert.deepEqual(grid?.flags, {
          scenarios: false,
          formatCells: true,
          sort: true,
          autoFilter: true,
        });
        assert.equal(grid?.credential?.algorithmName, 'SHA-512');
        assert.equal(grid?.credential?.spinCount, 100000);
        assert.equal(grid?.credential?.saltValue, 'uq7o93W/PeaNtWE6aZb8BA==');
      },
    },
    {
      name: 'sheet names, tab order, and hidden state survive the binary workbook part',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(api.xlsbSheets(), [
          {name: 'Quiet', state: 'hidden'},
          {name: 'Grid', state: 'visible'},
          {name: 'Values', state: 'visible'},
        ]);
      },
    },
    {
      name: 'the packed RK numeric encoding decodes to the number Excel stored',
      expect(api: CorpusApi, assert: Assert) {
        // Three encodings share one record type: a small integer, a two-decimal value scaled by 100,
        // and a negative integer. A reader that gets the flag bits wrong gets all three wrong.
        assert.equal(api.xlsbCell('Values', 'B2').value, 10);
        assert.equal(api.xlsbCell('Values', 'B3').value, 1.23);
        assert.equal(api.xlsbCell('Values', 'B5').value, -42);
        // Too wide for the packed form, so Excel stored a full double instead.
        assert.equal(api.xlsbCell('Values', 'B4').value, 1234.5678);
      },
    },
    {
      name: 'boolean, error, and Unicode string cells decode by their record type',
      expect(api: CorpusApi, assert: Assert) {
        assert.equal(api.xlsbCell('Values', 'B6').value, true);
        assert.deepEqual(api.xlsbCell('Values', 'B7').value, {error: '#DIV/0!'});
        // The dash is the fixture's own cell text, not prose: this asserts the binary reader
        // hands back exactly what Excel wrote into B12.
        // charcheck-disable-next-line no-em-dash-in-source
        assert.equal(api.xlsbCell('Values', 'B12').value, 'naïve \u2014 日本語');
      },
    },
    {
      name: 'a number under a date format reads as a date, as it does from the XML twin',
      expect(api: CorpusApi, assert: Assert) {
        const cell = api.xlsbCell('Values', 'B8');
        assert.equal(cell.value, '2020-01-02T00:00:00.000Z');
        assert.equal(cell.numFmt, 'yyyy\\-mm\\-dd');
      },
    },
    {
      name: 'a formula cell surfaces the result Excel cached for it',
      expect(api: CorpusApi, assert: Assert) {
        // The formula's own token stream is not decoded yet, so what a formula cell carries is the
        // value, which is the same value the XML reader takes from `<v>`. Locking it here means the
        // gap cannot widen into a *wrong* value while the token decoder is still to come.
        assert.equal(api.xlsbCell('Values', 'B9').value, 1245.7978);
        assert.equal(api.xlsbCell('Values', 'B10').value, 'WIDGET');
        assert.equal(api.xlsbCell('Values', 'B11').value, true);
      },
    },
    {
      name: 'a cell resolves its number format, font, fill, border, alignment, and protection',
      expect(api: CorpusApi, assert: Assert) {
        const facets = api.xlsbCell('Values', 'D2');
        assert.deepEqual(facets.font, {
          name: 'Consolas',
          size: 11,
          family: 3,
          color: {argb: 'FFC00000'},
          italic: true,
          underline: true,
        });
        assert.deepEqual(facets.alignment, {
          horizontal: 'left',
          vertical: 'top',
          wrapText: true,
          indent: 2,
        });
        // `locked` defaults to true, so only an explicitly unlocked cell carries the facet.
        assert.deepEqual(facets.protection, {locked: false});
        assert.deepEqual(facets.border, {
          left: {style: 'thin', color: {indexed: 64}},
          bottom: {style: 'dashed', color: {indexed: 64}},
        });
        // A solid fill's visible colour is the pattern foreground, beside the automatic background.
        assert.deepEqual(api.xlsbCell('Values', 'A1').fill, {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {argb: 'FFCCDDEE'},
          bgColor: {indexed: 64},
        });
        assert.equal(api.xlsbCell('Values', 'B13').numFmt, '0.00%');
      },
    },
    {
      name: 'an unset hatch pattern carries no colours, matching what the XML twin states',
      expect(api: CorpusApi, assert: Assert) {
        // The binary form always states both colours, using the legacy palette's automatic sentinels
        // where XML simply omits them; a reader that carries the sentinels through invents colours
        // the file does not have.
        assert.deepEqual(api.xlsbCell('Values', 'D6').fill, {
          type: 'pattern',
          pattern: 'lightTrellis',
        });
      },
    },
    {
      name: 'a cell with only a format and no value reads back styled and empty',
      expect(api: CorpusApi, assert: Assert) {
        const blank = api.xlsbCell('Values', 'B14');
        assert.equal(blank.value, null);
        // `Fill` is discriminated (pattern vs gradient), so reaching for `.pattern` means claiming
        // which arm this is. Stating that claim is a stronger assertion than the unchecked read was.
        const fill = blank.fill;
        assert.ok(fill?.type === 'pattern', 'the format-only cell carries a pattern fill');
        assert.equal(fill.pattern, 'solid');
      },
    },
    {
      name: 'a cell inherits the format its row or column declares when it has none of its own',
      expect(api: CorpusApi, assert: Assert) {
        // C1 sits in a bold column; A8 sits in an italic row. Neither cell carries the format.
        assert.equal(api.xlsbCell('Grid', 'C1').font?.bold, true);
        assert.equal(api.xlsbCell('Grid', 'A8').font?.italic, true);
      },
    },
    {
      name: 'column width, hidden state, and outline level survive the binary column records',
      expect(api: CorpusApi, assert: Assert) {
        const grid = api.xlsbGrid('Grid');
        const byIndex = new Map<number, Untyped>(
          grid.columns!.map((entry) => [entry.index, entry.properties]),
        );
        assert.equal(byIndex.get(1)?.width, 24.6328125);
        assert.equal(byIndex.get(2)?.hidden, true);
        assert.equal(byIndex.get(3)?.outlineLevel, 1);
        assert.equal(byIndex.get(4)?.outlineLevel, 1);
      },
    },
    {
      name: 'row height, hidden state, and outline level survive the binary row headers',
      expect(api: CorpusApi, assert: Assert) {
        const grid = api.xlsbGrid('Grid');
        const byNumber = new Map<number, Untyped>(
          grid.rows!.map((entry) => [entry.number, entry.properties]),
        );
        assert.equal(byNumber.get(1)?.height, 30);
        assert.equal(byNumber.get(3)?.hidden, true);
        assert.equal(byNumber.get(5)?.outlineLevel, 1);
        // Every row header restates the sheet's default height; a row that merely restates it has no
        // height of its own and must not read back with one.
        assert.equal(byNumber.get(8)?.height, undefined);
      },
    },
    {
      name: 'a merged range survives the binary merge records',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(api.xlsbGrid('Values').merges, ['F2:G3']);
      },
    },
    {
      name: 'a malformed binary workbook part fails with a typed parse error, not a crash',
      expect(api: CorpusApi, assert: Assert) {
        // A record header declaring a payload far longer than the part holds: the lever a hostile
        // file pulls to make a naive reader over-allocate or read out of bounds.
        const result = api.xlsbMalformedBinaryWorkbook();
        assert.equal(result.threw, true);
        assert.equal(result.errorName, 'XlsbParseError');
      },
    },
  ],
} satisfies Case;
