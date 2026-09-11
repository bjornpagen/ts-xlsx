import type {Assert, Case, CorpusApi} from '../case.ts';

// A binary `.xlsb` cites a built-in function by its index in a function table. The table [MS-XLS]
// publishes ends at `RTD`, but Excel 2007 took the Analysis ToolPak and CUBE functions into its engine
// and added `IFERROR` and the `*IFS` family, and an `.xlsb` cites those 105 by indices past it. They
// are some of the most used functions there are (`SUMIFS`, `IFERROR`, `EOMONTH`, `NETWORKDAYS`), and a
// reader that knows only the [MS-XLS] table decodes none of them: each such cell reads as its cached
// value with its formula gone.
//
// The fixtures are a pair Excel Desktop 16.0 (build 20326) saved from one in-memory workbook, one
// formula per function plus a few calls with an optional argument; see `author.ps1` beside them. The
// XML twin states each formula as text, so it is the oracle for what the token stream decodes to.
export default {
  id: 'xlsb-analysis-toolpak-and-2007-functions-decode',
  cluster: 'xlsx-io',
  description:
    'A binary .xlsb formula calling a function Excel 2007 added to its function table (the Analysis ' +
    'ToolPak and CUBE functions, IFERROR, COUNTIFS, SUMIFS, AVERAGEIF, AVERAGEIFS) reads back as the ' +
    'same formula text its .xlsx twin states.',
  provenance: {source: 'excel-desktop', excel: '16.0 build 20326'},
  behavior: [
    {
      name: 'every call to a function past RTD reads back as the text its XML twin states',
      expect(api: CorpusApi, assert: Assert) {
        const result = api.xlsbFormulaTextMatchesXlsxTwin(
          'xlsb-analysis-toolpak-and-2007-functions',
        );
        assert.deepEqual(result.differences, []);
        // Every function once, and the calls repeated with an optional argument.
        assert.equal(result.compared, 109);
      },
    },
  ],
} satisfies Case;
