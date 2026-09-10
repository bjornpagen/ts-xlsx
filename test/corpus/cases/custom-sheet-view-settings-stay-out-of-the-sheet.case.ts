// Cluster: xlsx-io
//
// Real-world scenario: Excel's View > Custom Views saves a named snapshot of a sheet's pane, print
// settings and filter, and writes each one into the worksheet part as a `<customSheetView>`. Its
// children are the very elements the sheet itself uses (`<pane>`, `<rowBreaks>`, `<pageMargins>`,
// `<headerFooter>`, `<autoFilter>` and more), and the block sits after the sheet's own `<sheetViews>`
// and `<autoFilter>`. A reader that dispatches on element name alone applies every saved view to the
// sheet: the saved pane replaces the real one, the saved filter replaces the real filter, the saved
// breaks join the sheet's own, and margins, page setup and a header appear where the sheet had none.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'custom-sheet-view-settings-stay-out-of-the-sheet',
  provenance: {source: 'hand-authored'},
  cluster: 'xlsx-io',
  description:
    'A worksheet carrying a saved custom view reads with its own pane, page breaks, margins, print ' +
    'options, page setup, header/footer and autofilter, exactly as it reads without the view: ' +
    'nothing the saved view declares is applied to the sheet.',

  behavior: [
    {
      name: 'the saved view changes nothing about how the sheet reads',
      expect(api: CorpusApi, assert: Assert) {
        const {withView, withoutView} = api.customSheetViewReport();
        assert.deepEqual(withView, withoutView);
      },
    },
    {
      name: "the sheet keeps its own filter and breaks rather than the view's",
      expect(api: CorpusApi, assert: Assert) {
        // Pinned outright as well, so two readings agreeing on the wrong answer cannot pass.
        const {withView} = api.customSheetViewReport();
        assert.equal(withView.autoFilterRef, 'A1:A2');
        assert.deepEqual(withView.rowBreaks, [3]);
        assert.equal(
          withView.viewState === 'frozen',
          false,
          'the saved frozen pane is not applied',
        );
        assert.equal(
          withView.headerFooter.oddHeader,
          undefined,
          "the saved header is not the sheet's",
        );
      },
    },
    {
      name: 'a workbook read with a saved view writes back out',
      expect(api: CorpusApi, assert: Assert) {
        assert.equal(api.customSheetViewReport().rewriteOk, true);
      },
    },
  ],
} satisfies Case;
