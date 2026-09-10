// Rows, columns, merges and the sheet geometry around them: insertion, splicing, outline
// levels, freeze panes, print areas and page breaks, and the print settings that ride alongside
// them: page margins, and the header/footer definition text.

import type {RowInput} from '../../../../src/core/worksheet.ts';
import {messageOf} from '../../thrown.ts';
import type {Untyped} from '../../untyped.ts';
import {type PartMap, partMapOf, partOf, patchedPackage, roundtrip} from './package-facts.ts';
import {
  decodeRange,
  encodeAddress,
  fixtureBytes,
  MAX_COLUMN,
  MAX_ROW,
  readFixture,
  readWorkbookStream,
  readXlsx,
  Workbook,
  type WorkbookInstance,
  writeXlsx,
} from './runtime.ts';
import {buildFrom, normalizeStreamValue, ONE_PX_PNG, printAreaRefersTo} from './spec-model.ts';
import {attrsOf, reloadPatched} from './xml-probes.ts';

export const grid = {
  cellColRowTypes(ref = 'B3') {
    const sheet = new Workbook().addWorksheet('S');
    const cell = sheet.getCell(ref);
    cell.value = 'x';
    return {col: cell.col, row: cell.row, colType: typeof cell.col, rowType: typeof cell.row};
  },

  // Author a sheet, round-trip, load, append more rows after the last populated row, round-trip again →
  // { loadedRowCount, finalRowCount, rows }. The load-bearing fact: a reloaded sheet reports its last
  // populated row so addRow lands new content at N+1 with no gap or overwrite.
  appendRowsAfterReload(initial: RowInput[] = [], append: RowInput[] = []) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    for (const row of initial) sheet.addRow(row);

    const loaded = roundtrip(workbook);
    const s = loaded.getWorksheet('S');
    const loadedRowCount = s!.rowCount;
    for (const row of append) s!.addRow(row);

    const final = roundtrip(loaded);
    const f = final.getWorksheet('S');
    // Mirror the oracle's `row.values.slice(1)` per-row array: each row is sized to its own populated
    // extent, holes are null, and an empty row is an empty array, indexed by row number so a gap shows.
    const rows: Untyped[] = Array.from({length: f!.rowCount}, () => []);
    for (const {number, cells} of f!.rows()) {
      const maxCol = cells.reduce((m: number, c: Untyped) => Math.max(m, c.col), 0);
      const arr = new Array(maxCol).fill(null);
      for (const cell of cells) arr[cell.col - 1] = normalizeStreamValue(cell.value);
      rows[number - 1] = arr;
    }
    return {loadedRowCount, finalRowCount: f!.rowCount, rows};
  },

  // Read a fixture's single _xlnm.Print_Area name (a comma-separated range list), re-write it, and read
  // it again → { sourceRangeCount, readPrintArea, rewrittenRangeCount }. Both disjoint ranges must
  // survive read and re-serialization, never truncated to the first.
  roundtripFixturePrintAreas(rel: string) {
    const printAreaOf = (wb: WorkbookInstance) =>
      wb.definedNames.find((n) => n.name === '_xlnm.Print_Area')?.refersTo ?? '';
    const source = readXlsx(fixtureBytes(rel));
    const readPrintArea = printAreaOf(source);
    const sourceRangeCount = readPrintArea.split(',').filter(Boolean).length;
    const rewritten = roundtrip(source);
    const rewrittenRangeCount = printAreaOf(rewritten).split(',').filter(Boolean).length;
    return {sourceRangeCount, readPrintArea, rewrittenRangeCount};
  },

  // Author a sheet-scoped _xlnm.Print_Area over a comma-separated area, round-trip, and report the
  // emitted ranges (sheet prefix stripped) → { ranges }. Two disjoint areas must emit two proper
  // rectangular ranges in one name, not a truncated single range.
  writePrintAreaDefinedName(area: string) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    workbook.defineName({
      name: '_xlnm.Print_Area',
      scope: sheet.name,
      refersTo: printAreaRefersTo(sheet.name, area),
    });
    const back = roundtrip(workbook);
    const refersTo = back.definedNames.find((n) => n.name === '_xlnm.Print_Area')?.refersTo ?? '';
    const ranges = refersTo.split(',').map((r) => r.split('!').pop());
    return {ranges};
  },

  // Author a sheet-scoped _xlnm.Print_Area over one area (whole-column or bounded), round-trip, and
  // report the written and recovered forms → { writtenDefinedName, reReadPrintArea, reloadOk }. A
  // column-only reference ($A:$D) must recover intact, never decoded to a NaN-mangled address.
  printAreaRoundtrip(area: string) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    workbook.defineName({
      name: '_xlnm.Print_Area',
      scope: sheet.name,
      refersTo: printAreaRefersTo(sheet.name, area),
    });
    let reloadOk = true;
    // Explicitly null rather than left unassigned: the `?.` below is load-bearing for the failed-read
    // path, and `any` let the declaration stay silent about the state that makes it so.
    let back: WorkbookInstance | null = null;
    try {
      back = roundtrip(workbook);
    } catch {
      reloadOk = false;
    }
    const refersTo = back?.definedNames.find((n) => n.name === '_xlnm.Print_Area')?.refersTo ?? '';
    const reReadPrintArea = refersTo.split('!').pop()?.replace(/\$/g, '') ?? '';
    return {writtenDefinedName: refersTo, reReadPrintArea, reloadOk};
  },

  // Freeze the first row, write, and report the emitted pane plus a round-trip → { paneEmitted,
  // reReadState, reReadYSplit, reReadXSplit }. A frozen-header view serializes a <pane> and reloads
  // as a frozen split of one row and no columns.
  frozenTopRowRoundtrip() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getCell('A1').value = 'header';
    sheet.freeze(1);
    const buffer = writeXlsx(workbook);
    const sheetXml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const paneEmitted = /<pane\b[^>]*ySplit="1"[^>]*state="frozen"/.test(sheetXml);
    const view = readXlsx(buffer).getWorksheet('S')!.view;
    return {
      paneEmitted,
      reReadState: view.state ?? 'normal',
      reReadYSplit: view.ySplit ?? 0,
      reReadXSplit: view.xSplit ?? 0,
    };
  },

  // Patch a frozen pane's split attributes with a malformed value, reload, and write the reloaded
  // workbook back out -> [{spelling, xSplit, ySplit, rewriteError}]. A file may spell a split as a
  // fraction, a negative, or a word; the authoring API (`freeze`) refuses all three, so a read that
  // stores one puts the model into a state its own writer cannot serialize.
  malformedFrozenSplitReport(spellings = ['abc', '1.5', '-3']) {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getCell('A1').value = 'header';
    sheet.freeze(1, 1);
    const bytes = writeXlsx(wb);
    return spellings.map((spelling) => {
      const back = reloadPatched(bytes, {
        'xl/worksheets/sheet1.xml': (xml) =>
          xml.replace(/(<pane\b[^>]*?)xSplit="[^"]*"/, `$1xSplit="${spelling}"`),
      });
      const {xSplit, ySplit} = back.getWorksheet('S')!.view;
      let rewriteError: string | null = null;
      try {
        writeXlsx(back);
      } catch (error) {
        rewriteError = messageOf(error);
      }
      return {spelling, xSplit: xSplit ?? null, ySplit: ySplit ?? null, rewriteError};
    });
  },

  // Author the shape a generated report has, a frozen top row above grouped, hidden columns on the
  // first of two sheets, then write it and report the view-initialisation facts of that package →
  // { bookViewCount, bookViewsBeforeSheets, windowWidth, windowHeight, selectedSheets,
  // outlineLevelCol, paneEmitted }. A consumer restores the document window from the workbook view
  // and lays the frozen pane and the column-outline bar out inside it; omitting those facts leaves
  // the top pane measured against an uninitialised window.
  frozenHeaderViewStateReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('Report');
    sheet.getCell('A1').value = 'header';
    sheet.getCell('A2').value = 'body';
    for (const index of [2, 3]) {
      sheet.getColumn(index).outlineLevel = 1;
      sheet.getColumn(index).hidden = true;
    }
    sheet.freeze(1);
    wb.addWorksheet('Notes').getCell('A1').value = 'notes';

    const parts = partMapOf(writeXlsx(wb));
    const workbookXml = parts['xl/workbook.xml'] || '';
    const sheetXml = parts['xl/worksheets/sheet1.xml'] || '';
    const bookView = (workbookXml.match(/<workbookView\b[^>]*\/>/) || [''])[0] || '';
    const attrNumber = (tag: string, name: string): number | null => {
      const found = tag.match(new RegExp(`\\b${name}="(-?\\d+(?:\\.\\d+)?)"`));
      return found ? Number(found[1]) : null;
    };
    const selectedSheets = ['Report', 'Notes'].filter((_, i) =>
      /<sheetView\b[^>]*\btabSelected="1"/.test(parts[`xl/worksheets/sheet${i + 1}.xml`] || ''),
    );
    const sheetFormat = (sheetXml.match(/<sheetFormatPr\b[^>]*\/>/) || [''])[0] || '';
    return {
      bookViewCount: (workbookXml.match(/<workbookView\b/g) || []).length,
      bookViewsBeforeSheets:
        workbookXml.indexOf('<bookViews>') >= 0 &&
        workbookXml.indexOf('<bookViews>') < workbookXml.indexOf('<sheets>'),
      windowWidth: attrNumber(bookView, 'windowWidth'),
      windowHeight: attrNumber(bookView, 'windowHeight'),
      selectedSheets,
      outlineLevelCol: attrNumber(sheetFormat, 'outlineLevelCol'),
      outlineLevelRow: attrNumber(sheetFormat, 'outlineLevelRow'),
      paneEmitted: /<pane\b[^>]*ySplit="1"[^>]*state="frozen"/.test(sheetXml),
    };
  },

  // Freeze a pane on one sheet, transplant that sheet's model onto another, write, and report what
  // reached the destination → { dstPaneEmitted, dstState, dstXSplit, dstYSplit, srcPaneEmitted }.
  // A frozen pane is workbook-independent sheet state, so it must ride the model transplant and reach
  // the written package as a <pane>; the source's own pane must survive the read of its model.
  frozenPaneSurvivesModelTransplant() {
    const workbook = new Workbook();
    const src = workbook.addWorksheet('Src');
    src.getCell('A1').value = 'header';
    src.freeze(2, 1);
    const dst = workbook.addWorksheet('Dst');
    dst.model = src.model;

    const parts = partMapOf(writeXlsx(workbook));
    const paneOf = (xml: string): boolean =>
      /<pane\b[^>]*xSplit="1"[^>]*ySplit="2"[^>]*state="frozen"/.test(xml);
    const view = roundtrip(workbook).getWorksheet('Dst')!.view;
    return {
      dstPaneEmitted: paneOf(parts['xl/worksheets/sheet2.xml'] || ''),
      dstState: view.state ?? 'normal',
      dstXSplit: view.xSplit ?? 0,
      dstYSplit: view.ySplit ?? 0,
      srcPaneEmitted: paneOf(parts['xl/worksheets/sheet1.xml'] || ''),
    };
  },

  // Freeze a view, then unfreeze it, and report the pane presence in each written form plus a reload
  // → { frozenHasPane, normalHasPane, reloadedState, reloadedHasSplit }. Unfreezing must leave no
  // leftover <pane> (which triggers Excel's repair prompt) and reload as a normal, unsplit view.
  unfreezeViewRoundtrip() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getCell('A1').value = 'header';
    sheet.freeze(1);
    const frozenHasPane = /<pane\b/.test(
      partMapOf(writeXlsx(workbook))['xl/worksheets/sheet1.xml'] || '',
    );

    sheet.unfreeze();
    const normalBuffer = writeXlsx(workbook);
    const normalHasPane = /<pane\b/.test(partMapOf(normalBuffer)['xl/worksheets/sheet1.xml'] || '');
    const view = readXlsx(normalBuffer).getWorksheet('S')!.view;
    return {
      frozenHasPane,
      normalHasPane,
      reloadedState: view.state ?? 'normal',
      reloadedHasSplit: (view.xSplit ?? 0) > 0 || (view.ySplit ?? 0) > 0,
    };
  },

  // Author three columns with distinct widths (one hidden), write, then REVERSE the order of the
  // emitted `<col>` tags, the shape foreign generators (excelize, jxls-poi) produce, and read the
  // patched package back → { w1, w2, w3, hidden2 }. Each column's width and hidden flag must bind to
  // the column its min/max names, regardless of document order.
  outOfOrderColumnsReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getColumn(1).width = 10;
    sheet.getColumn(2).width = 20;
    sheet.getColumn(2).hidden = true;
    sheet.getColumn(3).width = 30;
    const back = reloadPatched(writeXlsx(wb), {
      'xl/worksheets/sheet1.xml': (xml) =>
        xml.replace(/<cols>([\s\S]*?)<\/cols>/, (_, inner) => {
          const tags = inner.match(/<col\b[^>]*\/>/g) || [];
          return `<cols>${tags.reverse().join('')}</cols>`;
        }),
    }).getWorksheet('S')!;
    return {
      w1: back.getColumn(1).width,
      w2: back.getColumn(2).width,
      w3: back.getColumn(3).width,
      hidden2: back.getColumn(2).hidden ?? false,
    };
  },

  // Patch a written sheet's `<cols>`/`<sheetData>` with the out-of-grid spans a hostile or broken
  // producer emits, then read it back → { spanColumnCount, lastSpannedWidth, beyondColumnCount,
  // beyondRowCount }. A `<col max="99999999">` names more columns than the format has; a reader that
  // walks the span verbatim allocates until it dies, so the counts below are the allocation bound.
  hostileGridSpanReport(spanMax = 99999999) {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getCell('A1').value = 'x';
    const bytes = writeXlsx(wb);
    const patch = (cols: string, rows: string) =>
      reloadPatched(bytes, {
        'xl/worksheets/sheet1.xml': (xml) =>
          xml
            .replace(/<cols>[\s\S]*?<\/cols>/, '')
            .replace(/<sheetData>/, `${cols}<sheetData>`)
            .replace(/<\/sheetData>/, `${rows}</sheetData>`),
      }).getWorksheet('S')!;

    const span = patch(
      `<cols><col min="1" max="${spanMax}" width="12" customWidth="1"/></cols>`,
      '',
    );
    const beyond = patch(
      `<cols><col min="${MAX_COLUMN + 1}" max="${spanMax}" width="12" customWidth="1"/></cols>`,
      `<row r="${MAX_ROW + 1}" ht="30" customHeight="1"/>`,
    );
    return {
      spanColumnCount: [...span.columns()].length,
      lastSpannedWidth: span.getColumn(MAX_COLUMN).width,
      beyondColumnCount: [...beyond.columns()].length,
      // The base sheet holds A1, so count what lies past the grid rather than what lies in it.
      beyondRowCount: [...beyond.rows()].filter((row) => row.number > MAX_ROW).length,
      maxColumn: MAX_COLUMN,
    };
  },

  // Patch a written sheet with many full-grid `<col>` spans, then time both readers over it →
  // { xmlBytes, spans, bufferedMs, streamingMs, columnCount }. Clamping ONE span to the grid bounds
  // that span's loop and nothing else: the number of `<col>` elements is unbounded and each may span
  // the whole grid, so the cost is their product. 154 KB of worksheet XML holding 2,000 such spans cost
  // the buffered reader 12.3 s, which against the 512 MiB inflate ceiling extrapolates to hours of CPU
  // from a package compressing to a few hundred KB. The zip-bomb guard cannot see it, because after
  // inflation the payload really is small. Use it to assert the per-sheet work budget bounds the time.
  repeatedFullGridColumnSpanReport(spans = 4000) {
    const wb = new Workbook();
    wb.addWorksheet('S').getCell('A1').value = 'x';
    const cols =
      '<cols>' +
      Array.from(
        {length: spans},
        () => '<col min="1" max="99999999" width="12" customWidth="1" hidden="1" style="0"/>',
      ).join('') +
      '</cols>';
    const archive = patchedPackage(writeXlsx(wb), {
      edit: {'xl/worksheets/sheet1.xml': (xml) => xml.replace('<sheetData>', `${cols}<sheetData>`)},
    });

    const bufferedStart = performance.now();
    const back = readXlsx(archive).getWorksheet('S')!;
    const bufferedMs = performance.now() - bufferedStart;

    const streamingStart = performance.now();
    for (const streamed of readWorkbookStream(archive)) {
      for (const _row of streamed.rows()) {
        // Drained rather than sampled: the `<col>` handling runs inside the row scan.
      }
      void streamed.hiddenColumns;
    }
    const streamingMs = performance.now() - streamingStart;

    return {
      xmlBytes: partOf(archive, 'xl/worksheets/sheet1.xml').length,
      spans,
      bufferedMs,
      streamingMs,
      columnCount: [...back.columns()].length,
      maxColumn: MAX_COLUMN,
    };
  },

  // Put a `tokenLength`-character junk token with no `=` into A1's open tag, then time both readers
  // → { tokenLength, zippedBytes, bufferedMs, buffered, streamingMs, streamed }, where `buffered` and
  // `streamed` are A1's value as each reader saw it. The token is the input that made a backtracking
  // attribute scan quadratic; the `r` after it is what proves the scan still reads past it.
  junkAttributeTokenReport(tokenLength: number) {
    const wb = new Workbook();
    wb.addWorksheet('S').getCell('A1').value = 42;
    const archive = patchedPackage(writeXlsx(wb), {
      edit: {
        'xl/worksheets/sheet1.xml': (xml) =>
          xml.replace('<c r="A1"', `<c ${'j'.repeat(tokenLength)} r="A1"`),
      },
    });

    const bufferedStart = performance.now();
    const buffered = readXlsx(archive).getWorksheet('S')?.getCell('A1').value ?? null;
    const bufferedMs = performance.now() - bufferedStart;

    const streamingStart = performance.now();
    let streamed: unknown = null;
    for (const sheet of readWorkbookStream(archive)) {
      for (const row of sheet.rows()) {
        for (const cell of row.cells) if (cell.address === 'A1') streamed = cell.value;
      }
    }
    const streamingMs = performance.now() - streamingStart;

    return {tokenLength, zippedBytes: archive.length, bufferedMs, buffered, streamingMs, streamed};
  },

  // Assign an outline (grouping) level to a row and a column, write, and read back → { rowOutline,
  // colOutline }. The OOXML outlineLevel attribute on <row>/<col> must survive the round-trip on both
  // axes so a collapsible grouping is preserved on reopen.
  rowColumnOutlineLevelRoundtrip() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getCell('A1').value = 'x';
    sheet.getRow(2).outlineLevel = 1;
    sheet.getColumn(3).outlineLevel = 1;
    const back = roundtrip(wb).getWorksheet('S')!;
    return {
      rowOutline: back.getRow(2).outlineLevel ?? 0,
      colOutline: back.getColumn(3).outlineLevel ?? 0,
    };
  },

  // Set explicit widths on three columns (one of which coincides with the format's conventional
  // default width, 9), write, and report per-column whether a `<col>` with a customWidth flag was
  // emitted and what width each reads back → { emitted: {c1,c2,c3}, readBack: {c1,c2,c3} }. An
  // explicitly-set width must survive even when its value equals the magic default.
  columnWidthDefaultCollisionReport(widths = [8, 9, 10]) {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    widths.forEach((w, i) => {
      sheet.getColumn(i + 1).width = w;
    });
    const buffer = writeXlsx(wb);
    const sheetXml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    // A column emits an explicit width when its `<col>` carries a customWidth flag over its index.
    const emittedAt = (index: number) =>
      new RegExp(`<col\\b[^>]*\\bmin="${index}"[^>]*\\bmax="${index}"[^>]*\\bcustomWidth="1"`).test(
        sheetXml,
      );
    const back = readXlsx(buffer).getWorksheet('S')!;
    return {
      emitted: {c1: emittedAt(1), c2: emittedAt(2), c3: emittedAt(3)},
      readBack: {
        c1: back.getColumn(1).width,
        c2: back.getColumn(2).width,
        c3: back.getColumn(3).width,
      },
    };
  },

  // Append rows in every shape (dense array, sparse array, keyed object, mixed batch), round-trip, and
  // read them back letter-keyed by row number → { rows }. Column keys bind object values to columns;
  // dense/sparse arrays map positionally with holes left empty; a numeric/date value survives typed.
  appendRowShapes() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getColumn(1).key = 'k1';
    sheet.getColumn(2).key = 'k2';
    sheet.addRow(['header']); // row 1: keeps the checked rows at their stated numbers
    sheet.addRow(['a', 'b', 'c']); // row 2: dense positional array
    // oxlint-disable-next-line eslint/no-sparse-arrays -- a genuine array hole (not undefined) is the point: the gap at column B must be skipped positionally
    sheet.addRow(['x', , 'z']); // row 3: sparse array, gap at column B
    sheet.addRow({k1: 'o1', k2: 'o2'}); // row 4: keyed object
    sheet.addRow([7, new Date(Date.UTC(2021, 0, 2))]); // row 5: number + date
    sheet.addRows([['m1', 'm2'], {k1: 'n1'}]); // rows 6, 7: mixed batch

    const loaded = roundtrip(workbook);
    const s = loaded.getWorksheet('S');
    const rows: Record<string, Record<string, Untyped>> = {};
    for (const {number, cells} of s!.rows()) {
      const row: Record<string, Untyped> = {};
      for (const cell of cells)
        row[encodeAddress(cell.col, number).match(/^[A-Z]+/)![0]] = normalizeStreamValue(
          cell.value,
        );
      rows[number] = row;
    }
    // Every checked column reads as null when the round-trip left it empty, so a gap is visible.
    for (const n of Object.keys(rows)) for (const col of ['A', 'B', 'C']) rows[n]![col] ??= null;
    return {rows};
  },

  // Lay out a template the way an author does (a header, then a band of rows and columns styled
  // ahead of the data that will fill them), then append, and report where the append landed and what
  // survived → { rowCountBeforeAppend, appendedRowValue, styledRowValue, styledRowKeptFill,
  // columnCountBeforeAppend, appendedColumnValue, styledColumnValue, actualRowCount,
  // usedRangeAfterRoundtrip }. Formatting a cell is how a caller claims it; an append that treats a
  // pre-formatted line as free ground writes over the layout, silently merging two of the author's
  // rows into one.
  appendOverPreformattedBandReport() {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'header';
    ws.getCell('A2').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFFFF00'}};
    const rowCountBeforeAppend = ws.rowCount;
    ws.addRow(['appended']);

    ws.getCell('B1').font = {bold: true};
    const columnCountBeforeAppend = ws.columnCount;
    ws.addColumn(['appended-col']);

    const rt = roundtrip(wb).getWorksheet('S')!;
    return {
      rowCountBeforeAppend,
      appendedRowValue: ws.getCell('A3').value ?? null,
      styledRowValue: ws.getCell('A2').value ?? null,
      styledRowKeptFill: ws.getCell('A2').fill?.type ?? null,
      columnCountBeforeAppend,
      appendedColumnValue: ws.getCell('C1').value ?? null,
      styledColumnValue: ws.getCell('B1').value ?? null,
      // A formatting-only line bounds the used range but is not a populated row: the two counts
      // answer different questions and must not collapse into one.
      actualRowCount: ws.actualRowCount,
      usedRangeAfterRoundtrip: rt.usedRange?.address ?? null,
    };
  },

  // Insert and append the SAME column of values three ways and report what each wrote, after a
  // round-trip → { intoEmpty, intoShortGrid, appended }, each a per-row map of the inserted column
  // plus the value the sheet already held. A column's values are indexed by row, so most of the rows
  // one names do not exist yet on a sheet being filled in. The insert pass ran inside the loop over
  // the rows the grid already had, so it could only write onto those: inserting into an empty sheet
  // wrote nothing at all, and inserting beside a single cell wrote the first value and dropped the
  // rest, while `addColumn` handed the identical array materialised every row of it.
  columnInsertMaterialisesEveryValue() {
    const report = (build: (sheet: ReturnType<WorkbookInstance['addWorksheet']>) => void) => {
      const wb = new Workbook();
      build(wb.addWorksheet('S'));
      const s = roundtrip(wb).getWorksheet('S')!;
      return {
        rowCount: s.rowCount,
        column: ['A1', 'A2', 'A3'].map((ref) => s.getCell(ref).value ?? null),
        shifted: s.getCell('B1').value ?? null,
      };
    };
    return {
      intoEmpty: report((s) => {
        s.insertColumn(1, ['x', 'y', 'z']);
      }),
      intoShortGrid: report((s) => {
        s.getCell('A1').value = 'was here';
        s.insertColumn(1, ['x', 'y', 'z']);
      }),
      appended: report((s) => {
        s.addColumn(['x', 'y', 'z']);
      }),
    };
  },

  // Feed addRow an array built in another realm (a vm context): Array.isArray must recognize it so its
  // elements fill columns → { isArrayCrossRealm, a, b, c }. `instanceof Array` would miss it and place
  // nothing, walking it as a keyed object instead.
  async crossRealmArrayRow() {
    const vm = await import('node:vm');
    const arr = vm.runInNewContext('[10, 20, 30]');
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.addRow(arr);
    return {
      isArrayCrossRealm: Array.isArray(arr),
      a: sheet.getCell('A1').value ?? null,
      b: sheet.getCell('B1').value ?? null,
      c: sheet.getCell('C1').value ?? null,
    };
  },

  // Set row-level properties on rows that carry NO cell value and report what survives a round-trip
  // → { row3Hidden, row4Hidden, row4Height, row5Hidden }. A content-less row bearing a hidden flag,
  // a height, or an outline level must still be written (its <row> element materialised) so the
  // property is not lost: the failure mode is a blank hidden/grouped spacer row coming back visible.
  hiddenEmptyRowReport() {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getRow(3).hidden = true;
    ws.getRow(4).hidden = true;
    ws.getRow(4).height = 25;
    ws.getRow(5).hidden = true;
    ws.getRow(5).outlineLevel = 1;
    const rt = roundtrip(wb).getWorksheet('S')!;
    return {
      row3Hidden: rt.getRow(3).hidden ?? false,
      row4Hidden: rt.getRow(4).hidden ?? false,
      row4Height: rt.getRow(4).height ?? null,
      row5Hidden: rt.getRow(5).hidden ?? false,
    };
  },

  // Prove manual horizontal page breaks survive load and a load→save round-trip → { sourceBreaks,
  // loadedBreaks, rewrittenBreaks }, each the ascending list of break row ids. sourceBreaks reads the
  // raw fixture XML (the precondition); loadedBreaks/rewrittenBreaks come off the model after read and
  // after write→re-read, so a dropped-on-read or dropped-on-write regression shows as an empty list.
  roundtripFixtureRowBreaks(rel: string) {
    const rowBreakIds = (xml: string) => {
      const section = xml.match(/<rowBreaks[\s\S]*?<\/rowBreaks>/);
      if (section === null) return [];
      return [...section[0].matchAll(/<brk\b[^>]*\bid="(\d+)"/g)]
        .map((m) => Number(m[1]))
        .sort((a, b) => a - b);
    };
    const sheet1 = (parts: PartMap) => parts['xl/worksheets/sheet1.xml'] ?? '';
    const modelBreaks = (wb: WorkbookInstance) =>
      wb.worksheets[0]!.rowBreaks.map((brk) => brk.id).sort((a: number, b: number) => a - b);

    const sourceBreaks = rowBreakIds(sheet1(partMapOf(fixtureBytes(rel))));
    const loaded = readFixture(rel);
    const loadedBreaks = modelBreaks(loaded);
    const rewrittenBreaks = modelBreaks(roundtrip(loaded));
    return {sourceBreaks, loadedBreaks, rewrittenBreaks};
  },

  // Patch a written sheet with one manual row break and one automatic row break carrying a `min`, read
  // it, and write it again → { read, rewritten }: `read` is the model's breaks as plain objects, and
  // `rewritten` is the second write's `<rowBreaks>` `{count, manualBreakCount}` beside each `<brk>`'s
  // `{id, min, max, man}` attributes, an absent attribute reported as null.
  pageBreakKindsReport() {
    const wb = new Workbook();
    wb.addWorksheet('S').getCell('A1').value = 'x';
    const breaks =
      '<rowBreaks count="2" manualBreakCount="1">' +
      '<brk id="3" max="16383" man="1"/><brk id="7" min="2" max="16383"/></rowBreaks>';
    const read = reloadPatched(writeXlsx(wb), {
      'xl/worksheets/sheet1.xml': (xml) => xml.replace('</worksheet>', `${breaks}</worksheet>`),
    });
    const section =
      (partMapOf(writeXlsx(read))['xl/worksheets/sheet1.xml'] ?? '').match(
        /<rowBreaks\b[^>]*>[\s\S]*?<\/rowBreaks>/,
      )?.[0] ?? '<rowBreaks>';
    const container = attrsOf(section.match(/<rowBreaks\b[^>]*>/)?.[0] ?? '<rowBreaks>');
    return {
      read: (read.getWorksheet('S')?.rowBreaks ?? []).map((brk) => ({...brk})),
      rewritten: {
        counts: {
          count: container.count ?? null,
          manualBreakCount: container.manualBreakCount ?? null,
        },
        breaks: [...section.matchAll(/<brk\b[^>]*\/>/g)].map(([element]) => {
          const attrs = attrsOf(element);
          return {
            id: attrs.id ?? null,
            min: attrs.min ?? null,
            max: attrs.max ?? null,
            man: attrs.man ?? null,
          };
        }),
      },
    };
  },

  // Merge A1:B3 with values only in A1/A2, round-trip, then iterate every used-range position
  // (include-empty) and report → { rowCount, visited, a3: { visited, isMerged, master } }. A merge
  // reaching into an otherwise-empty trailing row keeps that row within the bounds, so A3 is visited
  // and resolves to its master A1 rather than being skipped.
  trailingMergedRowIterationReport() {
    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    s.getCell('A1').value = 'top';
    s.getCell('A2').value = 'data';
    s.mergeCells('A1:B3');
    const rs = roundtrip(wb).worksheets[0]!;
    const rects = rs.merges.map((range: Untyped) => {
      const {top, left, bottom, right} = decodeRange(range);
      return {top, left, bottom, right, masterRef: encodeAddress(left!, top!)};
    });
    const masterOf = (row: Untyped, col: Untyped) =>
      rects.find((r) => row >= r.top! && row <= r.bottom! && col >= r.left! && col <= r.right!);
    const visited: Untyped[] = [];
    for (let row = 1; row <= rs.rowCount; row++) {
      for (let col = 1; col <= rs.columnCount; col++) visited.push(encodeAddress(col, row));
    }
    const a3Rect = masterOf(3, 1);
    return {
      rowCount: rs.rowCount,
      visited,
      a3: {
        visited: visited.includes('A3'),
        isMerged: a3Rect !== undefined,
        master: a3Rect ? a3Rect.masterRef : null,
      },
    };
  },

  // Build a workbook from a spec, round-trip it, and for each requested row report the column indices
  // an include-empty iteration yields → { rows: { <n>: { cols } }, columnCount }. Positional iteration
  // walks 1..columnCount (the sheet's declared width), so interior *and* trailing empties are surfaced
  // and every row reconstructs to the header width, the alignment invariant a positional consumer needs.
  async readRowCellPresence(spec: Untyped, rowNumbers: number[] = []) {
    const sheet = roundtrip(buildFrom(spec)).worksheets[0]!;
    const columnCount = sheet.columnCount;
    const rows: Record<string, Untyped> = {};
    for (const rn of rowNumbers) {
      const cols: Untyped[] = [];
      for (let col = 1; col <= columnCount; col++) {
        sheet.getCell(encodeAddress(col, rn));
        cols.push(col);
      }
      rows[rn] = {cols};
    }
    return {rows, columnCount};
  },

  // Merge A1:B1 with the value in the master (A1), then read the display text of the master and of a
  // merged child (B1) → { masterText, childText, childThrew }. Addressing a covered cell resolves to
  // its region's master, so a merged child's text mirrors the master and never throws.
  mergedCellDisplayTextReport() {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'Group';
    ws.mergeCells('A1:B1');
    // `String` on a `CellValue` renders every non-primitive shape as `[object Object]`, which would
    // make a rich-text or formula read indistinguishable from any other. This probe only ever reads the
    // string it just wrote, so anything else arriving here is a failure worth being able to see.
    const textOf = (ref: string) => {
      const v = ws.getCell(ref).value;
      if (v === null || v === undefined) return '';
      return typeof v === 'object' ? JSON.stringify(v) : String(v);
    };
    let masterText: string | null;
    let childText = null;
    let childThrew = false;
    try {
      masterText = textOf('A1');
    } catch {
      masterText = null;
    }
    try {
      childText = textOf('B1');
    } catch {
      childThrew = true;
    }
    return {masterText, childText, childThrew};
  },

  // Merge a horizontal span with a value + alignment on the anchor, write, then read back →
  // { mergeCount, merges, populatedCoveredCells, anchorValue, anchorAlignment }. A clean merge
  // declares the range exactly once and emits a value only on the anchor, so the covered cells
  // carry no conflicting <v>, the shape that opens without Excel's repair prompt, and the
  // anchor's value and alignment survive the round-trip.
  mergeCleanReport({anchor = 'B1', range = 'B1:G1', value = 'Group Title'}: Untyped = {}) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    const cell = sheet.getCell(anchor);
    cell.value = value;
    cell.alignment = {horizontal: 'center'};
    sheet.mergeCells(range);
    const buffer = writeXlsx(workbook);
    const sheetXml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const merges = [...sheetXml.matchAll(/<mergeCell\b[^>]*ref="([^"]*)"/g)].map((m) => m[1]);
    const {left, right, top, bottom} = decodeRange(range);
    const populatedCoveredCells: Untyped[] = [];
    for (let r = top!; r <= bottom!; r++) {
      for (let c = left!; c <= right!; c++) {
        const ref = encodeAddress(c, r);
        if (ref === anchor) continue;
        if (new RegExp(`<c\\b[^>]*\\br="${ref}"[^>]*>[\\s\\S]*?<v>`).test(sheetXml))
          populatedCoveredCells.push(ref);
      }
    }
    const a = readXlsx(buffer).getWorksheet('S')!.getCell(anchor);
    return {
      mergeCount: merges.length,
      merges,
      populatedCoveredCells,
      anchorValue: a.value ?? null,
      anchorAlignment: a.alignment ? {...a.alignment} : null,
    };
  },

  // Populate covered non-anchor cells FIRST, then merge over them: the order that leaves stray
  // values behind. Write and read back → { anchorValue, populatedCoveredCells, coveredValuesOnRead }.
  // Excel keeps only the anchor's value on merge; a covered cell that still carries a <v> under the
  // <mergeCell> ref is the geometry that trips Excel's repair prompt.
  mergeOverPopulatedReport({anchor = 'B1', range = 'B1:G1', value = 'Group Title'}: Untyped = {}) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    const {left, right, top, bottom} = decodeRange(range);
    // Fill every cell of the range, anchor included, before the merge collapses it.
    for (let r = top!; r <= bottom!; r++) {
      for (let c = left!; c <= right!; c++) {
        sheet.getCell(encodeAddress(c, r)).value = c === left && r === top ? value : `covered-${c}`;
      }
    }
    sheet.mergeCells(range);
    const buffer = writeXlsx(workbook);
    const sheetXml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const populatedCoveredCells: Untyped[] = [];
    for (let r = top!; r <= bottom!; r++) {
      for (let c = left!; c <= right!; c++) {
        const ref = encodeAddress(c, r);
        if (ref === anchor) continue;
        if (new RegExp(`<c\\b[^>]*\\br="${ref}"[^>]*>[\\s\\S]*?<v>`).test(sheetXml))
          populatedCoveredCells.push(ref);
      }
    }
    const reread = readXlsx(buffer).getWorksheet('S')!;
    return {
      anchorValue: reread.getCell(anchor).value ?? null,
      populatedCoveredCells,
    };
  },

  // Give the top-left/master cell a border (+ numFmt + font), merge it into a region, round-trip,
  // then report the master's border/numFmt/font and the declared merges → for asserting a merge
  // does not strip the style the master needs to render the merged region's outline.
  mergeMasterBorderReport() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    const cell = sheet.getCell('A1');
    cell.value = 'x';
    cell.border = {top: {style: 'thin'}, bottom: {style: 'medium'}};
    cell.numFmt = '0.00';
    cell.font = {bold: true};
    sheet.mergeCells('A1:B2');
    const reread = roundtrip(workbook).getWorksheet('S')!;
    const m = reread.getCell('A1');
    const b = m.border || {};
    return {
      hasTopBorder: !!b.top?.style,
      hasBottomBorder: !!b.bottom?.style,
      topStyle: b.top ? (b.top.style ?? null) : null,
      bottomStyle: b.bottom ? (b.bottom.style ?? null) : null,
      numFmt: m.numFmt ?? null,
      fontBold: !!m.font?.bold,
      merges: [...reread.merges],
    };
  },

  // Merge a rectangular region (master = top-left), set a value by addressing a NON-master
  // (slave) cell inside it, write, and report which cells carry an independent <v> in the sheet
  // XML, the declared merges, and the re-read master/slave values → for asserting the slave write
  // resolves to the master (only the master carries a value; reading either address returns it).
  mergeSlaveWrite({range = 'A1:B2', slave = 'B2', value = 'slave-write'}: Untyped = {}) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.mergeCells(range);
    sheet.getCell(slave).value = value;
    const buffer = writeXlsx(workbook);
    const sheetXml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    // A cell "carries a value" if its element has value content: a number/bool/formula (<v>),
    // an inline string (<is>), or a formula (<f>). The writer serialises strings as inlineStr,
    // so keying on <v> alone would miss them; an empty covered cell is never emitted at all.
    const cellsWithValue = [
      ...sheetXml.matchAll(/<c\b[^>]*\br="([A-Z]+\d+)"[^>]*>([\s\S]*?)<\/c>/g),
    ]
      .filter((m) => /<(?:v|is|f)\b/.test(m[2]!))
      .map((m) => m[1]);
    const merges = [...sheetXml.matchAll(/<mergeCell\b[^>]*ref="([^"]*)"/g)].map((m) => m[1]);
    const master = range.split(':')[0]!;
    const s = readXlsx(buffer).getWorksheet('S')!;
    return {
      cellsWithValue,
      merges,
      masterValue: s.getCell(master).value ?? null,
      slaveValue: s.getCell(slave).value ?? null,
    };
  },

  // Duplicate a populated row with default args, then merge a range on the copy. The copy must be
  // faithful (values, not empty/NaN) and carry no phantom merge that would reject the merge.
  duplicateRowReport() {
    const sheet = new Workbook().addWorksheet('S');
    sheet.getCell('A1').value = 'a';
    sheet.getCell('B1').value = 'b';
    sheet.getCell('C1').value = 'c';
    let dupError = null;
    try {
      sheet.duplicateRow(1, {count: 1, insert: true});
    } catch (e) {
      dupError = messageOf(e);
    }
    const val = (ref: string) => sheet.getCell(ref).value ?? null;
    const row1 = [val('A1'), val('B1'), val('C1')];
    const row2 = [val('A2'), val('B2'), val('C2')];
    let mergeError = null;
    try {
      sheet.mergeCells('A2:C2');
    } catch (e) {
      mergeError = messageOf(e);
    }
    return {dupError, mergeError, rowCount: sheet.rowCount, row1, row2, merges: [...sheet.merges]};
  },

  // Insert a row then style a cell of it. The inserted cells must stay mutable (no frozen,
  // "object is not extensible" style object) regardless of the requested style-inheritance mode. The
  // rewrite's copy-on-write style model makes every cell mutable by construction, so the mode is
  // immaterial; it is accepted and ignored.
  insertRowThenStyle(_styleMode = 'i') {
    const sheet = new Workbook().addWorksheet('S');
    sheet.getCell('A1').value = 'header';
    sheet.getCell('A1').font = {bold: true};
    sheet.getCell('A2').value = 'data';
    let error = null;
    let numFmt = null;
    try {
      sheet.insertRow(2, ['inserted']);
      const cell = sheet.getCell('A2');
      cell.numFmt = '$#,##0.00;[Red]-$#,##0.00';
      cell.font = {...cell.font, bold: true};
      numFmt = cell.numFmt;
    } catch (e) {
      error = messageOf(e);
    }
    return {error, numFmt};
  },

  // Attach a cell note and an outline level, insert a row above them, and round-trip through the real
  // writer/reader → { dataShifted, noteFollowsRow, outlineFollowsRow }. Both the note and the outline
  // level must follow their logical row through the insert and survive serialization.
  rowInsertPreservesNoteAndOutline() {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'r1';
    ws.getCell('A2').value = 'r2';
    ws.getCell('A2').note = 'mynote';
    ws.getRow(2).outlineLevel = 1;
    ws.insertRow(1, ['new']); // r1 -> row 2, r2 (noted, outlined) -> row 3
    const s = roundtrip(wb).getWorksheet('S')!;
    return {
      dataShifted: s.getCell('A2').value === 'r1' && s.getCell('A3').value === 'r2',
      noteFollowsRow: !!s.getCell('A3').note,
      outlineFollowsRow: s.getRow(3).outlineLevel === 1,
    };
  },

  // Anchor a whole-column validation, a full-height conditional format, a full-height autofilter and a
  // merge on the very last rows, then insert rows above them, write, and read the package back → the
  // edges the reloaded file reports, on both axes. A whole column is written bounded to the last row
  // (that is what Excel writes), so every one of these regions starts on the grid's edge and an insert
  // above it has nowhere to push the edge to. A file naming a row past the last one opens in Excel
  // with its repair prompt, so no edge here may exceed the grid.
  spliceHoldsGeometryInsideTheGrid() {
    const LAST_ROW = MAX_ROW;
    const rowAxis = new Workbook();
    const rows = rowAxis.addWorksheet('S');
    rows.getCell('A1').value = 'top';
    rows.addDataValidation(`B1:B${LAST_ROW}`, {type: 'list', formulae: ['"a,b,c"']});
    rows.addConditionalFormatting({
      ref: `C1:C${LAST_ROW}`,
      rules: [{type: 'dataBar', priority: 1}],
    });
    rows.autoFilter = `A1:A${LAST_ROW}`;
    rows.mergeCells(`E${LAST_ROW - 1}:F${LAST_ROW}`);
    rows.spliceRows(2, 0, ['inserted'], ['also']);

    const colAxis = new Workbook();
    const cols = colAxis.addWorksheet('S');
    cols.getCell('A1').value = 'left';
    cols.addDataValidation('A1:XFD1', {type: 'list', formulae: ['"a,b,c"']});
    cols.autoFilter = 'A2:XFD2';
    cols.spliceColumns(2, 0, ['inserted']);

    // Line metadata is anchored to the grid exactly as the regions above it are, and was the last
    // participant still shifting by hand. A height on the last row plus an insert put a properties
    // entry at 1048577, which made the sheet report a row count `new Row` refuses to construct: the
    // sheet could then not be iterated, so it could not be written at all.
    const lineAxis = new Workbook();
    const lines = lineAxis.addWorksheet('S');
    lines.getCell('A1').value = 'top';
    lines.getRow(MAX_ROW).height = 20;
    lines.getColumn(MAX_COLUMN).width = 12;
    lines.spliceRows(2, 0, ['inserted']);
    lines.spliceColumns(2, 0, ['inserted']);

    const reread = (wb: WorkbookInstance) => {
      const s = roundtrip(wb).getWorksheet('S')!;
      return {
        validationRefs: s.dataValidations.map((entry) => entry.sqref),
        formattingRefs: s.conditionalFormattings.map((entry) => entry.ref),
        autoFilterRef: s.autoFilter?.ref ?? null,
        merges: [...s.merges],
      };
    };
    const rereadLines = () => {
      try {
        const s = roundtrip(lineAxis).getWorksheet('S')!;
        return {
          writeError: null,
          rowCount: s.rowCount,
          columnCount: s.columnCount,
          lastRowHeight: s.getRow(MAX_ROW).height ?? null,
          lastColumnWidth: s.getColumn(MAX_COLUMN).width ?? null,
        };
      } catch (error) {
        return {
          writeError: messageOf(error),
          rowCount: -1,
          columnCount: -1,
          lastRowHeight: null,
          lastColumnWidth: null,
        };
      }
    };
    return {rows: reread(rowAxis), columns: reread(colAxis), lineProperties: rereadLines()};
  },

  // Anchor a data validation, a conditional format, a comment thread and an autofilter over the same
  // block, splice the row axis, write, and read the package back → the anchors as the reloaded file
  // reports them, beside the cell the moved content landed on. `inserted` puts a row above the block;
  // `deleted` removes the block's every row. Everything bound to a range must travel with the cells it
  // covers, and an anchor whose rows are gone must go with them rather than re-pointing at survivors.
  spliceReanchorsRangeBoundOverlays() {
    const AUTHOR = '{39236F6F-643D-4654-8264-DD21C8472F7F}';
    const anchored = () => {
      const wb = new Workbook();
      wb.addPerson({id: AUTHOR, displayName: 'Ada Lovelace', providerId: 'AD'});
      const s = wb.addWorksheet('S');
      s.getCell('B5').value = 'anchored';
      s.addDataValidation('B5:B6', {type: 'list', formulae: ['"a,b"']});
      s.addConditionalFormatting({ref: 'B5:B6', rules: [{type: 'dataBar', priority: 1}]});
      s.autoFilter = 'B5:C6';
      s.addCommentThread({
        ref: 'B5',
        resolved: false,
        comments: [
          {
            id: '{11111111-2222-3333-4444-555555555555}',
            personId: AUTHOR,
            text: 'about the block',
            mentions: [],
          },
        ],
      });
      return wb;
    };
    const report = (wb: WorkbookInstance, movedTo: string) => {
      const s = roundtrip(wb).getWorksheet('S')!;
      return {
        movedValue: s.getCell(movedTo).value,
        validationRefs: s.dataValidations.map((entry) => entry.sqref),
        formattingRefs: s.conditionalFormattings.map((entry) => entry.ref),
        autoFilterRef: s.autoFilter?.ref ?? null,
        threadRefs: s.commentThreads.map((thread) => thread.ref),
      };
    };

    const insertedWb = anchored();
    insertedWb.getWorksheet('S')!.insertRow(1, ['header']);
    const deletedWb = anchored();
    deletedWb.getWorksheet('S')!.spliceRows(5, 2);

    return {inserted: report(insertedWb, 'B6'), deleted: report(deletedWb, 'B5')};
  },

  spliceShiftsRefs() {
    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    // Table occupies A3:B5 (header + 2 rows); image anchored from row 5 (0-based).
    s.addTable({name: 'T', ref: 'A3', columns: [{name: 'H1'}, {name: 'H2'}], rowCount: 2});
    const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
    s.addImage(id, {tl: {col: 0, row: 5}, br: {col: 2, row: 8}});
    s.spliceRows(1, 0, ['inserted']); // insert a row at the top → table and image shift down 1

    const parts = partMapOf(writeXlsx(wb));
    const tableXml = parts['xl/tables/table1.xml'] || '';
    const tableRef = (tableXml.match(/<table\b[^>]*\bref="([^"]*)"/) || [])[1] ?? null;
    const drawingXml = parts['xl/drawings/drawing1.xml'] || '';
    const imageFromRow = (drawingXml.match(/<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/) || [])[1];

    // Duplicate table column names authored separately: construction disambiguates them into a
    // unique set rather than emitting a corrupt table (the same repair the reader applies on load).
    const w2 = new Workbook();
    const dupTable = w2
      .addWorksheet('S')
      .addTable({name: 'T2', ref: 'A1', columns: [{name: 'Dup'}, {name: 'Dup'}], rowCount: 1});
    writeXlsx(w2);
    const dupColumnNames = dupTable.columns.map((c) => c.name);
    const dupColumnNamesUnique =
      new Set(dupColumnNames.map((n) => n.toLowerCase())).size === dupColumnNames.length;

    return {
      tableRef,
      imageFromRow: imageFromRow != null ? Number(imageFromRow) : null,
      dupColumnNames,
      dupColumnNamesUnique,
    };
  },

  // Define four adjacent columns with identical width and outline level, write, and report whether
  // the write and reload succeed and how many <col> spans the part carries. Equivalent adjacent
  // columns must coalesce into fewer <col> spans than columns, without the collapse pass throwing.
  equivalentColumnCollapseReport() {
    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    for (let i = 1; i <= 4; i++) {
      const col = s.getColumn(i);
      col.width = 12;
      col.outlineLevel = 1;
    }
    let writeOk = true;
    let writeError = null;
    let buffer: Uint8Array | null = null;
    try {
      buffer = writeXlsx(wb);
    } catch (e) {
      writeOk = false;
      writeError = messageOf(e);
    }
    if (!writeOk) return {writeOk, writeError, reloadOk: false, colSpanCount: null};

    const parts = partMapOf(buffer!);
    const sheetPart = Object.keys(parts).find((n) => /xl\/worksheets\/sheet\d+\.xml$/.test(n));
    const colsBlock = (parts[sheetPart!]!.match(/<cols>[\s\S]*?<\/cols>/) || [])[0] ?? '';
    const colSpanCount = (colsBlock.match(/<col\b/g) || []).length;

    let reloadOk = true;
    try {
      readXlsx(buffer!);
    } catch {
      reloadOk = false;
    }
    return {writeOk, writeError, reloadOk, colSpanCount};
  },

  // Write a sheet carrying its own autofilter, row break, footer and margins, patch in a
  // `<customSheetViews>` block whose saved view declares a frozen pane, different breaks, margins,
  // print options, page setup, header and autofilter, and read it back → the sheet's settings as the
  // model holds them, beside the same package read without the patch → { withView, withoutView,
  // rewriteOk }. A saved view is Excel's own and is not the sheet: the two readings must match.
  customSheetViewReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getCell('A1').value = 'header';
    sheet.getCell('A2').value = 1;
    sheet.autoFilter = 'A1:A2';
    sheet.rowBreaks.push({id: 3, max: 16383, man: true});
    sheet.headerFooter.oddFooter = 'sheet footer';
    sheet.pageMargins.left = 0.5;
    const written = writeXlsx(wb);
    const views =
      '<customSheetViews><customSheetView guid="{C3D4E5F6-0000-4000-8000-000000000001}">' +
      '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>' +
      '<rowBreaks count="1" manualBreakCount="1"><brk id="9" max="16383" man="1"/></rowBreaks>' +
      '<pageMargins left="2" right="2" top="2" bottom="2" header="1" footer="1"/>' +
      '<printOptions gridLines="1"/><pageSetup orientation="landscape"/>' +
      '<headerFooter><oddHeader>view header</oddHeader></headerFooter><autoFilter ref="A1:B9"/>' +
      '</customSheetView></customSheetViews>';
    const settings = (workbook: WorkbookInstance) => {
      const s = workbook.getWorksheet('S')!;
      return {
        viewState: s.view.state ?? null,
        topLeftCell: s.view.topLeftCell ?? null,
        rowBreaks: s.rowBreaks.map((brk) => brk.id),
        pageMargins: {...s.pageMargins},
        printGridLines: s.printOptions.gridLines ?? null,
        orientation: s.pageSetup.orientation ?? null,
        headerFooter: {...s.headerFooter},
        autoFilterRef: s.autoFilter?.ref ?? null,
        a2: s.getCell('A2').value,
      };
    };
    const withView = reloadPatched(written, {
      // `CT_Worksheet` puts `<customSheetViews>` directly after `<autoFilter>`.
      'xl/worksheets/sheet1.xml': (xml) => xml.replace(/(<autoFilter\b[^>]*\/>)/, `$1${views}`),
    });
    let rewriteOk = true;
    try {
      writeXlsx(withView);
    } catch {
      rewriteOk = false;
    }
    return {withView: settings(withView), withoutView: settings(readXlsx(written)), rewriteOk};
  },

  // Read a fixture whose `<headerFooter>` children hold `_xHHHH_` escapes → { eager, roundtrip }, each a
  // map of the six header/footer slots to the text the model carries. `eager` is the fixture as read;
  // `roundtrip` is that model written back through our own writer and re-read, which is what holds the
  // escape and the decode to being inverses rather than two independently plausible transformations.
  headerFooterEscapeReport(rel: string) {
    const textOf = (workbook: WorkbookInstance) => {
      const hf = workbook.worksheets[0]?.headerFooter;
      return Object.fromEntries(HEADER_FOOTER_SLOTS.map((slot) => [slot, hf?.[slot] ?? null]));
    };
    const source = readFixture(rel);
    return {eager: textOf(source), roundtrip: textOf(roundtrip(source))};
  },

  // Author `text` as a sheet's odd header, write, and read back → { emitted, rawInPart, read }.
  // `emitted` is the `<oddHeader>` body exactly as it reached the part, `rawInPart` names any emitted
  // part still carrying a character XML cannot hold, which would make the package malformed, and
  // `read` is what the reader gives back.
  authoredHeaderFooterEscape(text: string) {
    // oxlint-disable-next-line eslint/no-control-regex -- matching the control characters is the check: this asks whether an emitted part carries one
    const raw = /[\u{0}-\u{8}\u{B}\u{C}\u{E}-\u{1F}\u{FFFE}\u{FFFF}\u{D800}-\u{DFFF}]/u;
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getCell('A1').value = 'x';
    sheet.headerFooter.oddHeader = text;
    const bytes = writeXlsx(workbook);
    const parts = partMapOf(bytes);
    const sheetPart = Object.entries(parts).find(([name]) =>
      /xl\/worksheets\/sheet\d+\.xml$/.test(name),
    )?.[1];
    return {
      emitted: /<oddHeader>([\s\S]*?)<\/oddHeader>/.exec(sheetPart ?? '')?.[1] ?? null,
      rawInPart: Object.keys(parts)
        .filter((name) => raw.test(parts[name] ?? ''))
        .sort(),
      read: readXlsx(bytes).worksheets[0]?.headerFooter.oddHeader ?? null,
    };
  },
  // Build a merge whose covered cell was materialised BEFORE the merge, then report what a range over
  // the whole merged block sees and what clearing its style does → { addresses, numFmtsAfterClear }.
  // `Worksheet.getCell` resolves a covered address to the merge master, so a walk that tested
  // `hasCell(row, col)` and then fetched by address got the master once per covered position and never
  // saw the covered cell: `cells` reported the master twice, and `clearStyle` cleared it twice and left
  // the covered cell styled, the opposite of what it documents.
  rangeOverMergeReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getCell('B2').value = 'master';
    const covered = sheet.getCell('C2');
    covered.value = 'covered';
    covered.numFmt = '0.00';
    sheet.getCell('B2').numFmt = '0.00';
    sheet.mergeCells('B2:C2');

    const range = sheet.getRange('B2:C2');
    const addresses = range.cells.map((cell) => cell.address);
    range.clearStyle();

    // Read back off the sheet's own iteration, which is merge-blind, so a covered cell left styled is
    // visible rather than hidden behind the same resolution that caused the bug.
    const numFmtsAfterClear: Record<string, string | null> = {};
    for (const {cells} of sheet.rows()) {
      for (const cell of cells) numFmtsAfterClear[cell.address] = cell.numFmt ?? null;
    }
    return {addresses, numFmtsAfterClear};
  },
};

// The six `<headerFooter>` children, in CT_HeaderFooter order: the slots a header/footer report walks.
const HEADER_FOOTER_SLOTS = [
  'oddHeader',
  'oddFooter',
  'evenHeader',
  'evenFooter',
  'firstHeader',
  'firstFooter',
] as const;
