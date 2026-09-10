// Fonts, fills, borders, alignment, number formats and the style-deduplication that decides
// which of them survive a round-trip.

import {canonicalJson} from '../../canonical-json.ts';
import {messageOf} from '../../thrown.ts';
import type {Untyped} from '../../untyped.ts';
import {partMapOf, roundtrip} from './package-facts.ts';
import {
  type CellInstance,
  fixtureBytes,
  NAMED_STYLE_ID,
  readFixture,
  readXlsx,
  Workbook,
  type WorkbookInstance,
  type WorksheetInstance,
  writeXlsx,
} from './runtime.ts';
import {applyStyle, buildFrom} from './spec-model.ts';
import {reloadPatched} from './xml-probes.ts';

export const styles = {
  // Resolve a workbook's default font, optionally over a fixture, optionally after authoring a theme
  // body face and/or a default font, write it, and report what the package now says about font id 0
  // and the theme it claims to follow → { declared, resolved, font0, fontCount, themeMinor,
  // agreesWithTheme, reReadDeclared, cellFonts }.
  //
  // Font 0 is a claim spanning two parts: `<scheme val="minor"/>` says "I am the theme's body face",
  // while `<name>` says which face that is. The facts are therefore reported from the *written*
  // package with both parts read back, because the failure this guards is precisely the two
  // disagreeing while each part is individually well-formed.
  defaultFontReport({fixture = null, themeFonts = null, defaultFont = null, cells = []}: Untyped) {
    const workbook = fixture === null ? new Workbook() : readFixture(fixture);
    const declared = workbook.declaredDefaultFont ?? null;
    if (fixture === null) workbook.addWorksheet('S').getCell('A1').value = 'plain';
    const sheet = workbook.worksheets[0]!;
    // Each entry styles one cell, so a case can state what a cell said about its own font and check
    // which cells an authored default reaches.
    for (const {address, font = null, fill = null} of cells as Untyped[]) {
      const cell = sheet.getCell(address);
      cell.value = address;
      if (font !== null) cell.font = font;
      if (fill !== null) cell.fill = fill;
    }
    if (themeFonts !== null) workbook.setTheme({fonts: themeFonts});
    if (defaultFont !== null) workbook.setDefaultFont(defaultFont);
    const resolved = {...workbook.defaultFont};

    const buffer = writeXlsx(workbook);
    const parts = partMapOf(buffer);
    const stylesXml = parts['xl/styles.xml'] ?? '';
    const fontsBlock = /<fonts\b[^>]*>[\s\S]*?<\/fonts>/.exec(stylesXml)?.[0] ?? '';
    const font0 = /<font>[\s\S]*?<\/font>/.exec(fontsBlock)?.[0] ?? '';
    const themePath = Object.keys(parts).find((p) => /^xl\/theme\/theme\d+\.xml$/.test(p));
    const theme = themePath === undefined ? '' : (parts[themePath] ?? '');
    const themeMinor = /<a:minorFont>\s*<a:latin[^>]*\btypeface="([^"]*)"/.exec(theme)?.[1] ?? null;
    const font0Name = /<name val="([^"]*)"/.exec(font0)?.[1] ?? null;
    const reread = readXlsx(buffer);
    const rereadSheet = reread.worksheets[0];
    return {
      declared,
      resolved,
      font0,
      font0Name,
      font0Scheme: /<scheme val="([^"]*)"/.exec(font0)?.[1] ?? null,
      // A count above one after a plain re-write is the redundant-duplicate failure: the declared
      // default replaced by an assumed one and re-added beside it as a custom entry.
      fontCount: [...fontsBlock.matchAll(/<font>/g)].length,
      themeMinor,
      // The claim `scheme="minor"` makes, checked rather than trusted.
      agreesWithTheme: font0Name !== null && font0Name === themeMinor,
      reReadDeclared: reread.declaredDefaultFont ?? null,
      // What each styled cell resolves to after the round-trip, so a case can tell a cell that named
      // a face from one that merely inherited the file's default.
      cellFonts: Object.fromEntries(
        (cells as Untyped[]).map(({address}) => [
          address,
          rereadSheet?.getCell(address).font?.name ?? null,
        ]),
      ),
    };
  },

  // Style rectangular blocks through the range handle, write, and report what the package shows →
  // { cellXfs, styledCells, distinctStyleIds, styledAddresses, reloadedStyles, sharedFacet,
  //   materialisedInBlock, outsideBlock }.
  //
  // Two claims a per-cell loop makes silently and a range must keep: an empty cell inside the block
  // renders styled (so it has to be materialised, not skipped), and a uniformly styled block still
  // collapses to one shared cellXfs entry rather than one per cell.
  rangeStyleReport({blocks = [], values = [], probe = null}: Untyped) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    for (const {address, value} of values as Untyped[]) sheet.getCell(address).value = value;
    for (const {ref, style = null, facet = null, value = null} of blocks as Untyped[]) {
      const range = sheet.getRange(ref);
      if (style !== null) range.style = style;
      if (facet !== null) (range as Untyped)[facet] = value;
    }
    const sharedFacet =
      probe === null ? null : ((sheet.getRange(probe.ref) as Untyped)[probe.facet] ?? null);
    const materialisedInBlock = probe === null ? null : sheet.getRange(probe.ref).cells.length;

    const buffer = writeXlsx(workbook);
    const parts = partMapOf(buffer);
    const stylesXml = parts['xl/styles.xml'] ?? '';
    const sheetXml = parts['xl/worksheets/sheet1.xml'] ?? '';
    const styled = [...sheetXml.matchAll(/<c r="([A-Z]+\d+)"[^>]*\ss="(\d+)"/g)];
    const reread = readXlsx(buffer).worksheets[0];
    return {
      cellXfs: Number(/<cellXfs count="(\d+)"/.exec(stylesXml)?.[1] ?? 0),
      styledCells: styled.length,
      // A uniform block must not mint one entry per cell: the historical performance cliff, and the
      // reason the style table is interned at all.
      distinctStyleIds: new Set(styled.map((m) => m[2])).size,
      styledAddresses: styled.map((m) => m[1]),
      sharedFacet,
      materialisedInBlock,
      // The facets survive the round-trip on a cell that was empty when it was styled.
      //
      // The value is `Untyped` because the expression reports whichever facet happens to be set
      // rather than the one `probe.facet` asked for, so its static type is `Fill | Font | null` and no
      // caller can know which it got. Returning the probed facet instead would be the honest shape,
      // but it would also change what a probe for `font` reports on a cell that carries a fill,
      // a behavior change that does not belong in a typing pass.
      reloadedStyles: Object.fromEntries<Untyped>(
        (probe === null ? [] : [...sheet.getRange(probe.ref).addresses()]).map((address) => [
          address,
          reread?.getCell(address).fill ?? reread?.getCell(address).font ?? null,
        ]),
      ),
      outsideBlock: probe === null ? null : (reread?.getCell(probe.outside).fill ?? null),
    };
  },

  // Assign a valid format-code string (alongside font/alignment/protection) to one cell and a
  // structured OBJECT (a parsed `{id, formatCode}` shape) to another, write, and report the string's
  // survival plus whether the styles part was corrupted → { controlNumFmtReload, stylesHasObjectObject }.
  // The object must never serialize as `formatCode="[object Object]"`.
  // Write a custom number format carrying every character an XML attribute has to escape, and report
  // the emitted `formatCode` attribute alongside what reads back ->
  // { attribute, readBack }. XML 1.0 3.3.3 makes a conforming parser normalise a raw tab, line feed
  // or carriage return in *any* attribute value to a space, so a format code that carries one raw is
  // a format code Excel reads back different from the one that was written. Our own reader preserves
  // the raw character, which is exactly why the loss cannot be seen from a round-trip alone: the
  // assertion has to be on the bytes.
  formatCodeAttributeEscaping() {
    const code = '0.0"a\tb\nc\rd"e&f<g>h\'i';
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getCell('A1').value = 1;
    sheet.getCell('A1').numFmt = code;
    const bytes = writeXlsx(workbook);
    const styles = partMapOf(bytes)['xl/styles.xml'] ?? '';
    return {
      attribute: /<numFmt numFmtId="\d+" formatCode="([^"]*)"\/>/.exec(styles)?.[1] ?? null,
      readBack: readXlsx(bytes).getWorksheet('S')?.getCell('A1').numFmt ?? null,
      authored: code,
    };
  },

  numFmtObjectCorruptionReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    const control = sheet.getCell('A1');
    control.value = 45000;
    control.numFmt = 'yyyy-mmm-dd';
    control.font = {bold: true};
    control.alignment = {horizontal: 'center'};
    control.protection = {locked: false};
    const bad = sheet.getCell('A2');
    bad.value = 42;
    // A caller wrongly assigns the structured numFmt object Excel parses a cell's format into.
    bad.numFmt = {id: 164, formatCode: '0.00'} as Untyped;
    const stylesXml = partMapOf(writeXlsx(wb))['xl/styles.xml'] || '';
    const back = roundtrip(wb).getWorksheet('S')!;
    return {
      controlNumFmtReload: back.getCell('A1').numFmt ?? null,
      stylesHasObjectObject: stylesXml.includes('[object Object]'),
    };
  },

  // Mark a cell whose content looks like a formula with the quote-prefix flag, write, and report
  // whether its cell-format record carries quotePrefix and whether the flag survives a round-trip →
  // { writtenQuotePrefix, reloaded }. The flag forces formula-like content to be stored as literal text.
  quotePrefixReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    const cell = sheet.getCell('A1');
    cell.value = '=1+1';
    cell.quotePrefix = true;
    const buffer = writeXlsx(wb);
    const writtenQuotePrefix = /<xf\b[^>]*quotePrefix="1"/.test(
      partMapOf(buffer)['xl/styles.xml'] || '',
    );
    const reloaded = readXlsx(buffer).getWorksheet('S')!.getCell('A1').quotePrefix === true;
    return {writtenQuotePrefix, reloaded};
  },

  // Take a cell carrying every kind of formatting a cell can carry (a shared style facet, the
  // quote-prefix flag, a link to a named cell style), put it through each path that COPIES a cell (a
  // row splice, a column splice, duplicateRow, a model assignment onto another sheet), and report what
  // survives each → { spliceRows, spliceColumns, duplicateRow, modelCopy, written }, each
  // { value, numFmt, quotePrefix, namedStyle }. The named-style link is codec-only and has no public
  // accessor, so it is read back off the written package instead of off the model: `written` reports
  // whether the re-emitted xf still carries quotePrefix and an xfId. Use it to assert a structural edit
  // carries a cell's WHOLE formatting, not the six facets that happen to be in the shared style tuple.
  cellContentSurvivesCopies() {
    const seed = (sheet: WorksheetInstance, ref: string) => {
      const cell = sheet.getCell(ref);
      cell.value = '007';
      cell.numFmt = '0.00';
      cell.quotePrefix = true;
      (cell as Untyped)[NAMED_STYLE_ID] = 1;
    };
    const facts = (cell: Untyped) => ({
      value: cell.value as unknown,
      numFmt: cell.numFmt as unknown,
      quotePrefix: cell.quotePrefix === true,
      namedStyle: cell[NAMED_STYLE_ID] as unknown,
    });

    const wb = new Workbook();

    const rowSpliced = wb.addWorksheet('RowSplice');
    seed(rowSpliced, 'A1');
    rowSpliced.spliceRows(1, 0, ['inserted']);

    const columnSpliced = wb.addWorksheet('ColSplice');
    seed(columnSpliced, 'A1');
    columnSpliced.spliceColumns(1, 0, ['inserted']);

    const duplicated = wb.addWorksheet('Duplicate');
    seed(duplicated, 'A1');
    duplicated.duplicateRow(1, {count: 1, insert: true});

    const source = wb.addWorksheet('Source');
    seed(source, 'A1');
    const destination = wb.addWorksheet('Destination');
    destination.model = source.model;

    // The written package, so the two cell-only facets are checked where they actually land.
    const single = new Workbook();
    const only = single.addWorksheet('S');
    seed(only, 'A1');
    only.spliceRows(1, 0, ['inserted']);
    const styles = partMapOf(writeXlsx(single))['xl/styles.xml'] ?? '';
    const cellXfs = (styles.match(/<cellXfs[\s\S]*?<\/cellXfs>/) ?? [])[0] ?? '';

    return {
      spliceRows: facts(rowSpliced.getCell('A2') as Untyped),
      spliceColumns: facts(columnSpliced.getCell('B1') as Untyped),
      duplicateRow: facts(duplicated.getCell('A2') as Untyped),
      modelCopy: facts(destination.getCell('A1') as Untyped),
      written: {
        quotePrefix: /<xf\b[^>]*quotePrefix="1"/.test(cellXfs),
        xfId: /<xf\b[^>]*xfId="1"/.test(cellXfs),
      },
    };
  },

  // Load a fixture whose A1 fill lives only in a named cell style (cellXfs xfId → cellStyleXfs), and
  // report: the source's cellStyleXfs count, A1's resolved fill, and, after a load→save round-trip,
  // the re-emitted cellStyleXfs count and whether A1's cellXfs entry still links via xfId → so a case
  // asserts the named-style layer is honoured on read and preserved (with its link) on write.
  namedStyleFillReport(rel: string) {
    const countCellStyleXfs = (xml: string) =>
      Number((xml.match(/<cellStyleXfs count="(\d+)"/) || [])[1] || 0);
    const srcCellStyleXfsCount = countCellStyleXfs(
      partMapOf(fixtureBytes(rel))['xl/styles.xml'] || '',
    );

    const workbook = readXlsx(fixtureBytes(rel));
    const readFill = workbook.worksheets[0]?.getCell('A1').fill ?? null;

    const out = writeXlsx(workbook);
    const outParts = partMapOf(out);
    const outStyles = outParts['xl/styles.xml'] || '';
    const roundtripCellStyleXfsCount = countCellStyleXfs(outStyles);
    // Resolve A1's cellXfs entry via its `s` index and check it retains a non-zero xfId link.
    const sIndex = Number(
      (outParts['xl/worksheets/sheet1.xml']?.match(/<c r="A1"[^>]*\bs="(\d+)"/) || [])[1] || 0,
    );
    const cellXfsBlock =
      (outStyles.match(/<cellXfs count="\d+">([\s\S]*?)<\/cellXfs>/) || [])[1] || '';
    const cellXfEntries = cellXfsBlock.match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g) || [];
    const a1Xf = cellXfEntries[sIndex] || '';
    const xfIdMatch = a1Xf.match(/\bxfId="(\d+)"/);
    const roundtripCellHasXfIdLink = xfIdMatch !== null && Number(xfIdMatch[1]) > 0;

    return {srcCellStyleXfsCount, readFill, roundtripCellStyleXfsCount, roundtripCellHasXfIdLink};
  },

  // Write a value under a date number format and report the sheet XML's health → { ok, hasNaN,
  // hasInvalidDate, cellXml }. A string, a null (empty cell), or an Invalid Date under a date format
  // must never leak a bare "NaN" or "Invalid Date" into the cell value.
  dateNumFmtValueReport(kind: 'string' | 'null' | 'invalidDate') {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    const cell = sheet.getCell('A1');
    if (kind === 'string') cell.value = 'not a date';
    else if (kind === 'null') cell.value = null;
    else if (kind === 'invalidDate') cell.value = new Date(Number.NaN);
    cell.numFmt = 'yyyy-mm-dd';
    let ok = true;
    let sheetXml = '';
    try {
      sheetXml = partMapOf(writeXlsx(workbook))['xl/worksheets/sheet1.xml'] || '';
    } catch {
      ok = false;
    }
    const cellXml = (sheetXml.match(/<c r="A1"[\s\S]*?(?:\/>|<\/c>)/) || [])[0] ?? '';
    return {
      ok,
      hasNaN: /NaN/.test(sheetXml),
      hasInvalidDate: /Invalid Date/.test(sheetXml),
      cellXml,
    };
  },

  // Apply a font to each named cell, then read each requested cell's font back → { <ref>:
  // font|null }. Each cell owns its own font, so a font set on one cell is observable there
  // and nowhere else: the isolation this reports. In-memory, matching the contract: the
  // <fonts>-table write/read path is exercised by the io/xlsx unit tests.
  probeCellFonts({apply = [], read = []}: Untyped) {
    const sheet = new Workbook().addWorksheet('sheet');
    for (const {cell, font} of apply) sheet.getCell(cell).font = font;
    const fonts: Record<string, Untyped> = {};
    for (const address of read) fonts[address] = sheet.getCell(address).font ?? null;
    return JSON.parse(JSON.stringify(fonts));
  },

  // Write the spec and report the shared style table's size plus the index each requested cell
  // resolved to → { cellXfCount, indices: { <ref>: index|null } }. styles.xml is a SHARED table
  // referenced by index, so identically-styled cells must collapse to one <cellXfs> entry (one
  // shared index), dedup neither inflating to one entry per cell nor over-collapsing distinct
  // styles. A cell left at the default style carries no `s` and reports null.
  styleDedupReport(spec: Untyped, cells: string[] = []) {
    const parts = partMapOf(writeXlsx(buildFrom(spec)));
    const styles = parts['xl/styles.xml'] || '';
    const xfBlock = (styles.match(/<cellXfs\b[\s\S]*?<\/cellXfs>/) || [''])[0];
    const cellXfCount = (xfBlock.match(/<xf\b/g) || []).length;
    const sheetXml = parts['xl/worksheets/sheet1.xml'] || '';
    const indices: Record<string, Untyped> = {};
    for (const ref of cells) {
      const m = sheetXml.match(new RegExp(`<c\\b[^>]*\\br="${ref}"[^>]*\\bs="(\\d+)"`));
      indices[ref] = m ? Number(m[1]) : null;
    }
    return {cellXfCount, indices};
  },

  // Write one plain-valued, unformatted cell, round-trip it, and report the font it reads back →
  // { hasFont, fontName, fontSize }. An unstyled cell renders in the workbook default font, so it
  // must resolve a concrete name/size rather than reporting no font at all.
  unstyledCellFontReport() {
    const wb = new Workbook();
    wb.addWorksheet('S').getCell('A1').value = 'hello';
    const cell = roundtrip(wb).getWorksheet('S')!.getCell('A1');
    const font = cell.font || null;
    return {
      hasFont: !!font,
      fontName: font ? (font.name ?? null) : null,
      fontSize: font ? (font.size ?? null) : null,
    };
  },

  // Colour a sheet's tab with an 8-digit ARGB (alpha first) alongside an uncoloured sheet, then
  // report the round-trip → { tabColorArgbWritten, reReadArgb, uncoloredHasTab }. A set tab colour
  // must survive verbatim while a sheet with none must not gain a spurious one.
  tabColorRoundtrip() {
    const wb = new Workbook();
    const colored = wb.addWorksheet('Colored');
    colored.tabColor = {argb: 'FFFF0000'};
    colored.getCell('A1').value = 'x';
    const plain = wb.addWorksheet('Plain');
    plain.getCell('A1').value = 'y';
    const buffer = writeXlsx(wb);
    const sheetXml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const written = (sheetXml.match(/<tabColor\b[^>]*rgb="([^"]*)"/) || [null, null])[1];
    const reload = readXlsx(buffer);
    const coloredTab = reload.getWorksheet('Colored')!.tabColor;
    return {
      tabColorArgbWritten: written,
      reReadArgb: coloredTab?.argb || null,
      uncoloredHasTab: !!reload.getWorksheet('Plain')!.tabColor,
    };
  },

  // Write a solid fill twice, once with a clean bare ARGB, once with a CSS-habit '#'-prefixed one,
  // and report the emitted <fgColor rgb="..."> for each → { validRgb, hashRgb }. Both must serialize
  // as valid 8-hex-digit values; a '#'-prefixed input must be normalized, never passed through as a
  // malformed 9-character colour.
  fillArgbHashPrefixReport() {
    const emittedFgColor = (argb: string) => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 'x';
      ws.getCell('A1').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb}};
      const stylesXml = partMapOf(writeXlsx(wb))['xl/styles.xml'] || '';
      return (stylesXml.match(/<fgColor rgb="([^"]*)"/) || [null, null])[1];
    };
    return {validRgb: emittedFgColor('FFBFBFBF'), hashRgb: emittedFgColor('#FFBFBFBF')};
  },

  // Author a solid fill with a 6-hex RGB (no alpha) and with a malformed value, and report how the
  // writer treats each → { sixHexRgb, rejectsMalformed }. A 6-hex RGB is the common "colour without
  // its alpha channel" case: it must be promoted to a valid opaque 8-hex ARGB, not emitted as a
  // 6-char rgb that Excel renders black. A value that is neither 6 nor 8 hex digits is a programming
  // error and must be rejected, never written as a colour Excel silently renders black.
  argbNormalizationReport() {
    const emittedFgColor = (argb: string) => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 'x';
      ws.getCell('A1').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb}};
      const stylesXml = partMapOf(writeXlsx(wb))['xl/styles.xml'] || '';
      return (stylesXml.match(/<fgColor rgb="([^"]*)"/) || [null, null])[1];
    };
    let rejectsMalformed = false;
    try {
      emittedFgColor('12345');
    } catch {
      rejectsMalformed = true;
    }
    return {sixHexRgb: emittedFgColor('00FF00'), rejectsMalformed};
  },

  outlinePropertiesRoundtrip() {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.outline.summaryBelow = false;
    ws.outline.summaryRight = false;
    ws.getCell('A1').value = 'x';
    const buffer = writeXlsx(wb);
    const sheetXml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const outlinePr = (sheetXml.match(/<outlinePr\b[^>]*\/?>/) || [''])[0];
    const reload = readXlsx(buffer);
    const outline = reload.getWorksheet('S')!.outline;
    return {
      outlinePrEmitted: /summaryBelow="0"/.test(outlinePr) && /summaryRight="0"/.test(outlinePr),
      reReadSummaryBelow: outline.summaryBelow ?? null,
      reReadSummaryRight: outline.summaryRight ?? null,
    };
  },

  // Author a bold cell, then rewrite the emitted <b/> flag to each explicit form and report how
  // the reader reads bold back → { bareTag, valOne, valZero }. A boolean font flag's `val` governs:
  // a bare tag or val="1" is on, val="0" is off; presence alone must not force true.
  fontExplicitFalseBoldReport() {
    const readBoldWith = (tag: string) => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 'x';
      ws.getCell('A1').font = {bold: true};
      // The bare tag is the control, and it is what the writer already emits: nothing to patch.
      const written = writeXlsx(wb);
      const reloaded =
        tag === '<b/>'
          ? readXlsx(written)
          : reloadPatched(written, {'xl/styles.xml': (xml) => xml.replace(/<b ?\/>/, tag)});
      const font = reloaded.getWorksheet('S')!.getCell('A1').font;
      return !!font?.bold;
    };
    return {
      bareTag: readBoldWith('<b/>'),
      valOne: readBoldWith('<b val="1"/>'),
      valZero: readBoldWith('<b val="0"/>'),
    };
  },

  // Author cells with italic/strike/underline on, rewrite each flag to its explicit-off form, and
  // report what the reader reads back → { italic, strike, underline }. val="0" turns a boolean flag
  // off; <u val="none"/> is the ABSENCE of an underline, so it must read back falsy, never the
  // truthy string "none".
  fontExplicitOffFlagsReport() {
    const readWith = (baseFont: Untyped, tagRe: RegExp, tag: string, field: string) => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 'x';
      ws.getCell('A1').font = baseFont;
      const reloaded = reloadPatched(writeXlsx(wb), {
        'xl/styles.xml': (xml) => xml.replace(tagRe, tag),
      });
      const font = reloaded.getWorksheet('S')!.getCell('A1').font || {};
      return (font as Record<string, Untyped>)[field] ?? null;
    };
    return {
      italic: readWith({italic: true}, /<i ?\/>/, '<i val="0"/>', 'italic'),
      strike: readWith({strike: true}, /<strike ?\/>/, '<strike val="0"/>', 'strike'),
      underline: readWith({underline: true}, /<u ?\/>/, '<u val="none"/>', 'underline'),
    };
  },

  // Inject an xf whose alignment element carries only an explicit-false boolean (wrapText="0" /
  // shrinkToFit="0"), point A1 at it, and report the alignment the reader surfaces → { wrapTextZero,
  // shrinkZero }. An all-false alignment carries no information and must read back as no alignment
  // at all: the raw "0" is a truthy JS string, so a reader guarding on the raw value rather than the
  // parsed boolean would wrongly report a present alignment.
  alignmentFalseBooleanReport() {
    const readWithAlignment = (alignAttr: Untyped) => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 'x';
      let injectedIndex = -1;
      const reloaded = reloadPatched(writeXlsx(wb), {
        'xl/styles.xml': (xml) => {
          const count = Number(xml.match(/<cellXfs count="(\d+)">/)![1]);
          injectedIndex = count;
          return xml
            .replace(/<cellXfs count="\d+">/, `<cellXfs count="${count + 1}">`)
            .replace(
              /<\/cellXfs>/,
              `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">` +
                `<alignment ${alignAttr}/></xf></cellXfs>`,
            );
        },
        'xl/worksheets/sheet1.xml': (xml) =>
          xml.replace(/<c r="A1"([^>]*)>/, `<c r="A1" s="${injectedIndex}"$1>`),
      });
      const alignment = reloaded.getWorksheet('S')!.getCell('A1').alignment;
      return alignment ? JSON.parse(JSON.stringify(alignment)) : null;
    };
    return {
      wrapTextZero: readWithAlignment('wrapText="0"'),
      shrinkZero: readWithAlignment('shrinkToFit="0"'),
    };
  },

  // Assign a vertical-alignment enum token to a cell, routed through the untyped adapter surface so a
  // value TypeScript would reject at compile time (e.g. the ExcelJS-era "middle") reaches the writer as
  // an untyped caller would smuggle it, then attempt a full write→read cycle → { writeThrew,
  // writeError, readBackVertical }. Captures the writer-boundary contract: a value the writer accepts
  // must survive read-back, so an out-of-contract token must be rejected at write or preserved, never
  // silently serialized into a schema-invalid file the library's own reader then discards.
  alignmentVerticalEnumReport(vertical: Untyped) {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    const cell = sheet.getCell('A1');
    cell.value = 'x';
    cell.alignment = {vertical};
    let writeThrew = false;
    let writeError: string | null = null;
    let readBackVertical: string | null = null;
    try {
      const back = roundtrip(wb).getWorksheet('S')!.getCell('A1').alignment;
      readBackVertical = back?.vertical ?? null;
    } catch (e) {
      writeThrew = true;
      writeError = messageOf(e);
    }
    return {writeThrew, writeError, readBackVertical};
  },

  // Give one cell a fill and another a border but NO value, leave a third cell entirely untouched,
  // round-trip, and report each → { filledArgb, filledValue, borderedStyle, borderedValue, untouched }.
  // A formatted-but-empty cell is a real cell Excel keeps: its style must survive the write and the
  // cell must stay value-less, while a cell with neither value nor style must not be fabricated.
  styledEmptyCellReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getCell('A1').value = 'anchor';
    sheet.getCell('B2').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF00FF00'}};
    sheet.getCell('C3').border = {top: {style: 'thin', color: {argb: 'FF000000'}}};
    sheet.getCell('D4'); // materialised but never given a value or style
    const back = roundtrip(wb).getWorksheet('S')!;
    const filled = back.getCell('B2');
    const bordered = back.getCell('C3');
    // fgColor lives on the pattern-fill variant, past the general Fill union surface.
    const filledFill = filled.fill as Untyped;
    return {
      filledArgb: filledFill?.fgColor ? filledFill.fgColor.argb : null,
      filledValue: filled.value,
      borderedStyle: bordered.border?.top ? bordered.border.top.style : null,
      borderedValue: bordered.value,
      untouched: !!(back.getCell('D4').fill || back.getCell('D4').value),
    };
  },

  readFixtureCellStyles(rel: string, cells: Untyped = []) {
    const wb = readFixture(rel);
    const out: Record<string, Untyped> = {};
    for (const key of cells) {
      const [sheetName, addr] = key.split('!');
      const sheet = wb.getWorksheet(sheetName);
      const cell = sheet ? sheet.getCell(addr) : null;
      out[key] = cell
        ? {
            fill: cell.fill?.type ? JSON.parse(JSON.stringify(cell.fill)) : null,
            fontColor: cell.font?.color ? JSON.parse(JSON.stringify(cell.font.color)) : null,
          }
        : null;
    }
    return out;
  },

  // Read a real fixture, write it straight back out, read the result, and report whether any
  // cell's visible fill colour or border-edge colours changed across that no-op round-trip →
  // { checked, fillMismatches, borderMismatches, fillSample, borderSample }. Theme+tint and
  // indexed-palette references must survive verbatim so the sheet renders identically after a
  // pure open-then-save. Colour comparison is key-order-insensitive.
  roundtripFixtureColorFidelity(rel: string) {
    const before = readFixture(rel);
    const after = roundtrip(before);

    const realFill = (cell: CellInstance) =>
      cell.fill && cell.fill.type === 'pattern' && cell.fill.pattern !== 'none' ? cell.fill : null;
    const borderColors = (cell: CellInstance) => {
      if (!cell.border) return null;
      const edges: Record<string, Untyped> = {};
      // `as const` so each name is a literal key of `Border` rather than a bare `string`, which
      // cannot index it, and the edge is bound once, so the colour read is narrowed by the guard
      // above it instead of being a second, independently-unnarrowed lookup.
      for (const edge of ['top', 'left', 'right', 'bottom'] as const) {
        const side = cell.border[edge];
        if (side?.color) edges[edge] = side.color;
      }
      return Object.keys(edges).length ? edges : null;
    };
    const norm = (v: unknown) => JSON.stringify(canonicalJson(v ?? null));

    let checked = 0;
    let fillMismatches = 0;
    let borderMismatches = 0;
    let fillSample = null;
    let borderSample = null;
    for (const sheet of before.worksheets) {
      const other = after.getWorksheet(sheet.name);
      for (const {cells} of sheet.rows()) {
        for (const cell of cells) {
          if (!realFill(cell) && !borderColors(cell)) continue;
          checked += 1;
          const oc = other ? other.getCell(cell.address) : null;
          const bf = norm(realFill(cell));
          const af = oc ? norm(realFill(oc)) : '(missing)';
          if (bf !== af) {
            fillMismatches += 1;
            if (!fillSample)
              fillSample = {cell: `${sheet.name}!${cell.address}`, before: bf, after: af};
          }
          const bb = norm(borderColors(cell));
          const ab = oc ? norm(borderColors(oc)) : '(missing)';
          if (bb !== ab) {
            borderMismatches += 1;
            if (!borderSample)
              borderSample = {cell: `${sheet.name}!${cell.address}`, before: bb, after: ab};
          }
        }
      }
    }
    return {checked, fillMismatches, borderMismatches, fillSample, borderSample};
  },

  // Give one column a right border and later columns only a width, then round-trip and report
  // each cell's right border → { a1, b1, c1 }. A column's border is a default for its own cells,
  // so the declaring column's cell carries it while columns without a style of their own get
  // nothing: column styles are independent, not bled into subsequent columns.
  columnBorderScopedReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getColumn(1).border = {right: {style: 'thin', color: {argb: 'FF000000'}}};
    sheet.getColumn(2).width = 10;
    sheet.getCell('A1').value = 'a';
    sheet.getCell('B1').value = 'b';
    sheet.getCell('C1').value = 'c';
    const s = roundtrip(wb).getWorksheet('S')!;
    const rightBorder = (ref: string) => {
      const b = s.getCell(ref).border;
      return !!b?.right?.style;
    };
    return {a1: rightBorder('A1'), b1: rightBorder('B1'), c1: rightBorder('C1')};
  },

  // Read a fixture, extract its column widths and pageSetup, write it straight back, re-read, and
  // report the same facts → { source, rewritten }. A faithful no-op round-trip must reproduce every
  // fractional column width and the print-scaling attributes the real file carries.
  roundtripFixtureStyleFacts(rel: string) {
    // Model-level facts (column widths, pageSetup, dxfs) come from the parsed workbook; the custom
    // indexed-color palette is a raw styles.xml fact, so it is read straight from the part bytes on
    // each side, matching the legacy oracle, which extracts the same block from the zip.
    const facts = (workbook: WorkbookInstance, stylesXml: string) => {
      const sheet = workbook.worksheets[0];
      const ps = sheet ? sheet.pageSetup : {};
      // Differential styles are preserved as verbatim `<dxf>` fragments; a rule's number format is
      // whatever formatCode the fragment carries, so a coerced "[object Object]" can never appear.
      const dxfs = workbook.differentialStyles;
      const dxfFormatCodes = dxfs.flatMap((f: Untyped) =>
        [...f.matchAll(/formatCode="([^"]*)"/g)].map((m) => m[1]),
      );
      const palette = stylesXml.match(/<indexedColors>([\s\S]*?)<\/indexedColors>/);
      return {
        columnWidths: sheet
          ? [...sheet.columns()].map((c) => c.properties?.width).filter((w) => w !== undefined)
          : [],
        pageSetup: {
          scale: ps.scale ?? null,
          fitToWidth: ps.fitToWidth ?? null,
          fitToHeight: ps.fitToHeight ?? null,
          pageOrder: ps.pageOrder ?? null,
          orientation: ps.orientation ?? null,
          paperSize: ps.paperSize ?? null,
        },
        dxfCount: dxfs.length,
        dxfFormatCodes,
        hasIndexedColors: !!palette,
        indexedColorSample: palette
          ? [...(palette[1] ?? '').matchAll(/rgb="([0-9a-fA-F]+)"/g)].slice(0, 6).map((m) => m[1])
          : [],
      };
    };
    const before = readFixture(rel);
    const source = facts(before, partMapOf(fixtureBytes(rel))['xl/styles.xml'] || '');
    const buffer = writeXlsx(before);
    const after = readXlsx(buffer);
    return {source, rewritten: facts(after, partMapOf(buffer)['xl/styles.xml'] || '')};
  },

  // Read a fixture, write it straight back, and report the theme part on each side →
  // { source, rewritten }, each { name, colors, majorFont, minorFont, relTargets, relTargetsResolve }.
  // The theme is what every `theme="n"` colour and `scheme="major|minor"` font in the package
  // resolves against, so a no-op round-trip that substitutes a different theme silently re-renders
  // the whole file. Read from the part bytes rather than through the model: the claim is about what
  // the package carries, and the model holds the theme opaquely by design.
  //
  // `relTargets` are the theme's own outbound relationship targets (a picture used as a themed fill)
  // and `relTargetsResolve` whether each one names a part the package actually holds: a theme
  // re-emitted without its closure would leave that `r:embed` dangling, which is worse than dropping
  // the theme outright.
  roundtripFixtureThemeFacts(rel: string) {
    // The theme is reached the way OPC reaches it, through the workbook's `.../theme` relationship,
    // whose target is relative to `xl/`, not by assuming the conventional `theme1.xml` name.
    const themePathOf = (parts: Record<string, string>) => {
      const rels = parts['xl/_rels/workbook.xml.rels'] ?? '';
      for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
        const tag = m[0];
        if (!/Type="[^"]*\/theme"/.test(tag)) continue;
        const target = /Target="([^"]*)"/.exec(tag)?.[1];
        if (target !== undefined) return `xl/${target.replace(/^\.?\//, '')}`;
      }
      return null;
    };
    const facts = (parts: Record<string, string>) => {
      const themePath = themePathOf(parts);
      const theme = themePath === null ? '' : (parts[themePath] ?? '');
      const scheme = (slot: string) =>
        new RegExp(`<a:${slot}>\\s*<a:(?:srgbClr|sysClr)[^>]*?(?:val|lastClr)="([^"]*)"`).exec(
          theme,
        )?.[1] ?? null;
      const face = (slot: string) =>
        new RegExp(`<a:${slot}>\\s*<a:latin typeface="([^"]*)"`).exec(theme)?.[1] ?? null;
      const relsPath = themePath === null ? null : themePath.replace(/([^/]+)$/, '_rels/$1.rels');
      const relsXml = relsPath === null ? '' : (parts[relsPath] ?? '');
      const relTargets = [...relsXml.matchAll(/Target="([^"]*)"/g)].map((m) => m[1] as string);
      return {
        path: themePath,
        present: themePath !== null && themePath in parts,
        name: /<a:theme[^>]*\sname="([^"]*)"/.exec(theme)?.[1] ?? null,
        colors: Object.fromEntries(
          ['dk2', 'lt2', 'accent1', 'accent2', 'accent6', 'hlink', 'folHlink'].map((slot) => [
            slot,
            scheme(slot),
          ]),
        ),
        majorFont: face('majorFont'),
        minorFont: face('minorFont'),
        relTargets,
        // A target is stated relative to the theme part's own directory, so `../media/x.png` on
        // `xl/theme/theme1.xml` resolves to `xl/media/x.png`.
        relTargetsResolve: relTargets.every((target) => {
          const segments = (themePath ?? '').split('/').slice(0, -1);
          for (const segment of target.split('/')) {
            if (segment === '..') segments.pop();
            else if (segment !== '.') segments.push(segment);
          }
          return segments.join('/') in parts;
        }),
      };
    };
    const bytes = fixtureBytes(rel);
    const rewritten = writeXlsx(readXlsx(bytes));
    return {source: facts(partMapOf(bytes)), rewritten: facts(partMapOf(rewritten))};
  },

  // Read a fixture, write it straight back, and report the tail blocks of styles.xml on each side →
  // { source, rewritten }. Those blocks, `<dxfs>`, `<tableStyles>`, `<colors>`, are the ones a
  // regenerating writer drops most easily, and they reference each other: a `tableStyleElement`'s
  // `dxfId` indexes the dxf table, and a table part's `tableStyleInfo/@name` names a `<tableStyle>`
  // by name. So the facts are reported *resolved*: `elementDxfs` is the dxf fragment each element's
  // `dxfId` actually lands on, and `tableStyleOnTableResolves` says whether the name the table asks
  // for is one the stylesheet still defines.
  //
  // `undeclaredPrefixes` guards the hazard of verbatim preservation: a fragment carries its namespace
  // prefixes with it, and one the re-emitted root never declares makes the whole part unparseable.
  roundtripFixtureStylesTailFacts(rel: string) {
    const facts = (parts: Record<string, string>) => {
      const xml = parts['xl/styles.xml'] ?? '';
      const block = (name: string) =>
        new RegExp(`<${name}\\b[^>]*/>|<${name}\\b[^>]*>[\\s\\S]*?</${name}>`).exec(xml)?.[0] ?? '';
      const tableStyles = block('tableStyles');
      const colors = block('colors');
      const dxfs = [...block('dxfs').matchAll(/<dxf\b[^>]*\/>|<dxf\b[^>]*>[\s\S]*?<\/dxf>/g)].map(
        (m) => m[0],
      );
      const elements = [...tableStyles.matchAll(/<tableStyleElement\b[^>]*\/>/g)].map((m) => {
        const tag = m[0];
        const dxfId = /\bdxfId="(\d+)"/.exec(tag)?.[1];
        return {
          type: /\btype="([^"]*)"/.exec(tag)?.[1] ?? null,
          dxfId: dxfId === undefined ? null : Number(dxfId),
        };
      });
      // The name the first table part asks its style by: the reference that dangles when the
      // definition is dropped.
      const tablePart = Object.keys(parts)
        .filter((p) => /^xl\/tables\/table\d+\.xml$/.test(p))
        .sort()[0];
      const nameOnTable =
        tablePart === undefined
          ? null
          : (/<tableStyleInfo\b[^>]*\bname="([^"]*)"/.exec(parts[tablePart] ?? '')?.[1] ?? null);
      const definedNames = [...tableStyles.matchAll(/<tableStyle\b[^>]*\bname="([^"]*)"/g)].map(
        (m) => m[1] as string,
      );
      const declared = new Set(
        [...xml.matchAll(/xmlns:([A-Za-z_][\w.-]*)\s*=/g)].map((m) => m[1] as string),
      );
      const usedPrefixes = new Set<string>();
      for (const tag of xml.matchAll(/<[^!?][^>]*>/g)) {
        for (const m of tag[0].matchAll(/[\s</]([A-Za-z_][\w.-]*):[A-Za-z_]/g)) {
          if (m[1] !== 'xmlns') usedPrefixes.add(m[1] as string);
        }
      }
      return {
        defaultTableStyle:
          /<tableStyles\b[^>]*\bdefaultTableStyle="([^"]*)"/.exec(xml)?.[1] ?? null,
        defaultPivotStyle:
          /<tableStyles\b[^>]*\bdefaultPivotStyle="([^"]*)"/.exec(xml)?.[1] ?? null,
        definedNames,
        elements,
        // Resolve each element's dxfId through the dxf table so a renumbered table shows up as a
        // changed fragment rather than as an unchanged index.
        elementDxfs: elements.map(({dxfId}) => (dxfId === null ? null : (dxfs[dxfId] ?? null))),
        dxfCount: dxfs.length,
        nameOnTable,
        tableStyleOnTableResolves: nameOnTable === null ? null : definedNames.includes(nameOnTable),
        mruColors: [...colors.matchAll(/<color\b[^>]*\brgb="([0-9a-fA-F]+)"/g)].map(
          (m) => m[1] as string,
        ),
        indexedColorCount: [...colors.matchAll(/<rgbColor\b/g)].length,
        // CT_Colors orders indexedColors before mruColors; a writer emitting them the other way round
        // produces a schema-invalid part.
        colorsChildren: [...colors.matchAll(/<(indexedColors|mruColors)\b/g)].map(
          (m) => m[1] as string,
        ),
        undeclaredPrefixes: [...usedPrefixes].filter((p) => !declared.has(p)),
      };
    };
    const bytes = fixtureBytes(rel);
    const rewritten = writeXlsx(readXlsx(bytes));
    return {source: facts(partMapOf(bytes)), rewritten: facts(partMapOf(rewritten))};
  },

  // Read a fixture and report, per requested cell, the raw colour encodings its fill/font carry
  // alongside what the workbook resolves each one to → { <ref>: {fill, fillResolved, font,
  // fontResolved}, themeColors }. An `indexed="n"` or `theme="n"` reference means nothing on its own;
  // resolution is what turns it into a colour a caller can render. The raw encoding is reported
  // beside the resolved value on purpose: resolution is a derived view, so the model must still be
  // holding the original reference.
  fixtureColorResolution(rel: string, cells: Untyped = []) {
    const workbook = readFixture(rel);
    const sheet = workbook.worksheets[0];
    const out: Record<string, Untyped> = {};
    for (const ref of cells) {
      const cell = sheet ? sheet.getCell(ref) : null;
      const fill = cell?.fill && cell.fill.type === 'pattern' ? (cell.fill.fgColor ?? null) : null;
      const font = cell?.font?.color ?? null;
      out[ref] = {
        fill: fill ? {...fill} : null,
        fillResolved: fill ? (workbook.resolveColor(fill) ?? null) : null,
        font: font ? {...font} : null,
        fontResolved: font ? (workbook.resolveColor(font) ?? null) : null,
      };
    }
    return {cells: out, themeColors: {...workbook.themeColors}};
  },

  // Resolve a colour reference against a workbook built from scratch → the 8-hex ARGB, or null. For
  // the claims that need no fixture: the built-in indexed palette, the system sentinels, tint.
  resolveColorOnEmptyWorkbook(color: Untyped) {
    return new Workbook().resolveColor(color) ?? null;
  },

  // Author theme colours/fonts on a workbook, from scratch, or over a fixture's own theme, write
  // it, and report what the emitted theme part carries → { scheme, fonts, schemeName, keptFmtScheme,
  // hasThemeRels, mediaParts, resolvedThemeColor, reReadScheme }. Authoring a palette has to reach
  // three places at once: the theme part, the cells that reference it by `theme="n"`, and whatever the
  // source theme already carried and must not lose.
  authorThemeReport({fixture = null, colors = {}, fonts = {}}: Untyped) {
    const workbook = fixture === null ? new Workbook() : readFixture(fixture);
    if (fixture === null) {
      const sheet = workbook.addWorksheet('Brand');
      const cell = sheet.getCell('A1');
      cell.value = 'themed';
      cell.fill = {type: 'pattern', pattern: 'solid', fgColor: {theme: 4}};
    }
    workbook.setTheme({colors, fonts});
    const buffer = writeXlsx(workbook);
    const parts = partMapOf(buffer);
    const theme = parts['xl/theme/theme1.xml'] ?? '';
    const clrScheme = /<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/.exec(theme)?.[0] ?? '';
    const slotValue = (slot: string) =>
      new RegExp(`<a:${slot}>\\s*<a:(?:srgbClr|sysClr)[^>]*?(?:val|lastClr)="([^"]*)"`).exec(
        clrScheme,
      )?.[1] ?? null;
    const face = (which: string) =>
      new RegExp(`<a:${which}>\\s*<a:latin typeface="([^"]*)"`).exec(theme)?.[1] ?? null;
    // A slot left unauthored must keep the *encoding* the source used, not just its value: dk1/lt1
    // are `<a:sysClr>` so they follow the viewer's window colours, and rewriting them as srgbClr
    // would pin them to one machine.
    const slotEncoding = (slot: string) =>
      new RegExp(`<a:${slot}>\\s*<a:(srgbClr|sysClr)`).exec(clrScheme)?.[1] ?? null;
    return {
      schemeName: /<a:clrScheme[^>]*\sname="([^"]*)"/.exec(theme)?.[1] ?? null,
      scheme: Object.fromEntries(
        ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent6', 'hlink'].map(
          (slot) => [slot, slotValue(slot)],
        ),
      ),
      encodings: {
        dk1: slotEncoding('dk1'),
        lt1: slotEncoding('lt1'),
        accent1: slotEncoding('accent1'),
      },
      fonts: {major: face('majorFont'), minor: face('minorFont')},
      // The format scheme is a designer's work, not something an API regenerates; it must ride
      // through an authored theme untouched.
      keptFmtScheme: /<a:fmtScheme\b/.test(theme),
      fmtSchemeGradientStops: [...theme.matchAll(/<a:gs\b/g)].length,
      hasThemeRels: 'xl/theme/_rels/theme1.xml.rels' in parts,
      mediaParts: Object.keys(parts).filter((p) => p.startsWith('xl/media/')).length,
      // The cell-facing half: a `theme="n"` reference must now resolve to the authored colour.
      resolvedThemeColor: workbook.resolveColor({theme: 4}) ?? null,
      // And the written package must say so too, not just the in-memory model.
      reReadScheme: {...readXlsx(buffer).themeColors},
    };
  },

  // Register custom table styles on a workbook (optionally one read from a fixture), point a table at
  // one of them, write, and report the cross-part wiring that has to hold → { definitions, elements,
  // elementDxfs, nameOnTable, resolves, dxfCount, styleCount }. The claim a table style makes spans
  // three parts: the table names a style, the styles part defines it, the dxf table backs each
  // element, so the facts are reported *resolved* rather than as raw indices.
  authorTableStyleReport({fixture = null, styles: authored = [], tableStyle = null}: Untyped) {
    const workbook = fixture === null ? new Workbook() : readFixture(fixture);
    if (fixture === null) {
      // A two-column, two-data-row table with no cell-level formatting at all, so anything the
      // written package says about its appearance can only have come from the style.
      const sheet = workbook.addWorksheet('Data');
      const rows = [
        ['Port', 'Tonnage'],
        ['Bilbao', 4120],
        ['Gdansk', 3380],
      ];
      rows.forEach((cells, r) => {
        cells.forEach((value, c) => {
          sheet.getCell(`${String.fromCharCode(65 + c)}${r + 1}`).value = value;
        });
      });
      sheet.addTable({
        name: 'Cargo',
        ref: 'A1',
        rowCount: 2,
        columns: [{name: 'Port'}, {name: 'Tonnage'}],
        // A table declares the style it wants when it is created; a fixture's table already names
        // its own, so `tableStyle` only applies to the from-scratch path.
        ...(tableStyle === null ? {} : {style: tableStyle}),
      });
    }
    for (const style of authored) workbook.addTableStyle(style);
    const parts = partMapOf(writeXlsx(workbook));
    const stylesXml = parts['xl/styles.xml'] ?? '';
    const block = (name: string) =>
      new RegExp(`<${name}\\b[^>]*/>|<${name}\\b[^>]*>[\\s\\S]*?</${name}>`).exec(stylesXml)?.[0] ??
      '';
    const tableStyles = block('tableStyles');
    const dxfs = [...block('dxfs').matchAll(/<dxf\b[^>]*\/>|<dxf\b[^>]*>[\s\S]*?<\/dxf>/g)].map(
      (m) => m[0],
    );
    const definitions = [...tableStyles.matchAll(/<tableStyle\b(?!s)[^>]*\bname="([^"]*)"/g)].map(
      (m) => m[1] as string,
    );
    const elements = [...tableStyles.matchAll(/<tableStyleElement\b[^>]*\/>/g)].map((m) => {
      const tag = m[0];
      const size = /\bsize="(\d+)"/.exec(tag)?.[1];
      const dxfId = /\bdxfId="(\d+)"/.exec(tag)?.[1];
      return {
        type: /\btype="([^"]*)"/.exec(tag)?.[1] ?? null,
        size: size === undefined ? null : Number(size),
        dxfId: dxfId === undefined ? null : Number(dxfId),
      };
    });
    const tablePart = Object.keys(parts)
      .filter((p) => /^xl\/tables\/table\d+\.xml$/.test(p))
      .sort()[0];
    const nameOnTable =
      tablePart === undefined
        ? null
        : (/<tableStyleInfo\b[^>]*\bname="([^"]*)"/.exec(parts[tablePart] ?? '')?.[1] ?? null);
    return {
      definitions,
      elements,
      // Each element's dxf, resolved: the half that makes the style actually paint anything.
      elementDxfs: elements.map(({dxfId}) => (dxfId === null ? null : (dxfs[dxfId] ?? null))),
      nameOnTable,
      resolves: nameOnTable === null ? null : definitions.includes(nameOnTable),
      declaredCount: Number(/<tableStyles\b[^>]*\bcount="(\d+)"/.exec(tableStyles)?.[1] ?? -1),
      dxfCount: dxfs.length,
    };
  },

  // Register a table style the library must refuse → the error message, or null if it was accepted.
  authorInvalidTableStyle(style: Untyped) {
    try {
      new Workbook().addTableStyle(style);
      return null;
    } catch (error) {
      return messageOf(error);
    }
  },

  // Author a theme colour the library must refuse → the error message, or null if it was accepted.
  // A malformed theme colour does not error in Excel: the slot renders as flat black.
  authorInvalidThemeColor(value: Untyped) {
    try {
      new Workbook().setTheme({colors: {accent1: value}});
      return null;
    } catch (error) {
      return messageOf(error);
    }
  },

  // Assign one base style object to two cells, then spread-reassign one cell's font color →
  // { a1Color, a2Color, bled }. The sibling given the same base must keep its original font.
  sharedBaseStyleFontMutation() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    const base = {font: {name: 'Arial', size: 11}};
    sheet.getCell('A1').value = 'YES';
    sheet.getCell('A2').value = 'NO';
    applyStyle(sheet.getCell('A1'), base);
    applyStyle(sheet.getCell('A2'), base);
    sheet.getCell('A1').font = {...sheet.getCell('A1').font, color: {argb: 'FF00FF00'}};
    const s = roundtrip(workbook).getWorksheet('S')!;
    const colorOf = (ref: string) => {
      const f = s.getCell(ref).font;
      return f?.color ? (f.color.argb ?? null) : null;
    };
    const a1Color = colorOf('A1');
    const a2Color = colorOf('A2');
    return {a1Color, a2Color, bled: a2Color === 'FF00FF00'};
  },

  // Author three cells sharing one style record, load, border ONE, round-trip →
  // { a1, a2, a3, bled }. Only the targeted cell may gain a border.
  loadMutateCellBorder() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    for (const r of [1, 2, 3]) {
      const c = sheet.getCell(`A${r}`);
      c.value = 'x';
      c.font = {bold: true};
    }
    const loaded = roundtrip(workbook);
    loaded.getWorksheet('S')!.getCell('A1').border = {
      top: {style: 'thin'},
      left: {style: 'thin'},
      bottom: {style: 'thin'},
      right: {style: 'thin'},
    };
    const s = roundtrip(loaded).getWorksheet('S')!;
    const hasBorder = (ref: string) => {
      const b = s.getCell(ref).border;
      return !!b?.top?.style;
    };
    return {
      a1: hasBorder('A1'),
      a2: hasBorder('A2'),
      a3: hasBorder('A3'),
      bled: hasBorder('A2') || hasBorder('A3'),
    };
  },

  // Author two cells with one shared fill, load, replace ONE cell's fill, read the sibling in
  // memory and after write-back → { sibling, mutatedTo, original, bled, diskSibling, diskBled }.
  loadMutateCellStyle({sharedFill = 'FFFF0000', mutateTo = 'FF00FF00'}: Untyped = {}) {
    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    s.getCell('A1').value = 'a';
    s.getCell('B1').value = 'b';
    const fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: sharedFill}} as Untyped;
    s.getCell('A1').fill = fill;
    s.getCell('B1').fill = fill; // identical formatting → one shared style index on disk
    // fgColor lives on the pattern-fill variant, past the general Fill union surface.
    const fgOf = (cell: CellInstance) =>
      (cell.fill as Untyped)?.fgColor ? ((cell.fill as Untyped).fgColor.argb ?? null) : null;

    const wb2 = roundtrip(wb);
    const s2 = wb2.getWorksheet('S')!;
    s2.getCell('A1').fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {argb: mutateTo},
    } as Untyped;
    const sibling = fgOf(s2.getCell('B1'));

    const diskSibling = fgOf(roundtrip(wb2).getWorksheet('S')!.getCell('B1'));
    return {
      sibling,
      mutatedTo: mutateTo,
      original: sharedFill,
      bled: sibling === mutateTo,
      diskSibling,
      diskBled: diskSibling === mutateTo,
    };
  },

  // Author two cells with one shared font, load, spread-reassign ONE cell's font color, read the
  // sibling → { edited, sibling, original, mutatedTo, bled }.
  loadMutateCellFont({original = 'FF000000', mutateTo = 'FFFF0000'}: Untyped = {}) {
    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    const font = {name: 'Arial', size: 12, color: {argb: original}};
    s.getCell('A1').value = 'a';
    s.getCell('A1').font = font;
    s.getCell('B1').value = 'b';
    s.getCell('B1').font = font; // identical formatting → one shared style index on disk

    const s2 = roundtrip(wb).getWorksheet('S')!;
    s2.getCell('A1').font = {...s2.getCell('A1').font, color: {argb: mutateTo}};
    const colorOf = (cell: CellInstance) =>
      cell.font?.color ? (cell.font.color.argb ?? null) : null;
    const sibling = colorOf(s2.getCell('B1'));
    return {
      edited: colorOf(s2.getCell('A1')),
      sibling,
      original,
      mutatedTo: mutateTo,
      bled: sibling === mutateTo,
    };
  },

  // Load two cells sharing one style record, set ONE style facet (alignment | numFmt | protection)
  // on one via its setter, and report whether it bled into the sibling in memory and on disk →
  // { facet, target, sibling, original, bled, diskSibling, diskBled }. The remaining facets of the
  // copy-on-write family, alongside fill/font/border above.
  loadMutateCellFacet(facet: Untyped = 'alignment') {
    const readFacet = (
      {
        alignment: (c: Untyped) => c.alignment?.horizontal || null,
        numFmt: (c: Untyped) => c.numFmt || null,
        protection: (c: Untyped) =>
          c.protection && typeof c.protection.locked === 'boolean' ? c.protection.locked : null,
      } as Record<string, Untyped>
    )[facet];
    const apply = (
      {
        alignment: (c: Untyped) => {
          c.alignment = {horizontal: 'center'};
        },
        numFmt: (c: Untyped) => {
          c.numFmt = '#,##0';
        },
        protection: (c: Untyped) => {
          c.protection = {locked: false};
        },
      } as Record<string, Untyped>
    )[facet];
    if (!readFacet) throw new Error(`unknown style facet: ${facet}`);

    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    s.getCell('A1').value = 'a';
    s.getCell('B1').value = 'b';
    const base = {numFmt: '0.00'}; // one identical non-default style → both cells dedup to one xf index
    applyStyle(s.getCell('A1'), base);
    applyStyle(s.getCell('B1'), base);

    const wb2 = roundtrip(wb);
    const s2 = wb2.getWorksheet('S')!;
    const original = readFacet(s2.getCell('B1'));
    apply(s2.getCell('A1'));
    const target = readFacet(s2.getCell('A1'));
    const sibling = readFacet(s2.getCell('B1'));

    const diskSibling = readFacet(roundtrip(wb2).getWorksheet('S')!.getCell('B1'));
    return {
      facet,
      target,
      sibling,
      original,
      bled: sibling !== original,
      diskSibling,
      diskBled: diskSibling !== original,
    };
  },

  // --- CSV (src/io/csv) -------------------------------------------------------------------------
  // The contract mirrors the oracle's ExcelJS-shaped options; here they translate onto the rewrite's
  // cleaner CsvReadOptions/CsvWriteOptions. A read yields a JSON-serializable 2-D array of typed
  // values (Date → { date: iso }, error → { error }, else the scalar or null), matching the oracle so
  // the same cases assert unchanged.
};
