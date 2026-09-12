// The buffered `.xlsx` reader: an OPC zip package in, a Workbook model out.
//
// An unrecognised construct is skipped rather than guessed, and content the model does not interpret
// is carried verbatim as preserved parts, so a foreign file reads without crashing and writes back
// with what it held.
//
// This module is the orchestrator, and now only that. It wires the parsed package parts together (the
// OPC/rel resolution in `../opc/read-opc.ts`, the style table in `./read-styles.ts`, each worksheet
// body in `./read-worksheet.ts`) and decides the order a workbook is assembled in, which is where the
// non-local constraints live.
//
// Two questions it used to answer itself moved out. *Which package part does a feature live in* is
// `./read-parts.ts`: sheet- and workbook-part discovery, and the closure capture that carries verbatim
// whatever the model does not interpret. *What does `xl/workbook.xml` say* is `./read-workbook-xml.ts`,
// beside `workbook-xml.ts` which writes exactly those elements, because this tree keeps both
// directions of one wire form together and the workbook part was the last place that was not true.
//
// Untrusted input: inflate is bounded by a running byte counter (`../opc/inflate.ts`) that caps
// actual decompressed output rather than trusting the archive's forgeable size headers, and
// the parser (ADR 0004) never expands entities.

import {INTERNAL} from '../../core/internal.ts';
import {Workbook} from '../../core/workbook.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {parseXmlPasses} from '../../xml/xml-read.ts';
import {UnsupportedFormatError} from '../opc/errors.ts';
import {
  contentTypeResolver,
  openSpreadsheetPackage,
  readPartRelationships,
} from '../opc/read-opc.ts';
import type {ReadPackageOptions} from '../opc/read-options.ts';
import {admitting, repairedSheetNames} from '../read-policy/read-repair.ts';
import type {XfStyle} from '../style/xf-style.ts';
import {readXlsbPackage} from '../xlsb/read.ts';
import {parseDynamicArrayCellMetadata} from './cell-metadata.ts';
import type {SharedString} from './cell-value.ts';
import {applyConditionalFormattings, conditionalFormattingPass} from './conditional-formatting.ts';
import {
  applyDataValidations,
  dataValidationPass,
  extendedDataValidationPass,
} from './data-validation.ts';
import {applyHyperlinks, sheetHyperlinkPass} from './hyperlinks.ts';
import {SHARED_STRINGS_PART, STYLES_PART} from './part-names.ts';
import {applyNotes} from './read-comments.ts';
import {
  type PackageReadContext,
  readRootPreservedReferences,
  readSheetBackground,
  readSheetComments,
  readSheetCommentThreads,
  readSheetImages,
  readSheetPivotTables,
  readSheetPreservedReferences,
  readSheetPrinterSettings,
  readSheetTables,
  readWorkbookPersons,
  externalReferenceRegistrationsPass,
  pivotCacheRegistrationsPass,
  readWorkbookPreservedReferences,
  readWorkbookTheme,
  worksheetReferencePass,
} from './read-parts.ts';
import {parseSharedStrings} from './read-shared-strings.ts';
import {parseStyleTable} from './read-styles.ts';
import {
  applyAppProperties,
  applyCoreProperties,
  definedNamesPass,
  workbookPropertiesPass,
  workbookProtectionPass,
  workbookSheetsPass,
  workbookViewPass,
} from './read-workbook-xml.ts';
import {worksheetPass} from './read-worksheet.ts';

/**
 * Read a spreadsheet package into a {@link Workbook}.
 *
 * Both OOXML serialisations are accepted: an XML `.xlsx`, and a binary `.xlsb` (BIFF12), which is the
 * same OPC container with binary office-document parts. The two are auto-detected from the package
 * itself rather than from a file extension, so a caller never branches on which form it holds, and
 * the model produced is the same either way. See `../xlsb/read.ts` for what the binary path does not
 * yet decode.
 *
 * @throws {UnsupportedFormatError} if the input is neither: a legacy `.xls` (`.format === 'xls'`) or
 *   an unrecognised/non-ZIP blob (`'unknown'`).
 * @throws {XlsbParseError} if a binary `.xlsb` part is malformed.
 * @throws {PackageReadError} if the input is a ZIP that cannot be unpacked: a corrupt or
 *   truncated archive, or one exceeding the inflate bound (a probable zip bomb).
 */
export function readXlsx(data: Uint8Array, options: ReadPackageOptions = {}): Workbook {
  const {files, pkg, documentPath, workbookXml} = openSpreadsheetPackage(
    data,
    options.maxUncompressedBytes,
  );
  const {partText, partBytes} = pkg;

  if (workbookXml === undefined) {
    // No XML office document. A binary one means this is an `.xlsb`, which reads through the BIFF12
    // codec over the very same model. The package is already inflated, so it is handed over as-is,
    // along with where its own relationship graph says the binary workbook lives.
    if (partBytes(documentPath) !== undefined) return readXlsbPackage(files, documentPath);
    throw new UnsupportedFormatError('unknown');
  }

  // A part's content type is needed to faithfully re-declare any part preserved verbatim for
  // round-tripping (a vector-shape drawing, a header/footer image and its VML). Resolve it the way
  // OPC does: an explicit `<Override>` for the exact part, else the `<Default>` for its extension.
  const contentTypeOf = contentTypeResolver(partText('[Content_Types].xml') ?? '');

  // One parse of the workbook's rels, queried by the sheet loop and by the two workbook-level part
  // readers below. It used to be held as a raw string and handed to three separate scanners, which
  // also left two different idioms for "reach a related part" side by side in one function.
  const workbookRels = readPartRelationships(documentPath, partText);
  const sharedStrings = parseSharedStrings(
    workbookRels.relatedTextOrPath('sharedStrings', SHARED_STRINGS_PART),
  );
  // The style table resolves a cell/row/column style index to its facets (fill, number
  // format); a package without one (a hand-rolled foreign file) yields an empty table and
  // every index reads as unstyled.
  const stylesXml = workbookRels.relatedTextOrPath('styles', STYLES_PART);
  const {cellXfs: xfStyles, namedStyles, defaultFont, preserved} = parseStyleTable(stylesXml);

  const workbook = new Workbook();
  // Everything the part readers below share for the whole of this read, built once and handed down.
  // The workbook-level readers run before the sheet loop and the sheet-level ones inside it, and all
  // of them resolve parts against the same package and intern media into the same map.
  const context = {
    pkg,
    workbook,
    contentTypeOf,
    sharedStrings,
    xfStyles,
    dynamicArrayCells: parseDynamicArrayCellMetadata(
      workbookRels.relatedText('sheetMetadata') ?? '',
    ),
    // A picture used on more than one sheet is one media part; caching by media path across the
    // whole loop keeps it a single workbook image so a re-write does not duplicate the bytes.
    imageIdByMediaPath: new Map<string, number>(),
  } satisfies Omit<SheetReadContext, 'definedNames'>;
  // The four sub-tables the stylesheet carries verbatim, all captured by the same read of the part
  // that resolved the xfs above rather than by four more scans of it.
  //
  // Preserve the differential-style table so conditional formatting's dxfId references stay valid,
  // and a foreign dxf's number format stays a real format code, across a re-write.
  workbook[INTERNAL].restoreDifferentialStyles([...preserved.dxfs]);
  // Preserve a custom indexed-color palette so an `indexed="…"` colour reference keeps its intended
  // RGB across a re-write instead of resolving to a different default-palette entry.
  workbook[INTERNAL].restoreIndexedColors([...preserved.indexedColors]);
  // Preserve the author's "Recent Colors" swatches, which the model never reads but re-writing would
  // otherwise discard.
  workbook[INTERNAL].restoreMruColors([...preserved.mruColors]);
  // Preserve the custom table-style definitions so a table referencing one by name still resolves to
  // a real definition after a re-write instead of rendering unstyled.
  workbook[INTERNAL].restoreTableStyles(preserved.tableStyles);
  // Preserve the theme part so a branded colour/font scheme is not overwritten by the default theme
  // the writer emits for a workbook that has none.
  readWorkbookTheme(context, workbookRels);
  // Preserve the named cell-style layer only when a file declares one beyond the Normal default, so an
  // ordinary workbook keeps an empty named-style table and emits just the default on write.
  if (namedStyles.length > 1) workbook[INTERNAL].restoreNamedStyles(namedStyles);
  // Preserve the declared default font (font id 0) so a re-write emits the face the file itself named
  // rather than an assumed Calibri, which would change every empty cell and the metric every
  // character-unit column width is expressed in.
  workbook[INTERNAL].restoreDefaultFont(defaultFont);
  const core = partText('docProps/core.xml');
  if (core !== undefined) applyCoreProperties(workbook, core);
  const app = partText('docProps/app.xml');
  if (app !== undefined) applyAppProperties(workbook, app);
  // Every reader of the workbook part, over one scan of it.
  //
  // The ordering that matters is unchanged and is now structural: all six see the whole part before
  // any sheet is read, so the date system `<workbookPr>` carries is in place before the first cell
  // decode. It used to be six scans, of which four matched no element each; the worksheet part had
  // been fixed for exactly this and the workbook part had got none of the treatment.
  const protection = workbookProtectionPass();
  const sheets = workbookSheetsPass();
  const definedNames = definedNamesPass();
  const pivotCaches = pivotCacheRegistrationsPass();
  const externalReferences = externalReferenceRegistrationsPass();
  parseXmlPasses(workbookXml, [
    protection,
    workbookPropertiesPass(workbook),
    workbookViewPass(workbook.view),
    sheets,
    definedNames,
    pivotCaches,
    externalReferences,
  ]);
  workbook.protection = protection.result();
  // The threaded-comment author registry is workbook-level, and every conversation on every sheet
  // resolves its authors and @mentions through it, so it is restored before the sheet loop that reads
  // those conversations, not alongside the other workbook-level parts below.
  readWorkbookPersons(context, workbookRels);
  // The names are known from the workbook part alone, which is what lets a sheet's formulas be read
  // against them although the names themselves are registered only after the sheets.
  const sheetContext: SheetReadContext = {...context, definedNames: definedNames.spellings()};

  const sheetOrder: string[] = [];
  // The name is repaired rather than trusted. `addWorksheet` refuses an empty, over-long, duplicate
  // or forbidden-character name, and a file is free to carry all four; refusing there would report a
  // corrupt package as the caller's mistake, and dropping the sheet would take its cells, its place
  // in the order, and every `localSheetId` that indexes past it. `sheetOrder` therefore carries the
  // name the model ended up with, which is what a scoped defined name has to resolve against.
  for (const {name, relId, state} of repairedSheetNames(sheets.result())) {
    const target = workbookRels.byId(relId)?.target;
    const sheet = workbook.addWorksheet(name, state === undefined ? undefined : {state});
    sheetOrder.push(sheet.name);
    readSheet(sheet, target === undefined ? undefined : workbookRels.pathOf(target), sheetContext);
  }

  readWorkbookPreservedReferences(context, workbookRels, {
    cacheIdByRelId: pivotCaches.result(),
    externalIndexByRelId: externalReferences.result(),
  });
  readRootPreservedReferences(context);

  // Defined names follow the sheets: a scoped name's `localSheetId` indexes the sheet order, which
  // is why the names are read only once every sheet is registered.
  for (const name of definedNames.result(sheetOrder)) {
    admitting(() => {
      workbook.defineName(name);
    });
  }
  return workbook;
}

/**
 * Everything a single sheet needs from the package around it: the package-wide read state, plus the
 * two tables only a *sheet* body decodes against. A cell's `t="s"` indexes the pool and its `s=`
 * indexes the xfs. Both are resolved before the workbook part is even scanned, so the whole read
 * shares one object, and {@link readSheet} takes it rather than a run of positional arguments.
 */
interface SheetReadContext extends PackageReadContext {
  readonly sharedStrings: readonly SharedString[];
  readonly xfStyles: readonly XfStyle[];
  /** The `cm` values the workbook's cell metadata marks as dynamic arrays, for the sheet's formulas. */
  readonly dynamicArrayCells: ReadonlySet<number>;
  /** Every name the workbook defines, as `definedNameKeys` spells them, for the sheet's formulas. */
  readonly definedNames: ReadonlySet<string>;
}

/**
 * Read one worksheet at `path`: its body, the four overlays that ride the same parse of the worksheet
 * part, and every part hanging off the sheet's own relationships.
 *
 * The stages are ordered, not merely sequential, and each constraint is non-local:
 *
 * - the overlays are gathered during the body's parse but *applied* only once the sheet's rels are in
 *   hand, because a hyperlink resolves its target through them;
 * - threads land before notes, because a threaded cell's comments-part entry is that thread's legacy
 *   fallback rather than a note, and `applyNotes` reads the restored threads to tell the two apart;
 * - preserved references are captured after the images, because that capture excludes what the image
 *   reader already modelled and would otherwise re-emit a drawing the writer also emits.
 *
 * Defined names are deliberately *not* read here: a sheet-scoped name indexes the workbook's sheet
 * order, so `readXlsx` reads them only once every sheet is registered.
 *
 * A sheet whose relationship is dangling (`path === undefined`) stays an empty sheet in its place in
 * the order rather than vanishing from the workbook.
 */
function readSheet(sheet: Worksheet, path: string | undefined, context: SheetReadContext): void {
  const {pkg, workbook, sharedStrings, xfStyles, definedNames, dynamicArrayCells} = context;
  const {partText} = pkg;
  const sheetXml = path === undefined ? undefined : partText(path);

  // Six readers want the worksheet part, and it is the largest in the package by a wide margin, so
  // they share one parse of it rather than scanning it once each. Only the body commits as it goes;
  // the other five gather, and are applied below in the order they were always applied.
  const hyperlinks = sheetHyperlinkPass();
  const validations = dataValidationPass(definedNames);
  const extendedValidations = extendedDataValidationPass(definedNames);
  const formattings = conditionalFormattingPass(definedNames, (fragment) =>
    workbook[INTERNAL].adoptDifferentialStyle(fragment),
  );
  const references = worksheetReferencePass();
  if (sheetXml !== undefined) {
    parseXmlPasses(sheetXml, [
      worksheetPass(
        sheet,
        sharedStrings,
        xfStyles,
        workbook.dateEpoch,
        definedNames,
        dynamicArrayCells,
      ),
      hyperlinks,
      validations,
      extendedValidations,
      formattings,
      references,
    ]);
  }
  if (path === undefined) return;

  // The sheet's rels are the index to nearly every part hanging off it, so they are parsed once here
  // and threaded through the readers below rather than re-read by each.
  const sheetRels = readPartRelationships(path, partText, pkg.partBytes);
  if (sheetXml !== undefined) {
    applyHyperlinks(sheet, hyperlinks.result(), (id) => sheetRels.byId(id)?.target);
    applyDataValidations(sheet, [...validations.result(), ...extendedValidations.result()]);
    applyConditionalFormattings(sheet, formattings.result());
  }

  const threads = readSheetCommentThreads(context, sheetRels);
  if (threads.length > 0) sheet[INTERNAL].restoreCommentThreads(threads);
  const comments = readSheetComments(sheetRels);
  if (comments !== undefined) applyNotes(sheet, comments);

  readSheetImages(context, sheetRels, sheet);
  readSheetBackground(context, sheetRels, sheet);
  if (sheetXml !== undefined) {
    readSheetPreservedReferences(context, sheetRels, references.result(), sheet);
  }

  readSheetTables(context, sheetRels, sheet, definedNames);
  readSheetPivotTables(context, sheetRels, sheet);
  const printerSettings = readSheetPrinterSettings(sheetRels);
  if (printerSettings !== undefined) sheet.pageSetup.printerSettings = printerSettings;
}
