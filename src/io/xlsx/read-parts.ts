// Everything hanging off a package's relationships: the parts a sheet or the workbook reaches, and the
// closure capture that carries verbatim whatever the model does not interpret.
//
// Split from `read.ts`, whose own header calls it "the orchestrator" and which was only that in its
// first two hundred lines. What discovery *is* is a separate question from what order a workbook is
// assembled in, and this is the half that answers it: given a part's relationships, which package part
// does each feature live in, and what comes back when it is read.

import type {CommentThread} from '../../core/comment-thread.ts';
import type {PictureProperties} from '../../core/image.ts';
import {INTERNAL} from '../../core/internal.ts';
import {mergesOverlappingTables} from '../../core/merge.ts';
import type {PreservedPart, PreservedWorksheetReference} from '../../core/preserved.ts';
import {Workbook} from '../../core/workbook.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {isAnyRelType} from '../../rel-type.ts';
import type {CollectingPass} from '../../xml/xml-read.ts';
import {localName} from '../../xml/xml-scan.ts';
import {relAttr} from '../opc/namespaces.ts';
import {extensionOf, resolveRelativePart} from '../opc/part-paths.ts';
import {
  capturePartClosure,
  type PackageAccessors,
  parseRelationshipRecords,
  type PartRelationships,
  readPartRelationships,
} from '../opc/read-opc.ts';
import {admitting} from '../read-policy/read-repair.ts';
import {type ParsedImageAnchor, parseDrawing} from './images.ts';
import {type ParsedComment, parseComments} from './read-comments.ts';
import {parsePivotTable} from './read-pivot.ts';
import {transitionalPart} from './strict-parts.ts';
import {parseTable} from './tables.ts';
import {parseThemeSchemes} from './theme-xml.ts';
import {buildCommentThreads, parsePersons, parseThreadedComments} from './threaded-comments.ts';

/**
 * What every part reader in one package read shares: the inflated package, the model being built, the
 * content-type resolver a preserved part's closure is captured through, and the interning map that
 * makes one workbook image per media part however many places point at it.
 *
 * None of the four varies within a read, and threading them positionally is what let `internImage`
 * end up taking the same three in a different order from both of its callers. What genuinely varies
 * -- a part's relationships, the sheet being filled -- stays a positional argument, and the sheet's
 * rels in particular are deliberately *not* in here: they are per sheet, and a stale copy of them
 * would send every sheet's parts to the first sheet's.
 */
export interface PackageReadContext {
  readonly pkg: PackageAccessors;
  readonly workbook: Workbook;
  readonly contentTypeOf: (path: string) => string;
  readonly imageIdByMediaPath: Map<string, number>;
}

// The closure of parts a preserved reference reaches, each translated out of ISO/IEC 29500 Strict as it
// is captured: the write that carries it is a Transitional package, and a part left in the other
// profile's names is one its consumers cannot place (see `strict-parts.ts`).
function capturePreservedClosure(
  {pkg, contentTypeOf}: PackageReadContext,
  entryPath: string,
): readonly PreservedPart[] | undefined {
  return capturePartClosure(entryPath, pkg.partText, pkg.partBytes, contentTypeOf)?.map((part) => {
    const bytes = transitionalPart(part.bytes, part.contentType);
    return bytes === part.bytes ? part : {...part, bytes};
  });
}

// A sheet's comments live in a comments part reached through the sheet's own relationships: the sheet
// declares a relationship of type `.../comments` whose target resolves (relative to the sheet's
// directory) to the comments part. A sheet declaring no such relationship simply has none.
export function readSheetComments(
  sheetRels: PartRelationships,
): Map<string, ParsedComment> | undefined {
  const commentsXml = sheetRels.relatedText('comments');
  return commentsXml === undefined ? undefined : parseComments(commentsXml);
}

// The workbook's threaded-comment identity registry: a relationship of type `.../person` names
// `xl/persons/person.xml`, whose entries every message's `personId` and every mention's
// `mentionpersonId` resolve through. A workbook with no threaded comments declares no such
// relationship and keeps an empty registry.
export function readWorkbookPersons(
  {workbook}: PackageReadContext,
  workbookRels: PartRelationships,
): void {
  const xml = workbookRels.relatedText('person');
  if (xml !== undefined) workbook[INTERNAL].restorePersons(parsePersons(xml));
}

// The workbook's theme part: the `<clrScheme>`/`<fontScheme>`/`<fmtScheme>` every `theme="n"` colour
// and every `scheme="major|minor"` font in the package resolves against. It is reached through the
// workbook's `.../theme` relationship rather than assumed at `xl/theme/theme1.xml`, because the target
// is rel-relative and a foreign package is free to name the part anything.
//
// Captured with its transitive part closure, not as a lone string: a theme can carry its own
// relationships (a picture used as a themed fill, wired by an `r:embed` into the theme's rels part),
// and re-emitting the theme body without them would leave that reference dangling, which Excel
// reports as a package needing repair. A package that declares no theme leaves the workbook on the
// library's default, which is also what a dangling relationship target degrades to.
export function readWorkbookTheme(
  context: PackageReadContext,
  workbookRels: PartRelationships,
): void {
  const {pkg, workbook} = context;
  const entryPath = workbookRels.targetPath('theme');
  if (entryPath === undefined) return;
  const parts = capturePreservedClosure(context, entryPath);
  if (parts === undefined) return;
  // The schemes are decoded here rather than on demand from the model: the part rides through the
  // model as opaque bytes, and only the codec knows how to read one.
  const xml = pkg.partText(entryPath) ?? '';
  workbook[INTERNAL].restoreThemePart({entryPath, parts}, parseThemeSchemes(xml));
}

// A sheet's threaded conversations live in a `xl/threadedComments/threadedComment{n}.xml` part reached
// through a relationship of type `.../threadedComment` on the sheet's own rels: the same discovery
// shape as the notes part above, and deliberately separate from it: a thread and a legacy note are
// different features that happen to share a sheet. The messages are grouped into threads and their
// authors resolved against the workbook registry, so each thread lands self-contained.
//
// What lands here IS what a re-write emits: the thread part is re-serialised from these threads, and so is
// the legacy fallback `<comment>` that binds each cell to its conversation (see `comments.ts`). Anything
// this reader drops is therefore dropped from the file, which is why a message too damaged to place is
// still kept wherever it can be, and why the anchor is canonicalised here rather than trusted downstream.
export function readSheetCommentThreads(
  {workbook}: PackageReadContext,
  sheetRels: PartRelationships,
): CommentThread[] {
  const xml = sheetRels.relatedText('threadedComment');
  if (xml === undefined) return [];
  return buildCommentThreads(parseThreadedComments(xml), (id) => workbook.getPerson(id));
}

// A sheet's printer-settings blob is an opaque binary part linked from `<pageSetup r:id>`: the sheet
// declares a relationship of type `.../printerSettings` whose target resolves to a `.bin` part. We
// keep the raw bytes verbatim: the DEVMODE inside is platform-specific and the model never
// interprets it, only round-trips it so re-writing the file preserves the user's print configuration.
// A sheet declaring no such relationship simply has none.
export function readSheetPrinterSettings(sheetRels: PartRelationships): Uint8Array | undefined {
  return sheetRels.relatedBytes('printerSettings');
}

// One workbook image per media part, however many places in the package point at that part. A sheet's
// background and a drawing's picture routinely name the same bytes, and modelling them as two images
// would write the media twice on the way out, so the map is the property both callers depend on and
// is what makes the re-write byte-count-stable. `undefined` means the part named a target the package
// does not carry, which is a broken relationship the caller skips over rather than fails on.
function internImage(context: PackageReadContext, mediaPath: string): number | undefined {
  const {workbook, pkg, imageIdByMediaPath} = context;
  const known = imageIdByMediaPath.get(mediaPath);
  if (known !== undefined) return known;
  const bytes = pkg.partBytes(mediaPath);
  if (bytes === undefined) return undefined;
  const id = workbook.addImage({buffer: bytes, extension: extensionOf(mediaPath)});
  imageIdByMediaPath.set(mediaPath, id);
  return id;
}

// A sheet's anchored images live in a drawing part reached through the sheet's own relationships: a
// relationship of type `.../drawing` names the drawing part, whose own relationships map each
// picture's embed id to a media part under `xl/media/`. Each anchor becomes a workbook image (deduped
// by media path) placed back on the sheet at its two-cell anchor.
export function readSheetImages(
  context: PackageReadContext,
  sheetRels: PartRelationships,
  sheet: Worksheet,
): void {
  const {partText} = context.pkg;
  const drawingPath = sheetRels.targetPath('drawing');
  if (drawingPath === undefined) return;
  const drawingXml = partText(drawingPath);
  if (drawingXml === undefined) return;
  const {anchors, fullyModeled} = parseDrawing(drawingXml);
  // A drawing holding anything the image model cannot write back (a chart, a shape, a picture anchored
  // to the page or linked rather than embedded) is preserved whole (see readSheetPreservedReferences),
  // so none of its pictures are modeled here: modeling them would leave the sheet with images, which
  // suppresses that preservation and drops the rest. Leaving `sheet.images` empty routes the entire
  // drawing, pictures included, through byte-preservation, keeping every anchor faithful.
  if (!fullyModeled) return;
  const drawingRels = readPartRelationships(drawingPath, partText);

  for (const anchor of anchors) {
    const embedded = drawingRels.byId(anchor.embed);
    if (embedded === undefined) continue;
    const mediaPath = drawingRels.pathOf(embedded.target);
    const id = internImage(context, mediaPath);
    if (id === undefined) continue;
    const rot = anchor.rotation !== undefined ? {rotation: anchor.rotation} : {};
    const properties = pictureOf(anchor, drawingRels);
    if (anchor.to !== undefined) {
      const mode = anchor.editAs !== undefined ? {editAs: anchor.editAs} : {};
      sheet.addImageAnchor(id, {from: anchor.from, to: anchor.to, ...mode, ...rot}, properties);
    } else if (anchor.ext !== undefined) {
      sheet.addImageAnchor(id, {from: anchor.from, ext: anchor.ext, ...rot}, properties);
    }
  }
}

// What a parsed picture says about itself, with its link's relationship resolved. A link whose relationship
// is missing, or that names nothing, is dropped rather than kept pointing nowhere. A relationship that is
// not external keeps its target as written, the `#Sheet1!C3` form Excel uses for a place in the workbook.
function pictureOf(anchor: ParsedImageAnchor, drawingRels: PartRelationships): PictureProperties {
  const {description, title, crop, hyperlinkId, tooltip} = anchor;
  const link = hyperlinkId === undefined ? undefined : drawingRels.byId(hyperlinkId);
  const target = link === undefined || link.target === '' ? undefined : link.target;
  return {
    ...(description === undefined ? {} : {description}),
    ...(title === undefined ? {} : {title}),
    ...(crop === undefined ? {} : {crop}),
    ...(target === undefined
      ? {}
      : {hyperlink: tooltip === undefined ? {target} : {target, tooltip}}),
  };
}

// A sheet background is a workbook image referenced by the worksheet's `<picture>` element through a
// sheet-local relationship of type `.../image`. Unlike an anchored image (whose image relationships
// live in the drawing part's own rels), the background's relationship sits directly on the sheet, so
// it is the sheet rels' sole image relationship. The bytes are deduped against images shared with a
// drawing, keeping one media part per picture across a re-write.
export function readSheetBackground(
  context: PackageReadContext,
  sheetRels: PartRelationships,
  sheet: Worksheet,
): void {
  const mediaPath = sheetRels.targetPath('image');
  if (mediaPath === undefined) return;
  const id = internImage(context, mediaPath);
  if (id !== undefined) sheet.addBackgroundImage(id);
}

// Capture the worksheet-level references to package content the model does not interpret, so a
// round-trip re-emits them verbatim instead of dropping them:
//   • `<drawing>`, but only when the reader modeled no anchored image from it: either a drawing that
//     holds no pictures at all (a chart or shape), or a mixed drawing whose pictures the reader
//     declined to model precisely so the whole part (chart included) rides here verbatim. A drawing
//     whose pictures were modeled is owned by the model and re-serialised from it; capturing it here
//     too would double-emit those pictures.
//   • `<legacyDrawingHF>`, a header/footer image's VML, which the model never interprets.
// Each reference's target part and the transitive closure of parts it reaches (a VML's image, a
// drawing's media) are captured with their bytes, content types, and relationships.
export function readSheetPreservedReferences(
  context: PackageReadContext,
  sheetRels: PartRelationships,
  referenceRelIds: WorksheetReferenceRelIds,
  sheet: Worksheet,
): void {
  const capture = (
    element: PreservedWorksheetReference['element'],
    relType: string,
    target: string,
  ): void => {
    const entryPath = sheetRels.pathOf(target);
    const parts = capturePreservedClosure(context, entryPath);
    if (parts !== undefined)
      sheet[INTERNAL].addPreservedReference({element, relType, entryPath, parts});
  };

  // Element-wired references: a `<drawing>`/`<legacyDrawingHF>` names its part by an `r:id` in the
  // sheet body. A `<drawing>` is preserved only when the reader modeled no picture from it (a
  // chart/shape-only drawing, or a mixed one the reader left unmodeled) since one whose pictures are
  // modeled is re-serialised from the model.
  const referenceElements: Array<'drawing' | 'legacyDrawingHF'> =
    sheet.images.length === 0 ? ['drawing', 'legacyDrawingHF'] : ['legacyDrawingHF'];
  for (const element of referenceElements) {
    const relId = referenceRelIds[element];
    const record = relId === undefined ? undefined : sheetRels.byId(relId);
    if (record !== undefined && !record.external) capture(element, record.type, record.target);
  }

  // Relationship-wired references: a pivot table or slicer is reached through a sheet relationship
  // with no worksheet child pointing at it; Excel discovers it by scanning the sheet's rels. Preserve
  // each so the pivots/slicers a fill-and-save workflow does not touch are not dropped.
  for (const record of sheetRels.records) {
    if (record.external) continue;
    if (isPreservedSheetRelType(record.type)) capture(undefined, record.type, record.target);
  }
}

// A sheet relationship the model does not consume but must round-trip: a pivot table or a slicer. Every
// other sheet rel kind (drawing, printerSettings, table, comments, threadedComment, hyperlinks, background
// image, the comment VML) is modeled and re-serialised from the model, so preserving it here would emit
// the part twice.
function isPreservedSheetRelType(type: string): boolean {
  return isAnyRelType(type, 'pivotTable', 'slicer');
}

// Capture the workbook-level references to package content the model does not interpret: pivot
// caches (`pivotCacheDefinition`), slicer caches (`slicerCache`), and external links (`externalLink`,
// each a link to a source workbook), so a round-trip re-emits them instead of dropping the pivots,
// slicers, and linked-workbook references they back. A pivot cache's `<pivotCaches>` registration (its
// `cacheId`) and an external link's `<externalReferences>` position (its `[n]` index) are captured
// alongside so the wiring a pivot table or a formula resolves through survives too.
export function readWorkbookPreservedReferences(
  context: PackageReadContext,
  workbookRels: PartRelationships,
  registrations: WorkbookRegistrations,
): void {
  const {workbook} = context;
  const {cacheIdByRelId, externalIndexByRelId} = registrations;
  for (const record of workbookRels.records) {
    if (record.external || !isPreservedWorkbookRelType(record.type)) continue;
    const entryPath = workbookRels.pathOf(record.target);
    const parts = capturePreservedClosure(context, entryPath);
    if (parts === undefined) continue;
    const cacheId = cacheIdByRelId.get(record.id);
    const externalReferenceIndex = externalIndexByRelId.get(record.id);
    workbook[INTERNAL].addPreservedReference({
      relType: record.type,
      entryPath,
      parts,
      ...(cacheId !== undefined ? {pivotCacheId: cacheId} : {}),
      ...(externalReferenceIndex !== undefined ? {externalReferenceIndex} : {}),
    });
  }
}

// Content wired from the package's own `_rels/.rels` that the writer does not regenerate from the
// model: the ribbon customUI parts, custom document properties, a thumbnail. The writer rebuilds the
// root rels for the parts it models (the workbook, and core/app properties), so every other root
// relationship's target would be dropped on write; capturing its closure here re-declares it verbatim.
// External targets and the three regenerated relationship types are skipped.
export function readRootPreservedReferences(context: PackageReadContext): void {
  const {workbook} = context;
  const relsXml = context.pkg.partText('_rels/.rels');
  if (relsXml === undefined) return;
  for (const record of parseRelationshipRecords(relsXml)) {
    if (record.external || isRegeneratedRootRelType(record.type)) continue;
    const entryPath = resolveRelativePart('', record.target);
    const parts = capturePreservedClosure(context, entryPath);
    if (parts === undefined) continue;
    workbook[INTERNAL].addPreservedRootReference({relType: record.type, entryPath, parts});
  }
}

// The three root relationships the writer regenerates from the model on every write: the office
// document and the core/extended document properties. Every other root relationship is unmodeled and
// is preserved verbatim by {@link readRootPreservedReferences} rather than dropped. A Strict package's
// `extendedProperties` arrives here already spelled `extended-properties` (see `io/opc/strict.ts`).
function isRegeneratedRootRelType(type: string): boolean {
  return isAnyRelType(type, 'officeDocument', 'core-properties', 'extended-properties');
}

// A workbook relationship the model does not consume but must round-trip: a pivot cache, a slicer
// cache, an external link (the pointer to a linked source workbook), or a macro-enabled workbook's VBA
// project. Worksheets, styles, theme, shared strings, and the threaded-comment person registry are modeled
// and re-serialised from the model. Preserving vbaProject here, rather than silently dropping it as an
// unrecognised relationship type otherwise would, is what keeps loading and re-saving a .xlsm from
// discarding its macros; the content-type override in workbook-xml.ts is the other half, so the re-emitted
// package still declares itself macro-enabled. Preserving externalLink is what keeps a formula's `[n]`
// external reference from dangling: the link part and its `<externalReferences>` registration are both
// re-emitted.
function isPreservedWorkbookRelType(type: string): boolean {
  return isAnyRelType(type, 'pivotCacheDefinition', 'slicerCache', 'vbaProject', 'externalLink');
}

/** The two `<pivotCaches>`/`<externalReferences>` registries a preserved reference is wired by,
 * gathered from the workbook part's own scan rather than from two more of it. */
export interface WorkbookRegistrations {
  readonly cacheIdByRelId: ReadonlyMap<string, string>;
  readonly externalIndexByRelId: ReadonlyMap<string, number>;
}

// Map each `<pivotCache>` registration in the workbook's `<pivotCaches>` to the relationship id that
// reaches its cache definition, so a preserved cache carries the `cacheId` a pivot table refers to.
export function pivotCacheRegistrationsPass(): CollectingPass<ReadonlyMap<string, string>> {
  const byRelId = new Map<string, string>();
  return {
    handlers: {
      onOpen(name, attrs, _selfClosing, scope) {
        if (localName(name) !== 'pivotCache') return;
        const relId = relAttr(scope, attrs, 'id');
        if (relId !== undefined && attrs.cacheId !== undefined) byRelId.set(relId, attrs.cacheId);
      },
    },
    result: () => byRelId,
  };
}

// Map each `<externalReference>` in the workbook's `<externalReferences>` to its 0-based position, keyed
// by the relationship id it wires. That position is the `[n]` index a formula or defined name resolves
// an external cell through (`[1]Sheet!$A$1`), so preserving it lets the writer re-emit the block in the
// original order and keep every `[n]` pointing at the same linked workbook.
export function externalReferenceRegistrationsPass(): CollectingPass<ReadonlyMap<string, number>> {
  const byRelId = new Map<string, number>();
  let index = 0;
  return {
    handlers: {
      onOpen(name, attrs, _selfClosing, scope) {
        if (localName(name) !== 'externalReference') return;
        const relId = relAttr(scope, attrs, 'id');
        if (relId !== undefined) byRelId.set(relId, index++);
      },
    },
    result: () => byRelId,
  };
}

/** The `r:id` each of the two element-wired references carries, or `undefined` where absent. */
export type WorksheetReferenceRelIds = Readonly<
  Partial<Record<'drawing' | 'legacyDrawingHF', string>>
>;

/**
 * A pass gathering the `r:id` of the first `<drawing>` and `<legacyDrawingHF>` in a worksheet.
 *
 * These two references live in the worksheet body rather than being distinguishable by relationship
 * Type, since a header/footer VML and a comment VML share the `vmlDrawing` type, so the specific
 * relationship is found by reading the element's own `r:id`.
 *
 * A pass rather than its own scan, because the sheet part is the largest in a package by a wide
 * margin and this ran a full extra scan of it, once or twice per sheet, for two attributes.
 * `readSheet` already builds a multi-pass single parse of it whose own comment records that five
 * separate scans "spent 45% of a large file's read on four scans that matched no element"; this was
 * the sixth scan, added later, that the same argument covers.
 */
export function worksheetReferencePass(): CollectingPass<WorksheetReferenceRelIds> {
  const relIds: {-readonly [K in keyof WorksheetReferenceRelIds]: string} = {};
  return {
    handlers: {
      onOpen(name, attrs, _selfClosing, scope) {
        const local = localName(name);
        if (local !== 'drawing' && local !== 'legacyDrawingHF') return;
        // The FIRST one wins, matching what a scan returning on its first hit did.
        if (relIds[local] !== undefined) return;
        const relId = relAttr(scope, attrs, 'id');
        if (relId !== undefined) relIds[local] = relId;
      },
    },
    result: () => relIds,
  };
}

// A sheet's tables live in `xl/tables/table{n}.xml` parts, each reached through a relationship of
// type `.../table` on the sheet's own rels. The writer emits one relationship per table; each part
// is parsed back into the model and re-registered in definition order. A part that fails to parse
// (missing name/ref/columns, which is Excel corruption) is skipped rather than crashing the whole read.
export function readSheetTables(
  {pkg}: PackageReadContext,
  sheetRels: PartRelationships,
  sheet: Worksheet,
): void {
  for (const tablePath of sheetRels.targetPaths('table')) {
    const tableXml = pkg.partText(tablePath);
    if (tableXml === undefined) continue;
    const options = parseTable(tableXml);
    // A table name is validated as an Excel identifier and bounded in length, and a file is free to
    // carry neither; the refusal is native (`SyntaxError`/`RangeError`), so an unguarded call put a
    // failure outside the `XlsxError` taxonomy on a path that faces untrusted input.
    if (options !== undefined) {
      admitting(() => {
        sheet.addTable(options);
      });
    }
  }
  dropMergesInsideTables(sheet);
}

// Reconstruct an inspectable model of each pivot table hosted on a sheet. A pivot is reached by a
// sheet relationship of type `.../pivotTable`; the pivot-table part carries its own relationship of
// type `.../pivotCacheDefinition` to the cache holding the field catalogue and source range. Both
// parts are parsed and combined into a read-only view registered on the sheet, separate from the
// byte-preservation that actually round-trips the pivot, so this never changes what is re-emitted.
// The read is lenient: a pivot whose cache is missing still yields a (partial) model rather than
// throwing, matching Excel's tolerance for a damaged package on load.
export function readSheetPivotTables(
  {pkg}: PackageReadContext,
  sheetRels: PartRelationships,
  sheet: Worksheet,
): void {
  const {partText} = pkg;
  for (const tablePath of sheetRels.targetPaths('pivotTable')) {
    const tableXml = partText(tablePath);
    if (tableXml === undefined) continue;
    const cacheXml =
      readPartRelationships(tablePath, partText).relatedText('pivotCacheDefinition') ?? '';
    sheet[INTERNAL].addLoadedPivotTable(parsePivotTable(tableXml, cacheXml));
  }
}

// Excel forbids a merged range inside a formatted table and repairs such a file on load by dropping
// the merge. A worksheet's merges are read before its tables, so a real file carrying that invalid
// geometry lands in the model intact; this applies the same repair once the tables are known, so a
// re-write does not surface the Excel-invalid geometry the writer (correctly) rejects.
function dropMergesInsideTables(sheet: Worksheet): void {
  // The overlapping merges come back as a list of their own, which is what lets this unmerge as it
  // goes: `sheet.merges` is the live backing array `unmergeCells` splices out of.
  for (const {merge} of mergesOverlappingTables(sheet.merges, sheet.tables)) {
    sheet.unmergeCells(merge);
  }
}
