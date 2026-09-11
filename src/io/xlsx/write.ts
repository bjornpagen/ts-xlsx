// The buffered `.xlsx` writer: a Workbook model in, an OPC zip package out.
//
// A value the writer cannot represent faithfully is refused rather than emitted as a lossy or corrupt
// package, and content the model does not interpret is carried through verbatim.
//
// This module is the orchestrator: it plans the package graph (via `package-plan.ts`) and
// stitches the serialised parts (from `workbook-xml.ts` and `worksheet-xml.ts`) into the
// part map. The row/cell renderer the streaming writer shares lives in `row-xml.ts`.

import {strToU8, zip, zipSync} from 'fflate';

import type {Person} from '../../core/comment-thread.ts';
import {formulaNamesInScope} from '../../core/formula.ts';
import {pictureProperties} from '../../core/image.ts';
import type {Workbook} from '../../core/workbook.ts';
import {refuseImagesBesideKeptDrawing, type Worksheet} from '../../core/worksheet.ts';
import {AuthoringError, InternalError, quoted} from '../../errors.ts';
import {isRelType} from '../../rel-type.ts';
import {relsPathFor, THEME_PART_PATH} from '../opc/part-paths.ts';
import {relsPartXml} from '../opc/rels.ts';
import {FIXED_ENTRY_MTIME} from '../opc/zip-mtime.ts';
import {collectComments, commentsXml, vmlDrawingXml} from './comments.ts';
import {collectHyperlinks, liveCells, planHyperlinks} from './hyperlinks.ts';
import {type DrawingImage, drawingXml} from './images.ts';
import {
  type BackgroundPlan,
  type CommentPlan,
  type DrawingPlan,
  type MediaPlan,
  type PivotPlan,
  type PlannedRelationship,
  type PreservedPartPlan,
  type PreservedPlan,
  type PreservedReferencePlan,
  type PreservedRootReferencePlan,
  type SheetPlan,
  type PreservedWorkbookReferencePlan,
  type PrinterSettingsPlan,
  planMedia,
  planPreservedParts,
  RelationshipLedger,
  type TablePlan,
  type ThreadedCommentPlan,
} from './package-plan.ts';
import {
  APP_PROPS_PART,
  commentsPart,
  CORE_PROPS_PART,
  drawingPart,
  mediaPart,
  PERSONS_PART,
  pivotCacheDefinitionPart,
  pivotCacheRecordsPart,
  pivotTablePart,
  printerSettingsPart,
  SHARED_STRINGS_PART,
  STYLES_PART,
  tablePart,
  threadedCommentsPart,
  vmlDrawingPart,
  WORKBOOK_PART,
  worksheetPart,
} from './part-names.ts';
import {pivotCacheDefinitionXml, pivotCacheRecordsXml, pivotTableXml} from './pivot.ts';
import {REL} from './relationships.ts';
import type {FlushedSheet} from './row-xml.ts';
import {SharedStringTable} from './shared-strings.ts';
import {StyleRegistry} from './styles.ts';
import {tableXml} from './tables.ts';
import {applyThemeOverrides, DEFAULT_THEME_XML} from './theme-xml.ts';
import {personsXml, threadedCommentsXml} from './threaded-comments.ts';
import {
  appPropsXml,
  contentTypesXml,
  corePropsXml,
  type WorkbookRelPlan,
  workbookXml,
} from './workbook-xml.ts';
import {type SheetReferences, worksheetXml} from './worksheet-xml.ts';

/** Options controlling how {@link writeXlsx} serialises a workbook. */
export interface WriteOptions {
  /**
   * Pool string and rich-text cell values into a shared-strings table (`xl/sharedStrings.xml`) that
   * cells reference by index, rather than storing each inline in its cell. Deduplicates repeated text
   * and matches Excel's own storage; a rich value is pooled as a `<si>` of its runs, so its formatting
   * survives. Off by default, which keeps strings inline and omits the part.
   */
  readonly useSharedStrings?: boolean;
}

/**
 * {@link WriteOptions} plus the streaming writer's internal wiring, so a buffered caller's options
 * object can never carry fields meant only for {@link WorkbookStreamWriter}'s own use. Not exported
 * from the public barrel: {@link WorkbookStreamWriter} is the only caller that fills these fields,
 * and it reaches them through {@link buildPackageParts}.
 */
export interface InternalWriteOptions extends WriteOptions {
  /**
   * The style registry to intern into, in place of a freshly-seeded one. The streaming writer
   * serialises each committed row eagerly (freeing its cells), so those rows' style ids must be
   * assigned by the very same registry that later emits `xl/styles.xml`; otherwise the ids in the
   * pre-rendered rows would not match the styles part. When omitted the buffered path seeds its own,
   * so its output is unchanged.
   */
  readonly styles?: StyleRegistry;

  /**
   * Per-sheet rows already serialised and evicted from the model by the streaming writer, keyed by
   * the model worksheet. Their XML is emitted ahead of the sheet's remaining live rows and their
   * extent folds into `<dimension>`. Absent for the buffered path, which holds every row live.
   */
  readonly flushed?: ReadonlyMap<Worksheet, FlushedSheet>;
}

/**
 * Serialise a workbook into an `.xlsx` package.
 *
 * The bytes are a pure function of the workbook: an unchanged model written twice produces two
 * identical archives, because entry timestamps are pinned to a fixed date rather than taken from the
 * clock. A committed `.xlsx` therefore only changes when something about it changed.
 *
 * @throws {AuthoringError} if the workbook has no worksheets (a zero-sheet package is corrupt),
 *   or holds a value the writer cannot yet represent.
 */
export function writeXlsx(workbook: Workbook, options: WriteOptions = {}): Uint8Array {
  return zipSync(buildPackageParts(workbook, options), {level: 6, mtime: FIXED_ENTRY_MTIME});
}

/**
 * Serialise a workbook into an `.xlsx` package, deflating off the calling thread.
 *
 * Produces the same package {@link writeXlsx} does, byte for byte and entry timestamps included, and
 * exists for one reason: DEFLATE dominates the cost of writing a large workbook, and {@link writeXlsx}
 * spends all of it on the caller's thread. Here `fflate` deflates each part in a worker, so the event
 * loop keeps turning (stalls drop from the whole write to tens of milliseconds) and parts compress in
 * parallel, which on a multi-sheet workbook also finishes sooner. On a single-sheet workbook there is
 * only one part to deflate, so expect responsiveness rather than speed.
 *
 * Building the parts still happens on the calling thread; only compression moves. That is why there
 * is no `readXlsxAsync` mirroring this: reading is dominated by XML parsing and model building, which
 * no worker can take, and the reader's zip-bomb ceiling is enforced by counting output between
 * synchronous input slices. See ADR-0024.
 *
 * @throws {AuthoringError}, as a rejection, under the same conditions as {@link writeXlsx};
 *   the part-building it shares happens before any worker is involved. A failure raised by the zip
 *   layer itself (including an environment that cannot spawn a worker) propagates unwrapped, exactly
 *   as it does from {@link writeXlsx}.
 */
export async function writeXlsxAsync(
  workbook: Workbook,
  options: WriteOptions = {},
): Promise<Uint8Array> {
  const parts = buildPackageParts(workbook, options);
  return await new Promise<Uint8Array>((resolve, reject) => {
    zip(parts, {level: 6, mtime: FIXED_ENTRY_MTIME}, (error, data) => {
      if (error) reject(error);
      else resolve(data);
    });
  });
}

/**
 * A style registry seeded from a workbook's read-in style layers (differential styles, named cell
 * styles, custom indexed palette), ready to intern authored styles after them. Both the buffered
 * pass and the streaming writer build their registry through here so a cell's style id means the
 * same thing whichever writer emits it.
 */
export function createStyleRegistry(workbook: Workbook): StyleRegistry {
  // Font id 0 is the workbook's own default face, resolved from what it declared, what was authored,
  // and its theme's body typeface. Never an assumed Calibri, which would re-face every empty cell
  // and change the metric every character-unit column width is expressed in.
  const styles = new StyleRegistry({
    defaultFont: workbook.defaultFont,
    ...(workbook.declaredDefaultFont === undefined
      ? {}
      : {declaredDefaultFont: workbook.declaredDefaultFont}),
  });
  // Seed the differential-style table with the fragments read from a source file so conditional
  // formatting's dxfId references stay valid; styles authored on rules append after them.
  styles.seedDifferentialStyles(workbook.differentialStyles);
  // Seed the named cell-style layer (cellStyleXfs/cellStyles) so each style's facets re-intern into the
  // rebuilt sub-tables and a cell's xfId link stays valid; without any, the default Normal alone emits.
  styles.seedNamedStyles(workbook.namedStyles);
  // Seed the custom indexed-color palette so it re-emits verbatim and an `indexed="…"` colour keeps
  // its intended RGB; a workbook that never overrode the palette seeds nothing and writes no <colors>.
  styles.seedIndexedColors(workbook.indexedColors);
  // Seed the author's "Recent Colors" swatches so they re-emit unchanged rather than being reset.
  styles.seedMruColors(workbook.mruColors);
  // Seed the custom table-style definitions so a table's `styleName` still names a real definition and
  // each element's dxfId still indexes the differential-style table seeded above at its original index.
  styles.seedTableStyles(workbook.tableStyles);
  // Authored styles append after the preserved ones, and intern their elements' formatting after the
  // seeded dxfs: the ordering that keeps every preserved dxfId pointing where it did.
  for (const style of workbook.customTableStyles) styles.addTableStyle(style);
  return styles;
}

// The workbook-global part counters the per-sheet planning advances: tables, drawings and pivot
// tables are numbered across the whole package, not per sheet. Passed as one mutable object so the
// shared state is part of {@link planSheet}'s signature rather than three variables it closes over.
interface PartNumbering {
  table: number;
  drawing: number;
  pivot: number;
}

/**
 * Plan one sheet's parts, recording every sheet-local relationship in that sheet's own ledger as its
 * id is taken, in the one canonical order the package wires them: tables, drawing, comments (VML +
 * comments part), threaded comments, printer settings, external hyperlinks, background, preserved
 * references, pivot tables.
 *
 * That order decides the ids, and the ledger is what keeps them gapless and collision-free: no step
 * re-derives its offset by summing the ones before it. A step moved above another here renumbers a
 * package, which the determinism tests and the corpus would notice, but no longer breaks one, because
 * the `.rels` part is the ledger rather than a second listing that has to agree with it.
 */
function planSheet(context: {
  readonly sheet: Worksheet;
  readonly index: number;
  readonly media: MediaPlan;
  readonly preserved: PreservedPlan;
  readonly numbering: PartNumbering;
  // What the streaming writer already flushed for this sheet, whose cells are gone from the model.
  readonly flushed?: FlushedSheet | undefined;
}): SheetPlan {
  const {sheet, index, media, preserved, numbering, flushed} = context;
  const rels = new RelationshipLedger(worksheetPart(index + 1));

  const tables: TablePlan[] = sheet.tables.map((table) => {
    const number = ++numbering.table;
    return {table, number, relId: rels.add(REL.table, tablePart(number))};
  });

  let drawing: DrawingPlan | null = null;
  if (sheet.images.length > 0) {
    // The authoring doors refuse this already; every write path comes through here, so this is the
    // guarantee that a kept drawing is never left unreferenced by a planned one.
    refuseImagesBesideKeptDrawing(sheet);
    const number = ++numbering.drawing;
    const path = drawingPart(number);
    // An image's embed relationship belongs to the drawing part, so it is recorded in the drawing's
    // own ledger rather than the sheet's.
    const drawingRels = new RelationshipLedger(path);
    const images = sheet.images.map((image): DrawingImage => {
      const {number: mediaNumber, image: registered} = media.resolve(image.imageId);
      const embedId = drawingRels.add(REL.image, mediaPart(mediaNumber, registered.extension));
      const properties = pictureProperties(image);
      const target = properties.hyperlink?.target;
      // A link to a place in this workbook is a hyperlink relationship whose target is the `#` location,
      // not an external one, which is how Excel writes it; a URL is an external relationship.
      if (target === undefined) return {anchor: image.anchor, embedId, properties};
      const hyperlinkId = target.startsWith('#')
        ? drawingRels.addInDocument(REL.hyperlink, target)
        : drawingRels.addExternal(REL.hyperlink, target);
      return {anchor: image.anchor, embedId, properties, hyperlinkId};
    });
    drawing = {
      number,
      relId: rels.add(REL.drawing, path),
      images,
      relationships: drawingRels.relationships,
    };
  }

  // A conversation and the legacy fallback `<comment>` that binds its cell to it are two halves of one
  // representation, so both are derived from this single list and neither can be emitted without the
  // other. Verified against desktop Excel: a `tc=` fallback whose thread part is absent shows as neither
  // a thread nor a note (the text disappears rather than degrading) and a thread part whose fallback is
  // absent is ignored, leaving the cell blank. A thread with no messages is not one of them: it has
  // nothing to say, and no head id for its replies or its fallback to hang off.
  const threads = sheet.commentThreads.filter((thread) => thread.comments.length > 0);
  const sheetComments = collectComments(liveCells(sheet), threads, flushed?.notes ?? []);
  let comments: CommentPlan | null = null;
  if (sheetComments.length > 0) {
    const vmlRelId = rels.add(REL.vmlDrawing, vmlDrawingPart(index + 1));
    rels.add(REL.comments, commentsPart(index + 1));
    comments = {number: index + 1, comments: sheetComments, vmlRelId};
  }
  let threadedComments: ThreadedCommentPlan | null = null;
  if (threads.length > 0) {
    rels.add(REL.threadedComment, threadedCommentsPart(index + 1));
    threadedComments = {number: index + 1, threads};
  }

  const printerData = sheet.pageSetup.printerSettings;
  const printerSettings: PrinterSettingsPlan | null =
    printerData === undefined
      ? null
      : {
          number: index + 1,
          data: printerData,
          relId: rels.add(REL.printerSettings, printerSettingsPart(index + 1)),
        };

  // The live rows plus whatever the streaming writer already flushed and evicted, merged back into
  // the row-major order Excel writes them in: a committed row's cells are gone from the model, so the
  // walk alone would silently drop its links.
  const hyperlinks = planHyperlinks(
    [...collectHyperlinks(liveCells(sheet)), ...(flushed?.hyperlinks ?? [])].sort(
      (a, b) => a.row - b.row || a.col - b.col,
    ),
    rels,
  );

  let background: BackgroundPlan | null = null;
  if (sheet.backgroundImageId !== undefined) {
    const {number, image} = media.resolve(sheet.backgroundImageId);
    background = {relId: rels.add(REL.image, mediaPart(number, image.extension))};
  }

  const preservedRefs: PreservedReferencePlan[] = (preserved.perSheet[index] ?? []).map(
    (reference) => ({...reference, relId: rels.add(reference.relType, reference.entryPath)}),
  );

  const pivots: PivotPlan[] = sheet.pivotTables.map((table) => {
    const number = ++numbering.pivot;
    // A pivot table is found by relationship type, so nothing in the sheet body cites this id. Each
    // pivot is numbered globally (its parts and its `cacheId` must be workbook-unique); the workbook
    // relationship reaching its cache is recorded once the modeled workbook rels are.
    rels.add(REL.pivotTable, pivotTablePart(number));
    return {number, cacheId: String(number), table, workbookRelId: ''};
  });

  return {
    relationships: rels.relationships,
    tables,
    drawing,
    comments,
    threadedComments,
    printerSettings,
    hyperlinks,
    background,
    preservedRefs,
    pivots,
  };
}

// The part numbers of whichever sheets carry a part of one kind. Four call sites spelled this as a
// map/filter/map triple with a type predicate whose only job was to restate the plan type it had just
// selected; one picker says the same thing without the predicate.
function numbersOf(
  plans: readonly SheetPlan[],
  pick: (plan: SheetPlan) => {readonly number: number} | null,
): number[] {
  const numbers: number[] = [];
  for (const plan of plans) {
    const part = pick(plan);
    if (part !== null) numbers.push(part.number);
  }
  return numbers;
}

// Resolve one sheet's tail reference ids (the `<drawing>`/`<legacyDrawing>`/`<legacyDrawingHF>`/
// `<picture>` slots and the slicer list) from its plan. A preserved `<drawing>` and a modeled one never
// share a sheet (`planSheet` refuses one that has both), so the drawing slot takes whichever exists; a
// comment's VML rides the legacy-
// drawing slot; and each preserved slicer surfaces its rel id so the `<x14:slicerList>` can reactivate
// the widget rather than orphan its part.
function resolveSheetReferences(plan: SheetPlan): SheetReferences {
  const refs = plan.preservedRefs;
  const preservedDrawingRelId = refs.find((ref) => ref.element === 'drawing')?.relId ?? null;
  const legacyDrawingHFRelId = refs.find((ref) => ref.element === 'legacyDrawingHF')?.relId ?? null;
  const slicerRelIds = refs
    .filter((ref) => isRelType(ref.relType, 'slicer'))
    .map((ref) => ref.relId);
  return {
    drawingRelId: plan.drawing?.relId ?? preservedDrawingRelId,
    legacyDrawingRelId: plan.comments?.vmlRelId ?? null,
    printerSettingsRelId: plan.printerSettings?.relId ?? null,
    backgroundRelId: plan.background?.relId ?? null,
    legacyDrawingHFRelId,
    slicerRelIds,
  };
}

/**
 * Record every workbook-level relationship, once, in the one canonical order the workbook part wires
 * them, and return the ids the workbook body cites along with the relationships its `.rels` declares.
 *
 * The order is the whole of it. The modeled rels come first (one per sheet, then the fixed
 * styles/theme pair, then shared strings when emitted), because they are the ones an existing package
 * already numbered: anything laid after them can be added without renumbering an id already in use.
 * The threaded-comment person registry follows, then the preserved workbook references, then the
 * generated pivot caches.
 *
 * A pivot's id is written back onto the shared {@link PivotPlan} rather than returned, because the
 * workbook body's `<pivotCaches>` registration reads it from there, beside the pivot's `cacheId`.
 */
function planWorkbookRelationships(context: {
  readonly sheetCount: number;
  readonly hasSharedStrings: boolean;
  readonly hasPersons: boolean;
  readonly preservedWorkbook: readonly PreservedWorkbookReferencePlan[];
  readonly pivots: readonly PivotPlan[];
}): WorkbookRelPlan {
  const {sheetCount, hasSharedStrings, hasPersons, preservedWorkbook, pivots} = context;
  const rels = new RelationshipLedger(WORKBOOK_PART);
  const sheetRelIds = Array.from({length: sheetCount}, (_, i) =>
    rels.add(REL.worksheet, worksheetPart(i + 1)),
  );
  rels.add(REL.styles, STYLES_PART);
  rels.add(REL.theme, THEME_PART_PATH);
  if (hasSharedStrings) rels.add(REL.sharedStrings, SHARED_STRINGS_PART);
  // The threaded-comment identity registry every conversation on every sheet resolves its authors and
  // @mentions through. Workbook-level and singular, so this one relationship serves all the sheets.
  if (hasPersons) rels.add(REL.person, PERSONS_PART);
  const preservedWorkbookRels = preservedWorkbook.map((ref) => ({
    ...ref,
    relId: rels.add(ref.relType, ref.entryPath),
  }));
  for (const pivot of pivots) {
    pivot.workbookRelId = rels.add(
      REL.pivotCacheDefinition,
      pivotCacheDefinitionPart(pivot.number),
    );
  }
  return {sheetRelIds, preservedWorkbookRels, relationships: rels.relationships};
}

// The package root's relationships: the three the writer regenerates from the model (the office
// document and the core/app properties), then any preserved root reference (a customUI ribbon part,
// custom properties, a thumbnail) after them, so a round trip keeps content wired from `_rels/.rels`
// that the model does not otherwise emit.
function planRootRelationships(
  rootRefs: readonly PreservedRootReferencePlan[],
): readonly PlannedRelationship[] {
  const rels = new RelationshipLedger('');
  rels.add(REL.officeDocument, WORKBOOK_PART);
  rels.add(REL.coreProps, CORE_PROPS_PART);
  rels.add(REL.extProps, APP_PROPS_PART);
  for (const ref of rootRefs) rels.add(ref.relType, ref.entryPath);
  return rels.relationships;
}

// Everything about a package that is resolved before any of its bytes exist: the media every sheet
// shares, the parts carried over verbatim, and each sheet's own parts with their relationship ids
// already allocated. Named because it is the boundary between the two halves of the writer: nothing
// above it serialises, and nothing below it decides what the package contains.
interface PackagePlan {
  readonly media: MediaPlan;
  readonly preserved: PreservedPlan;
  readonly perSheet: readonly SheetPlan[];
  readonly allTables: readonly TablePlan[];
  readonly allPivots: readonly PivotPlan[];
}

// Resolve the whole package graph: the media the sheets share, the verbatim-preserved parts numbered
// past the generated ones, then every sheet's parts in a single pass through {@link planSheet}.
function planPackage(
  workbook: Workbook,
  sheets: readonly Worksheet[],
  flushed: InternalWriteOptions['flushed'],
): PackagePlan {
  // Anchored images share workbook-wide media: every image a sheet references becomes one media part,
  // addressed by a global number. Resolved before the sheet loop so a drawing's embeds can target it.
  const media = planMedia(workbook, sheets);

  // Content the model does not interpret (a vector-shape drawing, a header/footer image, a pivot
  // table and its caches, a slicer) captured on read and re-emitted verbatim onto collision-proof
  // paths. Preserved parts are renumbered past the parts the writer generates of the same kind
  // (drawings, VML, media), so resolving them needs only those generated counts; each sheet's
  // preserved references take their sheet-local rel ids in canonical position below.
  const generatedDrawingCount = sheets.filter((sheet) => sheet.images.length > 0).length;
  // Pivots are numbered globally across the workbook, so the count of authored ones is what a
  // preserved pivot's parts must be renumbered past.
  const generatedPivotCount = sheets.reduce((total, sheet) => total + sheet.pivotTables.length, 0);
  const preserved = planPreservedParts(
    workbook,
    generatedDrawingCount,
    media.parts.length,
    generatedPivotCount,
  );

  // The part numbers that are global across the workbook (tables, drawings, pivots) run through one
  // shared counter, which is why this is a `map` over a mutable object rather than a pure one: the
  // shared state is in the signature instead of being three `let`s a callback happens to close over.
  const numbering: PartNumbering = {table: 0, drawing: 0, pivot: 0};
  const perSheet = sheets.map((sheet, index) =>
    planSheet({sheet, index, media, preserved, numbering, flushed: flushed?.get(sheet)}),
  );

  return {
    media,
    preserved,
    perSheet,
    allTables: perSheet.flatMap((plan) => plan.tables),
    allPivots: perSheet.flatMap((plan) => plan.pivots),
  };
}

/**
 * Serialise every worksheet, in sheet order.
 *
 * This runs before `xl/styles.xml` is generated, and must: interning a cell's or row's format into
 * the style table is a side effect of this pass, so the stylesheet is only complete once every sheet
 * has been through it. Emitting the styles part first would silently drop the styles of whatever had
 * not been serialised yet.
 */
function serialiseSheets(context: {
  readonly workbook: Workbook;
  readonly sheets: readonly Worksheet[];
  readonly plan: PackagePlan;
  readonly styles: StyleRegistry;
  readonly sharedStrings: SharedStringTable | null;
  readonly flushed: InternalWriteOptions['flushed'];
}): string[] {
  const {workbook, sheets, plan, styles, sharedStrings, flushed} = context;
  return sheets.map((sheet, i) => {
    const sheetPlan = plan.perSheet[i] as SheetPlan;
    return worksheetXml({
      sheet,
      tables: sheetPlan.tables,
      styles,
      references: resolveSheetReferences(sheetPlan),
      hyperlinks: sheetPlan.hyperlinks,
      sharedStrings,
      dateEpoch: workbook.dateEpoch,
      formulaNames: formulaNamesInScope(workbook.definedNames, sheet.name),
      // Exactly one sheet is marked selected; the model resolves which, so no package can ship with
      // none selected (no view initialised on open) or with several (an accidental group selection,
      // where an edit to one sheet lands on all of them).
      active: i === workbook.activeTabIndex,
      flushed: flushed?.get(sheet),
    });
  });
}

/**
 * Build the part map itself: every part path paired with its bytes.
 *
 * The workbook-level relationship ids are laid out here rather than in the plan because two of their
 * inputs are known only once the sheets are serialised: whether any string was interned into the
 * shared-strings pool, and whether any sheet carried a conversation for the person registry to serve.
 */
function emitPackageParts(context: {
  readonly workbook: Workbook;
  readonly sheets: readonly Worksheet[];
  readonly plan: PackagePlan;
  readonly styles: StyleRegistry;
  readonly sharedStrings: SharedStringTable | null;
  readonly sheetXml: readonly string[];
}): Record<string, Uint8Array> {
  const {workbook, sheets, plan, styles, sharedStrings, sheetXml} = context;
  const {media, preserved, perSheet, allTables, allPivots} = plan;

  // The pool is filled only once every sheet is serialised. Emit the part (and its rel + content
  // type) solely when the option is on and at least one string was interned, so a workbook with no
  // string cells never fabricates an empty table.
  const hasSharedStrings = sharedStrings !== null && !sharedStrings.isEmpty;

  const commentNumbers = numbersOf(perSheet, (sheetPlan) => sheetPlan.comments);
  const drawingNumbers = numbersOf(perSheet, (sheetPlan) => sheetPlan.drawing);
  const printerSettingsNumbers = numbersOf(perSheet, (sheetPlan) => sheetPlan.printerSettings);
  const threadedCommentNumbers = numbersOf(perSheet, (sheetPlan) => sheetPlan.threadedComments);

  // The identity registry is emitted only beside the thread parts that point into it. With no conversation
  // in the package nothing can reference a `<person>`, so the part would be a workbook-level relationship
  // to dead weight, and it is the messages, not the registry, that make an identity worth carrying.
  const persons = threadedCommentNumbers.length === 0 ? [] : workbook.persons;

  const workbookRels = planWorkbookRelationships({
    sheetCount: sheets.length,
    hasSharedStrings,
    hasPersons: persons.length > 0,
    preservedWorkbook: preserved.workbook,
    pivots: allPivots,
  });
  const {preservedWorkbookRels} = workbookRels;

  const files = new PackageFiles();
  files.add(
    '[Content_Types].xml',
    strToU8(
      contentTypesXml({
        sheetCount: sheets.length,
        tables: allTables,
        commentNumbers,
        drawingNumbers,
        printerSettingsNumbers,
        mediaExtensions: media.extensions,
        hasSharedStrings,
        preservedParts: preserved.parts,
        pivots: allPivots,
        preservedWorkbookRefs: preservedWorkbookRels,
        threadedCommentNumbers,
        hasPersons: persons.length > 0,
      }),
    ),
  );
  files.add(relsPathFor(''), strToU8(relsPartXml(planRootRelationships(preserved.root))));
  files.add(CORE_PROPS_PART, strToU8(corePropsXml(workbook.properties)));
  files.add(APP_PROPS_PART, strToU8(appPropsXml(workbook.properties)));
  files.add(WORKBOOK_PART, strToU8(workbookXml(workbook, workbookRels, allPivots)));
  files.add(relsPathFor(WORKBOOK_PART), strToU8(relsPartXml(workbookRels.relationships)));
  files.add(STYLES_PART, strToU8(styles.toXml()));
  // A theme read from a source package is emitted through the preserved-part path, closure and all,
  // with any authored overrides already spliced into its entry part by the planner. A workbook without
  // one gets its authored theme, or the library's default, which the stylesheet's `theme="1"` default
  // font still needs something to resolve against.
  if (!preserved.themeEmitted) {
    const overrides = workbook.themeOverrides;
    files.add(
      THEME_PART_PATH,
      strToU8(
        overrides === undefined
          ? DEFAULT_THEME_XML
          : applyThemeOverrides(DEFAULT_THEME_XML, overrides),
      ),
    );
  }
  if (hasSharedStrings) {
    files.add(SHARED_STRINGS_PART, strToU8(sharedStrings.toXml()));
  }
  // Singular and unnumbered, unlike the per-sheet thread parts: one registry serves the whole workbook.
  if (persons.length > 0) files.add(PERSONS_PART, strToU8(personsXml(persons)));
  for (const part of media.parts) {
    files.add(mediaPart(part.number, part.extension), part.data);
  }
  emitSheetParts(files, perSheet, sheetXml, (id) => workbook.getPerson(id));
  for (const {table, number} of allTables) {
    files.add(tablePart(number), strToU8(tableXml(table, number)));
  }
  emitPivotParts(files, allPivots);
  emitPreservedParts(files, preserved.parts);

  return files.toRecord();
}

/**
 * Assemble a workbook into the map of OPC package parts (part name → bytes) that make up an `.xlsx`,
 * short of zipping them. This is the whole serialisation (content types, relationships, workbook,
 * per-sheet XML, styles, theme, media, tables, and props) factored out of {@link writeXlsx} so the
 * streaming writer can drive the identical parts through a streamed zip container rather than
 * `zipSync`. Neither writer duplicates a byte of serialisation.
 *
 * Three steps, and their order is load-bearing. The package is planned in full before anything is
 * serialised, so a drawing's embed and a pivot's cache can target a part number that already exists.
 * The sheets are serialised before the parts are emitted, because interning a style is a side effect
 * of that pass and `xl/styles.xml` is one of the parts emitted in the third step.
 *
 * @throws {AuthoringError} if the workbook has no worksheets, or holds a value the writer cannot represent.
 */
export function buildPackageParts(
  workbook: Workbook,
  options: InternalWriteOptions = {},
): Record<string, Uint8Array> {
  const sheets = workbook.worksheets;
  if (sheets.length === 0) {
    throw new AuthoringError(
      'cannot write a workbook with no worksheets: a zero-sheet package is corrupt to Excel',
    );
  }

  // With the option on, plain string cell values are pooled into a shared-strings table interned
  // during the sheet pass (like the style registry); a null table keeps every string inline.
  const sharedStrings = options.useSharedStrings ? new SharedStringTable() : null;
  // The streaming writer supplies its own registry (already seeded, and already carrying its eagerly
  // flushed rows' styles); the buffered path seeds a fresh one here.
  const styles = options.styles ?? createStyleRegistry(workbook);

  const plan = planPackage(workbook, sheets, options.flushed);
  const sheetXml = serialiseSheets({
    workbook,
    sheets,
    plan,
    styles,
    sharedStrings,
    flushed: options.flushed,
  });
  return emitPackageParts({workbook, sheets, plan, styles, sharedStrings, sheetXml});
}

/**
 * The OPC part paths a package is being assembled from, accumulated by {@link buildPackageParts} and
 * its per-phase `emit*` helpers.
 *
 * A class rather than the plain record it used to be, for two reasons that are both about making a
 * silent failure loud.
 *
 * **A collision is refused.** Twenty sites write into this map, and the paths they choose come from
 * several independent numberings: the generated parts number from 1 per kind, and a preserved part
 * carries a path renumbered past them, or its own if the writer generates no part of that kind. When
 * one of those numberings was wrong, the later write simply won and the earlier part was gone from
 * the package with the content types still declaring both. Refusing here means the arithmetic is
 * checked at every write site at once rather than trusted at each.
 *
 * **The map has no prototype.** A preserved part can keep its original zip entry name, which comes
 * from an untrusted package, so a part path of `__proto__` is reachable. Assigned onto a plain object
 * it silently drops the part and re-points the map's prototype at its bytes. `io/opc/inflate.ts` and
 * `io/xlsx/edit-vba.ts` both defend against exactly this on the read side, and say why; this was the
 * writer-side gap in the same rule.
 */
class PackageFiles {
  readonly #files: Record<string, Uint8Array> = Object.create(null) as Record<string, Uint8Array>;

  /**
   * Add a part at `path`.
   *
   * @throws {InternalError} if a part is already emitted at that path. Two parts on one path is a
   *   numbering bug in this writer, never something a caller can cause, and the package it would
   *   produce carries one of them plus a duplicate `PartName` override that violates OPC M2.5.
   */
  add(path: string, bytes: Uint8Array): void {
    if (path in this.#files) {
      throw new InternalError(`two package parts claim the path ${quoted(path)}`);
    }
    this.#files[path] = bytes;
  }

  /** The assembled package, as the part-path → bytes map the zip layer takes. */
  toRecord(): Record<string, Uint8Array> {
    return this.#files;
  }
}

// Emit each sheet's own parts: the sheet XML, its rels part (only when the sheet references something),
// and the drawing/comment/printer-settings parts those relationships point at. `sheetXml[i]` is the
// already-serialised body for `perSheet[i]`, indexed in lockstep. `personById` is the workbook's
// registry, which every message's author has to be in.
function emitSheetParts(
  files: PackageFiles,
  perSheet: readonly SheetPlan[],
  sheetXml: readonly string[],
  personById: (id: string) => Person | undefined,
): void {
  perSheet.forEach((plan, i) => {
    const {relationships, drawing, comments, threadedComments, printerSettings} = plan;
    files.add(worksheetPart(i + 1), strToU8(sheetXml[i] as string));
    if (relationships.length > 0) {
      files.add(relsPathFor(worksheetPart(i + 1)), strToU8(relsPartXml(relationships)));
    }
    if (printerSettings !== null) {
      files.add(printerSettingsPart(printerSettings.number), printerSettings.data);
    }
    if (drawing !== null) {
      const drawingPath = drawingPart(drawing.number);
      files.add(drawingPath, strToU8(drawingXml(drawing.images)));
      files.add(relsPathFor(drawingPath), strToU8(relsPartXml(drawing.relationships)));
    }
    if (comments !== null) {
      files.add(commentsPart(comments.number), strToU8(commentsXml(comments.comments)));
      files.add(vmlDrawingPart(comments.number), strToU8(vmlDrawingXml(comments.comments)));
    }
    if (threadedComments !== null) {
      files.add(
        threadedCommentsPart(threadedComments.number),
        strToU8(threadedCommentsXml(threadedComments.threads, personById)),
      );
    }
  });
}

// Emit every pivot table's three chained parts. A pivot spans a pivot-table part (linked from its host
// sheet) that reaches its cache definition by relationship type, and a cache definition whose `r:id`
// cites the relationship to its cache records. Each link is recorded in its owner's ledger, and the
// definition is rendered with the id its ledger handed out rather than one both sides assume.
function emitPivotParts(files: PackageFiles, allPivots: readonly PivotPlan[]): void {
  for (const pivot of allPivots) {
    const {number, cacheId, table} = pivot;
    const tablePath = pivotTablePart(number);
    const definitionPath = pivotCacheDefinitionPart(number);
    const recordsPath = pivotCacheRecordsPart(number);

    const tableRels = new RelationshipLedger(tablePath);
    tableRels.add(REL.pivotCacheDefinition, definitionPath);
    files.add(tablePath, strToU8(pivotTableXml(table, `PivotTable${number}`, cacheId)));
    files.add(relsPathFor(tablePath), strToU8(relsPartXml(tableRels.relationships)));

    const definitionRels = new RelationshipLedger(definitionPath);
    const recordsRelId = definitionRels.add(REL.pivotCacheRecords, recordsPath);
    files.add(definitionPath, strToU8(pivotCacheDefinitionXml(table, recordsRelId)));
    files.add(relsPathFor(definitionPath), strToU8(relsPartXml(definitionRels.relationships)));

    files.add(recordsPath, strToU8(pivotCacheRecordsXml(table)));
  }
}

// Emit the verbatim-preserved parts (and their rewired rels) last: their paths are collision-proof, so
// ordering against the generated parts does not matter.
function emitPreservedParts(files: PackageFiles, parts: readonly PreservedPartPlan[]): void {
  for (const part of parts) {
    files.add(part.path, part.bytes);
    if (part.relsPath !== null && part.relsXml !== null) {
      files.add(part.relsPath, strToU8(part.relsXml));
    }
  }
}
