// The plan layer of the writer: pure graph resolution that turns a Workbook model into the numbered,
// cross-referenced set of parts an `.xlsx` package needs (media, preserved verbatim-carried parts,
// and the sheet-/workbook-local relationship ids that wire them) before any XML is serialised.

import type {CommentThread} from '../../core/comment-thread.ts';
import type {WorkbookImage} from '../../core/image.ts';
import type {PivotTable} from '../../core/pivot-table.ts';
import type {Table} from '../../core/table.ts';
import type {Workbook} from '../../core/workbook.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {AuthoringError, InternalError, quoted} from '../../errors.ts';
import {extensionOf, relativePartPath, relsPathFor, THEME_PART_PATH} from '../opc/part-paths.ts';
import {relsPartXml} from '../opc/rels.ts';
import type {CommentCell} from './comments.ts';
import type {HyperlinkPlan} from './hyperlinks.ts';
import type {DrawingImage} from './images.ts';
import {
  drawingPart,
  mediaPart,
  pivotCacheDefinitionPart,
  pivotCacheRecordsPart,
  pivotTablePart,
  vmlDrawingPart,
} from './part-names.ts';
import {applyThemeOverrides} from './theme-xml.ts';

/** One relationship a generated `.rels` part declares, its target already relative to the owner. */
export interface PlannedRelationship {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly external?: boolean;
}

/**
 * The relationships one part owns, recorded as their ids are handed out.
 *
 * An id and the relationship it names are one fact, so they are written down in one call. When the
 * ids came from a counter and the `.rels` part was listed somewhere else, a sheet's set was spelled
 * three times: once where ids were drawn, once in the rels renderer in another order, and once more in
 * a nine-way condition deciding whether the rels part existed at all. A tenth kind of part needed all
 * three edits, and missing the last dropped the `.rels` part while the sheet XML still cited the id.
 * Now the ledger is the rels part, and a part with an empty ledger has none.
 *
 * Ids run `rId1`, `rId2`, … in the order parts are added, so no id is derived by summing the ones
 * before it, and the `.rels` part lists them in that same order. One ledger per owning part; ids are
 * scoped to it.
 */
export class RelationshipLedger {
  readonly #owner: string;
  readonly #relationships: PlannedRelationship[] = [];

  /** @param owner the package path of the part whose `.rels` this is, `''` for the package root. */
  constructor(owner: string) {
    this.#owner = owner;
  }

  /** Record a relationship to a package part, named by its package path, and return its id. */
  add(type: string, partPath: string): string {
    return this.#record({type, target: relativePartPath(this.#owner, partPath)});
  }

  /** Record a relationship to a target outside the package (a hyperlink's URL) and return its id. */
  addExternal(type: string, url: string): string {
    return this.#record({type, target: url, external: true});
  }

  /** Record a relationship to a place inside this document (a picture link's `#Sheet1!C3`), kept as
   * written rather than resolved as a part path, and return its id. */
  addInDocument(type: string, location: string): string {
    return this.#record({type, target: location});
  }

  /** Every relationship recorded, in id order. */
  get relationships(): readonly PlannedRelationship[] {
    return this.#relationships;
  }

  #record(relationship: Omit<PlannedRelationship, 'id'>): string {
    const id = `rId${this.#relationships.length + 1}`;
    this.#relationships.push({id, ...relationship});
    return id;
  }
}

// A pivot table planned for emission: its global part number, the workbook-unique `cacheId` its
// `<pivotCaches>` registration and `pivotTableDefinition` agree on, and the workbook relationship
// reaching its cache definition (assigned once the modeled workbook rels are recorded).
export interface PivotPlan {
  readonly number: number;
  readonly cacheId: string;
  readonly table: PivotTable;
  workbookRelId: string;
}

// A sheet's comments (its cells' notes and one legacy fallback per threaded conversation) paired with
// the part number and the sheet-local relationship id its `<legacyDrawing>` element cites for the VML
// drawing. The comments part itself is found by relationship type, so nothing cites its id.
export interface CommentPlan {
  readonly number: number;
  readonly comments: readonly CommentCell[];
  readonly vmlRelId: string;
}

// A sheet's threaded conversations paired with the part number naming its
// `threadedComments/threadedComment{n}.xml`. No worksheet element points at that part. Excel discovers
// it by scanning the sheet's relationships, the way it finds a pivot table, so the relationship the
// sheet's ledger records is the whole of the wiring.
export interface ThreadedCommentPlan {
  readonly number: number;
  readonly threads: readonly CommentThread[];
}

// A sheet's opaque printer-settings blob paired with the part number naming its `.bin` part and the
// sheet-local relationship id that links the sheet's `<pageSetup r:id>` to it.
export interface PrinterSettingsPlan {
  readonly number: number;
  readonly data: Uint8Array;
  readonly relId: string;
}

// A table paired with the identifiers the package needs: a workbook-global part number
// and the sheet-local relationship id that links its worksheet to the table part.
export interface TablePlan {
  readonly table: Table;
  readonly number: number;
  readonly relId: string;
}

// A sheet background image resolved for serialisation: the sheet-local relationship id its
// `<picture>` element references. The relationship to the media part is in the sheet's ledger.
export interface BackgroundPlan {
  readonly relId: string;
}

// A verbatim-preserved package part resolved for serialisation: the collision-proof path it is
// emitted at, its bytes and content type, and, when it references other parts, the rels part
// linking it to their new paths.
export interface PreservedPartPlan {
  readonly path: string;
  readonly bytes: Uint8Array;
  readonly contentType: string;
  readonly relsPath: string | null;
  readonly relsXml: string | null;
}

// A preserved worksheet reference resolved for serialisation, short of its sheet-local relationship
// id: the worksheet element that wires it (`undefined` for a pivot-table/slicer reference the sheet
// carries by relationship alone), the relationship Type, and the new path of the entry part it
// targets. The id is recorded by the caller in the sheet's {@link RelationshipLedger}, after the
// sheet's generated parts, so a preserved reference never renumbers an id already threaded into the
// sheet XML.
export interface ResolvedPreservedReference {
  readonly element: 'drawing' | 'legacyDrawingHF' | undefined;
  readonly relType: string;
  readonly entryPath: string;
}

// A resolved preserved reference with its sheet-local relationship id assigned.
export interface PreservedReferencePlan extends ResolvedPreservedReference {
  readonly relId: string;
}

// A preserved workbook reference resolved for serialisation: its relationship Type, the new path of
// the entry part, and, for a pivot cache, the `cacheId` its `<pivotCaches>` registration carries.
// The workbook relationship id is recorded at emit time (it follows the modeled workbook rels, whose
// number depends on whether a shared-strings part and a person registry are emitted), so it is not
// fixed here.
export interface PreservedWorkbookReferencePlan {
  readonly relType: string;
  readonly entryPath: string;
  readonly pivotCacheId: string | undefined;
  readonly externalReferenceIndex: number | undefined;
}

// A preserved package-root reference resolved for serialisation: its relationship Type and the new
// path of the entry part it targets. Its root relationship is recorded at emit time, after the fixed
// root relationships (the office document and the core/app properties).
export interface PreservedRootReferencePlan {
  readonly relType: string;
  readonly entryPath: string;
}

// The whole workbook's preserved content resolved for serialisation: per-sheet references (parallel to
// the sheets) still awaiting their sheet-local relationship ids, the workbook-level reference plans,
// the package-root reference plans, and the flat, de-duplicated list of parts to emit. Kept together
// because the parts are numbered globally while the references are sheet-, workbook-, or root-local.
export interface PreservedPlan {
  readonly perSheet: readonly (readonly ResolvedPreservedReference[])[];
  readonly workbook: readonly PreservedWorkbookReferencePlan[];
  readonly root: readonly PreservedRootReferencePlan[];
  readonly parts: readonly PreservedPartPlan[];
  /**
   * Whether a source theme rides in {@link parts} at {@link THEME_PART_PATH}. The theme relationship
   * and content-type override are emitted unconditionally either way; this only tells the writer
   * whether to *also* generate its default theme body, which would otherwise clobber the preserved one.
   */
  readonly themeEmitted: boolean;
}

// A sheet's drawing part: its workbook-global number, the sheet-local relationship id linking the
// sheet's `<drawing>` element to it, the images it lays out, and the drawing's own relationships, one
// per image, which each image's `r:embed` cites.
export interface DrawingPlan {
  readonly number: number;
  readonly relId: string;
  readonly images: readonly DrawingImage[];
  readonly relationships: readonly PlannedRelationship[];
}

// One picture written to `xl/media/`: its global part number, extension, and bytes.
export interface MediaPart {
  readonly number: number;
  readonly extension: string;
  readonly data: Uint8Array;
}

// A referenced image, resolved: the media part number a relationship targets, and the registered
// image behind it. Both together, because every caller that wants the number also wants the
// extension, and looking the image up a second time is what would need an assertion to type.
export interface ResolvedMedia {
  readonly number: number;
  readonly image: WorkbookImage;
}

// The workbook's media, resolved for writing: the parts to emit, the resolution of a workbook image
// id (so a drawing embed can target its media part), and the distinct extensions in use (so content
// types can declare an image `<Default>` per extension).
export interface MediaPlan {
  readonly parts: readonly MediaPart[];
  readonly extensions: readonly string[];
  /** The media part number and registered image for a referenced id. Total over the ids
   * {@link planMedia} saw on the sheets it planned, because planning already threw on any it could
   * not resolve; an id from anywhere else is a caller bug and is reported as one. */
  resolve(id: number): ResolvedMedia;
}

// Per-kind counters for numbering preserved parts, each seeded past the generated parts of its kind.
interface PreservedNumbering {
  drawing: number;
  vml: number;
  media: number;
  pivotTable: number;
  pivotCacheDefinition: number;
  pivotCacheRecords: number;
}

// Where a sheet first referenced an image id, kept so a failure to resolve it can say which sheet
// and in which of the two roles. The role is carried as a token rather than a phrase because the
// two readings ("anchors image id 7" against "sets background image id 7") are one sentence apart
// and are better built at the throw than stored pre-worded.
interface MediaUse {
  readonly id: number;
  readonly sheetName: string;
  readonly role: 'anchor' | 'background';
}

// Gather the workbook images actually referenced by some sheet, either anchored in a drawing or set
// as a sheet background (an unreferenced image is not written), number them in first-use order, and
// record the extensions in play. This is the *only* place a sheet's image id is checked against the
// workbook's registry: a referenced-but-unregistered id is a programming error surfaced here, named
// by sheet and role, rather than emitted as a dangling relationship or re-checked downstream.
export function planMedia(workbook: Workbook, sheets: readonly Worksheet[]): MediaPlan {
  const uses: MediaUse[] = [];
  const seen = new Set<number>();
  const use = (id: number, sheetName: string, role: MediaUse['role']): void => {
    if (!seen.has(id)) {
      seen.add(id);
      uses.push({id, sheetName, role});
    }
  };
  for (const sheet of sheets) {
    for (const image of sheet.images) use(image.imageId, sheet.name, 'anchor');
    if (sheet.backgroundImageId !== undefined) {
      use(sheet.backgroundImageId, sheet.name, 'background');
    }
  }
  const parts: MediaPart[] = [];
  const byId = new Map<number, ResolvedMedia>();
  const extensions = new Set<string>();
  uses.forEach(({id, sheetName, role}, i) => {
    const image = workbook.getImage(id);
    if (image === undefined) {
      const reference =
        role === 'anchor' ? `anchors image id ${id}` : `sets background image id ${id}`;
      throw new AuthoringError(
        `sheet ${quoted(sheetName)} ${reference}, which is not registered on the workbook`,
      );
    }
    const number = i + 1;
    parts.push({number, extension: image.extension, data: image.data});
    byId.set(id, {number, image});
    extensions.add(image.extension);
  });
  return {
    parts,
    extensions: [...extensions],
    resolve(id: number): ResolvedMedia {
      const resolved = byId.get(id);
      if (resolved === undefined) {
        throw new AuthoringError(`image id ${id} was not planned into this workbook's media`);
      }
      return resolved;
    },
  };
}

// Resolve every sheet's verbatim-preserved worksheet references (a vector-shape drawing, a
// header/footer image) into the parts to emit and the per-sheet reference data that wires them. Each
// reference's captured part closure is re-numbered onto collision-proof `preservedP{n}` paths, so
// preserved content never clobbers a generated drawing/VML/media part, with the closure's internal
// relationships rewritten to the new sibling paths. Part numbering is the only cross-sheet concern
// here; each reference's sheet-local relationship is recorded by the caller in the sheet's
// {@link RelationshipLedger}, so this function stays free of sheet-local ids.
export function planPreservedParts(
  workbook: Workbook,
  generatedDrawingCount: number,
  generatedMediaCount: number,
  generatedPivotCount: number,
): PreservedPlan {
  const sheets = workbook.worksheets;
  // Every kind the writer generates of its own is re-numbered past the generated ones, so a preserved
  // part never clobbers a generated one: a preserved drawing never lands on an anchored drawing's
  // path, a preserved VML never on a comment's. Comment VML is numbered by sheet index, so
  // `sheets.length` bounds it. Pivots are the kind this list forgot: the writer emits a pivot table
  // part and both cache parts, numbered globally from 1, so a package that already carried a pivot
  // and then had one authored onto it wrote both at `pivotTable1.xml`. The preserved bytes won, the
  // new pivot's sheet relationship pointed at the old pivot's data, and the content types declared
  // the same `PartName` twice, which is a package Excel repairs. Kinds the writer really never
  // generates (a slicer, a chart) still keep their original path. See {@link preservedPartPath}.
  const numbering: PreservedNumbering = {
    drawing: generatedDrawingCount,
    vml: sheets.length,
    media: generatedMediaCount,
    pivotTable: generatedPivotCount,
    pivotCacheDefinition: generatedPivotCount,
    pivotCacheRecords: generatedPivotCount,
  };

  // One package-wide remap and one emitted-parts map: a part reached through more than one reference
  // (a pivot cache reached both from its pivot table and from the workbook) is numbered once and
  // emitted once, so overlapping closures collapse instead of duplicating parts.
  const remap = new Map<string, string>();
  // A preserved theme rides the same closure machinery as every other verbatim part: it can carry
  // relationships of its own (a picture used as a themed fill) that need the same renumbering and
  // rewiring. Its entry is pinned to the fixed theme path rather than left to {@link preservedPartPath},
  // because the workbook's theme relationship and the content-type override name that path
  // unconditionally; a source package that called its part `theme2.xml` would otherwise land it
  // somewhere neither points.
  const theme = workbook.themePart;
  if (theme !== undefined) remap.set(theme.entryPath, THEME_PART_PATH);
  const allReferences = [
    ...sheets.flatMap((sheet) => sheet.preservedReferences),
    ...workbook.preservedReferences,
    ...workbook.preservedRootReferences,
    ...(theme === undefined ? [] : [theme]),
  ];
  for (const reference of allReferences) {
    for (const part of reference.parts) {
      if (!remap.has(part.path)) remap.set(part.path, preservedPartPath(part.path, numbering));
    }
  }
  const emitted = new Map<string, PreservedPartPlan>();
  for (const reference of allReferences) {
    for (const part of reference.parts) {
      const newPath = resolveRemapped(remap, part.path);
      if (emitted.has(newPath)) continue;
      const rels = part.rels.flatMap((rel) => {
        // An external relationship (a linked workbook) is emitted verbatim: its target is outside the
        // package, so it is neither in the remap nor expressed relative to the new path.
        if (rel.external) {
          return [{id: rel.id, type: rel.type, target: rel.targetPath, external: true}];
        }
        const target = remap.get(rel.targetPath);
        return target === undefined
          ? []
          : [{id: rel.id, type: rel.type, target: relativePartPath(newPath, target)}];
      });
      // The one preserved part whose *bytes* can change: a theme the caller authored over is
      // composed onto the source part rather than carried verbatim, so the format scheme, the
      // unauthored slots' encoding, and the relationships below all still ride through; only the
      // authored elements differ.
      const overrides = newPath === THEME_PART_PATH ? workbook.themeOverrides : undefined;
      emitted.set(newPath, {
        path: newPath,
        bytes:
          overrides === undefined
            ? part.bytes
            : new TextEncoder().encode(
                applyThemeOverrides(new TextDecoder().decode(part.bytes), overrides),
              ),
        contentType: part.contentType,
        relsPath: rels.length === 0 ? null : relsPathFor(newPath),
        relsXml: rels.length === 0 ? null : relsPartXml(rels),
      });
    }
  }

  const perSheet = sheets.map((sheet): ResolvedPreservedReference[] =>
    sheet.preservedReferences.map((reference) => ({
      element: reference.element,
      relType: reference.relType,
      entryPath: resolveRemapped(remap, reference.entryPath),
    })),
  );

  const workbookRefs = workbook.preservedReferences.map(
    (reference): PreservedWorkbookReferencePlan => ({
      relType: reference.relType,
      entryPath: resolveRemapped(remap, reference.entryPath),
      pivotCacheId: reference.pivotCacheId,
      externalReferenceIndex: reference.externalReferenceIndex,
    }),
  );

  const rootRefs = workbook.preservedRootReferences.map(
    (reference): PreservedRootReferencePlan => ({
      relType: reference.relType,
      entryPath: resolveRemapped(remap, reference.entryPath),
    }),
  );

  return {
    perSheet,
    workbook: workbookRefs,
    root: rootRefs,
    parts: [...emitted.values()],
    themeEmitted: theme !== undefined,
  };
}

/**
 * The path a preserved part was renumbered onto.
 *
 * The invariant is held a module away: `capturePartClosure` returns `undefined` when the entry part is
 * absent, so a reference that reaches planning has every part of its closure in the remap. Asserting
 * that with a cast made the one shape a violation could take an `undefined` path interpolated into a
 * relationship target, where `escapeAttr` throws a bare `TypeError` from inside `relationship()` and
 * says nothing about which part went missing.
 *
 * @throws {InternalError} naming the path, since a miss can only be a bug in this planner.
 */
function resolveRemapped(remap: ReadonlyMap<string, string>, path: string): string {
  const resolved = remap.get(path);
  if (resolved === undefined) {
    throw new InternalError(`preserved part ${quoted(path)} was never assigned a new path`);
  }
  return resolved;
}

// The path a preserved part is emitted at. A kind the writer generates of its own (a drawing, a VML, a
// media image, a pivot table and its two cache parts) is re-numbered past the generated parts of that
// kind (see {@link planPreservedParts}) so it never clobbers one. A kind the writer never generates (a
// slicer or slicer cache, a chart) keeps its original path, leaving the package's standard part names
// intact and letting overlapping closures agree on a single path for a shared part.
function preservedPartPath(originalPath: string, numbering: PreservedNumbering): string {
  const ext = extensionOf(originalPath);
  if (ext === 'vml') return vmlDrawingPart(++numbering.vml);
  if (originalPath.startsWith('xl/media/')) return mediaPart(++numbering.media, ext);
  if (originalPath.startsWith('xl/drawings/') && ext === 'xml') {
    return drawingPart(++numbering.drawing);
  }
  // Matched on the generated part's own name rather than on its directory: `xl/pivotCache/` also holds
  // slicer caches, which the writer does not generate and must not renumber, and the two are told
  // apart by exactly this prefix.
  if (isNumberedPart(originalPath, PIVOT_TABLE_PREFIX)) {
    return pivotTablePart(++numbering.pivotTable);
  }
  if (isNumberedPart(originalPath, PIVOT_CACHE_DEFINITION_PREFIX)) {
    return pivotCacheDefinitionPart(++numbering.pivotCacheDefinition);
  }
  if (isNumberedPart(originalPath, PIVOT_CACHE_RECORDS_PREFIX)) {
    return pivotCacheRecordsPart(++numbering.pivotCacheRecords);
  }
  return originalPath;
}

// Whether a path is one of the writer's own numbered parts of a kind, i.e. the kind's fixed prefix
// followed by a run of digits and `.xml`. Built from the part-name builders themselves, so a rename
// there cannot leave this matching the old spelling.
const PIVOT_TABLE_PREFIX = numberedPartPrefix(pivotTablePart);
const PIVOT_CACHE_DEFINITION_PREFIX = numberedPartPrefix(pivotCacheDefinitionPart);
const PIVOT_CACHE_RECORDS_PREFIX = numberedPartPrefix(pivotCacheRecordsPart);

function numberedPartPrefix(part: (n: number) => string): string {
  return part(1).slice(0, -'1.xml'.length);
}

function isNumberedPart(path: string, prefix: string): boolean {
  return path.startsWith(prefix) && /^\d+\.xml$/.test(path.slice(prefix.length));
}

// One worksheet's planned package parts and the sheet-local relationships wiring them, produced in
// the single planning pass. Held as a struct per sheet rather than eight index-aligned arrays, so a
// downstream step reads one sheet's plan as a unit and cannot transpose two sheets by mis-indexing.
export interface SheetPlan {
  /** Every relationship the sheet's `.rels` part declares, in id order; empty when it needs none. */
  readonly relationships: readonly PlannedRelationship[];
  readonly tables: TablePlan[];
  readonly drawing: DrawingPlan | null;
  readonly comments: CommentPlan | null;
  readonly threadedComments: ThreadedCommentPlan | null;
  readonly printerSettings: PrinterSettingsPlan | null;
  readonly hyperlinks: HyperlinkPlan[];
  readonly background: BackgroundPlan | null;
  readonly preservedRefs: PreservedReferencePlan[];
  readonly pivots: PivotPlan[];
}
