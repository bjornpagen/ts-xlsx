// Workbook-level serialisation: the package's `[Content_Types].xml`, `xl/workbook.xml` (sheets,
// defined names, calc/protection settings, pivot-cache and slicer registrations), and the `docProps`
// core/app property parts. The root and workbook `.rels` parts are the writer's relationship ledgers.

import {formulaNamesInScope, mangleFormula, quoteSheetName} from '../../core/formula.ts';
import {WORKBOOK_PROTECTION_CREDENTIAL_ATTRS} from '../../core/workbook-protection.ts';
import type {Workbook, WorkbookProperties} from '../../core/workbook.ts';
import {isVisibility, type Worksheet} from '../../core/worksheet.ts';
import {AuthoringError, InternalError, quoted} from '../../errors.ts';
import {isRelType} from '../../rel-type.ts';
import {
  assertWritableDate,
  checkedToken,
  escapeAttr,
  escapeText,
  numAttr,
  textAttr,
  XML_DECLARATION,
} from '../../xml/xml.ts';
import {extensionOf, THEME_PART_PATH} from '../opc/part-paths.ts';
import {imageContentType} from './images.ts';
import {SLICER_CACHES_EXT_URI} from './namespaces.ts';
import type {
  PivotPlan,
  PlannedRelationship,
  PreservedPartPlan,
  PreservedWorkbookReferencePlan,
  TablePlan,
} from './package-plan.ts';
import {
  APP_PROPS_PART,
  commentsPart,
  CORE_PROPS_PART,
  drawingPart,
  PERSONS_PART,
  pivotCacheDefinitionPart,
  pivotCacheRecordsPart,
  pivotTablePart,
  SHARED_STRINGS_PART,
  STYLES_PART,
  tablePart,
  threadedCommentsPart,
  WORKBOOK_PART,
  worksheetPart,
} from './part-names.ts';
import {NS} from './relationships.ts';
import {x14Ext} from './x14-ext.ts';

const CT = {
  workbook: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  macroEnabledWorkbook: 'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
  worksheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
  styles: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
  theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
  core: 'application/vnd.openxmlformats-package.core-properties+xml',
  app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
  table: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
  comments: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
  vml: 'application/vnd.openxmlformats-officedocument.vmlDrawing',
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
  printerSettings: 'application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings',
  sharedStrings: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
  pivotTable: 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml',
  pivotCacheDefinition:
    'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml',
  pivotCacheRecords:
    'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml',
  // Both threaded-comment parts are Microsoft extensions, hence the `vnd.ms-excel` vendor prefix rather
  // than the `openxmlformats-officedocument` one every standard part carries.
  threadedComments: 'application/vnd.ms-excel.threadedcomments+xml',
  person: 'application/vnd.ms-excel.person+xml',
} as const;

// A preserved workbook reference with the relationship id recorded for it (see the writer's
// `planWorkbookRelationships`).
export type PreservedWorkbookRel = PreservedWorkbookReferencePlan & {readonly relId: string};

/**
 * The workbook part's relationships, recorded once by the writer's `planWorkbookRelationships`, with
 * the ids the workbook body cites picked out of them.
 *
 * Declared here because the body is the consumer of the ids: it references them from `<sheet r:id>`.
 * The body used to spell that reference `rId${i + 1}`, which agreed with the allocator only because
 * sheets happen to be the first ids it hands out. Anything laid ahead of them would have re-pointed
 * every sheet at the wrong part, and the package would have stayed schema-valid while doing it -- a
 * class of corruption no validator catches and no round-trip through this library notices, because
 * the reader resolves the same wrong ids consistently.
 */
export interface WorkbookRelPlan {
  readonly sheetRelIds: readonly string[];
  readonly preservedWorkbookRels: readonly PreservedWorkbookRel[];
  /** Every relationship the workbook's `.rels` part declares, in id order. */
  readonly relationships: readonly PlannedRelationship[];
}

/**
 * What `[Content_Types].xml` needs to know about a package: every part family it declares.
 *
 * One object rather than twelve positional parameters, eight of them arrays and four of those bare
 * `readonly number[]`, which two adjacent ones could be silently transposed between. That is the same
 * hazard `SheetReferences`, `SheetPlan` and `WorksheetXmlInputs` remove elsewhere in this writer, and
 * it travels the whole way down: the two functions that render from this take it as it stands rather
 * than re-scattering its fields across their own parameter lists.
 */
export interface ContentTypeInputs {
  readonly sheetCount: number;
  readonly tables: readonly TablePlan[];
  readonly commentNumbers: readonly number[];
  readonly drawingNumbers: readonly number[];
  readonly printerSettingsNumbers: readonly number[];
  readonly mediaExtensions: readonly string[];
  readonly hasSharedStrings: boolean;
  readonly preservedParts: readonly PreservedPartPlan[];
  readonly pivots: readonly PivotPlan[];
  readonly preservedWorkbookRefs: readonly PreservedWorkbookReferencePlan[];
  readonly threadedCommentNumbers: readonly number[];
  readonly hasPersons: boolean;
}

export function contentTypesXml(inputs: ContentTypeInputs): string {
  // One extension→default-content-type map both halves read: the defaults render it, the overrides
  // correct any preserved part whose own type differs from its extension's default. Sharing it is what
  // keeps a `<Default>` and its `<Override>`s from ever disagreeing. It stays a second argument to the
  // overrides rather than joining the inputs: it is derived from them here, and folding an output back
  // into the input bag would let a caller hand over a map that disagrees with the parts beside it.
  const extensionDefaults = buildExtensionDefaults(inputs);
  return (
    XML_DECLARATION +
    `<Types xmlns="${NS.contentTypes}">` +
    contentTypeDefaults(extensionDefaults) +
    contentTypeOverrides(inputs, extensionDefaults) +
    '</Types>'
  );
}

// The content type declared at the `<Default>` level for each file extension the package carries, in
// emission order: the built-in rels/xml pair, the vml/bin/image kinds the writer itself generates,
// then one default per *new* extension a preserved binary part introduces (first part of an extension
// wins). Keys are lower-cased for case-insensitive extension matching; each value keeps the extension
// token as it will be emitted so rendering stays byte-stable. A preserved `.xml` part is deliberately
// left out: it always carries its own type as a per-part override, never a generic `xml` default.
function buildExtensionDefaults(
  inputs: ContentTypeInputs,
): Map<string, {extension: string; contentType: string}> {
  const {commentNumbers, printerSettingsNumbers, mediaExtensions, preservedParts} = inputs;
  const defaults = new Map<string, {extension: string; contentType: string}>();
  const add = (extension: string, contentType: string): void => {
    const key = extension.toLowerCase();
    if (!defaults.has(key)) defaults.set(key, {extension, contentType});
  };
  add('rels', 'application/vnd.openxmlformats-package.relationships+xml');
  add('xml', 'application/xml');
  if (commentNumbers.length > 0) add('vml', CT.vml);
  if (printerSettingsNumbers.length > 0) add('bin', CT.printerSettings);
  for (const ext of mediaExtensions) add(ext, imageContentType(ext));
  for (const part of preservedParts) {
    const ext = extensionOf(part.path);
    if (ext === 'xml') continue;
    add(ext, part.contentType);
  }
  return defaults;
}

// A macro-enabled workbook's own VBA project is round-tripped as a preserved workbook reference
// (see `isPreservedWorkbookRelType` in read.ts), so its presence is what distinguishes a re-emitted
// .xlsm from a plain .xlsx: `xl/workbook.xml` must declare the macro-enabled content type or Excel
// flags the package as needing repair on open.
function isMacroEnabled(preservedWorkbookRefs: readonly PreservedWorkbookReferencePlan[]): boolean {
  return preservedWorkbookRefs.some((ref) => isRelType(ref.relType, 'vbaProject'));
}

// The extension-level `<Default>` declarations, rendered from the shared extension-default map (see
// {@link buildExtensionDefaults}) in insertion order. The raw bytes of a VML, a printer-settings blob,
// or an image carry no XML content type of their own, so each is declared once per extension rather
// than per part; a preserved binary part whose type differs from its extension's default is corrected
// by a per-part `<Override>` in {@link contentTypeOverrides}.
function contentTypeDefaults(
  extensionDefaults: ReadonlyMap<string, {extension: string; contentType: string}>,
): string {
  return [...extensionDefaults.values()]
    .map(({extension, contentType}) => defaultType(extension, contentType))
    .join('');
}

// The per-part `<Override>` declarations, in canonical package order: workbook, worksheets, tables,
// drawings, comments, the threaded-comment parts, each pivot's three parts, theme, styles, the optional
// shared strings and threaded-comment person registry, the doc-props pair, then any preserved part whose
// content type its extension's `<Default>` does not already carry: every `.xml` part (the generic `xml`
// default never matches a real part type) and any binary part sharing an extension with a differently-typed
// sibling (a `vbaProjectSignature.bin` next to a `vbaProject.bin`), which a lone extension default would
// otherwise mis-type.
function contentTypeOverrides(
  inputs: ContentTypeInputs,
  extensionDefaults: ReadonlyMap<string, {extension: string; contentType: string}>,
): string {
  const {
    sheetCount,
    tables,
    drawingNumbers,
    commentNumbers,
    threadedCommentNumbers,
    pivots,
    hasSharedStrings,
    hasPersons,
    preservedParts,
    preservedWorkbookRefs,
  } = inputs;
  const preservedOverrides = preservedParts
    // A preserved theme lands at the fixed theme path, whose override is already in the list below,
    // and OPC forbids declaring the same PartName twice. The fixed declaration is the right one to
    // keep: the part reached the model through a `.../theme` relationship, so its type is the theme
    // type whatever the source package happened to declare.
    .filter((part) => part.path !== THEME_PART_PATH)
    .filter(
      (part) => extensionDefaults.get(extensionOf(part.path))?.contentType !== part.contentType,
    )
    .map((part) => override(part.path, part.contentType));
  return [
    override(
      WORKBOOK_PART,
      isMacroEnabled(preservedWorkbookRefs) ? CT.macroEnabledWorkbook : CT.workbook,
    ),
    ...Array.from({length: sheetCount}, (_, i) => override(worksheetPart(i + 1), CT.worksheet)),
    ...tables.map(({number}) => override(tablePart(number), CT.table)),
    ...drawingNumbers.map((number) => override(drawingPart(number), CT.drawing)),
    ...commentNumbers.map((number) => override(commentsPart(number), CT.comments)),
    ...threadedCommentNumbers.map((number) =>
      override(threadedCommentsPart(number), CT.threadedComments),
    ),
    ...pivots.map(({number}) => override(pivotTablePart(number), CT.pivotTable)),
    ...pivots.map(({number}) =>
      override(pivotCacheDefinitionPart(number), CT.pivotCacheDefinition),
    ),
    ...pivots.map(({number}) => override(pivotCacheRecordsPart(number), CT.pivotCacheRecords)),
    override(THEME_PART_PATH, CT.theme),
    override(STYLES_PART, CT.styles),
    ...(hasSharedStrings ? [override(SHARED_STRINGS_PART, CT.sharedStrings)] : []),
    ...(hasPersons ? [override(PERSONS_PART, CT.person)] : []),
    override(CORE_PROPS_PART, CT.core),
    override(APP_PROPS_PART, CT.app),
    ...preservedOverrides,
  ].join('');
}

function override(partPath: string, contentType: string): string {
  // `PartName` is an OPC part *name*, which is the package-absolute path with a leading slash, where
  // everywhere else in the writer the same part is a bare path. Prefixing here rather than at the
  // twenty call sites is what lets those call sites hand over the very string the part is emitted
  // under, so the declaration and the part cannot drift apart by one character.
  return `<Override PartName="/${escapeAttr(partPath)}" ContentType="${escapeAttr(contentType)}"/>`;
}

// A `<Default>` content-type declaration binding a file extension to the type every part with that
// extension carries: the extension-level counterpart to {@link override}'s per-part declaration.
function defaultType(extension: string, contentType: string): string {
  return `<Default Extension="${escapeAttr(extension)}" ContentType="${escapeAttr(contentType)}"/>`;
}

export function workbookXml(
  workbook: Workbook,
  plan: WorkbookRelPlan,
  pivots: readonly PivotPlan[],
): string {
  const preservedRels = plan.preservedWorkbookRels;
  const sheets = workbook.worksheets;
  const entries = sheets
    .map((sheet, i) => {
      const relId = plan.sheetRelIds[i];
      if (relId === undefined) {
        throw new InternalError(
          `worksheet ${i + 1} of ${sheets.length} has no relationship id: the plan was drawn for ` +
            `${plan.sheetRelIds.length} sheet(s)`,
        );
      }
      const state =
        sheet.state === 'visible'
          ? ''
          : ` state="${checkedToken(sheet.state, isVisibility, 'sheet visibility')}"`;
      return `<sheet name="${escapeAttr(sheet.name)}" sheetId="${sheet.id}"${state} r:id="${relId}"/>`;
    })
    .join('');
  return (
    XML_DECLARATION +
    `<workbook xmlns="${NS.main}" xmlns:r="${NS.docRels}">` +
    workbookPrXml(workbook) +
    workbookProtectionXml(workbook) +
    bookViewsXml(workbook) +
    `<sheets>${entries}</sheets>` +
    externalReferencesXml(preservedRels) +
    definedNamesXml(workbook) +
    calcPrXml(workbook) +
    pivotCachesXml(preservedRels, pivots) +
    workbookExtLstXml(preservedRels) +
    '</workbook>'
  );
}

// `<workbookPr>` opens CT_Workbook's optional content (it precedes `<workbookProtection>`), and
// carries the two workbook-level declarations the model preserves rather than derives: which date
// system its serials count in, and what its VBA project calls it. Omitted when the workbook declares
// neither, so a file that never mentioned it stays byte-clean, and `date1904="0"` is never written:
// that is the schema default, and emitting it would put an attribute into every package Excel writes
// without one.
function workbookPrXml(workbook: Workbook): string {
  const attrs =
    (workbook.dateEpoch === 1904 ? ' date1904="1"' : '') +
    (workbook.codeName === undefined ? '' : ` codeName="${escapeAttr(workbook.codeName)}"`);
  return attrs === '' ? '' : `<workbookPr${attrs}/>`;
}

// `<bookViews>` follows `<workbookProtection>` and precedes `<sheets>` in CT_Workbook order. Unlike
// every other optional block here it is emitted unconditionally: a consumer restores the document
// window from this rect and lays the sheet's panes out inside it, so a package without one leaves a
// frozen split measured against an uninitialised window (see `DEFAULT_WORKBOOK_VIEW`). The
// visibility and minimised flags are written only when they differ from the schema default, and
// `activeTab` only for a non-first sheet: exactly the shape Excel writes for an ordinary window.
function bookViewsXml(workbook: Workbook): string {
  const view = workbook.view;
  const activeTab = workbook.activeTabIndex;
  const attrs =
    (view.visibility === undefined || view.visibility === 'visible'
      ? ''
      : ` visibility="${checkedToken(view.visibility, isVisibility, 'window visibility')}"`) +
    (view.minimized ? ' minimized="1"' : '') +
    numAttr('xWindow', view.x) +
    numAttr('yWindow', view.y) +
    numAttr('windowWidth', view.width) +
    numAttr('windowHeight', view.height) +
    (activeTab === 0 ? '' : ` activeTab="${activeTab}"`);
  return `<bookViews><workbookView${attrs}/></bookViews>`;
}

// The `<externalReferences>` element registers each preserved external link (a link to a source
// workbook), wired to the relationship reaching its `externalLink` part. It follows `<sheets>` in
// CT_Workbook order and precedes `<definedNames>`; its child order is the `[n]` index a formula or
// defined name resolves an external cell through, so entries are emitted by their captured
// `externalReferenceIndex`, keeping every `[1]`, `[2]`, … pointing at the same linked workbook it did
// before the round-trip. '' when no external link was preserved.
function externalReferencesXml(preservedRels: readonly PreservedWorkbookRel[]): string {
  const links = preservedRels
    .filter((ref) => ref.externalReferenceIndex !== undefined)
    .sort((a, b) => (a.externalReferenceIndex ?? 0) - (b.externalReferenceIndex ?? 0));
  if (links.length === 0) return '';
  const entries = links.map((ref) => `<externalReference r:id="${ref.relId}"/>`).join('');
  return `<externalReferences>${entries}</externalReferences>`;
}

// The workbook-body `<x14:slicerCaches>` extension that registers each preserved slicer cache, wired
// to the relationship reaching its cache part. Slicer caches (unlike pivot caches, which register in
// `<pivotCaches>`) live only in this extension block, so re-emitting it is what lets Excel rediscover
// the slicers. `<extLst>` is the final child of CT_Workbook. '' when no slicer cache was preserved.
function workbookExtLstXml(preservedRels: readonly PreservedWorkbookRel[]): string {
  const caches = preservedRels.filter((ref) => isRelType(ref.relType, 'slicerCache'));
  if (caches.length === 0) return '';
  const entries = caches.map((ref) => `<x14:slicerCache r:id="${ref.relId}"/>`).join('');
  return `<extLst>${x14Ext(SLICER_CACHES_EXT_URI, `<x14:slicerCaches>${entries}</x14:slicerCaches>`)}</extLst>`;
}

// The `<pivotCaches>` element registers each pivot cache under the `cacheId` a pivot table resolves
// its cache through, wired to the relationship that reaches the cache definition. It follows
// `<calcPr>` in CT_Workbook order and carries both preserved caches (passed through from a read file)
// and caches the writer generated for modeled pivot tables. A slicer cache (no `cacheId`) is
// registered in a workbook extension block, not here, so it is skipped.
function pivotCachesXml(
  preservedRels: readonly PreservedWorkbookRel[],
  pivots: readonly PivotPlan[],
): string {
  const preserved = preservedRels.flatMap((ref) =>
    ref.pivotCacheId === undefined
      ? []
      : [`<pivotCache cacheId="${escapeAttr(ref.pivotCacheId)}" r:id="${ref.relId}"/>`],
  );
  const generated = pivots.map(
    (pivot) => `<pivotCache cacheId="${escapeAttr(pivot.cacheId)}" r:id="${pivot.workbookRelId}"/>`,
  );
  const entries = [...preserved, ...generated];
  if (entries.length === 0) return '';
  return `<pivotCaches>${entries.join('')}</pivotCaches>`;
}

// `<workbookProtection>` precedes `<sheets>` in CT_Workbook order. It re-emits the workbook's
// structure/window lock flags (each written only when true, so an unlocked aspect stays absent) and
// the preserved password/agile-hash credential attributes verbatim. Emitted only when the workbook
// actually declares protection, either the flags or a credential, so an unprotected workbook stays clean.
function workbookProtectionXml(workbook: Workbook): string {
  const p = workbook.protection;
  if (p === undefined) return '';
  const attrs: string[] = [];
  if (p.lockStructure) attrs.push('lockStructure="1"');
  if (p.lockWindows) attrs.push('lockWindows="1"');
  if (p.lockRevision) attrs.push('lockRevision="1"');
  for (const key of WORKBOOK_PROTECTION_CREDENTIAL_ATTRS) {
    const value = p.credentials?.[key];
    if (value !== undefined) attrs.push(`${key}="${escapeAttr(value)}"`);
  }
  if (attrs.length === 0) return '';
  return `<workbookProtection ${attrs.join(' ')}/>`;
}

// The `calcId` desktop Excel stamps into `<calcPr>`: the build number of the calc engine that last
// evaluated the workbook. Consumers only compare it to decide whether cached results are stale; we
// emit a fixed recent value alongside `fullCalcOnLoad`, which already forces a full recalculation.
const EXCEL_CALC_ID = '171027';

// `<calcPr>` follows `<definedNames>` in CT_Workbook order and carries the calculation settings.
// Today the model exposes a single one: `fullCalcOnLoad`, which tells the consumer to recalculate
// every formula on open instead of trusting the cached results. Emitted only when set, so an
// unmarked workbook keeps the element (and its `calcId`) out of the file entirely.
function calcPrXml(workbook: Workbook): string {
  return workbook.fullCalcOnLoad ? `<calcPr calcId="${EXCEL_CALC_ID}" fullCalcOnLoad="1"/>` : '';
}

// The `<definedNames>` block follows `<sheets>` in the schema. A sheet-scoped name carries a
// `localSheetId`, the 0-based position of its sheet among the `<sheet>` entries, NOT the sheet's
// own id, so the index is resolved against the worksheet order here. The refersTo formula is the
// element's text content, run through the same function mangling the writer applies to a cell formula
// so a name defined as a modern function (a LAMBDA, an XLOOKUP-based name) is stored under the prefix
// Excel requires; a plain reference has no function call and passes through untouched. A function the
// formula passes as a value resolves against the names visible from the name's own scope. Only names
// that are actually set emit anything.
// The 0-based position of a scoped name's sheet among the `<sheet>` entries.
//
// Matched case-insensitively, because that is how a sheet name is identified everywhere else: sheet
// names are unique case-insensitively, `Workbook.getWorksheet` resolves that way, and so does
// `defineName`'s own scope check. An exact match here made `defineName` accept a scope this then
// could not resolve.
//
// A scope with no sheet at all is refused rather than written. `definedNames` is a live collection, so
// a name can be pushed onto it without passing `defineName`'s check, and `localSheetId` is an
// `xsd:unsignedInt` for which the `-1` a lookup miss returns is a package Excel offers to repair.
function localSheetId(sheets: readonly Worksheet[], name: string, scope: string): number {
  const wanted = scope.toLowerCase();
  const index = sheets.findIndex((sheet) => sheet.name.toLowerCase() === wanted);
  if (index === -1) {
    throw new AuthoringError(
      `defined name ${quoted(name)} is scoped to worksheet ${quoted(scope)}, which this workbook does not contain`,
    );
  }
  return index;
}

function definedNamesXml(workbook: Workbook): string {
  const sheets = workbook.worksheets;
  const userEntries = workbook.definedNames.map((name) => {
    const scopeAttr =
      name.scope === undefined
        ? ''
        : ` localSheetId="${localSheetId(sheets, name.name, name.scope)}"`;
    const commentAttr = textAttr('comment', name.comment);
    const hiddenAttr = name.hidden ? ' hidden="1"' : '';
    return (
      `<definedName name="${escapeAttr(name.name)}"${scopeAttr}${commentAttr}${hiddenAttr}>` +
      `${escapeText(mangleFormula(name.refersTo, formulaNamesInScope(workbook.definedNames, name.scope)))}</definedName>`
    );
  });
  // Every sheet-level autofilter contributes the hidden, sheet-scoped `_FilterDatabase` built-in that
  // Excel derives from its range. The reader drops these on load and rebuilds them from the sheet's
  // `<autoFilter>`, so `Worksheet.autoFilter` stays the single source of truth and a round-trip never
  // duplicates them.
  const filterEntries = sheets.flatMap((sheet, index) =>
    sheet.autoFilter === undefined
      ? []
      : [
          `<definedName name="_xlnm._FilterDatabase" localSheetId="${index}" hidden="1">` +
            `${escapeText(filterDatabaseRefersTo(sheet.name, sheet.autoFilter.ref))}</definedName>`,
        ],
  );

  const entries = [...userEntries, ...filterEntries];
  if (entries.length === 0) return '';
  return `<definedNames>${entries.join('')}</definedNames>`;
}

// Build the sheet-qualified, fully-absolute reference a `_FilterDatabase` name carries
// (`'Sheet 1'!$A$1:$C$10`) from a sheet name and its already-canonical `A1:C10` autofilter range.
function filterDatabaseRefersTo(sheetName: string, range: string): string {
  const absolute = range.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2');
  return `${quoteSheetName(sheetName)}!${absolute}`;
}

export function corePropsXml(properties: WorkbookProperties): string {
  const parts: string[] = [];
  // `dc:title` precedes `dc:creator`: `cp:coreProperties` is a sequence, not a bag, and Excel
  // repairs a file whose children are out of schema order.
  if (properties.title !== undefined) {
    parts.push(`<dc:title>${escapeText(properties.title)}</dc:title>`);
  }
  if (properties.creator !== undefined) {
    parts.push(`<dc:creator>${escapeText(properties.creator)}</dc:creator>`);
  }
  if (properties.lastModifiedBy !== undefined) {
    parts.push(`<cp:lastModifiedBy>${escapeText(properties.lastModifiedBy)}</cp:lastModifiedBy>`);
  }
  if (properties.created !== undefined) {
    parts.push(w3cdtf('created', properties.created));
  }
  if (properties.modified !== undefined) {
    parts.push(w3cdtf('modified', properties.modified));
  }
  return (
    XML_DECLARATION +
    `<cp:coreProperties xmlns:cp="${NS.coreProps}" xmlns:dc="${NS.dc}" xmlns:dcterms="${NS.dcterms}" ` +
    `xmlns:dcmitype="${NS.dcmitype}" xmlns:xsi="${NS.xsi}">` +
    parts.join('') +
    '</cp:coreProperties>'
  );
}

// The millisecond field is trimmed because Excel writes seconds and nothing finer, and a document
// property is a timestamp a person reads, not an interval anything measures. `dcterms:W3CDTF` admits
// both forms, so this is a match with the producer rather than a correctness fix.
function w3cdtf(element: string, date: Date): string {
  assertWritableDate(date, `the document property ${quoted(element)}`);
  const stamp = date.toISOString().replace(/\.\d{3}Z$/, 'Z');
  return `<dcterms:${element} xsi:type="dcterms:W3CDTF">${stamp}</dcterms:${element}>`;
}

export function appPropsXml(properties: WorkbookProperties): string {
  // `Company` follows `Application` in the extended-properties sequence. It is the one document
  // property OOXML keeps in this part rather than in core.xml, which is why it arrives here.
  const company =
    properties.company === undefined ? '' : `<Company>${escapeText(properties.company)}</Company>`;
  return (
    XML_DECLARATION +
    `<Properties xmlns="${NS.extProps}"><Application>ts-xlsx</Application>${company}</Properties>`
  );
}
