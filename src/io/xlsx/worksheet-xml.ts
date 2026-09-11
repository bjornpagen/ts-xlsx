// Worksheet serialisation: a Worksheet model into its `xl/worksheets/sheetN.xml` part (the sheet's rels
// part is the writer's relationship ledger for the sheet). Orchestrates the whole sheet body: the
// print/page/view/protection blocks live in `sheet-properties.ts`, shared-formula planning in
// `shared-formulas.ts`, and rendering one row and its cells in `row-xml.ts`, each imported here rather
// than duplicated. Table *parts*
// (`xl/tables/tableN.xml`) are `tables.ts`'s concern, alongside their reader; this module only wires
// the sheet's `<tableParts>` back-references to them.
//
// The seam with `row-xml.ts` is not length. A row can be serialised the moment it is committed, with
// no sheet around it, which is exactly what the streaming writer does; everything here needs the whole
// sheet. The two halves shared nothing but the style registry, and `row-xml.ts` is precisely the
// surface that writer imports, which `write.ts` used to re-export on its behalf.

import {encodeRect} from '../../core/address.ts';
import type {DateEpoch} from '../../core/date.ts';
import {mergesOverlappingTables} from '../../core/merge.ts';
import {pickStyleFacets} from '../../core/style.ts';
import type {ColumnProperties, Worksheet, WorksheetProperties} from '../../core/worksheet.ts';
import {AuthoringError, InternalError, quoted} from '../../errors.ts';
import {escapeAttr, numberText, XML_DECLARATION} from '../../xml/xml.ts';
import {
  conditionalFormattingsExtXml,
  conditionalFormattingsXml,
  type DataBarExtLinks,
  dataBarExtLinks,
} from './conditional-formatting.ts';
import {dataValidationsExtXml, dataValidationsXml} from './data-validation.ts';
import {type HyperlinkPlan, hyperlinksXml} from './hyperlinks.ts';
import {SLICER_LIST_EXT_URI} from './namespaces.ts';
import type {TablePlan} from './package-plan.ts';
import {NS} from './relationships.ts';
import {
  assertWritableLevel,
  buildColumnDefaults,
  Extent,
  type FlushedRow,
  type FlushedSheet,
  outlineAttr,
  renderRow,
  rowOpenTag,
  type RowRenderContext,
} from './row-xml.ts';
import {planSharedFormulas} from './shared-formulas.ts';
import type {SharedStringTable} from './shared-strings.ts';
import {
  autoFilterXml,
  headerFooterXml,
  pageBreaksXml,
  pageMarginsXml,
  pageSetupXml,
  printOptionsXml,
  sheetProtectionXml,
  sheetPrXml,
  sheetViewsXml,
} from './sheet-properties.ts';
import type {StyleRegistry} from './styles.ts';
import {x14Ext} from './x14-ext.ts';

// The sheet-local relationship ids that wire a worksheet's tail elements to their parts, gathered into
// one struct so the caller hands them over as a named unit rather than a run of positional `string |
// null` arguments a mis-ordered call could silently transpose. Each id is `null` when the sheet carries
// no part of that kind. `slicerRelIds` is plural because a sheet may carry several preserved slicers.
export interface SheetReferences {
  readonly drawingRelId: string | null;
  readonly legacyDrawingRelId: string | null;
  readonly printerSettingsRelId: string | null;
  readonly backgroundRelId: string | null;
  readonly legacyDrawingHFRelId: string | null;
  readonly slicerRelIds: readonly string[];
}

/**
 * Everything one worksheet part is rendered from: the sheet itself, the parts of the package plan
 * that belong to it, and the two registries a sheet pass interns into.
 *
 * One object for the reason {@link SheetReferences} is one object, applied to the layer above it: a
 * run of positional arguments carrying two `readonly …[]`s, two nullable references and a bare
 * boolean is a run whose order the call site cannot be read against. `active:` is the clearest of
 * them -- as a positional it needed a comment at the call site to say which boolean it was.
 */
export interface WorksheetXmlInputs {
  readonly sheet: Worksheet;
  readonly tables: readonly TablePlan[];
  readonly styles: StyleRegistry;
  readonly references: SheetReferences;
  readonly hyperlinks: readonly HyperlinkPlan[];
  readonly sharedStrings: SharedStringTable | null;
  readonly dateEpoch: DateEpoch;
  /** Whether this is the workbook's one selected sheet. */
  readonly active: boolean;
  /** The rows the streaming writer already serialised and evicted, absent on the buffered path. */
  readonly flushed?: FlushedSheet | undefined;
}

export function worksheetXml(inputs: WorksheetXmlInputs): string {
  const {sheet, tables, styles, references, hyperlinks, sharedStrings, dateEpoch, active, flushed} =
    inputs;
  // A merge overlapping a table is Excel-invalid geometry; reject it before serialising
  // rather than emit a package a consumer repairs on open.
  validateMerges(sheet);

  const columnDefaults = buildColumnDefaults(sheet);

  // A cell filled from a shared formula is written as a master (seeding the group) or a clone
  // (referencing it by shared index); resolve every such role before the row loop so each cell knows
  // how to serialise its `<f>`. This also validates the master/clone geometry, throwing if a clone
  // precedes its master or its master carries no formula.
  const sharedRoles = planSharedFormulas(sheet);
  // One link map for the whole sheet, handed to both conditional-formatting passes. Built here beside
  // the shared-formula plan for the same reason: it is a fact about the sheet that two serialisers
  // must agree on, and two of them deriving it separately is agreement by coincidence.
  const extLinks = dataBarExtLinks(sheet.conditionalFormattings);

  // A fully-hidden outline group's collapse toggle belongs on its summary row; derive that set once
  // so the row loop can stamp it even onto a summary row that carries no properties of its own. The
  // same pass yields the sheet's deepest row outline level for `<sheetFormatPr>`.
  const rowOutline = scanRowOutline(sheet, flushed);
  const collapsedSummaries = rowOutline.collapsedSummaries;

  const context: RowRenderContext = {
    columnDefaults,
    styles,
    sharedStrings,
    sharedRoles,
    collapsedSummaries,
    dateEpoch,
  };

  const liveRows: {number: number; xml: string}[] = [];
  // Seed the used-cell extent with any rows the streaming writer already serialised and evicted, so
  // the dimension spans both them and the live rows below.
  const extent = new Extent(flushed?.extent);

  for (const entry of sheet.rows()) {
    const {xml, minCol, maxCol} = renderRow(entry, context);
    if (xml === '') continue;
    liveRows.push({number: entry.number, xml});
    extent.add(entry.number, minCol, maxCol);
  }

  const dimensionRef = extent.isEmpty ? 'A1' : encodeRect(extent);
  // Merge the streaming writer's pre-rendered rows with the live ones into ascending row order. A
  // flushed row can carry any number, and rows may be committed out of order. The buffered path has no
  // flushed rows, so it skips the merge and its sort entirely.
  const orderedRows = flushed
    ? [...flushed.rows.map((row) => completeCollapsed(row, collapsedSummaries)), ...liveRows].sort(
        (a, b) => a.number - b.number,
      )
    : liveRows;
  const bodyXml = orderedRows.map((row) => row.xml).join('');
  const sheetData = bodyXml === '' ? '<sheetData/>' : `<sheetData>${bodyXml}</sheetData>`;

  return (
    XML_DECLARATION +
    `<worksheet xmlns="${NS.main}" xmlns:r="${NS.docRels}">` +
    sheetPrXml(sheet) +
    `<dimension ref="${dimensionRef}"/>` +
    sheetViewsXml(sheet.view, active) +
    sheetFormatPr(sheet.properties, {
      col: maxColumnOutlineLevel(sheet),
      // A streamed sheet's flushed rows are gone from the model; their deepest level rides along on
      // the flush record so the header still reports the whole sheet's outline.
      row: rowOutline.maxLevel,
    }) +
    colsXml(sheet, styles) +
    sheetData +
    sheetProtectionXml(sheet.protection) +
    // CT_Worksheet order: <autoFilter> follows <sheetProtection> (and the scenarios block) and
    // precedes <mergeCells>. Its `_FilterDatabase` companion is emitted in the workbook part.
    autoFilterXml(sheet.autoFilter) +
    mergeCellsXml(sheet.merges) +
    // CT_Worksheet order: <conditionalFormatting> blocks follow <mergeCells>, then <dataValidations>,
    // then <hyperlinks>, all of which precede the print settings.
    conditionalFormattingsXml(sheet.conditionalFormattings, styles, extLinks) +
    dataValidationsXml(sheet.dataValidations) +
    hyperlinksXml(hyperlinks) +
    // CT_Worksheet order: <printOptions> precedes <pageMargins>, which precedes <pageSetup>.
    printOptionsXml(sheet.printOptions) +
    pageMarginsXml(sheet.pageMargins) +
    pageSetupXml(sheet.pageSetup, references.printerSettingsRelId) +
    headerFooterXml(sheet.headerFooter) +
    // CT_Worksheet order: <rowBreaks> follows <headerFooter>, <colBreaks> follows <rowBreaks>, and
    // both precede the drawing block.
    pageBreaksXml(sheet.rowBreaks, 'rowBreaks') +
    pageBreaksXml(sheet.columnBreaks, 'colBreaks') +
    // Schema order near the tail: <drawing> (the images), then <legacyDrawing> (the VML holding the
    // note boxes), then <legacyDrawingHF> (a preserved header/footer image's VML), then <picture>
    // (the sheet background), then <tableParts>.
    refElement('drawing', references.drawingRelId) +
    refElement('legacyDrawing', references.legacyDrawingRelId) +
    refElement('legacyDrawingHF', references.legacyDrawingHFRelId) +
    refElement('picture', references.backgroundRelId) +
    tablePartsXml(tables) +
    // `<extLst>` is the final child of CT_Worksheet and a worksheet may carry at most one. Both the
    // x14 conditional-formatting extensions (data-bar gradient/negative-fill/axis) and the extended
    // (x14) data validations ride inside it as sibling `<ext>` blocks, so they are gathered here into
    // a single `<extLst>` rather than each emitting its own.
    worksheetExtLstXml(sheet, references.slicerRelIds, extLinks) +
    '</worksheet>'
  );
}

// Assemble the worksheet's single `<extLst>` from every x14 extension the sheet carries, or '' when it
// carries none. Each producer returns a bare `<ext>` so they compose without nesting an `<extLst>`.
function worksheetExtLstXml(
  sheet: Worksheet,
  slicerRelIds: readonly string[],
  extLinks: DataBarExtLinks,
): string {
  const exts = [
    conditionalFormattingsExtXml(sheet.conditionalFormattings, extLinks),
    dataValidationsExtXml(sheet.dataValidations),
    slicerListExtXml(slicerRelIds),
  ].filter((ext) => ext !== '');
  return exts.length === 0 ? '' : `<extLst>${exts.join('')}</extLst>`;
}

// The worksheet-body `<x14:slicerList>` extension that reconnects a sheet to its preserved slicer
// parts. Each `<x14:slicer>` names the sheet-local relationship id its slicer rel was re-emitted under,
// so the wiring stays consistent even though the id is reassigned on write. '' when the sheet has none.
function slicerListExtXml(slicerRelIds: readonly string[]): string {
  if (slicerRelIds.length === 0) return '';
  const slicers = slicerRelIds.map((relId) => `<x14:slicer r:id="${relId}"/>`).join('');
  return x14Ext(SLICER_LIST_EXT_URI, `<x14:slicerList>${slicers}</x14:slicerList>`);
}

// Excel forbids a merged range from intersecting a formatted table; such a file opens as
// corrupt. The writer is the OOXML gatekeeper for this cross-feature geometry conflict.
function validateMerges(sheet: Worksheet): void {
  const [conflict] = mergesOverlappingTables(sheet.merges, sheet.tables);
  if (conflict === undefined) return;
  const {merge, table} = conflict;
  throw new AuthoringError(
    `merged range ${merge} overlaps table ${quoted(table.name)} (${table.range}): Excel forbids a merge inside a table`,
  );
}

// The model stores every merge in canonical form, so the ranges are written as they are held.
function mergeCellsXml(merges: readonly string[]): string {
  if (merges.length === 0) return '';
  const cells = merges.map((range) => `<mergeCell ref="${escapeAttr(range)}"/>`).join('');
  return `<mergeCells count="${merges.length}">${cells}</mergeCells>`;
}

// A tail reference element (`<drawing r:id="…"/>` and its `<legacyDrawing>`/`<legacyDrawingHF>`/
// `<picture>` siblings) wiring the sheet to a part by relationship id, or '' when the sheet carries no
// part of that kind, in which case each such id is null.
function refElement(tag: string, relId: string | null): string {
  return relId === null ? '' : `<${tag} r:id="${relId}"/>`;
}

function tablePartsXml(tables: readonly TablePlan[]): string {
  if (tables.length === 0) return '';
  const parts = tables.map(({relId}) => `<tablePart r:id="${relId}"/>`).join('');
  return `<tableParts count="${tables.length}">${parts}</tableParts>`;
}

// Excel's standard row height in points, emitted as the `defaultRowHeight` when the sheet does not
// override it so a reader sees the same baseline Excel would write.
const DEFAULT_ROW_HEIGHT = 15;

// `<sheetFormatPr>` carries the sheet's grid defaults and, when the sheet groups anything, the depth
// of its deepest outline. A consumer sizes the outline bars from those depths, the strips that sit
// above the column headers and left of the row headers, so a grouped sheet that omits them lays its
// grid out with no room reserved for a bar it then has to draw. Both are omitted at zero, as Excel
// does, so an ungrouped sheet stays byte-clean.
function sheetFormatPr(
  properties: WorksheetProperties,
  outlineLevel: {readonly col: number; readonly row: number},
): string {
  const rowHeight = properties.defaultRowHeight ?? DEFAULT_ROW_HEIGHT;
  let attrs = ` defaultRowHeight="${numberText(rowHeight)}"`;
  if (properties.defaultColWidth !== undefined) {
    attrs += ` defaultColWidth="${numberText(properties.defaultColWidth)}"`;
  }
  // A non-standard default row height is only honoured by Excel when customHeight is set.
  if (properties.defaultRowHeight !== undefined) attrs += ' customHeight="1"';
  attrs += outlineAttr('outlineLevelCol', outlineLevel.col);
  attrs += outlineAttr('outlineLevelRow', outlineLevel.row);
  return `<sheetFormatPr${attrs}/>`;
}

// The deepest column outline level the sheet declares: the `outlineLevelCol` its `<sheetFormatPr>`
// reports.
function maxColumnOutlineLevel(sheet: Worksheet): number {
  let max = 0;
  for (const column of sheet.columns()) max = Math.max(max, column.outlineLevel ?? 0);
  return max;
}

function colsXml(sheet: Worksheet, styles: StyleRegistry): string {
  // Runs of adjacent columns that carry identical definitions are coalesced into a single
  // `<col min max>` span. Excel writes columns this way, and it keeps the part compact for a
  // sheet whose columns share a width or outline level. A gap in the indices or any difference
  // in the emitted attributes breaks the run.
  const runs: {min: number; max: number; body: string}[] = [];
  for (const {index, properties} of sheet.columns()) {
    const body = colBody(properties ?? {}, styles);
    // A <col> with no width, visibility, or style says nothing; omit it entirely. That also covers
    // the column with no format record at all, which `columns()` does not in fact yield.
    if (body === null) continue;
    const last = runs[runs.length - 1];
    if (last !== undefined && last.max === index - 1 && last.body === body) {
      last.max = index;
    } else {
      runs.push({min: index, max: index, body});
    }
  }
  if (runs.length === 0) return '';
  const cols = runs.map((run) => `<col min="${run.min}" max="${run.max}"${run.body}/>`).join('');
  return `<cols>${cols}</cols>`;
}

// The attributes of a `<col>` sans its `min`/`max` span (each with a leading space), or `null` when
// the column declares nothing worth emitting. Two columns with the same body are interchangeable, so
// the body doubles as the equivalence key that {@link colsXml} coalesces adjacent runs by.
function colBody(properties: ColumnProperties, styles: StyleRegistry): string | null {
  let attrs = '';
  let meaningful = false;
  if (properties.width !== undefined) {
    attrs += ` width="${numberText(properties.width)}" customWidth="1"`;
    meaningful = true;
  }
  if (properties.hidden) {
    attrs += ' hidden="1"';
    meaningful = true;
  }
  const outline = outlineAttr('outlineLevel', properties.outlineLevel);
  if (outline !== '') {
    attrs += outline;
    meaningful = true;
  }
  if (properties.collapsed) {
    attrs += ' collapsed="1"';
    meaningful = true;
  }
  // The column's style facets are carried as its own `<col>` style; its populated cells inherit
  // them via the composition above, and this `style` makes Excel apply them to the column's empty
  // cells too.
  const style = styles.styleId(pickStyleFacets(properties));
  if (style !== 0) {
    attrs += ` style="${style}"`;
    meaningful = true;
  }
  return meaningful ? attrs : null;
}

// The two row-outline facts the serialiser needs, from one walk over the rows: which summary rows
// terminate a fully-collapsed group, and how deep the sheet's grouping goes. They share a pass
// because the pass is the expensive part: the streaming writer must not be made to traverse rows
// twice just to fill in a header attribute.
interface RowOutline {
  readonly collapsedSummaries: Set<number>;
  readonly maxLevel: number;
}

// A collapsed outline group is two coordinated facts: its detail rows carry outlineLevel and are
// hidden, AND the summary row that terminates the group carries `collapsed`. Authors typically set
// only outlineLevel + hidden on the detail rows, so the summary flag is derived here rather than
// demanded of the caller: a row is a collapsed summary iff its adjacent detail run, the contiguous
// higher-outline-level rows on the summary side, is non-empty and every row in it is hidden.
// Placement follows the sheet's summaryBelow flag (Excel's default is summary below the detail); the
// walk stops at the first row of level <= the summary's own, so a gap or a boundary ends the group.
/**
 * Add `collapsed="1"` to a flushed summary row whose whole detail group turned out to be hidden.
 *
 * Every other row attribute is a fact about that row alone, so a row can be rendered the moment it is
 * committed; this one is a fact about the rows *after* it, which the row cannot know when it flushes.
 * Rather than refuse an outlined row on a streamed sheet, or leave the group silently rendering
 * expanded, the one attribute that needs the look-ahead is added once the look-ahead is possible.
 *
 * Both halves of the patch are answered by the flushed row's own fields rather than by reading its
 * markup back. *Is the attribute already there* is `attrs`, because the row's `xml` also holds cell
 * text and `escapeText` leaves a double quote verbatim: a cell whose value was the literal string
 * ` collapsed="1"` used to convince the patcher the row already carried the attribute, and the
 * outline group rendered expanded. *Where does it go* is {@link rowOpenTag}, the emitter's own
 * spelling of the tag this row began with, so the splice point is shared rather than re-derived by
 * scanning for a `>` that a cell value could also supply.
 */
function completeCollapsed(
  row: FlushedRow,
  collapsedSummaries: ReadonlySet<number>,
): {number: number; xml: string} {
  if (!collapsedSummaries.has(row.number) || row.attrs.includes(' collapsed="1"')) {
    return {number: row.number, xml: row.xml};
  }
  const open = rowOpenTag(row.number);
  if (!row.xml.startsWith(open)) {
    throw new InternalError(`flushed row ${row.number} does not open with ${quoted(open)}`);
  }
  return {
    number: row.number,
    xml: `${open} collapsed="1"${row.xml.slice(open.length)}`,
  };
}

function scanRowOutline(sheet: Worksheet, flushed: FlushedSheet | undefined): RowOutline {
  const level = new Map<number, number>();
  const hidden = new Map<number, boolean>();
  let maxLevel = 0;
  const note = (number: number, rowLevel: number, rowHidden: boolean): void => {
    // Refused here and not only where the attribute is written, because the walk below compares
    // levels to find a group's end: against a negative level every comparison holds and the walk runs
    // off the sheet forever, and against `NaN` none does and the group ends before it starts.
    assertWritableLevel(`row ${number} outlineLevel`, rowLevel);
    level.set(number, rowLevel);
    hidden.set(number, rowHidden);
    if (rowLevel > maxLevel) maxLevel = rowLevel;
  };
  // The flushed rows first, so a live row carrying the same number (which cannot happen, but the map
  // has to answer something) wins, matching the row merge below.
  for (const [number, entry] of flushed?.rowOutline ?? []) {
    note(number, entry.outlineLevel, entry.hidden);
  }
  for (const {number, properties} of sheet.rows()) {
    note(number, properties?.outlineLevel ?? 0, properties?.hidden ?? false);
  }
  const levelOf = (row: number): number => level.get(row) ?? 0;
  const step = sheet.outline.summaryBelow === false ? 1 : -1;
  const collapsedSummaries = new Set<number>();
  for (const [summary, summaryLevel] of level) {
    let detail = summary + step;
    let sawDetail = false;
    let allHidden = true;
    // The level map is what bounds this walk: it ends at the first row outside the group, and a row
    // nobody noted is outside every group whatever the arithmetic above it says.
    while (level.get(detail) !== undefined && levelOf(detail) > summaryLevel) {
      sawDetail = true;
      if (!hidden.get(detail)) allHidden = false;
      detail += step;
    }
    if (sawDetail && allHidden) collapsedSummaries.add(summary);
  }
  return {collapsedSummaries, maxLevel};
}
