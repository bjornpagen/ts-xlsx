// The worksheet-body reader: one `xl/worksheets/sheetN.xml` in, its cells/columns/rows applied onto
// a {@link Worksheet}. It is a single streaming pass driving a handful of accumulators (the cell
// being read, shared-formula masters, an autofilter draft, the current page-break axis) so each
// element commits its state as it closes. Style indices resolve through the parsed style table.
//
// What this module *is* is the dispatch and the cell body. The layout and print blocks it dispatches
// to (`<sheetPr>`, `<sheetView>`, `<sheetProtection>`, the print settings, the page breaks) are read
// by `sheet-properties.ts`, beside the writer that emits them.

import {boundedRect, tryDecodeRange} from '../../core/address.ts';
import {
  type CustomFilterPredicate,
  type FilterColumn,
  type FilterCriteria,
  isCustomFilterOperator,
} from '../../core/autofilter.ts';
import type {DateEpoch} from '../../core/date.ts';
import {INTERNAL} from '../../core/internal.ts';
import {HEADER_FOOTER_ELEMENTS} from '../../core/page-setup.ts';
import {assignStyleFacets} from '../../core/style.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {decodeSpreadsheetText, numFinite, numInteger} from '../../xml/xml-attrs.ts';
import {type SaxHandlers, type SaxPass, TextCapture} from '../../xml/xml-read.ts';
import {boolPresent, boolStrict, localName, type XmlAttributes} from '../../xml/xml-scan.ts';
import {ColumnRecordBudget} from '../read-policy/column-budget.ts';
import {admitting} from '../read-policy/read-repair.ts';
import {CellStyleResolver} from '../style/cell-style-resolution.ts';
import type {XfStyle} from '../style/xf-style.ts';
import {CellAccumulator, WORKSHEET_BODY_EMPTY_CLOSES} from './cell-accumulator.ts';
import type {SharedString} from './cell-value.ts';
import {takeColumnSpan} from './column-span.ts';
import {RowPositionTracker} from './row-position.ts';
import {
  applySheetProperties,
  PageBreakAccumulator,
  parseSheetProtection,
} from './sheet-properties.ts';

// Membership, not order: the reader meets a `<headerFooter>` child by name and needs only to know
// whether it is one, on both the open (start capturing) and the close (commit). The order the tuple
// carries is the writer's concern.
const HEADER_FOOTER_CHILDREN: ReadonlySet<string> = new Set(HEADER_FOOTER_ELEMENTS);

function isHeaderFooterElement(local: string): local is (typeof HEADER_FOOTER_ELEMENTS)[number] {
  return HEADER_FOOTER_CHILDREN.has(local);
}

// The body elements both readers expand to open+close, plus a criteria-free self-closing
// `<autoFilter/>`, which this reader also commits from its close.
const WORKSHEET_EMPTY_CLOSES: ReadonlySet<string> = new Set([
  ...WORKSHEET_BODY_EMPTY_CLOSES,
  'autoFilter',
]);

// Fold a filter column's accumulated `<filters>` or `<customFilters>` state into one criteria value,
// or null when it carried nothing filterable. A `<filters>` block with no values and no blank flag,
// or a `<customFilters>` with no predicates, is a no-op that would round-trip as noise, so it drops.
function pendingFilterCriteria(
  values: string[] | null,
  blank: boolean,
  predicates: CustomFilterPredicate[] | null,
  and: boolean,
): FilterCriteria | null {
  if (values !== null && (values.length > 0 || blank)) {
    return {kind: 'values', values, blank};
  }
  if (predicates !== null && predicates.length > 0) {
    return {kind: 'custom', and, predicates: predicates.slice(0, 2)};
  }
  return null;
}

// Autofilter accumulation. The sheet `<autoFilter ref>` seeds a draft; each `<filterColumn colId>`
// opens a column whose criteria (`<filters>` values or `<customFilters>` predicates) stream into the
// draft until `</filterColumn>`, and `</autoFilter>` commits the whole thing to the sheet.
class AutoFilterAccumulator {
  #ref: string | null = null;
  #columns: FilterColumn[] = [];
  #colId = -1;
  #values: string[] | null = null;
  #blank = false;
  #predicates: CustomFilterPredicate[] | null = null;
  #and = false;

  // Seed a draft from the range. (A table's own `<autoFilter>` also matches, but table sheets route
  // through parseTable, so this only ever sees the sheet-level one.)
  begin(attrs: XmlAttributes): void {
    this.#ref = attrs.ref !== undefined && attrs.ref !== '' ? attrs.ref : null;
    this.#columns = [];
  }

  // Open a criteria block for one column, offset `colId` from the range's left edge. Reset the
  // per-column accumulators; whichever child (`<filters>`/`<customFilters>`) opens fills one.
  beginColumn(attrs: XmlAttributes): void {
    this.#colId = numInteger(attrs.colId, 0) ?? -1;
    this.#values = null;
    this.#blank = false;
    this.#predicates = null;
    this.#and = false;
  }

  beginValues(attrs: XmlAttributes): void {
    this.#values = [];
    // `blank` defaults off when absent, so presence is required first; then it reads as an
    // on-when-present flag. (`boolPresent` alone would treat the absent attribute as on.)
    this.#blank = attrs.blank !== undefined && boolPresent(attrs.blank);
  }

  addValue(attrs: XmlAttributes): void {
    if (this.#values !== null && attrs.val !== undefined) this.#values.push(attrs.val);
  }

  beginCustom(attrs: XmlAttributes): void {
    this.#predicates = [];
    this.#and = attrs.and !== undefined && boolPresent(attrs.and);
  }

  // The operator attribute defaults to `equal` when absent (per CT_CustomFilter); an operand is
  // likewise optional. An unrecognised operator drops the predicate rather than guessing.
  addCustom(attrs: XmlAttributes): void {
    if (this.#predicates === null) return;
    const operator = attrs.operator ?? 'equal';
    if (isCustomFilterOperator(operator)) {
      this.#predicates.push({operator, val: attrs.val ?? ''});
    }
  }

  // Assemble this column's criteria from whichever accumulator filled. A column whose colId is
  // negative, or whose criteria are empty (no values, no blank, no predicates), carries nothing
  // filterable and is dropped so a re-write stays clean: load-repair, not authoring.
  endColumn(): void {
    const criteria = pendingFilterCriteria(this.#values, this.#blank, this.#predicates, this.#and);
    if (this.#colId >= 0 && criteria !== null) {
      this.#columns.push({colId: this.#colId, criteria});
    }
  }

  // Commit the accumulated autofilter to the sheet. Runs on `</autoFilter>`, including the synthesized
  // close of a criteria-free self-closing `<autoFilter/>`. Columns whose colId falls outside the range
  // are dropped here so the strict setter never trips on hostile input.
  commit(sheet: Worksheet): void {
    if (this.#ref === null) return;
    const decoded = tryDecodeRange(this.#ref);
    const rect = decoded === undefined ? undefined : boundedRect(decoded);
    // A filter needs a bounded rectangle, so an unbounded or unreadable range leaves the sheet
    // without one; `canonicalizeAutoFilter` would refuse it anyway, and refusing here keeps the
    // authoring guard a guard rather than a control-flow path. All four corners, not the two
    // columns: `ref="A:C"` has both and no rows, and it reached the setter and aborted the read.
    if (rect !== undefined) {
      const width = rect.right - rect.left + 1;
      sheet.autoFilter = {ref: this.#ref, columns: this.#columns.filter((c) => c.colId < width)};
    }
    this.#ref = null;
    this.#columns = [];
  }
}
/**
 * A pass reading a worksheet part into `sheet`: cells and their styles, the column and row metadata,
 * merges, the autofilter, page breaks, and the view and print layout. It commits as it goes rather
 * than gathering, so it has nothing to hand back once the parse ends.
 *
 * Offered as a pass because the worksheet is the largest part in a package and four other readers
 * want the same events, so the five share one parse of it rather than scanning it once each.
 */
export function worksheetPass(
  sheet: Worksheet,
  sharedStrings: readonly SharedString[],
  xfStyles: ReadonlyArray<XfStyle>,
  dateEpoch: DateEpoch,
): SaxPass {
  // The one `<c>` currently being read: its address/type/style, formula, value, inline text, rich
  // runs, and the sheet-spanning shared-formula master map. Each `<c>` resets it and commits it.
  const cell = new CellAccumulator({richRuns: true, dateEpoch});
  // A row with customFormat="1" supplies a default style for its cells that carry no `s`.
  const autoFilter = new AutoFilterAccumulator();
  const pageBreaks = new PageBreakAccumulator();
  const styleResolution = new CellStyleResolver();
  const columnBudget = new ColumnRecordBudget();
  const rowPosition = new RowPositionTracker();
  // A `<headerFooter>` child's text, gathered whole: the `&`-prefixed section and format tokens
  // (`&C&"Arial"&G`) are what a header image's `&G` and every other directive ride in.
  const headerFooterText = new TextCapture(HEADER_FOOTER_ELEMENTS);
  // How deep the parse is inside `<customSheetViews>`, 0 when outside it. Each saved view repeats the
  // sheet's own pane, breaks, margins, print options, page setup, header/footer and autofilter under
  // the same local names (`CT_CustomSheetView`), and the block follows the sheet's own `<sheetViews>`
  // and `<autoFilter>`. Dispatched on name alone, a saved view replaced the sheet's pane and filter
  // and joined its breaks. The model keeps no custom views, so the whole subtree is skipped. Counted
  // on non-self-closing opens and on every close, which balances for an expanded `<x/>` too: that
  // arrives as an open that is not self-closing followed by a close.
  let customViewDepth = 0;

  // Commit the cell held in the accumulator under the style it resolves to: its own `s`, then its row's
  // (when customFormat), then its column's, then xf 0. That is the order Excel applies and the one rule
  // the streaming and BIFF12 readers share, so no two of them decode the same cell to different types.
  // Runs on `</c>` close, including the synthesized close of a self-closing `<c/>` formatted-but-empty
  // cell.
  const finalizeCellFromState = (): void => {
    cell.finalize(
      sheet,
      sharedStrings,
      styleResolution.styleFor(cell.col, cell.styleIndex, xfStyles),
    );
  };

  const handlers: SaxHandlers = {
    onOpen(name, attrs, selfClosing) {
      const local = localName(name);
      if (customViewDepth > 0) {
        if (!selfClosing) customViewDepth++;
        return;
      }
      if (local === 'customSheetViews') {
        if (!selfClosing) customViewDepth = 1;
        return;
      }
      if (cell.openElement(local, attrs, selfClosing)) return;
      if (isHeaderFooterElement(local)) {
        headerFooterText.open(local, selfClosing);
        return;
      }
      switch (local) {
        case 'col':
          applyColumn(sheet, attrs, xfStyles, styleResolution, columnBudget);
          break;
        case 'row': {
          const {number, inGrid} = rowPosition.open(attrs);
          if (inGrid) applyRow(sheet, number, attrs);
          // The cell machine is told about the row whether or not it is in the grid: it keeps
          // reading the row's cells either way (that is what keeps it in step with the element
          // stream), and it is the one that declines to place them when the row is past the grid.
          cell.openRow(number, inGrid);
          styleResolution.openRow(numInteger(attrs.s, 0) ?? -1, boolStrict(attrs.customFormat));
          break;
        }
        case 'mergeCell':
          // A well-formed file never declares overlapping merges; a corrupt one might. Reject the
          // bad range at the model boundary, but don't let one abort the whole parse: drop it and
          // keep reading the valid geometry.
          if (attrs.ref !== undefined && attrs.ref !== '') {
            const ref = attrs.ref;
            admitting(() => {
              sheet.mergeCells(ref);
            });
          }
          break;
        case 'sheetPr':
        case 'tabColor':
        case 'outlinePr':
        case 'sheetView':
        case 'pane':
        case 'pageSetUpPr':
        case 'printOptions':
        case 'pageMargins':
        case 'pageSetup':
          applySheetProperties(local, attrs, sheet);
          break;
        case 'rowBreaks':
          pageBreaks.begin(sheet.rowBreaks);
          break;
        case 'colBreaks':
          pageBreaks.begin(sheet.columnBreaks);
          break;
        case 'brk':
          pageBreaks.add(attrs);
          break;
        case 'sheetProtection': {
          const protection = parseSheetProtection(attrs);
          if (protection !== undefined) sheet[INTERNAL].restoreProtection(protection);
          break;
        }
        case 'autoFilter':
          autoFilter.begin(attrs);
          break;
        case 'filterColumn':
          autoFilter.beginColumn(attrs);
          break;
        case 'filters':
          autoFilter.beginValues(attrs);
          break;
        case 'filter':
          autoFilter.addValue(attrs);
          break;
        case 'customFilters':
          autoFilter.beginCustom(attrs);
          break;
        case 'customFilter':
          autoFilter.addCustom(attrs);
          break;
      }
    },
    onText(chunk) {
      if (customViewDepth > 0) return;
      cell.appendChunk(chunk);
      headerFooterText.text(chunk);
    },
    onClose(name) {
      if (customViewDepth > 0) {
        customViewDepth--;
        return;
      }
      const local = localName(name);
      const claimed = cell.closeElement(local);
      if (claimed === 'cell') {
        finalizeCellFromState();
        return;
      }
      if (claimed === 'claimed') return;
      const headerFooter = headerFooterText.close(local);
      if (headerFooter !== undefined && isHeaderFooterElement(local)) {
        // Header text carries the `_xHHHH_` convention, same as a cell value: Excel decodes it
        // here and re-emits it on save (measured: a patched `_x0001_` reads back over COM as
        // U+0001, and a `_x005F_x0041_` as the literal `_x0041_`). The decode is on the whole
        // element text, never on a SAX chunk. See {@link decodeSpreadsheetText}.
        sheet.headerFooter[local] = decodeSpreadsheetText(headerFooter);
        return;
      }
      switch (local) {
        case 'row':
          styleResolution.closeRow();
          cell.closeRow();
          break;
        case 'filterColumn':
          autoFilter.endColumn();
          break;
        case 'autoFilter':
          autoFilter.commit(sheet);
          break;
        case 'rowBreaks':
        case 'colBreaks':
          pageBreaks.end();
          break;
        default:
          break;
      }
    },
  };
  return {handlers, closeEmptyElements: WORKSHEET_EMPTY_CLOSES};
}
function applyColumn(
  sheet: Worksheet,
  attrs: XmlAttributes,
  xfStyles: ReadonlyArray<XfStyle>,
  styleResolution: CellStyleResolver,
  budget: ColumnRecordBudget,
): void {
  // Which columns the span reaches is `takeColumnSpan`'s answer, shared with the streaming reader so
  // the two cannot read one `<col>` as covering different columns.
  const span = takeColumnSpan(attrs, budget);
  if (span === undefined) return;
  const {first: min, last} = span;
  const width = numFinite(attrs.width);
  const hidden = boolStrict(attrs.hidden);
  const styleIndex = numInteger(attrs.style, 0) ?? -1;
  // The column's style resolves to the same facet bundle a cell's does; mirror all of it onto the
  // column model so `getColumn(i)` reflects the declared default, not just its number format.
  const style = styleIndex >= 0 ? xfStyles[styleIndex] : undefined;
  for (let index = min; index <= last; index++) {
    const column = sheet.getColumn(index);
    if (width !== undefined && boolPresent(attrs.customWidth)) column.width = width;
    if (hidden) column.hidden = true;
    const outlineLevel = numInteger(attrs.outlineLevel, 1);
    if (outlineLevel !== undefined) column.outlineLevel = outlineLevel;
    if (boolStrict(attrs.collapsed)) column.collapsed = true;
    if (style !== undefined) assignStyleFacets(column, style);
  }
  // Record the column's style so a bare cell in it can inherit the full column format on read. Noted
  // over the whole affordable span rather than per index inside the loop above, which is the same
  // work said once.
  styleResolution.noteColumnSpan(min, last, numInteger(attrs.style, 0) ?? -1);
}

function applyRow(sheet: Worksheet, number: number, attrs: XmlAttributes): void {
  // A `<row>` that states no attribute at all leaves no format record behind: the handle creates
  // one only when something is written through it. That is the right reading: a bare `<row r="5"/>`
  // carries no formatting to round-trip, and fabricating an empty record for it would put row 5 in
  // the used range on the strength of an element that says nothing.
  const row = sheet.getRow(number);
  const height = numFinite(attrs.ht);
  if (height !== undefined && boolPresent(attrs.customHeight)) row.height = height;
  if (boolStrict(attrs.hidden)) row.hidden = true;
  const outlineLevel = numInteger(attrs.outlineLevel, 1);
  if (outlineLevel !== undefined) row.outlineLevel = outlineLevel;
  if (boolStrict(attrs.collapsed)) row.collapsed = true;
}
