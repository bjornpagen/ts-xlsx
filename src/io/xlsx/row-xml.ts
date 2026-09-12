// Rendering one row and its cells: the `<row>` element, the `<c>` elements inside it, and everything
// a cell's own value and format contribute to them.
//
// Split from `worksheet-xml.ts`, which stays the sheet orchestrator: the two halves shared nothing but
// the style registry, and this is precisely the surface the streaming writer imports. That writer
// serialises a row the moment it is committed, long before the sheet around it exists, so "what does
// one row look like" is a question with an answer independent of the sheet, and a module boundary is
// the honest way to say so.

import {encodeRange, tryDecodeAnchoredRange} from '../../core/address.ts';
import {type Cell, cellHasOwnStyle} from '../../core/cell.ts';
import {type DateEpoch, DEFAULT_DATE_NUMFMT, dateToSerial} from '../../core/date.ts';
import {mangleFormula} from '../../core/formula.ts';
import {NAMED_STYLE_ID} from '../../core/internal.ts';
import {CELL_STYLE_FACETS, type CellStyle, type Fill} from '../../core/style.ts';
import {
  type FormulaResult,
  detectValueType,
  isArrayFormulaValue,
  isDataTableFormulaValue,
  isErrorCode,
  isErrorValue,
  isFormulaValue,
  isRichTextValue,
  isSharedFormulaValue,
} from '../../core/value.ts';
import type {ColumnProperties, RowProperties, Worksheet} from '../../core/worksheet.ts';
import {AuthoringError, InternalError, quoted} from '../../errors.ts';
import {
  boolAttr,
  checkedToken,
  escapeAttr,
  escapeSpreadsheetText,
  escapeText,
  numberText,
  nonDefaultNumAttr,
  textAttr,
  textElement,
} from '../../xml/xml.ts';
import type {XfStyle} from '../style/xf-style.ts';
import type {CellMetadataTable} from './cell-metadata.ts';
import type {CommentCell} from './comments.ts';
import {richTextRunsXml} from './rich-text.ts';
import type {SharedFormulaRole} from './shared-formulas.ts';
import type {SharedStringTable} from './shared-strings.ts';
import type {StyleRegistry} from './styles.ts';

/**
 * The used-cell extent of a sheet: the top-left/bottom-right grid bounds that fold into the
 * `<dimension>`. Rows carrying only formatting (a row height, an outline level) do not extend the
 * used range, matching how Excel records `<dimension>`, so {@link add} ignores them. A fresh extent
 * holds the `Infinity`/`-Infinity` sentinels; {@link isEmpty} reports that no used cell has been seen.
 */
export class Extent {
  top = Infinity;
  left = Infinity;
  bottom = -Infinity;
  right = -Infinity;

  // Seed from a prior extent (the rows a streaming writer already flushed and evicted) so the buffered
  // pass folds its live rows onto the same bounds; unseeded, it starts empty.
  constructor(seed?: Extent) {
    if (seed) {
      this.top = seed.top;
      this.left = seed.left;
      this.bottom = seed.bottom;
      this.right = seed.right;
    }
  }

  /** Whether no used cell has been folded in yet, in which case the dimension is the lone cell `A1`. */
  get isEmpty(): boolean {
    return this.bottom === -Infinity;
  }

  /** Fold a rendered row's used-column span into the extent. `minCol` is `Infinity` when the row
   * carried no cells (only formatting), which extends nothing. */
  add(row: number, minCol: number, maxCol: number): void {
    if (minCol === Infinity) return;
    if (row < this.top) this.top = row;
    if (row > this.bottom) this.bottom = row;
    if (minCol < this.left) this.left = minCol;
    if (maxCol > this.right) this.right = maxCol;
  }
}

/**
 * A row the streaming writer already serialised, kept with the attribute string it was built from.
 *
 * The attributes are carried rather than re-read out of `xml`, and that is the whole point of the
 * field. `collapsed="1"` is decided after the row is gone (see {@link FlushedSheet.rowOutline}), so
 * the summary row has to be patched, and the patcher used to ask `xml.includes(' collapsed="1"')`
 * whether the attribute was already there. The `<row>` tag's shape is known, but the *cell text*
 * embedded in it is not: `escapeText` leaves a double quote verbatim, so a cell whose value was the
 * literal string ` collapsed="1"` answered that question for the row and left the outline group
 * rendering expanded. A file cannot forge a field.
 */
export interface FlushedRow {
  readonly number: number;
  readonly xml: string;
  /** The `<row>` element's own attributes, exactly as {@link renderRow} emitted them. */
  readonly attrs: string;
}

/**
 * A worksheet's eagerly-serialised rows: each row's `<row>` XML tagged with its number (so it merges
 * into ascending order with the sheet's remaining live rows, whatever order it was committed in), plus
 * the used-cell {@link Extent} they span. The buffered pass folds that extent into the sheet's dimension.
 */
export interface FlushedSheet {
  readonly rows: readonly FlushedRow[];
  readonly extent: Extent;
  /**
   * Each flushed row's outline level and hidden flag. Carried across the eviction because both feed
   * whole-sheet derivations made long after the row is gone from the model: `<sheetFormatPr
   * outlineLevelRow>` is the deepest level on the sheet, and `collapsed="1"` rides a summary row only
   * when its whole detail group is hidden, which is a question about *other* rows.
   *
   * Carried as the two inputs rather than the derived flag, because the derivation is a look-ahead a
   * row being finalised cannot do, while its inputs are per-row and known exactly when the row flushes.
   */
  readonly rowOutline: ReadonlyMap<
    number,
    {readonly outlineLevel: number; readonly hidden: boolean}
  >;
  /**
   * The notes the flushed rows carried, on the same terms and for the same reason. A note is
   * serialised outside the `<row>`, into the comments and VML parts, and the buffered pass gathers
   * notes by walking the sheet's rows at commit time, which finds nothing on a row whose cells have
   * already been evicted. A hyperlink needs no such record: it lives beside the grid, on the sheet.
   */
  readonly notes: readonly CommentCell[];
}

/**
 * A column's style facets are defaults its cells inherit unless they override them; the writer
 * composes each cell's full style up front (cell over row over column, per facet) so a cell that
 * overrides one facet still carries the column's others, rather than silently dropping them. Frozen
 * once by the streaming writer at its first flush so every eagerly-rendered row sees the same defaults.
 */
export function buildColumnDefaults(sheet: Worksheet): Map<number, ColumnProperties> {
  const columnDefaults = new Map<number, ColumnProperties>();
  // `columns()` yields only columns that carry a format record, so the fallback is unreachable.
  // it is here because the handle's `properties` is honestly optional, not because a defined
  // column can lack one.
  for (const {index, properties} of sheet.columns()) columnDefaults.set(index, properties ?? {});
  return columnDefaults;
}

/** The whole-sheet context a single row needs to serialise: the column defaults it inherits, the
 * style/string tables it interns into, the shared-formula roles its cells play, and the collapsed
 * outline summaries whose toggle it must stamp. The streaming writer supplies empty shared-formula
 * and collapsed-summary sets, since those are whole-sheet derivations a flushed row cannot join. */
export interface RowRenderContext {
  readonly columnDefaults: ReadonlyMap<number, ColumnProperties>;
  readonly styles: StyleRegistry;
  readonly sharedStrings: SharedStringTable | null;
  readonly sharedRoles: ReadonlyMap<string, SharedFormulaRole>;
  readonly collapsedSummaries: ReadonlySet<number>;
  /** The workbook's date system, which is what a `Date` cell's serial counts from. */
  readonly dateEpoch: DateEpoch;
  /** The defined names a bare name in this sheet's formulas resolves to, as `formulaNamesInScope`
   * gives them: a function passed as a value is written bare where one of them captures its name. */
  readonly formulaNames: ReadonlySet<string>;
  /** The workbook's cell metadata, which a dynamic-array formula's cell points into. Shared by every
   * row of every sheet, since the part it becomes is the workbook's. */
  readonly cellMetadata: CellMetadataTable;
}

/**
 * Serialise one row to its `<row>` element, or '' when the row has neither data nor its own
 * formatting. Returns the used-column bounds (`Infinity`/`-Infinity` when nothing was rendered) so a
 * caller can fold them into the sheet dimension, and the attribute string separately, so a later pass
 * can ask what this row declared without reading it back out of markup that also holds cell text.
 * Shared by the buffered sheet pass and the streaming writer's eager flush, so both emit
 * byte-identical rows.
 */
export function renderRow(
  entry: {
    readonly number: number;
    readonly cells: readonly Cell[];
    readonly properties: RowProperties | undefined;
  },
  ctx: RowRenderContext,
): {xml: string; attrs: string; minCol: number; maxCol: number} {
  const {number, cells, properties} = entry;
  // A cell earns a <c> element if it holds a value OR carries its own style: a formatted-but-empty
  // cell (a fill/border on a null value) is a real cell to Excel, and dropping it would lose the
  // formatting. A cell with neither is inherited from its row/column and needs no element of its own.
  const rendered = cells.filter((cell) => cell.value !== null || cellHasOwnStyle(cell));
  const attrs = rowAttrs(properties, ctx.styles, ctx.collapsedSummaries.has(number));
  // A row with neither data nor its own formatting has nothing to serialise.
  if (rendered.length === 0 && attrs === '') {
    return {xml: '', attrs: '', minCol: Infinity, maxCol: -Infinity};
  }
  const rowFill = properties?.fill;
  const cellsXml = rendered
    .map((cell) => {
      const style = ctx.styles.styleId(
        composeCellStyle(cell, rowFill, ctx.columnDefaults.get(cell.col)),
      );
      return cellXml(cell, style, ctx.sharedRoles.get(cell.address), ctx);
    })
    .join('');
  let minCol = Infinity;
  let maxCol = -Infinity;
  for (const cell of rendered) {
    if (cell.col < minCol) minCol = cell.col;
    if (cell.col > maxCol) maxCol = cell.col;
  }
  return {xml: `${rowOpenTag(number)}${attrs}>${cellsXml}</row>`, attrs, minCol, maxCol};
}

/**
 * The `<row>` element's opening tag up to its first attribute: the one thing a later pass may rely on
 * about a rendered row's shape, spelled here so it is the emitter's own string rather than a
 * re-parse of the emitter's output.
 */
export function rowOpenTag(number: number): string {
  return `<row r="${number}"`;
}

// Compose a cell's full style by resolving each facet cell-over-row-over-column, so a cell that
// overrides one facet still carries the row's fill and the column's other facets rather than silently
// dropping them: the per-facet precedence Excel applies.
//
// Cell-over-column is the default, taken from the facet registry, so a seventh facet added to
// `CellStyle` composes correctly on the day it joins. The two facets that resolve against more than
// that overwrite themselves afterwards rather than being exempted from the loop: the loop stays the
// exhaustive half, and a special case has to be written down to exist. Quote-prefix and the
// named-style link are cell-only, with no row/column default to inherit, and are deliberately not
// facets.
function composeCellStyle(
  cell: Cell,
  rowFill: Fill | undefined,
  colDef: ColumnProperties | undefined,
): XfStyle {
  const style: {-readonly [K in keyof XfStyle]?: XfStyle[K]} = {};
  for (const facet of CELL_STYLE_FACETS) inheritFacet(style, facet, cell, colDef);
  // A row carries a fill and nothing else, so it sits between the cell and the column on this one
  // facet alone.
  style.fill = cell.fill ?? rowFill ?? colDef?.fill;
  // A bare Date carries no format of its own, so it renders as a raw serial and reads back as a
  // number unless we apply a date format. An explicit cell/column format wins.
  style.numFmt = cell.numFmt ?? colDef?.numFmt ?? dateDefaultNumFmt(cell.value);
  if (cell.quotePrefix !== undefined) style.quotePrefix = cell.quotePrefix;
  // Preserved so a round-trip keeps the cell tied to its named style rather than flattening it into
  // a purely-direct format.
  const xfId = cell[NAMED_STYLE_ID];
  if (xfId !== undefined) style.xfId = xfId;
  return style;
}

// One facet at a time, so `target[key] = a[key] ?? b[key]` typechecks without a cast: the
// correlated-key access TS cannot verify when the key is the whole union. The same shape, and the
// same reason, as `copyFacet` in `core/style.ts`.
function inheritFacet<K extends keyof CellStyle>(
  target: CellStyle,
  key: K,
  cell: Readonly<CellStyle>,
  colDef: Readonly<CellStyle> | undefined,
): void {
  target[key] = cell[key] ?? colDef?.[key];
}

function rowAttrs(
  properties: RowProperties | undefined,
  styles: StyleRegistry,
  collapsedSummary: boolean,
): string {
  if (properties === undefined) return collapsedSummary ? ' collapsed="1"' : '';
  let attrs = '';
  if (properties.height !== undefined)
    attrs += ` ht="${numberText(properties.height)}" customHeight="1"`;
  if (properties.hidden) attrs += ' hidden="1"';
  attrs += outlineAttr('outlineLevel', properties.outlineLevel);
  // The collapse toggle is set explicitly by the author, or derived onto a summary row whose whole
  // detail group is hidden (the `collapsedSummaries` set `scanRowOutline` in `worksheet-xml.ts`
  // builds). It rides the summary row, never the detail rows.
  if (properties.collapsed || collapsedSummary) attrs += ' collapsed="1"';
  // A row-level fill is a default format for the row's cells; customFormat="1" is what makes
  // Excel honour the row's `s`, and a cell without its own `s` then inherits it.
  const style = styles.styleId({fill: properties.fill});
  if (style !== 0) attrs += ` s="${style}" customFormat="1"`;
  return attrs;
}

// A valid Date, whether the cell's own value or a formula's cached result, with no format of its
// own gets the default date format so it renders and reads back as a date rather than a bare serial.
// An Invalid Date and every non-date value contribute nothing here.
function dateDefaultNumFmt(value: Cell['value']): string | undefined {
  const date =
    value instanceof Date
      ? value
      : (isFormulaValue(value) || isSharedFormulaValue(value) || isArrayFormulaValue(value)) &&
          value.result instanceof Date
        ? value.result
        : undefined;
  return date !== undefined && !Number.isNaN(date.getTime()) ? DEFAULT_DATE_NUMFMT : undefined;
}

// The `<c>` element, assembled in one place. It was assembled in nineteen: every arm of the value
// dispatch and every arm of the cached-result dispatch rebuilt it from scratch, so the attribute
// order, the self-closing decision and the `<v>` wrapper were each re-derived per arm and each free
// to drift per arm.
//
// An empty `body` is the self-closing form. A formatted-but-empty cell, a value with no OOXML
// spelling, and a formula whose result was not cached all arrive here that way: they differ in why
// there is nothing to say, not in what Excel reads back. `cm` is the cell-metadata attribute, after
// `t` where Excel writes it.
function cellElement(ref: string, s: string, type: string, body: string, cm = ''): string {
  const t = type === '' ? '' : ` t="${type}"`;
  return body === '' ? `<c r="${ref}"${s}${t}${cm}/>` : `<c r="${ref}"${s}${t}${cm}>${body}</c>`;
}

// A cell's type token and its `<v>` text. `v` is null when there is no `<v>` at all, which is not the
// same as an empty one: a formula whose cached result is the empty string caches `<v></v>`, and
// collapsing that to a self-closing cell would lose the fact that it was calculated.
interface CellBody {
  readonly type: string;
  readonly v: string | null;
}

// A value the format has no way to spell, kept as a styled but empty cell rather than emitted as a
// bare `NaN`/`Infinity` token, so one bad value never corrupts the sheet or takes the export down.
const UNWRITABLE: CellBody = {type: '', v: null};

const vElement = (v: string | null): string => (v === null ? '' : `<v>${v}</v>`);

/**
 * The four value kinds a bare cell and a cached formula result spell identically.
 *
 * That they do was previously a claim in a comment ("typing the cell by the result's kind exactly as a
 * bare value of that kind would be") restated by hand in seven arms on one side and five on the other.
 * Sharing the function makes it structural: the two dispatches cannot disagree about how a boolean is
 * typed, because there is only one answer to give.
 *
 * `undefined` is "not one of the four": a string, rich text, a hyperlink label. Those genuinely differ
 * between the callers -- a bare string may be pooled into the shared table, a cached one is always
 * `t="str"` -- so each caller spells its own.
 */
function valueBody(value: Cell['value'] | FormulaResult, epoch: DateEpoch): CellBody | undefined {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? {type: '', v: numberText(value)} : UNWRITABLE;
  }
  if (typeof value === 'boolean') return {type: 'b', v: value ? '1' : '0'};
  if (value instanceof Date) {
    // An Invalid Date (new Date(NaN)) has no serial. A valid one caches the serial the cell's date
    // number format (applied when its style is composed) reads back as a Date.
    return Number.isNaN(value.getTime())
      ? UNWRITABLE
      : {type: '', v: numberText(dateToSerial(value, epoch))};
  }
  // The code goes into `<v>` unescaped, which is sound only for the canonical spellings in ERROR_CODES.
  // `isErrorValue` asks for an `error` key and nothing more, so the set is checked here, as every closed
  // token is at the write boundary: a code smuggled past the type closed the cell and wrote markup of
  // its own into the sheet.
  if (isErrorValue(value)) {
    return {type: 'e', v: checkedToken(value.error, isErrorCode, 'cell error value')};
  }
  return undefined;
}

function cellXml(
  cell: Cell,
  style: number,
  shared: SharedFormulaRole | undefined,
  ctx: RowRenderContext,
): string {
  const {sharedStrings, dateEpoch: epoch} = ctx;
  const ref = cell.address;
  const value = cell.value;
  const s = style !== 0 ? ` s="${style}"` : '';

  const formula = cellFormulaXml(cell, s, shared, ctx);
  if (formula !== undefined) return formula;

  const body = valueBody(value, epoch);
  if (body !== undefined) return cellElement(ref, s, body.type, vElement(body.v));

  // A string and rich text are one arm, not two: with shared strings on, both are pooled as an `<si>`
  // and the cell holds only the pool index (`t="s"`); with them off, both live inline in the cell.
  // Either way the read decodes back to what was written.
  if (typeof value === 'string' || isRichTextValue(value)) {
    if (sharedStrings !== null) {
      return cellElement(ref, s, 's', vElement(`${sharedStrings.intern(value)}`));
    }
    const inline = typeof value === 'string' ? textElement(value) : richTextRunsXml(value.richText);
    return cellElement(ref, s, 'inlineStr', `<is>${inline}</is>`);
  }
  // A null value only reaches here for a formatted-but-empty cell (the row loop keeps it for its
  // style); emit the styled cell with no <v>, exactly how Excel stores a formatted blank.
  if (value === null) return cellElement(ref, s, '', '');
  // Every ValueType kind is served by an arm above (a formula routes through its own writer), so
  // this is unreachable. It exists because the union is not exhaustively narrowed here.
  throw new InternalError(
    `writing a ${detectValueType(value)} cell value has no arm: every CellValue kind is handled above`,
  );
}

// Serialise a formula cell (a shared-formula master or clone, a What-If data table, an array formula,
// or a plain formula) into its `<c>` element, or return undefined when the value is not a formula so
// `cellXml` falls through to its value dispatch.
function cellFormulaXml(
  cell: Cell,
  s: string,
  shared: SharedFormulaRole | undefined,
  ctx: RowRenderContext,
): string | undefined {
  const {dateEpoch: epoch, formulaNames} = ctx;
  const {address: ref, value} = cell;
  // A shared-formula master seeds the group with its formula text under `t="shared" ref si`; a clone
  // carries no text of its own, only a back-reference to the master's `si`. Its cached result still
  // travels with the cell.
  if (shared !== undefined) {
    if (shared.ref !== undefined && isFormulaValue(value)) {
      const f = `<f t="shared" ref="${shared.ref}" si="${shared.si}">${escapeText(mangleFormula(value.formula, formulaNames))}</f>`;
      return formulaBodyXml(ref, s, f, value.result, epoch);
    }
    const result = isSharedFormulaValue(value) ? value.result : undefined;
    return formulaBodyXml(ref, s, `<f t="shared" si="${shared.si}"/>`, result, epoch);
  }
  if (isDataTableFormulaValue(value)) {
    // A data-table formula carries no expression text, only its declaration attributes, which we
    // re-emit verbatim so a read-modify-write cycle preserves the What-If kind the library never
    // evaluates. The cached result travels as any formula result does.
    const attrs =
      `ref="${escapeAttr(value.ref)}"` +
      boolAttr('dt2D', value.dataTable2D) +
      boolAttr('dtr', value.dataTableRow) +
      (value.r1Deleted === true ? ' del1="1"' : '') +
      (value.r2Deleted === true ? ' del2="1"' : '') +
      textAttr('r1', value.r1) +
      textAttr('r2', value.r2);
    return formulaBodyXml(ref, s, `<f t="dataTable" ${attrs}/>`, value.result, epoch);
  }
  if (isArrayFormulaValue(value)) {
    // A dynamic array differs from a legacy array formula only by the cell metadata its `cm` points
    // at; without that link Excel opens it as a Ctrl+Shift+Enter formula that no longer spills.
    const cm = value.dynamic === true ? ` cm="${ctx.cellMetadata.markDynamicArray()}"` : '';
    const f = `<f t="array" ref="${arrayFormulaRange(cell, value.ref)}">${escapeText(mangleFormula(value.formula, formulaNames))}</f>`;
    return formulaBodyXml(ref, s, f, value.result, epoch, cm);
  }
  if (isFormulaValue(value)) {
    return formulaBodyXml(
      ref,
      s,
      `<f>${escapeText(mangleFormula(value.formula, formulaNames))}</f>`,
      value.result,
      epoch,
    );
  }
  return undefined;
}

// The range an array formula fills, as Excel spells it. The range starts at the cell holding the
// formula, the only cell Excel states it on, so a `ref` starting anywhere else describes no formula
// this cell can hold and is refused rather than written.
function arrayFormulaRange(cell: Cell, ref: string): string {
  const range = tryDecodeAnchoredRange(ref, cell.col, cell.row);
  if (range === undefined) {
    throw new AuthoringError(
      `the array formula in ${cell.address} fills ${quoted(ref)}, which is not a range starting at ${cell.address}`,
    );
  }
  return encodeRange(range);
}

// Wrap a prepared `<f>` element (a plain formula, an array formula, or a shared master/slave `<f>`)
// with the cell element and its cached result.
function formulaBodyXml(
  ref: string,
  s: string,
  f: string,
  result: FormulaResult | undefined,
  epoch: DateEpoch,
  cm = '',
): string {
  // An uncalculated formula caches nothing, and the cell is the formula alone.
  if (result === undefined) return cellElement(ref, s, '', f, cm);
  const body = valueBody(result, epoch);
  if (body !== undefined) return cellElement(ref, s, body.type, f + vElement(body.v), cm);
  if (typeof result === 'string') {
    // The cached result of a string formula is a cell value, not structure, so it carries the
    // `_xHHHH_` escape a `<t>` does, and Excel decodes it here too (verified over COM: a `<v>` of
    // `_x0041_` under t="str" reads back as "A" with calculation held manual).
    return cellElement(ref, s, 'str', f + vElement(escapeSpreadsheetText(result)), cm);
  }
  // Every FormulaResult kind is handled above; this guards a value that reached here past the model.
  throw new InternalError(
    'writing a non-primitive formula result has no arm: every FormulaResult kind is handled above',
  );
}

/**
 * Refuse an outline depth that is not a non-negative integer, naming what carried it.
 *
 * `xsd:unsignedInt` is the attribute's type, so this is what the schema already says. It is asserted
 * rather than clamped because a negative level is also a hang: the writer finds a group's end by
 * walking outwards while the neighbouring level exceeds the summary's, and every unmapped row
 * answers `0`, which exceeds `-1` for as long as there are rows to walk.
 */
export function assertWritableLevel(name: string, level: number): void {
  if (Number.isInteger(level) && level >= 0) return;
  throw new AuthoringError(`${name} must be a non-negative integer, not ${level}`);
}

// An outline depth is written only above the default of zero. That zero test answers whether the
// level is worth recording, not whether it can be recorded at all, and keeping the two apart is what
// makes an unwritable level loud: `NaN > 0` is false, so a gate on its own would drop it silently
// while `Infinity` sailed through into an `xsd:unsignedInt` attribute.
export function outlineAttr(name: string, level: number | undefined): string {
  if (level === undefined) return '';
  assertWritableLevel(name, level);
  return nonDefaultNumAttr(name, level, 0);
}
