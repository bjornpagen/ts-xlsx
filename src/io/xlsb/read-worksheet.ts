// The BIFF12 worksheet-body reader: one `xl/worksheets/sheetN.bin` in, one populated {@link Worksheet}
// out: the binary counterpart of `../xlsx/read-worksheet.ts`, filling the very same model.
//
// The cell table is a flat, row-major run of records: a `BrtRowHdr` opens a row and every cell record
// until the next one belongs to it, so the open row is a single variable rather than nested state.
// Each cell record is a `Cell` header followed by a value shaped by the record's own type, which is
// what makes the binary form quick to parse: there is no `t=` attribute to interpret, the record
// number *is* the type.
//
// A formula cell carries both halves of what the XML form spells in `<f>` and `<v>`: a `Ptg` token
// stream, decoded back to text by `./formula.ts`, and the result Excel last computed. The two are
// filled independently: a formula whose stream uses a token this reader does not decode still
// surfaces its cached value, which is exactly what the reader could see before the decoder existed.
//
// One shape needs a second look at the sheet. An array formula's member cells hold a `PtgExp`
// pointing at the group's top-left, and the `BrtArrFmla` record carrying the group's actual formula
// comes *after* those cells in the stream. Those cells are therefore parked and resolved once the
// whole part has been read.

import {encodeAddress, encodeRect, MAX_COLUMN_INDEX, MAX_ROW_INDEX} from '../../core/address.ts';
import type {Cell} from '../../core/cell.ts';
import {coerceDateSerial, type DateEpoch} from '../../core/date.ts';
import {unmangleFunctions} from '../../core/formula.ts';
import {assignStyleFacets} from '../../core/style.ts';
import type {CellValue, ErrorValue, FormulaResult} from '../../core/value.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {ColumnRecordBudget, clampColumnSpan} from '../read-policy/column-budget.ts';
import {admitting} from '../read-policy/read-repair.ts';
import {CellStyleResolver} from '../style/cell-style-resolution.ts';
import {applyXfToCell, type XfStyle} from '../style/xf-style.ts';
import {decodeFormula, type FormulaScope, formulaAnchor} from './formula.ts';
import {errorCodeFor, RecordReader} from './primitives.ts';
import {readRecords} from './record-stream.ts';
import {BRT} from './record-types.ts';

/** Everything reading one cell's payload needs that is not the cursor itself. */
interface ValueContext {
  readonly sharedStrings: readonly string[];
  readonly numFmt: string | undefined;
  readonly epoch: DateEpoch;
}

/**
 * How one cell-bearing record's payload is read: as a value, or as a formula's cached result followed
 * by the token stream that produced it.
 *
 * One table rather than two sets and two switches. Membership and decoding were separate, so a record
 * type added to the set and forgotten in the switch fell through a `default:` that also meant
 * `BrtCellBlank` -- and read every such cell as a formatted blank, silently, for as long as nobody
 * compared the file with what Excel showed. Here `BrtCellBlank` says `() => null` in as many words,
 * the `default:` is gone, and a record with no entry is skipped whole rather than mistaken for a cell.
 */
type CellRecord =
  | {readonly kind: 'value'; readonly read: (reader: RecordReader, ctx: ValueContext) => CellValue}
  | {
      readonly kind: 'formula';
      readonly read: (reader: RecordReader, ctx: ValueContext) => FormulaResult | undefined;
    };

const value = (read: (reader: RecordReader, ctx: ValueContext) => CellValue): CellRecord => ({
  kind: 'value',
  read,
});
const formula = (
  read: (reader: RecordReader, ctx: ValueContext) => FormulaResult | undefined,
): CellRecord => ({kind: 'formula', read});

const CELL_RECORDS: ReadonlyMap<number, CellRecord> = new Map<number, CellRecord>([
  // Formatted but empty. The style is already applied; the value is genuinely none.
  [BRT.CellBlank, value(() => null)],
  // A number under a date format is a date serial, surfaced as a Date so an `.xlsb` reads the value
  // its `.xlsx` twin does. The rule is `core/date.ts`'s, shared with the XML decoders.
  [BRT.CellRk, value((r, c) => coerceDateSerial(r.rk(), c.numFmt, c.epoch))],
  [BRT.CellReal, value((r, c) => coerceDateSerial(r.f64(), c.numFmt, c.epoch))],
  [BRT.CellBool, value((r) => r.u8() !== 0)],
  // An unrecognised error byte keeps the cell non-empty without inventing an error the model does not
  // define; there is no text form to fall back to as there is in XML.
  [BRT.CellError, value((r) => errorValueOrNull(r.u8()))],
  [BRT.CellSt, value((r) => r.wideString())],
  // Rich runs are not modelled in this cut; the flattened text is what a consumer sees.
  [BRT.CellRString, value((r) => r.richString())],
  [BRT.CellIsst, value((r, c) => c.sharedStrings[r.u32()] ?? '')],
  // A formula's cached numeric result honours the cell's date format exactly as a bare number does,
  // so a date-valued formula reads back as a Date rather than a serial.
  [BRT.FmlaNum, formula((r, c) => coerceDateSerial(r.f64(), c.numFmt, c.epoch))],
  [BRT.FmlaBool, formula((r) => r.u8() !== 0)],
  [BRT.FmlaError, formula((r) => errorValueOrNull(r.u8()) ?? undefined)],
  [BRT.FmlaString, formula((r) => r.wideString())],
]);

// A BErr byte as the model's error value, or null when the byte names no error this library defines.
function errorValueOrNull(code: number): ErrorValue | null {
  const error = errorCodeFor(code);
  return error === undefined ? null : {error};
}

// A cell whose formula defers to a group's top-left, held until the record naming that group's
// formula has been read.
interface DeferredFormula {
  readonly cell: Cell;
  readonly row: number;
  readonly column: number;
  readonly anchorRow: number;
  readonly anchorColumn: number;
  readonly result: FormulaResult | undefined;
}

/**
 * Read a worksheet part into `sheet`: its column and row geometry, its merged ranges, and every
 * non-empty cell with the style its index resolves to in `xfStyles` and, for a formula cell, the text
 * its token stream decodes to through `scope`.
 */
export function parseWorksheet(part: Uint8Array, context: WorksheetReadContext): void {
  const {sheet, xfStyles, scope} = context;
  // The open row, one-based as the model counts them. -1 means none is open, which a cell record
  // arriving before any row header (a malformed sheet) is dropped against rather than guessed at.
  let row = -1;
  // Which format a cell takes: its own, its row's, its column's, then xf 0, by the rule both XML
  // readers drive too. Column records precede the cell table and a row header precedes its cells,
  // which is the order the resolver is fed in; the next row header replaces the open row's format.
  const styleResolution = new CellStyleResolver();
  // The sheet's default row height, in twips. Every row header restates its height whether or not the
  // row has one of its own, so the default is what tells the two apart. See {@link applyRow}.
  // `BrtWsFmtInfo` precedes the cell table, so it is always known by the time a row is read.
  let defaultRowHeight = -1;
  // The formula of each array-formula group, keyed by the group's top-left cell, and the member cells
  // waiting on one. Both are needed because `BrtArrFmla` follows the cells it speaks for.
  const groups = new Map<string, {rgce: Uint8Array; rgcb: Uint8Array}>();
  const deferred: DeferredFormula[] = [];
  const columnBudget = new ColumnRecordBudget();

  for (const record of readRecords(part)) {
    const cellRecord = CELL_RECORDS.get(record.type);
    const reader = new RecordReader(record.data);
    if (record.type === BRT.WsFmtInfo) {
      reader.skip(6); // dxGCol, cchDefColWidth: the default *column* width, which the model does not read.
      defaultRowHeight = reader.u16();
    } else if (record.type === BRT.ColInfo) {
      applyColumn(reader, sheet, xfStyles, styleResolution, columnBudget);
    } else if (record.type === BRT.RowHdr) {
      const header = applyRow(reader, sheet, defaultRowHeight);
      row = header.row;
      styleResolution.openRow(header.styleIndex, header.customFormat);
    } else if (record.type === BRT.MergeCell) {
      const {rowFirst, rowLast, colFirst, colLast} = reader.range();
      // A merge overlapping one already read is refused by the model; as in the XML reader, the one
      // bad range is dropped and the rest of the sheet's geometry still reads.
      if (inGrid(colFirst, rowFirst) && inGrid(colLast, rowLast)) {
        const ref = encodeRect({
          top: rowFirst + 1,
          left: colFirst + 1,
          bottom: rowLast + 1,
          right: colLast + 1,
        });
        admitting(() => {
          sheet.mergeCells(ref);
        });
      }
    } else if (record.type === BRT.ArrFmla) {
      const {rowFirst, colFirst} = reader.range();
      reader.skip(1); // fAlwaysCalc: a recalculation hint, not part of the formula.
      groups.set(groupKey(rowFirst, colFirst), {
        rgce: reader.bytes(reader.u32()),
        rgcb: reader.bytes(reader.u32()),
      });
    } else if (cellRecord !== undefined) {
      const member = readCellRecord(cellRecord, reader, {...context, row, styleResolution});
      if (member !== undefined) deferred.push(member);
    }
  }

  for (const member of deferred) {
    // Only the group's top-left cell states the formula; the rest carry the value it produced, which
    // is exactly what the XML form writes for them.
    const group = groups.get(groupKey(member.anchorRow, member.anchorColumn));
    const own = member.row === member.anchorRow && member.column === member.anchorColumn;
    member.cell.value =
      group === undefined || !own
        ? (member.result ?? null)
        : formulaValue(decodeFormula(group.rgce, group.rgcb, scope), member.result);
  }
}

/**
 * What reading a worksheet part needs around it: the sheet being filled and the three tables its
 * records resolve through, plus the date system a serial under a date format counts from.
 *
 * One value rather than five positional arguments, two of them arrays and two of those `readonly
 * string[]`/`ReadonlyArray<XfStyle>` -- a run whose order the call site cannot be read against, on a
 * reader that decides what every cell in the sheet says. It is also exactly the half of
 * {@link CellRecordContext} that does not change between records, which is why that one extends it
 * and the record loop spreads this one into it.
 */
export interface WorksheetReadContext {
  readonly sheet: Worksheet;
  readonly sharedStrings: readonly string[];
  readonly xfStyles: ReadonlyArray<XfStyle>;
  readonly scope: FormulaScope;
  /** The workbook's date system, from `BrtWbProp`: what a serial under a date format counts from. */
  readonly dateEpoch: DateEpoch;
}

/** What reading one cell record needs: the sheet around it, and the row it is currently inside. */
interface CellRecordContext extends WorksheetReadContext {
  /** The open row, one-based; -1 when none is, which is a malformed sheet. */
  readonly row: number;
  readonly styleResolution: CellStyleResolver;
}

/**
 * Read one cell or formula record onto its cell, returning the deferred entry when the formula is an
 * array-group member whose group has not been seen yet.
 *
 * Lifted out of the record loop, where it was thirty-three lines three levels deep doing style
 * resolution, value decoding, formula decoding and deferred-group bookkeeping in one block. The loop
 * is now six one-line dispatches, and this is testable on its own, which the block was not.
 */
function readCellRecord(
  record: CellRecord,
  reader: RecordReader,
  context: CellRecordContext,
): DeferredFormula | undefined {
  const {sheet, sharedStrings, xfStyles, scope, dateEpoch, row, styleResolution} = context;
  // A cell record arriving before any row header is dropped rather than guessed at.
  if (row <= 0) return undefined;
  const {column, styleIndex} = reader.cell();
  if (!inGrid(column, row - 1)) return undefined;
  // BIFF12 writes ixfe 0 where XML omits `s` and has no other way to say "none", so 0 is read as no
  // format of the cell's own and lets its row's and column's through. The record cannot tell that
  // from a cell that meant xf 0 on a formatted row, which therefore reads with the row's format.
  const style = styleResolution.styleFor(column + 1, styleIndex > 0 ? styleIndex : -1, xfStyles);
  const cell = sheet.getCell(encodeAddress(column + 1, row));
  applyXfToCell(cell, style);

  const ctx: ValueContext = {sharedStrings, numFmt: style?.numFmt, epoch: dateEpoch};
  if (record.kind === 'value') {
    cell.value = record.read(reader, ctx);
    return undefined;
  }

  const result = record.read(reader, ctx);
  reader.skip(2); // grbitFlags: per-cell recalculation hints the model does not carry.
  const rgce = reader.bytes(reader.u32());
  const rgcb = reader.bytes(reader.u32());
  const anchor = formulaAnchor(rgce, rgcb);
  if (anchor === undefined) {
    cell.value = formulaValue(decodeFormula(rgce, rgcb, scope), result);
    return undefined;
  }
  return {
    cell,
    row: row - 1,
    column,
    anchorRow: anchor.row,
    anchorColumn: anchor.column,
    result,
  };
}

function groupKey(row: number, column: number): string {
  return `${row}:${column}`;
}

// Pair a decoded formula with its cached result, in the shape the XML reader produces for the same
// cell. A formula the decoder could not read leaves the value alone: the cached result is still true,
// and is what this reader surfaced before formulas were decoded at all.
function formulaValue(formula: string | undefined, result: FormulaResult | undefined): CellValue {
  if (formula === undefined) return result ?? null;
  // Strip the `_xlfn.`/`_xlpm.` on-disk mangling, as the XML reader does, so the model never holds it.
  const stored = unmangleFunctions(formula);
  return result === undefined ? {formula: stored} : {formula: stored, result};
}

// Excel's grid bounds, zero-based as the binary format counts. [MS-XLSB] states them as MUST
// constraints, which is exactly why a reader has to check them: a damaged or hostile file states
// whatever it likes, and an address beyond the grid has nowhere to go. Everything positional funnels
// through here before it reaches the model, so an out-of-grid record is dropped rather than turned
// into an unrepresentable address (which the address encoder would reject) or, worse, a column loop
// four billion iterations long.
function inGrid(column: number, row: number): boolean {
  return column >= 0 && column <= MAX_COLUMN_INDEX && row >= 0 && row <= MAX_ROW_INDEX;
}

// `BrtRowHdr` ([MS-XLSB] 2.4.770): the row index, its default format, its height, and a byte of
// layout flags. Returns the open row (one-based), the format it names, and whether its cells inherit
// that format.
function applyRow(
  reader: RecordReader,
  sheet: Worksheet,
  defaultRowHeight: number,
): {row: number; styleIndex: number; customFormat: boolean} {
  const index = reader.u32();
  const styleIndex = reader.u32();
  const height = reader.u16();
  reader.skip(1); // fExtraAsc/fExtraDsc: border padding, a rendering hint the model does not carry.
  const flags = reader.u8();
  // A row beyond the grid closes the open row without opening another, so its cells are dropped too.
  if (index > MAX_ROW_INDEX) return {row: -1, styleIndex: -1, customFormat: false};

  const row = index + 1;
  const handle = sheet.getRow(row);
  // Every row header restates a height; only a row whose height is its *own* has one to record. That
  // is a row the user sized by hand, or one Excel auto-fitted to a taller font or wrapped text. Both
  // differ from the sheet default, which is exactly when XML emits `ht`. A row merely restating the
  // default carries no height, so it must not read back with one.
  if ((flags & ROW_CUSTOM_HEIGHT) !== 0 || height !== defaultRowHeight) {
    handle.height = height / TWIPS_PER_POINT;
  }
  if ((flags & ROW_HIDDEN) !== 0) handle.hidden = true;
  const outlineLevel = flags & ROW_OUTLINE_LEVEL;
  if (outlineLevel > 0) handle.outlineLevel = outlineLevel;
  if ((flags & ROW_COLLAPSED) !== 0) handle.collapsed = true;
  // The flag is XML's `customFormat="1"`: without it the row's format is not its cells'.
  return {row, styleIndex, customFormat: (flags & ROW_CUSTOM_FORMAT) !== 0};
}

const TWIPS_PER_POINT = 20;
const ROW_OUTLINE_LEVEL = 0b0000_0111;
const ROW_COLLAPSED = 0b0000_1000;
const ROW_HIDDEN = 0b0001_0000;
const ROW_CUSTOM_HEIGHT = 0b0010_0000;
const ROW_CUSTOM_FORMAT = 0b0100_0000;

// `BrtColInfo` ([MS-XLSB] 2.4.319): one record per run of identically-sized columns.
function applyColumn(
  reader: RecordReader,
  sheet: Worksheet,
  xfStyles: ReadonlyArray<XfStyle>,
  styleResolution: CellStyleResolver,
  budget: ColumnRecordBudget,
): void {
  const first = reader.u32();
  const last = reader.u32();
  const width = reader.u32();
  const styleIndex = reader.u32();
  const flags = reader.u16();
  const style = styleIndex > 0 ? xfStyles[styleIndex] : undefined;
  const outlineLevel = (flags >> COLUMN_OUTLINE_SHIFT) & 0b111;

  // A run can name every column to the right of the data; materialising all 16 384 of them would
  // turn a two-column sheet into a 16 384-entry model, so a run is only applied where it carries
  // something a default column does not.
  if (
    style === undefined &&
    outlineLevel === 0 &&
    (flags & (COLUMN_HIDDEN | COLUMN_CUSTOM_WIDTH | COLUMN_COLLAPSED)) === 0
  ) {
    return;
  }
  // The loop bound comes from the file, so it goes through the rule every reader shares before it is
  // one: clamped to the grid, because a run declaring four billion columns is a denial of service and
  // not a wide sheet, and charged to the sheet's budget, because a thousand full-width runs are the
  // same attack spelled a thousand times.
  const span = clampColumnSpan(first + 1, last + 1, budget);
  if (span === undefined) return;
  for (let index = span.first; index <= span.last; index++) {
    const column = sheet.getColumn(index);
    // The stored width is taken whether or not the file marks it user-set, matching the XML reader:
    // a `<col>`/`BrtColInfo` exists only for a column that differs from the sheet default in *some*
    // way, and it always states the width that column actually has.
    column.width = width / COLUMN_WIDTH_UNITS;
    if ((flags & COLUMN_HIDDEN) !== 0) column.hidden = true;
    if (outlineLevel > 0) column.outlineLevel = outlineLevel;
    if ((flags & COLUMN_COLLAPSED) !== 0) column.collapsed = true;
    if (style !== undefined) assignStyleFacets(column, style);
  }
  // ixfe 0 is the default xf here as on a cell, so only a nonzero one is a format the column's cells take.
  if (styleIndex > 0) styleResolution.noteColumnSpan(span.first, span.last, styleIndex);
}

// Column width is stored in 1/256ths of a character, where XML states the character count directly.
const COLUMN_WIDTH_UNITS = 256;
const COLUMN_HIDDEN = 0b0000_0001;
const COLUMN_CUSTOM_WIDTH = 0b0000_0010;
const COLUMN_COLLAPSED = 0b1000_0000_0000;
const COLUMN_OUTLINE_SHIFT = 8;
