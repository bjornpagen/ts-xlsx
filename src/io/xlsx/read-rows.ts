// Streaming row reader: yield a worksheet's rows one at a time, without ever building the whole
// {@link Workbook} model.
//
// `readXlsx` materialises every cell of every sheet as a live `Cell` object held in nested Maps,
// fine for editing, but for a large sheet read purely to extract its data it holds the entire grid
// in memory at once. This reader instead *pulls* the sheet's XML through `xmlEvents` and yields a
// plain {@link StreamedRow} at each `</row>`, retaining only the row currently in hand. Peak model
// memory is one row, not the sheet.
//
// Two entry points sit on the same scanner:
//   - {@link readSheetRows} streams a single selected sheet's rows (the terse data-extraction case).
//   - {@link readWorkbookStream} yields a {@link StreamedSheet} per worksheet in workbook order, so a
//     caller can walk every sheet, and each sheet's rows still stream one at a time.
//
// Scope of this slice: the package is still inflated whole (bounded by the running counter in
// `./inflate.ts`) and shared strings / styles are read as whole parts, both being legitimately
// document-sized and cheap. What this avoids is retaining N materialised cells. A later slice can
// make the inflate itself per-part lazy; the pull primitive this stands on (`xmlEvents`) is the
// same one that path will use.

import type {DateEpoch} from '../../core/date.ts';
import type {CellValue} from '../../core/value.ts';
import {Workbook} from '../../core/workbook.ts';
import {WorksheetMerges} from '../../core/worksheet-merges.ts';
import {AuthoringError, quoted} from '../../errors.ts';
import {numInteger} from '../../xml/xml-attrs.ts';
import {closeEmptyElements, parseXmlPasses} from '../../xml/xml-read.ts';
import {boolStrict, localName, type XmlAttributes, xmlEvents} from '../../xml/xml-scan.ts';
import type {CellMetadataIndex} from '../cell-metadata/metadata.ts';
import {openSpreadsheetPackage, readPartRelationships} from '../opc/read-opc.ts';
import type {ReadPackageOptions} from '../opc/read-options.ts';
import {unsupportedWorkbookPart} from '../opc/sniff-format.ts';
import {ColumnRecordBudget} from '../read-policy/column-budget.ts';
import {admitting, repairedSheetNames} from '../read-policy/read-repair.ts';
import {CellStyleResolver} from '../style/cell-style-resolution.ts';
import type {XfStyle} from '../style/xf-style.ts';
import {CellAccumulator, WORKSHEET_BODY_EMPTY_CLOSES} from './cell-accumulator.ts';
import {readCellMetadata} from './cell-metadata.ts';
import type {SharedString} from './cell-value.ts';
import {takeColumnSpan} from './column-span.ts';
import {XlsxParseError} from './errors.ts';
import {SHARED_STRINGS_PART, STYLES_PART} from './part-names.ts';
import {parseSharedStrings} from './read-shared-strings.ts';
import {parseStyleTable} from './read-styles.ts';
import {
  definedNamesPass,
  type SheetEntry,
  workbookPropertiesPass,
  workbookSheetsPass,
} from './read-workbook-xml.ts';
import {RowPositionTracker} from './row-position.ts';

export interface ReadSheetRowsOptions extends ReadPackageOptions {
  /**
   * Which worksheet to stream: its name, or its 1-based position in the workbook. Defaults to the
   * first sheet.
   */
  readonly sheet?: string | number;
}

/**
 * The resolved style facets of a streamed cell: its own `<c s>` format, failing that its row's (when
 * the row is marked `customFormat`), failing that its column's, flattened exactly as the buffered
 * reader resolves it. Present only when the cell, its row or its column declares a format; a consumer
 * can copy these straight onto a writer cell to preserve its look through a streaming read→write.
 *
 * @unpublished A named commitment this surface is not ready to make. `entries/xlsx.ts` records the
 * decision: the streaming reader's granular output shapes stay inferred structural types while the
 * surface settles, so a consumer destructures them and never has to spell them.
 */
export type StreamedCellStyle = XfStyle;

/** One non-empty cell in a {@link StreamedRow}.
 *
 * @unpublished A named commitment this surface is not ready to make. `entries/xlsx.ts` records the
 * decision: the streaming reader's granular output shapes stay inferred structural types while the
 * surface settles, so a consumer destructures them and never has to spell them.
 */
export interface StreamedCell {
  /** 1-based column index. */
  readonly col: number;
  /** Canonical A1 address (`"B3"`). */
  readonly address: string;
  /** The decoded value, identical to what `readXlsx` would produce for the same cell. */
  readonly value: CellValue;
  /** The cell's resolved style facets, including a format it inherits from its row or column; absent
   * when none of the three declares one. */
  readonly style?: StreamedCellStyle;
}

/** One worksheet row, as yielded by {@link readSheetRows} / {@link StreamedSheet.rows}. */
export interface StreamedRow {
  /** 1-based row index. */
  readonly number: number;
  /** Whether the row declares itself hidden. */
  readonly hidden: boolean;
  /** The row's non-empty cells, in column order. An empty (or purely style-only) row yields none. */
  readonly cells: readonly StreamedCell[];
}

/**
 * One worksheet, as yielded by {@link readWorkbookStream}. The sheet's {@link rows} stream one at a
 * time; its {@link hiddenColumns} and {@link merges} are populated by that same single pass.
 *
 * The two summaries are resolved lazily: reading either accessor drives a full scan of the sheet if
 * its rows have not already been consumed, so their order relative to `rows()` never matters. (When
 * rows *are* consumed first, the streaming idiom, the accessors reuse that pass and re-scan
 * nothing.)
 *
 * @unpublished A named commitment this surface is not ready to make. `entries/xlsx.ts` records the
 * decision: the streaming reader's granular output shapes stay inferred structural types while the
 * surface settles, so a consumer destructures them and never has to spell them.
 */
export interface StreamedSheet {
  /**
   * The worksheet's name, joined from the workbook part and repaired exactly as `readXlsx` repairs it
   * (a duplicate becomes `S (2)`, a forbidden character is removed), so the two readers name a sheet
   * the same way. Never a positional placeholder.
   */
  readonly name: string;
  /** Stream this sheet's rows, one at a time, in sheet order. */
  // The two extra arguments are not decoration: a bare `Generator<T>` defaults its return and next
  // types to `any`, and that `any` reaches the caller the moment they touch `.next()` rather than
  // `for…of`. `void, undefined` says what these generators actually do, ending with nothing and
  // taking nothing back, and keeps the streaming API free of `any`.
  rows(): Generator<StreamedRow, void, undefined>;
  /** 1-based indices of columns the sheet declares hidden, ascending. */
  readonly hiddenColumns: readonly number[];
  /**
   * The sheet's merged ranges, as canonical A1 range strings, in declaration order: the ones
   * `readXlsx` admits, so an unreadable range, or one overlapping a range declared before it, is
   * dropped.
   */
  readonly merges: readonly string[];
}

/**
 * Stream a worksheet's rows from an `.xlsx` package, yielding each in sheet order without building
 * the workbook model. Only rows the sheet actually declares are yielded, and within a row only its
 * non-empty cells: a blank or style-only cell contributes nothing, matching the intent of a data
 * read.
 *
 * @param data The raw `.xlsx` bytes.
 * @param options Sheet selector and the inflate bound (see {@link ReadSheetRowsOptions}).
 * @throws {UnsupportedFormatError} if the input is not a readable `.xlsx` package (a legacy `.xls`, a
 *   binary `.xlsb`, or an unrecognised/non-ZIP blob; branch on `.format`).
 * @throws {PackageReadError} if the input is a ZIP that cannot be unpacked: a corrupt or
 *   truncated archive, or one exceeding the inflate bound (a probable zip bomb).
 * @throws {XlsxParseError} if the package's workbook part declares no worksheets.
 * @throws {RangeError} / {@link AuthoringError} if `options.sheet` selects a position, or a name,
 *   that no worksheet has.
 */
export function* readSheetRows(
  data: Uint8Array,
  options: ReadSheetRowsOptions = {},
): Generator<StreamedRow, void, undefined> {
  const pkg = openPackage(data, options.maxUncompressedBytes);
  const chosen = pickSheet(pkg.sheets, options.sheet);
  const sheetXml = pkg.sheetXml(chosen.relId);
  // The sheet is named but its part is missing (a truncated or foreign package), so it has no rows.
  if (sheetXml === undefined) return;
  yield* scanSheet(sheetXml, pkg, new Set(), new WorksheetMerges());
}

/**
 * Stream every worksheet of an `.xlsx` package in workbook order, without building the workbook
 * model. Each yielded {@link StreamedSheet} carries the declared sheet name and lets the caller
 * stream that sheet's rows and read its hidden-column and merge summaries: the streaming analogue
 * of walking `readXlsx(data).worksheets`.
 *
 * @param data The raw `.xlsx` bytes.
 * @param options The inflate bound (see {@link ReadPackageOptions}).
 * @throws {UnsupportedFormatError} if the input is not a readable `.xlsx` package (a legacy `.xls`, a
 *   binary `.xlsb`, or an unrecognised/non-ZIP blob; branch on `.format`).
 * @throws {PackageReadError} if the input is a ZIP that cannot be unpacked: a corrupt or
 *   truncated archive, or one exceeding the inflate bound (a probable zip bomb).
 */
export function* readWorkbookStream(
  data: Uint8Array,
  options: ReadPackageOptions = {},
): Generator<StreamedSheet, void, undefined> {
  const pkg = openPackage(data, options.maxUncompressedBytes);
  for (const sheet of pkg.sheets) {
    // A named sheet whose part is missing (truncated/foreign package) still surfaces, with no rows,
    // no hidden columns, and no merges, rather than vanishing from the workbook's sheet list.
    const xml = pkg.sheetXml(sheet.relId) ?? '';
    yield new StreamedSheetReader(sheet.name, xml, pkg);
  }
}

// The shared parts every streaming read needs: the sheet directory (name + rel id, in workbook
// order), the shared-string and style tables, and a resolver from a sheet's rel id to its XML. The
// package is inflated once; sheet XML is fetched lazily so a sheet the caller never visits is never
// stringified.
interface OpenPackage extends SheetTables {
  readonly sheets: ReadonlyArray<{name: string; relId: string}>;
  sheetXml(relId: string): string | undefined;
}

// What every sheet's cells decode against, the same for each sheet of the package.
interface SheetTables {
  readonly sharedStrings: readonly SharedString[];
  readonly xfStyles: ReadonlyArray<XfStyle>;
  /** The workbook's declared date system, read from `<workbookPr date1904>` like the buffered
   * reader's: a streamed cell whose serial counted from another day is a different date. */
  readonly dateEpoch: DateEpoch;
  /** Every name the workbook defines, as `definedNameKeys` spells them: whether a function a formula
   * passes as a value sheds its `_xleta.` depends on them, as it does in the buffered reader. */
  readonly definedNames: ReadonlySet<string>;
  /** What the workbook's `cm` and `vm` values resolve to, as the buffered reader resolves them: a
   * streamed array formula or rich-value error is the same value it is there. */
  readonly cellMetadata: CellMetadataIndex;
}

function openPackage(data: Uint8Array, maxUncompressedBytes: number | undefined): OpenPackage {
  const {pkg, documentPath, workbookXml} = openSpreadsheetPackage(data, maxUncompressedBytes);
  const {partText: text, partBytes} = pkg;

  // A binary `.xlsb` is a workbook this library *can* read, just not through here. Row streaming is
  // built on the XML worksheet parser, so the binary cell table has no streaming path yet; say so,
  // rather than reporting the format as unreadable when `readXlsx` would take the very same bytes.
  if (workbookXml === undefined) {
    throw unsupportedWorkbookPart(
      partBytes,
      documentPath,
      'the binary .xlsb format (BIFF12) cannot be row-streamed yet; read it with readXlsx or readXlsb',
    );
  }

  // The model is never built here, so `<workbookPr>` is read onto a bare workbook rather than out of
  // one: `workbookPropertiesPass` is the same reader the buffered path runs, which is what keeps the
  // two from disagreeing about which calendar a sheet's serials are in. Both ride one scan, as they
  // do there.
  const properties = new Workbook();
  const sheetsPass = workbookSheetsPass();
  const namesPass = definedNamesPass();
  parseXmlPasses(workbookXml, [sheetsPass, workbookPropertiesPass(properties), namesPass]);
  // Repaired by the same function `readXlsx` runs, so a streamed sheet carries the name the buffered
  // model gives it, and selecting a sheet by that name works in both. Raw, a file with two sheets named
  // `S` streamed two of them, `readSheetRows({sheet: 'S (2)'})` refused a name `readXlsx` reports, and a
  // streamed name handed to the streaming writer was refused there.
  const sheets = repairedSheetNames(sheetsPass.result());
  const rels = readPartRelationships(documentPath, text);
  const sharedStrings = parseSharedStrings(
    rels.relatedTextOrPath('sharedStrings', SHARED_STRINGS_PART),
  );
  const {cellXfs: xfStyles} = parseStyleTable(rels.relatedTextOrPath('styles', STYLES_PART));

  return {
    sheets,
    sharedStrings,
    xfStyles,
    dateEpoch: properties.dateEpoch,
    definedNames: namesPass.spellings(),
    cellMetadata: readCellMetadata(rels),
    sheetXml(relId: string): string | undefined {
      const target = rels.byId(relId)?.target;
      return target === undefined ? undefined : text(rels.pathOf(target));
    },
  };
}

function pickSheet(
  sheets: ReadonlyArray<SheetEntry>,
  selector: string | number | undefined,
): SheetEntry {
  const first = sheets[0];
  if (first === undefined) throw new XlsxParseError('workbook names no worksheets');
  if (selector === undefined) return first;
  if (typeof selector === 'number') {
    const sheet = sheets[selector - 1];
    if (sheet === undefined) throw new RangeError(`no worksheet at position ${selector}`);
    return sheet;
  }
  const sheet = sheets.find((candidate) => candidate.name === selector);
  if (sheet === undefined) throw new AuthoringError(`no worksheet named ${quoted(selector)}`);
  return sheet;
}

// A single worksheet exposed by readWorkbookStream. Its rows() re-scans on each call (a fresh pass,
// so it is safely re-iterable); the hidden-column and merge accessors reuse a completed scan or, if
// the rows were never drained, drive one of their own. The hidden/merge state is filled in by the
// same scanSheet pass that yields the rows.
class StreamedSheetReader implements StreamedSheet {
  readonly name: string;
  readonly #xml: string;
  readonly #tables: SheetTables;
  #hiddenColumns = new Set<number>();
  #merges = new WorksheetMerges();
  #scanned = false;

  constructor(name: string, xml: string, tables: SheetTables) {
    this.name = name;
    this.#xml = xml;
    this.#tables = tables;
  }

  *rows(): Generator<StreamedRow, void, undefined> {
    this.#hiddenColumns = new Set();
    this.#merges = new WorksheetMerges();
    this.#scanned = false;
    yield* scanSheet(this.#xml, this.#tables, this.#hiddenColumns, this.#merges);
    this.#scanned = true;
  }

  get hiddenColumns(): readonly number[] {
    this.#ensureScanned();
    return [...this.#hiddenColumns].sort((a, b) => a - b);
  }

  get merges(): readonly string[] {
    this.#ensureScanned();
    // Copied, like `hiddenColumns` above, and for the reason the model's collection accessors are
    // not: `rows()` assigns a fresh array on each iteration, so a caller holding the live one across
    // a second pass would be holding a detached snapshot without ever having been told. A copy makes
    // that explicit at the one place it can happen. A `Worksheet` accessor hands back its live array
    // because it *is* the owner and the array outlives the call; see the collection-accessor rule in
    // `docs/architecture.md`.
    return [...this.#merges.ranges];
  }

  // Drain a scan purely for its summaries when the caller reads them without (or before) iterating
  // rows. A completed row iteration already set #scanned, so this re-scans nothing in the common
  // streaming idiom.
  #ensureScanned(): void {
    if (this.#scanned) return;
    for (const _row of this.rows()) {
      // The rows themselves are irrelevant here; we only want the hidden/merge side effects.
    }
  }
}

// Pull the sheet XML through the event stream, yielding a StreamedRow at each `</row>`, while
// recording the sheet's hidden columns (from `<col hidden>`, before <sheetData>) and merged ranges
// (from `<mergeCells>`, after <sheetData>) into the caller-supplied collectors. The `<c>` machine is
// the accumulator's own, the same one the buffered reader drives, so the two cannot read a cell
// differently. What differs is what committing means: this one pushes into a row buffer that is
// handed off and discarded per row rather than into a persistent Worksheet, and that hand-off is
// what bounds retained memory to one row.
function* scanSheet(
  xml: string,
  tables: SheetTables,
  hiddenColumns: Set<number>,
  merges: WorksheetMerges,
): Generator<StreamedRow, void, undefined> {
  const {sharedStrings, xfStyles, dateEpoch, definedNames, cellMetadata} = tables;
  let rowNumber = 0;
  let rowHidden = false;
  let cells: StreamedCell[] = [];
  const columnBudget = new ColumnRecordBudget();
  const styleResolution = new CellStyleResolver();
  const rowPosition = new RowPositionTracker();

  // The in-flight `<c>`, gathered exactly as the buffered reader gathers it, then taken as the
  // cell's plain decoded value (via decode) rather than through the shared-formula / data-table
  // resolution the buffered finalize adds, which a data read does not want. A rich string keeps its
  // runs here as it does there, whether the producer inlined it or pooled it.
  const cell = new CellAccumulator({dateEpoch, definedNames, cellMetadata});

  const finalizeCell = (): void => {
    // Whether a cell was placed at all, a row past the grid included, is the accumulator's decision.
    if (cell.ref === '') return;
    // Through the shared resolution, xf 0 included, not the cell's own `s` alone. `decodeCellContent`
    // reads `numFmt` off the resolved style to tell a date serial from a plain number, so anything less
    // decodes a cell to a different *type* than the buffered reader does.
    const style = styleResolution.styleFor(cell.col, cell.styleIndex, xfStyles);
    const value = cell.decode(sharedStrings, style);
    // A blank or purely style-only cell decodes to null; a data read wants only cells that carry
    // something (a formula object, an empty string, a false, and a 0 all count; only null drops).
    if (value !== null) {
      const {col, ref} = cell;
      // xf 0 is every cell's floor rather than a format anything declared, so repeating it on each
      // streamed cell would tell a consumer nothing; the style is reported when something declared it.
      cells.push(
        style !== undefined && styleResolution.declaresFormat(col, cell.styleIndex)
          ? {col, address: ref, value, style}
          : {col, address: ref, value},
      );
    }
  };

  for (const event of closeEmptyElements(xmlEvents(xml), WORKSHEET_BODY_EMPTY_CLOSES)) {
    if (event.kind === 'text') {
      cell.appendChunk(event.text);
      continue;
    }
    if (event.kind === 'open') {
      const local = localName(event.name);
      if (cell.openElement(local, event.attrs, event.selfClosing)) continue;
      switch (local) {
        case 'row': {
          // A row past the grid is dropped whole, the same reading the buffered reader takes: an `r`
          // names one row, so there is nothing to clamp it onto, and yielding a `number` of 1048577
          // would hand the consumer an address no `getCell` will accept. The `<c>` machine still runs
          // over its cells, because it is what keeps the reader in step with the element stream, but
          // it places none of them and no row is handed off.
          const position = rowPosition.open(event.attrs);
          rowNumber = position.number;
          rowHidden = boolStrict(event.attrs.hidden);
          cell.openRow(position.number, position.inGrid);
          styleResolution.openRow(
            numInteger(event.attrs.s, 0) ?? -1,
            boolStrict(event.attrs.customFormat),
          );
          break;
        }
        case 'col':
          collectColumn(event.attrs, hiddenColumns, styleResolution, columnBudget);
          break;
        case 'mergeCell': {
          // Admitted through the buffered reader's own collection, not pushed raw: it canonicalises
          // the range and refuses an unreadable, sheet-qualified or overlapping one, and its index
          // keeps that overlap check from growing with the square of the file's merges.
          const ref = event.attrs.ref;
          if (ref !== undefined) admitting(() => merges.add(ref));
          break;
        }
        default:
          break;
      }
      continue;
    }
    // close
    const local = localName(event.name);
    const claimed = cell.closeElement(local);
    if (claimed === 'cell') finalizeCell();
    else if (claimed === 'other' && local === 'row') {
      styleResolution.closeRow();
      const inGrid = cell.rowOpen;
      cell.closeRow();
      if (inGrid) yield {number: rowNumber, hidden: rowHidden, cells};
      // A fresh buffer at the close rather than at the next open: the one just yielded belongs to the
      // consumer now, and a stray `<c>` between rows was being pushed into it after the hand-off.
      cells = [];
    }
  }
}

// Take what a `<col min max hidden style>` element says: which columns it hides, and the cell-format
// default its cells inherit. The span is clamped to the format's column ceiling and the hidden columns
// gathered into a Set, so even a hostile file full of full-width hidden spans can add at most
// MAX_COLUMN distinct entries, never an unbounded allocation. Memory was never the whole question
// though: a Set bounded at 16,384 entries still costs one insertion per column per element, and
// nothing bounds the element count, so the per-sheet budget bounds the time too.
function collectColumn(
  attrs: XmlAttributes,
  hiddenColumns: Set<number>,
  styleResolution: CellStyleResolver,
  budget: ColumnRecordBudget,
): void {
  const span = takeColumnSpan(attrs, budget);
  if (span === undefined) return;
  const {first: min, last} = span;
  // The span's cell-format default, which a bare `<c>` in these columns inherits: the streaming
  // reader ignored it entirely, which is what made it decode a date column's cells as numbers.
  styleResolution.noteColumnSpan(min, last, numInteger(attrs.style, 0) ?? -1);
  if (!boolStrict(attrs.hidden)) return;
  for (let index = min; index <= last; index++) hiddenColumns.add(index);
}
