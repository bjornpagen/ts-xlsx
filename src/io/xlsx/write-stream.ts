// The streaming `.xlsx` writer: author a workbook incrementally and emit its package through a Node
// stream, rather than holding the finished bytes in one buffer as {@link writeXlsx} does.
//
// A producer adds worksheets, appends rows as it generates them, commits each sheet, then commits the
// workbook, at which point the package is assembled and streamed out. The output goes to whichever
// destination the caller chose at construction: its own pull-based `stream` to pipe, a caller-owned
// `Writable` sink (an outbound upload), or a `filename` the writer opens. In every case `commit()`
// settles: it resolves once a supplied sink has flushed the whole package and rejects if that sink
// errors (an unopenable file), never hanging on a finish signal that will not come. The output rides a
// genuinely
// streamed zip container (fflate's `Zip`/`ZipDeflate`), which computes each entry's CRC-32
// incrementally, so the archive is well-formed by construction, and the defect the upstream "streaming
// writer emits a corrupt zip" reports describe is structurally absent here. The bytes reload
// identically to a whole-file write because both writers share `buildPackageParts` for every part.
//
// Peak memory: a row committed with `row.commit()` is serialised to its `<row>` XML immediately and
// its cells evicted from the model, so an append-driven producer holds only the rows still in flight
// rather than the whole sheet's cell graph. The eagerly-rendered rows intern into the workbook's live
// style registry, the same one that emits `xl/styles.xml`, so their style ids stay correct; their
// XML is handed to `buildPackageParts` and emitted at the head of `<sheetData>`. This eager path runs
// with strings inline (a shared-strings pool is inherently whole-workbook, so it defeats bounding);
// turning `useSharedStrings` on falls back to holding every row live until commit. A flushed row is a
// finished row: it cannot join whole-sheet derivations, so a shared-formula clone in a committed row
// is rejected, as is a `getCell` reaching back into a committed row (its `<row>` is already rendered,
// so the cell could only be emitted as a second row of that number), while rows reached only through
// `getCell` (never `row.commit()`) stay live and serialise the ordinary way. The package bytes
// themselves are still assembled once at commit; a later slice can flush each sheet's `<sheetData>`
// straight into its streamed zip entry to bound that half too.
//
// What `WorksheetStreamWriter` does *not* expose, since a hand-maintained subset of `Worksheet` grows
// its gaps silently: a verb added to the sheet model simply never appears here, and nothing mechanical
// would say so. An exhaustiveness proof would be the wrong artifact, because most of `Worksheet` must
// not be streamable, so the honest one is this list. Two classes are absent because a streamed sheet
// cannot answer them honestly. Structural edits (`spliceRows`, `insertRow`, `duplicateRow`, and the
// column verbs) re-anchor rows an eager writer has already rendered and evicted, so they would apply
// to the live tail alone while looking like they applied to the sheet. Whole-sheet reads (`usedRange`,
// `getRange`, `hasCell`, `actualRowCount`, `dataValidationAt`) have the same defect one step later:
// each would compute a plausible answer over whichever rows happen to still be in memory. `rowCount`
// is the exception that proves it, and is why it keeps its own high-water mark instead of asking the
// model.
//
// The third class is different, and calling it absent-by-design would be a fiction: merges, the frozen
// pane, column widths, page setup, tables and comment threads are sheet-level state the streamed
// package already emits, because both writers assemble their parts through `buildPackageParts` (a
// merge, a freeze and a column width applied to the model survive a streamed write; measured). They
// are missing because no streaming producer has asked for them yet. Adding one is a delegation and a
// test, subject to the single ordering rule this writer imposes: column defaults are frozen at the
// first flushed row, so a column defined after that is still written to `<cols>` but is composed into
// no eagerly-rendered cell.

import {createWriteStream} from 'node:fs';
import {PassThrough, type Readable, type Writable} from 'node:stream';

import {Zip, ZipDeflate} from 'fflate';

import {concat} from '../../bytes.ts';
import {encodeAddress, tryDecodeCellRef} from '../../core/address.ts';
import type {AutoFilter} from '../../core/autofilter.ts';
import type {Cell} from '../../core/cell.ts';
import type {ConditionalFormatting} from '../../core/conditional-formatting.ts';
import type {DataValidation} from '../../core/data-validation.ts';
import type {DateEpoch} from '../../core/date.ts';
import type {AnchorPoint} from '../../core/image.ts';
import {INTERNAL} from '../../core/internal.ts';
import type {SheetProtectionOptions} from '../../core/protection.ts';
import {type CellValue, isSharedFormulaValue} from '../../core/value.ts';
import {type AddImageOptions, type AddWorksheetOptions, Workbook} from '../../core/workbook.ts';
import type {ColumnProperties, Worksheet} from '../../core/worksheet.ts';
import {AuthoringError, quoted} from '../../errors.ts';
import {FIXED_ENTRY_MTIME} from '../opc/zip-mtime.ts';
import {type CommentCell, collectNotes} from './comments.ts';
import {type CollectedHyperlink, collectHyperlinks} from './hyperlinks.ts';
import {
  buildColumnDefaults,
  Extent,
  type FlushedRow,
  type FlushedSheet,
  renderRow,
} from './row-xml.ts';
import type {StyleRegistry} from './styles.ts';
import {buildPackageParts, createStyleRegistry, type WriteOptions} from './write.ts';

/** Calculation settings applied to the streamed workbook. Mirrors the {@link Workbook} flags. */
export interface CalcProperties {
  /** Ask the consumer to recalculate every formula on open: the OOXML `fullCalcOnLoad` flag. */
  fullCalcOnLoad?: boolean;
}

// A caller may supply a sink to write to (a stream or a filename) but never both. This is enforced as a
// type-level union rather than only the runtime check the constructor also carries, so passing both is
// a compile error, not a first-run surprise.
type SinkOptions =
  | {
      /**
       * Write the package to a caller-owned {@link Writable} sink: an outbound upload, a cloud-SDK
       * stream, any destination the caller controls. {@link WorkbookStreamWriter.commit} pushes every
       * chunk into it and settles only once the sink has finished (or rejects if it errors), so a
       * caller can deterministically sequence work after the upload completes. Mutually exclusive
       * with {@link filename}.
       */
      readonly stream: Writable;
      readonly filename?: never;
    }
  | {
      readonly stream?: never;
      /**
       * Write the package to a file at this path. The writer opens a `fs.createWriteStream` for it; if
       * the destination cannot be opened (bad path, name too long) the stream errors and
       * {@link WorkbookStreamWriter.commit} rejects with that I/O error rather than hanging. Mutually
       * exclusive with {@link stream}.
       */
      readonly filename: string;
    }
  | {readonly stream?: never; readonly filename?: never};

/** Options fixed at construction that shape the whole streamed package. */
export type WorkbookStreamWriterOptions = SinkOptions & {
  /**
   * Pool plain string cell values into a shared-strings table rather than storing each inline: the
   * same {@link WriteOptions.useSharedStrings} the buffered writer exposes. Off by default.
   */
  readonly useSharedStrings?: boolean;
  /**
   * Which date system this workbook's serials count in ({@link Workbook.dateEpoch}). A construction
   * option rather than a settable property, unlike on the buffered writer: an eager writer serialises
   * each row as it is committed, so a system changed part-way through would leave the rows before the
   * change counting from a different day than the rows after it.
   */
  readonly dateEpoch?: DateEpoch;
};

/**
 * A row appended to a {@link WorksheetStreamWriter}. Style its cells through {@link cells}, then call
 * {@link commit} to mark it finished. In an eager (inline-strings) writer, committing serialises the
 * row and frees its cells from the model, bounding peak memory; with `useSharedStrings` on it is a
 * no-op and the row stays live until the workbook commits.
 */
export class StreamedRow {
  readonly #cells: readonly Cell[];
  readonly #sheet: WorksheetStreamWriter | null;
  readonly #number: number;
  #committed = false;

  // Private, and reached from `WorksheetStreamWriter.addRow` through the static channel below. A
  // public constructor let `new StreamedRow([], sheetWriter, n).commit()` evict row `n` from the model
  // and write nothing for it: the number is the writer's to hand out, not a caller's to name.
  private constructor(cells: readonly Cell[], sheet: WorksheetStreamWriter | null, number: number) {
    this.#cells = cells;
    this.#sheet = sheet;
    this.#number = number;
  }

  /** The sheet writer's construction channel, keyed like its own; see `core/internal.ts`. */
  static readonly [INTERNAL]: StreamedRowFactory = {
    create(cells, sheet, number) {
      return new StreamedRow(cells, sheet, number);
    },
  };

  /** The cells `addRow` materialised, for styling before it is committed. A cell added to the row
   * through {@link WorksheetStreamWriter.getCell} is not among them, and is committed all the same. */
  get cells(): readonly Cell[] {
    return this.#cells;
  }

  /** Finalise the row: an eager writer serialises it now and releases its cells; otherwise a no-op.
   * Committing twice is harmless: the second call does nothing rather than re-emitting the row. */
  commit(): void {
    if (this.#committed) return;
    this.#committed = true;
    this.#sheet?.[INTERNAL].flushRow(this.#number);
  }
}

/**
 * How a {@link StreamedRow} comes into being. A row belongs to the sheet writer that numbered it, and
 * only that writer knows which number is still open. `sheet` is `null` for a writer that keeps every
 * row live until the workbook commits, where committing a row has nothing to do. Reached as
 * `StreamedRow[INTERNAL]`.
 */
export interface StreamedRowFactory {
  create(cells: readonly Cell[], sheet: WorksheetStreamWriter | null, number: number): StreamedRow;
}

/**
 * A worksheet being written incrementally. Append rows with {@link addRow}/{@link addRows}, style
 * cells through {@link getCell}, then {@link commit} to freeze it, after which any further mutation
 * is rejected with a legible error rather than silently accepted or crashing.
 */
export class WorksheetStreamWriter {
  readonly #sheet: Worksheet;
  readonly #eager: boolean;
  readonly #styles: StyleRegistry;
  #committed = false;
  // The last row number this writer appended, tracked independently of the model: an eager writer
  // evicts flushed rows, which lowers the model's used range, so leaning on it would reuse numbers.
  #lastRow = 0;
  // The column defaults an eagerly-rendered row inherits, frozen at the first flush so every flushed
  // row composes against the same columns even as later ones are defined.
  #columnDefaults: ReadonlyMap<number, ColumnProperties> | undefined;
  readonly #flushedRows: FlushedRow[] = [];
  // Every row number this writer has flushed. A flushed row's `<row>` is already rendered and its
  // cells are gone from the model, so re-materialising one would emit a second element with the same
  // number; the set is what lets `getCell` refuse that rather than produce it.
  readonly #flushedNumbers = new Set<number>();
  readonly #extent = new Extent();
  readonly #rowOutline = new Map<number, {outlineLevel: number; hidden: boolean}>();
  // What a flushed row carried that is serialised outside its `<row>`, taken before eviction.
  readonly #hyperlinks: CollectedHyperlink[] = [];
  readonly #notes: CommentCell[] = [];

  readonly #dateEpoch: DateEpoch;

  // Private, and reached from `WorkbookStreamWriter.addWorksheet` through the static channel below.
  // A caller never builds one of these -- they receive it from `addWorksheet` -- and the parameters
  // say why that matters: a `StyleRegistry` is the writer's own interning table, so a public
  // constructor made it a type a consumer could hold and could not name.
  private constructor(
    sheet: Worksheet,
    eager: boolean,
    styles: StyleRegistry,
    dateEpoch: DateEpoch,
  ) {
    this.#sheet = sheet;
    this.#eager = eager;
    this.#styles = styles;
    this.#dateEpoch = dateEpoch;
  }

  /**
   * The workbook writer's construction channel. A static member keyed by the same symbol as the
   * instance one below, so both halves of this class's internal wiring are reached the one way; see
   * `core/internal.ts`.
   */
  static readonly [INTERNAL]: WorksheetStreamWriterFactory = {
    create(sheet, eager, styles, dateEpoch) {
      return new WorksheetStreamWriter(sheet, eager, styles, dateEpoch);
    },
  };

  /** The sheet's name. */
  get name(): string {
    return this.#sheet.name;
  }

  /** The number of rows written so far. Spans gaps and formatted-only rows, like the model, and
   * survives the eviction of eagerly-flushed rows. */
  get rowCount(): number {
    return Math.max(this.#lastRow, this.#sheet.rowCount);
  }

  /** Append one row of values after the last used row; the cells are returned for styling. */
  addRow(values: CellValue[]): StreamedRow {
    this.#assertOpen();
    if (!this.#eager) return StreamedRow[INTERNAL].create(this.#sheet.addRow(values), null, 0);
    const number = this.#nextRowNumber();
    return StreamedRow[INTERNAL].create(this.#placeRow(number, values), this, number);
  }

  /** Append a batch of rows in one call, each landing directly below the previous. */
  addRows(rows: CellValue[][]): StreamedRow[] {
    this.#assertOpen();
    if (!this.#eager)
      return this.#sheet.addRows(rows).map((cells) => StreamedRow[INTERNAL].create(cells, null, 0));
    return rows.map((values) => this.addRow(values));
  }

  // The next append position: past both this writer's own high-water mark and any rows a `getCell`
  // materialised, so appends never collide with random-access edits or with already-evicted rows.
  #nextRowNumber(): number {
    this.#lastRow = Math.max(this.#lastRow, this.#sheet.rowCount) + 1;
    return this.#lastRow;
  }

  // Materialise a positional row at an explicit number (the model's own append would reuse numbers once
  // eviction shrinks its used range), returning the cells for styling. A hole leaves its cell absent.
  #placeRow(number: number, values: CellValue[]): Cell[] {
    const cells: Cell[] = [];
    values.forEach((value, index) => {
      const cell = this.#sheet.getCell(encodeAddress(index + 1, number));
      cell.value = value;
      cells.push(cell);
    });
    return cells;
  }

  // The row is read off the model as it stands now rather than as `addRow` left it: `getCell` may add
  // cells to an open row, and the eviction below takes every cell the model holds for it, so a cell
  // missing from what is rendered here would be gone from the file with nothing to say so.
  #flushRow(number: number): void {
    const row = this.#sheet.getRow(number);
    const {cells} = row;
    for (const cell of cells) {
      if (isSharedFormulaValue(cell.value)) {
        throw new AuthoringError(
          `row ${number} of streamed sheet ${quoted(this.#sheet.name)} carries a shared-formula cell; a ` +
            'committed row is finalised before the sheet is planned, so author shared formulas through ' +
            'getCell (leaving the row uncommitted) instead',
        );
      }
    }
    this.#columnDefaults ??= buildColumnDefaults(this.#sheet);
    const {properties} = row;

    // The row's outline level and hidden flag are read off here because eviction is about to take its
    // properties with it, and both feed whole-sheet derivations made long afterwards: `<sheetFormatPr
    // outlineLevelRow>` is the deepest level on the sheet, and `collapsed="1"` rides a summary row only
    // when its whole detail group is hidden. Both are questions about other rows, so the inputs are
    // carried and the derivation runs at commit.
    this.#rowOutline.set(number, {
      outlineLevel: properties?.outlineLevel ?? 0,
      hidden: properties?.hidden ?? false,
    });
    // A hyperlink and a note are serialised outside the `<row>`, into the sheet's `<hyperlinks>`
    // element and into the comments/VML parts, and the buffered pass gathers both by walking the
    // sheet's rows at commit. Eviction is about to make that walk find nothing, so they are taken
    // here for the same reason the outline level is.
    this.#hyperlinks.push(...collectHyperlinks(cells));
    this.#notes.push(...collectNotes(cells));
    const {xml, attrs, minCol, maxCol} = renderRow(
      {number, cells, properties},
      {
        columnDefaults: this.#columnDefaults,
        styles: this.#styles,
        sharedStrings: null,
        sharedRoles: new Map(),
        collapsedSummaries: new Set(),
        dateEpoch: this.#dateEpoch,
      },
    );
    if (xml !== '') {
      this.#flushedRows.push({number, xml, attrs});
      this.#extent.add(number, minCol, maxCol);
    }
    this.#flushedNumbers.add(number);
    this.#sheet[INTERNAL].evictRow(number);
  }

  #flushedSheet(): FlushedSheet | undefined {
    if (this.#flushedRows.length === 0) return undefined;
    return {
      rows: this.#flushedRows,
      extent: this.#extent,
      rowOutline: this.#rowOutline,
      hyperlinks: this.#hyperlinks,
      notes: this.#notes,
    };
  }

  /**
   * Address a cell by its A1 reference to read or style it before the sheet is committed. The row it
   * names must still be live: a row whose {@link StreamedRow.commit} has run is finished.
   *
   * @throws {AuthoringError} if the reference names an already-committed row. That row's `<row>` is
   *   rendered and its cells are released, so the cell this would materialise could only be written
   *   as a second row carrying the same number.
   */
  getCell(reference: string): Cell {
    this.#assertOpen();
    const target = tryDecodeCellRef(reference);
    if (target !== undefined && this.#flushedNumbers.has(target.row)) {
      throw new AuthoringError(
        `cell ${reference} of streamed sheet ${quoted(this.#sheet.name)} is in row ${target.row}, which is ` +
          'already committed: its bytes are written and its cells released, so this value could only ' +
          'be emitted as a second row with that number. Style a row before committing it.',
      );
    }
    return this.#sheet.getCell(reference);
  }

  /**
   * Attach a data validation to a range before the sheet is committed. Delegates to the model, so the
   * streamed package emits the `<dataValidations>` block in its CT_Worksheet position, before
   * `<hyperlinks>`, because both writers share one worksheet serializer.
   */
  addDataValidation(sqref: string, rule: DataValidation, options: {extended?: boolean} = {}): void {
    this.#assertOpen();
    this.#sheet.addDataValidation(sqref, rule, options);
  }

  /**
   * Attach a conditional formatting to a range before the sheet is committed. Like every other block,
   * it lands in its schema-mandated slot, after `<mergeCells>` and before `<dataValidations>` and
   * `<hyperlinks>`, since the streamed sheet is serialized through the same path as a buffered write.
   */
  addConditionalFormatting(formatting: ConditionalFormatting): void {
    this.#assertOpen();
    this.#sheet.addConditionalFormatting(formatting);
  }

  /**
   * Anchor a workbook image (the id from {@link WorkbookStreamWriter.addImage}) to this sheet,
   * spanning the rectangle from the top-left grid point `tl` to the bottom-right `br`. The streamed
   * package emits the drawing part, its media relationship, and the sheet's `<drawing>` reference
   * exactly as a buffered write does: both writers share `buildPackageParts`.
   */
  addImage(imageId: number, anchor: {readonly tl: AnchorPoint; readonly br: AnchorPoint}): void {
    this.#assertOpen();
    this.#sheet.addImage(imageId, anchor);
  }

  /**
   * Apply the sheet's autofilter before it is committed; mirrors {@link Worksheet.autoFilter}. The
   * streamed package emits `<autoFilter>` in its CT_Worksheet slot, after `<sheetProtection>`, and
   * contributes the hidden `_FilterDatabase` defined name, exactly as a buffered write does.
   */
  set autoFilter(filter: string | AutoFilter | undefined) {
    this.#assertOpen();
    this.#sheet.autoFilter = filter;
  }

  get autoFilter(): AutoFilter | undefined {
    return this.#sheet.autoFilter;
  }

  /**
   * Apply sheet-level protection before the sheet is committed; mirrors {@link Worksheet.protect}. The
   * shared serializer places `<sheetProtection>` ahead of `<autoFilter>` per CT_Worksheet, so a
   * streamed sheet carrying both stays valid rather than corrupt.
   */
  protect(password?: string, options: SheetProtectionOptions = {}): void {
    this.#assertOpen();
    this.#sheet.protect(password, options);
  }

  /** Freeze the sheet: no more rows or edits may be added after this. */
  commit(): void {
    this.#committed = true;
  }

  /** Whether the sheet has been committed. */
  get committed(): boolean {
    return this.#committed;
  }

  // Internal: the underlying model sheet, so the workbook writer can serialise it at commit time.
  get model(): Worksheet {
    return this.#sheet;
  }

  readonly [INTERNAL]: WorksheetStreamWriterInternals = {
    flushRow: (number) => {
      this.#flushRow(number);
    },
    flushedSheet: () => this.#flushedSheet(),
  };

  #assertOpen(): void {
    if (this.#committed) {
      throw new AuthoringError(
        `worksheet ${quoted(this.#sheet.name)} is already committed: its rows are finalised and no more can be added`,
      );
    }
  }
}

/**
 * How a {@link WorksheetStreamWriter} comes into being. Its parameters are the writer's own plumbing
 * -- the interning table its rows' style ids are assigned from, the epoch its dates are serialised
 * against -- and a consumer supplies neither, because a consumer receives the sheet writer from
 * {@link WorkbookStreamWriter.addWorksheet}. Reached as `WorksheetStreamWriter[INTERNAL]`.
 */
export interface WorksheetStreamWriterFactory {
  create(
    sheet: Worksheet,
    eager: boolean,
    styles: StyleRegistry,
    dateEpoch: DateEpoch,
  ): WorksheetStreamWriter;
}

/**
 * What the streaming writer's own machinery may do to a sheet writer that a caller may not, the same
 * boundary `sheet[INTERNAL]` draws around the model. Reached as `writer[INTERNAL]`; see
 * `core/internal.ts`.
 *
 * Both members were public until they were not. Neither is callable by anyone but this module, and
 * between them they named seven writer-internal types -- `StyleRegistry`, `FlushedSheet` and the
 * shapes those reach -- on a published class's signature, so a consumer could hold values whose types
 * the package would not export. They carried `@unpublished` to record that, which is a note rather
 * than a boundary; this is the boundary.
 */
export interface WorksheetStreamWriterInternals {
  /**
   * Serialise an eagerly-committed row and release its cells from the model. Called by
   * {@link StreamedRow.commit}; the row's `<row>` XML is retained (interned into the workbook's live
   * style registry so its ids stay valid) and the cell graph is dropped, bounding peak memory. The
   * row is rendered from every cell the model holds for it, including any `getCell` added after
   * `addRow` returned.
   *
   * @throws {AuthoringError} if the row carries a shared-formula cell: a finished row cannot join the
   *   whole-sheet formula planning, so shared formulas must be authored through
   *   {@link WorksheetStreamWriter.getCell}.
   */
  flushRow(number: number): void;

  /** The rows this writer flushed, or undefined if none. Handed to `buildPackageParts` at commit. */
  flushedSheet(): FlushedSheet | undefined;
}

/**
 * A workbook written incrementally to a Node stream. Add worksheets, append their rows, commit each
 * sheet, then {@link commit} the workbook to assemble and stream the package. The produced bytes are
 * available both as the resolved value of `commit()` and through {@link stream} (a Node `Readable`
 * that a caller can `pipe`).
 */
export class WorkbookStreamWriter {
  readonly #workbook = new Workbook();
  readonly #sheets: WorksheetStreamWriter[] = [];
  readonly #writeOptions: WriteOptions;
  readonly #sink: Writable | undefined;
  // The single style registry shared by the eager per-row flush and the commit-time serialisation, so a
  // flushed row's style ids match the styles.xml built from the same table.
  readonly #styles: StyleRegistry;
  // Eager per-row flushing runs with strings inline; a shared-strings pool is inherently whole-workbook,
  // so it cannot bound memory: turning it on keeps every row live until commit.
  readonly #eager: boolean;
  #stream: PassThrough | undefined;
  #committed = false;

  /** Calculation settings for the workbook; set `fullCalcOnLoad` before committing to emit it. */
  readonly calcProperties: CalcProperties = {};

  constructor(options: WorkbookStreamWriterOptions = {}) {
    this.#writeOptions = {useSharedStrings: options.useSharedStrings ?? false};
    this.#eager = !this.#writeOptions.useSharedStrings;
    if (options.dateEpoch !== undefined) this.#workbook.dateEpoch = options.dateEpoch;
    this.#styles = createStyleRegistry(this.#workbook);
    if (options.stream && options.filename) {
      throw new AuthoringError(
        'provide either a stream or a filename to the streaming writer, not both',
      );
    }
    this.#sink =
      options.stream ?? (options.filename ? createWriteStream(options.filename) : undefined);
  }

  /** Document-level metadata written to the package's core properties. */
  get properties(): Workbook['properties'] {
    return this.#workbook.properties;
  }

  /**
   * The output stream carrying the package bytes. A caller drives it with Node's standard idiom,
   * `writer.stream.pipe(out)`, which composes because `pipe` returns its destination. The stream is
   * created lazily on first access so a caller handing the writer its own sink is still free to
   * ignore this one.
   */
  get stream(): Readable {
    this.#stream ??= new PassThrough();
    return this.#stream;
  }

  /**
   * Register a picture's bytes on the workbook's shared media registry and return its id, to anchor
   * on any sheet with {@link WorksheetStreamWriter.addImage}. Mirrors {@link Workbook.addImage}: one
   * media part backs an image anchored on several sheets. Rejected once the workbook is committed.
   */
  addImage(options: AddImageOptions): number {
    this.#assertOpen('no more images can be registered');
    return this.#workbook.addImage(options);
  }

  /** Create a worksheet and append it to the workbook. */
  addWorksheet(name: string, options: AddWorksheetOptions = {}): WorksheetStreamWriter {
    this.#assertOpen('no more worksheets can be added');
    const sheet = WorksheetStreamWriter[INTERNAL].create(
      this.#workbook.addWorksheet(name, options),
      this.#eager,
      this.#styles,
      this.#workbook.dateEpoch,
    );
    this.#sheets.push(sheet);
    return sheet;
  }

  /**
   * Assemble the workbook into its package and stream the bytes out. Every sheet is frozen first, so
   * a row added after this rejects legibly. Idempotent only in that a second call throws rather than
   * re-emitting.
   *
   * **What it resolves with, and why it can be `undefined`.** The archive is handed back only when
   * somebody asked for it: when no sink was supplied, or when {@link stream} was touched, since a
   * `PassThrough` nobody drained would otherwise be the only copy. A caller who passed `{filename}`
   * or their own `stream` and never reached for {@link stream} gets `undefined`, and the package is
   * never materialised as a whole: retaining every chunk to concatenate them at the end costs a
   * second full-size buffer on top of the part map and the archive itself, which is exactly the
   * memory the caller passed a sink to avoid. The type says so rather than the prose alone, because
   * a promise that resolves with a value the caller was told to ignore is a promise they will use.
   *
   * If assembling or zipping the package fails, the returned promise rejects *and* every stream this
   * writer was given or handed out is destroyed with that error. A caller piping {@link stream}, or
   * awaiting its own sink, therefore sees an `error` rather than waiting on a stream that will never
   * end; and it is an `error` rather than a clean `end` because the bytes written so far are a
   * truncated package that must not be read as a whole one.
   */
  async commit(): Promise<Uint8Array | undefined> {
    this.#assertOpen('its package is already assembled and cannot be re-emitted');
    this.#committed = true;
    for (const sheet of this.#sheets) sheet.commit();
    if (this.calcProperties.fullCalcOnLoad) this.#workbook.fullCalcOnLoad = true;

    // Hand every sheet's eagerly-flushed rows to the shared serialiser, which emits them alongside the
    // rows still live in the model. The style registry is the same one the flushed rows interned into,
    // so styles.xml stays consistent with the ids already baked into their XML.
    const flushed = new Map<Worksheet, FlushedSheet>();
    for (const sheet of this.#sheets) {
      const sheetFlushed = sheet[INTERNAL].flushedSheet();
      if (sheetFlushed) flushed.set(sheet.model, sheetFlushed);
    }
    const owned = this.#stream;
    const sink = this.#sink;
    // Nobody is waiting for the whole archive when the caller supplied a sink and never reached for
    // `stream`. Read here, beside `owned`, so the decision and the streams it is about are one
    // statement rather than two that could be made at different moments.
    const collect = sink === undefined || owned !== undefined;
    // Track the caller sink's terminal state before writing a byte, so an open failure that errors on a
    // later tick (a bad filename) is caught rather than lost, which is the whole point of the reject-not-hang
    // contract. Ahead of the serialisation too, which throws routinely: a rejection handler created
    // after the first possible throw is a handler the failure path cannot reach.
    const sinkSettled = sink ? settleOnFinish(sink) : undefined;

    try {
      const parts = buildPackageParts(this.#workbook, {
        ...this.#writeOptions,
        styles: this.#styles,
        flushed,
      });
      const bytes = await streamZipPackage(
        parts,
        (chunk) => {
          owned?.write(chunk);
          sink?.write(chunk);
        },
        collect,
      );
      owned?.end();
      sink?.end();

      await sinkSettled;
      return bytes;
    } catch (error) {
      // `destroy(err)`, never `end()`: ending says the package is complete, and a consumer handed a
      // truncated archive that ended cleanly will read it as whole. Destroying propagates through
      // `pipe` as the destination's `error`, which is what a `pipeline()` caller already handles.
      const failure = error instanceof Error ? error : new Error(String(error));
      // The sink is about to error because we are destroying it, and that is not the failure worth
      // reporting: the original is. Claim the rejection so it does not surface as unhandled.
      sinkSettled?.catch(() => {});
      owned?.destroy(failure);
      sink?.destroy(failure);
      throw error;
    }
  }

  #assertOpen(what: string): void {
    if (this.#committed) {
      throw new AuthoringError(`the workbook is already committed: ${what}`);
    }
  }
}

// Resolve when a caller-supplied sink has flushed the whole package (`finish`), or reject if it errors
// (`error`): the commit promise must settle either way, never hang. Whichever fires first wins; the
// other is ignored.
function settleOnFinish(sink: Writable): Promise<void> {
  return new Promise((resolve, reject) => {
    sink.once('finish', resolve);
    sink.once('error', reject);
  });
}

// Zip the package parts through fflate's streaming container, forwarding each output chunk to `onChunk`
// as it is produced and resolving with the whole archive once the final chunk arrives. `ZipDeflate`
// deflates synchronously, so the callback fires inline as each part is pushed, and the CRC-32 fflate
// stamps into every entry's header therefore always matches the bytes it just compressed.
//
// `collect` is what makes this a streaming writer on the output side as well as the input side. Every
// chunk used to be retained whatever the caller wanted, and `concat` then allocated a second
// full-size buffer to join them, so a caller who passed a sink precisely to avoid holding the package
// held two copies of it.
function streamZipPackage(
  parts: Record<string, Uint8Array>,
  onChunk: (chunk: Uint8Array) => void,
  collect: boolean,
): Promise<Uint8Array | undefined> {
  return new Promise((resolve, reject) => {
    const collected: Uint8Array[] = [];
    const zip = new Zip((err, chunk, final) => {
      if (err) {
        reject(err);
        return;
      }
      if (collect) collected.push(chunk);
      onChunk(chunk);
      if (final) resolve(collect ? concat(collected) : undefined);
    });
    for (const [name, data] of Object.entries(parts)) {
      const entry = new ZipDeflate(name, {level: 6});
      // `mtime` is a field on the entry rather than a compression option (`ZipDeflate` forwards its
      // options to the deflater alone) and the header is written when the entry is added, so it has
      // to be set before that call rather than before the first push.
      entry.mtime = FIXED_ENTRY_MTIME;
      zip.add(entry);
      entry.push(data, true);
    }
    zip.end();
  });
}
