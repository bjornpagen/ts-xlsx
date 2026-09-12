/**
 * The three things the playground actually does: write, read, and round-trip.
 *
 * Every lane returns a result rather than throwing, because the caller is a page and a
 * thrown error there becomes a blank panel. A failure is a result too, and one worth
 * showing: a reader learning that this library refuses their file with a named error has
 * learned something true about it.
 *
 * Elapsed time is measured around the library call and nothing else. A number on the page
 * that includes the caller's own work is a number that lies.
 *
 * Nothing here touches a DOM, and nothing here makes a network call.
 */

import {
  type Cell,
  type CellValue,
  isDataTableFormulaValue,
  isErrorValue,
  isFormulaValue,
  isRichTextValue,
  isSharedFormulaValue,
  readXlsx,
  type Workbook,
  type Worksheet,
  writeXlsx,
} from '../../src/index.ts';

export type Lane<T> =
  | {readonly ok: true; readonly value: T}
  | {readonly ok: false; readonly error: string};

export interface WriteFacts {
  readonly bytes: Uint8Array;
  readonly byteLength: number;
  readonly elapsedMs: number;
}

export interface SheetFacts {
  readonly name: string;
  readonly rowCount: number;
  readonly columnCount: number;
  /** The A1 range the sheet actually occupies, or nothing when the sheet is empty. */
  readonly usedRange: string | undefined;
}

export interface ReadFacts {
  readonly workbook: Workbook;
  readonly sheets: readonly SheetFacts[];
  /**
   * Parts the model does not interpret and carries verbatim, by package path. This is the
   * claim worth showing: what the library does not understand, it does not throw away.
   */
  readonly preservedParts: readonly string[];
  readonly elapsedMs: number;
}

export interface RoundTripFacts {
  readonly agrees: boolean;
  /** Where the two reads first disagree, as a path into the compared model. */
  readonly firstDisagreement: string | undefined;
  readonly firstByteLength: number;
  readonly secondByteLength: number;
  /** True when the two writes produced identical bytes, which this writer guarantees. */
  readonly bytesIdentical: boolean;
  readonly elapsedMs: number;
}

const message = (err: unknown): string =>
  err instanceof Error ? `${err.name}: ${err.message}` : String(err);

export function write(workbook: Workbook): Lane<WriteFacts> {
  try {
    const started = performance.now();
    const bytes = writeXlsx(workbook);
    const elapsedMs = performance.now() - started;
    return {ok: true, value: {bytes, byteLength: bytes.length, elapsedMs}};
  } catch (err: unknown) {
    return {ok: false, error: message(err)};
  }
}

export function read(bytes: Uint8Array): Lane<ReadFacts> {
  try {
    const started = performance.now();
    const workbook = readXlsx(bytes);
    const elapsedMs = performance.now() - started;
    return {
      ok: true,
      value: {
        workbook,
        sheets: sheetFacts(workbook),
        preservedParts: preservedPaths(workbook),
        elapsedMs,
      },
    };
  } catch (err: unknown) {
    return {ok: false, error: message(err)};
  }
}

/**
 * Read, write, read again, and say whether the two reads agree.
 *
 * The models are compared, never the bytes, because the bytes are allowed to differ: entry
 * order and a shared-strings table are both choices a writer may make differently from the
 * one that produced the input. What must not differ is what a consumer sees. The two writes
 * of *our* output are compared byte for byte as well, since this writer does promise that,
 * and a promise that is cheap to check should be checked.
 */
export function roundTrip(bytes: Uint8Array): Lane<RoundTripFacts> {
  try {
    const started = performance.now();
    const first = readXlsx(bytes);
    const firstBytes = writeXlsx(first);
    const second = readXlsx(firstBytes);
    const secondBytes = writeXlsx(second);
    const elapsedMs = performance.now() - started;
    const firstDisagreement = difference(comparable(first), comparable(second), '');
    return {
      ok: true,
      value: {
        agrees: firstDisagreement === undefined,
        firstDisagreement,
        firstByteLength: firstBytes.length,
        secondByteLength: secondBytes.length,
        bytesIdentical: sameBytes(firstBytes, secondBytes),
        elapsedMs,
      },
    };
  } catch (err: unknown) {
    return {ok: false, error: message(err)};
  }
}

function sheetFacts(workbook: Workbook): SheetFacts[] {
  return workbook.worksheets.map((sheet) => ({
    name: sheet.name,
    rowCount: sheet.actualRowCount,
    columnCount: sheet.columnCount,
    usedRange: sheet.usedRange?.address,
  }));
}

function preservedPaths(workbook: Workbook): string[] {
  const paths = new Set<string>();
  for (const reference of workbook.preservedReferences) {
    for (const part of reference.parts) paths.add(part.path);
  }
  for (const reference of workbook.preservedRootReferences) {
    for (const part of reference.parts) paths.add(part.path);
  }
  for (const sheet of workbook.worksheets) {
    for (const reference of sheet.preservedReferences) {
      for (const part of reference.parts) paths.add(part.path);
    }
  }
  return [...paths].sort();
}

// ---------------------------------------------------------------------------
// The comparable model
//
// A hand-written projection rather than the corpus's normalisers, which are deliberately
// untyped at the adapter boundary (ADR 0038) and would drag that boundary into a tree where
// `any` is not allowed. What is compared is what a consumer of the model can observe: the
// values, the style facets a cell actually carries, and the sheet-level structure. What is
// left out is left out because it is legally allowed to move: entry order, timestamps, and
// the ids the writer assigns.
// ---------------------------------------------------------------------------

export type Comparable =
  | null
  | boolean
  | number
  | string
  | readonly Comparable[]
  | {readonly [key: string]: Comparable};

function comparable(workbook: Workbook): Comparable {
  return {sheets: workbook.worksheets.map(comparableSheet)};
}

function comparableSheet(sheet: Worksheet): Comparable {
  const cells: Record<string, Comparable> = {};
  const used = sheet.usedRange;
  if (used !== undefined) {
    for (const cell of used.cells) {
      const projected = comparableCell(cell);
      if (projected !== null) cells[cell.address] = projected;
    }
  }
  return {
    name: sheet.name,
    state: sheet.state,
    merges: [...sheet.merges].sort(),
    usedRange: used?.address ?? null,
    columns: [...sheet.columns()].map((column) => ({
      index: column.index,
      width: column.width ?? null,
      hidden: column.hidden ?? false,
      numFmt: column.numFmt ?? null,
    })),
    freeze: {
      xSplit: sheet.view.xSplit ?? 0,
      ySplit: sheet.view.ySplit ?? 0,
      state: sheet.view.state ?? null,
    },
    tables: sheet.tables.map((table) => ({name: table.name, ref: table.range})),
    hyperlinks: sheet.hyperlinks.map(({ref, target, tooltip}) => ({
      ref,
      target,
      tooltip: tooltip ?? null,
    })),
    cells,
  };
}

/** A cell with neither a value nor a style contributes nothing, so it is not compared. */
function comparableCell(cell: Cell): Comparable | null {
  const value = comparableValue(cell.value);
  const style = cell.style;
  const facets: Record<string, Comparable> = {};
  if (style.numFmt !== undefined) facets['numFmt'] = style.numFmt;
  if (style.font !== undefined) facets['font'] = json(style.font);
  if (style.fill !== undefined) facets['fill'] = json(style.fill);
  if (style.border !== undefined) facets['border'] = json(style.border);
  if (style.alignment !== undefined) facets['alignment'] = json(style.alignment);
  if (style.protection !== undefined) facets['protection'] = json(style.protection);
  if (cell.note !== undefined) facets['note'] = cell.note;
  if (value === null && Object.keys(facets).length === 0) return null;
  return {value, ...facets};
}

function comparableValue(value: CellValue): Comparable {
  if (
    value === null ||
    typeof value === 'number' ||
    typeof value === 'string' ||
    typeof value === 'boolean'
  ) {
    return value;
  }
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? null : value.toISOString();
  }
  if (isErrorValue(value)) return {error: value.error};
  if (isSharedFormulaValue(value)) {
    return {
      sharedFormula: value.sharedFormula,
      formula: value.formula ?? null,
      result: comparableResult(value.result),
    };
  }
  if (isDataTableFormulaValue(value)) {
    return {dataTable: value.ref, result: comparableResult(value.result)};
  }
  if (isFormulaValue(value))
    return {formula: value.formula, result: comparableResult(value.result)};
  if (isRichTextValue(value)) {
    return {
      richText: value.richText.map((run) => ({text: run.text, font: json(run.font ?? null)})),
    };
  }
  return null;
}

function comparableResult(result: CellValue | undefined): Comparable {
  return result === undefined ? null : comparableValue(result);
}

/** Structural facets are plain data; this drops `undefined` members so two shapes compare equal. */
function json(value: unknown): Comparable {
  return JSON.parse(JSON.stringify(value ?? null)) as Comparable;
}

/**
 * The first path at which two comparable models differ, or nothing when they agree.
 *
 * A path rather than a boolean, because "the round-trip disagreed" is not a finding anyone
 * can act on and `sheets[0].cells.B2.numFmt` is.
 */
export function difference(left: Comparable, right: Comparable, path: string): string | undefined {
  if (left === right) return undefined;
  const at = path === '' ? '(root)' : path;
  if (isList(left) !== isList(right)) return at;
  if (isList(left) && isList(right)) {
    if (left.length !== right.length) return `${at}.length`;
    for (const [index, item] of left.entries()) {
      const found = difference(item, right[index] ?? null, `${path}[${index}]`);
      if (found !== undefined) return found;
    }
    return undefined;
  }
  if (isRecord(left) && isRecord(right)) {
    const keys = [...new Set([...Object.keys(left), ...Object.keys(right)])].sort();
    for (const key of keys) {
      const found = difference(
        left[key] ?? null,
        right[key] ?? null,
        path === '' ? key : `${path}.${key}`,
      );
      if (found !== undefined) return found;
    }
    return undefined;
  }
  return at;
}

/**
 * `Array.isArray` on its own narrows a readonly array to `any[]`, which turns every element
 * read below into an unchecked one. Naming the predicate keeps the element type.
 */
function isList(value: Comparable): value is readonly Comparable[] {
  return Array.isArray(value);
}

function isRecord(value: Comparable): value is {readonly [key: string]: Comparable} {
  return typeof value === 'object' && value !== null && !isList(value);
}

function sameBytes(left: Uint8Array, right: Uint8Array): boolean {
  if (left.length !== right.length) return false;
  for (const [index, byte] of left.entries()) {
    if (byte !== right[index]) return false;
  }
  return true;
}
