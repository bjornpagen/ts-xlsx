// What a reader does when a file describes something the model refuses.
//
// The taxonomy says an `AuthoringError` "is always the calling code that is wrong, never the input
// file; a malformed file raises a 'malformed-input' error instead" (`src/errors.ts`). The read path
// broke that every time it handed a file-derived string to a model method that validates: a
// duplicate sheet name, a 40-character one, a table called `1 bad`, and `readXlsx` threw
// `AuthoringError`, `SyntaxError` or `RangeError` at a caller who had done nothing but open a file.
// The last two are the worse half, because they are *native*: `catch (e) { if (e instanceof
// XlsxError) … }`, the one-line answer the taxonomy promises, does not see them at all, so the
// caller cannot tell "your file is broken" from "my own code threw".
//
// Excel repairs every one of those inputs on load, and that is the standard this module holds every
// reader to. It sits below both codecs because the rule binds both: while it lived inside the XML
// codec the layering gate kept the BIFF12 reader from importing it, and that reader handed sheet
// names, merges and defined names to the model unrepaired. There are two ways to meet the standard
// and they are not interchangeable:
//
//   - {@link repairSheetName}, where the model's rule is a *naming* rule and there is an obviously
//     right answer. A tab has to be called something, and a sheet dropped for its name would take
//     its cells, its position in the order, and every `localSheetId` that indexes past it.
//   - {@link admitting}, where the refusal is about the construct itself and there is nothing to
//     repair. A table whose name is not an identifier, a defined name scoped to a sheet that is not
//     there: the honest reading is that the file does not really carry that feature, so the feature
//     is dropped and the rest of the workbook is read.
//
// `admitting` is deliberately the *only* place the read path swallows a model refusal, so the set of
// constructs a corrupt file may silently lose is a list of call sites rather than a habit. It
// catches exactly the three types a model method raises about its own input. Anything else goes
// straight through: an `XlsxError` from a layer below, an `InternalError` of ours.

import {arrayRangeRepair, formulaPlacement} from '../../core/array-formula-ranges.ts';
import {INVALID_SHEET_NAME_CHARS, MAX_SHEET_NAME_LENGTH} from '../../core/limits.ts';
import {type CellValue, isFormulaValue, isSharedFormulaValue} from '../../core/value.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {AuthoringError} from '../../errors.ts';

/** Every occurrence, where {@link INVALID_SHEET_NAME_CHARS} tests for the first. */
const INVALID_SHEET_NAME_CHARS_GLOBAL = new RegExp(INVALID_SHEET_NAME_CHARS.source, 'g');

/**
 * A sheet name from a file, rewritten into one the model will accept.
 *
 * Excel's own repair, in the order the constraints interact: forbidden characters go first (they
 * can be anywhere), then the length ceiling, then the apostrophe rule *again*, because truncating
 * can expose an interior apostrophe at the new edge. An empty result becomes `Sheet{n}`, and a
 * collision takes a ` (2)` suffix that is itself made to fit inside the 31-character limit.
 *
 * @param taken the lower-cased names already in the workbook; sheet names collide
 *   case-insensitively, which is the comparison `Workbook.getWorksheet` makes.
 */
export function repairSheetName(name: string, taken: ReadonlySet<string>): string {
  const stripped = trimApostrophes(name.replace(INVALID_SHEET_NAME_CHARS_GLOBAL, '')).slice(
    0,
    MAX_SHEET_NAME_LENGTH,
  );
  const base = trimApostrophes(stripped);
  if (base.length === 0) return firstFree((n) => `Sheet${n}`, taken);
  if (!taken.has(base.toLowerCase())) return base;
  return firstFree((n) => withSuffix(base, ` (${n})`), taken, 2);
}

/**
 * A workbook's declared sheets in order, each carrying the name the model will accept: repaired by
 * {@link repairSheetName} against every name repaired before it.
 *
 * Both codecs declare their sheets as an ordered list and both have to know the repaired names before
 * anything cites a sheet by position, since a scoped defined name and a 3-D reference both index into
 * that order. The loop threading the taken set through the list is one loop, so it is written once.
 */
export function repairedSheetNames<T extends {readonly name: string}>(declared: readonly T[]): T[] {
  const taken = new Set<string>();
  return declared.map((entry) => {
    const name = repairSheetName(entry.name, taken);
    taken.add(name.toLowerCase());
    return {...entry, name};
  });
}

/** An apostrophe at either edge cannot be told from the quoting of a sheet-qualified reference. */
function trimApostrophes(name: string): string {
  return name.replace(/^'+/, '').replace(/'+$/, '');
}

/** The first `n` for which `candidate(n)` is not already taken, and the name it produced. */
function firstFree(candidate: (n: number) => string, taken: ReadonlySet<string>, from = 1): string {
  for (let n = from; ; n++) {
    const name = candidate(n);
    if (!taken.has(name.toLowerCase())) return name;
  }
}

/** `base` with `suffix` appended, trimming `base` so the whole still fits the length ceiling. */
function withSuffix(base: string, suffix: string): string {
  const room = MAX_SHEET_NAME_LENGTH - suffix.length;
  return trimApostrophes(base.slice(0, room)) + suffix;
}

/**
 * Read a sheet whose array formulas cannot stand as Excel repairs it (`core/array-formula-ranges.ts`). A
 * file is free to carry an array formula whose range holds another formula, or shares a cell with another
 * array formula's range, and the writer refuses both, since Excel offers to repair such a package.
 *
 * Excel 16.0 keeps the array formula decided first. A formula inside its range loses its formula and its
 * value, keeping its style, and a shared formula whose master that was loses its formula too, keeping its
 * value; an array formula whose range only shares cells with it loses its formula, keeping its value. The
 * value Excel then shows in a cell it emptied is what the kept array formula computes there, which the
 * library does not compute, so the cell is left empty.
 */
export function admitArrayRanges(sheet: Worksheet): void {
  const placements = [];
  for (const {cells} of sheet.rows()) {
    for (const cell of cells) {
      const placement = formulaPlacement(cell.address, cell.col, cell.row, cell.value);
      if (placement !== undefined) placements.push(placement);
    }
  }
  const {cellsRemoved, formulasRemoved} = arrayRangeRepair(placements);
  if (cellsRemoved.length === 0 && formulasRemoved.length === 0) return;
  const mastersRemoved = new Set<string>();
  for (const {address} of cellsRemoved) {
    const cell = sheet.getCell(address);
    if (isFormulaValue(cell.value)) mastersRemoved.add(address);
    cell.value = null;
  }
  for (const {address} of formulasRemoved) {
    const cell = sheet.getCell(address);
    cell.value = cachedResult(cell.value);
  }
  if (mastersRemoved.size === 0) return;
  for (const {cells} of sheet.rows()) {
    for (const cell of cells) {
      const value = cell.value;
      if (isSharedFormulaValue(value) && mastersRemoved.has(value.sharedFormula)) {
        cell.value = cachedResult(value);
      }
    }
  }
}

// The value a formula cell cached, standing in for the formula a repair took; empty when it cached none.
function cachedResult(value: CellValue): CellValue {
  return typeof value === 'object' &&
    value !== null &&
    'result' in value &&
    value.result !== undefined
    ? value.result
    : null;
}

/**
 * Run a model call that is being handed foreign input, and answer `undefined` where the model
 * refuses it.
 *
 * The three caught types are the three a model method raises *about its argument*: an
 * `AuthoringError` for a composite that cannot exist, and the native `RangeError` / `SyntaxError`
 * the taxonomy assigns to a single scalar that is out of range or does not parse. Every one of them
 * is a statement about the caller, and on this path the caller is a file.
 *
 * Nothing else is caught. An `XlsxError` raised by a layer below is a real failure of the read and
 * keeps its identity; an `InternalError` is a bug of ours and must not be swallowed by a reader
 * being tolerant about somebody else's file.
 */
export function admitting<T>(build: () => T): T | undefined {
  try {
    return build();
  } catch (error) {
    if (
      error instanceof AuthoringError ||
      error instanceof RangeError ||
      error instanceof SyntaxError
    ) {
      return undefined;
    }
    throw error;
  }
}
