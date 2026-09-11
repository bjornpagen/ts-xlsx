// The buffered `.xlsb` reader: a binary OPC package in, a {@link Workbook} model out.
//
// An `.xlsb` is the *same* OPC/ZIP container and the *same* relationship graph as an `.xlsx`: the
// workbook part points at its sheets, its shared strings, and its style sheet through ordinary
// `.rels` XML. Only the office-document parts differ: `xl/workbook.bin`, `xl/worksheets/sheetN.bin`
// and friends are BIFF12 record streams instead of XML. So this module reuses the container layer
// wholesale (`../opc/`, the bounded inflater) and swaps only the part parsers, which is
// exactly the "two codecs over one model" shape the format note argues for.
//
// The model produced is the one `readXlsx` produces, not a parallel one: the same `Workbook`, the
// same `XfStyle` table, the same cells. That is what lets a caller convert between the two forms,
// and what the corpus asserts, by reading a workbook Excel saved in both forms and comparing.
//
// Not yet decoded (each its own slice of work, none silently wrong): rich-text runs, tables, pivots,
// and conditional formatting.

import type {DateEpoch} from '../../core/date.ts';
import {definedNameKeys, unmangleFunctions} from '../../core/formula.ts';
import {INTERNAL} from '../../core/internal.ts';
import {type DefinedName, Workbook} from '../../core/workbook.ts';
import type {WorksheetState} from '../../core/worksheet.ts';
import {quoted} from '../../errors.ts';
import {UnsupportedFormatError} from '../opc/errors.ts';
import {openSpreadsheetPackage, packageAccessors, readPartRelationships} from '../opc/read-opc.ts';
import type {ReadPackageOptions} from '../opc/read-options.ts';
import {admitting, repairedSheetNames} from '../read-policy/read-repair.ts';
import {XlsbParseError} from './errors.ts';
import {decodeFormula, type ExternSheetRef, type FormulaScope} from './formula.ts';
import {RecordReader} from './primitives.ts';
import {parseSharedStrings} from './read-shared-strings.ts';
import {parseStyleTable} from './read-styles.ts';
import {parseWorksheet} from './read-worksheet.ts';
import {blockTracker, readRecords} from './record-stream.ts';
import {BRT} from './record-types.ts';

/** The office-document part every `.xlsb` package is entered through. */
export const XLSB_WORKBOOK_PART = 'xl/workbook.bin';

/**
 * Read an `.xlsb` (binary BIFF12) package into a {@link Workbook}.
 *
 * @throws {UnsupportedFormatError} if the input is not an `.xlsb` package: a legacy `.xls`
 *   (`.format === 'xls'`), an XML `.xlsx` or unrecognised blob (`'unknown'`).
 * @throws {XlsbParseError} if a binary part is malformed.
 * @throws {PackageReadError} if the input is a ZIP that cannot be unpacked: a corrupt or
 *   truncated archive, or one exceeding the inflate bound (a probable zip bomb).
 */
export function readXlsb(data: Uint8Array, options: ReadPackageOptions = {}): Workbook {
  const {files, documentPath, workbookXml} = openSpreadsheetPackage(
    data,
    options.maxUncompressedBytes,
  );
  // An XML office document is a package this library reads, through the other codec. Said here rather
  // than left to the record reader, which would report the first byte of `<workbook` as a malformed
  // BIFF12 record: the format is what is wrong, not the bytes.
  if (workbookXml !== undefined) {
    throw new UnsupportedFormatError(
      'unknown',
      `not a valid .xlsb package: its office document ${quoted(documentPath)} is XML, not the binary ${XLSB_WORKBOOK_PART} a .xlsb carries; read it with readXlsx`,
    );
  }
  return readXlsbPackage(files, documentPath);
}

/**
 * Build the model from an already-inflated `.xlsb` package. Separate from {@link readXlsb} so the
 * `.xlsx` reader can hand over a package it has already inflated and classified, rather than
 * inflating the same bytes twice.
 *
 * `documentPath` is where the package's own `_rels/.rels` says its workbook lives; it defaults to the
 * conventional path for a caller holding nothing but the parts.
 */
export function readXlsbPackage(
  files: Record<string, Uint8Array>,
  documentPath: string = XLSB_WORKBOOK_PART,
): Workbook {
  const {partText, partBytes} = packageAccessors(files);
  const workbookPart = partBytes(documentPath);
  if (workbookPart === undefined) {
    throw new UnsupportedFormatError(
      'unknown',
      `not a valid .xlsb package: ${quoted(documentPath)} is missing`,
    );
  }

  // The pool and the stylesheet are reached through the workbook's relationships, with the
  // conventional paths as the fallback: the same resolution the XML reader does, for the same reason.
  // A package is free to name these parts anything its relationship graph points at.
  const rels = readPartRelationships(documentPath, partText, partBytes);
  const sharedStrings = parseSharedStrings(
    rels.relatedBytes('sharedStrings') ?? partBytes('xl/sharedStrings.bin'),
  );
  const {cellXfs, namedStyles, defaultFont} = parseStyleTable(
    rels.relatedBytes('styles') ?? partBytes('xl/styles.bin'),
  );

  const workbook = new Workbook();
  // As in the XML reader, the named-style layer is restored only when a file declares more than the
  // Normal default, so an ordinary workbook keeps an empty table and writes just that default back.
  if (namedStyles.length > 1) workbook[INTERNAL].restoreNamedStyles(namedStyles);
  // As in the XML reader, font 0 is the workbook's declared default and must survive a re-write; an
  // assumed Calibri in its place changes every empty cell and every character-unit column width.
  workbook[INTERNAL].restoreDefaultFont(defaultFont);

  const declaration = readWorkbookPart(workbookPart);
  workbook.dateEpoch = declaration.dateEpoch;
  // Repaired before anything cites them, exactly as the XML reader does: `addWorksheet` refuses a
  // duplicate, over-long or forbidden name that a file is free to carry, and a 3-D reference or a
  // scoped defined name resolves its sheet by position, so the scope has to hold the names the model
  // actually ended up with.
  const sheets = repairedSheetNames(declaration.sheets);
  const scope: FormulaScope = {
    sheetNames: sheets.map((sheet) => sheet.name),
    externSheets: declaration.externSheets,
    selfSupBook: declaration.selfSupBook,
    names: declaration.names.map((name) => name.name),
  };
  // Whether a function a formula passes as a value sheds its `_xleta.` depends on the names the
  // workbook defines, so they are known before any formula is decoded, from the records alone.
  const namesInWorkbook = definedNameKeys(declaration.names.filter(isWorkbookName));

  for (const declared of sheets) {
    const sheet = workbook.addWorksheet(declared.name, {state: declared.state});
    const target = declared.relId === undefined ? undefined : rels.byId(declared.relId)?.target;
    const part = target === undefined ? undefined : partBytes(rels.pathOf(target));
    if (part !== undefined)
      parseWorksheet(part, {
        sheet,
        sharedStrings,
        xfStyles: cellXfs,
        scope,
        dateEpoch: declaration.dateEpoch,
        definedNames: namesInWorkbook,
      });
  }
  // A name the model refuses, such as an empty one, is a name the file does not really carry.
  for (const defined of definedNames(declaration, scope, namesInWorkbook)) {
    admitting(() => {
      workbook.defineName(defined);
    });
  }
  return workbook;
}

// One sheet as `xl/workbook.bin` declares it, in workbook (tab) order.
//
// The xlsx codec models the same thing as `SheetEntry` in `read-workbook-xml.ts`, with the same
// three fields differing only in how they spell optionality. Kept apart on purpose: the codecs are
// peers, neither imports the other, and a shared type would buy one interface at the cost of a
// dependency edge that does not otherwise exist. A sheet declaration *is* a workbook-part concept,
// so if a third codec ever needs one, hoist it to `opc` rather than making one of these two import
// the other.
interface SheetDeclaration {
  readonly name: string;
  readonly relId: string | undefined;
  readonly state: WorksheetState['state'];
}

// One `BrtName`, still in its on-disk form: the target is a token stream that cannot be decoded until
// every sheet is known, so the record is gathered first and resolved after the pass.
interface NameDeclaration {
  readonly name: string;
  /** Zero-based index of the sheet the name is scoped to, or `undefined` for a workbook-global name. */
  readonly scopeSheet: number | undefined;
  /** Whether the name registers a callable rather than a target. See {@link definedNames}. */
  readonly isFunction: boolean;
  readonly rgce: Uint8Array;
  readonly rgcb: Uint8Array;
}

interface WorkbookDeclaration {
  readonly sheets: readonly SheetDeclaration[];
  readonly names: readonly NameDeclaration[];
  readonly externSheets: readonly ExternSheetRef[];
  readonly selfSupBook: number | undefined;
  /** The binary spelling of `<workbookPr date1904>`: what every serial in the book counts from. */
  readonly dateEpoch: DateEpoch;
}

// The Begin/End blocks the workbook part carries: the sheet bundle, and the externals block a 3-D
// reference resolves through. Siblings rather than nested, and tracked independently, so a damaged
// file that leaves one open cannot close the other.
type WorkbookBlock = 'bundle' | 'externals';

const WORKBOOK_BLOCKS: readonly (readonly [number, number, WorkbookBlock])[] = [
  [BRT.BeginBundleShs, BRT.EndBundleShs, 'bundle'],
  [BRT.BeginExternals, BRT.EndExternals, 'externals'],
];

// One pass over `xl/workbook.bin`, gathering everything the rest of the read depends on: the sheet
// bundle, the externals block a 3-D reference resolves through, and the defined names.
function readWorkbookPart(part: Uint8Array): WorkbookDeclaration {
  const sheets: SheetDeclaration[] = [];
  const names: NameDeclaration[] = [];
  let externSheets: readonly ExternSheetRef[] = [];
  // The two Begin/End blocks this pass reads, tracked by the same helper the style reader uses. It
  // used to be four booleans woven into the chain below, where the *order* of the arms was
  // load-bearing and unstated: moving `EndExternals` under the catch-all that counts records inside
  // the externals block would have silently miscounted the supporting books. Asking "is this a block
  // boundary" before asking what the record means removes that constraint instead of documenting it.
  const blocks = blockTracker(WORKBOOK_BLOCKS);
  // A workbook with no external links declares exactly one supporting book: itself. Rather than
  // enumerate every record type that could open another, and risk miscounting into a *wrong* sheet
  // name, anything else inside the externals block disqualifies the whole table.
  let supportingBooks = 0;
  let selfSupBook: number | undefined;
  let dateEpoch: DateEpoch = 1900;

  for (const record of readRecords(part)) {
    if (blocks.boundary(record.type)) continue;
    if (record.type === BRT.WbProp) dateEpoch = readDateEpoch(record.data);
    else if (record.type === BRT.BundleSh && blocks.isOpen('bundle'))
      sheets.push(readSheet(record.data));
    else if (record.type === BRT.ExternSheet) externSheets = readExternSheets(record.data);
    else if (record.type === BRT.SupSelf) selfSupBook = supportingBooks++;
    // Still ordered on purpose, and now only where the semantics require it: the two arms above name
    // records the externals block itself carries, and everything else inside it is a supporting book.
    else if (blocks.isOpen('externals')) supportingBooks++;
    else if (record.type === BRT.Name) names.push(readName(record.data));
  }
  return {
    sheets,
    names,
    externSheets,
    selfSupBook: supportingBooks === 1 ? selfSupBook : undefined,
    dateEpoch,
  };
}

// `BrtWbProp` ([MS-XLSB] 2.4.823): a 4-byte bit field of workbook settings, of which this reader
// wants one bit. Established against Excel Desktop rather than read off a table: the same workbook
// saved as `.xlsb` with the 1904 date system on and off differs in this record's first byte alone,
// 0x21 against 0x20, which puts `f1904` in bit 0. A record too short to hold the field is a damaged
// one, and the Windows default is the reading that loses least.
function readDateEpoch(data: Uint8Array): DateEpoch {
  if (data.length < 4) return 1900;
  return (new RecordReader(data).u32() & 1) === 0 ? 1900 : 1904;
}

// `BrtBundleSh` ([MS-XLSB] 2.4.303): the binary spelling of `<sheet name state r:id/>`.
function readSheet(data: Uint8Array): SheetDeclaration {
  const reader = new RecordReader(data);
  const state = SHEET_STATES[reader.u32()] ?? 'visible';
  reader.skip(4); // iTabId: the sheet's stable id, which the model assigns itself.
  const relId = reader.nullableWideString();
  return {name: reader.wideString(), relId, state};
}

// `BrtExternSheet` ([MS-XLSB] 2.4.677): the `Xti` table every 3-D reference indexes into.
function readExternSheets(data: Uint8Array): ExternSheetRef[] {
  const reader = new RecordReader(data);
  const count = reader.u32();
  // Each entry is three 4-byte fields, so a count the record cannot hold is a lie about the record and
  // is refused, the way every other length in this reader is. It used to return an empty table
  // instead, which reads as "this workbook declares no external sheets": every 3-D reference in it
  // then decodes to nothing and every formula carrying one silently falls back to its cached value,
  // across the whole workbook, with nothing reported.
  if (count * XTI_BYTES > reader.remaining) {
    throw new XlsbParseError(
      `BIFF12 externSheet table declares ${count} entries needing ${count * XTI_BYTES} bytes but ` +
        `only ${reader.remaining} remain in the record`,
    );
  }
  const entries: ExternSheetRef[] = [];
  for (let index = 0; index < count; index++) {
    entries.push({supBook: reader.u32(), firstSheet: reader.i32(), lastSheet: reader.i32()});
  }
  return entries;
}

const XTI_BYTES = 12;

// `BrtName` ([MS-XLSB] 2.4.673). The target is left undecoded here: it is a token stream that may cite
// a sheet by index, and the sheet bundle is not necessarily complete at this point in the stream.
function readName(data: Uint8Array): NameDeclaration {
  const reader = new RecordReader(data);
  const flags = reader.u32();
  reader.skip(1); // chKey: the Alt-key shortcut a macro name can carry.
  const itab = reader.u32();
  const name = reader.wideString();
  const rgce = reader.bytes(reader.u32());
  const rgcb = reader.bytes(reader.u32());
  return {
    name,
    scopeSheet: itab === GLOBAL_NAME_SCOPE ? undefined : itab,
    isFunction: (flags & NAME_IS_FUNCTION) !== 0,
    rgce,
    rgcb,
  };
}

// `itab` for a workbook-global name; any other value is a zero-based sheet index.
const GLOBAL_NAME_SCOPE = 0xffffffff;
// `fFunc`: the name registers something callable rather than a range.
const NAME_IS_FUNCTION = 0x00000002;

// The workbook's defined names, as the model holds them.
//
// Only the names {@link isWorkbookName} admits are kept. A name whose target uses a token this reader
// cannot decode is dropped too, rather than surfaced with a target that is a guess.
function definedNames(
  declaration: WorkbookDeclaration,
  scope: FormulaScope,
  namesInWorkbook: ReadonlySet<string>,
): DefinedName[] {
  const names: DefinedName[] = [];
  for (const declared of declaration.names) {
    if (!isWorkbookName(declared)) continue;
    const refersTo = decodeFormula(declared.rgce, declared.rgcb, scope);
    if (refersTo === undefined) continue;
    const sheet =
      declared.scopeSheet === undefined ? undefined : scope.sheetNames[declared.scopeSheet];
    names.push({
      name: declared.name,
      ...(sheet === undefined ? {} : {scope: sheet}),
      // Stripped back to the readable form, the same normalisation the XML reader applies.
      refersTo: unmangleFunctions(refersTo, namesInWorkbook),
    });
  }
  return names;
}

// Whether a `BrtName` is one of the workbook's defined names, rather than a record Excel keeps for its
// own formulas. Three kinds are not, and the XML form persists none of them, so carrying them through
// would make the two readings of one workbook disagree:
//
// - a *function* name, Excel's registration of a callable (every post-2007 function gets one,
//   `_xlfn.TEXTJOIN` and friends), whose target is the placeholder `#NAME?`;
// - the name a `PtgName` cites for a LET or LAMBDA parameter (`_xlpm.x`) or a function passed as a
//   value (`_xleta.SUM`), which Excel 16.0 (build 20326) registers without the function flag and with
//   the same `#NAME?` target;
// - `_xlnm._FilterDatabase`, the built-in Excel derives from a sheet's autofilter, which the model
//   reconstructs from the autofilter itself.
function isWorkbookName(declared: NameDeclaration): boolean {
  return (
    !declared.isFunction &&
    !FORMULA_PLACEHOLDER_NAME.test(declared.name) &&
    declared.name !== FILTER_DATABASE_NAME
  );
}

const FORMULA_PLACEHOLDER_NAME = /^_xl(?:fn|pm|eta)\./;
const FILTER_DATABASE_NAME = '_xlnm._FilterDatabase';

// `hsState` ([MS-XLSB] 2.4.303), indexed by its stored value.
const SHEET_STATES: ReadonlyArray<WorksheetState['state']> = ['visible', 'hidden', 'veryHidden'];
