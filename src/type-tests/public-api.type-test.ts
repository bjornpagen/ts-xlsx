// Type-level tests over the public barrel (`src/index.ts`). Each imported symbol
// that no longer exists breaks compilation here (an export-removal guard), and each
// `Expect<Equal<...>>` locks a contract the runtime tests cannot see: synchronous
// I/O, the optionality of an address's col/row, the membership of CellValue.

import type {
  AddImageOptions,
  AuthoringError,
  AutoFilter,
  CellAddress,
  CellValue,
  Comment,
  CommentThread,
  CsvParseError,
  CustomUiParseError,
  DataTableFormulaValue,
  DefinedName,
  decodeAddress,
  ErrorValue,
  FilterColumn,
  FilterCriteria,
  FormulaValue,
  Hyperlink,
  InternalError,
  isDataTableFormulaValue,
  isErrorValue,
  isFormulaValue,
  isRichTextValue,
  isSharedFormulaValue,
  GridRect,
  Mention,
  MentionRef,
  PackageReadError,
  PageBreak,
  PageSetup,
  Person,
  PrintOptions,
  Range,
  RichTextValue,
  readXlsx,
  SharedFormulaValue,
  SheetView,
  TableRegion,
  UnsupportedFormat,
  UnsupportedFormatError,
  VbaAuthorError,
  VbaParseError,
  Workbook,
  Worksheet,
  writeXlsx,
  XlsbParseError,
  XlsxError,
  XlsxErrorCode,
  XlsxParseError,
  XmlParseError,
} from '../index.ts';
import type {StreamedRow} from '../io/xlsx/write-stream.ts';
import type {Equal, Expect, Extends} from './expect.ts';

// decodeAddress yields the canonical CellAddress, whose col/row stay optional so a
// column-only (`$A`) or row-only (`$1`) reference is representable.
export type AddressContracts = [
  Expect<Equal<ReturnType<typeof decodeAddress>, CellAddress>>,
  Expect<Equal<CellAddress['address'], string>>,
  Expect<Equal<CellAddress['col'], number | undefined>>,
  Expect<Equal<CellAddress['row'], number | undefined>>,
];

// CellValue admits the primitive/Date leaves and never `undefined`: an absent cell
// is `null`, not `undefined`, and the writer relies on that distinction.
export type ValueContracts = [
  Expect<Extends<null, CellValue>>,
  Expect<Extends<number, CellValue>>,
  Expect<Extends<string, CellValue>>,
  Expect<Extends<boolean, CellValue>>,
  Expect<Extends<Date, CellValue>>,
  Expect<Equal<Extends<undefined, CellValue>, false>>,
];

// Each object-shaped CellValue kind has a guard on the public barrel, and the whole signature is
// the contract: it accepts any CellValue (so it can be the *first* question asked about an unknown
// cell), and it narrows to exactly its own member of the union. Writing the predicate out is what
// pins the narrowing target; `ReturnType` would only ever say `boolean`. These are the discipline
// that keeps a consumer from hand-rolling `'richText' in value`, which narrows nothing useful.
export type ValueGuardContracts = [
  Expect<Equal<typeof isErrorValue, (value: CellValue) => value is ErrorValue>>,
  Expect<Equal<typeof isFormulaValue, (value: CellValue) => value is FormulaValue>>,
  Expect<Equal<typeof isSharedFormulaValue, (value: CellValue) => value is SharedFormulaValue>>,
  Expect<
    Equal<typeof isDataTableFormulaValue, (value: CellValue) => value is DataTableFormulaValue>
  >,
  Expect<Equal<typeof isRichTextValue, (value: CellValue) => value is RichTextValue>>,
];

// A hyperlink is sheet state beside the grid, not a kind of cell value, so no CellValue shape carries
// one: a value that can hold a link cannot also be the number or date the link sits on.
export type HyperlinkContracts = [
  Expect<Equal<Extract<CellValue, {readonly hyperlink: unknown}>, never>>,
  Expect<
    Equal<Hyperlink, {readonly ref: string; readonly target: string; readonly tooltip?: string}>
  >,
];

// The buffered I/O surface is synchronous: writeXlsx returns bytes and readXlsx a
// Workbook directly, never a Promise. getWorksheet is partial (a miss is undefined).
export type IoContracts = [
  Expect<Equal<ReturnType<typeof writeXlsx>, Uint8Array>>,
  Expect<Equal<ReturnType<typeof readXlsx>, Workbook>>,
  Expect<Equal<ReturnType<Workbook['getWorksheet']>, Worksheet | undefined>>,
  // ...and requireWorksheet is the total counterpart: it is only worth having if the caller is
  // spared the narrowing, so a `Worksheet | undefined` here would defeat the whole point of it.
  Expect<Equal<ReturnType<Workbook['requireWorksheet']>, Worksheet>>,
  Expect<Equal<Parameters<Workbook['requireWorksheet']>, Parameters<Workbook['getWorksheet']>>>,
];

// The used range is partial by design: a sheet spanning no rectangle says so, rather than handing
// back a phantom A1:A1 that a caller would go on to style, filter or iterate.
export type UsedRangeContracts = [Expect<Equal<Worksheet['usedRange'], Range | undefined>>];

// Every range-shaped thing in the library is the same four inclusive 1-based bounds, so a caller who
// writes one function over `GridRect` can pass it a range handle, a table's extent, or a merge.
export type GridRectContracts = [
  Expect<Extends<Range, GridRect>>,
  Expect<Equal<TableRegion, GridRect>>,
];

// A streamed row comes only from its sheet writer, which alone knows which row numbers are still
// open. A constructible one could commit a number the writer never handed out, evicting that row from
// the model with nothing written for it, so a private constructor is the contract, not a style.
export type StreamingConstructionContracts = [
  Expect<Equal<Extends<typeof StreamedRow, new (...args: never) => unknown>, false>>,
];

// Export-presence guards for the core feature types now on the barrel: importing each locks it into
// the public surface (its removal would break this compilation), and a self-`Extends` references it.
// FilterColumn/FilterCriteria come along so a constructed AutoFilter is fully nameable by callers.
export type FeatureSurface = [
  Expect<Extends<AutoFilter, AutoFilter>>,
  Expect<Extends<FilterColumn, FilterColumn>>,
  Expect<Extends<FilterCriteria, FilterCriteria>>,
  Expect<Extends<PageSetup, PageSetup>>,
  Expect<Extends<PrintOptions, PrintOptions>>,
  Expect<Extends<PageBreak, PageBreak>>,
  Expect<Extends<SheetView, SheetView>>,
  Expect<Extends<DefinedName, DefinedName>>,
  Expect<Extends<AddImageOptions, AddImageOptions>>,
];

// Threaded comments read back as a fully-resolved, immutable tree: a thread's messages and a message's
// mentions are readonly arrays (an inspection view, not an authoring handle), `resolved` is the
// thread's own boolean, and an identity lookup is partial: a message may name a person the registry
// does not hold, which is why `Comment.author`/`Mention.person` stay optional.
export type CommentThreadContracts = [
  Expect<Equal<CommentThread['comments'], readonly Comment[]>>,
  Expect<Equal<CommentThread['resolved'], boolean>>,
  Expect<Equal<Comment['mentions'], readonly Mention[]>>,
  Expect<Equal<Comment['author'], Person | undefined>>,
  Expect<Equal<Mention['person'], Person | undefined>>,
  // The resolved mention is the wire shape plus that one field: a caller can hand a `Mention` to
  // anything taking a `MentionRef`, and the codec's own parsed shape is the same type.
  Expect<Extends<Mention, MentionRef>>,
  Expect<Equal<Omit<Mention, 'person'>, MentionRef>>,
  Expect<Equal<ReturnType<Workbook['getPerson']>, Person | undefined>>,
  Expect<Equal<Workbook['persons'], readonly Person[]>>,
  Expect<Equal<ReturnType<Worksheet['commentThreadAt']>, CommentThread | undefined>>,
  Expect<Equal<Worksheet['commentThreads'], readonly CommentThread[]>>,
  // Authoring takes the model's own shapes: no wire-level surrogate a caller has to translate into.
  Expect<Equal<Parameters<Worksheet['addCommentThread']>, [CommentThread]>>,
  Expect<Equal<Parameters<Workbook['addPerson']>, [Person]>>,
];

// The failure taxonomy is a discriminated union over `code`, not a bag of unrelated classes: each
// subclass pins `code` to a literal, so narrowing an `XlsxError` on it narrows the *type*, and a
// class that widened its `code` back to `XlsxErrorCode` would break these rather than silently make
// every branch reachable. The `Extends` rows are the ancestry contract: one `catch` clause answers
// "was that us?" for every class the barrel exports.
export type ErrorTaxonomyContracts = [
  Expect<Equal<AuthoringError['code'], 'authoring'>>,
  Expect<Equal<InternalError['code'], 'internal'>>,
  Expect<Equal<UnsupportedFormatError['code'], 'unsupported-format'>>,
  Expect<Equal<XmlParseError['code'], 'malformed-input'>>,
  Expect<Equal<XlsxParseError['code'], 'malformed-input'>>,
  Expect<Equal<XlsbParseError['code'], 'malformed-input'>>,
  Expect<Equal<CsvParseError['code'], 'malformed-input'>>,
  Expect<Equal<PackageReadError['code'], 'malformed-input'>>,
  Expect<Equal<CustomUiParseError['code'], 'malformed-input'>>,
  Expect<Equal<VbaParseError['code'], 'malformed-input'>>,
  Expect<Equal<VbaAuthorError['code'], 'authoring'>>,
  Expect<Equal<XlsxError['code'], XlsxErrorCode>>,
  Expect<Extends<AuthoringError, XlsxError>>,
  Expect<Extends<UnsupportedFormatError, XlsxError>>,
  Expect<Extends<XlsxError, Error>>,
  // `format` survives the move onto the base: it is the branch for *which* unsupported input,
  // where `code` is only the branch for what kind of failure.
  Expect<Equal<UnsupportedFormatError['format'], UnsupportedFormat>>,
];
