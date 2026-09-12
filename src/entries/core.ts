// Entry barrel for `@shbernal/ts-xlsx/core`: the document model, with no serialisation attached.
//
// This is what a consumer imports to build or inspect a workbook in memory: the classes, the
// property shapes, and the value vocabulary. It pulls in no ZIP, no XML writer and no BIFF12
// decoder. Errors are not here: the whole failure taxonomy lives behind `/errors`, so that a
// caller who only needs to branch on a failure never loads a codec to get the class.
//
// **Every closed token union published here publishes its narrowing guard.** The unions are the
// spellings OOXML allows for an attribute, and a caller holding a `string` (from a config file, a
// CLI flag, another library's model) has to get from it to the union somehow; without the guard the
// only route is a cast, which is the one thing the union exists to prevent. The guards are derived
// from the same table the union is (`src/token-set.ts`), so they cannot narrow to the wrong one, and
// they are already inside this entry's closure: publishing them costs no bytes.
//
// Two of them used to be published and eighteen were not, with nothing stating a rule that admitted
// `isTableStyleElementType` and refused `isBorderStyle`. That was drift. The rule is the whole list
// or none of it, and this is the whole list.

export {
  type CellAddress,
  type CellPosition,
  columnToNumber,
  decodeAddress,
  decodeRange,
  encodeAddress,
  type GridRect,
  MAX_COLUMN,
  MAX_ROW,
  numberToColumn,
  type RangeAddress,
} from '../core/address.ts';
export {
  type AutoFilter,
  type CustomFilter,
  type CustomFilterOperator,
  type CustomFilterPredicate,
  type FilterColumn,
  type FilterCriteria,
  isCustomFilterOperator,
  type ValuesFilter,
} from '../core/autofilter.ts';
export {Cell} from '../core/cell.ts';
export {
  applyTint,
  type ColorResolutionContext,
  DEFAULT_INDEXED_COLORS,
  resolveColor,
  SYSTEM_INDEXED_COLORS,
} from '../core/color-resolution.ts';
export {AxisHandle} from '../core/axis-handle.ts';
export {Column} from '../core/column.ts';
export type {DateEpoch} from '../core/date.ts';
export type {Comment, CommentThread, Mention, MentionRef, Person} from '../core/comment-thread.ts';
export {
  type CfIcon,
  type CfIconSetType,
  type CfTimePeriod,
  type CfValueObject,
  type CfValueObjectType,
  type ConditionalFormatting,
  type ConditionalFormattingOperator,
  type ConditionalFormattingRule,
  type ConditionalFormattingType,
  type DataBarAxisPosition,
  type DataBarDirection,
  type IconSetType,
  isCfIconSetType,
  isCfTimePeriod,
  isCfValueObjectType,
  isConditionalFormattingOperator,
  isConditionalFormattingType,
  isDataBarAxisPosition,
  isDataBarDirection,
  isIconSetType,
} from '../core/conditional-formatting.ts';
export {type Hyperlink} from '../core/hyperlink.ts';
export {
  type DataValidation,
  type DataValidationEntry,
  type DataValidationErrorStyle,
  type DataValidationImeMode,
  type DataValidationOperator,
  type DataValidationType,
  isDataValidationErrorStyle,
  isDataValidationImeMode,
  isDataValidationOperator,
  isDataValidationType,
} from '../core/data-validation.ts';
// The descriptor a structural edit is: `Table.shiftRows`/`shiftColumns` name it, so a caller that
// re-pins a table itself needs to be able to write its type.
export type {AxisSplice} from '../core/grid-shift.ts';
export {
  type AnchoredImage,
  type AnchorPoint,
  type Extent,
  type ImageAnchor,
  type ImageCrop,
  type ImageEditAs,
  type ImageHyperlink,
  isImageEditAs,
  isOneCellAnchor,
  type OneCellAnchor,
  type PictureProperties,
  type PortableImage,
  PX_TO_EMU,
  type TwoCellAnchor,
  type WorkbookImage,
  type WorksheetImages,
} from '../core/image.ts';
// The OPC container's inflate bound, shared by every reader (`readXlsx`, `readXlsb`, the row
// streamer). It is here rather than on a codec subpath so a `/xlsb` or `/csv` consumer can name it
// without importing the XML codec, and because a container bound is a fact about the package rather
// than about either serialisation inside it.
export type {ReadPackageOptions} from '../io/opc/read-options.ts';
// The pixel-unit anchor `Worksheet.addImage` takes, so a caller building one ahead of the call can
// write its type.
export type {PixelAnchor} from '../core/worksheet-pictures.ts';
export {
  INVALID_SHEET_NAME_CHARS,
  MAX_COLUMN_WIDTH,
  MAX_ROW_HEIGHT,
  MAX_SHEET_NAME_LENGTH,
  MAX_TABLE_NAME_LENGTH,
  TABLE_NAME_PATTERN,
} from '../core/limits.ts';
export {
  type HeaderFooter,
  isPageOrder,
  isPageOrientation,
  type PageBreak,
  type PageMargins,
  type PageOrder,
  type PageOrientation,
  type PageSetup,
  type PrintOptions,
} from '../core/page-setup.ts';
export {
  type ParsedPivotField,
  type ParsedPivotSource,
  type ParsedPivotTable,
  type PivotCacheField,
  type PivotItem,
  type PivotMetric,
  type PivotNumericSummary,
  isDeclarablePivotSourceKind,
  type PivotRecordCell,
  type PivotSourceKind,
  PivotTable,
  type PivotTableOptions,
} from '../core/pivot-table.ts';
export type {
  PreservedPart,
  PreservedRelationship,
  PreservedRootReference,
  PreservedWorksheetReference,
} from '../core/preserved.ts';
export type {
  SheetProtection,
  SheetProtectionCredential,
  SheetProtectionFlags,
  SheetProtectionOptions,
} from '../core/protection.ts';
export {Range} from '../core/range.ts';
export {Row} from '../core/row.ts';
export {
  type Alignment,
  type Border,
  type BorderEdge,
  type BorderStyle,
  type CellContent,
  type CellStyle,
  type Color,
  type Fill,
  type FillPatternType,
  type Font,
  type FontScheme,
  type FontVerticalAlignment,
  type GradientFill,
  type GradientStop,
  type HorizontalAlignment,
  isBorderStyle,
  isFillPatternType,
  isFontScheme,
  isFontVerticalAlignment,
  isHorizontalAlignment,
  isNamedUnderlineStyle,
  isVerticalAlignment,
  type PatternFill,
  type Protection,
  type UnderlineStyle,
  type VerticalAlignment,
} from '../core/style.ts';
export {
  Table,
  type TableColumn,
  type TableColumnStyle,
  isTotalsRowFunction,
  type TableGrid,
  type TableOptions,
  type TableRegion,
  type TableStyleInfo,
  type TotalsRowFunction,
} from '../core/table.ts';
// The workbook style *tables*: shapes that compose `CellStyle` without being one, so they live beside
// the slice that owns them rather than beside `Fill`. Only the declaration moved; the public surface
// is unchanged.
export type {
  DifferentialStyle,
  NamedCellStyle,
  TableStyleDefinition,
  TableStyleNamespace,
  TableStyleTable,
} from '../core/workbook-styles.ts';
export {
  isTableStyleElementType,
  STRIPE_ELEMENT_TYPES,
  TABLE_STYLE_ELEMENT_TYPES,
  type TableStyle,
  type TableStyleElement,
  type TableStyleElementType,
} from '../core/table-style.ts';
export {estimateWrappedLines} from '../core/text-metrics.ts';
export {
  DEFAULT_THEME_COLOR_SCHEME,
  DEFAULT_THEME_FONTS,
  isThemeColorSlot,
  THEME_COLOR_SLOTS,
  type ThemeColorScheme,
  type ThemeColorSlot,
  type ThemeFontScheme,
  type ThemeOverrides,
} from '../core/theme.ts';
export {
  type CellValue,
  cellValueToText,
  coerceCellValue,
  type DataTableFormulaValue,
  detectValueType,
  ERROR_CODES,
  type ErrorCode,
  type ErrorValue,
  type FormulaResult,
  type FormulaValue,
  isDataTableFormulaValue,
  isErrorCode,
  isErrorValue,
  isFormulaValue,
  isRichTextValue,
  isSharedFormulaValue,
  type RichTextRun,
  type RichTextValue,
  richTextToPlain,
  type SharedFormulaValue,
  ValueType,
} from '../core/value.ts';
export {
  type AddImageOptions,
  type AddWorksheetOptions,
  DEFAULT_WORKBOOK_VIEW,
  type DefinedName,
  type PreservedTheme,
  type PreservedWorkbookReference,
  Workbook,
  type WorkbookProperties,
  type WorkbookView,
} from '../core/workbook.ts';
export type {
  WorkbookProtection,
  WorkbookProtectionCredentialAttr,
} from '../core/workbook-protection.ts';
export {
  type CellModel,
  type ColumnProperties,
  type OutlineProperties,
  type RowInput,
  type RowProperties,
  type SheetView,
  isVisibility,
  type Visibility,
  Worksheet,
  type WorksheetModel,
  type WorksheetProperties,
  type WorksheetState,
} from '../core/worksheet.ts';
