// The part names the `.xlsx` writer generates, and the two other forms each of them is written in.
//
// Every generated part is named three times over, in three modules that never see each other's
// output. `write.ts` emits it under its package-absolute path; `workbook-xml.ts` declares that path
// again, slash-prefixed, as a `<Override PartName>` in `[Content_Types].xml`; and the relationship
// ledger in `package-plan.ts` names it a third time as a relationship Target, spelled relative to
// whichever part carries the relationship. A package where those three disagree is not malformed - it
// is well-formed and wrong, which is the expensive kind: the schema validates, and Excel then reports
// a part it cannot find, or silently drops the feature the relationship was reaching.
//
// So the path is written once, here, and the other two forms are derived from it: the content-type
// name by the slash, the relationship Target by the same `relativePartPath` the preserved references
// already went through. Renumbering a part kind or moving a directory is then one edit.
//
// SpreadsheetML-specific, and therefore here rather than in `io/opc/part-paths.ts`: those are the
// rules of the container (an extension, a `_rels` location, the arithmetic between two paths), which
// hold for any OPC package. These are this format's conventions about where a spreadsheet keeps its
// worksheets. The reader consults none of them, because it follows relationships instead of
// predicting names, and that asymmetry is deliberate: a file Excel wrote may put its parts anywhere
// OPC allows, so a reader that guessed would be wrong on real files.

/** The office document part: the workbook every other part hangs off. */
export const WORKBOOK_PART = 'xl/workbook.xml';
export const STYLES_PART = 'xl/styles.xml';
export const SHARED_STRINGS_PART = 'xl/sharedStrings.xml';
/** Singular and unnumbered, unlike the per-sheet thread parts: one identity registry serves the
 * whole workbook's threaded comments. */
export const PERSONS_PART = 'xl/persons/person.xml';
/** The metadata a `<c cm>` and `<c vm>` index: what marks a formula a dynamic array, and what names
 * the error behind a `#VALUE!`. */
export const METADATA_PART = 'xl/metadata.xml';
/** The rich values a value-metadata block names, and the structures they are read against. */
export const RICH_VALUES_PART = 'xl/richData/rdrichvalue.xml';
export const RICH_VALUE_STRUCTURES_PART = 'xl/richData/rdrichvaluestructure.xml';
export const CORE_PROPS_PART = 'docProps/core.xml';
export const APP_PROPS_PART = 'docProps/app.xml';

/** `xl/worksheets/sheet{n}.xml`, numbered by the sheet's position in the workbook (1-based). */
export function worksheetPart(number: number): string {
  return `xl/worksheets/sheet${number}.xml`;
}

/** `xl/tables/table{n}.xml`, numbered across the whole workbook rather than per sheet. */
export function tablePart(number: number): string {
  return `xl/tables/table${number}.xml`;
}

/** `xl/media/image{n}.{ext}`: the one part name that carries the source file's own extension, since
 * the `<Default>` content type is registered per extension. */
export function mediaPart(number: number, extension: string): string {
  return `xl/media/image${number}.${extension}`;
}

/** `xl/drawings/drawing{n}.xml`: the DrawingML anchors for one sheet's images. */
export function drawingPart(number: number): string {
  return `xl/drawings/drawing${number}.xml`;
}

/** `xl/drawings/vmlDrawing{n}.vml`: the legacy shape a note's box is drawn as. Numbered with the
 * comments part it accompanies, not with the DrawingML drawings it shares a directory with. */
export function vmlDrawingPart(number: number): string {
  return `xl/drawings/vmlDrawing${number}.vml`;
}

/** `xl/comments{n}.xml`: directly under `xl/`, unlike every other numbered part. */
export function commentsPart(number: number): string {
  return `xl/comments${number}.xml`;
}

/** `xl/threadedComments/threadedComment{n}.xml`: one sheet's modern conversations. */
export function threadedCommentsPart(number: number): string {
  return `xl/threadedComments/threadedComment${number}.xml`;
}

/** `xl/printerSettings/printerSettings{n}.bin`: an opaque blob carried through verbatim. */
export function printerSettingsPart(number: number): string {
  return `xl/printerSettings/printerSettings${number}.bin`;
}

/** `xl/pivotTables/pivotTable{n}.xml`: the pivot as its host sheet reaches it. */
export function pivotTablePart(number: number): string {
  return `xl/pivotTables/pivotTable${number}.xml`;
}

/** `xl/pivotCache/pivotCacheDefinition{n}.xml`: the field catalogue and source reference. */
export function pivotCacheDefinitionPart(number: number): string {
  return `xl/pivotCache/pivotCacheDefinition${number}.xml`;
}

/** `xl/pivotCache/pivotCacheRecords{n}.xml`: the cached rows the definition points at. */
export function pivotCacheRecordsPart(number: number): string {
  return `xl/pivotCache/pivotCacheRecords${number}.xml`;
}
