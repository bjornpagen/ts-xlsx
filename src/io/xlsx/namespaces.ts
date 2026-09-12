/**
 * Canonical SpreadsheetML namespace URIs and well-known extension GUIDs.
 *
 * These are wire-format constants: Excel keys its parsing off the exact URI or
 * GUID, so a producer must reproduce each one byte-for-byte. Centralizing them
 * keeps the writer and reader from drifting apart and retires the
 * `NS_MAIN`/`MAIN_NS`/`main` naming fork that had grown across the io modules.
 *
 * The package-level URIs (`.rels`, content types, the relationship vocabulary) describe the
 * container rather than the spreadsheet inside it, and live in `../opc/namespaces.ts`.
 */

import type {NamespaceScope} from '../../xml/xml-namespaces.ts';

/**
 * SpreadsheetML main namespace: the default `xmlns` of the workbook,
 * worksheet, styles, sharedStrings, comments, table and pivot parts.
 */
export const SPREADSHEETML_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/**
 * Markup-compatibility namespace (`mc:`), whose `mc:Ignorable` attribute lists the prefixes a
 * consumer that does not know them may skip rather than reject the part over.
 */
export const MARKUP_COMPATIBILITY_NS =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';

/**
 * The 2014 revision namespace (`xr:`), which scopes the `xr:uid` Excel stamps on a comment. Declared
 * `mc:Ignorable` wherever it appears, so a consumer that ignores it still reads the part.
 */
export const REVISION_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision';

/**
 * The 2018 threaded-comments namespace, shared by both parts of the feature: a sheet's
 * `threadedComment{n}.xml` and the workbook's `person.xml`. Note the plural `threadedcomments`, all
 * lower-case: Excel matches the URI exactly and reads neither part under any other spelling.
 */
export const THREADED_COMMENTS_NS =
  'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';

/** DrawingML shared graphics namespace (`a:`). */
export const DRAWINGML_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** Spreadsheet-drawing anchor namespace (`xdr:`) used by the worksheet drawing part. */
export const XDR_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';

/**
 * The 2009 Microsoft extension namespace. `x14` scopes the feature elements
 * Excel tucks inside `<ext>` blocks (conditional formatting, data validation,
 * slicers); it is declared inline on those elements exactly as Excel writes
 * them, so a worksheet root never needs an extra namespace declaration.
 */
export const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';

/** Scopes the `<xm:sqref>`/`<xm:f>` references the x14 feature elements carry. */
export const XM_NS = 'http://schemas.microsoft.com/office/excel/2006/main';

/**
 * Well-known `<ext uri=…>` GUIDs. Each `<ext>` block is opaque to a consumer
 * that does not recognize its GUID, so a producer must emit these exact values
 * for Excel to rediscover the feature.
 */
export const CF_EXT_URI = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}';
export const DATABAR_LINK_EXT_URI = '{B025F937-C7B1-47D3-B67F-A62EFF666E3E}';
export const DATA_VALIDATION_EXT_URI = '{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}';
export const SLICER_LIST_EXT_URI = '{A8765BA9-456A-4dab-B4F3-ACF838C121DE}';
export const SLICER_CACHES_EXT_URI = '{BBE1A952-AA13-448e-AADC-164F8A28A991}';
/** The extension a dynamic-array formula's cell-metadata block carries its properties in. */
export const DYNAMIC_ARRAY_PROPERTIES_EXT_URI = '{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}';

/** The 2017 dynamic-array namespace (`xda:`), which scopes those properties in `xl/metadata.xml`. */
export const DYNAMIC_ARRAY_NS =
  'http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray';

/** The extension a rich-value metadata block names its rich value in (`<xlrd:rvb i>`). */
export const RICH_VALUE_BLOCK_EXT_URI = '{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}';

/** The 2017 rich-data namespace (`xlrd:`): the rich values, their structures, and the blocks naming them. */
export const RICH_DATA_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata';

/**
 * Whether an element belongs to the standard SpreadsheetML vocabulary rather than to an extension.
 *
 * This used to be spelled `!name.includes(':')`, on the assumption that only extension elements carry
 * a prefix. A worksheet may bind the *main* namespace to a prefix instead, which is legal and which
 * real toolchains emit, and then every element has a colon in it: such a file had every one of its
 * data validations and conditional formats read as an unknown extension and dropped, silently.
 *
 * An unprefixed element counts even when the part declares no default namespace. A part with no
 * `xmlns` at all is not conforming OOXML, but this scanner does not reject it, such parts were read
 * before, and a fix that gains files by resolving namespaces should not lose files that have none.
 */
export function isMainNamespaceElement(scope: NamespaceScope, name: string): boolean {
  const namespace = scope.elementNamespace(name);
  return namespace === SPREADSHEETML_NS || (namespace === undefined && !name.includes(':'));
}

/** The complement of {@link isMainNamespaceElement}: an element of one of the extension vocabularies. */
export function isExtensionElement(scope: NamespaceScope, name: string): boolean {
  return !isMainNamespaceElement(scope, name);
}
