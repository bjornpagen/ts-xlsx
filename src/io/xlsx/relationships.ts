// The relationship-type URIs and namespace groupings the writer's serialisers reference. The
// canonical namespace set lives in `namespaces.ts`; the writer-local groupings here (the props
// vocabularies, the content-types namespace, and the rel-type URIs derived from them) are plumbing
// every part of the writer references.
//
// The `.rels` part envelope these types are written into is container-level, not SpreadsheetML, and
// lives in `../opc/rels.ts`.

import {CONTENT_TYPES_NS, PKG_RELS_NS, RELATIONSHIPS_NS} from '../opc/namespaces.ts';
import {SPREADSHEETML_NS} from './namespaces.ts';

export const NS = {
  contentTypes: CONTENT_TYPES_NS,
  packageRels: PKG_RELS_NS,
  main: SPREADSHEETML_NS,
  docRels: RELATIONSHIPS_NS,
  coreProps: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
  extProps: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
  dc: 'http://purl.org/dc/elements/1.1/',
  dcterms: 'http://purl.org/dc/terms/',
  dcmitype: 'http://purl.org/dc/dcmitype/',
  xsi: 'http://www.w3.org/2001/XMLSchema-instance',
} as const;

// The Office extension relationships namespace the 2018 threaded-comment feature wires both of its parts
// through. Local to this table: nothing outside it needs the base URI.
const MS_OFFICE_2017_RELS_NS = 'http://schemas.microsoft.com/office/2017/10/relationships';

// The Office extension relationships namespace the workbook reaches its rich values through.
const RICH_DATA_RELS_NS = 'http://schemas.microsoft.com/office/2017/06/relationships';

export const REL = {
  worksheet: `${NS.docRels}/worksheet`,
  styles: `${NS.docRels}/styles`,
  theme: `${NS.docRels}/theme`,
  officeDocument: `${NS.docRels}/officeDocument`,
  coreProps: `${NS.packageRels}/metadata/core-properties`,
  extProps: `${NS.docRels}/extended-properties`,
  table: `${NS.docRels}/table`,
  comments: `${NS.docRels}/comments`,
  vmlDrawing: `${NS.docRels}/vmlDrawing`,
  drawing: `${NS.docRels}/drawing`,
  printerSettings: `${NS.docRels}/printerSettings`,
  image: `${NS.docRels}/image`,
  hyperlink: `${NS.docRels}/hyperlink`,
  sharedStrings: `${NS.docRels}/sharedStrings`,
  pivotTable: `${NS.docRels}/pivotTable`,
  pivotCacheDefinition: `${NS.docRels}/pivotCacheDefinition`,
  pivotCacheRecords: `${NS.docRels}/pivotCacheRecords`,
  sheetMetadata: `${NS.docRels}/sheetMetadata`,
  rdRichValue: `${RICH_DATA_RELS_NS}/rdRichValue`,
  rdRichValueStructure: `${RICH_DATA_RELS_NS}/rdRichValueStructure`,
  // Threaded comments are a Microsoft extension, so both types live under the 2017/10 Office
  // relationships namespace rather than the standard officeDocument one.
  threadedComment: `${MS_OFFICE_2017_RELS_NS}/threadedComment`,
  person: `${MS_OFFICE_2017_RELS_NS}/person`,
} as const;
