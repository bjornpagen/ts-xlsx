// OOXML rendering for a modeled pivot table: the three parts a pivot needs, the cache definition
// (field catalogue), the cache records (a copy of the source rows), and the pivot table definition
// (the layout). The semantic computation lives in `core/pivot-table.ts`; this file only serialises.
//
// Every string that reaches an attribute is run through `escapeAttr`, so source strings carrying XML
// specials (`Smith & Co`, `<West>`, `It's "best"`) become well-formed markup rather than corrupting
// the package, which is the whole point of the shared-item escaping this module guarantees. Every
// number reaching an attribute goes through `numberText`, which is the other half of that boundary:
// the source-derived values (a shared item, a cache field's min/max) can only be finite because
// `core/pivot-table.ts` blanks a non-finite cell first, and this makes the writer say so itself
// rather than lean on an invariant three modules away.

import {encodeAddress} from '../../core/address.ts';
import type {
  PivotCacheField,
  PivotItem,
  PivotMetric,
  PivotRecordCell,
  PivotTable,
} from '../../core/pivot-table.ts';
import {escapeAttr, numberText, XML_DECLARATION} from '../../xml/xml.ts';
import {RELATIONSHIPS_NS} from '../opc/namespaces.ts';
import {SPREADSHEETML_NS} from './namespaces.ts';

// Excel's default caption prefix for each aggregation ("Sum of Amount", "Average of Amount"). A
// metric's name is also its `subtotal` value, which is why the record key equals the enum member.
const METRIC_CAPTIONS: Record<PivotMetric, string> = {
  sum: 'Sum',
  count: 'Count',
  countNums: 'Count',
  average: 'Average',
  max: 'Max',
  min: 'Min',
  product: 'Product',
  stdDev: 'StdDev',
  stdDevp: 'StdDevp',
  var: 'Var',
  varp: 'Varp',
};

/** The `pivotCacheDefinition` part: the source reference and the field catalogue. `recordsRelId` is
 * the id this cache's own rels part gives the cache-records part, which `r:id` cites.
 * `refreshOnLoad="1"` tells Excel to rebuild the pivot body from the cache on open, so the layout stays
 * correct without us computing it to the pixel. `recordCount` is the number of source data rows. */
export function pivotCacheDefinitionXml(table: PivotTable, recordsRelId: string): string {
  const fields = table.cacheFields
    .map((field) => {
      const shared = field.sharedItems;
      const kinds = sharedItemsKinds(field);
      const descriptor =
        shared === null
          ? `<sharedItems${kinds}/>`
          : `<sharedItems${kinds} count="${shared.length}">${shared.map(sharedItemXml).join('')}</sharedItems>`;
      return `<cacheField name="${escapeAttr(field.name)}" numFmtId="0">${descriptor}</cacheField>`;
    })
    .join('');
  return (
    XML_DECLARATION +
    `<pivotCacheDefinition xmlns="${SPREADSHEETML_NS}" xmlns:r="${RELATIONSHIPS_NS}" r:id="${escapeAttr(recordsRelId)}" refreshOnLoad="1" ` +
    `refreshedBy="ts-xlsx" createdVersion="8" refreshedVersion="8" minRefreshableVersion="3" ` +
    `recordCount="${table.records.length}">` +
    `<cacheSource type="worksheet">` +
    `<worksheetSource ref="${escapeAttr(table.sourceRef)}" sheet="${escapeAttr(table.sourceSheetName)}"/>` +
    `</cacheSource>` +
    `<cacheFields count="${table.cacheFields.length}">${fields}</cacheFields>` +
    `</pivotCacheDefinition>`
  );
}

/** The `pivotCacheRecords` part: one `<r>` per source data row, each cell either an index into an
 * axis field's shared-items catalogue (`<x>`) or an inline value (`<n>`/`<s>`/`<m>`). */
export function pivotCacheRecordsXml(table: PivotTable): string {
  const rows = table.records
    .map((record) => `<r>${record.map(recordCellXml).join('')}</r>`)
    .join('');
  return (
    XML_DECLARATION +
    `<pivotCacheRecords xmlns="${SPREADSHEETML_NS}" xmlns:r="${RELATIONSHIPS_NS}" count="${table.records.length}">` +
    rows +
    `</pivotCacheRecords>`
  );
}

/** The `pivotTableDefinition` part placed on the destination sheet: the field layout that binds the
 * cache (by `cacheId`) to the row/column axes and the summed value field. */
export function pivotTableXml(table: PivotTable, name: string, cacheId: string): string {
  // `PivotTable`'s constructor refuses a pivot with no row or column field, so both are present. Read
  // through a default rather than asserted: an assertion that is right for a reason stated in another
  // module is one the compiler cannot check, and field 0 is the honest fallback if it ever were wrong.
  const rowField = table.rowFields[0] ?? 0;
  const columnField = table.columnFields[0] ?? 0;
  const rowGroups = table.cacheFields[rowField]?.sharedItems?.length ?? 1;
  const columnGroups = table.cacheFields[columnField]?.sharedItems?.length ?? 1;
  // A generous bounding box on the destination sheet: a row-label column plus one column per column
  // group plus a grand-total column; two header rows plus one row per row group plus a grand total.
  // Excel recomputes the exact extent from the cache on refresh, so this only has to be valid. That
  // holds for nested fields too: sized from the first row and column field alone, with grand-total-only
  // items and `firstDataRow="2"`, a pivot with two row fields and two column fields opens clean in
  // Excel 16.0, although Excel itself writes `firstDataRow="3"` and every item for that layout.
  const location = `A1:${encodeAddress(2 + columnGroups, 3 + rowGroups)}`;

  const pivotFields = table.cacheFields
    .map((field, index) => {
      // The value field may also sit on an axis, and then it carries both attributes: Excel repairs a
      // package whose `<dataField fld>` points at a `<pivotField>` not flagged as one.
      const dataField = index === table.valueField ? ' dataField="1"' : '';
      if (table.rowFields.includes(index) || table.columnFields.includes(index)) {
        const axis = table.rowFields.includes(index) ? 'axisRow' : 'axisCol';
        const items = field.sharedItems ?? [];
        const entries = `${items.map((_item, i) => `<item x="${i}"/>`).join('')}<item t="default"/>`;
        return `<pivotField axis="${axis}"${dataField} showAll="0"><items count="${items.length + 1}">${entries}</items></pivotField>`;
      }
      return `<pivotField${dataField} showAll="0"/>`;
    })
    .join('');

  const rowFields = table.rowFields.map((index) => `<field x="${index}"/>`).join('');
  const columnFields = table.columnFields.map((index) => `<field x="${index}"/>`).join('');

  return (
    XML_DECLARATION +
    `<pivotTableDefinition xmlns="${SPREADSHEETML_NS}" xmlns:r="${RELATIONSHIPS_NS}" name="${escapeAttr(name)}" ` +
    `cacheId="${escapeAttr(cacheId)}" applyNumberFormats="0" applyBorderFormats="0" ` +
    `applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" ` +
    `applyWidthHeightFormats="1" dataCaption="Values" updatedVersion="8" minRefreshableVersion="3" ` +
    `useAutoFormatting="1" itemPrintTitles="1" createdVersion="8" indent="0" outline="1" ` +
    `outlineData="1" multipleFieldFilters="0">` +
    `<location ref="${location}" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>` +
    `<pivotFields count="${table.cacheFields.length}">${pivotFields}</pivotFields>` +
    `<rowFields count="${table.rowFields.length}">${rowFields}</rowFields>` +
    `<rowItems count="1"><i t="grand"><x/></i></rowItems>` +
    `<colFields count="${table.columnFields.length}">${columnFields}</colFields>` +
    `<colItems count="1"><i t="grand"><x/></i></colItems>` +
    `<dataFields count="1">` +
    dataFieldXml(table) +
    `</dataFields>` +
    `<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" ` +
    `showRowStripes="0" showColStripes="0" showLastColumn="1"/>` +
    `</pivotTableDefinition>`
  );
}

/** The `<dataField>` that names the aggregated column and selects its function. `sum` is Excel's
 * implicit default, so its `subtotal` attribute is omitted; every other metric names itself. */
function dataFieldXml(table: PivotTable): string {
  const caption = `${METRIC_CAPTIONS[table.metric]} of ${table.valueFieldName}`;
  const subtotal = table.metric === 'sum' ? '' : ` subtotal="${table.metric}"`;
  return (
    `<dataField name="${escapeAttr(caption)}" fld="${table.valueField}"${subtotal} ` +
    `baseField="0" baseItem="0"/>`
  );
}

// Which kinds of value a cache field holds, spelled as Excel 16.0 spells it for every mix of strings,
// numbers and blanks, in schema order. Each flag defaults to what a field of strings would say, so only a
// departure is written: `containsString="0"` for a field with none, `containsSemiMixedTypes="0"` only when
// it holds nothing but numbers (a blank counts as the other type), `containsMixedTypes="1"` when strings
// and numbers share it, and `containsNumber` with the range and integrality whenever it holds a number.
// The same attributes describe a catalogue and a field stored inline; a catalogue of numbers that
// omitted them opened with Excel's repair prompt.
function sharedItemsKinds(field: PivotCacheField): string {
  const {containsString, containsBlank, numeric} = field;
  const hasNumber = numeric !== null;
  return (
    (!containsString && !containsBlank && hasNumber ? ' containsSemiMixedTypes="0"' : '') +
    (containsString ? '' : ' containsString="0"') +
    (containsBlank ? ' containsBlank="1"' : '') +
    (containsString && hasNumber ? ' containsMixedTypes="1"' : '') +
    (numeric === null
      ? ''
      : ' containsNumber="1"' +
        (numeric.allInteger ? ' containsInteger="1"' : '') +
        ` minValue="${numberText(numeric.min)}" maxValue="${numberText(numeric.max)}"`)
  );
}

function sharedItemXml(item: PivotItem): string {
  switch (item.kind) {
    case 'string':
      return `<s v="${escapeAttr(item.value)}"/>`;
    case 'number':
      return `<n v="${numberText(item.value)}"/>`;
    case 'blank':
      return '<m/>';
  }
}

function recordCellXml(cell: PivotRecordCell): string {
  if (cell.kind === 'index') return `<x v="${cell.index}"/>`;
  return sharedItemXml(cell);
}
