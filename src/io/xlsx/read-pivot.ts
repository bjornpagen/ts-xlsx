// Reading a pivot table back from its OOXML parts: the inverse of `pivot.ts`.
//
// A written pivot round-trips today by byte-preservation: the reader keeps the pivot parts verbatim
// and re-emits them. That keeps the file faithful, but leaves the pivot opaque to the model: a
// `.model` copy cannot carry a pivot it cannot see. This module reconstructs the *semantic* shape of
// a pivot (its source, field roles, value field, and aggregation) from the `pivotTableDefinition`
// and its `pivotCacheDefinition`, so a loaded pivot becomes inspectable data rather than an opaque
// blob.
//
// Read leniently: these parts describe a file that already exists, so a missing or unrecognised
// attribute yields a sensible default rather than a throw. The strict invariants belong on the
// authoring path (`core/pivot-table.ts`), not here. A hostile part therefore degrades to an
// incomplete model; it never crashes the reader.

import {
  isDeclarablePivotSourceKind,
  type ParsedPivotField,
  type ParsedPivotSource,
  type ParsedPivotTable,
  type PivotMetric,
  type PivotSourceKind,
  pivotMetricFromSubtotal,
} from '../../core/pivot-table.ts';
import {numInteger} from '../../xml/xml-attrs.ts';
import {parseXml} from '../../xml/xml-read.ts';
import {localName} from '../../xml/xml-scan.ts';
import {relAttr} from '../opc/namespaces.ts';

/** Reconstruct a pivot's semantic model from its two definition parts. The records part is not
 * consulted: the cache's field catalogue and the table's field roles fully describe the pivot's
 * shape, and the aggregated values are Excel's to compute on refresh. */
export function parsePivotTable(tableXml: string, cacheXml: string): ParsedPivotTable {
  const {fields, source} = parsePivotCacheDefinition(cacheXml);
  const def = parsePivotTableDefinition(tableXml);
  return {
    name: def.name,
    cacheId: def.cacheId,
    source,
    fields,
    rowFields: def.rowFields,
    columnFields: def.columnFields,
    valueField: def.valueField,
    valueFieldName: fields[def.valueField]?.name ?? '',
    valueCaption: def.valueCaption,
    metric: def.metric,
  };
}

/** The field catalogue and worksheet source from a `pivotCacheDefinition`. Fields are collected in
 * document order, the order a table's `fld`/`x` indices address them by. */
function parsePivotCacheDefinition(cacheXml: string): {
  fields: ParsedPivotField[];
  source: ParsedPivotSource;
} {
  const fields: ParsedPivotField[] = [];
  // A worksheet source is the assumed default until proven otherwise: it is what our writer emits and
  // the overwhelmingly common shape, and its `<worksheetSource>` child fills in the coordinates. A
  // `<cacheSource type>` we recognise overrides the kind; an unrecognised one degrades to `unknown`.
  let source: ParsedPivotSource = {kind: 'worksheet', sheet: '', ref: '', inAnotherWorkbook: false};
  parseXml(cacheXml, {
    onOpen(name, attrs, _selfClosing, scope) {
      const local = localName(name);
      if (local === 'cacheField' && attrs.name !== undefined) {
        fields.push({name: attrs.name});
      } else if (local === 'cacheSource') {
        source = {...source, kind: sourceKind(attrs.type)};
      } else if (local === 'worksheetSource') {
        source = {
          ...source,
          sheet: attrs.sheet ?? '',
          ref: attrs.ref ?? '',
          inAnotherWorkbook: relAttr(scope, attrs, 'id') !== undefined,
        };
      }
    },
  });
  return {fields, source};
}

/** The layout half of a pivot: its name, cache id, axis field roles, and the single value field.
 * `<field x>` appears identically inside `<rowFields>` and `<colFields>`, so the current container is
 * tracked to route each into the right axis. Only the first `<dataField>` is modeled (the authoring
 * model supports one value field) and any further ones are ignored rather than rejected. */
function parsePivotTableDefinition(tableXml: string): {
  name: string;
  cacheId: string;
  rowFields: number[];
  columnFields: number[];
  valueField: number;
  valueCaption: string;
  metric: PivotMetric;
} {
  let name = '';
  let cacheId = '';
  const rowFields: number[] = [];
  const columnFields: number[] = [];
  let valueField = -1;
  let valueCaption = '';
  let metric: PivotMetric = 'sum';
  let seenDataField = false;
  let axis: 'row' | 'col' | null = null;

  parseXml(tableXml, {
    onOpen(elementName, attrs) {
      switch (localName(elementName)) {
        case 'pivotTableDefinition':
          name = attrs.name ?? '';
          cacheId = attrs.cacheId ?? '';
          break;
        case 'rowFields':
          axis = 'row';
          break;
        case 'colFields':
          axis = 'col';
          break;
        case 'field': {
          if (axis === null) break;
          const index = toIndex(attrs.x);
          if (index >= 0) (axis === 'row' ? rowFields : columnFields).push(index);
          break;
        }
        case 'dataField': {
          if (seenDataField) break;
          seenDataField = true;
          valueField = toIndex(attrs.fld);
          valueCaption = attrs.name ?? '';
          metric = pivotMetricFromSubtotal(attrs.subtotal);
          break;
        }
      }
    },
    onClose(elementName) {
      const local = localName(elementName);
      if (local === 'rowFields' || local === 'colFields') axis = null;
    },
  });

  return {name, cacheId, rowFields, columnFields, valueField, valueCaption, metric};
}

/** Map a `<cacheSource type>` to a known kind. Absent reads as `worksheet` (the spec default and what
 * our writer emits); an unrecognised value reads as `unknown` rather than throwing, keeping the read
 * lenient while still telling a consumer the declared source is not one we model. */
function sourceKind(type: string | undefined): PivotSourceKind {
  if (type === undefined) return 'worksheet';
  return isDeclarablePivotSourceKind(type) ? type : 'unknown';
}

/** Parse a non-negative field index attribute, or -1 when it is absent or not a whole number, so a
 * hostile `x="../etc"` can never become a wild array index this way. */
function toIndex(value: string | undefined): number {
  return numInteger(value, 0) ?? -1;
}
