// Worksheet tables (OOXML `<table>` parts, `xl/tables/table{n}.xml`), both directions. The writer
// (`tableXml`) turns a `Table` into its part; the reader (`parseTable`) is its inverse, turning a
// stored part back into the `TableOptions` a worksheet re-registers.
//
// The part stores the table's *full* occupied range (`ref="A1:B3"`), whereas the model anchors at a
// single top-left cell plus a data-row count. The two are equivalent: the data-row count is the
// range height minus the header row (present unless `headerRowCount="0"`) and the totals row (present
// only when `totalsRowCount` is positive), so reconstructing one from the other is lossless.

import {encodeAddress, tryDecodeRange} from '../../core/address.ts';
import {
  isTotalsRowFunction,
  type Table,
  TABLE_STYLE_FLAGS,
  type TableColumn,
  type TableOptions,
  type TableStyleInfo,
  type TotalsRowFunction,
} from '../../core/table.ts';
import {numInteger} from '../../xml/xml-attrs.ts';
import {parseXml, TextCapture} from '../../xml/xml-read.ts';
import {boolTristate, localName} from '../../xml/xml-scan.ts';
import {boolAttr, checkedToken, escapeAttr, escapeText, XML_DECLARATION} from '../../xml/xml.ts';
import {NS} from './relationships.ts';

export function tableXml(table: Table, id: number): string {
  const name = escapeAttr(table.name);
  const displayName = escapeAttr(table.displayName);
  // headerRowCount defaults to 1 in OOXML, so only a headerless table needs it stated.
  const headerRowCount = table.headerRow ? '' : ' headerRowCount="0"';
  // A present totals row implies it is shown, so it only needs the count. Without a totals row the
  // model's tri-state totalsRowShown decides: emit the flag Excel recorded, or nothing when the
  // source omitted it. Injecting `totalsRowShown="0"` onto a table that lacked the attribute is
  // exactly the spurious change that makes Excel treat an otherwise-valid table as corrupt.
  const totals = table.totalsRow
    ? ' totalsRowCount="1"'
    : boolAttr('totalsRowShown', table.totalsRowShown);
  const autoFilter =
    table.autoFilterRef !== undefined ? `<autoFilter ref="${table.autoFilterRef}"/>` : '';
  const columns = table.columns.map((column, i) => tableColumnXml(column, i + 1)).join('');
  return (
    XML_DECLARATION +
    `<table xmlns="${NS.main}" id="${id}" name="${name}" displayName="${displayName}" ` +
    `ref="${table.range}"${headerRowCount}${totals}>` +
    autoFilter +
    `<tableColumns count="${table.columns.length}">${columns}</tableColumns>` +
    tableStyleInfoXml(table.style) +
    '</table>'
  );
}

// Excel's default table appearance, written for a table that carries no style of its own.
const DEFAULT_TABLE_STYLE =
  '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" ' +
  'showRowStripes="1" showColumnStripes="0"/>';

// Emit `<tableStyleInfo>` from the model's style, or the default when none was captured. Each
// attribute is written only when the model holds it, so a style read without (say) a `name`, or a
// part that omitted a banding flag, re-emits exactly as it arrived rather than gaining an attribute.
function tableStyleInfoXml(style: TableStyleInfo | undefined): string {
  if (style === undefined) return DEFAULT_TABLE_STYLE;
  let attrs = '';
  if (style.name !== undefined) attrs += ` name="${escapeAttr(style.name)}"`;
  attrs += TABLE_STYLE_FLAGS.map((flag) => boolAttr(flag, style[flag])).join('');
  return `<tableStyleInfo${attrs}/>`;
}

// A `<tableColumn>` as the reader accumulates it: its attributes, then a `<totalsRowFormula>` child.
type TableColumnDraft = {
  name: string;
  totalsRowLabel?: string;
  totalsRowFunction?: TotalsRowFunction;
  totalsRowFormula?: string;
};

function tableColumnXml(column: TableColumn, id: number): string {
  let attrs = `id="${id}" name="${escapeAttr(column.name)}"`;
  if (column.totalsRowLabel !== undefined) {
    attrs += ` totalsRowLabel="${escapeAttr(column.totalsRowLabel)}"`;
  }
  if (column.totalsRowFunction !== undefined) {
    attrs += ` totalsRowFunction="${checkedToken(column.totalsRowFunction, isTotalsRowFunction, 'totals row function')}"`;
  }
  // A `custom` total is carried by a `<totalsRowFormula>` child rather than a built-in function, so
  // the element is non-self-closing when one is present. The formula is stored without a leading `=`,
  // matching how Excel writes it.
  if (column.totalsRowFormula !== undefined) {
    return `<tableColumn ${attrs}><totalsRowFormula>${escapeText(column.totalsRowFormula)}</totalsRowFormula></tableColumn>`;
  }
  return `<tableColumn ${attrs}/>`;
}

/**
 * Parse a `<table>` part into the options that reconstruct it, or `undefined` when the XML is not a
 * usable table (no name, no ref, or no columns: Excel treats such a part as corrupt, so we drop it
 * rather than fabricate a degenerate table). Duplicate column names are not resolved here, since the
 * {@link Table} constructor disambiguates them, so authoring and loading share one implementation.
 */
export function parseTable(xml: string): TableOptions | undefined {
  let name: string | undefined;
  let displayName: string | undefined;
  let ref: string | undefined;
  let headerRowCount = 1; // OOXML default: a table carries a header row unless it says otherwise.
  let totalsRowCount = 0; // OOXML default: no totals row.
  let totalsRowShown: boolean | undefined; // Absent unless the part states the attribute.
  let style: TableStyleInfo | undefined; // Absent unless the part carries a `<tableStyleInfo>`.
  let hasAutoFilter = false; // Only present when the part carries an `<autoFilter>` element.
  const columns: TableColumnDraft[] = [];
  // The column a `<totalsRowFormula>` belongs to. Unset while inside a column that was skipped, so
  // its formula lands nowhere rather than on the column before it.
  let currentColumn: TableColumnDraft | undefined;

  // A `<totalsRowFormula>` is a text child of the current `<tableColumn>`, so it is captured across
  // open/text/close rather than from an attribute. `calculatedColumnFormula` is a sibling child of
  // the same type (CT_TableFormula), which is why the capture answers only for its own element.
  const totalsFormula = new TextCapture('totalsRowFormula');

  parseXml(xml, {
    onOpen(elementName, attrs, selfClosing) {
      switch (localName(elementName)) {
        case 'table':
          // OOXML makes `displayName` the required identifier and `name` an optional alias; the
          // model inverts the roles (`name` is the formula identifier, `displayName` the label),
          // so read each from its own attribute and fall back across the pair when one is absent.
          name = attrs.name ?? attrs.displayName;
          displayName = attrs.displayName ?? attrs.name;
          ref = attrs.ref;
          headerRowCount = numInteger(attrs.headerRowCount, 0) ?? headerRowCount;
          totalsRowCount = numInteger(attrs.totalsRowCount, 0) ?? totalsRowCount;
          // Capture the flag verbatim so it re-emits exactly (or, absent, stays absent) rather
          // than being normalised. An unrecognised token is dropped: read as present-and-true, it
          // was written back as `"1"`.
          totalsRowShown = boolTristate(attrs.totalsRowShown);
          break;
        case 'autoFilter':
          hasAutoFilter = true;
          break;
        case 'tableStyleInfo': {
          // Keep each attribute off the literal so an absent one stays absent (not `key: undefined`),
          // preserving the round-trip: the writer re-emits only the attributes we actually saw.
          const captured: {-readonly [K in keyof TableStyleInfo]: TableStyleInfo[K]} = {};
          if (attrs.name !== undefined) captured.name = attrs.name;
          for (const flag of TABLE_STYLE_FLAGS) {
            const value = boolTristate(attrs[flag]);
            if (value !== undefined) captured[flag] = value;
          }
          style = captured;
          break;
        }
        case 'tableColumn': {
          if (attrs.name === undefined) {
            currentColumn = undefined;
            break;
          }
          const column: TableColumnDraft = {name: attrs.name};
          currentColumn = column;
          if (attrs.totalsRowLabel !== undefined) column.totalsRowLabel = attrs.totalsRowLabel;
          // An unrecognised totalsRowFunction is dropped rather than trusted in verbatim: the token
          // is a closed OOXML enumeration, so a foreign value is malformed input, not a future Excel
          // addition to accommodate.
          if (
            attrs.totalsRowFunction !== undefined &&
            isTotalsRowFunction(attrs.totalsRowFunction)
          ) {
            column.totalsRowFunction = attrs.totalsRowFunction;
          }
          columns.push(column);
          break;
        }
        case 'totalsRowFormula':
          totalsFormula.open('totalsRowFormula', selfClosing);
          break;
      }
    },
    onText(text) {
      totalsFormula.text(text);
    },
    onClose(elementName) {
      const formula = totalsFormula.close(localName(elementName));
      if (formula === undefined) return;
      // Attach to the column currently being parsed. Excel writes the child only for
      // `totalsRowFunction="custom"`, so a formula on any other column is meaningless, but preserving
      // whatever the part carried keeps the round-trip faithful rather than second-guessing.
      if (currentColumn !== undefined) currentColumn.totalsRowFormula = formula;
    },
  });

  if (name === undefined || ref === undefined || columns.length === 0) return undefined;

  // The `ref` is the coordinate every other field is read relative to, so a table whose anchor is
  // unreadable is dropped whole rather than rebuilt around a guessed origin. Re-encoding it here is
  // also what keeps the `Table` constructor's authoring-facing throw out of the reader's path.
  const decoded = tryDecodeRange(ref);
  if (decoded === undefined) return undefined;
  const {top, left, bottom, right} = decoded;
  if (top === undefined || left === undefined || bottom === undefined || right === undefined) {
    return undefined;
  }
  // The model lays its columns out from the anchor, one per `<tableColumn>`, so the columns must fill
  // the ref's width exactly. A column skipped for want of a name used to shift every column after it
  // one place left, shrinking the range and handing the next column's totals formula to the previous
  // one; a part declaring more columns than its ref is wide could also run past the grid. Either way
  // the table is not the one the part describes, so it is dropped whole like one with an unreadable
  // anchor.
  if (columns.length !== right - left + 1) return undefined;

  const headerRow = headerRowCount !== 0;
  const totalsRow = totalsRowCount > 0;
  const dataRows = bottom - top + 1 - (headerRow ? 1 : 0) - (totalsRow ? 1 : 0);

  const options: TableOptions = {
    name,
    displayName: displayName ?? name,
    ref: encodeAddress(left, top),
    columns,
    rowCount: Math.max(0, dataRows),
    headerRow,
    totalsRow,
    // Reconstruct the autoFilter state explicitly from the part: a header table read without an
    // `<autoFilter>` must not have one fabricated on the next write.
    autoFilter: hasAutoFilter,
  };
  // Kept off the literal so an absent attribute stays absent (not `totalsRowShown: undefined`),
  // preserving the round-trip: a table that never stated the flag must not gain one.
  if (totalsRowShown !== undefined) options.totalsRowShown = totalsRowShown;
  if (style !== undefined) options.style = style;
  return options;
}
