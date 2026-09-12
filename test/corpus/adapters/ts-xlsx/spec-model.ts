// The corpus spec vocabulary ↔ the rewrite's model.
//
// A case describes a workbook declaratively; `buildFrom` is the one place that turns that
// description into live objects, and the key sets below are its vocabulary: a spec reaching for a key
// nobody has wired fails loudly (see `UnsupportedSpecError`) rather than being quietly skipped.

import type {Untyped} from '../../untyped.ts';
import {
  decodeAddress,
  decodeRange,
  encodeAddress,
  Workbook,
  type WorksheetInstance,
} from './runtime.ts';

// Translate a corpus image range, a string like "B2:D6", or a {tl, br?/ext?, editAs?} object, into
// the model's typed addImage call. A one-cell anchor is a point plus a fixed pixel extent (editAs is a
// two-cell-only attribute the model drops by construction); a two-cell anchor spans tl..br. A
// fractional grid coordinate (col 3.5) is passed through: the model floors it to the cell and derives
// the sub-cell EMU offset from that cell's real width/height.
export function anchorSpecImage(sheet: Untyped, imageId: Untyped, range: Untyped) {
  if (typeof range === 'string') {
    const {left, top, right, bottom} = decodeRange(range);
    sheet.addImage(imageId, {tl: {col: left! - 1, row: top! - 1}, br: {col: right, row: bottom}});
    return;
  }
  const {tl, br, ext, editAs} = range;
  if (ext !== undefined) {
    sheet.addImage(imageId, {tl, ext: {width: ext.width, height: ext.height}});
  } else {
    sheet.addImage(imageId, editAs !== undefined ? {tl, br, editAs} : {tl, br});
  }
}

// The 1-based `row.values` array a full-load reader exposes, rebuilt from a streamed row's cells:
// index 0 is an empty leading slot and column A lands at index 1, so streaming and buffered reads
// index identically. Gaps (and the leading slot) are null, every present value normalized.
export const streamedRowValues = (cells: Untyped[]) => {
  const width = cells.reduce((max: number, cell: Untyped) => Math.max(max, cell.col), 0);
  const values = new Array(width + 1).fill(null);
  for (const cell of cells) values[cell.col] = normalizeStreamValue(cell.value);
  return values;
};

// A 1×1 PNG: a minimal image payload for anchoring on a sheet.
export const ONE_PX_PNG = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
    'base64',
  ),
);

/**
 * A spec reached for vocabulary this adapter does not map.
 *
 * It used to carry a `notImplemented` flag that made the runner report the behavior as **skipped**,
 * built when the library was incomplete and a case could legitimately outrun it. It cannot happen for
 * that reason any more, so a skip now means only one thing: a case asked for a key nobody wired, and
 * the corpus quietly declined to test it. That is the silent cap CLAUDE.md §3 forbids, so this is an
 * ordinary loud failure and the message says what to extend.
 *
 * It stays a distinct class because one caller still needs to tell it apart: `tryWriteWorkbook`
 * deliberately reports a *writer* error as data for a case to assert on, and an adapter gap must not
 * be able to masquerade as one.
 */
export class UnsupportedSpecError extends Error {
  constructor(message: string) {
    super(`corpus adapter: ${message}; extend adapters/ts-xlsx/spec-model.ts to cover it`);
    this.name = 'UnsupportedSpecError';
  }
}

export const unsupportedSpec = (message: string): UnsupportedSpecError =>
  new UnsupportedSpecError(message);

export const SUPPORTED_TOP_KEYS = new Set(['sheets', 'properties', 'definedNames']);
export const SUPPORTED_PROP_KEYS = new Set(['creator', 'lastModifiedBy', 'created', 'modified']);
export const SUPPORTED_SHEET_KEYS = new Set([
  'name',
  'state',
  'cells',
  'columns',
  'rows',
  'properties',
  'pageSetup',
  'pageMargins',
  'headerFooter',
  'tables',
  'merges',
  'autoFilter',
  'images',
  'background',
]);
export const SUPPORTED_CELL_KEYS = new Set([
  'ref',
  'value',
  'formula',
  'sharedFormula',
  'result',
  'hyperlink',
  'text',
  'tooltip',
  'fill',
  'numFmt',
  'font',
  'border',
  'alignment',
  'protection',
  'note',
]);
export const SUPPORTED_SHEET_PROP_KEYS = new Set(['defaultRowHeight', 'defaultColWidth']);
export const SUPPORTED_COLUMN_KEYS = new Set([
  'index',
  'width',
  'hidden',
  'numFmt',
  'fill',
  'font',
  'border',
  'alignment',
  'protection',
]);
export const SUPPORTED_ROW_KEYS = new Set([
  'index',
  'height',
  'hidden',
  'outlineLevel',
  'collapsed',
  'fill',
]);
export const SUPPORTED_PAGE_MARGIN_KEYS = new Set([
  'left',
  'right',
  'top',
  'bottom',
  'header',
  'footer',
]);
export const SUPPORTED_PAGE_SETUP_KEYS = new Set([
  'fitToPage',
  'fitToWidth',
  'fitToHeight',
  'scale',
  'orientation',
  'pageOrder',
  'paperSize',
]);
export const SUPPORTED_HEADER_FOOTER_KEYS = new Set([
  'oddHeader',
  'oddFooter',
  'evenHeader',
  'evenFooter',
  'firstHeader',
  'firstFooter',
]);
export const SUPPORTED_TABLE_KEYS = new Set([
  'name',
  'ref',
  'headers',
  'columnDefs',
  'rows',
  'headerRow',
  'totalsRow',
]);
export const SUPPORTED_TABLE_COLUMN_KEYS = new Set([
  'name',
  'totalsRowLabel',
  'totalsRowFunction',
  'totalsRowFormula',
]);

// Build an _xlnm.Print_Area refersTo from a comma-separated area (e.g. 'A1:F10,A12:F21' or 'A:D'):
// each range is made absolute ($-prefixed on every column and row bound) and sheet-qualified, exactly
// how Excel records a print area. A whole-column range keeps its column-only shape ($A:$D).
export const absolutizeRef = (ref: string) =>
  ref.replace(/([A-Z]+)/g, '$$$1').replace(/(\d+)/g, '$$$1');
export const printAreaRefersTo = (sheetName: string, area: string) =>
  area
    .split(',')
    .map((range) => `${sheetName}!${absolutizeRef(range)}`)
    .join(',');

export const toDate = (v: Untyped) =>
  v && typeof v === 'object' && v.invalidDate ? new Date(Number.NaN) : new Date(v);
export const isoOrNull = (d: Untyped) =>
  d instanceof Date && !Number.isNaN(d.getTime()) ? d.toISOString() : null;

// A JSON-serializable view of a read-back cell value: a Date becomes { date: iso } (null when
// invalid), every other object is deep-cloned, and a scalar passes through. Mirrors the oracle.
export const normalizeStreamValue = (v: Untyped) => {
  if (v instanceof Date) return {date: Number.isNaN(v.getTime()) ? null : v.toISOString()};
  if (v && typeof v === 'object') return JSON.parse(JSON.stringify(v));
  return v ?? null;
};

// Some specs express a rich-text run in the flat inline shape `{ text, bold, italic, … }`, while the
// rewrite models a run as `{ text, font: { … } }`. Translate a spec value into the model shape on the
// way in…
export const specValueToModel = (value: Untyped) => {
  if (value && typeof value === 'object' && Array.isArray(value.richText)) {
    return {
      richText: value.richText.map((run: Untyped) => {
        const {text, ...font} = run;
        return Object.keys(font).length ? {text, font} : {text};
      }),
    };
  }
  return value;
};

// …and flatten a read-back run's `font` facets back onto the run on the way out, so a spec asserting
// on `run.bold` sees the shape it wrote.
export const modelValueToSpec = (value: Untyped) => {
  if (value && typeof value === 'object' && Array.isArray(value.richText)) {
    return {
      richText: value.richText.map(({text, font}: Untyped) => ({text, ...font})),
    };
  }
  return value;
};

// Map a declarative spec onto the library's Workbook model, throwing `UnsupportedSpecError` the moment
// the spec uses vocabulary this adapter does not map.
export function buildFrom(spec: Untyped = {}) {
  for (const k of Object.keys(spec)) {
    if (!SUPPORTED_TOP_KEYS.has(k)) throw unsupportedSpec(`spec.${k} not supported yet`);
  }
  const workbook = new Workbook();

  const p = spec.properties || {};
  for (const k of Object.keys(p)) {
    if (!SUPPORTED_PROP_KEYS.has(k)) throw unsupportedSpec(`properties.${k} not supported yet`);
  }
  if (p.creator !== undefined) workbook.properties.creator = p.creator;
  if (p.lastModifiedBy !== undefined) workbook.properties.lastModifiedBy = p.lastModifiedBy;
  if (p.created !== undefined) workbook.properties.created = toDate(p.created);
  if (p.modified !== undefined) workbook.properties.modified = toDate(p.modified);

  for (const s of spec.sheets || []) {
    for (const k of Object.keys(s)) {
      if (!SUPPORTED_SHEET_KEYS.has(k)) throw unsupportedSpec(`sheet.${k} not supported yet`);
    }
    const sheet = workbook.addWorksheet(s.name, s.state ? {state: s.state} : undefined);

    const sp = s.properties || {};
    for (const k of Object.keys(sp)) {
      if (!SUPPORTED_SHEET_PROP_KEYS.has(k))
        throw unsupportedSpec(`sheet.properties.${k} not supported yet`);
    }
    if (sp.defaultRowHeight !== undefined) sheet.properties.defaultRowHeight = sp.defaultRowHeight;
    if (sp.defaultColWidth !== undefined) sheet.properties.defaultColWidth = sp.defaultColWidth;

    const pm = s.pageMargins || {};
    for (const k of Object.keys(pm)) {
      if (!SUPPORTED_PAGE_MARGIN_KEYS.has(k))
        throw unsupportedSpec(`pageMargins.${k} not supported yet`);
      (sheet.pageMargins as Record<string, unknown>)[k] = pm[k];
    }

    const psu = s.pageSetup || {};
    for (const k of Object.keys(psu)) {
      if (!SUPPORTED_PAGE_SETUP_KEYS.has(k))
        throw unsupportedSpec(`pageSetup.${k} not supported yet`);
      (sheet.pageSetup as Record<string, unknown>)[k] = psu[k];
    }

    const hf = s.headerFooter || {};
    for (const k of Object.keys(hf)) {
      if (!SUPPORTED_HEADER_FOOTER_KEYS.has(k))
        throw unsupportedSpec(`headerFooter.${k} not supported yet`);
      (sheet.headerFooter as Record<string, unknown>)[k] = hf[k];
    }

    for (const t of s.tables || []) {
      for (const k of Object.keys(t)) {
        if (!SUPPORTED_TABLE_KEYS.has(k)) throw unsupportedSpec(`table.${k} not supported yet`);
      }
      let columns: Untyped[];
      if (t.columnDefs) {
        for (const cd of t.columnDefs) {
          for (const k of Object.keys(cd)) {
            if (!SUPPORTED_TABLE_COLUMN_KEYS.has(k))
              throw unsupportedSpec(`table.columnDefs.${k} not supported yet`);
          }
        }
        columns = t.columnDefs.map((cd: Untyped) => {
          const col: Untyped = {name: cd.name};
          if (cd.totalsRowLabel !== undefined) col.totalsRowLabel = cd.totalsRowLabel;
          if (cd.totalsRowFunction !== undefined) col.totalsRowFunction = cd.totalsRowFunction;
          if (cd.totalsRowFormula !== undefined) col.totalsRowFormula = cd.totalsRowFormula;
          return col;
        });
      } else {
        columns = (t.headers || []).map((name: Untyped) => ({name}));
      }
      // A spec may express a table ref as the full occupied range (`A1:B3`), the shape the oracle
      // accepts, while the model anchors at the single top-left cell and derives the range from the
      // row count. Take the anchor; the declared row count reconstructs the same range.
      const options: Untyped = {
        name: t.name,
        ref: t.ref.split(':')[0],
        columns,
        rowCount: (t.rows || []).length,
      };
      if (t.headerRow !== undefined) options.headerRow = t.headerRow;
      if (t.totalsRow !== undefined) options.totalsRow = t.totalsRow;
      sheet.addTable(options);

      // Materialize the declared data rows into the grid. `rowCount` above only sizes the table's
      // range; a body cell exists in `sheetData` only if something writes it. Excel writes the body
      // cells and the table range as one fact, so a spec asserting over body content (column styles
      // reaching data cells, dimension, shared strings) needs the cells actually present. Write below
      // the header row (materialized by addTable), anchorRow for a headerless table, one below it
      // otherwise, and leave the totals row (if any) to its own materialization. A later `s.cells`
      // entry still wins, since cells are applied after tables.
      const anchor = decodeAddress(options.ref);
      const dataTop = (anchor.row ?? 1) + (t.headerRow === false ? 0 : 1);
      (t.rows || []).forEach((rowValues: Untyped, r: number) => {
        rowValues.forEach((value: Untyped, c: number) => {
          if (value === undefined || value === null) return;
          sheet.getCell(encodeAddress((anchor.col ?? 1) + c, dataTop + r)).value = value;
        });
      });
    }

    for (const range of s.merges || []) sheet.mergeCells(range);

    // A spec autoFilter is either a bare range string or the structured {ref, columns} shape; the
    // model's setter accepts both, so pass it through verbatim.
    if (s.autoFilter !== undefined) sheet.autoFilter = s.autoFilter;

    for (const col of s.columns || []) {
      for (const k of Object.keys(col)) {
        if (!SUPPORTED_COLUMN_KEYS.has(k)) throw unsupportedSpec(`column.${k} not supported yet`);
      }
      const target = sheet.getColumn(col.index);
      if (col.width !== undefined) target.width = col.width;
      if (col.hidden !== undefined) target.hidden = col.hidden;
      if (col.numFmt !== undefined) target.numFmt = col.numFmt;
      if (col.fill !== undefined) target.fill = col.fill;
      if (col.font !== undefined) target.font = col.font;
      if (col.border !== undefined) target.border = col.border;
      if (col.alignment !== undefined) target.alignment = col.alignment;
      if (col.protection !== undefined) target.protection = col.protection;
    }

    for (const row of s.rows || []) {
      for (const k of Object.keys(row)) {
        if (!SUPPORTED_ROW_KEYS.has(k)) throw unsupportedSpec(`row.${k} not supported yet`);
      }
      const target = sheet.getRow(row.index);
      if (row.height !== undefined) target.height = row.height;
      if (row.hidden !== undefined) target.hidden = row.hidden;
      if (row.outlineLevel !== undefined) target.outlineLevel = row.outlineLevel;
      if (row.collapsed !== undefined) target.collapsed = row.collapsed;
      if (row.fill !== undefined) target.fill = row.fill;
    }

    // Images build after columns and rows so a fractional anchor resolves against the sheet's real
    // column widths and row heights (the model floors + offsets at addImage time).
    for (const img of s.images || []) {
      // A spec omits `extension` to mean the default 'png'; it sets the key (to a dirty or missing
      // value) on purpose to exercise the library's write-side extension sanitisation. Pass the raw
      // value through: `workbook.addImage` normalises a leading dot / query string / missing hint.
      const options =
        'extension' in img ? {buffer: ONE_PX_PNG, extension: img.extension} : {buffer: ONE_PX_PNG};
      anchorSpecImage(sheet, workbook.addImage(options), img.range);
    }
    // A sheet background is a workbook image tiled behind the grid, not anchored: it rides its own
    // worksheet `<picture>` relationship, so a case can assert it coexists with comment/VML parts.
    if (s.background) {
      sheet.addBackgroundImage(
        workbook.addImage({buffer: ONE_PX_PNG, extension: s.background.extension || 'png'}),
      );
    }

    for (const c of s.cells || []) {
      for (const k of Object.keys(c)) {
        if (!SUPPORTED_CELL_KEYS.has(k)) throw unsupportedSpec(`cell.${k} not supported yet`);
      }
      const cell = sheet.getCell(c.ref);
      if ('hyperlink' in c) {
        // The label is the cell's own value, a plain string or rich text, and the link sits beside it.
        if (c.text !== undefined) cell.value = c.text;
        sheet.addHyperlink({
          ref: c.ref,
          target: c.hyperlink,
          ...(c.tooltip !== undefined ? {tooltip: c.tooltip} : {}),
        });
      } else if ('formula' in c) {
        cell.value = 'result' in c ? {formula: c.formula, result: c.result} : {formula: c.formula};
      } else if ('sharedFormula' in c) {
        // A shared-formula clone names its master by address; the master is a plain formula cell.
        cell.value =
          'result' in c
            ? {sharedFormula: c.sharedFormula, result: c.result}
            : {sharedFormula: c.sharedFormula};
      } else if ('value' in c) {
        const v = c.value;
        if (v !== null && typeof v === 'object') {
          // A structured date value materializes a Date; every other object shape is a value
          // kind the writer does not model yet, so skip the behavior rather than mis-serialize.
          if (v.invalidDate) cell.value = new Date(Number.NaN);
          else if (v.date) cell.value = toDate(v.date);
          else if (Array.isArray(v.richText)) cell.value = {richText: v.richText};
          else throw unsupportedSpec(`cell value shape ${JSON.stringify(v)} not supported yet`);
        } else {
          cell.value = v;
        }
      }
      if (c.fill !== undefined) cell.fill = c.fill;
      if (c.numFmt !== undefined) cell.numFmt = c.numFmt;
      if (c.font !== undefined) cell.font = c.font;
      if (c.border !== undefined) cell.border = c.border;
      if (c.alignment !== undefined) cell.alignment = c.alignment;
      if (c.protection !== undefined) cell.protection = c.protection;
      // A note attaches a comments part + legacy VML drawing; a case pairs it with a background image
      // to assert the two features' worksheet relationships never collide.
      if (c.note !== undefined) cell.note = c.note;
    }
  }

  // Workbook-level defined names are added after every sheet exists, since a scoped name targets a
  // sheet by name. The corpus spec expresses a name as one-or-more ranges; the model stores a single
  // refersTo formula, so the ranges join into the comma-separated union OOXML persists in one element.
  for (const dn of spec.definedNames || []) {
    workbook.defineName({
      name: dn.name,
      refersTo: (dn.ranges || []).join(','),
      ...(dn.scope !== undefined ? {scope: dn.scope} : {}),
    });
  }
  return workbook;
}

// Mirror current.mjs's normalizeCell for the rewrite's Cell: a plain JSON view of the
// value that survived the round-trip. Style facets are absent until the reader reads
// them, matching the contract that an unmaterialized facet is simply not present.
// A cell a hyperlink covers is reported as { hyperlink, text, tooltip? }: the link's target and tooltip,
// which live on the sheet, and the cell's own value as the label.
export function normalizeRewriteCell(cell: Untyped, sheet: WorksheetInstance) {
  const v = cell.value;
  const link = sheet.hyperlinkAt(cell.address);
  let out: Untyped;
  if (link !== undefined) {
    out = {
      hyperlink: link.target,
      text: v,
      ...(link.tooltip !== undefined ? {tooltip: link.tooltip} : {}),
    };
  } else if (v && typeof v === 'object' && 'sharedFormula' in v) {
    out = {sharedFormula: v.sharedFormula, formula: v.formula ?? null, result: v.result ?? null};
  } else if (v && typeof v === 'object' && 'formula' in v)
    out = {formula: v.formula, result: v.result ?? null};
  else if (v instanceof Date) out = {value: Number.isNaN(v.getTime()) ? null : v.toISOString()};
  else out = {value: v ?? null};
  // A style facet is reported only when the round-trip materialized it, matching the
  // contract that an unset facet is simply absent (never an empty placeholder).
  if (cell.fill !== undefined) out.fill = cell.fill;
  if (cell.numFmt) out.numFmt = cell.numFmt;
  if (cell.font !== undefined) out.font = cell.font;
  if (cell.border !== undefined) out.border = cell.border;
  if (cell.alignment !== undefined) out.alignment = cell.alignment;
  if (cell.protection !== undefined) out.protection = cell.protection;
  // A note is cell metadata, reported only when the round-trip preserved one; mirrors the oracle so a
  // case can assert a comment survives alongside a table/background rather than reading undefined.
  if (cell.note !== undefined) out.note = cell.note;
  return out;
}

// Decompose an ExcelJS-shaped aggregate style object onto the rewrite's per-facet setters.
// The rewrite has no `.style` aggregate: each cell owns independent facet fields, so "assign
// one base style to two cells" (the shared-style aliasing setup) is just assigning each facet
// present. Assigning the SAME base object to two cells shares the facet references: exactly the
// aliasing a copy-on-write setter must not let bleed when one cell is later mutated.
export function applyStyle(cell: Untyped, style: Untyped) {
  if (style.fill !== undefined) cell.fill = style.fill;
  if (style.numFmt !== undefined) cell.numFmt = style.numFmt;
  if (style.font !== undefined) cell.font = style.font;
  if (style.border !== undefined) cell.border = style.border;
  if (style.alignment !== undefined) cell.alignment = style.alignment;
  if (style.protection !== undefined) cell.protection = style.protection;
}
