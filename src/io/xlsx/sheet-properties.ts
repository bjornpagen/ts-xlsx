// The worksheet's non-cell property blocks, both directions: `<sheetPr>`, `<sheetViews>`,
// `<sheetProtection>`, `<autoFilter>`, and the print settings (`<printOptions>`, `<pageMargins>`,
// `<pageSetup>`, `<headerFooter>`, `<rowBreaks>`/`<colBreaks>`).
//
// The writing half renders one CT_Worksheet child (or child group) at a time, independently of the
// row/cell body `worksheet-xml.ts` orchestrates them alongside. The reading half sits below it and
// says why it is here rather than in the reader.

import {encodeAddress} from '../../core/address.ts';
import {
  type AutoFilter,
  type FilterColumn,
  type FilterCriteria,
  isCustomFilterOperator,
} from '../../core/autofilter.ts';
import {
  type HeaderFooter,
  HEADER_FOOTER_ELEMENTS,
  MARGIN_SIDES,
  PAGE_SETUP_FACETS,
  type PageBreak,
  type PageMargins,
  type PageSetup,
  PRINT_OPTION_FLAGS,
  type PrintOptions,
} from '../../core/page-setup.ts';
import {
  isLegacyPasswordHash,
  SHEET_PROTECTION_FLAGS,
  type SheetProtection,
  type SheetProtectionCredential,
  type SheetProtectionFlags,
} from '../../core/protection.ts';
import type {OutlineProperties, SheetView, Worksheet} from '../../core/worksheet.ts';
import {numFinite, numInteger} from '../../xml/xml-attrs.ts';
import {boolStrict, boolTristate, type XmlAttributes} from '../../xml/xml-scan.ts';
import {
  boolAttr,
  checkedToken,
  escapeAttr,
  escapeSpreadsheetText,
  numAttr,
  numberText,
} from '../../xml/xml.ts';
import {colorAttrs, parseColor} from './color-xml.ts';

// `<sheetViews>` holds the sheet's single view. A frozen view adds a `<pane>` recording the split
// and a `<selection>` naming the pane the split activates, exactly as Excel writes it. A normal
// view carries neither, so unfreezing leaves no leftover `<pane>` that would trip a repair prompt.
// The active pane is whichever scrolling region the freeze creates: bottom-right when both axes are
// frozen, else top-right (columns only) or bottom-left (rows only).
//
// `active` marks this sheet as the one selected on open (`tabSelected`). Exactly one sheet in a
// workbook carries it: with none, the consumer opens with no sheet view initialised; with several,
// the sheets form a *group selection*, where an edit to one is applied to all of them. The caller
// (`worksheetXml`, fed from `Workbook.activeTabIndex`) is what guarantees the "exactly one".
export function sheetViewsXml(view: SheetView, active: boolean): string {
  const selected = active ? ' tabSelected="1"' : '';
  // Excel defaults the grid on, so only an explicit `false` is worth an attribute; leaving it unset
  // keeps a sheet that never asked about gridlines byte-clean through a round-trip.
  const gridLines = view.showGridLines === false ? ' showGridLines="0"' : '';
  const xSplit = view.xSplit ?? 0;
  const ySplit = view.ySplit ?? 0;
  if (view.state !== 'frozen' || (xSplit === 0 && ySplit === 0)) {
    return `<sheetViews><sheetView${gridLines}${selected} workbookViewId="0"/></sheetViews>`;
  }
  const topLeftCell = view.topLeftCell ?? encodeAddress(xSplit + 1, ySplit + 1);
  const activePane =
    xSplit > 0 && ySplit > 0 ? 'bottomRight' : xSplit > 0 ? 'topRight' : 'bottomLeft';
  const pane =
    '<pane' +
    (xSplit > 0 ? ` xSplit="${xSplit}"` : '') +
    (ySplit > 0 ? ` ySplit="${ySplit}"` : '') +
    ` topLeftCell="${escapeAttr(topLeftCell)}" activePane="${activePane}" state="frozen"/>`;
  const selection = `<selection pane="${activePane}" activeCell="${escapeAttr(topLeftCell)}" sqref="${escapeAttr(topLeftCell)}"/>`;
  return `<sheetViews><sheetView${gridLines}${selected} workbookViewId="0">${pane}${selection}</sheetView></sheetViews>`;
}

// `<sheetPr>` carries the sheet's appearance properties: the tab colour, the outline
// summary-position flags, and the fit-to-page flag, plus the sheet's VBA identity as an attribute of
// its own. It is the first child of `<worksheet>` in CT_Worksheet order; its own children follow
// CT_SheetPr order: `<tabColor>`, `<outlinePr>`, then `<pageSetUpPr>`. Omitted entirely when the
// sheet carries none of them, so an unadorned sheet stays byte-clean.
export function sheetPrXml(sheet: Worksheet): string {
  const attrs = sheet.codeName === undefined ? '' : ` codeName="${escapeAttr(sheet.codeName)}"`;
  const children =
    (sheet.tabColor !== undefined ? `<tabColor ${colorAttrs(sheet.tabColor)}/>` : '') +
    outlinePrXml(sheet.outline) +
    pageSetUpPrXml(sheet.pageSetup);
  if (attrs === '' && children === '') return '';
  // A code name with nothing else to say is still an element, and a self-closing one: `<sheetPr/>`
  // with no children is what Excel writes for a macro-enabled sheet that is otherwise plain.
  return children === '' ? `<sheetPr${attrs}/>` : `<sheetPr${attrs}>${children}</sheetPr>`;
}

// `<pageSetUpPr>` holds the fit-to-page toggle, which lives on the sheet properties rather than on
// `<pageSetup>`. Excel reads it from here to decide whether the `fitToWidth`/`fitToHeight` counts
// or the fixed `scale` govern printing. Emitted only when the author set the flag.
function pageSetUpPrXml(pageSetup: PageSetup): string {
  return pageSetup.fitToPage ? '<pageSetUpPr fitToPage="1"/>' : '';
}

// `<outlinePr>` carries only the summary-position flags today. Each is emitted solely when the
// caller set it, so an inverted placement (`summaryBelow="0"`) is honoured while an untouched sheet
// keeps the element out of the file entirely.
function outlinePrXml(outline: OutlineProperties): string {
  const attrs =
    boolAttr('summaryBelow', outline.summaryBelow) + boolAttr('summaryRight', outline.summaryRight);
  return attrs === '' ? '' : `<outlinePr${attrs}/>`;
}

// Each sheet-protection flag maps to a `<sheetProtection>` attribute whose value is INVERTED
// from the author-facing allow-flag: the attribute records that an operation is *forbidden*
// ("1"), so `allow: true` serialises as "0". Only a value that differs from OOXML's per-
// attribute default (see SHEET_PROTECTION_FLAGS) is written: most editing operations default
// to forbidden under protection, while selecting cells defaults to permitted.
//
// <sheetProtection> is what makes the per-cell locked/hidden flags bite. `sheet="1"` marks the
// sheet protected; the password credential (when present) guards lifting it; the flag attributes
// carve out the operations that stay available. Every credential string is escaped: they arrive
// verbatim from a foreign package or straight from a caller-supplied `SheetProtectionCredential`,
// and `algorithmName` is prose even where the base64 salt and hash would not need it.
export function sheetProtectionXml(protection: SheetProtection | undefined): string {
  if (protection === undefined) return '';
  const {flags, credential, legacyPasswordHash} = protection;
  let attrs = '';
  // A legacy hash is only ever one a file carried, and it is interpolated, so it is held to its schema
  // type here: anything but four hex digits reached the model some way other than the reader.
  if (legacyPasswordHash !== undefined) {
    attrs += ` password="${checkedToken(legacyPasswordHash, isLegacyPasswordHash, 'legacy sheet password hash')}"`;
  }
  if (credential !== undefined) {
    attrs +=
      ` algorithmName="${escapeAttr(credential.algorithmName)}"` +
      ` hashValue="${escapeAttr(credential.hashValue)}"` +
      ` saltValue="${escapeAttr(credential.saltValue)}"` +
      numAttr('spinCount', credential.spinCount);
  }
  attrs += ' sheet="1"';
  for (const {key, defaultForbidden} of SHEET_PROTECTION_FLAGS) {
    const allow = flags[key];
    if (allow === undefined) continue;
    const forbidden = !allow;
    if (forbidden === defaultForbidden) continue;
    attrs += boolAttr(key, forbidden);
  }
  return `<sheetProtection${attrs}/>`;
}

// The sheet's autofilter: `<autoFilter ref="A1:C10"/>` when it only draws dropdowns, or with nested
// `<filterColumn>` children when columns carry criteria. Its companion `_FilterDatabase` defined name
// (the range Excel derives filtering from) is written in the workbook part, so a sheet with no filter
// emits nothing here and nothing there.
export function autoFilterXml(filter: AutoFilter | undefined): string {
  if (filter === undefined) return '';
  const ref = escapeAttr(filter.ref);
  if (filter.columns.length === 0) return `<autoFilter ref="${ref}"/>`;
  return `<autoFilter ref="${ref}">${filter.columns.map(filterColumnXml).join('')}</autoFilter>`;
}

function filterColumnXml(column: FilterColumn): string {
  return `<filterColumn colId="${numberText(column.colId)}">${filterCriteriaXml(column.criteria)}</filterColumn>`;
}

// A values filter is `<filters>` with a `<filter val>` per allowed value (and `blank="1"` to admit
// empty cells); a custom filter is `<customFilters>` with one or two `<customFilter operator val>`
// predicates, `and="1"` when they are AND-combined rather than OR.
function filterCriteriaXml(criteria: FilterCriteria): string {
  if (criteria.kind === 'values') {
    const blankAttr = criteria.blank ? ' blank="1"' : '';
    const filters = criteria.values.map((value) => `<filter val="${escapeAttr(value)}"/>`).join('');
    return `<filters${blankAttr}>${filters}</filters>`;
  }
  const andAttr = criteria.and ? ' and="1"' : '';
  const predicates = criteria.predicates
    .map(
      (p) =>
        `<customFilter operator="${checkedToken(p.operator, isCustomFilterOperator, 'custom filter operator')}" val="${escapeAttr(p.val)}"/>`,
    )
    .join('');
  return `<customFilters${andAttr}>${predicates}</customFilters>`;
}

// The even- and first-page variants are silently ignored by Excel unless differentOddEven /
// differentFirst are set, so the writer derives each flag from whether any variant in its class was
// provided.
export function headerFooterXml(hf: HeaderFooter): string {
  const children = HEADER_FOOTER_ELEMENTS.filter((key) => hf[key] !== undefined);
  if (children.length === 0) return '';
  const differentOddEven = hf.evenHeader !== undefined || hf.evenFooter !== undefined;
  const differentFirst = hf.firstHeader !== undefined || hf.firstFooter !== undefined;
  let attrs = '';
  if (differentOddEven) attrs += ' differentOddEven="1"';
  if (differentFirst) attrs += ' differentFirst="1"';
  const body = children
    // `escapeSpreadsheetText`, not `escapeText`: Excel applies the `_xHHHH_` convention to header
    // text exactly as it does to a cell value: it decodes an escape on load and writes one back on
    // save (measured over COM). So a header may carry a character XML itself cannot, and a header
    // that legitimately reads `_x0041_` must have its underscore escaped or it would decode to `A`.
    // The `?? ''` is unreachable: `children` is the keys `hf` carries. TypeScript cannot see that
    // across the filter, and a default is a smaller claim than an assertion.
    .map((key) => `<${key}>${escapeSpreadsheetText(hf[key] ?? '')}</${key}>`)
    .join('');
  return `<headerFooter${attrs}>${body}</headerFooter>`;
}

// Excel's "Normal" margins, in inches: the defaults Excel writes for an untouched sheet.
const DEFAULT_MARGINS = {
  left: 0.7,
  right: 0.7,
  top: 0.75,
  bottom: 0.75,
  header: 0.3,
  footer: 0.3,
} as const;

// `<printOptions>` carries the print-toggle flags and sits just before `<pageMargins>` in
// CT_Worksheet order. Each attribute is emitted only when the model carries it, as an explicit
// `="1"`/`="0"` so a caller can force a flag off against Excel's default, and an untouched sheet
// keeps the element out of the file entirely.
export function printOptionsXml(printOptions: PrintOptions): string {
  const attrs = PRINT_OPTION_FLAGS.map((flag) => boolAttr(flag, printOptions[flag])).join('');
  return attrs === '' ? '' : `<printOptions${attrs}/>`;
}

// OOXML's <pageMargins> is all-or-nothing: setting any one margin requires all six, or Excel
// repairs the file. So the element is emitted only when the caller set at least one, and the
// untouched sides fall back to the Normal-preset defaults.
export function pageMarginsXml(margins: PageMargins): string {
  if (MARGIN_SIDES.every((side) => margins[side] === undefined)) return '';
  const attrs = MARGIN_SIDES.map(
    (side) => `${side}="${numberText(margins[side] ?? DEFAULT_MARGINS[side])}"`,
  ).join(' ');
  return `<pageMargins ${attrs}/>`;
}

// `<pageSetup>` carries the print-scaling attributes (all but `fitToPage`, which is a `<sheetPr>`
// flag). It sits between `<pageMargins>` and `<headerFooter>` in CT_Worksheet order. Each attribute
// is emitted only when the author set it, so an untouched sheet keeps the element out of the file
// and a partially-set one never fabricates the counts Excel would otherwise default. A non-null
// `printerSettingsRelId` links the sheet's opaque printer-settings blob and forces the element out
// even when no scaling attribute is set: the reference is the only thing the model has to carry.
export function pageSetupXml(pageSetup: PageSetup, printerSettingsRelId: string | null): string {
  let attrs = '';
  for (const facet of PAGE_SETUP_FACETS) {
    switch (facet.kind) {
      case 'count':
        attrs += numAttr(facet.key, pageSetup[facet.key]);
        break;
      case 'token': {
        const value = pageSetup[facet.key];
        if (value !== undefined) {
          attrs += ` ${facet.key}="${checkedToken(value, facet.isValid, facet.label)}"`;
        }
        break;
      }
    }
  }
  if (printerSettingsRelId !== null) attrs += ` r:id="${printerSettingsRelId}"`;
  return attrs === '' ? '' : `<pageSetup${attrs}/>`;
}

// Page breaks (`<rowBreaks>`/`<colBreaks>`): one `<brk>` per break, attributes in CT_Break's order.
// Excel records the running total (`count`) and the manual subset (`manualBreakCount`), and the second
// counts only breaks carrying `man`. A break a file declared automatic is written back without it; one
// authored with no `man` at all is manual, since a caller pushing `{id: 3}` asked for a break and Excel
// stores no other kind. Row and column breaks share this shape, differing only in the wrapping element.
export function pageBreaksXml(
  breaks: readonly PageBreak[],
  element: 'rowBreaks' | 'colBreaks',
): string {
  if (breaks.length === 0) return '';
  const manual = breaks.filter((brk) => brk.man !== false).length;
  const brks = breaks
    .map(
      (brk) =>
        `<brk${numAttr('id', brk.id)}${numAttr('min', brk.min)}${numAttr('max', brk.max)}` +
        `${brk.man === false ? '' : ' man="1"'}/>`,
    )
    .join('');
  return `<${element} count="${breaks.length}" manualBreakCount="${manual}">${brks}</${element}>`;
}

// ---------------------------------------------------------------------------------------------
// The reading half of the same blocks.
//
// It lived in `read-worksheet.ts`, which meant a `<pageSetup>` change was an edit to two files that
// never saw each other, and left that module 548 lines of cell state machine, autofilter draft,
// page-break accumulator and six layout readers. Every other wire form in this tree keeps both
// directions together (`color-xml.ts`, `theme-xml.ts`, `tables.ts`, `hyperlinks.ts`,
// `data-validation.ts`, `conditional-formatting.ts`); `read-workbook-xml.ts` states the rule and
// claimed the workbook part was the last exception. It was not.
//
// The facet tables the two halves drive off (`PAGE_SETUP_FACETS`, `MARGIN_SIDES`,
// `PRINT_OPTION_FLAGS`, `SHEET_PROTECTION_FLAGS`) already live in `core/`, so what changes here is
// only which file the readings sit in.
// ---------------------------------------------------------------------------------------------

// Read a <sheetProtection> element back into a SheetProtection: the deserialization mirror of the
// writer. `sheet` defaults to false, so an element that omits it, or sets it false, records an
// *un*protected sheet and nothing is restored; reading absence as protected locked, on every save, a
// sheet its file had left open. Each flag attribute is the INVERSE of the author's allow-flag ("1"
// forbids, "0" permits), and only attributes actually present are carried, so an omitted
// (default-valued) flag stays absent, exactly what the writer emitted. A password is preserved
// verbatim, as the agile credential (algorithm, hash, salt, spin count) and as the legacy 16-bit
// `password` hash when that is four hex digits; there is no plaintext to recover, so neither is re-hashed.
export function parseSheetProtection(attrs: XmlAttributes): SheetProtection | undefined {
  if (!boolStrict(attrs.sheet)) return undefined;
  const flags: {-readonly [K in keyof SheetProtectionFlags]?: boolean} = {};
  for (const {key} of SHEET_PROTECTION_FLAGS) {
    const raw = attrs[key];
    if (raw !== undefined) flags[key] = !boolStrict(raw);
  }
  const {algorithmName, hashValue, saltValue, spinCount} = attrs;
  const spin = numInteger(spinCount, 0);
  if (
    algorithmName !== undefined &&
    hashValue !== undefined &&
    saltValue !== undefined &&
    spin !== undefined
  ) {
    const credential: SheetProtectionCredential = {
      algorithmName,
      hashValue,
      saltValue,
      spinCount: spin,
    };
    return {flags, credential, ...legacyHash(attrs.password)};
  }
  return {flags, ...legacyHash(attrs.password)};
}

// The legacy hash as a field to spread, present only when the file spelled it as its schema type.
function legacyHash(password: string | undefined): {legacyPasswordHash?: string} {
  return password !== undefined && isLegacyPasswordHash(password)
    ? {legacyPasswordHash: password}
    : {};
}

// Page-break accumulation. `<brk>` elements appear under both `<rowBreaks>` and `<colBreaks>`; a break
// container's open points the accumulator at that axis's list (null outside any container), so a
// `<brk>` lands on the right axis, and the matching close clears it. A self-closing
// `<rowBreaks/>`/`<colBreaks/>` fires no close, so a new open simply reassigns the target.
export class PageBreakAccumulator {
  #target: PageBreak[] | null = null;

  begin(target: PageBreak[]): void {
    this.#target = target;
  }

  end(): void {
    this.#target = null;
  }

  // `id` is the row/column the layout splits before; a non-positive or non-integer id is hostile input
  // and dropped rather than trusted. A `<brk>` outside any break container has no axis and is ignored.
  add(attrs: XmlAttributes): void {
    if (this.#target === null) return;
    const id = numInteger(attrs.id, 1);
    if (id === undefined) return;
    // `man` is recorded either way: a break the file left automatic has to stay automatic on write,
    // which an absent flag could not tell from a caller who never said.
    const brk: {id: number; min?: number; max?: number; man: boolean} = {
      id,
      man: boolStrict(attrs.man),
    };
    const min = numInteger(attrs.min, 0);
    if (min !== undefined) brk.min = min;
    const max = numInteger(attrs.max, 0);
    if (max !== undefined) brk.max = max;
    this.#target.push(brk);
  }
}

// Apply one `<sheetPr>` / `<sheetView>` / print-setup child to the sheet. These are the worksheet's
// layout and print metadata; grouping them here keeps the cell-reading switch a pure dispatch. Each
// records only what the source carried, so a file missing a facet leaves it unset and a re-write
// stays byte-clean. Each is read on open, from its attributes alone, which is why `<sheetView>`
// belongs here despite wrapping children: what this reads of it is attributes, and its `<pane>`
// child arrives as its own dispatch.
export function applySheetProperties(local: string, attrs: XmlAttributes, sheet: Worksheet): void {
  switch (local) {
    case 'sheetPr':
      // The element itself, for the one thing it carries as an attribute rather than a child: the
      // sheet's VBA identity. Nothing here reads it, but writing a `.xlsm` back without it leaves the
      // macros bound to a sheet name the project no longer finds.
      if (attrs.codeName !== undefined) sheet.codeName = attrs.codeName;
      break;
    case 'tabColor':
      // A `<sheetPr>` child.
      sheet.tabColor = parseColor(attrs);
      break;
    case 'outlinePr':
      // A `<sheetPr>` child.
      // An unrecognised token is dropped rather than read as present-and-true and written as `"1"`.
      for (const flag of ['summaryBelow', 'summaryRight'] as const) {
        const value = boolTristate(attrs[flag]);
        if (value !== undefined) sheet.outline[flag] = value;
      }
      break;
    case 'sheetView':
      // Excel omits `showGridLines` when the grid is on, so only a present-and-false attribute is
      // recorded. Leaving it unset otherwise is what keeps a re-write from fabricating the
      // attribute on every sheet that never mentioned it.
      if (boolTristate(attrs.showGridLines) === false) {
        sheet.view.showGridLines = false;
      }
      break;
    case 'pane':
      // A `<sheetView>` child recording a frozen (or split) pane. Only a frozen pane maps onto the
      // model's view; a source without one leaves `view` empty, so a re-write emits no pane.
      if (attrs.state === 'frozen' || attrs.state === 'frozenSplit') {
        sheet.view.state = 'frozen';
        // A split that is not a non-negative integer is dropped, not stored: `Worksheet.freeze()`
        // refuses the same value, and storing it here only defers the failure to the writer, which
        // adds the split to a cell ordinal and throws about a column the caller never named.
        const xSplit = numInteger(attrs.xSplit, 0);
        const ySplit = numInteger(attrs.ySplit, 0);
        if (xSplit !== undefined) sheet.view.xSplit = xSplit;
        if (ySplit !== undefined) sheet.view.ySplit = ySplit;
        if (attrs.topLeftCell !== undefined) sheet.view.topLeftCell = attrs.topLeftCell;
      }
      break;
    case 'pageSetUpPr':
      // The fit-to-page flag, a `<sheetPr>` child. Recorded only when the attribute is present, so a
      // `<pageSetUpPr>` present for other reasons (e.g. `autoPageBreaks`) leaves `fitToPage` unset.
      if (boolTristate(attrs.fitToPage) !== undefined) {
        sheet.pageSetup.fitToPage = boolStrict(attrs.fitToPage);
      }
      break;
    case 'printOptions':
      applyPrintOptions(sheet.printOptions, attrs);
      break;
    case 'pageMargins':
      applyMargins(sheet.pageMargins, attrs);
      break;
    case 'pageSetup':
      applyPageSetup(sheet.pageSetup, attrs);
      break;
  }
}

// Read the `<printOptions>` boolean toggles back onto the model, storing only the ones the source
// carried so a re-write stays byte-clean. An OOXML boolean is `1`/`true` for on and `0`/`false` for
// off; a present-but-unrecognised token is dropped rather than coerced.
function applyPrintOptions(printOptions: PrintOptions, attrs: XmlAttributes): void {
  for (const flag of PRINT_OPTION_FLAGS) {
    const value = boolTristate(attrs[flag]);
    if (value !== undefined) printOptions[flag] = value;
  }
}

function applyMargins(margins: PageMargins, attrs: XmlAttributes): void {
  for (const side of MARGIN_SIDES) {
    const value = numFinite(attrs[side]);
    if (value !== undefined) margins[side] = value;
  }
}

// Read the `<pageSetup>` print-scaling attributes back onto the model, setting only those the
// source carried so a re-write stays byte-clean. Each of the four is a count or an enumeration id,
// so a fractional or negative one carries no meaning and is dropped; the enumerated string
// attributes are trusted verbatim (an unexpected token round-trips harmlessly as an unknown string).
function applyPageSetup(pageSetup: PageSetup, attrs: XmlAttributes): void {
  for (const facet of PAGE_SETUP_FACETS) {
    const raw = attrs[facet.key];
    switch (facet.kind) {
      case 'count': {
        const value = numInteger(raw, 0);
        if (value !== undefined) pageSetup[facet.key] = value;
        break;
      }
      case 'token':
        if (raw !== undefined && facet.isValid(raw))
          assignPageSetupToken(pageSetup, facet.key, raw);
        break;
    }
  }
}

// One token attribute at a time, so the write's key type is a single member rather than the whole
// union and `pageSetup[key] = value` typechecks: the correlated-key access TypeScript cannot verify
// when the key is a union, the same shape `assignAlignmentToken` takes in read-styles.ts. The cast
// restates the guard's own proof: `isValid` has already accepted `raw` for this facet's
// enumeration, which the table cannot say in a type because both token entries share one shape.
function assignPageSetupToken<K extends 'pageOrder' | 'orientation'>(
  pageSetup: PageSetup,
  key: K,
  raw: string,
): void {
  pageSetup[key] = raw as PageSetup[K];
}
