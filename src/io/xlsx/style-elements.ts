// The stateless half of the stylesheet writer: one function per `styles.xml` element, each turning a
// model facet into its markup and holding no state at all.
//
// Split from `styles.ts`, which keeps the interning tables and the registry that assembles the part.
// The line between them is exactly that: nothing here touches a `StyleRegistry`, and the registry
// calls in here for every fragment it interns.
//
// The `<font>` element is not here. It is one wire form shared with a rich-text run's `<rPr>`, and it
// lives in `font-xml.ts` with the reader that consumes it.

import {
  ALIGNMENT_FACETS,
  type Alignment,
  type Border,
  type BorderEdge,
  type Fill,
  type GradientFill,
  isBorderStyle,
  isFillPatternType,
  type Protection,
} from '../../core/style.ts';
import type {DifferentialStyle} from '../../core/workbook-styles.ts';
import {
  checkedToken,
  escapeAttr,
  escapeFormatCode,
  nonDefaultNumAttr,
  numberText,
} from '../../xml/xml.ts';
import {colorAttrs} from './color-xml.ts';
import {fontXml} from './font-xml.ts';

// numFmt ids below 164 are reserved by ECMA-376 for the built-in formats every consumer
// knows implicitly; custom format codes are numbered from 164 up. Id 0 is General (no code).
export const CUSTOM_NUMFMT_BASE = 164;

// The Office default font's inner fragment, in the exact child order `fontXml` emits: the font 0 a
// registry built without a workbook falls back to. A registry built *with* one derives font 0 from
// its `defaultFont` instead, and for a plain `new Workbook()` that derivation lands on exactly this
// string; `styles.test.ts` guards the two against drifting.
export const DEFAULT_FONT_BODY =
  '<sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/>';
// The empty border: all five edges present but styleless. A border that overrides no edge
// serialises to exactly this, so it interns to the default border id 0 rather than a new one.
export const DEFAULT_BORDER = '<border><left/><right/><top/><bottom/><diagonal/></border>';

// One interned cell format. `fillId` 0 is no fill; `numFmtId` 0 is the General format;
// `fontId` 0 is the default font; `borderId` 0 is the empty border. `alignment` and `protection`
// hold the serialised `<alignment>`/`<protection>` attribute strings (empty when the cell has no
// explicit facet), carried inline because both are children of the xf rather than shared sub-tables.
export interface CellFormat {
  readonly fillId: number;
  readonly numFmtId: number;
  readonly fontId: number;
  readonly borderId: number;
  readonly alignment: string;
  readonly protection: string;
  readonly quotePrefix: boolean;
  // The named-style link (`xfId`) a cellXfs entry carries; 0 = Normal. A cellStyleXfs entry does not
  // itself nest, so it is serialised with this omitted.
  readonly xfId: number;
}

// The default xf: no facet, General format, linked to the Normal named style (xfId 0). Shared as the
// first entry of both the cell-format and named-style tables; never mutated (formats only append).
export const DEFAULT_FORMAT: CellFormat = {
  fillId: 0,
  numFmtId: 0,
  fontId: 0,
  borderId: 0,
  alignment: '',
  protection: '',
  quotePrefix: false,
  xfId: 0,
};

// A stable, collision-free key for a composed format so identical formats intern to one cellXfs entry.
export function formatSignature(format: CellFormat): string {
  return (
    `fill:${format.fillId}|numFmt:${format.numFmtId}|font:${format.fontId}|border:${format.borderId}|` +
    `align:${format.alignment}|protect:${format.protection}|quote:${format.quotePrefix}|xfId:${format.xfId}`
  );
}

const DEFAULT_FORMAT_SIGNATURE = formatSignature(DEFAULT_FORMAT);

/**
 * Whether a format is the do-nothing default: no facet, General number format, no quote prefix, and
 * linked to the Normal named style. Such a cellXfs entry adds nothing, so its owner needs no `s`.
 *
 * Asked of the signature rather than field by field, because the two questions have to be answered
 * from the same field set and both answer silently when they are not. A facet the signature forgets
 * makes two distinct formats intern to one entry, and every cell carrying the facet quietly loses it;
 * a facet *this* forgets forces a spurious entry and an `s` attribute onto every cell in the book.
 * Neither is a type error, and neither is visible in the file until someone compares two of them.
 */
export function isDefaultFormat(format: CellFormat): boolean {
  return formatSignature(format) === DEFAULT_FORMAT_SIGNATURE;
}

// Serialise one `<xf>`. A cellXfs entry passes its named-style link as `xfId`; a cellStyleXfs entry
// (the base a cell links *to*) passes `null` so the attribute is omitted, since it nests no further.
export function xfXml(format: CellFormat, xfId: number | null): string {
  const applyNumberFormat = format.numFmtId !== 0 ? ' applyNumberFormat="1"' : '';
  const applyFont = format.fontId !== 0 ? ' applyFont="1"' : '';
  const applyFill = format.fillId !== 0 ? ' applyFill="1"' : '';
  const applyBorder = format.borderId !== 0 ? ' applyBorder="1"' : '';
  const applyAlignment = format.alignment !== '' ? ' applyAlignment="1"' : '';
  const applyProtection = format.protection !== '' ? ' applyProtection="1"' : '';
  // `quotePrefix` is a CT_Xf attribute (after xfId, before the apply flags in schema order); it is
  // its own switch (there is no `applyQuotePrefix` flag) so it is emitted only when set.
  const quotePrefix = format.quotePrefix ? ' quotePrefix="1"' : '';
  const xfIdAttr = xfId === null ? '' : ` xfId="${xfId}"`;
  const open =
    `<xf numFmtId="${format.numFmtId}" fontId="${format.fontId}" fillId="${format.fillId}" ` +
    `borderId="${format.borderId}"${xfIdAttr}${quotePrefix}` +
    `${applyNumberFormat}${applyFont}${applyFill}${applyBorder}${applyAlignment}${applyProtection}`;
  // Alignment and protection are child elements of the xf, in that schema order; an xf carrying
  // either (or both) is not self-closing, while a plain one stays self-closing as before.
  const body =
    (format.alignment === '' ? '' : `<alignment ${format.alignment}/>`) +
    (format.protection === '' ? '' : `<protection ${format.protection}/>`);
  return body === '' ? `${open}/>` : `${open}>${body}</xf>`;
}

// One `<cellStyle>` entry mapping a name (and, for a built-in, its gallery id) to a cellStyleXfs index.
export function cellStyleTag(entry: {name: string; builtinId?: number; xfId: number}): string {
  const builtin = entry.builtinId === undefined ? '' : ` builtinId="${entry.builtinId}"`;
  return `<cellStyle name="${escapeAttr(entry.name)}" xfId="${entry.xfId}"${builtin}/>`;
}

// Serialise a cell's alignment as `<alignment>` attributes, driven by ALIGNMENT_FACETS so the writer
// and the reader cannot disagree on a facet's name, kind, or default. The table is in CT_CellAlignment
// order, which is the order ECMA-376 requires the attributes in. A facet at its default contributes
// nothing; an all-default alignment yields the empty string, so it forces neither an <alignment> child
// nor a distinct xf.
export function alignmentAttrs(alignment: Alignment): string {
  const parts: string[] = [];
  for (const facet of ALIGNMENT_FACETS) {
    switch (facet.kind) {
      case 'token': {
        const value = alignment[facet.key];
        if (value === undefined || value === facet.omit) break;
        parts.push(`${facet.key}="${checkedToken(value, facet.isValid, facet.label)}"`);
        break;
      }
      case 'number': {
        const attr = nonDefaultNumAttr(facet.key, alignment[facet.key], 0);
        if (attr !== '') parts.push(attr.trimStart());
        break;
      }
      case 'flag':
        if (alignment[facet.key]) parts.push(`${facet.key}="1"`);
        break;
    }
  }
  return parts.join(' ');
}

// Serialise a cell's protection as `<protection>` attributes. `locked` defaults to true in OOXML,
// so only an explicitly unlocked cell writes `locked="0"`; `hidden` defaults to false, so only a
// hidden cell writes `hidden="1"`. An all-default protection yields the empty string, forcing
// neither a <protection> child nor a distinct xf.
export function protectionAttrs(protection: Protection): string {
  const parts: string[] = [];
  if (protection.locked === false) parts.push('locked="0"');
  if (protection.hidden === true) parts.push('hidden="1"');
  return parts.join(' ');
}

// Serialise a differential style (CT_Dxf) in schema child order: font, numFmt, fill, border. Only the
// facets present are emitted: a dxf overrides exactly what it names and lets the cell's own style show
// through the rest. A dxf's pattern fill states the highlight through `bgColor`, matching how Excel
// writes a "fill with colour" conditional format.
export function dxfXml(style: DifferentialStyle): string {
  const parts: string[] = [];
  if (style.font !== undefined) {
    const font = fontXml(style.font);
    if (font !== '') parts.push(`<font>${font}</font>`);
  }
  // A dxf numFmt still needs an id; the code is what matters (dxf formats are not shared by id like
  // cell formats), so a fixed custom id carries it without a <numFmts> entry.
  if (typeof style.numFmt === 'string' && style.numFmt !== '') {
    parts.push(
      `<numFmt numFmtId="${CUSTOM_NUMFMT_BASE}" formatCode="${escapeFormatCode(style.numFmt)}"/>`,
    );
  }
  if (style.fill !== undefined) parts.push(patternFillXml(style.fill, {solidBgFallback: false}));
  if (style.border !== undefined) parts.push(borderXml(style.border));
  return `<dxf>${parts.join('')}</dxf>`;
}

// The gradient element shared by cell fills and dxf fills. Linear gradients carry a `degree`; path
// gradients carry inner-rectangle insets. A zero-valued attribute is its OOXML default and is omitted.
function gradientFillXml(fill: GradientFill): string {
  const attrs =
    (fill.gradient === 'path' ? ' type="path"' : '') +
    nonDefaultNumAttr('degree', fill.degree, 0) +
    nonDefaultNumAttr('left', fill.left, 0) +
    nonDefaultNumAttr('right', fill.right, 0) +
    nonDefaultNumAttr('top', fill.top, 0) +
    nonDefaultNumAttr('bottom', fill.bottom, 0);
  const stops = fill.stops
    .map(
      (stop) =>
        `<stop position="${numberText(stop.position)}"><color ${colorAttrs(stop.color)}/></stop>`,
    )
    .join('');
  return `<gradientFill${attrs}>${stops}</gradientFill>`;
}

// Serialise a border in ECMA-376 CT_Border child order (left, right, top, bottom, diagonal).
// Every edge element is always present, since a styleless `<left/>` is how OOXML says "no left
// border", so an all-absent border round-trips to the empty default rather than a new id.
export function borderXml(border: Border): string {
  const attrs =
    (border.diagonalUp ? ' diagonalUp="1"' : '') + (border.diagonalDown ? ' diagonalDown="1"' : '');
  return (
    `<border${attrs}>` +
    edgeXml('left', border.left) +
    edgeXml('right', border.right) +
    edgeXml('top', border.top) +
    edgeXml('bottom', border.bottom) +
    edgeXml('diagonal', border.diagonal) +
    '</border>'
  );
}

// One border edge: a styleless self-closing tag when absent, else the style attribute plus an
// optional colour child.
function edgeXml(tag: string, edge: BorderEdge | undefined): string {
  if (edge === undefined) return `<${tag}/>`;
  const style = checkedToken(edge.style, isBorderStyle, 'border style');
  if (edge.color === undefined) return `<${tag} style="${style}"/>`;
  return `<${tag} style="${style}"><color ${colorAttrs(edge.color)}/></${tag}>`;
}

// The `<fill>` element for a pattern or gradient fill. The two callers differ only in the solid-fill
// background fallback: a cell fill forces the automatic indexed placeholder onto a solid pattern that
// names no background (omitting it makes Excel render the fill as flat black) whereas a dxf states
// only the overrides it carries, so `solidBgFallback` gates that placeholder.
export function patternFillXml(fill: Fill, {solidBgFallback}: {solidBgFallback: boolean}): string {
  if (fill.type === 'gradient') return `<fill>${gradientFillXml(fill)}</fill>`;
  const fg = fill.fgColor ? `<fgColor ${colorAttrs(fill.fgColor)}/>` : '';
  const bg = fill.bgColor
    ? `<bgColor ${colorAttrs(fill.bgColor)}/>`
    : solidBgFallback && fill.pattern === 'solid'
      ? '<bgColor indexed="64"/>'
      : '';
  const pattern = checkedToken(fill.pattern, isFillPatternType, 'fill pattern');
  return `<fill><patternFill patternType="${pattern}">${fg}${bg}</patternFill></fill>`;
}
