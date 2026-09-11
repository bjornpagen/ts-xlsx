// The style-table reader: `xl/styles.xml` in, a flat table of resolved cell formats out. It is a
// single streaming pass over the shared sub-tables (`<numFmts>`, `<fills>`, `<fonts>`, `<borders>`)
// and the two xf tables (`<cellXfs>`, `<cellStyleXfs>`), flattening the id-indirection so a cell's
// `s` index maps straight to its facets. A construct it does not recognise is skipped, never guessed.
//
// Only the parsing is XML-specific. What an xf resolves *to* (`XfStyle`, the built-in number
// formats, layering a direct xf over the named style it links to, applying an xf to a cell) is a
// property of the OOXML style model rather than of its spelling, and lives above both codecs in
// `../style/xf-style.ts`; the `.xlsb` style reader parses BIFF12 records into the same four tables
// and finishes through the same `resolveStyleTable`.

import {
  ALIGNMENT_FACETS,
  type Alignment,
  type Border,
  type Color,
  type Fill,
  type Font,
  type GradientFill,
  type GradientStop,
  isBorderStyle,
  isFillPatternType,
  type Protection,
} from '../../core/style.ts';
import type {
  TableStyleDefinition,
  TableStyleNamespace,
  TableStyleTable,
} from '../../core/workbook-styles.ts';
import {numFinite, numInteger} from '../../xml/xml-attrs.ts';
import {
  closeEmptyElements,
  elementSubtrees,
  openElements,
  type SubtreeSelection,
} from '../../xml/xml-read.ts';
import {
  boolStrict,
  boolTristate,
  localName,
  type XmlAttributes,
  type XmlEvent,
  xmlEvents,
} from '../../xml/xml-scan.ts';
import {
  NO_PRESERVED_STYLE_TABLES,
  numFmtCodeFor,
  type PreservedStyleTables,
  protectionFrom,
  resolveStyleTable,
  type StyleLabel,
  type StyleTable,
  type XfDeps,
  type XfStyle,
} from '../style/xf-style.ts';
import {parseColor} from './color-xml.ts';
import {applyFontChild, type FontDraft} from './font-xml.ts';

// A mutable xf accumulator while an <xf> element streams in: its facet ids resolve on open, but
// the <alignment>/<protection> children (when present) arrive before the element closes, so the
// xf is held here and pushed on close rather than on open.
type XfDraft = {-readonly [K in keyof XfStyle]?: XfStyle[K]};

// A mutable border accumulator while a <border> element's edges stream in; frozen into a
// Border on close. The five edges match Border's; a bare styleless edge is simply never set.
type BorderDraft = {-readonly [K in keyof Border]?: Border[K]};

// The <fill> being read. It should hold one body, a <patternFill> or a <gradientFill>; `bodyRead` turns
// when the first closes, so a second body in a malformed fill is skipped rather than replacing it.
type FillDraft = {body: Fill | undefined; bodyRead: boolean};

// A <patternFill> being read: its pattern token and the colours its <fgColor>/<bgColor> children carry.
type PatternDraft = {pattern: string; fgColor: Color | undefined; bgColor: Color | undefined};

// A mutable gradient accumulator while a <gradientFill> streams in. `fill` builds up the frozen
// GradientFill (its stops appended as <stop>/<color> pairs close); `stopPosition`/`stopColor` hold the
// current <stop> until its close commits a {position, color} pair.
type GradientDraft = {
  fill: {-readonly [K in keyof GradientFill]: GradientFill[K]};
  stopPosition: number | null;
  stopColor: Color | undefined;
};

// The four sides plus the diagonal: the edge elements a <border> can hold, in the order the
// schema lists them. This one tuple drives the edge-name union, the membership set (which drives
// edge parsing without a per-name branch), and the "does any edge carry a style" scan below.
const BORDER_EDGE_NAMES = ['left', 'right', 'top', 'bottom', 'diagonal'] as const;
type BorderEdgeName = (typeof BORDER_EDGE_NAMES)[number];
const BORDER_EDGES = new Set<string>(BORDER_EDGE_NAMES);
const isBorderEdgeName = (name: string): name is BorderEdgeName => BORDER_EDGES.has(name);

// Style-table elements that commit on their close: a bare <font/>/<border/>/<fill/>/<patternFill/>/
// <gradientFill/>/<xf/> or a self-closing border edge is expanded to open+close so each commits
// exactly once in onClose, never in a duplicated (and easily-forgotten) self-closing branch.
const STYLE_EMPTY_CLOSES: ReadonlySet<string> = new Set([
  'font',
  'border',
  'fill',
  'patternFill',
  'gradientFill',
  'xf',
  ...BORDER_EDGES,
]);

// styles.xml is a shared table: <numFmts> defines custom format codes by id, <fills> lists
// the fills, and <cellXfs> lists the cell formats, each naming a fill and a number format by
// id. We flatten that indirection into one array, cellXfs index → resolved {fill, numFmt},
// so a cell/row/column style index maps straight to its facets. The schema orders <numFmts>
// and <fills> before <cellXfs>, so both lookups are complete before an xf references them.
export function parseStyleTable(xml: string): StyleTable {
  if (xml === '') return {cellXfs: [], namedStyles: [], preserved: NO_PRESERVED_STYLE_TABLES};
  let fills: ReadonlyArray<Fill | undefined> = [];
  let fonts: ReadonlyArray<Font | undefined> = [];
  let borders: ReadonlyArray<Border | undefined> = [];
  let numFmtCodes: ReadonlyMap<number, string> = new Map();
  const xfStyles: XfStyle[] = [];
  // The named-style layer: <cellStyleXfs> holds the base formats a cell's xfId links to; <cellStyles>
  // labels them by name/builtinId. Parsed in parallel with cellXfs, then zipped and merged below.
  const namedXfs: XfStyle[] = [];
  let cellStyleNames: ReadonlyArray<StyleLabel> = [];

  // One streaming pass, but each top-level sub-table drives its own focused sub-parser over the slice
  // of events between its open and close. The schema orders the shared tables (<numFmts>, <fonts>,
  // <fills>, <borders>) before the xf tables, so their results are complete before an <xf> resolves
  // against them. Every recognised container name is plural and unique to the styleSheet root, and none
  // appears inside a <dxf>'s singular <font>/<fill>/<border> children, so skipping an unrecognised
  // section here drops exactly what the old flat pass gated off with its `in*` flags.
  const events = closeEmptyElements(xmlEvents(xml), STYLE_EMPTY_CLOSES);
  let next = events.next();
  while (next.done !== true) {
    const event = next.value;
    if (event.kind === 'open' && !event.selfClosing) {
      switch (localName(event.name)) {
        case 'numFmts':
          numFmtCodes = parseNumFmts(events);
          break;
        case 'fonts':
          fonts = parseFonts(events);
          break;
        case 'fills':
          fills = parseFills(events);
          break;
        case 'borders':
          borders = parseBorders(events);
          break;
        case 'cellStyleXfs':
          namedXfs.push(
            ...parseXfTable(events, 'cellStyleXfs', {fills, fonts, borders, numFmtCodes}),
          );
          break;
        case 'cellXfs':
          xfStyles.push(...parseXfTable(events, 'cellXfs', {fills, fonts, borders, numFmtCodes}));
          break;
        case 'cellStyles':
          cellStyleNames = parseCellStyles(events);
          break;
      }
    }
    next = events.next();
  }

  // The four preserved sub-tables come out of one scan *of the four*, which is what four callers used
  // to take four regular expressions and four scans of their own to get. It is a second scan of the
  // part, deliberately: `elementSubtrees` captures verbatim source spans and the walk above builds a
  // model, so one pass cannot do both without the event stream carrying offsets it does not carry.
  // ADR-0004 records that seam and what it would cost to close.
  const {fragments, attributes} = elementSubtrees(xml, PRESERVED_SUBTREES);
  const preserved: PreservedStyleTables = {
    dxfs: fragments.get('dxfs') ?? [],
    indexedColors: fragments.get('indexedColors') ?? [],
    mruColors: fragments.get('mruColors') ?? [],
    tableStyles: buildTableStyleTable(
      xml,
      fragments.get('tableStyles') ?? [],
      attributes.get('tableStyles'),
    ),
  };
  return {
    ...resolveStyleTable({directXfs: xfStyles, namedXfs, labels: cellStyleNames, fonts}),
    preserved,
  };
}

// Pull events off the shared stream up to, and consuming, the close of `container`, yielding only
// those strictly inside it. A sub-table parser loops this to completion (never breaking), so it drives
// its own small state machine over exactly its section without closing the underlying generator, and
// the outer pass resumes at the element after the container's close.
function* until(events: Iterator<XmlEvent>, container: string): Generator<XmlEvent> {
  let next = events.next();
  while (next.done !== true) {
    const event = next.value;
    if (event.kind === 'close' && localName(event.name) === container) return;
    yield event;
    next = events.next();
  }
}

// <numFmts> entries are self-closing, so they are read on open. A code with no id, or the General
// id 0, contributes nothing.
function parseNumFmts(events: Iterator<XmlEvent>): ReadonlyMap<number, string> {
  const codes = new Map<number, string>();
  for (const event of until(events, 'numFmts')) {
    if (event.kind === 'open' && localName(event.name) === 'numFmt') {
      const id = numInteger(event.attrs.numFmtId, 1);
      if (id !== undefined && event.attrs.formatCode !== undefined) {
        codes.set(id, event.attrs.formatCode);
      }
    }
  }
  return codes;
}

function parseFonts(events: Iterator<XmlEvent>): ReadonlyArray<Font | undefined> {
  const fonts: Array<Font | undefined> = [];
  let fontDraft: FontDraft | null = null;
  for (const event of until(events, 'fonts')) {
    if (event.kind === 'open') {
      const local = localName(event.name);
      // A <font>'s children are self-closing, so they are read here on open.
      if (local === 'font') fontDraft = {};
      else if (fontDraft !== null) applyFontChild(fontDraft, local, event.attrs);
    } else if (event.kind === 'close' && localName(event.name) === 'font' && fontDraft !== null) {
      fonts.push(Object.keys(fontDraft).length > 0 ? fontDraft : undefined);
      fontDraft = null;
    }
  }
  return fonts;
}

// `fillId` is an index into this table, so a slot is committed on each <fill>'s close and nowhere else:
// exactly one per <fill>, whether its body parsed, was unrecognised, or was followed by a second body. A
// body outside any <fill> belongs to no slot and is skipped, so it cannot shift the fills after it.
function parseFills(events: Iterator<XmlEvent>): ReadonlyArray<Fill | undefined> {
  const fills: Array<Fill | undefined> = [];
  let fillDraft: FillDraft | null = null;
  let patternDraft: PatternDraft | null = null;
  let gradientDraft: GradientDraft | null = null;
  for (const event of until(events, 'fills')) {
    if (event.kind === 'open') {
      const attrs = event.attrs;
      // A body is read only while its <fill> has not yet had one.
      const bodyExpected = fillDraft !== null && !fillDraft.bodyRead;
      switch (localName(event.name)) {
        case 'fill':
          fillDraft = {body: undefined, bodyRead: false};
          break;
        case 'patternFill':
          if (!bodyExpected) break;
          patternDraft = {
            pattern: attrs.patternType ?? 'none',
            fgColor: undefined,
            bgColor: undefined,
          };
          break;
        case 'fgColor':
          if (patternDraft !== null) patternDraft.fgColor = parseColor(attrs);
          break;
        case 'bgColor':
          if (patternDraft !== null) patternDraft.bgColor = parseColor(attrs);
          break;
        case 'gradientFill':
          if (!bodyExpected) break;
          gradientDraft = {
            fill: {
              type: 'gradient',
              gradient: attrs.type === 'path' ? 'path' : 'linear',
              stops: [],
            },
            stopPosition: null,
            stopColor: undefined,
          };
          assignGradientNumbers(gradientDraft.fill, attrs);
          break;
        case 'stop':
          if (gradientDraft !== null) {
            gradientDraft.stopPosition = numFinite(attrs.position) ?? 0;
            gradientDraft.stopColor = undefined;
          }
          break;
        case 'color':
          // The colour of the open <stop>; committed to a GradientStop when the stop closes.
          if (gradientDraft !== null) gradientDraft.stopColor = parseColor(attrs);
          break;
      }
    } else if (event.kind === 'close') {
      switch (localName(event.name)) {
        case 'patternFill':
          if (fillDraft !== null && patternDraft !== null) {
            const {pattern, fgColor, bgColor} = patternDraft;
            fillDraft.body = toFill(pattern, fgColor, bgColor);
            fillDraft.bodyRead = true;
            patternDraft = null;
          }
          break;
        case 'stop':
          if (gradientDraft !== null && gradientDraft.stopPosition !== null) {
            const stop: GradientStop = {
              position: gradientDraft.stopPosition,
              color: gradientDraft.stopColor ?? {},
            };
            gradientDraft.fill.stops = [...gradientDraft.fill.stops, stop];
            gradientDraft.stopPosition = null;
            gradientDraft.stopColor = undefined;
          }
          break;
        case 'gradientFill':
          if (fillDraft !== null && gradientDraft !== null) {
            fillDraft.body = gradientDraft.fill;
            fillDraft.bodyRead = true;
            gradientDraft = null;
          }
          break;
        case 'fill':
          if (fillDraft !== null) fills.push(fillDraft.body);
          // A body left open by a malformed <fill> must not carry over into the next one.
          fillDraft = null;
          patternDraft = null;
          gradientDraft = null;
          break;
      }
    }
  }
  return fills;
}

function parseBorders(events: Iterator<XmlEvent>): ReadonlyArray<Border | undefined> {
  const borders: Array<Border | undefined> = [];
  let borderDraft: BorderDraft | null = null;
  // Which edge of the current border a nested <color> belongs to; null between edges.
  let currentEdge: BorderEdgeName | null = null;
  for (const event of until(events, 'borders')) {
    if (event.kind === 'open') {
      const local = localName(event.name);
      const attrs = event.attrs;
      if (local === 'border') {
        borderDraft = {};
        currentEdge = null;
        if (boolStrict(attrs.diagonalUp)) borderDraft.diagonalUp = true;
        if (boolStrict(attrs.diagonalDown)) borderDraft.diagonalDown = true;
      } else if (borderDraft !== null) {
        // A border's edges and their <color> children are all read on open (each is self-closing bar a
        // coloured edge, whose colour child is itself self-closing). An edge whose style is absent or
        // an unrecognised token is dropped: the side simply carries no border.
        if (isBorderEdgeName(local)) {
          if (attrs.style !== undefined && isBorderStyle(attrs.style)) {
            currentEdge = local;
            borderDraft[currentEdge] = {style: attrs.style};
          } else {
            currentEdge = null;
          }
        } else if (local === 'color' && currentEdge !== null) {
          const edge = borderDraft[currentEdge];
          if (edge !== undefined)
            borderDraft[currentEdge] = {style: edge.style, color: parseColor(attrs)};
        }
      }
    } else if (event.kind === 'close') {
      const local = localName(event.name);
      if (local === 'border') {
        if (borderDraft !== null) {
          borders.push(borderToStyle(borderDraft));
          borderDraft = null;
          currentEdge = null;
        }
      } else if (borderDraft !== null && BORDER_EDGES.has(local)) {
        // A coloured edge closes after its <color> child; drop the edge context so a stray later
        // <color> cannot attach to it.
        currentEdge = null;
      }
    }
  }
  return borders;
}

// Both <cellXfs> and <cellStyleXfs> hold <xf> with identical structure; they differ only in which
// table the result lands in and whether an xfId link is meaningful (only cellXfs entries link to a
// named style). One parser serves both, told by `container` which it is reading.
function parseXfTable(
  events: Iterator<XmlEvent>,
  container: 'cellXfs' | 'cellStyleXfs',
  deps: XfDeps,
): XfStyle[] {
  const xfs: XfStyle[] = [];
  const captureXfId = container === 'cellXfs';
  // The xf being read; held from open to close so its <alignment>/<protection> children can attach
  // before it is committed. null outside an <xf>.
  let pendingXf: XfDraft | null = null;
  for (const event of until(events, container)) {
    if (event.kind === 'open') {
      const local = localName(event.name);
      if (local === 'xf') {
        // Hold the xf open until its close so an <alignment>/<protection> child can attach first; a
        // self-closing <xf/> is expanded to a close, so it commits there too, child-free.
        pendingXf = resolveXf(event.attrs, deps, captureXfId);
      } else if (pendingXf !== null && local === 'alignment') {
        // An xf's <alignment> child arrives before the xf closes; attach it to the pending xf.
        const alignment = parseAlignment(event.attrs);
        if (alignment !== undefined) pendingXf.alignment = alignment;
      } else if (pendingXf !== null && local === 'protection') {
        // An xf's <protection> child likewise arrives before the xf closes.
        const protection = parseProtection(event.attrs);
        if (protection !== undefined) pendingXf.protection = protection;
      }
    } else if (event.kind === 'close' && localName(event.name) === 'xf' && pendingXf !== null) {
      xfs.push(pendingXf);
      pendingXf = null;
    }
  }
  return xfs;
}

// Resolve an <xf>'s facet ids against the shared sub-tables into a draft. `captureXfId` is set only
// for cellXfs entries, the sole table whose xfId links to a named style.
function resolveXf(attrs: XmlAttributes, deps: XfDeps, captureXfId: boolean): XfDraft {
  const fillId = numInteger(attrs.fillId, 0);
  const fill = fillId !== undefined ? deps.fills[fillId] : undefined;
  const fontId = numInteger(attrs.fontId, 0);
  // Font id 0 is the workbook default font (a real Calibri-11-style face), not an absence, unlike
  // border id 0, which is a genuinely empty border. So an xf naming font 0 resolves to that default
  // face, giving every cell a concrete font to render.
  const font = fontId !== undefined ? deps.fonts[fontId] : undefined;
  // Border id 0 is the empty default; only a custom border (id > 0) is an explicit one.
  const borderId = numInteger(attrs.borderId, 1);
  const border = borderId !== undefined ? deps.borders[borderId] : undefined;
  const numFmt = resolveNumFmt(attrs.numFmtId, deps.numFmtCodes);
  const draft: XfDraft = {};
  if (fill) draft.fill = fill;
  if (numFmt !== undefined) draft.numFmt = numFmt;
  if (font) draft.font = font;
  if (border) draft.border = border;
  // The quote-prefix flag is an attribute on the xf itself (no shared sub-table); carry it only when
  // set so an ordinary cell does not gain a spurious `quotePrefix: false`.
  if (boolStrict(attrs.quotePrefix)) draft.quotePrefix = true;
  // A cellXfs entry's xfId links it to a named style; capture it only when it points beyond the Normal
  // default (0), so an ordinary cell carries no spurious named-style link.
  if (captureXfId) {
    const xfId = numInteger(attrs.xfId, 1);
    if (xfId !== undefined) draft.xfId = xfId;
  }
  return draft;
}

// A <cellStyle> (inside <cellStyles>) names a cellStyleXfs entry by xfId; it is self-closing, so it
// is read on open.
function parseCellStyles(events: Iterator<XmlEvent>): ReadonlyArray<StyleLabel> {
  const names: StyleLabel[] = [];
  for (const event of until(events, 'cellStyles')) {
    if (event.kind === 'open' && localName(event.name) === 'cellStyle') {
      const attrs = event.attrs;
      const xfId = numInteger(attrs.xfId);
      if (xfId !== undefined) {
        const entry: {xfId: number; name?: string; builtinId?: number} = {xfId};
        if (attrs.name !== undefined) entry.name = attrs.name;
        const builtinId = numInteger(attrs.builtinId);
        if (builtinId !== undefined) entry.builtinId = builtinId;
        names.push(entry);
      }
    }
  }
  return names;
}

// An xf's numFmtId resolves against the custom codes first, then the built-in table; the
// General format (id 0) and any unrecognised id mean the cell carries no explicit format.
function resolveNumFmt(
  raw: string | undefined,
  custom: ReadonlyMap<number, string>,
): string | undefined {
  const id = numInteger(raw, 1);
  return id === undefined ? undefined : numFmtCodeFor(id, custom);
}

function toFill(
  pattern: string,
  fgColor: Color | undefined,
  bgColor: Color | undefined,
): Fill | undefined {
  // `none` (and an absent patternType) is the absence of a fill; an unrecognised token is dropped the
  // same way, like the border-edge style above, so a foreign pattern we do not model leaves the cell
  // unfilled rather than propagating a token the writer would later re-emit unvalidated.
  if (!isFillPatternType(pattern) || pattern === 'none') return undefined;
  return {
    type: 'pattern',
    pattern,
    ...(fgColor ? {fgColor} : {}),
    ...(bgColor ? {bgColor} : {}),
  };
}

// Copy the numeric <gradientFill> attributes (degree; the path insets) onto a gradient draft, keeping
// only the finite ones so an absent or malformed attribute leaves the field its OOXML default (unset).
function assignGradientNumbers(fill: GradientDraft['fill'], attrs: XmlAttributes): void {
  for (const key of ['degree', 'left', 'right', 'top', 'bottom'] as const) {
    const value = numFinite(attrs[key]);
    if (value !== undefined) fill[key] = value;
  }
}

// An accumulated border with no styled edge and no diagonal direction is the empty default:
// it carries nothing, so it resolves to undefined rather than an all-empty Border object.
function borderToStyle(draft: BorderDraft): Border | undefined {
  const hasEdge = BORDER_EDGE_NAMES.some((edge): boolean => draft[edge] !== undefined);
  if (!hasEdge && draft.diagonalUp === undefined && draft.diagonalDown === undefined)
    return undefined;
  return draft;
}

// Read an <alignment> element's attributes into an Alignment, driven by the same ALIGNMENT_FACETS the
// writer emits from, so neither direction can carry a facet the other does not. Only facets that
// differ from the default are kept: boolean flags honour their parsed value (wrapText="0" is off, so
// it must not fabricate a { wrapText: false } alignment), an unrecognised token (an out-of-enum
// vertical one, say) is dropped rather than trusted into the model, and an element carrying only
// defaults yields undefined rather than an empty alignment object.
function parseAlignment(attrs: XmlAttributes): Alignment | undefined {
  const out: {-readonly [K in keyof Alignment]?: Alignment[K]} = {};
  for (const facet of ALIGNMENT_FACETS) {
    const raw = attrs[facet.key];
    switch (facet.kind) {
      case 'token':
        if (raw !== undefined && raw !== facet.omit && facet.isValid(raw)) {
          assignAlignmentToken(out, facet.key, raw);
        }
        break;
      case 'number': {
        const value = numFinite(raw);
        if (value !== undefined && value !== 0) out[facet.key] = value;
        break;
      }
      case 'flag':
        if (boolStrict(raw)) out[facet.key] = true;
        break;
    }
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

// One token facet at a time, so the write's key type is a single member rather than the whole union
// and `out[key] = value` typechecks: the correlated-key access TypeScript cannot verify when the key
// is a union, the same shape `copyFacet` takes in core/style.ts. The cast restates the guard's own
// proof: `isValid` has already accepted `raw` for this facet's enumeration, which the table cannot
// say in a type because all its token entries share one shape.
function assignAlignmentToken<K extends 'horizontal' | 'vertical'>(
  out: {-readonly [P in keyof Alignment]?: Alignment[P]},
  key: K,
  raw: string,
): void {
  out[key] = raw as Alignment[K];
}

// Read a <protection> element into a Protection. The attributes are this codec's; which of their
// readings carry information is `protectionFrom`'s, shared with the binary reader. `locked` is read
// tri-state because an explicit `locked="1"` must read back as *nothing* rather than as a set flag,
// which is a distinction the boolean the binary reader has cannot make and the attribute can.
function parseProtection(attrs: XmlAttributes): Protection | undefined {
  return protectionFrom({locked: boolTristate(attrs.locked), hidden: boolStrict(attrs.hidden)});
}

// ---------------------------------------------------------------------------------------------
// The preserved sub-tables.
//
// `<dxfs>`, `<indexedColors>`, `<mruColors>` and `<tableStyles>` are read out of styles.xml verbatim
// rather than modelled, because the workbook needs them re-emitted unchanged and has no use for their
// contents: the exact entries (count, order, casing) a source file declared survive a round trip, so
// every `indexed="…"` reference keeps its RGB and every `dxfId` its differential style.
// ---------------------------------------------------------------------------------------------

// Assemble the table from what a capture of `<tableStyles>` yielded: the definitions verbatim (see
// {@link TableStyleTable} for why they stay raw while the two names are decoded), the container's two
// nominated default names, and the namespace declarations the definitions depend on. A file with no
// such block, or with the self-closing `count="0"` container Excel writes when it has only defaults to
// state, yields no definitions and whichever names it did carry.
function buildTableStyleTable(
  stylesXml: string,
  styles: readonly string[],
  containerAttrs: XmlAttributes | undefined,
): TableStyleTable {
  const table: {
    styles: TableStyleDefinition[];
    defaultTableStyle?: string;
    defaultPivotStyle?: string;
    namespaces?: TableStyleNamespace[];
  } = {styles: styles.map((xml) => ({name: tableStyleNameOf(xml), xml}))};
  if (containerAttrs?.defaultTableStyle !== undefined) {
    table.defaultTableStyle = containerAttrs.defaultTableStyle;
  }
  if (containerAttrs?.defaultPivotStyle !== undefined) {
    table.defaultPivotStyle = containerAttrs.defaultPivotStyle;
  }
  const namespaces = fragmentNamespaces(stylesXml, styles);
  if (namespaces.length > 0) table.namespaces = namespaces;
  return table;
}

// A captured `<tableStyle>`'s `name`, from the scanner rather than a pattern over its bytes: decoded,
// and in whichever quotes the file used. Empty when the element declares none, which CT_TableStyle
// requires, so such a fragment is one no table can reach and no authored style can collide with.
function tableStyleNameOf(fragment: string): string {
  for (const {attrs} of openElements(fragment, 'tableStyle')) return attrs.name ?? '';
  return '';
}

// The namespace declarations the verbatim `<tableStyle>` fragments depend on, resolved against the
// stylesheet root that scoped them. Only prefixes a fragment actually uses are carried, so an
// ordinary file (whose fragments use none) adds nothing to the re-emitted root; a prefix a fragment
// uses but the root never declared is skipped, because there is no URI to re-declare it with: the
// source was already unparseable there and inventing a URI would not repair it.
//
// `ignorable` is copied from the source's own `mc:Ignorable` rather than assumed: a prefix the source
// did *not* mark ignorable carries meaning the consumer must not skip, and marking it here would tell
// every consumer to throw that meaning away.
function fragmentNamespaces(
  stylesXml: string,
  fragments: readonly string[],
): TableStyleNamespace[] {
  if (fragments.length === 0) return [];
  const declared = new Map<string, string>();
  const ignorable = new Set<string>();
  for (const {attrs} of openElements(stylesXml, 'styleSheet')) {
    for (const [name, value] of Object.entries(attrs)) {
      if (name.startsWith('xmlns:')) declared.set(name.slice('xmlns:'.length), value);
    }
    for (const prefix of (attrs['mc:Ignorable'] ?? '').split(/\s+/)) {
      if (prefix !== '') ignorable.add(prefix);
    }
    break;
  }
  const used = new Set<string>();
  for (const fragment of fragments) {
    // A prefix appears either on an element (`<p:tag`, `</p:tag`) or on an attribute (` p:attr=`).
    for (const match of fragment.matchAll(/[\s</]([A-Za-z_][\w.-]*):[A-Za-z_]/g)) {
      used.add(match[1] as string);
    }
  }
  return [...used]
    .filter((prefix) => declared.has(prefix))
    .map((prefix) => ({
      prefix,
      uri: declared.get(prefix) as string,
      ignorable: ignorable.has(prefix),
    }));
}

// Every preserved styles sub-table is the verbatim children of one container, so the four of them
// are named once here and {@link parseStyleTable} captures them in a single scan.
const PRESERVED_SUBTREES: SubtreeSelection = new Map([
  ['dxfs', 'dxf'],
  ['indexedColors', 'rgbColor'],
  ['mruColors', 'color'],
  ['tableStyles', 'tableStyle'],
]);
