// Cell hyperlinks: the sheet-level `<hyperlinks>` element, its external relationships, and the
// reader that folds a link back onto its cell's value.
//
// A hyperlink is not stored inside the cell in OOXML: the `<c>` holds only the visible label (a
// normal string value), while a separate `<hyperlink>` child of `<worksheet>` ties an A1 reference
// to a destination. An EXTERNAL destination (a URL) is reached indirectly, through a sheet
// relationship carrying `TargetMode="External"` that the `<hyperlink>` names by `r:id`. An INTERNAL
// destination (a location inside the same workbook, which the author writes as a `#`-prefixed value)
// is held directly in a `location` attribute with NO relationship. Emitting an internal target as an
// external relationship makes a strict consumer resolve both the rel and the location and render the
// destination doubled.

import {tryDecodeRange} from '../../core/address.ts';
import type {Cell} from '../../core/cell.ts';
import {type HyperlinkValue, isHyperlinkValue, isRichTextValue} from '../../core/value.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {type CollectingPass} from '../../xml/xml-read.ts';
import {localName} from '../../xml/xml-scan.ts';
import {escapeAttr, textAttr} from '../../xml/xml.ts';
import {relAttr} from '../opc/namespaces.ts';
import type {RelIdAllocator} from './package-plan.ts';

/** A hyperlink gathered from a sheet for serialisation: the cell it sits on, its target, and an
 * optional tooltip. The visible label is the cell's own value and is serialised as that value.
 *
 * `row`/`col` are the anchor's position, kept so links gathered from separate passes (the sheet's live
 * rows, and the rows the streaming writer already flushed and evicted) can be merged back into the
 * row-major order Excel writes them in. They are not serialised; `ref` is.
 */
export interface CollectedHyperlink {
  readonly ref: string;
  readonly row: number;
  readonly col: number;
  readonly target: string;
  readonly tooltip?: string;
}

/** A hyperlink resolved for serialisation. An external target carries a `relId` (the sheet
 * relationship holding the URL) plus that `target`; an internal target carries a `location` (the
 * in-workbook reference). Exactly one of `relId`/`location` is ever set. */
export interface HyperlinkPlan {
  readonly ref: string;
  readonly relId?: string;
  readonly target?: string;
  readonly location?: string;
  readonly tooltip?: string;
}

/**
 * Gather every hyperlink among a run of cells.
 *
 * Takes the cells rather than the sheet, because the streaming writer has to ask this question of a
 * row at the moment it commits: after that the row's cells are evicted, and a post-hoc walk of
 * `sheet.rows()` finds nothing. Both writers therefore ask the same function, over whatever cells
 * they still hold.
 */
export function collectHyperlinks(cells: Iterable<Cell>): CollectedHyperlink[] {
  const links: CollectedHyperlink[] = [];
  for (const cell of cells) {
    const value = cell.value;
    if (isHyperlinkValue(value)) {
      // A link that spans a range carries its extent in `range`; the anchor cell (this one) is the
      // range's top-left. Emit that extent as `ref` so the clickable area survives, falling back to
      // the single cell for an ordinary link.
      links.push({
        ref: value.range ?? cell.address,
        row: cell.row,
        col: cell.col,
        target: value.hyperlink,
        ...(value.tooltip !== undefined ? {tooltip: value.tooltip} : {}),
      });
    }
  }
  return links;
}

/** Every cell a sheet still holds, row-major: the live half of what a writer must gather. */
export function* liveCells(sheet: Worksheet): Generator<Cell, void, undefined> {
  for (const {cells} of sheet.rows()) yield* cells;
}

/** An external hyperlink: the two fields a `TargetMode="External"` relationship needs, both present. */
export type ExternalHyperlinkPlan = HyperlinkPlan & {
  readonly relId: string;
  readonly target: string;
};

/**
 * Whether a planned hyperlink is the external kind, narrowing it so the relationship writer reads
 * `relId` and `target` as the strings they are.
 *
 * A predicate rather than a filter plus two casts: the filter already proved both fields present, and
 * a cast repeating that proof one line later is a claim the compiler cannot check against the filter
 * it is supposed to be echoing. `planHyperlinks` sets exactly one of `relId`/`location`, so testing
 * either field is testing the kind.
 */
export function isExternalHyperlink(link: HyperlinkPlan): link is ExternalHyperlinkPlan {
  return link.relId !== undefined && link.target !== undefined;
}

/** Split collected links into internal (location, no rel) and external (relationship) forms, drawing
 * each external link's relationship id from the sheet's allocator so external ids follow every other
 * sheet-local relationship in canonical order. An internal ('#'-prefixed) link consumes no id. */
export function planHyperlinks(
  links: readonly CollectedHyperlink[],
  rels: RelIdAllocator,
): HyperlinkPlan[] {
  return links.map((link) => {
    const tooltip = link.tooltip !== undefined ? {tooltip: link.tooltip} : {};
    // A '#'-prefixed target is an internal document location: held verbatim in `location`, with no
    // relationship. Everything else is an external URL reached through a relationship.
    if (link.target.startsWith('#')) {
      return {ref: link.ref, location: link.target.slice(1), ...tooltip};
    }
    return {ref: link.ref, relId: rels.next(), target: link.target, ...tooltip};
  });
}

/** The `<hyperlinks>` element, or '' when the sheet has none. Attribute order follows CT_Hyperlink:
 * `ref`, `r:id`, `location`, `tooltip`. */
export function hyperlinksXml(links: readonly HyperlinkPlan[]): string {
  if (links.length === 0) return '';
  const items = links
    .map((link) => {
      const rid = link.relId !== undefined ? ` r:id="${link.relId}"` : '';
      const location = textAttr('location', link.location);
      const tooltip = textAttr('tooltip', link.tooltip);
      return `<hyperlink ref="${escapeAttr(link.ref)}"${rid}${location}${tooltip}/>`;
    })
    .join('');
  return `<hyperlinks>${items}</hyperlinks>`;
}

/** A hyperlink parsed from a sheet: its cell reference plus whichever of `rid`/`location`/`tooltip`
 * the `<hyperlink>` element carried. */
interface ParsedHyperlink {
  readonly ref: string;
  readonly rid?: string;
  readonly location?: string;
  readonly tooltip?: string;
}

/** A pass gathering every `<hyperlink>` element of a worksheet part, for a caller reading the part
 * alongside its other readers in one parse. */
export function sheetHyperlinkPass(): CollectingPass<ParsedHyperlink[]> {
  const links: ParsedHyperlink[] = [];
  return {
    handlers: {
      onOpen(name, attrs, _selfClosing, scope) {
        if (localName(name) !== 'hyperlink') return;
        const ref = attrs.ref;
        if (ref === undefined) return;
        const rid = relAttr(scope, attrs, 'id');
        links.push({
          ref,
          ...(rid !== undefined ? {rid} : {}),
          ...(attrs.location !== undefined ? {location: attrs.location} : {}),
          ...(attrs.tooltip !== undefined ? {tooltip: attrs.tooltip} : {}),
        });
      },
    },
    result: () => links,
  };
}

/** Fold parsed hyperlinks onto a sheet's cells, wrapping each cell's existing value (its visible
 * label) into a {@link HyperlinkValue}. A link over a cell whose value is not text (a number, a date, a
 * boolean, a formula) is dropped and the value kept. `targetOf` resolves a relationship id to its raw
 * Target: a URL for the external links hyperlinks almost always are, so it must stay unresolved against
 * the package rather than being handed over as a part path. */
export function applyHyperlinks(
  sheet: Worksheet,
  links: readonly ParsedHyperlink[],
  targetOf: (relId: string) => string | undefined,
): void {
  for (const link of links) {
    const target = resolveTarget(link, targetOf);
    if (target === undefined) continue;
    // A hyperlink may span a range (`ref="D1:H1"`); Excel anchors the link at the range's top-left
    // cell. Decode once so a multi-cell link folds onto that anchor rather than asking the sheet for
    // a range address it cannot resolve. A ref that does not name a cell is skipped, not fatal.
    const decoded = tryDecodeRange(link.ref);
    // An unbounded ref (`A:A`) decodes but names no anchor, so it is dropped alongside the garbage:
    // there is no single cell to hang the link on.
    if (decoded === undefined || decoded.tl.col === undefined || decoded.tl.row === undefined)
      continue;
    const cell = sheet.getCell(decoded.tl.address);
    // The visible label is the cell's own value: a plain string, or rich text when the label carried
    // per-run formatting, and an empty cell labels its link with nothing. Any other value (a number, a
    // date, a boolean, a formula, or a link already folded onto the cell) cannot become a label without
    // being destroyed, and the model has nowhere to put a link beside it, so the value stays and the
    // link goes: losing an attribute beats losing content.
    const cellValue = cell.value;
    if (cellValue !== null && typeof cellValue !== 'string' && !isRichTextValue(cellValue))
      continue;
    const text = cellValue ?? '';
    // Record the extent only when the link genuinely spans more than the anchor, so an ordinary
    // single-cell link stays a plain value and the range survives verbatim for a multi-cell one.
    const spansRange = decoded.tl.address !== decoded.br.address;
    const value: HyperlinkValue = {
      hyperlink: target,
      text,
      ...(link.tooltip !== undefined ? {tooltip: link.tooltip} : {}),
      ...(spansRange ? {range: link.ref} : {}),
    };
    cell.value = value;
  }
}

function resolveTarget(
  link: ParsedHyperlink,
  targetOf: (relId: string) => string | undefined,
): string | undefined {
  if (link.rid !== undefined) {
    const base = targetOf(link.rid);
    if (base === undefined) return undefined;
    // A foreign file may split an external URL's fragment into the `location` attribute, apart from
    // the relationship Target; rejoin them so the whole URL survives. Our own writer keeps the
    // fragment in the Target, so a link we wrote never carries both.
    return link.location !== undefined ? `${base}#${link.location}` : base;
  }
  // No relationship: an internal ('#'-prefixed) target held verbatim in `location`.
  return link.location !== undefined ? `#${link.location}` : undefined;
}
