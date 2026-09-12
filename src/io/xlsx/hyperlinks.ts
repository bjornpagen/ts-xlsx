// Worksheet hyperlinks: the sheet-level `<hyperlinks>` element, the relationships an external link
// needs, and the reader that puts each link back on the sheet.
//
// A hyperlink is not stored in a cell. The `<hyperlink>` element names the cells it covers by `ref`,
// and the model holds it the same way, beside the grid. An EXTERNAL destination (a URL) is reached
// indirectly, through a sheet relationship carrying `TargetMode="External"` that the `<hyperlink>`
// names by `r:id`. An INTERNAL destination (a place in the same workbook, which the author writes as a
// `#`-prefixed target) is held directly in a `location` attribute with NO relationship. Emitting an
// internal target as an external relationship makes a strict consumer resolve both the rel and the
// location and render the destination doubled.

import {boundedRect, tryDecodeRange} from '../../core/address.ts';
import type {Hyperlink} from '../../core/hyperlink.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {type CollectingPass} from '../../xml/xml-read.ts';
import {localName} from '../../xml/xml-scan.ts';
import {escapeAttr, textAttr} from '../../xml/xml.ts';
import {relAttr} from '../opc/namespaces.ts';
import type {RelationshipLedger} from './package-plan.ts';
import {REL} from './relationships.ts';

/** A hyperlink resolved for serialisation. An external target carries a `relId` (the sheet
 * relationship holding the URL); an internal target carries a `location` (the in-workbook
 * reference). Exactly one of `relId`/`location` is ever set. */
export interface HyperlinkPlan {
  readonly ref: string;
  readonly relId?: string;
  readonly location?: string;
  readonly tooltip?: string;
}

/**
 * Split a sheet's links into internal (location, no rel) and external (relationship) forms, recording
 * each external link's URL in the sheet's ledger so its relationship follows every other sheet-local
 * relationship in canonical order. An internal (`#`-prefixed) link records nothing.
 *
 * The links keep the order the sheet holds them in. That order is what decides which of two links over
 * one cell `Worksheet.hyperlinkAt` reports, and a re-read adds them in document order, so writing them
 * any other way would change the answer across a save.
 */
export function planHyperlinks(
  links: readonly Hyperlink[],
  rels: RelationshipLedger,
): HyperlinkPlan[] {
  return links.map((link) => {
    const tooltip = link.tooltip !== undefined ? {tooltip: link.tooltip} : {};
    if (link.target.startsWith('#')) {
      return {ref: link.ref, location: link.target.slice(1), ...tooltip};
    }
    return {ref: link.ref, relId: rels.addExternal(REL.hyperlink, link.target), ...tooltip};
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

/**
 * Put parsed hyperlinks on a sheet, in document order, whatever the cells they cover hold. `targetOf`
 * resolves a relationship id to its raw Target: a URL for the external links hyperlinks almost always
 * are, so it must stay unresolved against the package rather than being handed over as a part path.
 */
export function applyHyperlinks(
  sheet: Worksheet,
  links: readonly ParsedHyperlink[],
  targetOf: (relId: string) => string | undefined,
): void {
  for (const link of links) {
    const target = resolveTarget(link, targetOf);
    if (target === undefined) continue;
    // A ref that names no cell, or a whole row or column (`A:A`), gives the link nothing to cover. It is
    // dropped here, at the reader's boundary, so the guard behind `addHyperlink` stays a guard rather
    // than a control-flow path.
    const decoded = tryDecodeRange(link.ref);
    if (decoded === undefined || boundedRect(decoded) === undefined) continue;
    sheet.addHyperlink({
      ref: link.ref,
      target,
      ...(link.tooltip !== undefined ? {tooltip: link.tooltip} : {}),
    });
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
