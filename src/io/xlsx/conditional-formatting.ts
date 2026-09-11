// Serialising conditional formatting to the classic `<conditionalFormatting>` worksheet element.
//
// Each block names its target range(s) in a `sqref` attribute and holds one or more `<cfRule>`
// children. A rule's shape depends on its type: a `dataBar`/`colorScale`/`iconSet` carries a scale
// element (its `<cfvo>` anchors and colours), while a `cellIs`/`expression`/`top10`/… carries its
// operands as `<formula>` children and points at a differential style by `dxfId`. A rule the library
// does not model in depth still round-trips its attributes, so nothing is silently dropped on save.
//
// A data bar's richer facets (its gradient fill, its negative-value fill colour, its axis colour)
// have no home in the classic `<dataBar>` element; Excel stores them only in the 2009 x14 extension.
// So a data-bar rule carrying any of them is written twice: the classic element (its anchors and bar
// colour, understood by every consumer) plus an `<x14:dataBar>` in the worksheet `<extLst>` carrying
// the extras, the two linked by a shared id. The reader folds the extension back onto the classic
// rule, so the gradient flag and the two extra colours survive a round-trip rather than being dropped.

import {
  type CfValueObject,
  type ConditionalFormatting,
  type ConditionalFormattingOperator,
  type ConditionalFormattingRule,
  type ConditionalFormattingType,
  type IconSetType,
  isCfValueObjectType,
  isConditionalFormattingOperator,
  isConditionalFormattingType,
  isIconSetType,
  isCfTimePeriod,
  type CfTimePeriod,
} from '../../core/conditional-formatting.ts';
import {stripFormulaEquals} from '../../core/formula.ts';
import {decodeSqrefRects} from '../../core/merge.ts';
import type {Color} from '../../core/style.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {coerceNumericLiteral, enumToken, numInteger} from '../../xml/xml-attrs.ts';
import {type CollectingPass, type SaxHandlers, TextCapture} from '../../xml/xml-read.ts';
import {boolPresent, boolStrict, boolTristate, localName} from '../../xml/xml-scan.ts';
import {
  boolAttr,
  checkedToken,
  escapeAttr,
  escapeText,
  intAttr,
  numberText,
  textAttr,
} from '../../xml/xml.ts';
import {admitting} from '../read-policy/read-repair.ts';
import {colorAttrs, parseColor} from './color-xml.ts';
// The x14/xm extension namespaces and ext-URI GUIDs are declared inline on the `<ext>` elements
// exactly as Excel writes them, so no worksheet-root xmlns is needed. `CF_EXT_URI` scopes the
// worksheet's x14 conditional formattings; `DATABAR_LINK_EXT_URI` scopes the `<x14:id>` link a
// classic cfRule carries to name its extension.
import {CF_EXT_URI, DATABAR_LINK_EXT_URI, isExtensionElement, XM_NS} from './namespaces.ts';
import type {StyleRegistry} from './styles.ts';
import {x14Ext} from './x14-ext.ts';

// Excel's default data bar when the author supplies none: a min/max anchor pair and its standard blue.
const DEFAULT_DATABAR_CFVO: readonly CfValueObject[] = [{type: 'min'}, {type: 'max'}];
const DEFAULT_DATABAR_COLOR: Color = {argb: 'FF638EC6'};

// A data bar needs the x14 extension only when it carries a facet the classic element cannot express.
// A plain data bar (anchors and bar colour alone) stays classic-only, so an unadorned rule never
// fabricates an empty extension block.
function needsDataBarExt(rule: ConditionalFormattingRule): boolean {
  return (
    rule.gradient !== undefined ||
    rule.negativeFillColor !== undefined ||
    rule.axisColor !== undefined
  );
}

// The synthetic id linking a classic cfRule to its x14 extension. Excel uses a random GUID; any unique
// token that matches on both ends works, so a deterministic per-sheet index keeps the output stable
// and testable.
function dataBarExtGuid(index: number): string {
  return `{00000000-0000-0000-0000-${String(index + 1).padStart(12, '0')}}`;
}

/** Each extended data bar's link id, keyed by rule identity: the one map both passes are handed. */
export type DataBarExtLinks = ReadonlyMap<ConditionalFormattingRule, string>;

// Assign every extended data bar its link id once, keyed by rule identity. The classic pass (which
// emits the `<extLst>` link on the cfRule) and the extension pass (which emits the `<x14:cfRule id>`)
// are handed the SAME map, which is what makes the two ends of a link agree by construction rather
// than by the two passes happening to walk the rules in the same order. Each used to build its own,
// so the agreement was really by determinism, and the comment claiming otherwise described an
// invariant the code did not have. A rule absent from the map needs no extension.
export function dataBarExtLinks(formattings: readonly ConditionalFormatting[]): DataBarExtLinks {
  const links = new Map<ConditionalFormattingRule, string>();
  let index = 0;
  for (const cf of formattings) {
    for (const rule of cf.rules) {
      if (rule.type === 'dataBar' && needsDataBarExt(rule))
        links.set(rule, dataBarExtGuid(index++));
    }
  }
  return links;
}

// The three built-in visual rules. Each renders a built-in visual and carries no differential
// style (so the write side skips the dxf), and each nests its `<color>` children differently (so
// the read side, {@link ScaleKind}, tracks which one it is inside). This tuple drives both.
const SCALE_KINDS = ['dataBar', 'colorScale', 'iconSet'] as const;
const SCALE_TYPES = new Set<string>(SCALE_KINDS);

/**
 * Serialise every conditional formatting on a sheet into its `<conditionalFormatting>` blocks, in
 * insertion order. Returns '' when the sheet has none. A rule missing a `priority` is assigned the
 * next free one, so the output always satisfies Excel's requirement that every cfRule carry one.
 */
export function conditionalFormattingsXml(
  formattings: readonly ConditionalFormatting[],
  styles: StyleRegistry,
  extLinks: DataBarExtLinks,
): string {
  if (formattings.length === 0) return '';
  const priority = {next: 1};
  return formattings.map((cf) => blockXml(cf, styles, priority, extLinks)).join('');
}

/**
 * The worksheet `<extLst>` `<ext>` carrying the x14 data-bar extensions, or '' when no data bar needs
 * one. Each extension's `<x14:cfRule id>` is read from the same {@link dataBarExtLinks} map the classic
 * pass uses for the `<x14:id>` its cfRule carries, so the two ends of a link always match. Emitted bare
 * (no `<extLst>` wrapper) so the worksheet serialiser can gather it into a single `<extLst>` beside the
 * data-validation extension.
 */
export function conditionalFormattingsExtXml(
  formattings: readonly ConditionalFormatting[],
  extLinks: DataBarExtLinks,
): string {
  const items: string[] = [];
  for (const cf of formattings) {
    for (const rule of cf.rules) {
      const guid = extLinks.get(rule);
      if (guid !== undefined) items.push(x14DataBarXml(cf.ref, rule, guid));
    }
  }
  if (items.length === 0) return '';
  return x14Ext(
    CF_EXT_URI,
    `<x14:conditionalFormattings>${items.join('')}</x14:conditionalFormattings>`,
  );
}

// One `<x14:conditionalFormatting>`: an `<x14:cfRule type="dataBar" id>` mirroring the classic anchors
// as `<x14:cfvo>` and adding the facets the classic element cannot carry (gradient, negative-fill and
// axis colours), with the target range in an `<xm:sqref>` child: the shape Excel writes.
function x14DataBarXml(ref: string, rule: ConditionalFormattingRule, guid: string): string {
  const anchors = dataBarAnchors(rule).map(x14Cfvo).join('');
  const gradient = boolAttr('gradient', rule.gradient);
  const negative =
    rule.negativeFillColor !== undefined
      ? `<x14:negativeFillColor ${colorAttrs(rule.negativeFillColor)}/>`
      : '';
  const axis = rule.axisColor !== undefined ? `<x14:axisColor ${colorAttrs(rule.axisColor)}/>` : '';
  return (
    `<x14:conditionalFormatting xmlns:xm="${XM_NS}">` +
    `<x14:cfRule type="dataBar" id="${guid}">` +
    `<x14:dataBar${gradient}>${anchors}${negative}${axis}</x14:dataBar>` +
    `</x14:cfRule><xm:sqref>${escapeText(ref)}</xm:sqref></x14:conditionalFormatting>`
  );
}

/**
 * A data bar's anchors, defaults included.
 *
 * The classic element and its x14 extension describe *one* bar, so they must show the same low and
 * high anchors or Excel repairs the sheet. Both halves used to decide that for themselves, with the
 * same expression written twice: the pairing held by coincidence, which is precisely the coupling
 * ADR-0003 identifies one level up for the link id and fixes there with a shared map.
 *
 * The minimal call (no cfvo at all) gains Excel's own min/max pair rather than an invalid empty
 * element, and it must gain the *same* pair on both sides.
 */
function dataBarAnchors(rule: ConditionalFormattingRule): readonly CfValueObject[] {
  return rule.cfvo && rule.cfvo.length > 0 ? rule.cfvo : DEFAULT_DATABAR_CFVO;
}

/**
 * Write one scale anchor in the classic form or the x14 one.
 *
 * A `min`/`max` anchor carries no value and self-closes; every other kind states its value, and the
 * two forms state it in different places. The classic element uses a `val` attribute; the x14
 * extension stores every anchor value as an `<xm:f>` formula child. So the escape has to follow the
 * form -- an attribute value and element text are not escaped alike -- and two separate writers is
 * where that stops being true.
 *
 * Curried rather than taking the form as a second parameter, because every caller is a `.map`, and
 * `.map(cfvoXml)` would quietly hand the array index in as the form.
 */
function cfvoWriter(form: 'classic' | 'x14'): (cfvo: CfValueObject) => string {
  return (cfvo) => {
    const type = checkedToken(cfvo.type, isCfValueObjectType, 'conditional format value type');
    const tag = form === 'classic' ? 'cfvo' : 'x14:cfvo';
    if (cfvo.value === undefined) return `<${tag} type="${type}"/>`;
    // A numeric anchor goes through the number check: `String(NaN)` wrote `val="NaN"`.
    const value = typeof cfvo.value === 'number' ? numberText(cfvo.value) : cfvo.value;
    return form === 'classic'
      ? `<${tag} type="${type}"${textAttr('val', value)}/>`
      : `<${tag} type="${type}"><xm:f>${escapeText(value)}</xm:f></${tag}>`;
  };
}

const cfvoXml = cfvoWriter('classic');
const x14Cfvo = cfvoWriter('x14');

// The `<extLst>` a classic data-bar cfRule carries to name its x14 extension by shared id.
function cfRuleExtLinkXml(guid: string): string {
  return `<extLst>${x14Ext(DATABAR_LINK_EXT_URI, `<x14:id>${guid}</x14:id>`)}</extLst>`;
}

function blockXml(
  cf: ConditionalFormatting,
  styles: StyleRegistry,
  priority: {next: number},
  extLinks: DataBarExtLinks,
): string {
  const rules = cf.rules.map((rule) => ruleXml(rule, styles, priority, extLinks)).join('');
  // `CT_ConditionalFormatting` requires at least one `<cfRule>`, so a block with none is omitted
  // rather than emitted empty. A caller can author one, and the reader produces one when every rule
  // in a foreign block named a type the enumeration does not allow.
  if (rules === '') return '';
  return `<conditionalFormatting sqref="${escapeAttr(cf.ref)}">${rules}</conditionalFormatting>`;
}

function ruleXml(
  rule: ConditionalFormattingRule,
  styles: StyleRegistry,
  priority: {next: number},
  extLinks: DataBarExtLinks,
): string {
  const p = rule.priority ?? priority.next;
  // Checked before it feeds the counter: an authored `NaN` spread through `Math.max` into every
  // auto-assigned priority after it, so one bad rule wrote `priority="NaN"` on all that followed.
  const priorityAttr = intAttr('priority', p).trim();
  // Keep the running counter ahead of any explicit priority so later auto-assigned ones stay unique.
  priority.next = Math.max(priority.next, p) + 1;

  const attrs = [
    `type="${checkedToken(rule.type, isConditionalFormattingType, 'conditional formatting type')}"`,
  ];
  const dxfId = resolveDxfId(rule, styles);
  if (dxfId !== undefined) attrs.push(`dxfId="${dxfId}"`);
  attrs.push(priorityAttr);
  if (rule.stopIfTrue) attrs.push('stopIfTrue="1"');
  if (rule.aboveAverage === false) attrs.push('aboveAverage="0"');
  if (rule.equalAverage) attrs.push('equalAverage="1"');
  if (rule.bottom) attrs.push('bottom="1"');
  if (rule.percent) attrs.push('percent="1"');
  if (rule.operator !== undefined) {
    attrs.push(
      `operator="${checkedToken(rule.operator, isConditionalFormattingOperator, 'conditional formatting operator')}"`,
    );
  }
  if (rule.text !== undefined) attrs.push(`text="${escapeAttr(rule.text)}"`);
  if (rule.timePeriod !== undefined) {
    attrs.push(`timePeriod="${checkedToken(rule.timePeriod, isCfTimePeriod, 'time period')}"`);
  }
  if (rule.rank !== undefined) attrs.push(intAttr('rank', rule.rank, 0).trim());
  if (rule.stdDev !== undefined) attrs.push(intAttr('stdDev', rule.stdDev).trim());

  let body = SCALE_TYPES.has(rule.type) ? scaleXml(rule) : formulaeXml(rule.formulae);
  // A data bar with x14-only facets links to its extension by the id assigned in dataBarExtLinks; the
  // extension itself rides in the worksheet <extLst>. The link is the cfRule's last child, after the
  // dataBar. A rule absent from the map carries no extension.
  const extGuid = extLinks.get(rule);
  if (extGuid !== undefined) body += cfRuleExtLinkXml(extGuid);
  return body === ''
    ? `<cfRule ${attrs.join(' ')}/>`
    : `<cfRule ${attrs.join(' ')}>${body}</cfRule>`;
}

// A rule points at a differential style either by a preserved index read from a file (`dxfId`) or by
// a style authored on the rule (interned here). The preserved index wins: it references the original
// file's dxf table, which the writer re-emits verbatim. Scale rules never carry one.
function resolveDxfId(rule: ConditionalFormattingRule, styles: StyleRegistry): number | undefined {
  if (SCALE_TYPES.has(rule.type)) return undefined;
  // `dxfId` is a public `string`, so an authored rule can carry anything. It is read through the
  // same predicate the reader uses, so a value the reader would have refused is not one an author
  // can smuggle in: without this the rule emits `dxfId="NaN"` into styles-referencing XML.
  const preserved = dxfIndex(rule.dxfId);
  if (preserved !== undefined) return preserved;
  if (rule.style !== undefined) return styles.differentialStyleId(rule.style);
  return undefined;
}

function formulaeXml(formulae: readonly (string | number)[] | undefined): string {
  if (formulae === undefined) return '';
  return formulae.map((f) => `<formula>${escapeText(stripFormulaEquals(f))}</formula>`).join('');
}

function scaleXml(rule: ConditionalFormattingRule): string {
  if (rule.type === 'dataBar') return dataBarXml(rule);
  if (rule.type === 'colorScale') return colorScaleXml(rule);
  return iconSetXml(rule);
}

// A data bar states its low and high anchors and its bar colour. The minimal call (no cfvo, no colour)
// gains Excel's own defaults, a min/max anchor pair and the standard blue, rather than an invalid
// empty element. The gradient flag and the negative-fill/axis colours have no home in this classic
// element; they ride in the x14 extension (see {@link conditionalFormattingsExtXml}), linked from the
// cfRule that wraps this by a shared id.
function dataBarXml(rule: ConditionalFormattingRule): string {
  const color = rule.color ?? DEFAULT_DATABAR_COLOR;
  const anchors = dataBarAnchors(rule).map(cfvoXml).join('');
  return `<dataBar>${anchors}<color ${colorAttrs(color)}/></dataBar>`;
}

// A colour scale pairs each anchor with a colour; a missing colour list falls back to none, still a
// well-formed (if plain) element.
function colorScaleXml(rule: ConditionalFormattingRule): string {
  const anchors = (rule.cfvo ?? []).map(cfvoXml).join('');
  const colors = (rule.colors ?? []).map((c) => `<color ${colorAttrs(c)}/>`).join('');
  return `<colorScale>${anchors}${colors}</colorScale>`;
}

function iconSetXml(rule: ConditionalFormattingRule): string {
  const name =
    rule.iconSet === undefined
      ? ''
      : ` iconSet="${checkedToken(rule.iconSet, isIconSetType, 'icon set')}"`;
  const anchors = (rule.cfvo ?? []).map(cfvoXml).join('');
  return `<iconSet${name}>${anchors}</iconSet>`;
}

// Which scale element a parsed `<color>` belongs to: a data bar names one bar colour, a colour scale
// a colour per anchor. Tracking which element we are inside routes a parsed `<color>` to the right slot.
type ScaleKind = (typeof SCALE_KINDS)[number];

// A rule under construction: fields accumulate across the cfRule's attributes and children, then are
// finalised into a ConditionalFormattingRule on the closing tag. The array/collection fields are
// always present here (empty until filled) and pruned to `undefined` when empty at finalisation.
interface RuleDraft {
  type: ConditionalFormattingType | undefined;
  priority: number | undefined;
  stopIfTrue: boolean;
  operator: ConditionalFormattingOperator | undefined;
  text: string | undefined;
  timePeriod: CfTimePeriod | undefined;
  rank: number | undefined;
  stdDev: number | undefined;
  percent: boolean;
  bottom: boolean;
  aboveAverage: boolean | undefined;
  equalAverage: boolean;
  dxfId: string | undefined;
  iconSet: IconSetType | undefined;
  formulae: (string | number)[];
  cfvo: CfValueObject[];
  colors: Color[];
  color: Color | undefined;
  // The `<x14:id>` a data-bar cfRule carries to name its extension. Transient: it links this rule to
  // its `<x14:dataBar>` during parsing and is dropped once the extension's facets are folded in.
  x14Id: string | undefined;
}

// The facets an `<x14:dataBar>` adds over the classic element, gathered by the id its `<x14:cfRule>`
// carries so they can be matched to the classic rule that links to the same id.
interface DataBarExt {
  gradient: boolean | undefined;
  negativeFillColor: Color | undefined;
  axisColor: Color | undefined;
}

/**
 * A pass reading a worksheet's conditional formatting into the model. The classic `<conditionalFormatting>`
 * blocks supply every rule; the x14 extension (`<x14:conditionalFormatting>` inside `<extLst>`) is
 * read only to enrich a classic data bar with the facets the classic element cannot carry (the
 * gradient flag and the negative-fill and axis colours) matched by the shared id the two ends link
 * on. An extension rule with no classic counterpart (a rule that lives only in x14) is ignored, so it
 * is never half-read into a broken classic rule.
 */
export function conditionalFormattingPass(): CollectingPass<ConditionalFormatting[]> {
  const blocks: ConditionalFormatting[] = [];
  let block: ConditionalFormatting | undefined;
  let draft: RuleDraft | undefined;
  let scale: ScaleKind | undefined;
  const formulaCapture = new TextCapture('formula');

  // Classic data-bar rules that named an extension, paired with the id they linked on, plus the
  // extensions gathered from the worksheet <extLst>. The two are married after the pass: the
  // extension always follows the classic blocks in document order, so it is known by then.
  const linked: {rule: ConditionalFormattingRule; id: string}[] = [];
  const extById = new Map<string, DataBarExt>();
  let x14Ext: DataBarExt | undefined;
  let x14ExtId: string | undefined;
  const x14IdCapture = new TextCapture('id');

  const handlers: SaxHandlers = {
    onOpen(name, attrs, selfClosing, scope) {
      const ln = localName(name);
      // In the x14 extension namespace, not merely prefixed. The prefix test was true of every
      // element in a worksheet that binds the MAIN namespace to a prefix, which is legal and which
      // real toolchains emit, so every conditional format in such a file was read as an extension
      // element and dropped.
      if (isExtensionElement(scope, name)) {
        // The `<x14:id>` a classic data bar carries to name its extension: capture its text into the
        // open draft. The rest are the worksheet extension's own elements.
        if (ln === 'id' && draft !== undefined) {
          x14IdCapture.open(ln, selfClosing);
        } else if (ln === 'cfRule') {
          x14Ext = attrs.type === 'dataBar' && attrs.id !== undefined ? emptyExt() : undefined;
          x14ExtId = attrs.id;
        } else if (x14Ext !== undefined && ln === 'dataBar') {
          // gradient defaults to true in the x14 schema, so an absent attribute reads as a gradient.
          x14Ext.gradient = boolPresent(attrs.gradient);
        } else if (x14Ext !== undefined && ln === 'negativeFillColor') {
          x14Ext.negativeFillColor = parseColor(attrs);
        } else if (x14Ext !== undefined && ln === 'axisColor') {
          x14Ext.axisColor = parseColor(attrs);
        }
        return;
      }
      if (ln === 'conditionalFormatting') {
        block = {ref: attrs.sqref ?? '', rules: []};
      } else if (ln === 'cfRule' && block !== undefined) {
        // A rule with no operands (e.g. duplicateValues) is a self-closing element that fires no
        // close event, so it must be finalised here; one with children waits for its </cfRule>.
        if (selfClosing) {
          const rule = finalizeRule(newDraft(attrs));
          if (rule !== undefined) block.rules.push(rule);
        } else {
          draft = newDraft(attrs);
          scale = undefined;
        }
      } else if (
        draft !== undefined &&
        (ln === 'dataBar' || ln === 'colorScale' || ln === 'iconSet')
      ) {
        scale = ln;
        if (ln === 'iconSet') draft.iconSet = enumToken(attrs.iconSet, isIconSetType);
      } else if (draft !== undefined && ln === 'cfvo') {
        draft.cfvo.push(parseCfvo(attrs));
      } else if (draft !== undefined && ln === 'color') {
        const color = parseColor(attrs);
        if (scale === 'dataBar') draft.color = color;
        else draft.colors.push(color);
      } else if (draft !== undefined && ln === 'formula') {
        formulaCapture.open(ln, selfClosing);
      }
    },
    onText(chunk) {
      formulaCapture.text(chunk);
      x14IdCapture.text(chunk);
    },
    onClose(name, scope) {
      const ln = localName(name);
      if (isExtensionElement(scope, name)) {
        const id = x14IdCapture.close(ln);
        if (id !== undefined) {
          if (draft !== undefined) draft.x14Id = id;
        } else if (ln === 'cfRule' && x14Ext !== undefined && x14ExtId !== undefined) {
          extById.set(x14ExtId, x14Ext);
          x14Ext = undefined;
          x14ExtId = undefined;
        }
        return;
      }
      const formula = formulaCapture.close(ln);
      if (formula !== undefined) {
        if (draft !== undefined) draft.formulae.push(coerceNumericLiteral(formula));
      } else if (ln === 'dataBar' || ln === 'colorScale' || ln === 'iconSet') {
        scale = undefined;
      } else if (ln === 'cfRule' && draft !== undefined) {
        const rule = finalizeRule(draft);
        if (rule !== undefined) {
          if (block !== undefined) block.rules.push(rule);
          if (draft.x14Id !== undefined) linked.push({rule, id: draft.x14Id});
        }
        draft = undefined;
      } else if (ln === 'conditionalFormatting' && block !== undefined) {
        blocks.push(block);
        block = undefined;
      }
    },
  };

  // The marrying step is the pass's result rather than part of the pass, which is what lets it share
  // a parse: the extension always follows the classic blocks in document order, so both ends are in
  // hand by the time the parse ends, whoever else was reading alongside.
  const result = (): ConditionalFormatting[] => {
    for (const {rule, id} of linked) {
      const ext = extById.get(id);
      if (ext === undefined) continue;
      if (ext.gradient !== undefined) rule.gradient = ext.gradient;
      if (ext.negativeFillColor !== undefined) rule.negativeFillColor = ext.negativeFillColor;
      if (ext.axisColor !== undefined) rule.axisColor = ext.axisColor;
    }
    return blocks;
  };
  return {handlers, result};
}

/** Fold parsed conditional formattings onto a sheet, each bound to its original range. */
export function applyConditionalFormattings(
  sheet: Worksheet,
  blocks: readonly ConditionalFormatting[],
): void {
  for (const block of blocks) {
    // A `sqref` no area of which decodes, an absent one included, names no cells to format, and
    // re-emitting it would put the file's own unreadable text back on the wire. Dropped here, at the
    // reader's boundary, so the authoring guard behind `addConditionalFormatting` stays a guard
    // rather than a control-flow path.
    if (decodeSqrefRects(block.ref).length === 0) continue;
    admitting(() => {
      sheet.addConditionalFormatting(block);
    });
  }
}

function emptyExt(): DataBarExt {
  return {gradient: undefined, negativeFillColor: undefined, axisColor: undefined};
}

function newDraft(attrs: Record<string, string>): RuleDraft {
  return {
    type: enumToken(attrs.type, isConditionalFormattingType),
    // `priority` and `stdDev` are `xsd:int` and `rank` is `xsd:unsignedInt`, so a fraction or a negative
    // rank is not a value the attribute can hold, and keeping one wrote it straight back.
    priority: numInteger(attrs.priority),
    stopIfTrue: boolStrict(attrs.stopIfTrue),
    operator: enumToken(attrs.operator, isConditionalFormattingOperator),
    text: attrs.text,
    timePeriod: enumToken(attrs.timePeriod, isCfTimePeriod),
    rank: numInteger(attrs.rank, 0),
    stdDev: numInteger(attrs.stdDev),
    percent: boolStrict(attrs.percent),
    bottom: boolStrict(attrs.bottom),
    // aboveAverage defaults to true in OOXML; only an explicit "0"/"false" means below-average, and an
    // unrecognised token is dropped, which leaves that default in force.
    aboveAverage: boolTristate(attrs.aboveAverage),
    equalAverage: boolStrict(attrs.equalAverage),
    dxfId: parseIndexAttr(attrs.dxfId),
    iconSet: undefined,
    formulae: [],
    cfvo: [],
    colors: [],
    color: undefined,
    x14Id: undefined,
  };
}

// The one reading of a `dxfId`, shared by the reader that preserves one and the writer that emits
// one, so the two cannot come to disagree about what a usable index is: a slot in the dxf table is a
// non-negative integer or it is nothing.
function dxfIndex(value: string | undefined): number | undefined {
  return numInteger(value, 0);
}

// dxfId is preserved as the raw string (not renumbered) so it keeps pointing at the same slot in the
// dxf table on re-write; a malformed value is dropped rather than later coercing to `dxfId="NaN"`.
function parseIndexAttr(value: string | undefined): string | undefined {
  return dxfIndex(value) === undefined ? undefined : value;
}

// A rule whose `type` the enumeration does not allow is dropped whole rather than half-read: `type`
// is the one attribute `<cfRule>` requires, every other field is read relative to it, and a rule the
// writer would refuse to emit is not one worth building a model for.
function finalizeRule(draft: RuleDraft): ConditionalFormattingRule | undefined {
  if (draft.type === undefined) return undefined;
  const rule: ConditionalFormattingRule = {type: draft.type};
  if (draft.priority !== undefined) rule.priority = draft.priority;
  if (draft.stopIfTrue) rule.stopIfTrue = true;
  if (draft.operator !== undefined) rule.operator = draft.operator;
  if (draft.text !== undefined) rule.text = draft.text;
  if (draft.timePeriod !== undefined) rule.timePeriod = draft.timePeriod;
  if (draft.rank !== undefined) rule.rank = draft.rank;
  if (draft.stdDev !== undefined) rule.stdDev = draft.stdDev;
  if (draft.percent) rule.percent = true;
  if (draft.bottom) rule.bottom = true;
  if (draft.aboveAverage !== undefined) rule.aboveAverage = draft.aboveAverage;
  if (draft.equalAverage) rule.equalAverage = true;
  if (draft.dxfId !== undefined) rule.dxfId = draft.dxfId;
  if (draft.iconSet !== undefined) rule.iconSet = draft.iconSet;
  if (draft.formulae.length > 0) rule.formulae = draft.formulae;
  if (draft.cfvo.length > 0) rule.cfvo = draft.cfvo;
  if (draft.colors.length > 0) rule.colors = draft.colors;
  if (draft.color !== undefined) rule.color = draft.color;
  return rule;
}

function parseCfvo(attrs: Record<string, string>): CfValueObject {
  // An anchor whose type is absent reads as `num`, the schema's default; an anchor whose type is
  // unrecognised takes the same fallback rather than being dropped, because an anchor missing from a
  // scale would leave the rule with fewer than the anchors its type requires.
  const raw = attrs.type;
  const type = raw !== undefined && isCfValueObjectType(raw) ? raw : 'num';
  const cfvo: CfValueObject = {type};
  if (attrs.val !== undefined) {
    // A `formula` anchor's value is an expression and stays a string; the rest are numeric.
    cfvo.value = type === 'formula' ? attrs.val : coerceNumericLiteral(attrs.val);
  }
  return cfvo;
}
