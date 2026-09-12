// Serialising conditional formatting to a worksheet, in both of the forms Excel stores it in.
//
// The classic form is the `<conditionalFormatting>` element. Each block names its target range(s) in a
// `sqref` attribute and holds one or more `<cfRule>` children. A rule's shape depends on its type: a
// `dataBar`/`colorScale`/`iconSet` carries a scale element (its `<cfvo>` anchors and colours), while a
// `cellIs`/`expression`/`top10`/… carries its operands as `<formula>` children and points at a
// differential style by `dxfId`. A rule the library does not model in depth still round-trips its
// attributes, so nothing is silently dropped on save.
//
// The extension form is `<x14:conditionalFormatting>` in the worksheet `<extLst>`, the 2009 schema's
// carrier for what the classic element cannot spell or Excel 2007 must not see: a rule whose formula
// reaches another sheet, a 2009 icon family, custom icons. It keeps the target in an `<xm:sqref>` child,
// every operand and anchor value in an `<xm:f>`, and the differential style inline as an `<x14:dxf>`
// rather than by index. A set read from it is marked `extended` so it is written back there, and its
// inline style is adopted into the workbook's table so the rule holds a `dxfId` as a classic one does.
//
// A data bar's 2010 facets (its gradient, border, direction, axis, negative-bar colours and automatic
// anchors) have no home in the classic `<dataBar>` element either. A classic data-bar rule carrying any
// of them is written twice: the classic element (its anchors and bar colour, understood by every
// consumer) plus an `<x14:dataBar>` in the extension carrying the rest, the two linked by a shared id.
// The reader folds the extension back onto the classic rule.

import {
  type CfIcon,
  type CfTimePeriod,
  type CfValueObject,
  type CfValueObjectType,
  type ConditionalFormatting,
  type ConditionalFormattingOperator,
  type ConditionalFormattingRule,
  type ConditionalFormattingType,
  type IconSetType,
  isCfIconSetType,
  isCfTimePeriod,
  isCfValueObjectType,
  isConditionalFormattingOperator,
  isConditionalFormattingType,
  isDataBarAxisPosition,
  isDataBarDirection,
  isIconSetType,
  ruleNeedsExtension,
} from '../../core/conditional-formatting.ts';
import {mangleFormula, stripFormulaEquals, unmangleFunctions} from '../../core/formula.ts';
import {decodeSqrefRects} from '../../core/merge.ts';
import type {Color} from '../../core/style.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {InternalError} from '../../errors.ts';
import {coerceNumericLiteral, enumToken, numInteger} from '../../xml/xml-attrs.ts';
import type {NamespaceScope} from '../../xml/xml-namespaces.ts';
import {
  type CollectingPass,
  elementRange,
  type SaxHandlers,
  TextCapture,
} from '../../xml/xml-read.ts';
import {boolStrict, boolTristate, localName} from '../../xml/xml-scan.ts';
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
import {
  CF_EXT_URI,
  DATABAR_LINK_EXT_URI,
  isExtensionElement,
  isMainNamespaceElement,
  XM_NS,
} from './namespaces.ts';
import type {StyleRegistry} from './styles.ts';
import {x14Ext} from './x14-ext.ts';

// Excel's default data bar when the author supplies none: a min/max anchor pair and its standard blue.
const DEFAULT_DATABAR_CFVO: readonly CfValueObject[] = [{type: 'min'}, {type: 'max'}];
const DEFAULT_DATABAR_COLOR: Color = {argb: 'FF638EC6'};

// The anchors the classic form has no spelling for, and the one it states in their place, as Excel's
// own save does.
const CLASSIC_ANCHOR: Partial<Record<CfValueObjectType, CfValueObjectType>> = {
  autoMin: 'min',
  autoMax: 'max',
};

// The facets of a data bar only its `<x14:dataBar>` carries, gathered as read. Its lengths are here too,
// though the classic element can carry them, because a bar with an extension keeps them in the
// extension, as Excel writes it.
type DataBarFacets = Pick<
  ConditionalFormattingRule,
  | 'minLength'
  | 'maxLength'
  | 'gradient'
  | 'border'
  | 'borderColor'
  | 'direction'
  | 'negativeFillColor'
  | 'negativeBorderColor'
  | 'negativeBarColorSameAsPositive'
  | 'negativeBarBorderColorSameAsPositive'
  | 'axisPosition'
  | 'axisColor'
>;

// A data bar needs the x14 extension only when it carries a facet the classic element cannot express.
// Its lengths are not one: Excel keeps a bar that differs from a plain one only in its lengths classic,
// with the lengths on the classic element. A plain data bar stays classic-only, so an unadorned rule
// never fabricates an empty extension block.
function needsDataBarExt(rule: ConditionalFormattingRule): boolean {
  return (
    rule.gradient !== undefined ||
    rule.border !== undefined ||
    rule.borderColor !== undefined ||
    rule.direction !== undefined ||
    rule.negativeFillColor !== undefined ||
    rule.negativeBorderColor !== undefined ||
    rule.negativeBarColorSameAsPositive !== undefined ||
    rule.negativeBarBorderColorSameAsPositive !== undefined ||
    rule.axisPosition !== undefined ||
    rule.axisColor !== undefined ||
    (rule.cfvo ?? []).some((anchor) => CLASSIC_ANCHOR[anchor.type] !== undefined)
  );
}

// Whether a rule is written wholly in the extension form: its set was read from there, or the rule
// says something only that form can.
function inExtension(cf: ConditionalFormatting, rule: ConditionalFormattingRule): boolean {
  return cf.extended === true || ruleNeedsExtension(rule);
}

// The synthetic id an `<x14:cfRule>` carries. Excel uses a random GUID; any token unique in the sheet
// works, so a deterministic per-sheet index keeps the output stable and testable.
function extensionGuid(index: number): string {
  return `{00000000-0000-0000-0000-${String(index + 1).padStart(12, '0')}}`;
}

/**
 * What the classic pass and the extension pass must agree on about a sheet's conditional formats,
 * decided once and handed to both. A classic data bar's link and the `<x14:cfRule>` it names have to
 * carry one id, and two rules in different forms must not be numbered alike, and each pass deriving its
 * half itself makes both agreements hold by the two walks happening to visit the rules in one order.
 */
export interface ConditionalFormattingPlan {
  /** Every rule's priority: its own, or the next one free after every priority before it. */
  readonly priorities: ReadonlyMap<ConditionalFormattingRule, number>;
  /**
   * The id of every rule with an `<x14:cfRule>`: a rule written wholly in the extension, and a classic
   * data bar whose extra facets ride there. A rule absent from the map needs no extension.
   */
  readonly extensionIds: ReadonlyMap<ConditionalFormattingRule, string>;
}

/** Plan a sheet's conditional formats, in insertion order, for both serialisers. */
export function planConditionalFormatting(
  formattings: readonly ConditionalFormatting[],
): ConditionalFormattingPlan {
  const priorities = new Map<ConditionalFormattingRule, number>();
  const extensionIds = new Map<ConditionalFormattingRule, string>();
  let next = 1;
  for (const cf of formattings) {
    for (const rule of cf.rules) {
      const priority = rule.priority ?? next;
      // Checked before it feeds the counter: an authored `NaN` spread through `Math.max` into every
      // auto-assigned priority after it, so one bad rule wrote `priority="NaN"` on all that followed.
      intAttr('priority', priority);
      // Kept ahead of any explicit priority so later auto-assigned ones stay unique.
      next = Math.max(next, priority) + 1;
      priorities.set(rule, priority);
      if (inExtension(cf, rule) || (rule.type === 'dataBar' && needsDataBarExt(rule))) {
        extensionIds.set(rule, extensionGuid(extensionIds.size));
      }
    }
  }
  return {priorities, extensionIds};
}

function priorityOf(plan: ConditionalFormattingPlan, rule: ConditionalFormattingRule): number {
  const priority = plan.priorities.get(rule);
  if (priority === undefined) {
    throw new InternalError('a conditional format rule reached a serialiser unplanned');
  }
  return priority;
}

// The three built-in visual rules. Each renders a built-in visual and carries no differential
// style (so the write side skips the dxf), and each nests its `<color>` children differently (so
// the read side, {@link ScaleKind}, tracks which one it is inside). This tuple drives both.
const SCALE_KINDS = ['dataBar', 'colorScale', 'iconSet'] as const;
const SCALE_TYPES = new Set<string>(SCALE_KINDS);

/**
 * Serialise the classic `<conditionalFormatting>` blocks of a sheet, in insertion order, leaving out
 * every rule {@link conditionalFormattingsExtXml} writes in the extension form. Returns '' when no rule
 * remains.
 *
 * A rule's formulas, its operands and a `formula` anchor alike, are stored with the function prefixes
 * a cell formula takes, as Excel stores them; `formulaNames` is `formulaNamesInScope` for the sheet.
 */
export function conditionalFormattingsXml(
  formattings: readonly ConditionalFormatting[],
  styles: StyleRegistry,
  plan: ConditionalFormattingPlan,
  formulaNames: ReadonlySet<string>,
): string {
  return formattings.map((cf) => blockXml(cf, styles, plan, formulaNames)).join('');
}

/**
 * The worksheet `<extLst>` `<ext>` carrying the sheet's extension-form conditional formats, or '' when
 * it needs none: every rule written wholly there, grouped under its set's range, and every classic data
 * bar's extra facets, under the id the plan gave the link its cfRule carries. Emitted bare (no
 * `<extLst>` wrapper) so the worksheet serialiser can gather it into a single `<extLst>` beside the
 * data-validation extension.
 */
export function conditionalFormattingsExtXml(
  formattings: readonly ConditionalFormatting[],
  styles: StyleRegistry,
  plan: ConditionalFormattingPlan,
  formulaNames: ReadonlySet<string>,
): string {
  const items: string[] = [];
  for (const cf of formattings) {
    const whole = cf.rules.filter((rule) => inExtension(cf, rule));
    if (whole.length > 0) {
      const rules = whole.map((rule) => extensionRuleXml(rule, styles, plan, formulaNames));
      items.push(extensionBlockXml(cf.ref, rules.join('')));
    }
    for (const rule of cf.rules) {
      const guid = plan.extensionIds.get(rule);
      if (guid === undefined || inExtension(cf, rule)) continue;
      items.push(
        extensionBlockXml(
          cf.ref,
          `<x14:cfRule type="dataBar" id="${guid}">${x14DataBarXml(rule, formulaNames, false)}</x14:cfRule>`,
        ),
      );
    }
  }
  if (items.length === 0) return '';
  return x14Ext(
    CF_EXT_URI,
    `<x14:conditionalFormattings>${items.join('')}</x14:conditionalFormattings>`,
  );
}

// One `<x14:conditionalFormatting>`: its rules, then the target range in an `<xm:sqref>` child, the
// shape Excel writes.
function extensionBlockXml(ref: string, rules: string): string {
  return (
    `<x14:conditionalFormatting xmlns:xm="${XM_NS}">${rules}` +
    `<xm:sqref>${escapeText(ref)}</xm:sqref></x14:conditionalFormatting>`
  );
}

// A rule written wholly in the extension: the attributes a classic cfRule carries less the `dxfId` this
// form has no use for, then its id; its operands as `<xm:f>`, its scale, and its style inline, in the
// order `x14:CT_CfRule` sequences them.
function extensionRuleXml(
  rule: ConditionalFormattingRule,
  styles: StyleRegistry,
  plan: ConditionalFormattingPlan,
  formulaNames: ReadonlySet<string>,
): string {
  const attrs = [typeAttr(rule), `priority="${priorityOf(plan, rule)}"`, ...ruleFlagAttrs(rule)];
  const id = plan.extensionIds.get(rule);
  if (id !== undefined) attrs.push(`id="${id}"`);
  const body =
    (SCALE_TYPES.has(rule.type)
      ? extensionScaleXml(rule, formulaNames)
      : operandsXml(rule.formulae, 'xm:f', formulaNames)) + inlineDxfXml(rule, styles);
  return body === ''
    ? `<x14:cfRule ${attrs.join(' ')}/>`
    : `<x14:cfRule ${attrs.join(' ')}>${body}</x14:cfRule>`;
}

// The extension form names no `dxfId`, so the style goes inline as the table entry the rule reaches,
// which carries a preserved style and an authored one alike. An authored style still lands in the
// table on the way, as Excel's own save puts every inline style there too.
function inlineDxfXml(rule: ConditionalFormattingRule, styles: StyleRegistry): string {
  const id = resolveDxfId(rule, styles);
  const fragment = id === undefined ? undefined : styles.differentialStyleFragment(id);
  const range = fragment === undefined ? undefined : elementRange(fragment, ['dxf']);
  if (fragment === undefined || range === undefined) return '';
  return `<x14:dxf>${fragment.slice(range.contentStart, range.contentEnd)}</x14:dxf>`;
}

function extensionScaleXml(
  rule: ConditionalFormattingRule,
  formulaNames: ReadonlySet<string>,
): string {
  if (rule.type === 'dataBar') return x14DataBarXml(rule, formulaNames, true);
  const cfvoXml = cfvoWriter('x14', formulaNames);
  if (rule.type === 'colorScale') {
    const anchors = (rule.cfvo ?? []).map(cfvoXml).join('');
    const colors = (rule.colors ?? []).map((c) => `<x14:color ${colorAttrs(c)}/>`).join('');
    return `<x14:colorScale>${anchors}${colors}</x14:colorScale>`;
  }
  return iconSetXml(rule, cfvoXml, 'x14');
}

// An `<x14:dataBar>`: the anchors mirrored as `<x14:cfvo>`, and the facets the classic element cannot
// carry, attributes and colours alike in the order `x14:CT_DataBar` sequences them. A bar written wholly
// in the extension has no classic element to hold its colour or its hidden value, so it states them
// here too.
function x14DataBarXml(
  rule: ConditionalFormattingRule,
  formulaNames: ReadonlySet<string>,
  whole: boolean,
): string {
  const attrs =
    intAttr('minLength', rule.minLength, 0) +
    intAttr('maxLength', rule.maxLength, 0) +
    (whole && rule.showValue === false ? ' showValue="0"' : '') +
    boolAttr('border', rule.border) +
    boolAttr('gradient', rule.gradient) +
    tokenAttr('direction', rule.direction, isDataBarDirection, 'data bar direction') +
    boolAttr('negativeBarColorSameAsPositive', rule.negativeBarColorSameAsPositive) +
    boolAttr('negativeBarBorderColorSameAsPositive', rule.negativeBarBorderColorSameAsPositive) +
    tokenAttr('axisPosition', rule.axisPosition, isDataBarAxisPosition, 'data bar axis position');
  const children =
    dataBarAnchors(rule).map(cfvoWriter('x14', formulaNames)).join('') +
    (whole ? x14ColorXml('fillColor', rule.color ?? DEFAULT_DATABAR_COLOR) : '') +
    x14ColorXml('borderColor', rule.borderColor) +
    x14ColorXml('negativeFillColor', rule.negativeFillColor) +
    x14ColorXml('negativeBorderColor', rule.negativeBorderColor) +
    x14ColorXml('axisColor', rule.axisColor);
  return `<x14:dataBar${attrs}>${children}</x14:dataBar>`;
}

function x14ColorXml(name: string, color: Color | undefined): string {
  return color === undefined ? '' : `<x14:${name} ${colorAttrs(color)}/>`;
}

function tokenAttr(
  name: string,
  value: string | undefined,
  isValid: (candidate: string) => boolean,
  kind: string,
): string {
  return value === undefined ? '' : ` ${name}="${checkedToken(value, isValid, kind)}"`;
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
 * Curried rather than taking the anchor as a further parameter, because every caller is a `.map`, and
 * a writer taking the form after the anchor would be handed the array index in its place.
 */
function cfvoWriter(
  form: 'classic' | 'x14',
  formulaNames: ReadonlySet<string>,
): (cfvo: CfValueObject) => string {
  return (cfvo) => {
    const checked = checkedToken(cfvo.type, isCfValueObjectType, 'conditional format value type');
    const type = form === 'classic' ? (CLASSIC_ANCHOR[cfvo.type] ?? checked) : checked;
    const tag = form === 'classic' ? 'cfvo' : 'x14:cfvo';
    // `gte` defaults to true, so only a strict threshold states it.
    const gte = cfvo.gte === false ? ' gte="0"' : '';
    if (cfvo.value === undefined) return `<${tag} type="${type}"${gte}/>`;
    // A numeric anchor goes through the number check: `String(NaN)` wrote `val="NaN"`. A `formula`
    // anchor is formula text, and takes the function prefixes Excel stores it with in either form.
    const value =
      typeof cfvo.value === 'number'
        ? numberText(cfvo.value)
        : type === 'formula'
          ? mangleFormula(cfvo.value, formulaNames)
          : cfvo.value;
    return form === 'classic'
      ? `<${tag} type="${type}"${textAttr('val', value)}${gte}/>`
      : `<${tag} type="${type}"${gte}><xm:f>${escapeText(value)}</xm:f></${tag}>`;
  };
}

// The `<extLst>` a classic data-bar cfRule carries to name its x14 extension by shared id.
function cfRuleExtLinkXml(guid: string): string {
  return `<extLst>${x14Ext(DATABAR_LINK_EXT_URI, `<x14:id>${guid}</x14:id>`)}</extLst>`;
}

function blockXml(
  cf: ConditionalFormatting,
  styles: StyleRegistry,
  plan: ConditionalFormattingPlan,
  formulaNames: ReadonlySet<string>,
): string {
  const rules = cf.rules
    .filter((rule) => !inExtension(cf, rule))
    .map((rule) => ruleXml(rule, styles, plan, formulaNames))
    .join('');
  // `CT_ConditionalFormatting` requires at least one `<cfRule>`, so a block with none is omitted
  // rather than emitted empty. A caller can author one, the reader produces one when every rule in a
  // foreign block named a type the enumeration does not allow, and a set whose every rule is written in
  // the extension leaves none here.
  if (rules === '') return '';
  return `<conditionalFormatting sqref="${escapeAttr(cf.ref)}">${rules}</conditionalFormatting>`;
}

function ruleXml(
  rule: ConditionalFormattingRule,
  styles: StyleRegistry,
  plan: ConditionalFormattingPlan,
  formulaNames: ReadonlySet<string>,
): string {
  const attrs = [typeAttr(rule)];
  const dxfId = resolveDxfId(rule, styles);
  if (dxfId !== undefined) attrs.push(`dxfId="${dxfId}"`);
  attrs.push(`priority="${priorityOf(plan, rule)}"`, ...ruleFlagAttrs(rule));

  // A data bar with x14-only facets links to its extension by the id the plan assigned; the extension
  // itself rides in the worksheet <extLst>. The link is the cfRule's last child, after the dataBar. A
  // classic rule absent from the map carries no extension.
  const extGuid = plan.extensionIds.get(rule);
  let body = SCALE_TYPES.has(rule.type)
    ? scaleXml(rule, formulaNames, extGuid !== undefined)
    : operandsXml(rule.formulae, 'formula', formulaNames);
  if (extGuid !== undefined) body += cfRuleExtLinkXml(extGuid);
  return body === ''
    ? `<cfRule ${attrs.join(' ')}/>`
    : `<cfRule ${attrs.join(' ')}>${body}</cfRule>`;
}

function typeAttr(rule: ConditionalFormattingRule): string {
  return `type="${checkedToken(rule.type, isConditionalFormattingType, 'conditional formatting type')}"`;
}

// The attributes both forms of cfRule share past its type and priority, each written only where it
// departs from the schema default.
function ruleFlagAttrs(rule: ConditionalFormattingRule): string[] {
  const attrs: string[] = [];
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
  return attrs;
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

// A rule's operands, each in its own element: a classic `<formula>` or an extension `<xm:f>`.
function operandsXml(
  formulae: readonly (string | number)[] | undefined,
  tag: 'formula' | 'xm:f',
  formulaNames: ReadonlySet<string>,
): string {
  if (formulae === undefined) return '';
  return formulae
    .map(
      (f) => `<${tag}>${escapeText(mangleFormula(stripFormulaEquals(f), formulaNames))}</${tag}>`,
    )
    .join('');
}

function scaleXml(
  rule: ConditionalFormattingRule,
  formulaNames: ReadonlySet<string>,
  linked: boolean,
): string {
  const cfvoXml = cfvoWriter('classic', formulaNames);
  if (rule.type === 'dataBar') return dataBarXml(rule, cfvoXml, linked);
  if (rule.type === 'colorScale') return colorScaleXml(rule, cfvoXml);
  return iconSetXml(rule, cfvoXml, 'classic');
}

// A data bar states its low and high anchors and its bar colour. The minimal call (no cfvo, no colour)
// gains Excel's own defaults, a min/max anchor pair and the standard blue, rather than an invalid
// empty element. Its lengths are stated here only for a bar with no extension: one `linked` to an
// `<x14:dataBar>` keeps them there, beside the facets this element cannot carry.
function dataBarXml(
  rule: ConditionalFormattingRule,
  cfvoXml: (cfvo: CfValueObject) => string,
  linked: boolean,
): string {
  const color = rule.color ?? DEFAULT_DATABAR_COLOR;
  const anchors = dataBarAnchors(rule).map(cfvoXml).join('');
  const lengths = linked
    ? ''
    : intAttr('minLength', rule.minLength, 0) + intAttr('maxLength', rule.maxLength, 0);
  const showValue = rule.showValue === false ? ' showValue="0"' : '';
  return `<dataBar${lengths}${showValue}>${anchors}<color ${colorAttrs(color)}/></dataBar>`;
}

// A colour scale pairs each anchor with a colour; a missing colour list falls back to none, still a
// well-formed (if plain) element.
function colorScaleXml(
  rule: ConditionalFormattingRule,
  cfvoXml: (cfvo: CfValueObject) => string,
): string {
  const anchors = (rule.cfvo ?? []).map(cfvoXml).join('');
  const colors = (rule.colors ?? []).map((c) => `<color ${colorAttrs(c)}/>`).join('');
  return `<colorScale>${anchors}${colors}</colorScale>`;
}

// An icon set in either form: its family, whether it hides the value and runs in reverse, and a
// threshold per icon. Custom icons exist only in the extension form, which is why a rule carrying any
// never reaches the classic one.
function iconSetXml(
  rule: ConditionalFormattingRule,
  cfvoXml: (cfvo: CfValueObject) => string,
  form: 'classic' | 'x14',
): string {
  const tag = form === 'classic' ? 'iconSet' : 'x14:iconSet';
  const icons = form === 'x14' ? (rule.icons ?? []) : [];
  const attrs =
    tokenAttr('iconSet', rule.iconSet, isIconSetType, 'icon set') +
    (rule.showValue === false ? ' showValue="0"' : '') +
    (rule.reverse ? ' reverse="1"' : '') +
    (icons.length > 0 ? ' custom="1"' : '');
  const anchors = (rule.cfvo ?? []).map(cfvoXml).join('');
  return `<${tag}${attrs}>${anchors}${icons.map(cfIconXml).join('')}</${tag}>`;
}

function cfIconXml(icon: CfIcon): string {
  const iconSet = checkedToken(icon.iconSet, isCfIconSetType, 'custom icon set');
  return `<x14:cfIcon iconSet="${iconSet}"${intAttr('iconId', icon.iconId, 0)}/>`;
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
  // `undefined` once a custom icon failed to read: icons pair with thresholds by position, so the rest
  // can no longer be placed, and the set falls back to its family's own icons.
  icons: CfIcon[] | undefined;
  reverse: boolean;
  showValue: boolean | undefined;
  formulae: (string | number)[];
  cfvo: CfValueObject[];
  colors: Color[];
  color: Color | undefined;
  // A data bar's facets, holding only those the markup states, so they fold onto a rule as they stand.
  bar: DataBarFacets;
  // The `<x14:id>` a data-bar cfRule carries to name its extension. Transient: it links this rule to
  // its `<x14:dataBar>` during parsing and is dropped once the extension's facets are folded in.
  x14Id: string | undefined;
  // An extension-form rule's inline `<x14:dxf>`, rebuilt as a `<dxf>`. Transient: the table adopts it
  // and the rule keeps the index.
  dxf: string | undefined;
}

// What an `<x14:dataBar>` a classic rule links to adds to that rule: its facets, and its anchors, which
// can say `autoMin` and `autoMax` where the classic element had to say `min` and `max`.
interface DataBarExt {
  readonly bar: DataBarFacets;
  readonly cfvo: readonly CfValueObject[];
}

// An `<x14:cfRule>` being read, and the id a classic data bar may link to it by: when one does, the
// rule's data-bar facets belong to that classic rule rather than to a rule of their own.
interface ExtensionRuleDraft {
  readonly draft: RuleDraft;
  readonly id: string | undefined;
}

/**
 * An `<x14:dxf>`'s content, rebuilt as the `<dxf>` a styles part holds.
 *
 * The event stream hands back names and decoded attributes, never source text, so the fragment is
 * serialised again from them. A differential style is all attributes and no text, so nothing is lost,
 * and it comes out in the shape Excel writes, which is what lets the table find the copy Excel saved
 * there too. Only main-namespace elements are kept: the fragment lands in a part that binds no other
 * prefix, so an extension element inside the style is left out, subtree and all, as is its `<extLst>`.
 */
class DxfCapture {
  #xml = '<dxf>';
  // How deep the current element sits inside the `<x14:dxf>`, and the depth an ignored subtree began at.
  #depth = 0;
  #skippingFrom: number | undefined;

  open(
    name: string,
    attrs: Record<string, string>,
    selfClosing: boolean,
    scope: NamespaceScope,
  ): void {
    const local = localName(name);
    if (this.#skippingFrom === undefined) {
      if (isMainNamespaceElement(scope, name) && local !== 'extLst') {
        const attributes = Object.entries(attrs)
          .filter(([key]) => key !== 'xmlns' && !key.includes(':'))
          .map(([key, value]) => ` ${key}="${escapeAttr(value)}"`)
          .join('');
        this.#xml += `<${local}${attributes}${selfClosing ? '/>' : '>'}`;
      } else if (!selfClosing) {
        this.#skippingFrom = this.#depth;
      }
    }
    if (!selfClosing) this.#depth += 1;
  }

  /** Close one element: true once the close is the `<x14:dxf>`'s own, which ends the capture. */
  close(name: string): boolean {
    if (this.#depth === 0) {
      this.#xml += '</dxf>';
      return true;
    }
    this.#depth -= 1;
    if (this.#skippingFrom === undefined) this.#xml += `</${localName(name)}>`;
    else if (this.#skippingFrom === this.#depth) this.#skippingFrom = undefined;
    return false;
  }

  get fragment(): string {
    return this.#xml;
  }
}

/**
 * A pass reading a worksheet's conditional formatting into the model, from both forms.
 *
 * The classic `<conditionalFormatting>` blocks come first. Each `<x14:conditionalFormatting>` in the
 * worksheet extension that follows them is read as a set of its own, marked `extended`, unless its rule
 * is a data bar a classic rule links to by id: that one only completes the classic rule with the facets
 * and automatic anchors the classic element cannot carry. An extension rule's inline style is handed to
 * `adoptDifferentialStyle`, which answers the index the rule then holds as its `dxfId`.
 *
 * A rule's formulas shed their function prefixes as a cell formula's do, against `definedNames`, the
 * workbook's names as `definedNameKeys` spells them.
 */
export function conditionalFormattingPass(
  definedNames: ReadonlySet<string>,
  adoptDifferentialStyle: (fragment: string) => number,
): CollectingPass<ConditionalFormatting[]> {
  const blocks: ConditionalFormatting[] = [];
  let block: ConditionalFormatting | undefined;
  let draft: RuleDraft | undefined;
  let scale: ScaleKind | undefined;
  const formulaCapture = new TextCapture('formula');

  // Classic data-bar rules that named an extension, paired with the id they linked on, plus the
  // extensions gathered from the worksheet <extLst>. The two are married after the pass: the
  // extension always follows the classic blocks in document order, so it is known by then.
  const linked: {rule: ConditionalFormattingRule; id: string}[] = [];
  const linkedIds = new Set<string>();
  const extById = new Map<string, DataBarExt>();
  const x14IdCapture = new TextCapture('id');

  // The extension form: the set and the rule being read, the anchor an `<xm:f>` feeds when it sits in
  // one, the set's range, and the inline style being rebuilt.
  let extRules: ConditionalFormattingRule[] | undefined;
  let extRule: ExtensionRuleDraft | undefined;
  let extCfvo: CfValueObject | undefined;
  let extSqref = '';
  const extCapture = new TextCapture(['f', 'sqref']);
  let dxf: DxfCapture | undefined;

  const closeExtensionRule = (): void => {
    if (extRule === undefined) return;
    const {draft: ruleDraft, id} = extRule;
    extRule = undefined;
    extCfvo = undefined;
    if (id !== undefined && linkedIds.has(id)) {
      extById.set(id, {bar: ruleDraft.bar, cfvo: ruleDraft.cfvo});
      return;
    }
    if (ruleDraft.dxf !== undefined)
      ruleDraft.dxfId = String(adoptDifferentialStyle(ruleDraft.dxf));
    const rule = finalizeRule(ruleDraft);
    if (rule !== undefined) extRules?.push(rule);
  };

  const openExtension = (ln: string, attrs: Record<string, string>, selfClosing: boolean): void => {
    // The `<x14:id>` a classic data bar carries to name its extension.
    if (ln === 'id' && draft !== undefined) {
      x14IdCapture.open(ln, selfClosing);
    } else if (ln === 'conditionalFormatting') {
      extRules = [];
      extSqref = '';
    } else if (ln === 'cfRule') {
      extRule = {draft: newDraft(attrs), id: attrs.id};
      // A rule with no children fires no close, so it is finished here.
      if (selfClosing) closeExtensionRule();
    } else if (extRule === undefined) {
      if (extRules !== undefined && ln === 'sqref') extCapture.open(ln, selfClosing);
    } else {
      readExtensionRuleChild(extRule.draft, ln, attrs, selfClosing);
    }
  };

  const readExtensionRuleChild = (
    rule: RuleDraft,
    ln: string,
    attrs: Record<string, string>,
    selfClosing: boolean,
  ): void => {
    switch (ln) {
      case 'dataBar':
        readDataBarAttrs(rule, attrs);
        break;
      case 'iconSet':
        readIconSetAttrs(rule, attrs);
        break;
      case 'cfvo':
        extCfvo = parseCfvo(attrs, definedNames);
        rule.cfvo.push(extCfvo);
        if (selfClosing) extCfvo = undefined;
        break;
      case 'f':
        extCapture.open(ln, selfClosing);
        break;
      case 'color':
        rule.colors.push(parseColor(attrs));
        break;
      case 'fillColor':
        rule.color = parseColor(attrs);
        break;
      case 'borderColor':
      case 'negativeFillColor':
      case 'negativeBorderColor':
      case 'axisColor':
        rule.bar[ln] = parseColor(attrs);
        break;
      case 'cfIcon': {
        const icon = parseCfIcon(attrs);
        if (icon === undefined) rule.icons = undefined;
        else rule.icons?.push(icon);
        break;
      }
      case 'dxf':
        if (selfClosing) rule.dxf = '<dxf/>';
        else dxf = new DxfCapture();
        break;
    }
  };

  const closeExtension = (ln: string): void => {
    const id = x14IdCapture.close(ln);
    if (id !== undefined) {
      if (draft !== undefined) draft.x14Id = id;
      return;
    }
    const text = extCapture.close(ln);
    if (text !== undefined) {
      if (ln === 'sqref') extSqref = text;
      else if (extCfvo !== undefined) extCfvo.value = anchorValue(extCfvo.type, text, definedNames);
      else
        extRule?.draft.formulae.push(coerceNumericLiteral(unmangleFunctions(text, definedNames)));
    } else if (ln === 'cfvo') {
      extCfvo = undefined;
    } else if (ln === 'cfRule') {
      closeExtensionRule();
    } else if (ln === 'conditionalFormatting' && extRules !== undefined) {
      if (extRules.length > 0) blocks.push({ref: extSqref, rules: extRules, extended: true});
      extRules = undefined;
    }
  };

  const handlers: SaxHandlers = {
    onOpen(name, attrs, selfClosing, scope) {
      // An inline style is main-namespace markup inside an extension element, so it is routed before
      // either form looks at it: a `<color>` in there is a font's, not a colour scale's.
      if (dxf !== undefined) {
        dxf.open(name, attrs, selfClosing, scope);
        return;
      }
      const ln = localName(name);
      // In the x14 extension namespace, not merely prefixed. The prefix test was true of every
      // element in a worksheet that binds the MAIN namespace to a prefix, which is legal and which
      // real toolchains emit, so every conditional format in such a file was read as an extension
      // element and dropped.
      if (isExtensionElement(scope, name)) {
        openExtension(ln, attrs, selfClosing);
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
        if (ln === 'iconSet') readIconSetAttrs(draft, attrs);
        else if (ln === 'dataBar') readDataBarAttrs(draft, attrs);
      } else if (draft !== undefined && ln === 'cfvo') {
        draft.cfvo.push(parseCfvo(attrs, definedNames));
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
      extCapture.text(chunk);
    },
    onClose(name, scope) {
      if (dxf !== undefined) {
        if (dxf.close(name)) {
          if (extRule !== undefined) extRule.draft.dxf = dxf.fragment;
          dxf = undefined;
        }
        return;
      }
      const ln = localName(name);
      if (isExtensionElement(scope, name)) {
        closeExtension(ln);
        return;
      }
      const formula = formulaCapture.close(ln);
      if (formula !== undefined) {
        if (draft !== undefined) {
          draft.formulae.push(coerceNumericLiteral(unmangleFunctions(formula, definedNames)));
        }
      } else if (ln === 'dataBar' || ln === 'colorScale' || ln === 'iconSet') {
        scale = undefined;
      } else if (ln === 'cfRule' && draft !== undefined) {
        const rule = finalizeRule(draft);
        if (rule !== undefined) {
          if (block !== undefined) block.rules.push(rule);
          if (draft.x14Id !== undefined) {
            linked.push({rule, id: draft.x14Id});
            linkedIds.add(draft.x14Id);
          }
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
      Object.assign(rule, ext.bar);
      if (ext.cfvo.length > 0) rule.cfvo = [...ext.cfvo];
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
    icons: [],
    reverse: false,
    showValue: undefined,
    formulae: [],
    cfvo: [],
    colors: [],
    color: undefined,
    bar: {},
    x14Id: undefined,
    dxf: undefined,
  };
}

// The attributes an icon set carries in either form. `showValue` defaults to true, so only an explicit
// false is kept.
function readIconSetAttrs(draft: RuleDraft, attrs: Record<string, string>): void {
  draft.iconSet = enumToken(attrs.iconSet, isIconSetType);
  draft.showValue = boolTristate(attrs.showValue);
  draft.reverse = boolStrict(attrs.reverse);
}

// The attributes a data bar carries, the classic element's three and the extension's rest alike: the
// schema gives the classic element only the lengths and `showValue`, so an attribute absent there reads
// as absent. Each facet is kept only as the markup states it, an unrecognised token and an absent
// attribute both leaving the schema default in force.
function readDataBarAttrs(draft: RuleDraft, attrs: Record<string, string>): void {
  const {bar} = draft;
  draft.showValue = boolTristate(attrs.showValue);
  setDefined(bar, 'minLength', numInteger(attrs.minLength, 0));
  setDefined(bar, 'maxLength', numInteger(attrs.maxLength, 0));
  setDefined(bar, 'border', boolTristate(attrs.border));
  setDefined(bar, 'gradient', boolTristate(attrs.gradient));
  setDefined(bar, 'direction', enumToken(attrs.direction, isDataBarDirection));
  setDefined(
    bar,
    'negativeBarColorSameAsPositive',
    boolTristate(attrs.negativeBarColorSameAsPositive),
  );
  setDefined(
    bar,
    'negativeBarBorderColorSameAsPositive',
    boolTristate(attrs.negativeBarBorderColorSameAsPositive),
  );
  setDefined(bar, 'axisPosition', enumToken(attrs.axisPosition, isDataBarAxisPosition));
}

function setDefined<K extends keyof DataBarFacets>(
  bar: DataBarFacets,
  key: K,
  value: DataBarFacets[K] | undefined,
): void {
  if (value !== undefined) bar[key] = value;
}

function parseCfIcon(attrs: Record<string, string>): CfIcon | undefined {
  const iconSet = enumToken(attrs.iconSet, isCfIconSetType);
  const iconId = numInteger(attrs.iconId, 0);
  return iconSet === undefined || iconId === undefined ? undefined : {iconSet, iconId};
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
  if (draft.icons !== undefined && draft.icons.length > 0) rule.icons = draft.icons;
  if (draft.reverse) rule.reverse = true;
  if (draft.showValue === false) rule.showValue = false;
  if (draft.formulae.length > 0) rule.formulae = draft.formulae;
  if (draft.cfvo.length > 0) rule.cfvo = draft.cfvo;
  if (draft.colors.length > 0) rule.colors = draft.colors;
  if (draft.color !== undefined) rule.color = draft.color;
  if (draft.type === 'dataBar') Object.assign(rule, draft.bar);
  return rule;
}

function parseCfvo(
  attrs: Record<string, string>,
  definedNames: ReadonlySet<string>,
): CfValueObject {
  // An anchor whose type is absent reads as `num`, the schema's default; an anchor whose type is
  // unrecognised takes the same fallback rather than being dropped, because an anchor missing from a
  // scale would leave the rule with fewer than the anchors its type requires.
  const raw = attrs.type;
  const type = raw !== undefined && isCfValueObjectType(raw) ? raw : 'num';
  const cfvo: CfValueObject = {type};
  if (attrs.val !== undefined) cfvo.value = anchorValue(type, attrs.val, definedNames);
  if (boolTristate(attrs.gte) === false) cfvo.gte = false;
  return cfvo;
}

// A `formula` anchor's value is an expression and stays a string; the rest are numeric. The classic
// form states it in `val`, the extension in an `<xm:f>`, and both are read here.
function anchorValue(
  type: CfValueObjectType,
  text: string,
  definedNames: ReadonlySet<string>,
): string | number {
  return type === 'formula' ? unmangleFunctions(text, definedNames) : coerceNumericLiteral(text);
}
