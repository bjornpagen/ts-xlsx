// Conditional formatting: the rules that restyle a cell based on its value (a data bar, a colour
// scale, a "highlight cells greater than 10", a formula-driven expression). Like a data validation,
// it is a worksheet-level overlay keyed by a target range, not a facet owned by one cell: one rule
// covers a whole range, and several rules can layer on the same cells with an evaluation precedence.
//
// The model carries the operands each rule type needs and otherwise leaves them absent. A rule type
// the library does not interpret in depth still round-trips its `type`, `priority`, `operator`,
// `formulae`, and differential-style reference, so a read/write cycle never silently drops a rule,
// whichever of the two forms Excel stored it in.

import {tokenSet} from '../token-set.ts';
import {type ClonePlan, cloneWith} from './clone.ts';
import type {AssertNever} from './internal.ts';
import {cloneBorder, cloneFill, cloneFont, type Color} from './style.ts';
import type {DifferentialStyle} from './workbook-styles.ts';

/**
 * How a {@link CfValueObject} reads its `value`: `ST_CfvoType` verbatim, with the `autoMin` and
 * `autoMax` the 2009 extension adds.
 */
export type CfValueObjectType =
  | 'num'
  | 'percent'
  | 'max'
  | 'min'
  | 'percentile'
  | 'formula'
  | 'autoMin'
  | 'autoMax';

/** Narrow a raw `<cfvo type>` token to a known {@link CfValueObjectType}. */
export const isCfValueObjectType = tokenSet<CfValueObjectType>({
  num: true,
  percent: true,
  max: true,
  min: true,
  percentile: true,
  formula: true,
  autoMin: true,
  autoMax: true,
});

/**
 * One anchor of a colour-scale, data-bar, or icon-set scale: a "conditional format value object".
 * `type` names how `value` is read: a literal `num`, a `percent`/`percentile` of the range, a
 * `formula`, or the range's own `min`/`max` (which carry no value).
 *
 * `autoMin` and `autoMax` are the automatic anchors Excel 2010 gives a new data bar, and carry no value
 * either. Only the extension form spells them, so a data bar using one is written with its extension,
 * its classic element saying `min` or `max` in their place as Excel's does.
 */
export interface CfValueObject {
  type: CfValueObjectType;
  value?: number | string;
  /**
   * Whether a value equal to this threshold reaches it. `false` makes an icon set's threshold a strict
   * `>` rather than the `>=` the schema defaults to.
   */
  gte?: boolean;
}

/**
 * What a rule tests, as `<cfRule type>` carries it: `ST_CfType` verbatim.
 *
 * Closed, and stated in full rather than left as `string`, even though the library models only some
 * of these in depth. Depth of modelling and legality are different questions: a `timePeriod` rule
 * whose operands the library never inspects still round-trips, while a token outside this list is one
 * Excel refuses to open, so it is refused on the way in and on the way out alike.
 */
export type ConditionalFormattingType =
  | 'expression'
  | 'cellIs'
  | 'colorScale'
  | 'dataBar'
  | 'iconSet'
  | 'top10'
  | 'uniqueValues'
  | 'duplicateValues'
  | 'containsText'
  | 'notContainsText'
  | 'beginsWith'
  | 'endsWith'
  | 'containsBlanks'
  | 'notContainsBlanks'
  | 'containsErrors'
  | 'notContainsErrors'
  | 'timePeriod'
  | 'aboveAverage';

/** Narrow a raw `<cfRule type>` token to a known {@link ConditionalFormattingType}. */
export const isConditionalFormattingType = tokenSet<ConditionalFormattingType>({
  expression: true,
  cellIs: true,
  colorScale: true,
  dataBar: true,
  iconSet: true,
  top10: true,
  uniqueValues: true,
  duplicateValues: true,
  containsText: true,
  notContainsText: true,
  beginsWith: true,
  endsWith: true,
  containsBlanks: true,
  notContainsBlanks: true,
  containsErrors: true,
  notContainsErrors: true,
  timePeriod: true,
  aboveAverage: true,
});

/**
 * How a `cellIs` or text rule compares, as `ST_ConditionalFormattingOperator` enumerates it.
 *
 * Overlaps {@link import('./data-validation.ts').DataValidationOperator} in four members and diverges
 * in the rest: this one has the text comparisons a validation has no use for, and spells "does not
 * contain" as `notContains` where nothing else in the format does.
 */
export type ConditionalFormattingOperator =
  | 'lessThan'
  | 'lessThanOrEqual'
  | 'equal'
  | 'notEqual'
  | 'greaterThanOrEqual'
  | 'greaterThan'
  | 'between'
  | 'notBetween'
  | 'containsText'
  | 'notContains'
  | 'beginsWith'
  | 'endsWith';

/** Narrow a raw `<cfRule operator>` token to a known {@link ConditionalFormattingOperator}. */
export const isConditionalFormattingOperator = tokenSet<ConditionalFormattingOperator>({
  lessThan: true,
  lessThanOrEqual: true,
  equal: true,
  notEqual: true,
  greaterThanOrEqual: true,
  greaterThan: true,
  between: true,
  notBetween: true,
  containsText: true,
  notContains: true,
  beginsWith: true,
  endsWith: true,
});

/** The window a `timePeriod` rule matches against, relative to the day the sheet is recalculated. */
export type CfTimePeriod =
  | 'today'
  | 'yesterday'
  | 'tomorrow'
  | 'last7Days'
  | 'thisMonth'
  | 'lastMonth'
  | 'nextMonth'
  | 'thisWeek'
  | 'lastWeek'
  | 'nextWeek';

/** Narrow a raw `<cfRule timePeriod>` token to a known {@link CfTimePeriod}. */
export const isCfTimePeriod = tokenSet<CfTimePeriod>({
  today: true,
  yesterday: true,
  tomorrow: true,
  last7Days: true,
  thisMonth: true,
  lastMonth: true,
  nextMonth: true,
  thisWeek: true,
  lastWeek: true,
  nextWeek: true,
});

/**
 * Which way a data bar grows, as `ST_DataBarDirection` enumerates it: `context` follows the sheet's
 * reading direction.
 */
export type DataBarDirection = 'context' | 'leftToRight' | 'rightToLeft';

/** Narrow a raw `<x14:dataBar direction>` token to a known {@link DataBarDirection}. */
export const isDataBarDirection = tokenSet<DataBarDirection>({
  context: true,
  leftToRight: true,
  rightToLeft: true,
});

/**
 * Where a data bar draws the axis between its negative and positive bars, as `ST_DataBarAxisPosition`
 * enumerates it: where the values put it, at the middle of the cell, or nowhere, which grows a
 * negative bar the way a positive one grows.
 */
export type DataBarAxisPosition = 'automatic' | 'middle' | 'none';

/** Narrow a raw `<x14:dataBar axisPosition>` token to a known {@link DataBarAxisPosition}. */
export const isDataBarAxisPosition = tokenSet<DataBarAxisPosition>({
  automatic: true,
  middle: true,
  none: true,
});

/**
 * The named icon family an `iconSet` rule draws from, as `ST_IconSetType` enumerates it. The leading
 * digit is the number of icons, which is also how many {@link CfValueObject} anchors the rule needs.
 *
 * `3Stars`, `3Triangles` and `5Boxes` are the families the 2009 extension added. The classic
 * `<iconSet>` element cannot name them, so a rule drawing from one is written in the extension form.
 */
export type IconSetType =
  | '3Arrows'
  | '3ArrowsGray'
  | '3Flags'
  | '3TrafficLights1'
  | '3TrafficLights2'
  | '3Signs'
  | '3Symbols'
  | '3Symbols2'
  | '4Arrows'
  | '4ArrowsGray'
  | '4RedToBlack'
  | '4Rating'
  | '4TrafficLights'
  | '5Arrows'
  | '5ArrowsGray'
  | '5Rating'
  | '5Quarters'
  | '3Stars'
  | '3Triangles'
  | '5Boxes';

/** Narrow a raw `<iconSet iconSet>` token to a known {@link IconSetType}. */
export const isIconSetType = tokenSet<IconSetType>({
  '3Arrows': true,
  '3ArrowsGray': true,
  '3Flags': true,
  '3TrafficLights1': true,
  '3TrafficLights2': true,
  '3Signs': true,
  '3Symbols': true,
  '3Symbols2': true,
  '4Arrows': true,
  '4ArrowsGray': true,
  '4RedToBlack': true,
  '4Rating': true,
  '4TrafficLights': true,
  '5Arrows': true,
  '5ArrowsGray': true,
  '5Rating': true,
  '5Quarters': true,
  '3Stars': true,
  '3Triangles': true,
  '5Boxes': true,
});

/** The families only the 2009 extension can name: a rule drawing from one is written there. */
export function isExtensionIconSet(iconSet: IconSetType): boolean {
  return iconSet === '3Stars' || iconSet === '3Triangles' || iconSet === '5Boxes';
}

/** The family a custom icon comes from: any {@link IconSetType}, or `NoIcons` for a threshold showing none. */
export type CfIconSetType = IconSetType | 'NoIcons';

/** Narrow a raw `<x14:cfIcon iconSet>` token to a known {@link CfIconSetType}. */
export function isCfIconSetType(value: string): value is CfIconSetType {
  return value === 'NoIcons' || isIconSetType(value);
}

/** One threshold's icon in a custom icon set: icon number `iconId`, from 0, of the family `iconSet`. */
export interface CfIcon {
  iconSet: CfIconSetType;
  iconId: number;
}

/**
 * A single conditional-formatting rule. `type` is the OOXML cfRule type; the remaining fields carry
 * the operands that type needs and are absent otherwise. A rule the library does not model in depth
 * still preserves `type`, `priority`, `operator`, `formulae`, and `dxfId` across a round-trip.
 *
 * A data bar's facets marked as x14 properties have no place in the classic `<dataBar>` element, so a
 * bar carrying any of them, or an automatic anchor, is written with a linked `<x14:dataBar>` as well.
 */
export interface ConditionalFormattingRule {
  type: ConditionalFormattingType;
  /** Evaluation precedence; lower wins. Excel requires one, so the writer supplies it when absent. */
  priority?: number;
  /** Halt evaluation of lower-priority rules on any cell this rule matches. */
  stopIfTrue?: boolean;
  /** cellIs / text comparison operator (`greaterThan`, `between`, `beginsWith`, …). */
  operator?: ConditionalFormattingOperator;
  /** Formula operands: cellIs bounds, an expression predicate, a containsText target formula, … */
  formulae?: (string | number)[];
  /** The literal a containsText / beginsWith / endsWith rule searches for. */
  text?: string;
  /** A differential style authored inline, serialised into `<dxfs>` and referenced by the cfRule. */
  style?: DifferentialStyle;
  /** A differential-style reference by `<dxfs>` index, as read from a file (kept verbatim). */
  dxfId?: string;
  /** colorScale / dataBar / iconSet scale anchors, in order. */
  cfvo?: CfValueObject[];
  /** A dataBar's bar colour. */
  color?: Color;
  /** A colorScale's colours, one per {@link cfvo}. */
  colors?: Color[];
  /** A dataBar's shortest bar, as a percentage of the cell. */
  minLength?: number;
  /** A dataBar's longest bar, as a percentage of the cell. */
  maxLength?: number;
  /** A dataBar's gradient-fill flag. An x14 property. */
  gradient?: boolean;
  /** Whether a dataBar has a border, in {@link borderColor}. An x14 property. */
  border?: boolean;
  /** A dataBar's border colour. An x14 property. */
  borderColor?: Color;
  /** Which way a dataBar grows. An x14 property. */
  direction?: DataBarDirection;
  /** A dataBar's fill colour for negative values. An x14 property. */
  negativeFillColor?: Color;
  /** A dataBar's border colour for negative values. An x14 property. */
  negativeBorderColor?: Color;
  /**
   * Whether a dataBar fills a negative bar in its positive colour rather than
   * {@link negativeFillColor}. An x14 property.
   */
  negativeBarColorSameAsPositive?: boolean;
  /**
   * Whether a dataBar borders a negative bar in its positive border colour rather than
   * {@link negativeBorderColor}. An x14 property.
   */
  negativeBarBorderColorSameAsPositive?: boolean;
  /** Where a dataBar draws the axis between negative and positive bars. An x14 property. */
  axisPosition?: DataBarAxisPosition;
  /** A dataBar's axis colour (the zero line between positive and negative bars). An x14 property. */
  axisColor?: Color;
  /** An iconSet's named icon family (e.g. `3TrafficLights1`). */
  iconSet?: IconSetType;
  /**
   * A custom iconSet's icons, one per {@link cfvo} threshold in order, each replacing the icon
   * {@link iconSet} would show there. Only the extension form carries them, so a rule with icons is
   * written in it.
   */
  icons?: CfIcon[];
  /** iconSet: the icons in reverse order, the highest threshold showing the family's first icon. */
  reverse?: boolean;
  /** dataBar / iconSet: whether the cell still shows its value beside the bar or icon. */
  showValue?: boolean;
  /** top10 rank cutoff. */
  rank?: number;
  /** top10: the rank is a percentage rather than a count. */
  percent?: boolean;
  /** top10: rank from the bottom rather than the top. */
  bottom?: boolean;
  /** aboveAverage: match above (default) or below the average. */
  aboveAverage?: boolean;
  /** aboveAverage: include cells equal to the average. */
  equalAverage?: boolean;
  /** aboveAverage: match beyond this many standard deviations. */
  stdDev?: number;
  /** timePeriod window (`today`, `lastWeek`, …). */
  timePeriod?: CfTimePeriod;
}

/**
 * A set of rules bound to the range(s) they cover. `ref` is an OOXML `sqref`: one or more
 * space-separated areas (`"A1:C1 A3:C3 A5:C5"`), the shape Excel writes when one rule is applied to
 * several non-contiguous selections at once.
 *
 * `extended` marks a set stored in the 2009 extension form (`<x14:conditionalFormatting>` inside the
 * worksheet `<extLst>`), where Excel puts a rule whose formula reaches another sheet. The reader sets
 * it for a set found in that form so a round-trip writes the set back there. A rule only that form can
 * express, one drawing from a 2009 icon family or carrying custom icons, is written there whatever the
 * flag says.
 */
export interface ConditionalFormatting {
  ref: string;
  rules: ConditionalFormattingRule[];
  extended?: boolean;
}

/** A defensive deep copy, so a stored conditional formatting never aliases the caller's object nor
 * any of its nested arrays (rules, formulae, cfvo, colours) or the differential style. */
export function cloneConditionalFormatting(cf: ConditionalFormatting): ConditionalFormatting {
  const copy: ConditionalFormatting = {ref: cf.ref, rules: cf.rules.map(cloneRule)};
  if (cf.extended !== undefined) copy.extended = cf.extended;
  return copy;
}

/**
 * Whether a rule needs the extension form: it names a family or custom icons the classic element has
 * no way to spell.
 */
export function ruleNeedsExtension(rule: ConditionalFormattingRule): boolean {
  return (
    rule.type === 'iconSet' &&
    ((rule.iconSet !== undefined && isExtensionIconSet(rule.iconSet)) ||
      (rule.icons !== undefined && rule.icons.length > 0))
  );
}

// One entry per field, so a new one on the rule does not compile until someone says how deep its copy
// goes. The spread this replaces carried a new nested field BY REFERENCE, which meant the stored rule
// aliased the caller's object with nothing failing to build.
const RULE_CLONE: ClonePlan<ConditionalFormattingRule> = {
  type: 'value',
  priority: 'value',
  stopIfTrue: 'value',
  operator: 'value',
  formulae: 'values',
  text: 'value',
  style: cloneStyle,
  dxfId: 'value',
  cfvo: 'records',
  color: 'record',
  colors: 'records',
  minLength: 'value',
  maxLength: 'value',
  gradient: 'value',
  border: 'value',
  borderColor: 'record',
  direction: 'value',
  negativeFillColor: 'record',
  negativeBorderColor: 'record',
  negativeBarColorSameAsPositive: 'value',
  negativeBarBorderColorSameAsPositive: 'value',
  axisPosition: 'value',
  axisColor: 'record',
  iconSet: 'value',
  icons: 'records',
  reverse: 'value',
  showValue: 'value',
  rank: 'value',
  percent: 'value',
  bottom: 'value',
  aboveAverage: 'value',
  equalAverage: 'value',
  stdDev: 'value',
  timePeriod: 'value',
};
export type EveryRuleFieldIsCloned = AssertNever<
  Exclude<keyof Required<ConditionalFormattingRule>, keyof typeof RULE_CLONE>
>;

function cloneRule(rule: ConditionalFormattingRule): ConditionalFormattingRule {
  return cloneWith(rule, RULE_CLONE);
}

// Each of the three non-scalar facets copies through its own plan in `style.ts`. They were `'record'`
// -- a spread -- which `clone.ts` defines as the strategy for a flat object, and none of the three is
// flat: a font carries a nested colour, a border five nested edges each with their own, a gradient
// fill an array of stops. So the copy stopped one level short of what the doc above promises, and a
// stored rule shared its border edges and its font colour with the caller's object.
const STYLE_CLONE: ClonePlan<DifferentialStyle> = {
  font: cloneFont,
  fill: cloneFill,
  border: cloneBorder,
  numFmt: 'value',
};
export type EveryDifferentialStyleFieldIsCloned = AssertNever<
  Exclude<keyof Required<DifferentialStyle>, keyof typeof STYLE_CLONE>
>;

function cloneStyle(style: DifferentialStyle): DifferentialStyle {
  return cloneWith(style, STYLE_CLONE);
}
