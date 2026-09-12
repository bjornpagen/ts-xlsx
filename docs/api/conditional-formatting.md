# Conditional Formatting

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `CfIcon`

<sub>interface</sub>

One threshold's icon in a custom icon set: icon number `iconId`, from 0, of the family `iconSet`.

```ts
interface CfIcon {
  iconSet: CfIconSetType;
  iconId: number;
}
```

---

### `CfIconSetType`

<sub>type</sub>

The family a custom icon comes from: any [`IconSetType`](./conditional-formatting.md#iconsettype), or `NoIcons` for a threshold showing none.

```ts
type CfIconSetType = IconSetType | 'NoIcons';
```

---

### `CfTimePeriod`

<sub>type</sub>

The window a `timePeriod` rule matches against, relative to the day the sheet is recalculated.

```ts
type CfTimePeriod =
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
```

---

### `CfValueObject`

<sub>interface</sub>

One anchor of a colour-scale, data-bar, or icon-set scale: a "conditional format value object".
`type` names how `value` is read: a literal `num`, a `percent`/`percentile` of the range, a
`formula`, or the range's own `min`/`max` (which carry no value).

`autoMin` and `autoMax` are the automatic anchors Excel 2010 gives a new data bar, and carry no value
either. Only the extension form spells them, so a data bar using one is written with its extension,
its classic element saying `min` or `max` in their place as Excel's does.

```ts
interface CfValueObject {
  type: CfValueObjectType;
  value?: number | string;
  /**
   * Whether a value equal to this threshold reaches it. `false` makes an icon set's threshold a strict
   * `>` rather than the `>=` the schema defaults to.
   */
  gte?: boolean;
}
```

---

### `CfValueObjectType`

<sub>type</sub>

How a [`CfValueObject`](./conditional-formatting.md#cfvalueobject) reads its `value`: `ST_CfvoType` verbatim, with the `autoMin` and
`autoMax` the 2009 extension adds.

```ts
type CfValueObjectType =
  | 'num'
  | 'percent'
  | 'max'
  | 'min'
  | 'percentile'
  | 'formula'
  | 'autoMin'
  | 'autoMax';
```

---

### `ConditionalFormatting`

<sub>interface</sub>

A set of rules bound to the range(s) they cover. `ref` is an OOXML `sqref`: one or more
space-separated areas (`"A1:C1 A3:C3 A5:C5"`), the shape Excel writes when one rule is applied to
several non-contiguous selections at once.

`extended` marks a set stored in the 2009 extension form (`<x14:conditionalFormatting>` inside the
worksheet `<extLst>`), where Excel puts a rule whose formula reaches another sheet. The reader sets
it for a set found in that form so a round-trip writes the set back there. A rule only that form can
express, one drawing from a 2009 icon family or carrying custom icons, is written there whatever the
flag says.

```ts
interface ConditionalFormatting {
  ref: string;
  rules: ConditionalFormattingRule[];
  extended?: boolean;
}
```

---

### `ConditionalFormattingOperator`

<sub>type</sub>

How a `cellIs` or text rule compares, as `ST_ConditionalFormattingOperator` enumerates it.

Overlaps `import ('./data-validation.ts').DataValidationOperator` in four members and diverges
in the rest: this one has the text comparisons a validation has no use for, and spells "does not
contain" as `notContains` where nothing else in the format does.

```ts
type ConditionalFormattingOperator =
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
```

---

### `ConditionalFormattingRule`

<sub>interface</sub>

A single conditional-formatting rule. `type` is the OOXML cfRule type; the remaining fields carry
the operands that type needs and are absent otherwise. A rule the library does not model in depth
still preserves `type`, `priority`, `operator`, `formulae`, and `dxfId` across a round-trip.

A data bar's facets marked as x14 properties have no place in the classic `<dataBar>` element, so a
bar carrying any of them, or an automatic anchor, is written with a linked `<x14:dataBar>` as well.

```ts
interface ConditionalFormattingRule {
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
```

---

### `ConditionalFormattingType`

<sub>type</sub>

What a rule tests, as `<cfRule type>` carries it: `ST_CfType` verbatim.

Closed, and stated in full rather than left as `string`, even though the library models only some
of these in depth. Depth of modelling and legality are different questions: a `timePeriod` rule
whose operands the library never inspects still round-trips, while a token outside this list is one
Excel refuses to open, so it is refused on the way in and on the way out alike.

```ts
type ConditionalFormattingType =
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
```

---

### `DataBarAxisPosition`

<sub>type</sub>

Where a data bar draws the axis between its negative and positive bars, as `ST_DataBarAxisPosition`
enumerates it: where the values put it, at the middle of the cell, or nowhere, which grows a
negative bar the way a positive one grows.

```ts
type DataBarAxisPosition = 'automatic' | 'middle' | 'none';
```

---

### `DataBarDirection`

<sub>type</sub>

Which way a data bar grows, as `ST_DataBarDirection` enumerates it: `context` follows the sheet's
reading direction.

```ts
type DataBarDirection = 'context' | 'leftToRight' | 'rightToLeft';
```

---

### `IconSetType`

<sub>type</sub>

The named icon family an `iconSet` rule draws from, as `ST_IconSetType` enumerates it. The leading
digit is the number of icons, which is also how many [`CfValueObject`](./conditional-formatting.md#cfvalueobject) anchors the rule needs.

`3Stars`, `3Triangles` and `5Boxes` are the families the 2009 extension added. The classic
`<iconSet>` element cannot name them, so a rule drawing from one is written in the extension form.

```ts
type IconSetType =
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
```

---

### `isCfIconSetType`

<sub>function</sub>

Narrow a raw `<x14:cfIcon iconSet>` token to a known [`CfIconSetType`](./conditional-formatting.md#cficonsettype).

```ts
function isCfIconSetType(value: string): value is CfIconSetType;
```

---

### `isCfTimePeriod`

<sub>const</sub>

Narrow a raw `<cfRule timePeriod>` token to a known [`CfTimePeriod`](./conditional-formatting.md#cftimeperiod).

```ts
const isCfTimePeriod: (value: string) => value is CfTimePeriod
```

---

### `isCfValueObjectType`

<sub>const</sub>

Narrow a raw `<cfvo type>` token to a known [`CfValueObjectType`](./conditional-formatting.md#cfvalueobjecttype).

```ts
const isCfValueObjectType: (value: string) => value is CfValueObjectType
```

---

### `isConditionalFormattingOperator`

<sub>const</sub>

Narrow a raw `<cfRule operator>` token to a known [`ConditionalFormattingOperator`](./conditional-formatting.md#conditionalformattingoperator).

```ts
const isConditionalFormattingOperator: (value: string) => value is ConditionalFormattingOperator
```

---

### `isConditionalFormattingType`

<sub>const</sub>

Narrow a raw `<cfRule type>` token to a known [`ConditionalFormattingType`](./conditional-formatting.md#conditionalformattingtype).

```ts
const isConditionalFormattingType: (value: string) => value is ConditionalFormattingType
```

---

### `isDataBarAxisPosition`

<sub>const</sub>

Narrow a raw `<x14:dataBar axisPosition>` token to a known [`DataBarAxisPosition`](./conditional-formatting.md#databaraxisposition).

```ts
const isDataBarAxisPosition: (value: string) => value is DataBarAxisPosition
```

---

### `isDataBarDirection`

<sub>const</sub>

Narrow a raw `<x14:dataBar direction>` token to a known [`DataBarDirection`](./conditional-formatting.md#databardirection).

```ts
const isDataBarDirection: (value: string) => value is DataBarDirection
```

---

### `isIconSetType`

<sub>const</sub>

Narrow a raw `<iconSet iconSet>` token to a known [`IconSetType`](./conditional-formatting.md#iconsettype).

```ts
const isIconSetType: (value: string) => value is IconSetType
```
