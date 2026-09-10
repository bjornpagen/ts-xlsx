# Data Validation

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `DataValidation`

<sub>interface</sub>

One validation rule. `formulae` holds the operand(s), `formula1` then optional `formula2`: a
numeric literal is stored as a number, while a cell reference, defined name, or list source keeps
its verbatim string.

```ts
interface DataValidation {
  type: DataValidationType;
  operator?: DataValidationOperator;
  formulae?: (string | number)[];
  allowBlank?: boolean;
  /**
   * Hide the in-cell dropdown arrow of a `list` rule while still enforcing the list, the usual way to
   * validate against a list without offering a picker. Stored as `showDropDown="1"`, an attribute
   * whose name says the opposite of what it does.
   */
  suppressDropDown?: boolean;
  showInputMessage?: boolean;
  showErrorMessage?: boolean;
  errorStyle?: DataValidationErrorStyle;
  /** How the input method editor behaves in a covered cell; absent leaves it as the user had it. */
  imeMode?: DataValidationImeMode;
  error?: string;
  errorTitle?: string;
  prompt?: string;
  promptTitle?: string;
}
```

---

### `DataValidationEntry`

<sub>interface</sub>

A validation bound to the range(s) it covers. `sqref` is an OOXML `sqref`: one or more
space-separated ranges. `extended` marks a rule stored in the 2009 extension form
(`<x14:dataValidation>` inside the worksheet `<extLst>`), Excel's carrier for validations a
legacy `<dataValidation>` cannot express, such as a list source on another sheet. The flag is how
a rule read from that form remembers to be written back to it, rather than downgraded to the
standard element (which would corrupt a cross-sheet reference).

```ts
interface DataValidationEntry {
  sqref: string;
  rule: DataValidation;
  extended?: boolean;
}
```

---

### `DataValidationErrorStyle`

<sub>type</sub>

How Excel reacts to input that fails the rule.

```ts
type DataValidationErrorStyle = 'stop' | 'warning' | 'information';
```

---

### `DataValidationImeMode`

<sub>type</sub>

How the input method editor behaves while a covered cell is edited. Only an East Asian input method
acts on it: `hiragana`, the katakana, alpha and hangul widths switch its mode, `on` and `off` turn it
on and off, `disabled` turns it off and keeps it off, and `noControl` leaves it as the user had it.

```ts
type DataValidationImeMode =
  | 'noControl'
  | 'off'
  | 'on'
  | 'disabled'
  | 'hiragana'
  | 'fullKatakana'
  | 'halfKatakana'
  | 'fullAlpha'
  | 'halfAlpha'
  | 'fullHangul'
  | 'halfHangul';
```

---

### `DataValidationOperator`

<sub>type</sub>

How a typed validation compares its operand(s). Absent on a `list`/`custom` rule; defaults to
`between` on a typed rule (the value Excel omits from the XML).

```ts
type DataValidationOperator =
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'greaterThan'
  | 'lessThan'
  | 'greaterThanOrEqual'
  | 'lessThanOrEqual';
```

---

### `DataValidationType`

<sub>type</sub>

The kind of constraint a validation enforces. `list` is a dropdown; `custom` is an arbitrary
boolean formula; `none` constrains nothing and exists only to carry the rule's messages; the rest
bound a typed value (`whole`/`decimal`/`date`/`time`/`textLength`).

```ts
type DataValidationType =
  | 'none'
  | 'list'
  | 'whole'
  | 'decimal'
  | 'date'
  | 'time'
  | 'textLength'
  | 'custom';
```

---

### `isDataValidationErrorStyle`

<sub>const</sub>

Narrow a raw `<dataValidation errorStyle>` token to a known [`DataValidationErrorStyle`](./data-validation.md#datavalidationerrorstyle).

```ts
const isDataValidationErrorStyle: (value: string) => value is DataValidationErrorStyle
```

---

### `isDataValidationImeMode`

<sub>const</sub>

Narrow a raw `<dataValidation imeMode>` token to a known [`DataValidationImeMode`](./data-validation.md#datavalidationimemode).

```ts
const isDataValidationImeMode: (value: string) => value is DataValidationImeMode
```

---

### `isDataValidationOperator`

<sub>const</sub>

Narrow a raw `<dataValidation operator>` token to a known [`DataValidationOperator`](./data-validation.md#datavalidationoperator).

```ts
const isDataValidationOperator: (value: string) => value is DataValidationOperator
```

---

### `isDataValidationType`

<sub>const</sub>

Narrow a raw `<dataValidation type>` token to a known [`DataValidationType`](./data-validation.md#datavalidationtype).

```ts
const isDataValidationType: (value: string) => value is DataValidationType
```
