# A unified accessor for a cell's raw and displayed value

Cluster: types

## Scenario

A consumer iterating cells wants "the value" without branching over every value-type variant. Today a
cell's `value` is a discriminated union (a number, a string, a `Date`, a rich-text object, a
hyperlink object `{text, hyperlink}`, a formula object `{formula, result}`, an error, a boolean) so a
caller who just wants the underlying scalar, or just the string a spreadsheet would display, must
hand-write a switch over all of them. The recurring ask is for one accessor that returns the _raw_
value (the underlying scalar, formula result unwrapped, hyperlink text extracted) or the
_displayed_ value (the string the application renders, number-format applied).

> Spec note, not a corpus case: this is an API-ergonomics design decision with no failing current
> behavior to baseline. The union is already exposed and correct, and the gap is a convenience layer on
> top of it. Recording the shape and the open questions feeds Phase 3 design.

## Desired behavior

Offer two distinct, clearly-named accessors so a caller never has to destructure the union by hand:

- **numFmt preserved and exposed on read is the precondition.** A cell's number-format code must be
  faithfully preserved and exposed alongside its raw stored value on read, independent of any display
  rendering. A consumer who opens a file where a cell stores a number under a format that renders it
  as `"8"` (rounded or scaled) must be able to see _both_ the raw value and the numFmt that says how it
  is meant to display. Returning the raw value is correct and lossless by default; the recurring
  surprise ("the format was not applied to the output") is resolved not by changing the raw value but
  by making the numFmt visible and offering the displayed accessor below. Faithful numFmt round-trip
  is the hard invariant; rendering formatted text is the optional layer on top.

- **Raw value.** The underlying scalar, with wrapper objects unwrapped: a formula cell yields its
  cached `result`, a hyperlink cell yields its display text (or a structured `{text, target}` if the
  caller wants the link), rich text collapses to its concatenated plain text, a date stays a `Date`, a
  number stays a number, and an error surfaces as a typed error value. This is what a caller means by
  "just give me the data" for export and serialization.

- **Displayed value.** The string the spreadsheet application would render for the cell: the raw value
  with the cell's effective number format applied, so `0.5` under a percent format reads `"50%"` and a
  date serial under a date format reads the formatted date. This requires a number-format
  formatter and must honor the workbook's locale and `date1904` epoch, consistent with the read-time
  date policy (see `xlsx-date-detection-control`).

- **Kind inspection.** Pair the accessors with a way to ask a value's kind (number, string, date,
  formula, hyperlink, rich-text, boolean, error) without brittle `typeof` or `instanceof` probing, so
  a caller can validate a column's contents robustly. This is the same affordance the
  date-detection note calls for, generalized to every value type.

The type surface must make both accessors precisely typed: the raw accessor's return is the unwrapped
union; the displayed accessor returns `string`, with a defined result for empty or null cells.

## Open questions

- ~~Naming: does `cell.text` become the "displayed value" accessor, or do we introduce distinct
  `raw`/`display` members?~~ **Answered: `cell.text` is the _raw_ value's plain text, and applies no
  number format.** `cellValueToText` (`src/core/value.ts`) is the function, total over `CellValue`,
  and `Cell.text` is it applied to the cell in hand. Overloading was avoided from the other end than
  this note expected: rather than giving the fuzzy name the richer meaning, the name was pinned to
  the meaning that needs no formatter, and its doc says so outright ("a currency cell's text carries
  no currency sign"). The displayed accessor is therefore still unbuilt, and now needs a name of its
  own, `displayText` or similar, which is the honest cost of this choice.
- ~~Kind inspection without brittle `typeof`/`instanceof` probing.~~ **Answered:** `detectValueType`
  classifies, and the six `isXxxValue` guards, public since the same release, narrow. A caller
  validating a column has both the total switch and the per-kind predicate.
- Does the displayed accessor build in a full number-format formatter (locale-aware), or start with a
  documented subset and defer exotic format codes? **Still open, and now the only thing between the
  two accessors:** everything except the formatter exists.
- ~~Hyperlink and rich-text cells: does "raw" flatten to a plain scalar/string by default, with the
  structured form available on request, or the reverse?~~ **Answered: neither is a default, because
  they are separate members.** `cell.value` is always the structured form and `cell.text` always the
  flattened one, so no caller has to opt in to the shape it wanted.
- ~~Should these be lazy accessors computed on read, or precomputed?~~ **Answered: computed on
  access, and the question dissolves without the formatter.** Flattening runs is cheap enough that
  caching would cost more invalidation than it saves. It returns if the displayed accessor lands.

Related: `xlsx-date-detection-control`, `column-level-value-type`,
`formula-cell-value-type-minimal-required-fields`, `html-fragment-to-rich-text-cell-value`,
`public-type-surface-matches-runtime`.
