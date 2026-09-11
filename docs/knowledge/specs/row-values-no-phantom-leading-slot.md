# A row's value collection must not carry a phantom leading empty slot

Cluster: types / API

## Scenario

A caller reads a worksheet row and inspects its values as an array. In the legacy API the array
is 1-based: index 0 is always an empty or `null` placeholder and `values[1]` is column A's value, so
a row with data in columns A to C reads back as `[null, a, b, c]` with length 4. Callers who iterate
or destructure the array trip over the phantom leading element, and the mapping from array index to
column is off by one from every other 0-based collection in the language.

## The tension

The 1-based array is deliberate in the legacy design, so `values[columnNumber]` works with
1-based Excel column numbers, but it is surprising, error-prone, and inconsistent with idiomatic
JavaScript. For a clean-break fork the question is what the row value surface _should_ be, not
preserving the quirk.

## Desired behavior (to decide)

The row value surface must give an **unambiguous, phantom-free** mapping from column to value:

- Iterating a row's populated cells visits only cells that hold values, in column order, each
  carrying its own column identity, with no leading `null` and no gaps materialized as `undefined`.
- If a plain array of values is offered, it is **0-based and dense**, or explicitly sparse with a
  documented convention, and never padded with a leading placeholder to fake 1-based indexing.
- A value read for a given column after a round-trip equals what was written to that column, so the
  column-to-value association is stable and lossless.

## Options

1. **0-based dense array** for contiguous leading data, plus a separate cell iterator for
   sparse or positional access. Idiomatic, but loses direct column-number indexing.
2. **Map or entries API** (`Map<columnLetter|number, value>` or `[{col, value}]`) as the primary
   surface, with arrays as an explicit opt-in. Unambiguous and gap-safe.
3. Keep a 1-based array but **document it loudly and type it** so the leading slot is intentional
   rather than a surprise. The weakest option for a fork whose principle is "correct use is the easy
   path".

## Open questions

- Which surface is primary, dense array or entries/map, given the fork's TypeScript-first,
  least-surprise stance?
- How are genuinely empty in-between cells represented: omitted, `null`, or a sentinel?
- Must the same shape serve both read and the `addRow` and assignment write paths so round-trips are
  symmetric?
