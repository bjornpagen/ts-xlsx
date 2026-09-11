# A "time" data-validation type, parallel to the date type

Cluster: data-validation

## Scenario

A spreadsheet author wants to restrict a cell to a valid time-of-day, the same as Excel's
"Allow: Time" data-validation option, with an operator and time-valued bounds: "the entered
time must fall between 00:00 and 23:59", or "after 09:00". The library exposes a `date` validation
type but no dedicated `time` type, so this constraint cannot be expressed and opened by a spreadsheet
app as a Time validation. Authors fall back to a `custom` rule or hand-written XML.

> Spec note, not a corpus case: this is a feature proposal for a new validation type, not a bug with
> a reproduction. The durable value is the format grounding and the API and type shape; a corpus case
> follows once the authoring surface exists, and can then assert the serialized `type="time"` element
> round-trips.

## Desired behavior

- **A `time` validation type sits alongside `date`.** A validation such as
  `{ type: 'time', operator: 'between', formulae: ['00:00', '23:59'], allowBlank, showErrorMessage,
errorStyle, errorTitle, error }` is writable on a cell and produces an OOXML `<dataValidation>`
  whose `type` attribute is `time`, with the operator and formulae serialized so a spreadsheet app
  presents it as an "Allow: Time" rule. On read, that element parses back into the same typed model.
- **Time values are day-fraction serials.** OOXML and Excel store a time as a fraction of a day, so
  12:00 is 0.5, and the persisted `<formula1>` and `<formula2>` carry those fractions. The authoring
  surface should accept human-readable `"hh:mm"` strings and numeric fractions, with a defined,
  documented, timezone-free, locale-independent conversion rule from strings to fractions.
- **The full comparison operator set applies** (`between`, `notBetween`, `equal`, `notEqual`,
  `greaterThan`, `lessThan`, `greaterThanOrEqual`, `lessThanOrEqual`) mirroring the date type.
- **The public validation type union adds `'time'`** so the surface is precisely typed. The types are
  the docs, and a `time` validation is a first-class member of `ST_DataValidationType`
  (`date`, `time`, `decimal`, `whole`, `list`, `textLength`, `custom`).

## Open questions

- Formula representation: accept `"hh:mm"` strings, numeric day-fractions, or both? Excel persists the
  fraction, and a string-to-fraction convenience conversion is desirable but needs a
  locale-independent, timezone-free rule, plus a decision on seconds (`"hh:mm:ss"`).
- Number format pairing: a time validation usually pairs with a time number format on the cell, so do
  we auto-apply one, or leave it to the author?
- Share the serialization path with the existing date validation, the closest analog, so operator and
  two-bound formulae handling is not re-implemented.

Related: `multiselect-dropdown-validation`, `whole-column-data-validation-bounded-memory`,
`list-validation-inline-formula-length-limit`, `date-validation-formula-never-serializes-nan`.
