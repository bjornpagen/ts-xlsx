# Control automatic date detection when reading xlsx

Cluster: types

## Scenario

OOXML stores dates as a numeric serial value plus a date-shaped number-format code, and there is no
intrinsic "date type" on the cell. To return a JS `Date`, the library applies a heuristic on the
number format. A user reading files where that heuristic guesses wrong, a numeric code that looks
date-ish but is not, or a genuine date they would rather handle as a raw serial, needs a way to make
the policy predictable and to opt out. The older API also conflated this xlsx read option with the
CSV-only date-parsing list, causing confusion when a moment-style format string was passed to the
wrong path.

> Spec note, not a corpus case: several corpus cases already lock specific date-detection outcomes
> (`date-value-written-as-serial-not-text`, `builtin-cjk-date-numfmt-ids-resolve-to-date-format`,
> `numfmt-date-detection-literal-m-scaling`, `strict-mode-iso8601-date-parses-correctly`). This note
> is the umbrella policy and opt-out those cases operate under.

## Desired behavior

- **A documented, deterministic rule** for which number-format codes trigger date coercion on read:
  the built-in date and time format ids, plus custom formats containing date tokens, with the
  literal-`m` minutes-against-months disambiguation already locked separately, and correct handling of
  the workbook's **`date1904` epoch flag**.
- **An opt-out**: a consumer can disable automatic date coercion and receive the **raw numeric serial**
  for every cell, applying their own interpretation. Per-cell or per-read granularity is still to be
  decided.
- **Independence from CSV date parsing**: this xlsx read policy is separate from the CSV-only
  `dateFormats` parsing list, and the two must not share an option name, so passing a moment-style
  format string to one does not silently affect the other.
- **A literal string cell stays a string, even when it looks like a date, and classification is
  uniform.** A reported "inconsistency", where the same column returns some cells as `Date` and others
  as strings, is expected rather than a reader bug: a cell is a `Date` **iff** its stored value is a
  numeric serial _and_ its effective number format is a date format, and a cell whose stored content is
  a literal inline or shared string remains a string no matter how date-like the text reads. The rule
  must be applied so that two cells with the _same_ underlying representation always decode to the same
  value type, and any observed variation must be explainable purely by differences in the source cells
  (number-against-string storage, date-format-against-not) and never by position within the column or
  nondeterminism in the reader. The API should also make a returned value's kind inspectable, genuine
  date against string, so a caller can validate a column without brittle `typeof` or `instanceof`
  probing.

## Open questions

- Granularity of the opt-out: a global read flag (`{ dates: 'raw' | 'auto' }`), a predicate the caller
  supplies (`isDateFormat(numFmtCode) => boolean`), or both?
- Does opting out still expose the number-format code so the caller can run the same heuristic
  themselves?
- Should the default remain auto-detect, the ecosystem expectation, with opt-out, or flip to
  raw-by-default for predictability? Auto-detect-by-default with a documented rule plus opt-out is the
  least-surprise path.
- Interaction with the streaming reader, which must apply the same policy (see
  `streaming-read-applies-date-format`).

Related: `date-value-written-as-serial-not-text`, `numfmt-date-detection-literal-m-scaling`,
`builtin-cjk-date-numfmt-ids-resolve-to-date-format`, `strict-mode-iso8601-date-parses-correctly`,
`date-value-timezone-conversion`, `streaming-read-applies-date-format`.
