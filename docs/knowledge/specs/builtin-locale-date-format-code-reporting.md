# What number-format code to report for locale-dependent builtin date format ids

Cluster: styles

## Scenario

A user opens a spreadsheet whose date column was formatted through the app's Format Cells dialog to a
locale-specific short date (say `dd.mm.yyyy`, showing `21.10.2014`). Reading that cell's number
format string back through the library returns a different code, the canonical English expansion
`mm-dd-yy`, which does not match what the non-US author saw. The confusion is that OOXML stores only
the numeric builtin id (here 14) with no explicit `formatCode`, so the visible day/month/separator
ordering is a function of the reader's locale, not the file.

> Spec note, not a corpus case: this is an open design question about what string to _report_ for a
> locale-sensitive builtin id, not a malformed-serialization bug. Related corpus cases already lock
> that builtin date ids are _detected as dates_ (`builtin-cjk-date-numfmt-ids-resolve-to-date-format`)
> and the literal-`m` scaling (`numfmt-date-detection-literal-m-scaling`); this note is the reporting
> policy they sit under, alongside `xlsx-date-detection-control`.

## Facts

- OOXML defines builtin `numFmtId`s, and several are explicitly **locale-sensitive**: short date (14),
  long date/time (15 to 22), and the elapsed-time group (45 to 47). A file authored in a non-US locale
  still stores only the numeric id, no `formatCode`.
- The canonical English expansion of builtin 14 is `mm-dd-yy`. Reporting that literally is faithful to
  the id but does not match what a non-US author saw.
- The file generally carries **no recoverable authoring locale**. It may have workbook or document
  locale metadata, but the day/month order is resolved by the _reader's_ locale at display time.

## Desired behavior (to decide)

- Pick a **documented, deterministic** policy for the reported format code of a locale-dependent
  builtin id, and apply it consistently. Candidate policies:
  1. Report the canonical English expansion. Faithful to the id and stable, but not "what the author
     saw". Simplest and locale-independent.
  2. Report the id itself (`builtin:14`) and expose the expansion separately, so a caller is not
     misled into thinking a concrete localized code was stored.
  3. Expand against a caller-supplied locale, defaulting to canonical English.
- Whatever is chosen, the write side must round-trip the **builtin id**, not a lossy re-expanded
  string that pins a locale the file never declared.

## Open questions

- Default policy: canonical-English expansion (least machinery) or id-passthrough (most honest about
  what the file stores)?
- Does the API expose both the raw builtin id and a rendered format code so callers can choose?
- Interaction with the `date1904` epoch and with the date-detection opt-out
  (`xlsx-date-detection-control`).

Related: `xlsx-date-detection-control`, `builtin-cjk-date-numfmt-ids-resolve-to-date-format`,
`numfmt-date-detection-literal-m-scaling`, `custom-numfmt-string-roundtrips-verbatim`.
