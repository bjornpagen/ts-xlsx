# ADR 0041: One date-format vocabulary, and it is the spreadsheet's

**Status:** Accepted 2026-09-03

## Context

This library had two public properties that take a date format string, and they spoke different
languages.

`Cell.numFmt` holds an Excel number-format code. So does `DEFAULT_DATE_NUMFMT` (`'yyyy-mm-dd'`), so
does everything `isDateFormat` classifies, and so does every format code that arrives from or leaves
for a `.xlsx` file. The vocabulary is lowercase and case-insensitive, and its one genuinely subtle
rule is that `m` means _month_ or _minute_ depending on where it sits: minutes beside an hour or a
seconds run, months anywhere else.

`CsvWriteOptions.dateFormat` was a moment.js-style token table. It is case-_sensitive_, the month is
uppercase `MM`, and lowercase `mm` means minutes unconditionally. Feeding it this library's own
vocabulary produced silent garbage, and the results were reproduced against the table in the source
rather than reasoned about:

    "yyyy-mm-dd"           ->  "yyyy-45-dd"
    "dd/mm/yyyy"           ->  "dd/45/yyyy"
    "yyyy-mm-dd hh:mm:ss"  ->  "yyyy-45-dd hh:45:07"

No throw, no warning. `writeCsv(wb, {dateFormat: cell.numFmt})` -- which is the obvious thing to
write, and the thing a caller exporting a sheet to CSV actually wants -- emitted `2024-45-dd` for
every date cell in the file.

There is no reading of that which is defensible. Whichever of the two vocabularies is "right", a
library cannot hold both under one name for one concept and expect a caller to know which property
speaks which.

## Decision

**`src/core/date.ts` is the single owner of number-format-code interpretation, in both directions.**
It already held the classifier (`isDateFormat`) and the epoch conversions; it now holds the renderer:

    export function formatSerialDate(date: Date, code: string, utc: boolean): string

driven by the same lowercase `y/m/d/h/s` grammar `isDateFormat` recognises. `formatDate` and
`DATE_TOKENS` are deleted from `src/io/csv/write.ts`, and `CsvWriteOptions.dateFormat` is documented
as what it now is: an Excel number-format code, the same one `Cell.numFmt` holds.

The renderer covers what a date format actually contains: `y`/`m`/`d`/`h`/`s` runs of any length,
month and weekday names (`mmm`, `mmmm`, `mmmmm`, `ddd`, `dddd`), the `AM/PM` and `A/P` meridiems and
the twelve-hour clock they imply (rendered as Excel Desktop renders them: the long form in capitals
however it is spelled, the short form in the case each of its letters is written), fractional seconds (`ss.00`), quoted literals, backslash escapes,
bracketed directives, and the first of a code's `;`-separated sections. Minute-versus-month is
decided by position, as Excel decides it, which is precisely the rule a token table cannot express
and the reason the CSV writer had needed a second vocabulary at all.

**What it deliberately does not cover**, because each is a distinct feature rather than a rounding of
this one:

- **Elapsed-time forms** (`[h]`, `[mm]`, `[ss]`), which measure a duration rather than name a moment.
  A `Date` is the wrong input for them; a duration renderer takes a serial.
- **The locale directive** (`[$-409]`). It is dropped rather than honoured, and month and weekday
  names render in English. Honouring it means a locale table; guessing from `Intl` means the same
  workbook rendering differently on a developer's machine and on a build server, which is not a step
  toward honouring it.

**`formatSerialDate` was built for a second consumer.**
`docs/knowledge/specs/cell-value-raw-and-displayed-accessor.md` asks for a displayed-value accessor,
"the raw value with the cell's effective number format applied". Its date half is this function, and
it takes `utc` for the reason that spec's dates need: a serial carries no timezone, so a `Date` that
came from one is UTC wall-clock.

## Why this break, in the worst shape a break comes in

This changes the _meaning_ of an existing public option without changing its type. A consumer passing
`"MM/DD/YYYY"` gets no compile error and no runtime error; they get a different string. That is the
worst shape of break there is, and it is worth naming rather than glossing.

It is still right, for three reasons.

1. **The alternative is a library with two date vocabularies under one concept.** Every route out of
   that keeps the ambiguity: renaming the CSV option leaves two renderers and two grammars; accepting
   both means guessing from the string, and `MM` is legal in both.
2. **The failure it removes is silent, and the one it introduces mostly is not.** Under the old
   reading, a caller who passed a format code got plausible-looking wrong output that survives review.
   Under the new one, a caller who passed a moment token gets `MM/DD/YYYY` rendered as an Excel code,
   which for the overwhelmingly common formats is _the same string_: `MM/DD/YYYY`, `YYYY-MM-DD` and
   `DD/MM/YYYY` all render identically under both readings, because the tokens differ only in case and
   the new grammar is case-insensitive. The formats that do change are the ones containing a time, and
   there `mm` moves from minutes to months -- visible in the first row of output.
3. **`CHANGELOG.md` and this record are the notice**, and the version is a major. The corpus case
   `csv-write-date-format-honored` now asserts a code where the two readings _disagree_, so the
   contract is pinned rather than left to coincidence.

## Consequences

- `CsvWriteOptions.dateFormat` and `Cell.numFmt` are interchangeable. `writeCsv(wb, {dateFormat:
cell.numFmt})` renders what the cell shows, which is what it always looked like it did.
- The CSV delimiter is validated by one function (`src/io/csv/delimiter.ts`) called from both
  `readCsv` and `writeCsvText`, closing a smaller version of the same defect: the reader refused a
  multi-character delimiter and the writer accepted anything, so `readCsv(writeCsv(wb, {delimiter:
'||'}))` threw on text this codec had just produced, and `{delimiter: ''}` quoted every field and
  emitted no separators at all.
- A displayed-value accessor now has its date renderer waiting for it, rather than a third
  implementation to write.
