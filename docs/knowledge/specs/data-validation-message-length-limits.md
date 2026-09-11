# Data-validation prompt/error message length limits must not silently corrupt the file

Cluster: data-validation

## Scenario

An author attaches a data-validation rule to a cell and gives it an informational prompt (the
tooltip that appears when the cell is selected) or a custom error message, supplying a couple hundred
characters of guidance text. When the text exceeds the format's maximum length for that field, the
written workbook becomes **corrupt** and a spreadsheet application refuses to open it, with no
indication of why. The author's question ("what is the maximum limit?") is really a symptom: the
library let them cross a hard format limit silently and emitted a broken file.

This is a distinct field and a distinct failure from the inline list-validation formula limit
(`list-validation-inline-formula-length-limit`, whose overflow yields a _structurally valid_ file
with a silently-missing dropdown). Here the overflow corrupts the package outright.

> Spec note, not a corpus case: the durable value is the set of limits and the "never emit a corrupt
> file" contract, plus an undecided design choice between throwing, truncating and documenting.
> Producing a corrupt package on overflow is the bug; the fix is input validation at the authoring or
> write boundary, which is a Phase 3 API-shape decision. It becomes assertable, as a case that authors
> an over-limit message and asserts a clear thrown error or a bounded, still-valid package, once that
> validation exists.

## Desired behavior

- **Never silently emit a corrupt workbook on message overflow.** When a data-validation `prompt`
  (tooltip), `promptTitle`, `error`, or `errorTitle` exceeds its allowed maximum, the library must
  either reject the input at write time with a clear, actionable error that names the offending field
  and its limit, or enforce the limit as a validated property. Silent production of a file that
  applications reject is never acceptable, and the correctness-first stance favors a validated throw
  over silent truncation, since truncation quietly alters the author's intent.
- **The limits are known and enforced by construction.** In the OOXML data-validation model these
  strings are `ST_Xstring`-typed and bounded as the spreadsheet UI enforces them: the prompt/tooltip
  body and the error-message body are limited to **255 characters**, and the prompt title and error
  title to **32 characters**. The authoring API should carry these bounds so an over-limit value
  is caught at author time, not discovered as a corrupt file downstream.
- **Same discipline for every message field.** The four fields (two bodies at 255, two titles at 32)
  are validated uniformly, and control characters or newlines that independently corrupt an `Xstring`
  are guarded by the same validation layer rather than only the length.

## Open questions

- Throw or truncate-with-warning on overflow? A validated throw is preferred here, but confirm
  whether any caller genuinely wants best-effort truncation for bulk imports.
- Confirm the exact title cap (32) and body cap (255) empirically against the format rather than
  trusting the UI numbers, and whether older or newer schema revisions differ.
- Where the validation lives: at the point a validation object is attached to a cell or range (fail
  fast, best error locality), or at serialization time (catches every path but reports late).
  Attaching-time validation with a serialization-time backstop is the likely answer.

Related: `list-validation-inline-formula-length-limit`, `multiselect-dropdown-validation`,
`time-data-validation-type`, `excel-repair-on-open-structural-constraints`.
