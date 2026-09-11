# Streaming read must resolve styles before emitting cells, and must not stall

Cluster: streaming

## Scenario

Consumers of the streaming reader hit a cluster of related failures on real files. Sometimes a
read **stalls** and never completes; sometimes a file can only be read after being opened and
re-saved in Excel first; and sometimes **date cells come back as raw numbers** instead of dates.
The date failure is a classic ordering race: a date cell carries a bare serial value plus a
style-driven number format, so it can only be recognized as a date once the styles part has been
parsed, but the streaming reader begins emitting worksheet cells before the styles part is
guaranteed to have been consumed. When the worksheet parts precede the styles part in the package,
a legitimate, generator-dependent ZIP ordering, the cell is emitted with no format applied and
surfaces as a number. The "open and re-save fixes it" reports are the same root: Excel rewrites the
package into an ordering the reader happens to tolerate.

> The date-application _behavior_ is locked as a corpus case
> (`streaming-read-applies-date-format`). This note captures the surrounding streaming-reader design
> requirements, meaning ordering-independence, default style interpretation, and guaranteed
> termination, including the ones that cannot be a corpus case because they are hang or latency
> properties.

## Desired behavior

- **Style resolution precedes cell emission regardless of physical part order.** A date-formatted
  cell must be surfaced as a date whether the styles part appears before or after the worksheet
  parts in the ZIP. The reader must buffer or defer cell interpretation until the styles table is
  available, rather than racing the two.
- **Styles are interpreted by default.** A consumer should not have to opt in to number-format
  interpretation to get dates as dates, and the default option set must match what the documentation
  promises, since default-argument drift from the docs was part of the reported bug.
- **The parse always terminates.** A well-formed foreign file must never stall the reader. The
  reported stall was a SAX-parser flush bug, a close-tag path returning instead of breaking and
  leaving the parser paused forever, so the rewrite must guarantee monotonic forward progress and a
  bounded, terminating parse, consistent with the fork's hostile-input stance (see the VML and
  comments hang and the bounded-memory notes).
- A streaming read of such a workbook yields **all** rows and sheets, not a truncated prefix.

## Open questions

- Should the streaming reader make styles a hard prerequisite, parsing styles fully before any sheet,
  or interleave with a deferred-resolution buffer keyed by style id? The latter preserves streaming
  for the common case where styles come first.
- Where is termination asserted? A perf or security harness with a wall-clock cap, since a
  potentially-stalling read cannot live in the behavior corpus.
- Do other style-dependent surfaces, number formats generally rather than just dates, share the same
  race, meaning is this a date-specific patch or a general "resolve styles before emit" guarantee?

Related: `streaming-read-applies-date-format` (the locked date behavior),
`streaming-read-resolves-shared-strings`, `streaming-read-emits-all-worksheets`,
`reading-comments-with-vml-drawing-must-terminate`, `bounded-memory-large-workbook-read`.
