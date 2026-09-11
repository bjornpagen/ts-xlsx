# Streaming reader options: styles caching, its default, and the "dates are a style" footgun

Cluster: streaming

## Scenario

A user reads a workbook with the streaming reader and finds that date cells come back as raw serial
numbers instead of dates, or that number formats are missing. The cause is non-obvious: the streaming
reader defaults some caching options **off** for memory reasons, and, critically, a cell's date-ness
is carried by its _style_, the `numFmt`, not by the cell value. So without style caching enabled the
reader has no way to know a numeric cell is a date, and every date silently reads as a bare number.
The option name (`styles`) does not hint that turning it off also disables date interpretation, and the
published docs and types disagreed with the implementation about what the defaults even are.

> Spec note, not a corpus case: the streaming-read-with-styles behavior is already exercised elsewhere
> in the corpus. The durable value here is the _option contract_, meaning precise names, precise
> defaults, and the documented consequence of each, which is an API-design decision to settle in
> Phase 3 rather than a current-behavior assertion.

## Desired behavior

- **Option names and defaults are documented and match the implementation exactly.** The streaming
  reader's construction options (worksheet emission mode, shared-strings caching, styles caching,
  hyperlink handling, …) each have a stated default, and the types, docs, and runtime agree. A doc or
  type that disagrees with the code is itself a defect.
- **The "styles" option's effect on values is spelled out.** Because number formats live in the style
  table, style resolution governs whether date and number cells are interpreted, with a serial number
  becoming a `Date`, not merely whether visual formatting is available. The contract states this
  explicitly so a caller disabling styles knows they are opting out of typed dates, not just cosmetics.
- **The defaults are chosen to avoid the silent-wrong-data footgun.** Defaulting styles-caching off
  trades a memory saving for silently mis-typed dates, a poor default for a correctness-first library.
  The rewrite should either default to correct typed reads, or make the degraded mode loud, with a
  typed result that marks such cells as "raw, unresolved" rather than handing back a
  plausible-but-wrong number, so a caller cannot get wrong data without having chosen it.
- **Memory-conscious modes remain available and explicit.** Callers who genuinely want the
  minimal-memory, values-only stream can still ask for it by name, accepting the documented
  consequences.

## Open questions

- The exact option surface and defaults for the rewritten streaming reader, and whether date typing is
  decoupled from full style retention, caching just the numFmt-to-isDate mapping cheaply without
  holding the whole style table, so correct dates need not cost full style memory.
- Whether the degraded, styles-off path returns raw values, marks them unresolved, or is removed
  entirely in favor of an always-correct-but-tunable-memory reader.
- How this interacts with the shared-strings caching option and the overall streaming memory budget.

Related: `streaming-read-emits-all-worksheets`, `streaming-read-resolves-shared-strings`,
`streaming-read-styles-before-cells`, `dates-round-trip-with-number-format`,
`public-type-surface-matches-runtime`.
