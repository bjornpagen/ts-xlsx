# Conditional formatting: string-literal quoting and per-cell expression translation

Cluster: conditional-formatting

## Scenario

A user wants a fill applied to cells whose value equals a short string (say `"T"`) via a
conditional-formatting rule. They reach for a `cellIs` rule with operator `equal` and a `formulae`
entry of `T`, but the resulting workbook applies the fill to the wrong cells, notably all blank
cells, because the bare token `T` is emitted as a defined-name or reference rather than a quoted
string literal. Users find that switching to an `expression` rule with a relative formula like
`A2="T"` works, because the relative address, anchored at the range's top-left, is translated
per-cell across the range so each cell is compared to its own value. A separate confusion in the
thread is a viewer rendering difference between applications, not a library defect.

> Spec note, not a corpus case: the durable value is the CF-rule authoring contract, meaning how a
> string literal must be quoted and how relative references in an expression rule are translated,
> rather than a single malformed-output assertion. The corpus already round-trips CF rules
> (`databar-conditional-formatting-roundtrip`, `conditional-formatting-multi-area-ref-survives`,
> `extension-only-conditional-formats-round-trip`); this note pins the semantics.

## Desired behavior

- **`cellIs` string literals are quoted.** For a `cellIs` rule with `equal`, and the other comparison
  operators, an operand intended as a string literal is serialized as an OOXML **quoted string**
  (`"T"`), not a bare token Excel resolves as a defined name or reference. A bare unquoted token
  currently produces a formula that matches empty and blank cells, a silent correctness bug.
- **Expression rules translate relative references per-cell.** For an `expression` rule, formulae with
  relative cell references such as `A2="T"` are anchored at the **top-left cell of the rule's ref
  range** and translated per-cell across the whole range, matching Excel's shared-formula-style
  relative translation for conditional formatting. This is what makes a per-row rule evaluate each
  row against its own cell.
- The API makes the string-literal-against-reference distinction **unambiguous** so a caller does not
  have to know OOXML quoting rules to compare against a literal.

## Open questions

- Does the library auto-quote a `cellIs` operand that is a plain string (least-surprise), require
  the caller to pass a quoted literal explicitly, or offer a typed `{literal: 'T'}` against
  `{ref: 'A2'}` discriminator?
- How is numeric-against-string operand intent disambiguated for `cellIs equal`, meaning
  compare-to-number against compare-to-text?
- Per-cell translation: is the anchor always the range's top-left, and how are absolute `$`
  references left untranslated as expected?

Related: `databar-conditional-formatting-roundtrip`, `conditional-formatting-multi-area-ref-survives`,
`extension-only-conditional-formats-round-trip`, `conditional-format-numfmt-roundtrip`.
