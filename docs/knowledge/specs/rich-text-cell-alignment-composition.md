# How rich text composes with cell alignment: alignment is cell-level, runs are character-level

Cluster: styles

## Scenario

An author builds a cell whose value is rich text, several runs with different fonts, and wants to
control its alignment: horizontal, vertical, wrap, indent, rotation. They reach for a per-run
alignment property and find it has no effect, and are unsure where alignment belongs when the value
is composed of several runs. The confusion is structural: alignment is a property of the _cell_, not
of any run, but the rich-text shape, an array of runs each with its own formatting, invites the
assumption that layout can be set per run.

> Spec note, not a corpus case: this is a how-to and design question with no failing serialization to
> baseline. The durable value is pinning where each property legitimately lives, so the type surface
> and docs make the correct usage the easy path. A corpus case can lock "cell alignment survives on a
> rich-text cell" once the declarative spec's cell shape can express a `richText` value.

## Desired behavior

- **Cell alignment governs the whole composed rich-text string.** When a cell's value is rich text,
  the cell-level alignment (`horizontal`, `vertical`, `wrapText`, `indent`, `textRotation`) set on the
  cell's style applies to the entire composed string exactly as it does for a plain-string cell. Rich
  text does not disable or override cell alignment.
- **Runs carry only character-level formatting.** A rich-text run's valid properties are character
  formatting: font `name`, `size`, `color`, `bold`, `italic`, `underline`, `strike`, and
  `vertAlign` for superscript and subscript. These map to the OOXML run-properties element (`<rPr>`),
  which has no layout or paragraph-alignment attributes.
- **Layout-shaped properties on a run are meaningless and must not silently "work".** Horizontal,
  vertical, wrap, indent and rotation are not valid run properties. Placing one on a run should be a
  no-op at minimum, and ideally rejected by the type surface, with a precise `RichTextRun` type that
  does not admit alignment fields, so a caller learns at compile time that alignment belongs on the
  cell rather than setting it per run, seeing no effect, and assuming a bug.

## Open questions

- Type surface: make run alignment _unrepresentable_, narrowing `RichTextRun` to character-format
  fields only, or accept and ignore it? The former prevents the confusion class at compile time for
  TypeScript callers and is preferred.
- `vertAlign` overlap: a run's `vertAlign`, superscript or subscript, is a legitimate character
  property and must not be confused with cell _vertical alignment_, so the naming should keep them
  distinct.
- Read-back: confirm a rich-text cell read from a file surfaces cell alignment on the cell and only
  character formatting on each run, with no phantom per-run alignment.

Related: `rich-text-cell-value-writer-robustness`, `hyperlink-display-text-can-be-rich-text`,
`html-fragment-to-rich-text-cell-value`, `public-type-surface-matches-runtime`,
`rich-text-run-font-is-all-or-nothing`, the _font_ half of the same "what does a run inherit from
its cell" question, which answers differently: alignment lives on the cell and reaches the whole
value, while a run's font inherits nothing at all.
