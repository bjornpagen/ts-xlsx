# The streaming writer's worksheet type is a narrowed, forward-only reflection

Cluster: types

## Scenario

A TypeScript caller opens a streaming, forward-only workbook writer and adds a worksheet, expecting
the same rich worksheet the buffered API returns. The published types say so, typing the streaming
worksheet as the full random-access worksheet, but at runtime it is a narrower,
forward-only object. Methods the type promises (`insertRow`, and other random-access, splice and
back-reference operations) either do not exist or throw once earlier rows have been flushed. The
type lies, so the compiler green-lights code that fails at runtime.

The same drift appears at the workbook level: the streaming writer is typed as if it extended the
buffered workbook, but it does not. It is a distinct forward-only surface with its own, smaller set
of capabilities.

> Spec note, not a corpus case: the defect is a type-surface fidelity gap on a forward-only object,
> pinned by type-level tests, not a malformed-output bug reproducible from a data file.

## Desired behavior

- **The streaming worksheet has its own precise type** that exposes _only_ what a forward-only writer
  can honor. Random-access operations that require rewriting already-flushed rows (inserting or
  splicing a row above the write frontier, re-reading a committed cell, arbitrary back-reference) are
  **absent from the type**, so misuse is a compile error rather than a runtime surprise.
- **The streaming workbook writer is typed as a distinct surface**, not as a subtype of the buffered
  workbook. Shared capabilities may come from a common minimal interface, but the writer never claims
  operations it cannot perform on a forward-only stream.
- **Commit and forward-only semantics are visible in the types**: adding rows and cells and committing
  a worksheet or the workbook are the vocabulary, and the forward-only constraint is expressed by
  omission, so the ergonomic loop (stream rows, commit, move on) is exactly what the types make easy.
- Where a buffered-only operation has a streaming analogue with different mechanics, the analogue is
  named honestly rather than aliased to the buffered method it does not fully implement.

## Open questions

- The exact minimal interface shared between buffered and streaming worksheets, against what stays
  buffered-only, driven by the unification question in `unified-streaming-and-buffered-io`.
- Whether any "insert above the frontier" operation is supportable at all on a forward-only writer,
  via a buffer-until-commit window, or is categorically out of scope.
- Naming for the streaming worksheet and writer types so the forward-only nature is discoverable at
  the call site, not just in prose.

Related: `unified-streaming-and-buffered-io`, `public-type-surface-matches-runtime`,
`streaming-write-sheet-protection-before-autofilter`, `write-buffer-return-type-contract`.
