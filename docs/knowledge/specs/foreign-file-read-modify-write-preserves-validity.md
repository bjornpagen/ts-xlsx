# Read-modify-write of a foreign-generated file must stay valid

Cluster: xlsx-io / robustness

## Scenario

The most basic mutation workflow: open an existing `.xlsx` produced by some other tool, change one
cell's value, and write it back. Reporters find that after this round-trip the output is
**corrupt**, so the application can no longer re-read it and Excel prompts to repair on open. The
corruption is not specific to the edit: reading a foreign file and writing it straight back out,
with no change at all, already produces the broken package. The trigger is the reader dropping or
mangling parts of the source it does not fully model, so the re-serialized package ends up
internally inconsistent (dangling or duplicated relationship ids, a worksheet referencing a part
that was not re-emitted, a shared-strings or table mismatch), even though each individual part looks
plausible.

> Spec note rather than a corpus case: faithfully reproducing this needs a specific foreign-authored
> fixture exhibiting the inconsistency, and the harvested attachments for these reports are no longer
> available. The durable _requirement_ is recorded here; when a concrete corrupting fixture surfaces
> it should be promoted into a corpus case under this spec. The general "spec-built workbook survives
> mutate-then-write" direction is already locked by `themed-workbook-mutate-write-stays-valid`; what
> is distinct here is a _foreign_ source with parts the reader under-models.

## Desired behavior

- **Open-then-save is lossless enough to stay valid.** Reading any well-formed foreign-generated
  workbook and writing it back unchanged must yield a package that re-opens successfully in this
  library and in Excel, with no repair prompt.
- **A single-cell edit preserves validity.** Changing one cell value and writing must not corrupt
  the rest of the document, and unrelated parts survive untouched.
- **Referential integrity is maintained on write.** Relationship ids are unique and every reference
  resolves; the worksheet, shared-strings, styles and table parts remain mutually consistent; parts the
  reader does not model are passed through rather than half-emitted, the same unmodeled-part
  passthrough principle as the drawing, VML and pivot notes.
- Failure, if a source truly cannot be preserved, is loud and specific, never a silently corrupt
  output the caller discovers only when re-reading fails.

## Prior art / notes

- Several reporters traced their specific corruption to an adjacent bug, such as adding a sheet whose
  name already existed, which shows the class is "the writer emits a structurally inconsistent
  package," reachable by multiple paths. So the guarantee is best framed as a package-integrity
  invariant on write, validated by re-reading the written bytes, not as a fix for one path.
- This is the mutation-side companion to the read-side robustness already captured for foreign
  files, where reads must not crash on prefixed namespaces, missing `sheetFormatPr`, missing company
  property, or mixed shared strings.

## Open questions

- Which under-modeled parts most often cause the inconsistency: rels, content-types, calcChain, or
  shared strings? A corpus of foreign fixtures would rank them.
- Should the writer run a cheap self-consistency check (all rels resolve, no duplicate ids) before
  finalizing, and refuse loudly rather than emit a package it can prove is broken?

Related: `foreign-generator-workbooks-read-without-crashing`,
`foreign-generator-boolean-and-mixed-sharedstring`, `themed-workbook-mutate-write-stays-valid`,
`excel-repair-on-open-structural-constraints`, `pivot-table-round-trip-preservation`.
