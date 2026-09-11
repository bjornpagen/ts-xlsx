# Duplicating and merging worksheets must be a safe, first-class, deep-copy operation

Cluster: core-model / packaging

## Scenario

Two closely related needs recur with no first-class API:

- **Clone within a workbook.** A user has a sheet, often a loaded template, and wants an independent
  copy to edit, or wants to stamp out N copies of a template sheet into the same workbook.
- **Merge across workbooks.** A user builds sheets in separate workers or processes, to parallelize a
  large workbook, and wants to assemble the finished sheets into one output file, or simply combine
  two workbooks into one.

With no supported operation, users approximate it two ways, both unsound. Cloning by overwriting a
blank sheet's serialized `model` with the source's model: assigning by reference, or by shallow copy,
shares the `name` and triggers a duplicate-name error, or leaves workbook metadata inconsistent so
the file opens with a recovery prompt, for example a named range still pointing at the original sheet.
A deep clone with the name reconciled works for the common case, but reporters find **embedded images
and merged cells do not reliably carry over**, and cross-workbook cloning is unreliable. Merging by
iterating a source sheet's rows and re-adding them to a destination fails too: the object read back
from a row is not the shape the add-row API consumes, and copying only values silently drops
styling, column widths, merges, notes, images, tables, and data validations.

> Spec note, not a corpus case: there is no clone or merge API yet, so there is nothing to assert
> against. The `dst.model = {...src.model}` workaround's specific merge-loss is captured as a corpus
> case (`worksheet-model-preserves-merged-cells`); this note records the durable _requirement_ for
> the real operation that should replace the workaround.

## Desired behavior

- A first-class, safe way to **duplicate a worksheet** into a workbook, cloning under a new,
  auto-deduplicated name, and to **append or merge** a worksheet from another workbook. The clone is a
  **deep, independent copy**, so mutating the copy never affects the source.
- **All sheet content survives**, not just cell values: values and formulas including shared
  formulas, number formats, fonts, fills, borders and alignment, column widths and hidden state, row
  heights, **merged ranges**, notes and comments, **images and media**, tables, and data validations.
- **Workbook-scoped state is reconciled** so the result opens without repair: shared-string entries
  remapped or rebuilt; style indices translated into the destination's style table, deduplicating
  identical ones; sheet ids and names kept unique under a defined collision policy; relationship ids
  for images, tables and hyperlinks re-issued so they never clash with the destination's existing rels;
  defined names, sheet-scoped references, and print areas re-pointed or duplicated rather than left
  dangling at the original sheet.
- **Cross-sheet references** a moved sheet depends on are handled by a documented policy. At minimum,
  formulas referencing the moved sheet by its possibly renamed own name stay valid.
- **Name collisions never silently corrupt**: either reject with a clear error or auto-suffix
  deterministically.
- Optionally, a **portable single-sheet serialization** so sheets built in separate processes can be
  marshalled and reassembled; if exposed, it must round-trip losslessly through append and merge.

## Prior art / observed workarounds

- Deep-clone the serialized worksheet model, overwrite `name`, assign it back. Works for values,
  columns and styles, but **loses merges and media**, and corrupts named ranges if metadata is not
  reconciled.
- The row-by-row re-add approach misuses the add-row API, since a read-back row is not an add-row
  input, and drops all formatting. Any merge API must not rely on that pattern.
- The underlying "large single-sheet generation is slow" complaint is a _separate_ concern; merge is
  requested to parallelize, not as a performance guarantee the library must make.

## Open questions

- Public surface: a workbook-level `cloneWorksheet(source, {name})` returning the new sheet, and a
  separate `importWorksheet(fromWorkbook, sheet)`?
- Collision policy for sheets and defined names: throw, auto-suffix, or caller-specified, and is it
  configurable?
- Sheet-scoped defined names and inter-sheet formula references pointing at the source: rewrite to
  the clone, or leave pointing at the original?
- Cross-workbook merge of artifacts that are not sheet-scoped (theme, calc properties, workbook views,
  workbook-level defined names): merge, prefer-destination, or error on conflict?
- Is a streaming or worker-friendly portable-sheet serialization in scope now, or a later addition?

Related: `worksheet-model-preserves-merged-cells` (the model-copy merge-loss locked as a case),
`foreign-file-read-modify-write-preserves-validity`, `defined-name-scope-must-be-per-sheet`,
`excel-repair-on-open-structural-constraints`, `streaming-write-per-sheet-memory-release`.
