# Dynamic-array (spill) formulas must not be downgraded to legacy CSE array formulas

Cluster: formulas

## Scenario

A workbook authored in a modern spreadsheet application contains a **dynamic-array formula** in a
single cell, for example a `LET` wrapping `FILTER`:

```
=LET(results, FILTER(Sheet3!$B$2:$B$1000, VALUE(Sheet3!$A$2:$A$1000) = VALUE(INDIRECT("Sheet1!S" & Sheet2!B1))), IFERROR(results, ""))
```

In modern spreadsheet semantics this is a **plain single-cell formula** that automatically _spills_
its multiple return values into the neighbouring cells of the column. It is **not** a legacy
Ctrl+Shift+Enter (CSE) array formula and carries no surrounding braces. A user reads such a workbook,
does nothing to that cell, and writes it back out. On reopening, the formula has been wrapped in
`{ … }` and now behaves as a legacy single-cell array formula: it no longer spills, and only the
top-most value is returned. The rewrite silently downgraded a dynamic-array formula to a CSE array
formula, destroying the spill.

> Spec note, not a corpus case: the library has no first-class concept of a dynamic-array (spill)
> formula distinct from a legacy CSE array formula, so this is a **model gap**. The fix is a new
> distinction in the formula model and its serialization, not a one-line assertion against current
> behaviour. It also needs a fixture authored by a real spreadsheet application, since the trigger is
> the dynamic-array metadata Excel writes, which we do not yet have. It becomes a corpus case once the
> model distinguishes the two formula kinds and a round-trip asserts the emitted formula element stays
> plain, with no braces and no single-cell `t="array"`.

## Desired behavior

- **The two formula kinds are distinct in the model.** A dynamic-array (spilling) formula and a legacy
  CSE array formula are represented differently, because they serialize differently and behave
  differently in the host application. Reading a plain cell formula must not classify it as an array
  formula.
- **A dynamic-array formula round-trips as a plain cell formula.** Read then write with no edit to the
  cell must emit `<f>…</f>` with the formula text verbatim and **no** wrapping braces and **no**
  `t="array"` anchored to a single cell. The spill behaviour survives because the emitted element is
  unchanged in kind.
- **Formula text is preserved verbatim**, including `LET`, `FILTER`, `INDIRECT` and nested references,
  with no re-quoting and no reference rewriting on an untouched cell.
- **A genuine CSE array formula still round-trips as an array formula**, so the legacy shape is not
  lost either, and the two are preserved as authored, each in its own kind.

## Open questions

- OOXML surface: modern dynamic-array formulas are plain `<f>` elements, with the spill metadata
  living in the cell-metadata, `metadata.xml` and `xlrd`/`xlda` rich-data parts rather than on the
  formula element itself. How much of that metadata must be preserved for the host to treat the
  formula as dynamic rather than legacy? Is keeping the formula plain, with no `t="array"`, sufficient,
  or must the cell metadata be carried through too?
- Detection on read: how is a dynamic-array formula recognised on input, purely by the absence of
  `t="array"`, or by the presence of the dynamic-array cell metadata? The reader must not manufacture
  an array wrapper that was never in the source.
- Authoring: should the public API let a caller _create_ a dynamic-array formula (spilling) distinctly
  from a CSE array formula, or is preservation-on-round-trip the first milestone?
- Rich-value errors ride the same metadata. Excel 16.0 (build 20326) stores `#SPILL!` (errorType 8),
  `#CALC!` (13) and the other newer errors as `<v>#VALUE!</v>` with a `vm` index pointing through
  `xl/metadata.xml` at a rich value of structure `_error` in `xl/richData/rdrichvalue.xml`, and a
  spilling formula's cell carries `cm` and `vm` together (recorded in
  `test/corpus/fixtures/excel-oracle/error-literals.json`). Such a cell reads as `#VALUE!` today and the
  `vm` link and the rich-data parts are dropped on write, so the error belongs with the formula model:
  the reader that resolves `cm` through the metadata is the one that resolves `vm`, and the errorType
  numbers of the errors not yet observed must come from Excel rather than be guessed.

Related: `formula-cell-value-type-minimal-required-fields`, `formula-recalculation-expectations`,
`shared-formula-master-survives-roundtrip-and-splice`.
