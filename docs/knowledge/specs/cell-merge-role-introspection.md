# A cell should report its role in a merge

Cluster: core-model

## Scenario

Code that walks a worksheet's cells frequently needs to know whether a given cell participates in a
merged range and, if so, whether it is the _master_ (top-left, value-bearing) cell or a _child_
(covered) cell. Today that knowledge lives only in the worksheet's merge map; a caller must reach
into that map and re-derive the relationship by hand for every cell, which is the exact workaround
people write when handling merged ranges, for example when deciding whether to render or skip a cell.
The cell already knows enough to answer this; it just does not expose it.

> Spec note, not a corpus case: this is a small model addition (two derived accessors), a
> design decision about what the cell exposes rather than a malformed-output bug. It becomes a corpus
> case once the accessors exist and their round-trip is asserted.

## Desired behavior

- **A cell exposes its merge role directly**: whether it is part of a merge at all, whether it is the
  master, and whether it is a child, derived from the worksheet's merge map rather than stored
  redundantly on the cell.
- The three notions are consistent by construction: a cell in no merge is neither master nor child;
  every merged range has exactly one master; every other cell of the range is a child.
- **A child can reach its master**, and the master its range, so a consumer can fetch the display
  value and style once from the master while iterating children. This complements
  `merged-child-cell-text-mirrors-master`, which locks that a child's _text_ already mirrors the
  master.
- The accessors are cheap, side-effect-free reads, safe to call on every cell of a large sheet.

## Open questions

- Naming and shape: a boolean pair (`isMerged` plus `isMergeMaster`/`isMergeChild`) or a single role
  enum (`none` | `master` | `child`). The enum is harder to misuse, since it has no illegal
  both-true state.
- Whether the master/range back-reference is eager or lazily resolved against the merge map, and how
  it behaves if the merge is mutated mid-iteration.
- Whether these live on the cell only, or also as worksheet-level queries (`worksheet.mergeOf(addr)`).

Related: `merged-child-cell-text-mirrors-master`, `merged-cell-value-and-style-semantics`.
