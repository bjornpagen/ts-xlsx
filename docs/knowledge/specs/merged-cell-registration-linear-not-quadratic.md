# Merged-cell registration must scale linearly, not quadratically

Cluster: merged-cells / performance

## Scenario

A worksheet contains tens of thousands of merged ranges. Real files reach 30k and more, for example
one small horizontal merge per row across thousands of rows. Loading such a file takes minutes, or
effectively never finishes in constrained or browser environments, even though the data is trivial.
The correctness of merge preservation is covered by the
`many-merged-cells-preserved-and-overlap-rejected` corpus case; this note is about the parse _time_.

> Captured as a spec note, not a corpus case: exercising the real 30k-merge fixture would make the
> corpus run take minutes. A bounded-time check belongs in a perf harness with a hard timeout, not
> the behavior corpus.

## Root cause

The naive loader registers merges one at a time and, for each new merge, scans **all**
previously-registered merges to reject overlaps, an O(merges²) collision check. With tens of
thousands of merges this dominates load time.

## Desired behavior

- Registering merged ranges during load scales roughly linearly, or n·log n, with the number of
  merges rather than quadratically.
- A merge only ever conflicts with a cell already part of another merge, so the collision check
  should consult a **per-cell "already merged" index** where each cell records its master, making each
  registration proportional to its own area rather than to the count of prior merges.
- **Merge _lookup_ is efficient too, not just registration.** Determining whether a given cell
  participates in a merge, and resolving which range it belongs to, master against covered, is a
  frequent operation during read, write, and cell access. It must be an amortized-constant lookup
  against the same index, never a linear scan of all merge ranges per cell, which is the read and
  access sibling of the O(n²) registration cost. The result must be identical to a naive scan: the
  master cell, the covered cells, and the reported bounds are unchanged, and only the time complexity
  improves.

## Open questions

- Is overlap validation even required when parsing a well-formed foreign file? Excel guarantees
  non-overlapping merges in valid OOXML, so load could trust the input and only validate on
  user-driven `mergeCells` calls, keeping a spatial index for the mutation API.
- Should a load-time overlap conflict throw, or be tolerated and repaired, given the fork's
  foreign-generator-tolerance stance? Consider consistency with
  `namespace-agnostic-bom-tolerant-ooxml-reading` and the other tolerance notes.
- What is the target: a memory-bounded per-cell index or an interval tree, and which wins for
  the common "one merge per row" shape?
- Build the index lazily on first merge query, or maintain it incrementally as merges are added and
  removed? And how does it interact with streaming read and write, where the full merge set may not be
  known up front when a cell is first accessed?
- A large-merge-count fixture as a performance regression guard, behind an opt-in slow tag, so the
  quadratic behavior in registration or lookup cannot silently return.

Related: `whole-column-data-validation-bounded-memory` (the same "don't materialize per-cell state
for a range feature" principle).
