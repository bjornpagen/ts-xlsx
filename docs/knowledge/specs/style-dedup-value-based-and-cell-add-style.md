# Value-based style deduplication on write, and merging a style onto a cell

Cluster: styles

## Scenario

When writing large spreadsheets, every cell's style must be deduplicated into a shared table of cell
formats (`xf` records) so the same visual style is written once and referenced by index. A common
real-world pattern assigns styles inside loops, as per-cell style literals, or combines column, row
and cell styles, which produces many distinct-but-equal style _objects_. A cache keyed on object
_identity_ then fails to recognize these as equal: it grows unbounded with a near-zero hit rate, and
write time degrades severely. In one measured case a 200k-row by 7-column workbook took roughly 5×
longer to write with identity-keyed style caching than with styling disabled entirely, so the cache
actively hurt. A value-based dedup key, a canonical, order-stable encoding of the style's effective
properties, recovers most of that cost, cutting the penalty to under 2×. Separately, users want to add
or merge a style onto a cell without discarding the column- and row-level styles the cell inherits.

> Spec note, not a corpus case: the durable value is a design proposal, how to deduplicate cell
> styles efficiently and correctly plus a cell-merge ergonomic, not a current behavior to assert.
> The core correctness invariants, that equal styles collapse and distinct styles never collide, are
> flagged as behavior intents for corpus coverage once the write path exists.

## Desired behavior

- **Value-based dedup is the single default.** Two style _values_ that are structurally equal must map
  to the same shared `xf` record regardless of whether they are the same JS object. This makes the
  common patterns, per-cell literals in a loop and column-plus-row-plus-cell combinations, fast and
  memory-bounded with no user tuning. There is **no** user-facing cache-mode knob: an enum choosing
  between correct-and-fast and identity-based caching is exactly the legacy compatibility surface this
  fork rejects, and the canonical value-encoding is the one code path.

- **The dedup key is collision-free.** Two structurally _different_ styles must never encode to the
  same key, or cells would silently inherit the wrong formatting. This is a hard correctness invariant,
  independent of any performance concern, and must hold across the full effective style surface:
  number format, font, fill, border, alignment, and protection.

- **Merge a partial style onto a cell, respecting inheritance.** A caller can add or merge a partial
  style onto a cell so that only the explicitly provided properties override, while the column- and
  row-level properties the cell inherits are preserved. This contrasts with replacing `cell.style`
  wholesale, which drops inheritance. Naming aside, `addStyle` or `mergeStyle`, the semantics are the
  load-bearing part.

- **Styles are values, snapshotted at write.** Value-based dedup encodes each cell's _effective_ style
  at write time, so mutating a shared style object after it was assigned to several cells is a separate,
  documented concern: styles should be treated as immutable values.

## Behavior intents (for later corpus coverage once the write path exists)

- Two cells assigned equal-but-distinct style objects produce a single shared `xf` entry, not two.
- Two cells with genuinely different styles produce two distinct `xf` entries and never collide onto
  each other's formatting after a round-trip.
- Assigning a style inside a loop, with many independent equal style values, does not grow the style
  table proportionally to cell count.
- Merging a style onto a cell preserves the column- and row-level properties it inherits, overriding
  only the explicitly-set properties.

## Open questions

- Define the canonical style encoding precisely over the full effective-style surface so equality is
  total and collision-free, and cover it with type-level and unit tests.
- How to expose the merge operation, a `cell.addStyle(partial)` method or a merge helper, and how it
  composes with the worksheet-default and named-style layers.
- Whether the same value-based encoding should back read-time style _interning_, so equal styles read
  from a file share one model object, not just write-time dedup.

Related: `set-style-over-cell-range`, `cellstylexfs-named-style-fill-roundtrip`,
`worksheet-default-cell-protection-unlock`, `public-type-surface-matches-runtime`,
`column-declared-numfmt-reaches-cells`.
