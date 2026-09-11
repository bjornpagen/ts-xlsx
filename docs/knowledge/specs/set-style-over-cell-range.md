# Apply a style to a whole cell range in one call

Cluster: styles

> **Resolved.** `Worksheet.getRange` hands back a `Range` handle carrying the six style facets plus a
> composing `style` accessor and a `clearStyle()` verb. See `src/core/range.ts`, and the
> `range-styling-covers-empty-cells-and-still-dedups` corpus case for the locked behaviour. Every
> open question below is answered; the answers, and where they differ from this note's leanings, are
> recorded under each.

## Scenario

A user wants to apply a fill, or any style facet, uniformly to a rectangular block of cells, shading
a header band or highlighting a region, **without iterating cell-by-cell**. Today the only path is to
loop over every row and column in the range and assign the style to each cell individually: verbose,
and easy to get wrong, with off-by-one bounds and forgotten materialization of empty cells so an
unwritten cell in the block silently stays unstyled.

> Spec note, not a corpus case: this is an ergonomic API request with no failing behavior to assert,
> since the per-cell loop already works. The durable value is the desired convenience API and its
> design questions; the eventual behavior is assertable through the existing per-cell fill and style
> checks once the API exists, so no new adapter capability is needed now.

## Desired behavior

- A **range-styling call** that takes a range reference (`"B2:D5"`, or a `{tl, br}` pair) and a style
  payload and applies it to every cell in that block in one operation, including cells that were
  previously empty, which are materialized as styled-but-valueless rather than skipped.
- **Merge against replace is explicit.** The default composes the payload onto each cell's existing
  style, so setting a fill does not wipe fonts, borders or numFmt already there, and a replace mode
  overwrites the whole style. The caller chooses; the library does not guess.
- **Bounded and predictable.** Styling a large range does not silently balloon the style table,
  because it benefits from the same style deduplication every write path uses, so a uniform block
  resolves to one shared style index (see the dedup case).
- **Composable with existing single-cell styling**, so a subsequent per-cell edit inside the range
  overrides just that cell, and range styling never corrupts merged-cell masters within the block.

## Open questions

- ~~Surface shape~~ **Answered: a first-class `Range`**, the third handle beside `Row` and `Column`,
  with the same live-view contract. There is no `setRangeStyle` helper beside it, one surface rather
  than two, and having a real type is what leaves room for range-scoped values and formulas later.
- ~~Merge-against-replace default~~ **Answered: the same semantics `Cell` already has**, which turned
  out to settle the question rather than trade it off. `cell.style = {...}` _already_ composes facet by
  facet, laying on each facet named and leaving the rest, while `cell.fill = x` replaces that facet.
  `Range` mirrors both exactly, so there is no second convention to learn and no mode flag. Wholesale
  replace is `clearStyle()` then assign.
- ~~Empty-cell materialization~~ **Answered: eager, and the cost cliff is closed off at the door.**
  A styled-but-valueless cell is the only way an empty cell renders with a fill, so the holes have to
  be filled. The 10,000-row worry is answered by _refusing_ the shape that causes it: `getRange` takes
  only a bounded rectangle, and `A:A` and `1:1` throw, pointing at `getColumn` and `getRow`, which say
  the same thing in one `<col style>` or `<row s>` attribute. `Range.cellCount` is the exact number of
  cells a write will create, before it creates them.
- ~~Interaction with column and row-level styles and merged ranges~~ **Answered.** The column and row
  layer is where whole-axis defaults belong and is now the _only_ way to state one, so the two cannot
  compete. A block overlapping a merge resolves each address through `getCell`, which lands on the
  region's master, so the master is restyled and no covered position gains a cell whose style the
  serializer would then have to drop.
- ~~Address-iteration primitive~~ **Answered: yes, typed and public.** `Range.addresses()` is a
  generator, so walking a block materializes nothing and can be abandoned part-way. The eager cost
  this note flagged is a property of _styling_, not of iterating. The facet setters do not go through
  it as sugar; they materialize directly.

Related: `shared-styles-deduplicated-in-written-package`, `per-cell-fill-isolation`,
`per-cell-font-isolation`, `column-level-value-type`, `worksheet-get-columns-range-accessor`,
`declarative-nested-column-headers`.
