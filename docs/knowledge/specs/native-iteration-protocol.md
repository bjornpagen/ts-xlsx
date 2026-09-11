# Native iteration protocol on collection-shaped API objects

Cluster: api-ergonomics

## Scenario

The library's collection-shaped objects (a row's cells, a column's cells, a worksheet's rows and
columns, the workbook's sheets, the defined-names collection, the internal cell matrix) currently
expose only callback-based traversal via `each*` and `forEach` methods. Callbacks cannot be `await`ed,
cannot `break` early without throwing, and don't compose with the language's built-in iteration
facilities. A caller who wants `for (const cell of row) { … }`, to spread cells into an array,
destructure, or short-circuit on the first match has no ergonomic way to do so. The ask is that these
collections implement the standard JS iteration protocol so idiomatic `for...of`, spread, early
`break`, and, for the streaming reader, `for await` all work directly.

> Spec note, not a corpus case: this is an API-shape proposal, not a bug with a reproduction. There is
> no failing current behavior to baseline, since the callback forms work and the gap is an ergonomic
> layer on top of them. In a TS-first fork iteration should likely be the _primary_ traversal API,
> which is a Phase 3 design decision. It becomes corpus-covered by type-level and in-memory unit tests
> once the iterators exist.

## Desired behavior

- **Synchronous iteration on in-memory collections.** A row, a column, a worksheet's rows, a
  worksheet's columns, the workbook's sheets, and the defined-names collection each implement
  `[Symbol.iterator]`, so `for...of`, array spread (`[...row]`), and destructuring work directly.

- **Lazy, short-circuitable.** `break` or `return` inside a `for...of` loop stops traversal
  immediately, a strict improvement over callback `each`, where early exit requires throwing.

- **Async iteration on the streaming reader.** The streaming worksheet reader implements
  `[Symbol.asyncIterator]` so `for await (const row of stream)` works, serving the async-consumption
  motivation that synchronous callbacks cannot. This is the natural fit for incrementally-arriving
  rows and composes with the existing streaming read API.

- **Sparse against dense is explicit, not an argument.** Iteration must honor the established
  "populated cells only by default" contract, but `[Symbol.iterator]` is invoked by the runtime with
  no arguments and cannot receive options during `for...of`. So the include-empty variant is a
  _distinct_ method: a default `[Symbol.iterator]` over populated entries plus an explicit
  `cells({ includeEmpty: true })` generator, not a single option-bearing iterator.

- **Coordinates are preserved.** Consumers of the callback form relied on the second
  `(cell, colNumber)` or `(row, rowNumber)` argument. With iteration, each yielded element carries its
  own 1-based coordinate (`cell.col`, `row.number`), and an `entries()` iterator yields
  `[index, element]`, so no positional information is lost.

## Open questions

- Do we keep the legacy callback `each*` methods at all, or make iteration the sole traversal API?
  Fork policy permits dropping them; decide by whether thin wrappers (iterate then invoke) cost
  anything.
- Per-collection default emptiness policy: rows and cells currently differ in their `includeEmpty`
  defaults. Fix a single documented default per collection rather than inheriting the inconsistency.
- Column iteration: key on defined column keys, or on every occupied column? Confirm the durable
  contract; the proposal flags key-based.
- Whether to provide full `keys()`, `values()` and `entries()` triplets for Map and Array parity, or
  only the default iterator plus `entries()`.

Related: `async-iterable-types-compile-cleanly`, `row-iteration-early-termination`,
`row-values-no-phantom-leading-slot`, `worksheet-columns-mutable-array-ergonomics`,
`streaming-read-emits-all-worksheets`, `public-type-surface-matches-runtime`.
