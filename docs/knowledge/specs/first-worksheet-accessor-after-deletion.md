# A deletion-safe accessor for the first existing worksheet

Cluster: types

## Scenario

A user opens or builds a workbook, removes one or more worksheets, then wants to operate on "the
first sheet." Worksheet ids are stable identifiers that are intentionally _not_ renumbered when a
sheet is removed, so indexing the workbook by id 1 returns nothing once the sheet that originally
held id 1 has been deleted, even though other sheets still exist. Callers routinely conflate "id 1"
with "the first sheet" and are surprised when the by-id lookup yields undefined after a deletion.
They want a straightforward, deletion-safe way to reach the first surviving worksheet without
knowing which ids remain.

> Spec note, not a corpus case: this is an API-ergonomics and type-surface proposal with no failing
> runtime behavior to baseline, since the underlying model (stable ids, ordered positions) is already
> correct. The gap is a convenience-and-clarity layer on top of it. The upstream proposal adds a
> new public method; that diff targets code we discard, so what is durable is the intent and the
> model distinction, recorded here for Phase 3 design.

## Desired behavior

- **A reliable, deletion-safe first-worksheet accessor.** Obtaining "the first existing worksheet",
  the one with the lowest position in sheet order among sheets that still exist, must keep working
  after arbitrary deletions, independent of any sheet's stable id.
- **An honest empty-workbook contract.** When the workbook has no worksheets at all, the accessor
  returns an explicit empty result (`undefined`) rather than throwing, the natural, precisely-typed
  shape for a strict-typed API.
- **Id and order are distinct operations, and the API must make that obvious.** A worksheet's id is a
  stable handle deliberately not reassigned when other sheets are removed; a worksheet's
  position or order is contiguous. Accessing by id and accessing by order are both legitimate; the
  pitfall is conflating them. Order-based access must be the obvious, hard-to-misuse path, and the
  type and name surface should keep a caller from accidentally passing a position where an id is
  expected (see `indexing-convention-accessor-naming`).

## Open questions

- A dedicated first-sheet accessor, or a clearly-named order-indexed accessor (a worksheet-by-
  position member, explicitly documented as to base) that also covers the general "nth existing
  sheet" case? The latter may subsume this need with a smaller API.
- The empty-workbook contract: `undefined` against an option that throws. Pick one honest, precisely
  typed shape; `undefined` is the natural fit.
- Whether to disambiguate id-based from order-based access at the type or name level so the two cannot
  be confused at a call site.
- Whether the active or first-_visible_ sheet concept (active tab, `veryHidden` and `hidden` state)
  should factor in, since "first worksheet" and "first visible worksheet" can differ and consumers
  often want the latter.

Related: `indexing-convention-accessor-naming`, `worksheet-columns-mutable-array-ergonomics`,
`worksheet-name-existence-check-robust-against-array-prototype-pollution`,
`public-type-surface-matches-runtime`, `worksheet-hidden-state-preserved-on-write`.
