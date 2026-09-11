# AutoFilter on a table is explicitly controllable, independent of the header row

Cluster: tables

## Scenario

A developer wants a worksheet table purely for its consistent styling, banded rows and a table
theme, with no filter dropdowns and, in some cases, no header row at all: a clean data block that
Excel opens without a repair prompt. Two related frustrations recur:

1. A **header-less** table (`headerRow: false`) still gets an AutoFilter emitted against the
   now-absent header, producing an internally inconsistent table part that Excel repairs on open.
   That specific corruption is already pinned by the corpus case
   `table-headerless-omits-autofilter`, where a header-less table must set `headerRowCount=0` and
   emit no `autoFilter`.
2. Even a **header-bearing** table has its AutoFilter forced on implicitly, with no way to say
   "give me a styled table with a header but no filter UI." The only escape is to abandon the table
   feature and hand-style cells.

The durable question this note captures is the missing _design decision_: filter presence should be
an explicit, first-class choice, not an implicit consequence of whether a table has a header.

> Spec note, not a corpus case: the corruption path (header-less plus forced filter) is already locked
> as a corpus case. What remains is an API-shape decision, how a caller expresses "table without a
> filter", which is design material rather than a malformed-output bug.

## Desired behavior

- **Filter emission follows an explicit flag, not the header's presence.** A table can be written
  with a header and no AutoFilter, with no header and no AutoFilter, or with a header and an
  AutoFilter. All three are expressible and all three produce valid OOXML Excel opens without
  repair.
- **`headerRow: false` conflates two concerns today**, header presence _and_ filter presence, so
  separate them and make the two axes independent. A header-less table never carries an AutoFilter,
  since it is not legal there; a header-bearing table carries one only when the caller asks.
- **Round-trip preserves both axes**: a table written header-less stays header-less on reload, and a
  table written filter-less does not silently regain a filter.
- **Styling stays valid without a header**: banded-row and first-column style flags remain meaningful
  on a header-less table, and the table style definition is still emitted.

## Open questions

- The public API: a dedicated `filterButton` or `autoFilter` boolean on the table definition, against
  inferring from `headerRow`. A separate flag is clearer and lets "header plus no filter" exist.
- Default behavior when the flag is unset: match Excel, where a header-bearing table defaults to
  showing filter buttons, but never emit a filter for a header-less table. Document the chosen
  default.
- Interaction with a totals row: confirm totals-row handling stays consistent when the filter is
  absent.

Related: `table-column-width-in-definition`, `streaming-writer-table-support`,
`table-handle-direct-property-access`.
