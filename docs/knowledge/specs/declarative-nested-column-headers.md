# Declarative multi-level (nested) column headers

Cluster: core-model / DX

## Scenario

Reports frequently need **grouped, multi-level headers**: a top banner ("Q1") spanning several
sub-columns ("Jan", "Feb", "Mar"), possibly several levels deep. Producing these today means
dropping to low-level primitives, computing how many leaf columns each group spans, writing each
header row by hand, and manually merging every group cell across its span, for what is a
structurally regular, mechanical layout. The width of a parent header is just the aggregated leaf
count of its descendants, and the depth of the header block is the tree height. Doing this by hand is
error-prone (off-by-one spans, forgotten merges) and does not compose with dynamic data whose
grouping is only known at runtime.

## Desired behavior

- A **declarative** way to define columns as a tree, where each node has a label and either child
  nodes or a leaf binding (a `key` or value accessor plus optional width and style). The library walks
  the tree, computes each node's span as the aggregate leaf count of its subtree, emits one header row
  per tree level, and merges each non-leaf header cell across its span automatically.
- Leaf columns produced by the tree behave like ordinary columns for everything downstream: keyed
  row population, widths, number formats, and cell styling all apply to the leaves.
- The generated header block composes cleanly with the rest of the sheet: data rows begin
  immediately below the deepest header row, and freezing among the header rows works as if they had
  been authored by hand.
- Styling applies per node, so a group header and its leaves can be styled independently, and the
  merged group cells carry borders and fills consistently with adjacent groups.

## Prior art / notes

- The proposed shape is a recursive column tree walked into a flat leaf list plus a per-level
  header plan; that flat list is exactly the existing column model, so nested headers can be a
  _builder_ over the current primitives rather than a new core concept.
- The hard part is interaction with the rest of the feature set (merged cells, borders on merged
  regions, per-node styles, and shared column keys) which is why this belongs in the design phase
  rather than as a mechanical corpus lock.

## Open questions

- Tree input shape: nested objects and arrays, or a fluent builder? How are leaf `key`s spelled
  (see `nested-property-path-column-keys` for the record-binding side)?
- How do borders and fills resolve across a merged group header: inherited from the group node,
  or composed from the leaves?
- Does the feature own row freezing of the header block, or leave that to the caller?

Related: `nested-property-path-column-keys`, `column-key-roundtrip-persistence`,
`many-merged-cells-preserved-and-overlap-rejected`.
