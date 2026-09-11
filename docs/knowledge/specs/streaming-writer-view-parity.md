# The streaming writer must honor sheet views assigned after a worksheet is created

Cluster: streaming

## Scenario

A caller writes a workbook through the streaming writer, adds a worksheet, and then, because the
split position is computed dynamically and is not known at the moment the sheet is created, assigns
the worksheet's view configuration, say a frozen split at column 3, row 4, _after_ the worksheet
already exists. The buffered, in-memory writer honors both paths: views passed as a construction
option and views set on the property afterward. The streaming writer does not. Its worksheet view
configuration is fixed at creation time and cannot be changed later, so the natural
"add the sheet, then decide the panes" pattern silently loses the frozen split.

> Spec note, not a corpus case: probing the current streaming writer shows the worksheet's `views` is
> exposed as a **getter only**, so assigning to it after creation throws a `TypeError`
> (`Cannot set property views of #<WorksheetWriter> which has only a getter`) rather than accepting
> and then dropping the value. So there is no silent wrong-output to baseline: the after-creation
> authoring path does not exist at all. Views supplied as an `addWorksheet(name, { views })`
> construction option _do_ work and emit a correct `<pane>`, verified. The durable value is the
> parity requirement and the API-shape decision, which a corpus case can lock once the property is
> settable.

## Desired behavior

- **Writer parity.** The streaming writer and the buffered writer produce the same sheet-view result
  for a given view configuration, regardless of _when_ the caller supplies it. Assigning `views`
  after `addWorksheet` must be honored, exactly as passing it at construction time is, up until the
  worksheet is committed. After commit the header is already serialized and the configuration is
  necessarily frozen, and that boundary should be a clear error rather than a silent drop.
- **Frozen and split panes survive.** A frozen split (xSplit/ySplit with `state: 'frozen'`) set either
  way emits the same `<pane>` in the sheet view and reads back with the same split on load.
- **Consistent property surface.** The streaming worksheet's `views` should be a settable property
  before commit, matching the buffered worksheet's surface, so code that configures views is portable
  between the two writers without special-casing which one it targets.

## Open questions

- The commit boundary: is the intent to allow `views` assignment only before the first row or commit
  is flushed, or to buffer the sheet-view header until commit so it can be set at any point
  beforehand? Define and document the exact cutoff, and make crossing it a diagnostic rather than a
  getter-only `TypeError`.
- Should the construction-time `{ views }` option and the post-creation property assignment merge, or
  does the later assignment replace wholesale? Match whatever the buffered writer does.
- Portability: audit the rest of the streaming worksheet surface for the same getter-only asymmetry,
  properties settable on the buffered worksheet but frozen on the streaming one, so this is fixed as
  a class rather than a one-off.

Related: `streaming-writer-row-commit-backpressure`, `sheetview-boolean-flags-and-showformulas`,
`streaming-write-per-sheet-memory-release`.
