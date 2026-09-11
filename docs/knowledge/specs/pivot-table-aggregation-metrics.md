# PivotTable per-field aggregation metrics (Count and the rest, not just Sum)

Cluster: pivot

## Scenario

A user builds a pivot table and needs to choose, per value field, how it is aggregated. Spreadsheet
applications offer a family of aggregation functions (Sum, Count, Average, Max, Min, Product, and
more) and the pivot definition records which one applies to each value field. A library that can only
sum cannot express the very common case of **counting occurrences** in a field, meaning how many rows
fall into each category, which is the natural reduction for non-numeric fields where Sum is
meaningless.

> Spec note, not a corpus case: this extends the partial and unbuilt pivot-authoring feature, so the
> durable value is the per-field aggregation model and its OOXML mapping, not an assertion against
> current behaviour. It becomes a corpus case once the authoring API accepts a metric per field and the
> written pivot definition's `dataField/@subtotal` is asserted. This complements the _multi-measure_
> axis captured in `pivot-table-multiple-value-fields`, whose open question about per-measure
> aggregation functions is answered here.

## Desired behavior

- **Each value field carries its own aggregation function.** The authoring API lets a caller specify,
  per data or value field, which reduction to apply: at minimum **Count** in addition to **Sum**, and
  ideally the full set the format allows.
- **The chosen function is serialized into the pivot definition** so the host shows the correct
  reduction and recomputes it on refresh, rather than silently defaulting everything to Sum.
- **Counting works on non-numeric fields.** Count applies to text fields, where Sum is meaningless,
  and the two counting variants are distinguished (see below).
- **Sum remains the default** when no metric is specified, matching the format default and preserving
  existing single-measure behaviour.

## Prior art (OOXML)

Each value field in the pivot table definition (`pivotTableDefinition/dataFields/dataField`) carries a
`subtotal` attribute whose enumerated values name the aggregation function:

`sum`, `count` (counts all non-empty values, including text), `average`, `max`, `min`, `product`,
`countNums` (counts only numeric cells), `stdDev`, `stdDevp`, `var`, `varp`.

The default when the attribute is omitted is `sum`. The two counting variants matter: **`count`**
counts every non-empty cell and works on text fields, while **`countNums`** counts only numeric cells.
The base pivot-cache field type may also constrain what a host will meaningfully compute for a given
metric.

## Open questions

- API shape for expressing the metric per field: a string enum on each configured value field, or a
  structured object? And how to handle an unknown or invalid metric name: reject at author time, or
  pass through?
- Whether to validate that Count-family metrics are compatible with the underlying field data, or
  leave that judgement to the host application.
- How the pivot cache is populated for count against sum, and whether the emitted definition needs a
  corresponding cache-field adjustment.

Related: `pivot-table-multiple-value-fields`, `pivot-table-round-trip-preservation`,
`pivot-and-slicer-parts-survive-roundtrip`.
