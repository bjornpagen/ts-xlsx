# Column keys must address nested properties of row objects

Cluster: core-model / DX

## Scenario

A very common export shape is "an array of JSON records becomes a sheet", where each column is bound
to a record field by a `key`. Real records are rarely flat: a person carries a nested address,

```js
{ name: 'Adam', address: { line1: 'Street', postcode: 'NE1 3SA' } }
```

The author wants a column whose value is the postcode, that is, a key that reaches into the nested
object: `address.postcode`. Today the column `key` is treated as a single own-property name and
used to index the record directly, so a dotted key like `"address.postcode"` matches nothing and
the cell comes out empty. The only workarounds are to pre-flatten every record before handing it
to the library, or to abandon key-bound rows and address cells positionally, and both push
boilerplate back onto the caller for what is an everyday data shape.

## Desired behavior

- A column `key` may be a **property path** that addresses a nested value of the row record, not
  only a top-level own property. A dotted path (`address.postcode`) is the natural spelling; the
  design should decide whether to also accept an explicit path array (`['address', 'postcode']`)
  to sidestep genuine keys that contain a literal dot.
- Reading a row from a record resolves each column's path and writes the resolved scalar into the
  cell. A path that does not resolve, because of a missing intermediate, yields an empty cell, never
  a throw and never the literal string `"[object Object]"`.
- The mapping is symmetric where it can be: setting row values by a keyed object and reading them
  back by the same keys round-trips. Whether _writing_ through a nested path reconstructs the
  nested object on read-back is an open question (see below); the load-bearing requirement is the
  common export direction, record to cells.
- A plain, non-nested key keeps working exactly as before, and nested support is additive.

## Open questions

- Path syntax: dotted string only, or also an array form for keys containing dots? And are array
  bracket paths (`items[0].name`) in scope, or explicitly unsupported?
- Collision rule when a record has both a literal `"address.postcode"` own key and a nested
  `address.postcode`: which wins, and is that surprising?
- Is the reverse direction, where assigning to a nested-path column rebuilds the nested object graph,
  worth supporting, or is nested support read-only (record to cell) by design?

Related: `column-key-roundtrip-persistence`, `declarative-nested-column-headers`.
