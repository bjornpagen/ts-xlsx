# The worksheet column collection must be honest about mutability

Cluster: core-model

Reported as "Issue when attempting to push elements into the columns".

## Desired behavior

The worksheet column collection should present a single, coherent, discoverable API for adding, reading, and mutating columns. Today's shape, where the public `columns` accessor is a getter and setter over an internal representation, surprises users because the value it returns _reads_ like an array in most respects but is not a live, mutable array. Idiomatic in-place mutation, appending, inserting or splicing a column definition, either throws or silently corrupts internal state, leaving later `getColumn(...)` lookups to fail with an out-of-bounds "invalid column letter" error even though the column names appear to be stored.

The modern library should make the correct usage the easy path:

- Provide an explicit, well-typed way to **append a column** from a definition (header, key, width, style, and so on) that keeps all internal indices and lookup maps consistent, so a subsequent lookup by key or by index succeeds.
- Provide explicit **insert and remove** operations rather than expecting users to reach for raw `Array.prototype` mutators on a value that only looks like an array.
- If a read accessor returns a collection, it must be honest about mutability: either return a genuinely live, fully array-compatible collection whose mutations write through to the sheet, or return an immutable snapshot and force mutation through named methods. The current in-between, which looks mutable and isn't, is the trap to eliminate.
- Bulk assignment of an entire column set, the current reliable workaround, must remain supported, but must not be the _only_ reliable path.

## Prior art

ExcelJS exposes `worksheet.columns` as a getter and setter, `getColumn(keyOrLetterOrNumber)` for individual access, and `column.values` for the cell values down a column. The failure modes reported: `push` on the returned value throwing, and dynamically-added columns not being resolvable via `getColumn`, throwing "Out of bounds. Invalid column letter". Multiple users hit this over several years and only the "assign the whole array" workaround was reliable.

## Open questions

- Should the append operation take a single definition, or accept variadic or iterable definitions for batch append?
- What is the contract when appending a column whose `key` collides with an existing column's key: error, overwrite, or ignore?
- Should a returned column collection be immutable-by-default, aligning with the fork's immutable-by-default stance, with mutation exclusively via named worksheet methods, and is a live array-proxy ever worth the complexity and footguns?
- How should appending a column interact with rows that already have data? Does the new column start empty, and are existing row `.values` arrays extended lazily?

## Corpus follow-up

Once the concrete API is chosen, add regression cases asserting that appending a column then calling `getColumn(key)` resolves it without throwing; that appending a column then reading its letter or number is consistent with its position; and that the chosen mutation surface either mutates through or is frozen, never the silent-corruption middle state.
