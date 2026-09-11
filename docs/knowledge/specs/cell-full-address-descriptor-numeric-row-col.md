# Cell full-address descriptor: row and col are numbers

Cluster: types

Reported as "Bad type in Cell.fullAddress".

## Desired behavior

When a cell exposes a full-location descriptor (sheet name, address, row, column), the row and column fields MUST be 1-based integers matching the cell's position, and the public TypeScript type MUST declare them as `number`.

Example runtime shape for the top-left cell:

```
{ sheetName: "Sheet1", address: "A1", row: 1, col: 1 }
```

## Prior art

The legacy library shipped a declaration that typed these as the `Row` and `Column` object types:

```
fullAddress: {
  sheetName: string;
  address: Address;
  row: Row;    // wrong: runtime value is a number
  col: Column; // wrong: runtime value is a number
};
```

The runtime value has always been plain numbers, and the cell's own `row`/`col` accessors return numbers as well. The declaration was simply out of sync with reality.

## Guidance for the rewrite

- Whatever we name the per-cell location descriptor in ts-xlsx, its `row` and `col` must be typed as `number` (1-based), and the address as the A1-style string type.
- Cover this with a type-level test (`expectTypeOf`) so the descriptor's `row` and `col` can never silently regress to an object type again.
- Ensure runtime and type agree: a decoded `"A1"` yields `row: 1, col: 1`.

## Decoded address descriptor: the type must match the runtime exactly

### Desired behavior

The same drift bit the general **decoded-address** descriptor, the structured form produced when a
cell reference is decoded (row and column indices, column letter, sheet qualifier, absolute/relative
markers). Its published type must match the runtime value _field for field_:

- Every field present at runtime is declared; no field is declared that never appears at runtime.
- Each field's primitive is correct: row and column are 1-based **numbers**, the column letter and
  sheet name are strings.
- **Optionality reflects real optionality.** Conditional parts, the sheet qualifier and the
  `$`-anchored absolute/relative markers, are typed optional because they are genuinely absent for a
  bare `A1`; always-present parts (row index, column index, column letter) are typed required. The
  legacy hand-maintained `Address` interface was blanket-required or blanket-optional for convenience,
  producing either false compile errors or silent unsoundness (a field typed always-present that can
  be absent), defeating the type system exactly where it was supposed to help.

### Guidance for the rewrite

- Derive the descriptor's type from the decoder's implementation, or type-check it against it, rather
  than maintaining a separate ambient declaration, so this class of drift cannot recur.
- A type-level test asserts the decoder's return type _equals_ the published address type, so any
  future mismatch fails CI rather than surviving as declaration-only debt.

## Open questions

- Naming: whether ts-xlsx keeps a single "full address" descriptor or splits sheet-scoped from cell-scoped location info.
- Whether to also expose the object-typed row and column handles under a distinctly named field, so numeric coordinates and rich handles never share a field and confuse consumers.
- The exact canonical shape of the decoded address, meaning which fields are always present and which are conditional (sheet qualifier, absolute markers), and whether absolute/relative is one flag per coordinate or a combined marker.

Related: `public-type-surface-matches-runtime`.
