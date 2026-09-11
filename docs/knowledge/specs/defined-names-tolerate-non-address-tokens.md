# Defined names must tolerate tokens that are not cell references

## The scenario

Real `.xlsx` files created in Excel routinely carry **defined names** whose value is
not an A1-style cell reference: named ranges backing a data-validation dropdown, or
names that collide with function-like identifiers (`MONTH`, `REPT`, …). When such a
file is opened, the reader parses the workbook's defined-name table and, for each
entry, tries to decode the referenced range into concrete cells.

## The failure this must not reproduce

Decoding a defined name by feeding its raw token straight into address decoding is
unsafe: a token like `MONTH` is handed to the column-letter resolver, which has no
column past `XFD` and throws `Out of bounds. Invalid column letter: MONTH`. Because
this happens deep inside the workbook-model setter during load, the error surfaces as
an **unhandled rejection that aborts the entire file read**, so a file that opens fine
in Excel becomes completely unreadable through the library. Users hit this on
otherwise ordinary spreadsheets and are forced to abandon the library or pre-edit the
file.

## Desired behaviour

- Reading a workbook whose defined-name table contains a non-address token must
  **not throw**, and must not abort the load of the rest of the workbook.
- A defined name that does resolve to cells is decoded as today.
- A defined name that does **not** resolve to a cell reference is preserved
  opaquely, kept as its raw string form on the model, rather than discarded, so a
  subsequent write does not silently drop names the file legitimately declared.
- The distinction is made by _attempting_ the decode and treating a decode failure
  as "this name is not a cell reference", not by trying to pre-classify tokens.

## Root cause (legacy)

The legacy read path (`doc/defined-names.js`) maps every defined-name range through
`col-cache.decodeEx`, then `decodeAddress`, then `l2n`, and `l2n` throws on any token that is
not a valid column letter. There is no guard distinguishing "a range I should decode"
from "a name I should keep verbatim".

## Prior art / signals

- Multiple independent reports of the same crash with different tokens (`MONTH`,
  `REPT`), all tracing to the same `l2n` throw during defined-name model load.
- A common user workaround, assigning `worksheet.columns` first, only masks the
  symptom by changing which lookup path runs; it does not address the unsafe decode.

## Open questions for the rebuild

- Should non-reference defined names be surfaced in the public API, as a `definedNames`
  view that can hold non-range values, or normalised away? Preserving them is safer
  for round-trip fidelity.
- Data-validation lists that point at a named range need the name to survive read then
  write so the dropdown still works in the output file; this should have its own
  corpus case once a data-validation capability exists in the adapter.
