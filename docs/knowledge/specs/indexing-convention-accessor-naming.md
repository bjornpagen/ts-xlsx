# Positional accessors must make their 1-based vs 0-based convention unambiguous

Cluster: types

## Scenario

Row, column and worksheet accessors use 1-based positions: `getWorksheet(1)` returns the first sheet
and `getWorksheet(0)` returns nothing, because the spreadsheet grid is naturally 1-based (row 1,
column A is 1). Meanwhile the array-shaped views (`workbook.worksheets[0]`, cell iteration, and JS
itself) are 0-based. The accessor parameters are named and typed `index`, `indexOrName`, or
`indexOrKey`, which reads as a 0-based array index and contradicts the actual 1-based behavior. Some
methods document "1-indexed" in JSDoc and others do not, and at least one accessor conflates an ID
with a position, so a developer relying on IntelliSense cannot tell from the signature whether to pass
a 0-based index, a 1-based position, or an opaque ID.

> Spec note, not a corpus case: this is a developer-experience and type-surface requirement, not a
> runtime behavior bug. The accessors work; their _names_ lie. In a TS-first fork where the types are
> the primary documentation, a parameter named `index` that is actually 1-based is a real defect. It
> becomes enforceable through the public-type surface and type-level tests, not a runtime assertion.

## Desired behavior

- **Never name a 1-based parameter `index`.** Reserve `index` for genuine 0-based array positions. Use
  a name that carries the convention for 1-based positional accessors, such as `rowNumber`,
  `colNumber` or `position` for 1-based against `index` for 0-based, and `id` when the value is truly
  an identifier rather than a position. The name alone must disambiguate.

- **One consistent convention across every positional accessor.** The numbering convention is explicit
  and uniform across worksheet-by-position, `getRow`, `getColumn`, and their by-key and by-name
  overloads, not left to JSDoc on some methods and absent on others. A developer should never have to
  read prose to learn whether a method is 0- or 1-based.

- **Do not conflate ID and position in one parameter.** Where an accessor today overloads the same
  parameter to mean either an identifier or a position, split them into distinctly-named and typed
  entry points, or a clearly discriminated union or overload, so the type tells the caller which they
  are supplying.

- **Types are the contract.** The published signatures encode the convention (branded types or precise
  overloads for by-position, by-key and by-id), and a type-level test pins that the by-position
  accessor rejects being called as if 0-based-array-shaped, so the API cannot silently
  regress to the ambiguous naming.

## Open questions

- Naming scheme: `rowNumber`/`colNumber`/`position` for 1-based against `index` for 0-based, or a
  single consistent word (`position`) everywhere a 1-based value is accepted? Pick one and apply it
  uniformly.
- Should the fork consider making the array-shaped and accessor-shaped APIs agree on a single base
  (all 1-based, or expose both explicitly) rather than carrying the 0/1 split at all? That is a larger
  breaking decision the fork is free to take.
- Whether a branded numeric type (`RowNumber`, `ColNumber`) is worth the ergonomic cost, or precise
  naming plus JSDoc suffices for the type-as-docs goal.

Related: `public-type-surface-matches-runtime`, `nested-property-path-column-keys`,
`column-key-roundtrip-persistence`, `worksheet-columns-mutable-array-ergonomics`,
`row-values-no-phantom-leading-slot`, `native-iteration-protocol`.
