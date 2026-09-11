# ADR 0023: Seven subpath entry points, disjoint by construction, with the error taxonomy as its own face

**Status:** Accepted (2026-07-29) · packaging · extends ADR 0006 (docs generated from the barrel),
which assumed a single entry point.

## Context

`package.json` published exactly one specifier, `"."`, and `src/index.ts` re-exported 200 symbols
through it. A consumer writing a plain `.xlsx` therefore had the CFB reader, the MS-OVBA codec, the
ribbon parser, the BIFF12 decoder and CSV in the same module graph, and `"sideEffects": false` was
**not** declared, so a bundler had to assume every module was impure and keep all of it.
`scripts/size-budget.ts` measured one number for the whole `dist/`, which cannot see per-feature
cost and cannot notice a layer being crossed.

Measuring the static-import closure of each candidate entry against the emitted JS changed what the
step was worth doing _for_. The naive expectation, that splitting the barrel makes each codec
cheap, is mostly false here, and the measurements say why:

| entry       | closure | note                                                                   |
| ----------- | ------- | ---------------------------------------------------------------------- |
| everything  | 863 KB  |                                                                        |
| `/xlsx`     | 848 KB  | `readXlsx` sniffs and dispatches a binary package to the BIFF12 reader |
| `/xlsb`     | 435 KB  |                                                                        |
| `/csv`      | 308 KB  | `readCsv` builds a `Workbook`, so it pulls the model whole             |
| `/core`     | 299 KB  | of which ~150 KB is VBA plus ribbon: `Workbook` models a macro project |
| `/vba`      | 73 KB   |                                                                        |
| `/customui` | 26 KB   |                                                                        |
| `/errors`   | 12 KB   |                                                                        |

Two structural facts fell out of that table and are worth more than the split itself: the model
drags the whole VBA codec because `Workbook` parses and edits `vbaProject.bin` directly, and the
XML codec drags the whole BIFF12 codec because auto-detection is part of `readXlsx`'s contract
(architecture.md, "Two serialisations, one model") and the API is synchronous, so the dispatch
cannot be a dynamic import.

## Decision

1. **Seven entry barrels under `src/entries/`, one per published subpath** (`/core`, `/xlsx`,
   `/xlsb`, `/csv`, `/vba`, `/customui`, `/errors`), with `src/index.ts` unioning them so the bare
   package name still carries everything. The entries are _not_ the existing internal barrels:
   `src/vba/index.ts` and `src/customui/index.ts` carry the CFB writer, the MS-OVBA primitives and
   the part-path constants that the model and the codecs need, and those are implementation. A
   public face has to be able to be narrower than the module it fronts.

2. **Each symbol is listed in exactly one entry, and `src/index.ts` is `export *` over all seven.**
   The alternative, an explicit curated list at the root _and_ in each entry, means every new
   export needs two edits and a forgotten one is silent. The union costs nothing and cannot drift.

   The disjointness this requires is load-bearing, not tidiness: `export *` does not report an
   ambiguous re-export, it **drops the name**. A symbol exported from two entries would disappear
   from the root specifier with no diagnostic from `tsc`, from lint, or from any test that imports
   it by a subpath. `scripts/check-entries.ts` (a gate in `verify --full`) fails the build on a
   duplicate, on an entry `package.json` does not publish, on a published subpath whose module is
   gone, and on an entry the root barrel forgot to union.

3. **Every error class is exported from `/errors` and from nowhere else.** This follows from (2)
   rather than being an independent taste: a container-level failure belongs to no single codec,
   since `readXlsx` and `readXlsb` both raise `UnsupportedFormatError` and `XmlParseError` escapes any
   codec that parses a part, so shelving the classes with their codecs would have forced exactly
   the duplication the union cannot survive. Giving the taxonomy its own face turns that constraint
   into the best entry in the table: the classes reach nothing but each other, so a service that
   only classifies a failure (log it, map it to a status, decide whether to retry) pays 12 KB and
   not a parser. It also answers "what can this throw at me?" with one import.

4. **No `/streaming` subpath**, though the working plan proposed one. Measured, `read-rows` plus
   `write-stream` reach every module `read` plus `write` do, plus three. An entry point that costs
   what the codec costs is an alias, not a packaging boundary, so `readSheetRows`,
   `readWorkbookStream` and `WorkbookStreamWriter` are exported from `/xlsx`, where they belong.

5. **`"sideEffects": false`**, verified rather than asserted: no module under `src/` touches
   `globalThis`, a prototype, or `process` at import time, and every top-level statement in the
   emitted JS is a declaration. `scripts/check-layering.ts` was extended to see the bare
   `import '…'` form as well as `… from '…'`, so a side-effect-only import cannot enter unnoticed
   now that the manifest promises there are none.

6. **Budgets are per entry, and the total was raised from 600 KB to 950 KB** (**amended 2026-08-08**:
   every figure in this point was roughly halved, with the total at 530 KB, when `build` split into
   two `tsc` passes and the JS pass began stripping comments. Nothing left any closure; the numbers
   here had been ~47% comment prose, so they were measuring the wrong thing. Rebaselined onto
   comment-free emit they measure code, which is what makes the tripwire below able to do its job at
   all.) 600 KB had been the number since before the BIFF12 reader, the VBA codec and the ribbon
   parser landed; the build was measured at 861 KB and CI's `pnpm run size` step had been failing on
   it, unread. None of that growth was bloat, and a tripwire nobody can satisfy stops being read at
   all. The signal moves to the per-entry closures, which notice what a total cannot: a codec
   acquiring a _value_ import of something it previously needed only as a type moves one entry's
   number and leaves the total exactly where it was.

7. **`scripts/smoke-dist.ts` imports through the package name, not `../dist/`.** Node's
   self-reference resolves a package's own `exports` map, and that is the only thing in the repo
   that exercises it. The corpus's `dist` target loads emitted modules by file path, so a subpath
   resolving to nothing would pass every other gate and fail on a consumer's first install. The
   smoke test also asserts the two shape invariants the budgets state only as numbers: `/core`
   reaches nothing under `dist/io/`, and `/errors` reaches nothing but error modules.

8. **Every type reachable from a published signature is published, and every closed token union
   published here publishes its guard.** Disjointness (rule 3) is gated because `export *` drops an
   ambiguous name silently. The dual is just as silent and was not gated: a type named in a
   published signature that no entry exports. `TableColumn` was published and
   `TotalsRowFunction`, the type of its `totalsRowFunction`, was not, so a consumer could hold the
   value and had no way to write its type. Nothing reported it. The emitted `.d.ts` typechecks
   either way, because declarations import each other by relative path regardless of what
   `exports` publishes, and `docs:check` regenerates from the barrel, so it sees only what the
   barrel already lists. `scripts/check-public-types.ts` walks the type graph out of every entry
   and fails on what it reaches and cannot name; it found seventeen.

   A declaration may decline, by carrying `@unpublished` with its reason in its doc comment, and
   the run reports how many did. That is the same standing rule as a declined lint rule: a decline
   on the record is a decision, an absence is an accident. Four decline today, and all four are the
   streaming reader's granular output shapes, held as inferred structural types while that surface
   settles. The streaming _writer_'s plumbing (`StyleRegistry`, `FlushedSheet` and the rest) also
   declines, and its reason is a different one: those types are reachable only through
   `WorksheetStreamWriter`'s constructor and its `flushedSheet()`, which a consumer never calls,
   because a caller receives the writer from `WorkbookStreamWriter.sheet()`. The honest fix there
   is for those two members not to be on the public surface at all, which is a change to that
   class rather than to this rule.

   The guards are the same completeness question one level down. The unions are the spellings OOXML
   allows for an attribute; a caller holding a `string` has to get from it to the union somehow,
   and without the guard the only route is a cast, which is the thing the union exists to prevent.
   Two of the twenty were published and eighteen were not, with nothing stating a rule that admitted
   `isTableStyleElementType` and refused `isBorderStyle`. That was drift, so the rule is now stated
   in `entries/core.ts`'s own header and it is all of them. They are derived from the same table the
   union is, and they were already inside the entry's closure, so publishing them costs no bytes.

## Consequences

- **Additive; nothing breaks.** The root specifier exports the same 200 symbols it did before,
  proved by `docs:check`, which regenerates the API reference from the root barrel and showed a
  zero diff across the rewrite.
- **A consumer with a bundler should keep importing the bare name.** With `sideEffects: false` the
  bundler prunes better than a subpath can, because it works per symbol rather than per module. The
  subpaths are for the case where the graph itself should state the dependency: no bundler, or an
  architectural boundary worth making visible.
- **`/xlsx` being ~98% of the package is now a published fact rather than a surprise.** The README
  table says so, with the reason. Making it smaller means making `readXlsx`'s auto-detection
  lazy, which the synchronous API forbids. That is a real decision, not a tidy-up, and not this one.
- **`/core` carrying the VBA codec is the next honest target.** ~150 KB of the model's 299 KB is
  VBA and ribbon parsing, because `Workbook` owns `parseVbaProject`/`addVbaReference`/
  `removeVbaModule` directly. Moving that behind a boundary the model does not import at value level
  would halve the model's cost, and would be a breaking change to `Workbook`. It is deliberately
  out of scope here: this ADR is about how the package is published, not about what the model owns.
- **Revisit when:** an async read path exists (a dynamic import could then make the BIFF12 codec
  optional on the `/xlsx` path), or a consumer needs a subpath finer than a codec.
