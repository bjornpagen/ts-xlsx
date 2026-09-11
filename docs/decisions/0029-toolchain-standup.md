# ADR 0029: Toolchain standup, Biome for lint/format, `node --test` kept, tsc for type tests

**Status:** Accepted (2026-07-19) · Phase 4 · resolves the toolchain deferrals from [ADR 0001](./0001-rewrite-runtime-and-toolchain.md) · **lint/format half superseded (2026-08-25)** by [ADR 0036](./0036-oxlint-and-oxfmt-replace-biome.md); the `node --test` and tsc-for-type-tests halves stand

> **Renumbered 2026-08-08 from 0002**, which it shared with
> [ADR 0002](./0002-ooxml-validation-oracle.md) (the validation oracle, accepted eight days
> earlier). The number moves to the next free one rather than to a chronological slot, because
> renumbering everything after it would break far more references than the collision did. The
> acceptance date is unchanged and is the authority on ordering.

## Context

ADR 0001 deferred three tooling decisions until `src/` was large enough to justify
their config: a linter/formatter (Biome), a test runner (the `STRATEGY.md` default
was Vitest), and type-level tests. The rewrite is now ~76 modules with 647 unit
tests and a 671-behavior corpus, all green. The publishable build shipped in the
prior slice ([ADR 0001 addendum](./0001-rewrite-runtime-and-toolchain.md)). It is
time to stand the toolchain up, and to decide each piece on its merits now that
there is real code to measure against, rather than inherit the pre-rewrite defaults.

## Decision

### Biome is the lint and format toolchain

One binary, zero transitive dependencies, covering TS and the `.mjs` corpus harness.
`biome.json` mirrors the style the tree was hand-authored in (2-space, single quote,
semicolons, `bracketSpacing: false`, trailing commas, 100-col), so adoption was a
near-no-op format pass rather than a restyle. The linter runs Biome's `recommended`
preset. Scripts: `lint` (`biome check`), `lint:fix`, `format`. `lint` is now the
first gate in `npm test`.

Two deliberate config choices:

- **`noNonNullAssertion` is disabled for test files only** (an `overrides` block
  matching `**/*.test.ts` and `test/**/*.mjs`). In a test a non-null assertion is
  the honest idiom, because the fixture is known, so `getCell('A1')!.value` beats
  defensively narrowing a value the test itself just created. Production code keeps
  the rule (CLAUDE.md §2 prefers narrowing over "trust me" escapes); the one `src`
  site was fixed by using `String.prototype.charAt` (typed `string`, no assertion).
- **The unsafe autofixes were applied, then verified against the corpus.** Biome
  classes `useOptionalChain` and `useTemplate` as unsafe, because they _can_ change
  semantics. We applied them and leaned on the full gate (typecheck, 647 unit
  tests, 671 corpus behaviors) to prove behavior was preserved. One unsafe fix was
  reverted by hand: `noSparseArray` rewrote the intentional sparse literals
  `['x', , 'z']` in two `addRow` hole-skipping tests to `['x', undefined, 'z']`, a
  real semantic change, since `1 in arr` is `false` for a hole and `true` for
  `undefined`. Those two sites now carry a `biome-ignore` with the reason.

> **Gotcha, load-bearing:** Biome 2.5.4 **silently drops the entire `overrides`
> array when `biome.json` contains `//` comments.** The config parses, the rest of
> it applies, but overrides vanish, so the test-file rule relaxation evaporates and
> ~50 warnings reappear. `biome.json` is therefore kept as **comment-free JSON**;
> rationale lives here instead. Do not add comments to it. CI would catch the
> regression as a wall of warnings, but the cause is deeply non-obvious.

### `node --test` is kept; Vitest is rejected

The `STRATEGY.md` default was Vitest. We reject it. The 647 tests already run green
under Node's built-in runner on `.ts` sources with zero build step, and everything
Vitest would buy us we already have without its dependency tree:

- **Coverage** comes from `node --test --experimental-test-coverage`. (Superseded 2026-08-25 by
  [0035](./0035-coverage-is-the-union-of-both-suites.md): that command measures only the unit
  suite, so it reported confident wrong numbers for everything the corpus covers. `test:coverage`
  is gone; `pnpm run coverage` reports the union. The point stands, in that the capability is
  Node's, not Vitest's.)
- **Type-level tests** come from tsc (see below), not `expectTypeOf`.
- **Watch** is already free, since the inner loop is build-free and `node --test --watch` exists.

Vitest is a large transitive dependency set (Vite, esbuild, rollup, chokidar, …), exactly
the kind of weight the fork exists to shed (CLAUDE.md §2, §4). It earns
its place only if we need a capability Node's runner lacks; today we don't. Tests
stay in `node:assert/strict` style, so if that day comes the migration is mechanical.

### Type-level tests are hand-rolled, checked by tsc

`src/type-tests/` holds `Expect<Equal<…>>` assertions over the public barrel, with no
`expectTypeOf` or `tsd` dependency. The standard bivariance-safe `Equal` and an
`Expect<T extends true>` are ~10 lines (`expect.ts`); the assertions
(`public-api.type-test.ts`) lock contracts the runtime tests cannot see: that
`writeXlsx`/`readXlsx` are synchronous (not `Promise`-returning), that a
`CellAddress`'s `col`/`row` are `number | undefined`, that `CellValue` never admits
`undefined`, and that every public symbol still exports. They are enforced by the
existing `typecheck` gate (`tsc` over all of `src/`) and excluded from the published
build (`tsconfig.build.json`). A drifted contract fails typecheck, the type-level
analogue of a red test, verified here by a negative control.

## Consequences

- **Positive:** lint and format is one zero-dep binary; the whole authored tree
  (src, scripts, corpus) is consistently formatted and lint-clean; no runtime-runner
  dependency added; type contracts are now regression-guarded; `npm test` gained a
  `lint` gate up front and coverage is available on demand.
- **Negative / deferred:** Biome's `recommended` is the whole ruleset, with no
  project-specific rules curated yet; coverage is available but **not** yet a CI
  gate with a threshold (deliberate, since a threshold is a follow-up once we decide the
  floor; **the floor was decided 2026-08-25 in**
  [**0035**](./0035-coverage-is-the-union-of-both-suites.md), which also found that the number
  being deferred here was never the library's); the Biome comment gotcha is a sharp edge
  documented but not fixable by us.
- **Revisit when:** Node's test runner blocks a capability we need (reconsider
  Vitest, since the migration is pre-planned); ~~coverage should become an enforced floor~~
  (done, see [0035](./0035-coverage-is-the-union-of-both-suites.md));
  or the ruleset needs curation beyond `recommended`.
