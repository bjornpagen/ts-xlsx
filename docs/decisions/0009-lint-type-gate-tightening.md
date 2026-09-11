# ADR 0009: Tighten the lint/type gates where free; decline `isolatedDeclarations`

**Status:** Accepted (2026-07-20) · Phase 4 · extends the toolchain of [ADR 0029](./0029-toolchain-standup.md) · the posture below stands; the tool enforcing it changed (2026-08-25) with [ADR 0036](./0036-oxlint-and-oxfmt-replace-biome.md)

## Context

An audit of the lint/type configuration asked a narrow question: are there checks we
could enable that would add real safety without paying for a backlog of pre-existing
violations? The gates were already green at the strict bar CLAUDE.md §2 mandates
(`strict` plus the seven strict-adjacent `tsc` flags, Biome `recommended`). The audit
measured each candidate by actually running it against the tree rather than reasoning
about it, and sorted candidates by _current violation count_. A rule that fires zero
times today locks in a guarantee at zero code cost; one that fires hundreds of times
is a refactor backlog wearing a config toggle.

## Decision

### Adopt three checks that pass clean today (commit `baf90294`)

Each was measured at **zero current violations** across the relevant scope, so
enabling it changes no existing code. It only prevents a future regression:

- **`tsc` `noUncheckedSideEffectImports: true`** catches a broken or typo'd
  side-effect import (`import './x'`) instead of letting it resolve to nothing.
- **Biome `nursery/noFloatingPromises: error`** enforces that every promise is
  awaited. This is the single highest-value correctness rule for _this_ library:
  the reader and writer paths are async-heavy, and a dropped `await` is exactly the
  class of bug that survives unit tests but corrupts streaming output. Caveat noted:
  it is a `nursery` (still-stabilizing) rule; if a false positive ever appears, the
  fallback is a scoped `biome-ignore`, not disabling it globally.
- **Biome `suspicious/noConsole: error`** in `src` keeps stray debug logging out of
  the shipped library. Overridden **off** for `scripts/**/*.mjs` and `test/**/*.mjs`,
  where `console` is the legitimate idiom (the same `overrides` mechanism already
  used for `noNonNullAssertion` in ADR 0029).

### Reject `isolatedDeclarations`

Considered and **declined.** Enabling it (on the build config, which already emits
`.d.ts`) produces **26 findings, none of which touch the public API barrel**
(`src/index.ts`; its 19 exports are clean). The findings cluster in four internal
`io/xlsx/` modules, and 20 of the 26 fall on one hotspot: the `REL` table in
`relationships.ts`, which derives OPC relationship-type URIs from the `NS` table via
template literals (`` `${NS.docRels}/worksheet` ``). `isolatedDeclarations` cannot
infer those, and every way to satisfy it there is a _downgrade_:

1. annotating `REL: Record<…, string>` loses the exact-URI literal types the
   `as const` currently guarantees;
2. giving an explicit literal type per key duplicates all 18 URI strings verbatim;
3. inlining the full URIs instead of deriving from `NS` defeats the deliberate DRY
   derivation those tables exist for.

The value `isolatedDeclarations` offers, a guaranteed-correct, fast-to-emit
_published_ `.d.ts`, is already met here by other means: the public API produces
zero findings and is independently pinned by the `Expect<Equal>` type-tests in
`src/type-tests/` and the docs-from-barrel generator ([ADR 0006](./0006-docs-from-types.md)).
The performance argument (parallel declaration emit) is nil for a 97-file library that
builds instantly, and today's inference-based emit already works. So the cost lands on
internal plumbing where the benefit doesn't, and contorts the DRY constant tables to
satisfy a check whose payoff is already covered.

This is the inverse of the three checks we _did_ adopt: those lock in a guarantee at
zero code cost. `isolatedDeclarations` would make the code worse to satisfy a
redundant check, the opposite of CLAUDE.md §5's "harder, better thing", since here
the better thing is _not_ contorting `relationships.ts`.

## Consequences

- **Positive:** three regression classes (broken side-effect imports, floating
  promises, shipped `console` calls) are now impossible to reintroduce without a
  failing gate, with no existing code touched.
- **Neutral:** `biome.json` remains comment-free JSON (ADR 0029's load-bearing
  gotcha, where comments silently drop the `overrides` array); rationale lives here.
- **Revisit `isolatedDeclarations` when:** the project starts shipping hand-authored
  `.d.ts`, or the public API grows large enough that inference-based declaration emit
  becomes a correctness risk. Neither holds today. Do not re-litigate absent one of
  those triggers.
