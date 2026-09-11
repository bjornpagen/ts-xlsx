# ADR 0001: Rewrite runtime and toolchain, run `.ts` directly, defer the bundler

**Status:** Accepted (2026-07-11) · Phase 3 kickoff · _build slice resolved 2026-07-19 (see addendum)_

## Context

`STRATEGY.md` decided the rewrite stack up front: strict TypeScript, ESM-only,
Vitest, Biome, and a `tsup`/`unbuild`-class bundler emitting ESM plus `.d.ts`. Those
were sound defaults written before the rewrite began. On starting Phase 3 two facts
changed the cheapest correct path for the _first_ modules:

1. The dev environment runs **Node 24**, which executes `.ts` files directly via
   type-stripping, with no transpile step and no loader, and an `.mjs` can `import` a
   local `.ts` module.
2. The corpus is what pins the product's behavior and already runs via plain `node`. The
   rewrite only needs to be _reachable_ from a corpus adapter and _type-checked_; it does
   not need to be _bundled_ to be proven correct.

Standing up Vitest plus Biome plus a bundler now is real dependency and config weight for
no correctness gain on a single pure module, and it is the highest-drift work in the
plan (`PROGRESS.md` guardrail). Deferring it keeps the first slices lean and
dependency-clean (`CLAUDE.md` §2).

## Decision

- **Runtime/test path uses Node's native `.ts` execution, with no build step.** The
  `rewrite.mjs` corpus adapter imports `src/**/*.ts` directly; unit tests run under
  the built-in `node --test` runner on `.ts` files. Local TS imports use explicit
  `.ts` extensions.
- **`tsc` is the type-safety _gate_, not a build tool.** `npm run typecheck`
  (`tsc --noEmit -p tsconfig.json`) enforces the full strict flag set
  (`strict`, `noUncheckedIndexedAccess`, `exactOptionalPropertyTypes`,
  `noImplicitOverride`, `verbatimModuleSyntax`, …). TypeScript is pinned to 5.x.
- **The new tree is ESM, scoped by `src/package.json` (`"type": "module"`).** The
  legacy CommonJS root (`lib/`, `excel.js`) is left byte-for-byte intact, with no
  root `"type"` flip that would reinterpret legacy `.js` as ESM.
- **Vitest, Biome, and the ESM/`.d.ts` bundler are deferred** to a dedicated
  toolchain-standup slice, to run once enough of `src/` exists to justify the
  packaging/lint config and after the legacy Grunt/Babel/Mocha rip-out is scheduled.

## Consequences

- **Positive:** zero-build inner loop; no new runtime/test deps beyond `typescript`;
  legacy tree untouched so the freeze guardrail holds; strict typing still enforced.
- **Negative / deferred:** no published bundle or emitted `.d.ts` yet, which is fine
  because nothing is published in Phase 3; two test runners transiently (`node --test`
  for `src/`, Mocha for legacy `spec/`) until the swap; `node --test` assertions are
  written in a Vitest-portable style (`node:assert/strict`) to make that migration
  mechanical.
- **Revisit when:** a target runtime without `.ts` execution must run the code, a
  publishable artifact is needed (Phase 4), or `src/` is large enough that
  Biome's lint/format and Vitest's watch/coverage pay for their config.

## Addendum (2026-07-19): the deferred bundler resolved to _no bundler_

Phase 4's publishable-build slice revisited the deferred "`tsup`/`unbuild`-class
bundler" and rejected it. The emit requirement is narrow: rewrite the source's
mandatory `.ts` import specifiers (the no-build dev/test runtime needs them) to
`.js`, and emit `.d.ts`. TypeScript 5.7+ does exactly the first with
`rewriteRelativeImportExtensions`, and `tsc` already does the second. So the build
is **pure `tsc`** (`tsconfig.build.json` extends the strict gate config, flips
`noEmit` off, adds `declaration` plus `rewriteRelativeImportExtensions`, emits to
`dist/`). No bundler earns its place: no new dependency, no config to maintain, no
tree-shake/minify step we don't need for a Node-targeted ESM library.

**Amended 2026-08-08:** still pure `tsc`, now two invocations of it. `removeComments`
is whole-emit, so one pass cannot both strip the implementation prose from the JS and
keep the JSDoc in the `.d.ts`. The passes therefore split by audience, with `declaration`
moving to `tsconfig.build.dts.json`. This is the one thing a bundler would have been
asked for and it did not take one.

Decisions that rode along:

- **`exports`/`main`/`types` point at `dist/`;** `files` ships `dist` only (no maps, no
  `src`) to keep the tarball lean (~237 KB packed). Maps are omitted deliberately,
  because maintainers debug `src/` directly, never `dist/`.
- **`engines` split:** the _compiled artifact_ is ES2022 ESM and supports Node
  `>=18` (declared in `engines`); the _dev toolchain_ still needs Node 24 for
  `.ts` execution and is pinned via `.nvmrc`.
- **`private` dropped;** publish is guarded by `prepublishOnly` (build, full test,
  `smoke:dist`, `size`). The definitive package **name** remains the one human
  decision deferred to the rebrand slice.
- **Two new CI-enforced guards:** `smoke:dist` loads the _compiled_ artifact as a
  consumer would and asserts a write→read round-trip, catching emit-shaped breakage
  typecheck can't; `size` fails if the emitted runtime JS crosses a 600 KB budget
  (currently ~489 KB). Both run in a dedicated `Build` workflow.

## Addendum (2026-07-20): full corpus runs against the emitted artifact too

The stripping-against-transpilation question was revisited. No `.ts`-execution-free
runtime materialized, so type-stripping stays the dev/test inner loop, per §0's
framing. But it surfaced the one real gap: `test:src` and the corpus's default
adapter only ever exercise **stripped `src/`**, while consumers run the
**`tsc`-emitted `dist/`**. `smoke:dist` guarded a single round-trip against that
divergence; the full 671-behavior corpus did not.

Closed cheaply, with no new inner-loop cost: the `rewrite` corpus adapter now
retargets its implementation imports via `CORPUS_TARGET` (default `src`/`.ts`;
`dist`/`.js` mirrors the tree since `rootDir=src`). `pnpm run corpus:dist` runs the
entire behavioral corpus against the emitted JS, wired into the `Build` workflow
after `smoke:dist` so it reuses the single build. The zero-dep type-stripping loop
is unchanged for day-to-day work; emit parity is now proven by the whole corpus, not
one smoke case.

Vitest and Biome remain deferred to the toolchain-standup slice.

## Addendum (2026-07-21): the deferred package name is resolved

The one open human decision from the addendum above, the package name, is settled:
`@shbernal/ts-xlsx` is final. See [ADR-0015](0015-publishing-name-semver-and-first-version.md),
which also states the SemVer policy and confirms the first published version is `1.0.0`.
