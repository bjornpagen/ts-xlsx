# ADR 0011: Type-check the harness, migrating `test/` and `scripts/` to strict `.ts`

**Status:** Accepted (2026-07-20) · Phase 4 · extends the toolchain of [ADR 0029](./0029-toolchain-standup.md) and the gate philosophy of [ADR 0009](./0009-lint-type-gate-tightening.md)

## Context

The `src/` tree was held to the full strict bar (`strict` plus the strict-adjacent
`tsc` flags, Biome `recommended`), but the **harness was not type-checked at all**.
261 `.mjs` files, 8 tooling scripts and 253 under `test/` (249 regression-corpus
cases, the corpus runner, two adapters, the OpenXmlValidator runner), were
Biome-linted but never seen by `tsc`. A type error in the corpus surfaced only at
runtime under `pnpm run corpus`, and the highest-value seam in the whole project,
the `rewrite` adapter that binds the corpus to the `src` API, had zero compile-time
coupling to the types it exercised.

## Decision

### Migrate to `.ts`, do not bolt on `checkJs`

The first instinct was `checkJs` plus JSDoc as the "80/20". It was reversed after
measuring the delta: **both paths fix the same ~877 type errors at the same
strictness**, so that cost is shared. `checkJs`'s only _saving_ is avoiding a
mechanical rename, and in this repo that rename is tiny and measured (261 `git mv`;
cross-harness import seams were two patterns; the corpus runner had two loader
strings; nine `package.json` paths). Against that saving, `checkJs` would
institutionalize a permanent two-language halfway state (src = `.ts`,
harness = JS plus JSDoc) exactly where typing value is highest (adapters against `src`)
and JSDoc is weakest. CLAUDE.md ("TypeScript-first", "no half-migrations landed on
main", "legacy is not a reason to keep") points unambiguously to `.ts`. There is
no CI risk: Node runs `.ts` directly via type-stripping with no build step, the same
way `src/` already runs ([ADR 0001](./0001-rewrite-runtime-and-toolchain.md)).

Strictness **matches `src/`**: a `tsconfig.test.json` extends the strict base
config and adds `test/**/*.ts` and `scripts/**/*.ts` to its `include`. One new
script, `typecheck:test`, is wired into `pnpm test` (after `typecheck`) and into
the Corpus CI job.

### The corpus stays implementation-blind

The 249 cases assert observable behavior through an adapter (`api`) that
must never couple to one implementation. That contract is now _typed_: a shared
`test/corpus/case.ts` exports `Case`/`Behavior` plus `CorpusApi`, a named alias
for `any`, carrying a single `biome-ignore` and a comment explaining why. Every
case pins its default export with `satisfies Case`; implementation-blind values are
annotated `CorpusApi` (the named alias passes Biome's `noExplicitAny`; a literal
`any` would not). What the gate buys for the cases is real but modest: harness
typos, misused `assert` calls, and unguarded index access, not implementation
coupling. Behavior callbacks carry explicit param annotations because `assert`'s
assertion signatures reject contextually-typed call targets (TS2775).

### The adapter binds to `src`, which is the real prize

The `rewrite` adapter (today `test/corpus/adapters/rewrite/runtime.ts`) loads `src` modules through a generic
`loadModule<T>()` typed by `typeof import('../../../src/…')`, so the adapter is now
type-checked against the source API: a signature change in `src` becomes a compile
error in the corpus. The dist retarget (`CORPUS_TARGET=dist`) still works, because dist
mirrors src's public API, so the src type is accurate under the load-time cast. A
handful of documented `as` casts reach past a src object's _public_ API (two-cell
image anchors, data-table formula fields, pattern-fill `fgColor`); the corpus reads
those internals deliberately, and we cast at the read site rather than widen `src`
to expose them. **We never change `src` to satisfy the harness.**

## Consequences

- **Positive:** the corpus is now itself type-safe. The binding between the adapter
  and `src` turns a whole class of silent drift into a red gate. `typecheck:test` runs
  in `pnpm test` and CI at the same strict bar as `src/`.
- **Neutral:** `smoke-dist.ts` is excluded from the gate, because it imports the built
  `dist/` artifact, which does not exist before a build and never in the Corpus CI
  job; its type-safety is already covered by the typechecked `src` it is emitted
  from. `scripts/**/*.ts` and `test/**/*.ts` replace the old `*.mjs` globs in
  `biome.json` (the `noConsole`/`noNonNullAssertion` overrides ADR 0029 and ADR 0009
  describe now target `.ts`).
- **Revisit `CorpusApi`-as-`any` when:** a case genuinely benefits from a typed view
  of the adapter without coupling to an implementation. It does not today, because
  the blindness is the contract.
