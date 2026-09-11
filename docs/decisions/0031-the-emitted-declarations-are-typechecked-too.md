# ADR 0031: The emitted declarations are typechecked too; two more gate flags go on

**Status:** Accepted (2026-08-08) · extends the gate audit of [ADR 0009](./0009-lint-type-gate-tightening.md) · supplies the missing control for the `isolatedDeclarations` rejection recorded there

## Context

[ADR 0009](./0009-lint-type-gate-tightening.md) audited the lint/type configuration by the
right method: run every candidate against the tree, sort by _current violation count_, adopt
what fires zero times and decline what is a refactor backlog wearing a config toggle. That
audit predates [ADR 0028](./0028-typescript-7-adoption.md), which replaced the compiler, and it
asked only "which checks can we turn on?", never "what does the gate not look at?"

Re-running it against TypeScript 7 answered both. Two more flags now fire zero times, and one
whole artifact turned out to be unread.

## Decision

### Adopt two flags that pass clean today

- **`erasableSyntaxOnly: true`**, with zero violations (no enum, namespace, or parameter property
  anywhere in `src/`, `test/`, `scripts/`, `tools/`). This one is not hygiene; it is the flag
  that makes the no-build workflow's central assumption machine-checked. Every gate in this
  repo runs `.ts` through Node's type-stripping, which _refuses_ non-erasable syntax. Without
  the flag `tsc` accepts all three constructs happily and the failure arrives later, from
  `node --test`, as a parse error pointing at the syntax rather than at the decision. Verified
  by control: an `enum` added to `src/` now fails the gate with TS1294.

- **`skipLibCheck: false`**, with zero errors in both projects. The near-universal `true` is a
  performance workaround from an era when checking dependency declarations was the expensive
  part of a run, and ADR 0028 ended that era: measured cold, **0.451 s with it on, 0.447 s with
  it off**. Free. With one runtime dependency and Renovate opening bumps for CI to review
  ([ADR 0027](./0027-dependencies-are-updated-by-a-bot-and-ci-is-the-reviewer.md)), a dependency
  shipping broken declarations _should_ fail this gate rather than be skipped past.

### Typecheck the emitted `.d.ts` with `tsconfig.dist.json`, run in `build.yml`

The gap worth the ADR. `tsconfig.json` proves `src/` is sound; the corpus, run against `dist`
via `corpus:dist`, proves the emitted **JS** behaves. Nothing read a single line of the emitted
**declarations**, and those are the half of the artifact a consumer meets first, in an editor.

That hole sat directly beneath ADR 0009's rejection of `isolatedDeclarations`, which rested in
part on the claim that inference-based declaration emit "already works." The claim was probably
true. It was also unverified, and the rejection is the reason nothing else was going to verify it.

`scripts/smoke-dist.ts` turned out to be the whole fixture, already written. It is the only file
in the tree importing through the package _name_, so typechecking it resolves the published
`exports` map and pulls in the declarations behind every subpath. It was excluded from
`tsconfig.test.json` for the one reason that does not apply after a build: there, `dist/` does
not exist yet. So the check is a third project over that one file, run in `build.yml` where
`dist/` is fresh, rather than a gate in `scripts/verify.ts`, which must stay runnable on a tree
that has never been built.

`skipLibCheck: false` is load-bearing rather than incidental here: `dist/*.d.ts` _is_ the
declaration file under test, so skipping lib checks would skip the entire point of the project.

Cost is **0.78 s**, once per CI build, reusing the build that was already there.

### Verified by control, not by a passing run

A gate that passes proves nothing until something known-broken makes it fail, and the first
control written for this one was itself vacuous, patching a symbol in `dist/index.d.ts` that
is a pure re-export barrel and contains no declarations to corrupt. The control that counts
changed `writeXlsx`'s parameter type at its real declaration site, three hops down the
re-export graph (`index.d.ts` → `entries/xlsx.d.ts` → `io/xlsx/write.d.ts`):

```
scripts/smoke-dist.ts(30,31): error TS2345:
  Argument of type 'Workbook' is not assignable to parameter of type 'number'.
```

So the gate reads declaration _bodies_ across the full graph, not just resolves modules.
Deleting `dist/index.d.ts` outright fails it too, with TS7016, a weaker control that only
proves resolution, recorded here so the next agent does not mistake it for sufficient.

### Decline `noPropertyAccessFromIndexSignature`

**327 violations.** Declined by ADR 0009's own sorting rule. It would also fight
`src/customui/ribbon.ts` specifically, where an index signature is the honest model of
arbitrary ribbon attributes and bracket access at every use site would obscure that.

`allowUnreachableCode: false`, `allowUnusedLabels: false`, and `moduleDetection: "force"` also
measure zero, and are left off: Biome and `verbatimModuleSyntax` already cover that ground, and
a flag that duplicates an existing guarantee is config to maintain for no new invariant.

## Consequences

- **Positive:** non-erasable syntax can no longer reach `master` and break the no-build
  workflow; a dependency's broken declarations now fail a gate instead of being skipped; the
  published `.d.ts` is checked as a consumer sees it, through the `exports` map, every build.
- **Positive:** ADR 0009's `isolatedDeclarations` rejection keeps its "do not re-litigate"
  standing and now has the safety net it was implicitly assuming, on the consumer-facing
  declarations only, without contorting `relationships.ts`.
- **Neutral:** a third tsconfig. It earns its place by having a subject the other two cannot
  have: an artifact that does not exist until `build` runs.
- **Negative / accepted:** `typecheck:dist` is reachable locally only after `pnpm run build`,
  so it is absent from `verify --full` and runs on CI alone. That is deliberate, since `verify`
  must work on a never-built tree, but it does mean a declaration regression surfaces on the
  runner rather than at pre-push.
- **Revisit when:** the `unstable/*` API or a future compiler makes `skipLibCheck: false`
  measurably expensive again (re-measure before reverting; do not assume the old cost), or the
  public API grows enough that one smoke file stops reaching most of the declarations,
  at which point `isolatedDeclarations` deserves its re-litigation trigger.
