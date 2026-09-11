# ADR 0036: oxlint and oxfmt replace Biome, and the type-aware rules come from tsgolint

**Status:** Accepted 2026-08-25 · supersedes the toolchain half of
[0029](./0029-toolchain-standup.md); the lint-gate posture of
[0009](./0009-lint-type-gate-tightening.md) stands and is now enforced by a different tool ·
**the exemption counts below are superseded** 2026-08-26 by
[0037](./0037-the-linter-and-the-typechecker-read-the-same-tsconfig.md) and
[0038](./0038-the-corpus-keeps-its-untyped-boundary.md), which found that ten of the sixteen
rules this record turns off were measuring a tsconfig mismatch rather than the code; the
decision to adopt oxlint stands unchanged

## Context

ADR 0029 chose Biome for lint and format, and it was the right call at the time: one fast
binary, one config, no plugin archaeology. What it could not give us is the rule set this
project actually needs.

The rules that matter for a library parsing untrusted spreadsheets are the ones that
require type information: is this promise floating, is this `any` flowing into a typed
parameter, is this `catch` about to receive something that is not an `Error`. Biome has
four of them, `noFloatingPromises`, `noMisusedPromises`, `useAwaitThenable` and
`useExhaustiveSwitchCases`, and all four are in `nursery`. All four are what `biome.json`
opted into by name. `nursery` is by definition a moving target, so the most important
gates in the file were the least stable ones in the tool.

oxlint reaches the real `typescript-eslint` type-aware surface by shelling out to
tsgolint, a Go build of the TypeScript checker. That is the whole argument. Everything
else below is measurement.

## Decision

oxfmt formats, oxlint lints, tsgolint supplies the types. Biome is removed.

`--type-aware` is passed by the `lint` package script and by `verify`'s whole-tree gate,
and by nothing else. `.oxlintrc.jsonc` deliberately does **not** set
`options.typeAware`: the config option and the CLI flag are an OR rather than an AND
(measured against a planted control: four findings under config-only, flag-only, and
both), and oxlint has no flag to switch it back off. Setting it in the config would make
every invocation type-aware, the pre-commit hook included. That is not mainly a cost
question; a type-aware finding depends on files outside the staged set, so a
staged-scoped run reports differently from CI for reasons that have nothing to do with
the commit, and a hook that disagrees with CI is a hook people learn to skip.

## What it costs

It is faster, which is worth stating plainly because an earlier draft of this migration
claimed the opposite. Best of five, whole tree, 511 files:

| command                                   | ms   |
| ----------------------------------------- | ---- |
| `biome check` (lint only)                 | 1398 |
| `oxlint`, the adopted rule set, 122 rules | 322  |
| `oxlint --type-aware`                     | 1297 |
| `biome format` (check only)               | 1287 |
| `oxfmt --check`                           | 502  |

The retracted claim is recorded because it is the trap: oxlint benchmarked with
`restriction` and `style` enabled is slower than Biome, and at 23 576 findings most of
that time is formatting output nobody asked for. Anyone re-measuring against the wide
presets will reach the same wrong conclusion and think this table is lying.

## Why the rule set is named by hand

Categories were measured cumulatively, each added to the one above it: `correctness` 31
findings, `+suspicious` 172, `+pedantic` 1990, `+restriction` 6716, `+style` 23 576. Past
`correctness` these stop being latent bugs and become another linter's house style. So
the config takes `correctness` and names everything else, and a rule this project turns
down is recorded in the config as `"off"` with its count and the reason. An unrecorded
decline is worse than no decision, because the next sweep re-derives it.

Nine rules are declined that way. The two worth reading are
`strict-boolean-expressions` and `no-unnecessary-condition`; the config carries the
reasoning at the site.

## The two counts that look like emergencies and are not

> Superseded 2026-08-26. Both counts were measured while the type-aware rules were reading
> TypeScript's _default_ compiler options for every file outside `src/**`. See ADR 0037. The
> `no-floating-promises` result below survives the correction unchanged; the second bullet does
> not, and ADR 0038 carries the re-measured version. The section is kept because the trap it
> describes is real and the numbers are what a cold reader will find in `git log`.

A bare `oxlint --type-aware` over this tree, before the overrides, reports numbers that
would panic anyone reading them cold. Both are configuration, not code:

- **`no-floating-promises`, 1160.** Every one is `node:test`'s `test()` or `describe()`
  returning a promise nobody awaits. `allowForKnownSafeCalls`, which exists for exactly
  this, names them and the count goes to zero with no source change.
- **The `no-unsafe-*` family, 1406**, plus `await-thenable` at 699 and
  `no-unnecessary-type-assertion` at 301. All in `test/corpus/**`, all traceable to
  `test/corpus/untyped.ts`, which declares `Untyped = any` on purpose because a corpus
  case describes behaviour against a deliberately untyped adapter surface. The rules are
  off across the harness and the config says why at length.

## What the migration found out about the source

This is the strongest single result and it belongs on the record. Across the entire
44-rule type-aware surface, including every rule this ADR declines, production `src/**`
reports **zero** for the whole `no-unsafe-*` family, `no-floating-promises`,
`await-thenable`, `no-misused-promises`, `only-throw-error`, `restrict-plus-operands`,
`no-deprecated`, `no-confusing-void-expression` and `return-await`, and zero non-null
assertions outside test files.

Nineteen findings needed fixing in `src/`, and fifteen of them were one wrong
declaration: `PackageAccessors` declared its two members with method syntax, which claims
they may read `this`, so every reader destructuring them off the object reported an
unbound method. They are closures over the package map. Property syntax says so, and is
the stricter declaration besides.

## How to prove the typechecker is actually running

A type-aware rule that is not switched on reports nothing, which reads exactly like a
clean tree. `docs/agent-correctness-playbook.md` carries a planted control that must
report exactly four findings. It reports the same four from `test/` and `scripts/`, even
though the root `tsconfig.json` includes only `src/**`: tsgolint is not bound to a
project the way `parserOptions.project` was, and `--tsconfig` overrides import resolution
only. What that observation missed is _which_ options those trees were then being checked
under. The answer is TypeScript's defaults, and
[0037](./0037-the-linter-and-the-typechecker-read-the-same-tsconfig.md) is the fix.

Use it before believing any zero in this document.

## Consequences

- `biome.json` is gone and `@biomejs/biome` is no longer a dependency. Eight
  `biome-ignore` directives were translated to `oxlint-disable-next-line`; every one was
  load-bearing and kept its reason verbatim.
- Four new suppressions exist for one upstream gap: oxlint does not count a JSDoc
  `{@link}` reference as a use, where tsc does, so four imports in `src/core/workbook.ts`
  that exist only to resolve doc links pass `noUnusedLocals` and fail the linter. Filed
  upstream already as `oxc-project/oxc#11639`, open since 2025-06-12, with a drafted fix in
  #13989 that was not merged, so there is nothing here to report. This is also not oxlint
  diverging from the linter it replaced: typescript-eslint behaves identically and closed its
  own reports as not-planned, pointing at `jsdoc/no-undefined-types`. `noUnusedLocals` is the
  only checker in this toolchain that resolves a doc link.
- `lint` passes `--report-unused-disable-directives`, so a suppression that outlives its
  cause fails the gate rather than quietly silencing a rule that would now pass.
- Formatting moved first and separately (`.oxfmtrc.jsonc`), configured to mirror the
  retired Biome formatter exactly. At those settings the two disagree on 3 of 510 files.
  Import sorting landed as its own commit, 47 files, because a reordering hidden inside a
  reflow is a reordering nobody reads.

## Revisit when

- oxlint gains a way to disable type-aware rules per invocation. Then
  `options.typeAware` can move into the config, editors get the rules, and the hook can
  still opt out.
- ~~The corpus adapter's `any` boundary is revisited.~~ Done 2026-08-26 in
  [0038](./0038-the-corpus-keeps-its-untyped-boundary.md): the boundary stands, and it costs
  six rules over `test/corpus/**` rather than sixteen across the whole harness. The other ten
  were [0037](./0037-the-linter-and-the-typechecker-read-the-same-tsconfig.md)'s tsconfig
  mismatch wearing the boundary's clothes.
