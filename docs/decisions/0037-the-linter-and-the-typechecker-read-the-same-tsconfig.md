# ADR 0037: Every tree the linter type-checks has a `tsconfig.json` the linter can find

**Status:** Accepted 2026-08-26 · corrects a mechanism [0036](./0036-oxlint-and-oxfmt-replace-biome.md)
observed but did not diagnose; the counts in that record's "two counts that look like
emergencies" section were measured before this and are superseded by
[0038](./0038-the-corpus-keeps-its-untyped-boundary.md)

## Context

ADR 0036 noted, as a curiosity, that "tsgolint is not bound to a project the way
`parserOptions.project` was, and `--tsconfig` overrides import resolution only." That is
true and it is not a curiosity. It means the type-aware rules were reading compiler
options from somewhere, and nothing in the repo said where.

Where, it turns out, is **the nearest `tsconfig.json` that includes the file being
linted**. Every gate config this repo owns for the harness has a non-standard name,
`tsconfig.test.json`, which that search does not look for, and the one file actually
called `tsconfig.json` includes `src/**` and nothing else. So for `test/`, `scripts/` and
`tools/` the search found nothing that claimed the file and fell back to TypeScript's
**defaults**: no `strict`, no `noUncheckedIndexedAccess`, no `exactOptionalPropertyTypes`.

The linter was judging two thirds of the tree under options the typechecker does not use.

## The control

Two files, identical text, one under `src/` and one under `scripts/`:

```ts
const a: string[] = [];
export const x = a[0] as string;
```

Under `noUncheckedIndexedAccess`, `a[0]` is `string | undefined` and the assertion is
load-bearing. Run against `no-unnecessary-type-assertion`, the `src/` copy reported
nothing and the `scripts/` copy reported the assertion as unnecessary. Adding a
`tsconfig.json` beside the second file silenced it. `--tsconfig tsconfig.test.json` did
not, and neither did `--tsconfig tsconfig.json`. Both were measured.

That is the whole diagnosis. Everything below is what it cost.

## Decision

`test/`, `scripts/` and `tools/` each carry a `tsconfig.json` that extends
`tsconfig.test.json`, narrows `include` to its own directory, and turns incrementality
off with a nulled `tsBuildInfoFile` so it can never race the gate's cache. They exist for
tsgolint and for editors. Nothing runs `tsc` against them; `pnpm typecheck` still runs
`tsconfig.test.json` over all three directories at once, so there is exactly one place
the harness's options are written down.

## What it changed

Re-measuring the type-aware surface with the linter and the typechecker finally agreeing:

|                                                                                  | before the configs | after them | after the fixes |
| -------------------------------------------------------------------------------- | ------------------ | ---------- | --------------- |
| `no-unnecessary-type-assertion`, `test/corpus/**`                                | 301                | 22         | 0               |
| findings in `scripts/`, `tools/`, `src/**/*.test.ts`, `test/` outside the corpus | 36                 | 31         | 0               |
| type-aware rules switched off across the harness                                 | 16                 | 16         | 6               |

The 279 assertions that stopped being reported were never unnecessary; the linter could not
see the flag that made them necessary. What survived the correction was real, and fixing it
(31 findings outside the corpus, 35 inside it) is what let ten rules come back on
(ADR [0038](./0038-the-corpus-keeps-its-untyped-boundary.md)).

Two disagreements survived the fix and are worth naming, because both look like the rule
being wrong and only one is:

- **A constructor cast that assignability calls redundant.** `errors.test.ts` cast a
  `new (message?: string) => XlsxError` to the two-argument form before calling it with
  two arguments. The types are mutually assignable, so the rule is right that the cast
  changes nothing, and deleting it produces `TS2554: Expected 0-1 arguments, but got 2`.
  Fixed by widening the taxonomy table's declared constructor type, which is what the
  test was asserting anyway.
- **`let buffer = null`.** TypeScript's evolving-`let` inference gives that variable
  `Uint8Array | null` at the use site; tsgolint resolved it to `Uint8Array` and called the
  `!` unnecessary. Its autofix removed six of them and broke the typecheck. Annotating the
  declaration (`let buffer: Uint8Array | null = null`) makes both agree. **Do not run
  `--fix` for this rule without a `tsc` run behind it.**

## Consequences

- `charcheck.config.ts` is the one file left uncovered: it sits at the repo root, where the
  nearest `tsconfig.json` is the `src/**` gate, and a root config that included it would
  muddy exactly the gate that has to stay narrow. It reports zero findings, and the only
  risk it carries is a false positive rather than a missed one.
- `scripts/smoke-dist.ts` stays outside a project deliberately. It imports the package by
  name through the published `exports` map, so its types live in `dist/`, which does not
  exist before a build and never in the corpus CI job. `.oxlintrc.jsonc` turns the six
  `any`-sensitive rules off for that one file and says why; `tsconfig.dist.json` checks it
  properly after a build.
- Editors now resolve the strict options for `test/`, `scripts/` and `tools/` instead of
  inferring a loose project, which is the same fix arriving somewhere else.

## Revisit when

oxlint grows a way to point the type-aware rules at a project explicitly. Then these three
files can go and the flag can name `tsconfig.test.json` directly. Check `--tsconfig`
against the control above before believing it does that. Today it does not.
