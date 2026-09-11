# ADR 0038: The corpus keeps its untyped boundary, and it costs six rules rather than sixteen

**Status:** Accepted 2026-08-26 · answers the question
[0036](./0036-oxlint-and-oxfmt-replace-biome.md) deferred; depends on
[0037](./0037-the-linter-and-the-typechecker-read-the-same-tsconfig.md) for its numbers

## Context

ADR 0036 left an open question: sixteen type-aware rules were off across `test/**`,
`scripts/**` and `tools/**`, and the reason given for all of them was
`test/corpus/untyped.ts`, the module that declares `type Untyped = any` on purpose. If a
declared `any` costs sixteen rules over the whole harness, the declaration is worth
re-examining. So the question was put as "should the corpus adapter still have an `any`
boundary?" and parked.

It was the wrong question, because the premise was wrong. Ten of those sixteen rules were
not reporting the boundary at all. They were reporting the tsconfig mismatch ADR 0037
describes, and they went to zero the moment the linter could see the same compiler options
`tsc` uses. The exemption list was measured honestly and measured the wrong thing.

## What the boundary actually is

`CorpusApi` is not `any` and has not been for some time. It is derived from the adapter
object, so a case sees the name, the arguments and the return type of every capability it
calls. `Untyped` is what is left: 288 occurrences, and its module comment sorts them into
the two kinds that remain.

- **Case specs.** The declarative `{sheets: [{cells: [...]}]}` a case hands `buildFrom`,
  and the options bags built around it. Writing that type down is a real piece of design
  work, not an annotation: the spec language is the corpus's own DSL.
- **Report accumulators.** `Record<string, Untyped>`, the shape a capability builds its
  findings into. Widening it to `unknown` types nothing; it moves the debt into the 254
  cases that read the report.

Everything else has been paid off, and `untyped.ts` says so at the site, with the `grep`
that counts the balance.

## Decision

The boundary stands. The exemption it earns is narrowed to what it actually reaches:
`test/corpus/**` only, and six rules rather than sixteen.

| rule                      | findings | why it is off                                                                        |
| ------------------------- | -------- | ------------------------------------------------------------------------------------ |
| `no-unsafe-member-access` | 803      | a case reads its report through `Record<string, Untyped>`                            |
| `await-thenable`          | 699      | cases `await` every capability so one can turn async without touching 282 call sites |
| `no-unsafe-assignment`    | 373      | destructuring that report                                                            |
| `no-unsafe-argument`      | 122      | handing a spec back into a capability                                                |
| `no-unsafe-call`          | 66       | calling through the untyped surface on purpose                                       |
| `no-unsafe-return`        | 48       | an adapter capability returning its accumulator                                      |

2111 findings, and the sample checked by hand traces where the table says: a case's
`(await api.roundtripWorkbook(SPEC)).sheets.S.cells.A1` is three of them at once. None is
a defect; each is the design being reported.

Ten rules came back on: `no-unnecessary-type-assertion`, `no-base-to-string`,
`no-misused-spread`, `require-array-sort-compare`, `restrict-template-expressions`,
`only-throw-error`, `prefer-promise-reject-errors`, `restrict-plus-operands`,
`unbound-method` and `no-unsafe-enum-comparison`. All ten now report zero everywhere, the
corpus included. Sixty-six findings were fixed to get there, 35 inside `test/corpus/**` and 31
in `src/**/*.test.ts`, `scripts/` and `tools/`. Two were worth the trip on their own:
`[...s]` iterating code points where a CFB stream name is UTF-16 code units, and `String(v)`
on a `CellValue` rendering every non-primitive shape as `[object Object]` in a probe's own
failure output.

## Why not type it

An adapter returning fully typed values would let those six rules run over the harness,
which is where a regression in the corpus itself would show up. That is a real gain and it
is not free: it is a rewrite of `test/corpus/adapters/ts-xlsx/**` and of the report shapes
all 282 cases are written against, and the case-spec DSL would have to be designed rather
than inferred.

At sixteen rules across three trees that trade looked worth arguing about. At six rules
over `test/corpus/**`, a directory whose entire job is to describe behaviour from outside the
library, it does not. The corpus's value is that a case cannot see how the library is
built; the boundary is a weaker version of the same idea, and the rules it silences are
the ones that exist to complain about exactly that.

## Revisit when

- The case-spec DSL is written down as a type for some other reason. Then the largest of
  the two remaining classes is already paid for and the balance changes.
- The count in `untyped.ts` starts climbing. It is 288 today, and the module comment
  carries the command that measures it. A boundary that is being _spent_ rather than held
  is a different decision from this one.
