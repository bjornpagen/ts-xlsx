# ADR 0035: Coverage is the union of both suites, measured with node's own implementation

**Status:** Accepted 2026-08-25 · picks up the coverage thread deferred by
[0029](./0029-toolchain-standup.md)

## Context

ADR 0029 kept `node --test` over Vitest, and one of the capabilities it pointed at to
justify that was coverage: _"Coverage → `node --test --experimental-test-coverage`
(script `test:coverage`)."_ It then deferred the floor (_"coverage is available but
**not** yet a CI gate with a threshold (deliberate, since a threshold is a follow-up once we
decide the floor)"_) and listed _"coverage should become an enforced floor"_ under
**Revisit when**.

What went unnoticed is that the number being deferred was never the library's number.

This project tests through two deliberately separate suites (ADR 0012, and the topology
in `docs/architecture.md`): the white-box unit tests co-located as `src/**/*.test.ts`,
and the implementation-blind black-box corpus under `test/corpus/`, which reaches the
library only through an adapter. They are complementary _by design_. `test:coverage` ran
`node --test --experimental-test-coverage "src/**/*.test.ts"`, one of the two, and
reported the result as though it described the library.

It does not merely under-report. It reports specific, confident, wrong numbers:

| module                    | `test:coverage`                     | union of both suites      |
| ------------------------- | ----------------------------------- | ------------------------- |
| `src/core/table-style.ts` | 81.25 % lines, **0.00 % functions** | 97.22 % / 100.00 %        |
| `src/io/xlsb/formula.ts`  | 77.69 % lines, 66.67 % functions    | 97.31 % / 100.00 %        |
| `src/io/xlsb/read.ts`     | 82.83 % lines, 66.67 % functions    | 100.00 % / 100.00 %       |
| **all files**             | 97.34 / 90.64 / 96.44               | **98.90 / 92.95 / 98.84** |

`table-style.ts` reporting **0 % of functions covered** is the sharp end. Both of its
functions are exercised by the corpus (`test/corpus/adapters/ts-xlsx/styles.ts`), and
the report said nothing tested them. A zero that means _"the other suite covers this"_ is
indistinguishable from a zero that means _"nothing tests this"_, so the table actively
misdirects: an agent auditing the worst rows spends its time on modules that were never
uncovered, while a genuinely untested function sits mid-table looking fine. That is
precisely what happened to `iconSetXml`, a documented conditional-formatting rule type
that no test anywhere named, which the report showed as part of a 96.97 %-covered file.

This matters more here than in a human-paced repo. CLAUDE.md §6 optimizes for agents
working with low supervision, and §2 makes the machine-checkable net the primary
guarantor of correctness. A signal that is _confidently wrong_ costs more than no signal,
because it is acted on.

## Decision

**Coverage is measured over both suites at once, and reported as their union.**
`scripts/coverage.ts` (`pnpm run coverage`) replaces `test:coverage`.

Each suite runs as its own process with `NODE_V8_COVERAGE` pointed at one shared
directory, `.tmp/coverage/`. V8 writes one `coverage-*.json` per process keyed by pid and
timestamp, and node's coverage reader merges every such file it finds. The suites need no
knowledge of each other and stay separate processes, so **the corpus stays
implementation-blind**, the property that made the two-suite split worth having.

**The report is computed by node's own `TestCoverage`**, imported from
`internal/test_runner/coverage` under `--expose-internals`, and rendered by node's own
`getCoverageReport`. The script re-executes itself with that flag for the report phase.

**Floors are enforced on the totals**, defaulting to 98 % lines / 92 % branches / 98 %
functions, just under the measured union, so a real regression trips them and ordinary
churn does not. Overridable per-run with `--lines` / `--branches` / `--functions`.

**Floors apply to the union or to nothing.** `--suite unit` is available to see what one
suite contributes, but a partial run prints `NOT the library's coverage` and is held to no
floor. The tool will not lend its authority to a number that does not mean what it says.

**Modules absent from the table are listed by name.** V8 records only scripts it loaded,
so a module no suite imports does not appear as 0 %. It does not appear _at all_, which
reads as "no problem here". Today all twelve are type-only modules (`page-setup.ts`,
`preserved.ts`, `src/type-tests/`) or re-export barrels (`src/index.ts`,
`src/entries/*`), whose contracts are gated by `typecheck` and `check-entries.ts`
instead. A module with real behavior appearing in that list is the signal.

**It is not a `verify --full` gate.** Coverage instrumentation roughly doubles both
suites (16.8 s to 32.5 s for the units, 20.7 s to 39.8 s for the corpus) so folding it in
would take the primary gate from ~21 s wall to ~95 s. ADR 0022 bought that 21 s
deliberately; spending it 4.5× over to re-derive a number that moves a few hundredths of
a percent per change is a bad trade against the inner loop. Coverage is on demand.

## Why borrow node's implementation rather than compute our own

The numbers have to be _comparable_ to `node --test --experimental-test-coverage`, or
this tool replaces one misleading report with another: a reader seeing a different total
could not tell whether the delta was the corpus or our arithmetic, and "you cannot tell
which part of this number is real" is the exact defect being repaired.

Line coverage is not a quantity you can eyeball. It is byte-offset ranges mapped onto
lines, with `/* node:coverage ignore */` handling, block-coverage branch counting, and a
specific rule about which lines count at all. A reimplementation lands _close_, and close
is the failure mode. Borrowing the implementation makes parity true by construction: the
unit-only path through this script reproduces the old `test:coverage` figures to the
digit (97.34 / 90.64 / 96.44 on the tree where they were recorded).

The cost is a dependency on an unstable internal module. It is accepted because **it
fails loudly**: the `require` throws `Cannot find module`, and the script stops with a
message naming this ADR and refusing to approximate. Weighed against a defect whose whole
nature is _quiet_ wrongness, a break that announces itself on a node upgrade is the
cheaper risk. Node's type-stripping is what makes this work at all, because offsets in the
V8 data land on the original `.ts` bytes, since stripping preserves positions.

## Alternatives rejected

- **Reimplement the line/branch mapping.** Independent of node internals, but it produces
  numbers that silently disagree with node's, reintroducing the defect in a new place.
- **Run the corpus under `node --test` so one process measures both.** Would give a union
  for free, but only by making the corpus a child of the unit runner, collapsing the
  implementation-blind boundary that ADR 0012 exists to protect. Rejected outright.
- **Point `node --test --experimental-test-coverage` at a pre-populated directory.**
  Tried; node deliberately copies to a temp directory first, commented in its source as
  _"so that no preexisting coverage files interfere with the results"_. Not a bug to work
  around, but the runner protecting its own report.
- **Per-file floors.** Either so low they prove nothing or so exact they fail on every
  honest refactor. The per-file table is already where a single module slipping shows up.
- **Keep `test:coverage` alongside the new script.** Two names, one of which lies. The
  rename is the point; a stale invocation fails loudly with "script not found".

## Consequences

- **Positive:** the reported number is the library's number. The floor ADR 0029 deferred
  now exists and is enforceable. `--suite` makes the two suites' contributions visible
  and labels them honestly. Modules no suite loads are named rather than omitted.
  No new dependency, since the implementation is the one already in the runtime.
- **Negative / accepted:** a private node API is on the path, and will break on some
  future node, loudly and by design. A full run costs ~80 s, so it is on demand rather than
  in `verify --full`: a gate that doubles the local loop is a gate people stop running.
- **Enforced in CI since 2026-09-04.** The floors used to bind only when someone chose to
  run them, which meant a change dropping branch coverage from 95% to 60% was green
  everywhere a human or an agent looked. `corpus.yml` now carries a `coverage` job,
  separate from the gate job so it lengthens nothing local. This is the follow-up the
  paragraph above used to record as open.
- **The instrument is out of the denominator too.** `EXCLUDE` named only
  `src/**/*.test.ts`, so `package.test-support.ts` was a row in the table at 100/100/100 by
  construction and `src/type-tests/**` counted as well. It is now the same three patterns
  `tsconfig.build.json` excludes, which is the same rule said once: what is not shipped is
  not the library. The honest union is 99.29 lines / 95.33 branches / 98.65 functions over
  144 modules.
- **Revisit when:** node's internal coverage module moves (fix the import, do not
  approximate), or coverage instrumentation gets cheap enough to fold into `verify --full`.
