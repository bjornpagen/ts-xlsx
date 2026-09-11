# ADR 0022: Verification is one cached, parallel entrypoint

**Status:** Accepted (2026-07-27) · implements the dispatch table of [ADR 0010](./0010-agent-correctness-dispatch.md) over the toolchain of [ADR 0029](./0029-toolchain-standup.md)

## Context

The gate set is cheap per gate and expensive per session. Nothing was named "the gates".
Agents assembled the chain by hand each time, spelled it differently each time, and
sometimes dropped `docs:check` or `constitution:check` from it; the turn-boundary hook
then re-ran a subset of the same work seconds later. Measured over one long session,
verification took roughly half of all shell wall time to produce fewer than a dozen
genuine checkpoints.

The fix is not weaker gates. It is one name for the gate set, run concurrently, and
never re-run against a tree already proven green.

## Decision

### One entrypoint, two depths

`scripts/verify.ts` owns the gate set. `--full` **is** the definition: `pnpm test`,
lefthook's `pre-push`, CI's `corpus.yml`, and the Stop hook are all that same command, so
they cannot disagree. `--quick` (types, unit tests, lint scoped to changed files) is the
inner loop and is explicitly **not** a substitute, because it has no corpus. The granular
`package.json` scripts stay as the way to run _one_ gate by hand; nothing automated
composes them into a set.

CI enumerated the gates as one step each for a single real benefit, a failing gate naming
itself in the Checks UI rather than hiding in the log, and paid for it with a second
definition that nothing kept in step. Workflow commands buy the annotation without the
copy: under `GITHUB_ACTIONS`, verify emits `::error title=verify: <gate>::` per failure
with the diagnostics in a collapsed `::group::`. The enumeration is gone.

Gates are spawned from `node_modules/.bin` directly. `pnpm run <script>` costs ~1 s of
pure wrapper per invocation, which is why the docs tell agents to call `node
scripts/verify.ts` rather than `pnpm run verify`.

### Concurrency is bounded at two, and that is not core starvation

The gates are not single-core jobs: `node --test` forks a process per test file and
Biome saturates every core. Running all of them at once inflated each gate about 2× and
produced a _worse_ wall than two at a time. Measured on a 14-core box: jobs=1 29.8 s,
jobs=2 20.7 s, jobs=3 19.6 s, jobs=7 19.2 s (±25 % run to run), so everything past two is
inside the noise. Two `tsc` runs racing cost 1.87 s against 1.57 s alone, so the
contention is filesystem and scheduler, not CPU count. `--jobs` overrides for anyone who
wants to re-measure on different hardware; do not replace the constant with a
core-count heuristic without doing so.

The two typecheck configs are **one gate with two sequential steps**, not two gates:
they both re-read all of `src/` and were the worst-contending pair.

### Incremental typecheck, but not `composite`

Both configs set `incremental` with buildinfo under `node_modules/.cache/`, which halves a
warm `tsc`. `composite` plus `references`, which would let the harness project consume src's
declarations instead of re-elaborating them, **is not available here**: a referenced
project may not disable emit (`TS6310`), and these configs exist to gate a tree that is
never built ([ADR 0001](./0001-rewrite-runtime-and-toolchain.md)). Do not revisit without
also giving up the build-free dev path.

Two traps this laid, both now pinned by comments at the site:

- `tsconfig.build.json` extends the gate config and so inherited `incremental`. Buildinfo
  lives outside `dist/`, so after `clean` tsc concluded the unchanged sources needed no
  emit, and said nothing: **158 files, then 0 files, exit 0 both times.** It pins
  `incremental: false`.
- A diagnostic _replayed_ from buildinfo exits `1`; a fresh elaboration exits `2`.
  Everything here tests for non-zero. Anything that keys on `2` would break.

### Lint fails on warnings, or most of the rule set is enforced by nothing

Biome exits `0` when every diagnostic is a warning, and most of its `style` group, including
`noNonNullAssertion`, is a warning by default. The gate was therefore green on a tree carrying
18 of them, against CLAUDE.md §2's "No warnings". Every gate that runs Biome passes
`--error-on-warnings`, so severity is a presentation choice and never an enforcement one.

### Lint is scoped by explicit file list, and needs no confirming pass

`--quick` passes Biome the changed files (`git diff`, `--cached`, untracked), filtered to
the lint roots: ~0.5 s against ~6.5 s whole-tree, falling back to the whole tree past 100
files. **`biome check --changed --since` is not the way.** Measured at 5.3 s, the git
walk eats the saving. And `biome check --write` still exits non-zero when a diagnostic
survives, so a green `lint:fix` is already the proof and the confirming `lint` run is
waste.

### A cache hit is a proof, not a skip

`--cached` records a stamp under `.tmp/` keyed by the HEAD sha, the full `git diff HEAD
--binary`, and the path and content of every untracked file. Same key means the gates
would be handed the same bytes, so their verdict is already known. This is why the Stop
hook can be fast without becoming advisory.

The key is content-derived with **no path allow-list**: an unrelated edit costs one extra
run, which is the safe direction to be wrong in. A failing run records nothing, so a
regression cannot be cached green.

The hook runs `--full --cached`, not `--quick`. Pointing it at `--quick` would drop the
corpus and with it the "cannot end a turn green while regressing" property the
playbook advertises. The hook must get its speed from the cache, never from a smaller
gate set. It now checks _more_ than the hand-rolled chain it replaced (six gates instead
of three, and it fires on any dirty path rather than only `src/`) at less wall time.

Deliberately outside the cache: lefthook's `pre-push`, which runs once per push and is the
last check before code leaves the machine. Not covered by the key: a `node_modules` change
with no lockfile change.

### The OOXML validator builds on demand and is then invoked as an assembly

> **Obsolete since 2026-08-15.** There is no local build to be stale:
> [ADR 0033](./0033-the-ooxml-oracle-is-a-shared-package.md) moved the oracle into the
> `ooxml-validate` package, which resolves a prebuilt binary. `scripts/ooxml-validator.ts`
> and its `resolveValidator()` are gone. Kept because the reasoning below, namely the cost of
> re-evaluating a project per call and MSBuild reporting errors on the stream reserved
> for report JSON, is why the resolution problem was worth solving somewhere rather than
> living with.

`dotnet run --project` re-evaluates the project on every call, taking ~2 to 6 s for a tool whose
actual work is under a second. `scripts/ooxml-validator.ts` compares the DLL's mtime
against the newest source beside the project, builds only if stale, and otherwise hands
`dotnet` the assembly path. `resolveValidator()` is exported so the baseline harness
resolves the same assembly rather than keeping a second copy of the staleness rule.

The locked restore is folded in as `-p:RestoreLockedMode=true` on that build, which is
sound only because `packages.lock.json` is itself a staleness input. When it fails it
fails loudly: **MSBuild reports its errors on stdout**, which this process reserves for
report JSON, so build stdout is captured and replayed to stderr rather than discarded.

## Consequences

- **Positive:** one name for the gate set, so the hook, the hook's human, CI's intent and
  lefthook cannot drift apart. A clean-tree `--full` is ~14 s against ~27 s of serial
  work; a turn that changed nothing verifiable costs ~0.3 s.
- **No second definition left to drift:** `corpus.yml` is one `node scripts/verify.ts --full`
  step. Adding a gate is a one-line change in one file and CI picks it up. `build.yml` still
  enumerates its own steps, which is not the same duplication: it gates the _emitted_ package
  (build, smoke, corpus-on-dist, size budget), a set `verify` does not contain.
- **`--jobs` is not tuned for the runner.** The default of 2 was measured on a 14-core box; a
  2-core GitHub runner is a different machine. Measure before assuming it is optimal there.
- **Revisit `composite` when** the dev/test loop stops being build-free, not before, and
  not without reproducing `TS6310` first.
