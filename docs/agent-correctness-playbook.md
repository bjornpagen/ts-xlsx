# Agent correctness playbook

> One page, so an agent mid-task never has to reconstruct the decision tree. It maps
> **what you are doing** to **the check that proves it correct** and the exact command.
> The capabilities themselves are described in `docs/architecture.md` and the ADRs;
> this is the dispatch table on top of them.

The net is defense-in-depth. From cheapest/fastest to most authoritative:

| Layer | What it proves | Command | Needs |
| --- | --- | --- | --- |
| Types + unit | The code compiles under strict TS and units pass | `pnpm run typecheck && pnpm run test:src` | Node 24 |
| ↳ narrower | Only one tree, when iterating | `pnpm run typecheck:src` · `pnpm run typecheck:test` | Node 24 |
| ↳ emitted `.d.ts` | The published declarations typecheck as a consumer sees them | `pnpm run typecheck:dist` | Node 24 + `pnpm run build` |
| Lint | The rule gates: correctness, imports, suppression hygiene | `pnpm run lint` | Node 24 |
| ↳ layout | Every file is as oxfmt would write it | `pnpm run format:check` | Node 24 |
| Prose | No banned character in the authored docs | `pnpm run chars:check` | Node 24 |
| **Corpus** | Well-formed XML, package structure, and no behavior regression | `pnpm run corpus` | Node 24 |
| **OOXML oracle** | Schema + semantic conformance against Microsoft's own validator | `pnpm run validate:ooxml file.xlsx` | Node 24 + network on first call |
| Spec grounding | Ground a decision in the authoritative format | `ooxml-lookup` skill + Learn MCP + `docs/knowledge/specs/` | Node 24 |
| Coverage | Which lines/branches/functions **both** suites together ever enter | `pnpm run coverage` | Node 24 (~74 s) |

**`typecheck` means both trees.** There are two strict projects, `tsconfig.json` over `src/` and
`tsconfig.test.json` over `test/`, `scripts/` and `tools/`, and the `verify` gate has always run
both. The `typecheck` *script* used to run only the first, which made the obvious command silently
blind to the tree the regression corpus lives in: edit an adapter, get a green `typecheck`, and
learn nothing. It now runs both, and `typecheck:src` is there for when you genuinely want one.
The `tsconfig.json` inside `test/`, `scripts/` and `tools/` is not a third gate and nothing runs
`tsc` against it; it exists so the linter reads the same options `tsc` does, and the note under
the planted control below says what happens when it does not.

**`typecheck` does not mean the third tree.** `tsconfig.dist.json` typechecks the *emitted*
`.d.ts` through the package's `exports` map, and its subject only exists after `pnpm run build`,
so it cannot live in `verify`, which must run on a never-built tree. It runs in `build.yml`
instead (ADR 0031). Consequence worth knowing before you push: a change that breaks the published
declarations but not `src/` passes every local gate and fails on the runner. If you are editing
the public barrel or a type it re-exports, run `pnpm run build && pnpm run typecheck:dist` first,
which costs about 0.8 s on top of the build.

**`lint:fix` needs no confirming `lint` pass.** `oxlint --fix` applies what it can and *still
exits non-zero* if any diagnostic survives, so a green `lint:fix` already is the proof. Re-running
`lint` after it only re-checks a tree you have been told is clean.

**Read what `--fix` did before you keep it.** Not every autofix is meaning-preserving, and the
rule that flags a construct is not the rule that understands it. `prefer-string-starts-ends-with`
rewrites a regex into a string method, which is only equivalent while the pattern holds no
metacharacter. `no-useless-spread` would have unwrapped `[...sheet.merges]` in
`src/io/xlsx/read.ts`, a copy that exists because the loop body splices the array it is walking.
That one is a suggestion oxlint declines to apply, which is exactly why the suggestions deserve
reading rather than a blanket `--fix`.

**Do not autofix a non-null assertion.** `?.` on an assertion that was load-bearing turns a crash
into silent wrong output. A `!` is usually a signal that an index is being carried where the
object itself could be. See `src/vba/cfb-writer.ts`. `typescript/no-non-null-assertion` is on for
`src/**` and off under `test/**`, where the assertion is an assertion about the fixture.

**A warning fails the gate.** Every oxlint invocation here passes `--deny-warnings`. Nothing is
set to `"warn"` in `.oxlintrc.jsonc` today, so it changes no current outcome. It is there so the
first rule adopted at warning severity, to stage a migration, is a gate and not a message.

**A silent type-aware rule looks exactly like a clean tree.** The rules that need type
information run only when `--type-aware` is passed, which `pnpm run lint` and verify's whole-tree
gate do and nothing else does. A bare `oxlint` runs the syntax rules alone and reports a clean
tree with a straight face. If a count looks too good, plant this in `src/` and check that it
reports exactly four findings, `no-deprecated`, `no-floating-promises`, `only-throw-error` and
`unbound-method`, then delete it:

```ts
/** @deprecated use other */
export function old(): number { return 1; }
export function other(): number { return 2; }
export function useIt(): number { return old(); }
export async function f(): Promise<void> {}
export function g(): void { f(); }
export function h(): void { throw 'a string'; }
export function i(x: {a(): void}): unknown { return x.a; }
```

It reports the same four from `test/`, `test/corpus/`, `scripts/` and `tools/`, verified in all
five trees. That it does is not free: tsgolint reads compiler options from the nearest `tsconfig.json`
that *includes* the file, which is why those three directories each carry one (ADR
[0037](./decisions/0037-the-linter-and-the-typechecker-read-the-same-tsconfig.md)). Delete one and
the linter falls back to TypeScript's defaults for that tree and disagrees with `tsc` without
saying so. `--tsconfig` does not substitute; it overrides import resolution only.

**A suppression that has outlived its cause is a lie.** `pnpm run lint` passes
`--report-unused-disable-directives`, so an `// oxlint-disable-next-line` whose rule would now
pass fails the gate. Every suppression in this tree carries its reason after a `--`; if you add
one without a reason, you have recorded that you silenced something and not why.

**Run one corpus case, not the whole corpus, while you iterate.**
`node test/corpus/run.ts --case <id-or-cluster-glob>` is well under a second against tens of
seconds for the whole corpus, and prints the case in full. `--json` gives one machine-readable report
object. The summary line reaches stdout in *every* mode, so never pipe a run through `grep` to
find a case, and never run the corpus twice to get both the detail and the tally.

**Cost is not the only thing that separates these layers. Authority is (ADR 0012).** They witness
three different things, and a lower one cannot stand in for a higher one:

- **Self-consistency.** A write→read round-trip is a fixed point of *our own* code. It
  catches *unilateral* writer/reader bugs. It is structurally blind to *correlated* ones
  (both halves wrong in compensating directions) and to anything spanning two package
  parts. Sufficient only for **intra-model** claims (a value survives, a style does not bleed).
- **Spec-conformance.** The `OpenXmlValidator` oracle and the `inspectPackage` structural
  facts. An *independent* implementation, so it breaks the round-trip correlation, but it
  enforces what ECMA-376 *states*, not what Excel *does*. Required for **single-part
  conformance**.
- **Excel behavior.** What Excel Desktop actually does. The only ground truth for cross-part
  invariants the spec omits, for example that a table's header cells must exist and match its
  column names. On a Windows+Excel host, scriptable via the Excel-oracle harness for
  state-observable behavior (ADR 0013); recorded as
  `provenance: {source: 'excel-desktop-verification'}`.

For a **cross-part correspondence**, one Excel-Desktop verification *seeds* the invariant
and a corpus fact whose shape *is* the relationship *locks* it. The `inspectPackage`
vocabulary is partitioned by part and cannot phrase most cross-part relationships yet. ADR 0012
lists the open ones.

**To run the whole net at once, use `node scripts/verify.ts`.** That is every gate above plus
`docs:check` and `constitution:check`, run concurrently, reported as one table with
per-gate timing (~14 s wall against ~27 s of serial work). Prefer it over assembling the
chain by hand, which is how `docs:check` and `constitution:check` get silently dropped.
`--quick` is the inner loop: types, unit tests, and lint scoped to your changed files, no
corpus (~5 s). Faster, but **not** a substitute for the full run. Invoke it with `node`,
not `pnpm run`, to skip ~1 s of package-manager wrapper. `pnpm test`, lefthook's `pre-push`
hook and CI's `corpus.yml` are all the same full run, so there is no second list of gates to
keep in step and adding one here is a one-line change that CI picks up. Why it is shaped this
way, including the pool width, the cache key and the incremental-`tsc` traps, is
[ADR 0022](./decisions/0022-verification-is-one-cached-parallel-entrypoint.md).

The **Stop hook** runs `verify --full --cached` at each turn boundary, so you cannot end a
turn green while regressing the corpus. `--cached` exits immediately when the working tree
is byte-for-byte what it was the last time this gate set passed. A *hit means proven*, not
skipped, because the key is the HEAD commit plus the full diff and every untracked file. A
turn that changed nothing verifiable costs ~0.3 s; one that changed anything pays the real
~13 s. The OOXML oracle is **not** in the hook, because it is slower and it spawns a large
external binary. Invoke it yourself; see below.

**Write scratch to `.tmp/`.** Probes, dumps, generated workbooks, anything regenerable
(`$SCRATCH` and `$TMPDIR` both point there; CLAUDE.md makes it the rule). It is git-ignored,
so probing leaves `git status` clean *and* costs nothing at the turn boundary: an untracked
file anywhere else is part of the cache key and buys you a full re-verify.

## Situation → check

**A corpus case failed and the message does not say where.**
Read the frames under it. A failure prints the throw's stack indented beneath `FAILED:`, and the
footer prints the command to re-run exactly the cases that failed. That matters because a case
reaches through optional chains, so a round trip that loses a sheet surfaces as `Cannot read
properties of undefined (reading 'cells')`, and nothing in that sentence distinguishes a bug in the
case, in the adapter, and in the reader. Under `--json` the stack is a `stack` field on the behavior.

**You are writing a check under `scripts/`.**
Report through `verdict()` and crash through `reportCrash()`, never `process.exit`: it truncates an
unflushed pipe, which is how a gate loses the diagnostic it just printed on exactly the runs where
someone needs it, and `verify.ts` captures every gate through a pipe. Accumulate findings into a
`problems` array rather than throwing on the first, because for a boundary check the whole list is
the diagnostic. Give the gate the same name as its `*:check` package script; that correspondence is
what lets a failure name its own re-run command, and `check-tsconfig-coverage.ts` now enforces it.

**You want to know whether something is actually tested.**
Run `pnpm run coverage`, never `node --test --experimental-test-coverage` on its own.
The library is covered by two separate suites, and each one alone reports numbers that
are confidently wrong about everything the other covers: measured on the unit suite
alone, `src/core/table-style.ts` reads **0 % of functions covered** while the corpus
exercises both of them. `pnpm run coverage` runs both and reports the union, which is
the only total that means what it says. `--suite unit` / `--suite corpus` shows what one
contributes and says so in the output; those partial runs are held to no floor. Modules
no suite loads at all are listed by name under the table rather than omitted, and that
list is where a genuinely orphaned module shows up. See ADR 0035.

**You added or changed a writer path (anything that emits XML).**
Run `pnpm run corpus`. It parses the written package and asserts well-formedness,
part/relationship/content-type structure, and element ordering
(`test/corpus/adapters/ooxml-facts.ts`), plus every behavior regression. Then run the
schema/semantic oracle on a representative file: use the **`validate-ooxml` skill**,
which emits a workbook and runs `pnpm run validate:ooxml` for you. New behavior ships
with a corpus case in the same change (use the **`write-corpus-case` skill**).

**A generated file's *content* is right but its *layout* opens wrong**, meaning a frozen header
row unpainted until you click it, a missing outline bar, no sheet selected.
Suspect an omitted **view-initialisation** fact before you suspect the data or the styles.
Excel writes `<bookViews><workbookView/>`, `tabSelected="1"` on exactly one `<sheetView>`,
and `outlineLevelCol`/`outlineLevelRow` on `<sheetFormatPr>` into every file it saves;
consumers lay the pane geometry and the outline bars out against them, so omitting them
leaves that layout uninitialised. Such a package is still schema-valid and still opens
without a repair prompt, so **neither the oracle nor `open-verdict.ps1` will flag it**. The
writer emits all three unconditionally now (`DEFAULT_WORKBOOK_VIEW`, `src/core/workbook.ts`).
That they were omitted is certain; that any *one* of them causes a given paint glitch is
inference from the diff. The single-variable A/B that would isolate it was never run, and
the original report only ever reproduced under a window geometry we could not recreate. So
if this class of symptom recurs with all three present, the cause is elsewhere: reopen the
investigation rather than assuming it regressed here.
The general move for this whole class: round-trip your output through Excel's own `SaveAs`
over COM and diff the two packages. What Excel adds unprompted is what a consumer expects
to find.

**You are writing or editing a document under `docs/`, or any prose in `src/`.**
Run `pnpm run chars:check`, or just commit: the pre-commit hook runs it over the index and
the `invariants` gate runs it over the whole tree. It bans the em dash (U+2014) and its
lookalike (U+2015) in authored prose, and it declares no autofix on purpose, because the
replacement for an em dash is a full stop, a colon, or a pair of commas depending on the
sentence, and a tool that guessed would turn a red check into worse prose. Two rules: one
over `docs/**/*.md` (`docs/api/**` included, since those pages are generated from source
JSDoc that is itself clean) and one over `src/**/*.ts`. The source rule reads whole files,
comments and string literals alike, because charcheck has no comments-only scope; an error
message is prose too, so that is the right reach rather than a compromise. `scripts/`,
`test/` and `tools/` are not covered yet and still carry the character. Both trees are at
zero today, which is what makes the gate a floor rather than a wish. Config and reasoning:
`charcheck.config.ts`.

**You are adding a throw, or touching the read path's tolerance of a foreign file.**
Two rules, one gate each. A name inside a message goes through `quoted()` from
`src/errors.ts`, never an inline `JSON.stringify`; run `pnpm run error-messages:check`, or
just let the `invariants` gate do it. And a *file* may never provoke an `authoring` failure
or a native `RangeError`/`SyntaxError`: if you are handing a file-derived value to a model
method that validates, it goes through `src/io/read-policy/read-repair.ts` -- `repairSheetName`
where there is an obviously right rewrite, `admitting` where the honest answer is that the
file does not really carry that feature. The check that proves it is a corpus case shaped
like `a-hostile-name-costs-that-name-not-the-read`: patch one attribute of a written
package, read it back, and assert the read survived *and* the model re-writes. The second
half is the one that catches the interesting failures, because a reader is allowed to
produce only values the writer can serialise.

**You are cutting a release.**
Bump `version` in `package.json`, cut `CHANGELOG.md`'s `## [Unreleased]` into the new
version's section, commit, and push. Then let CI go green *before* tagging, because the
tag is what the release names and a tag that fails its own gates is the one thing you
cannot quietly redo. Tag `vX.Y.Z`, push it, and publish a GitHub release on it: that
release event is what publishes to npm (ADR-0026), authenticated by OIDC with no
credential in the repository. **Publishing that release is the point of no return.** There
is no reviewer holding the job any more, so it goes to the registry unattended and a version
number cannot be reused. Rehearse when anything about the release is unusual, by dispatching
`publish.yml` from the tag with `dry_run` on, but note the rehearsal reaches
`npm publish --dry-run` only for a version the registry does not already serve. A green
rehearsal does mean the identity was accepted: `--dry-run` alone reports a rejected one as a
warning and exits `0`, so the job checks the exchange itself rather than trusting npm's exit
code. If the publish job fails the "tag and version must be the same claim" step, fix
`package.json` and re-tag; do not weaken the check. If it fails at `npm publish` with a
**404** on a package that plainly exists, that is npm refusing the OIDC identity, not a
missing package: read the trusted publisher on npmjs.com and check its repository, workflow
filename (`publish.yml`) and environment (`npm-publish`) against the job. A publish that has
never once succeeded is far more likely misconfigured there than here, so do not start
editing the workflow.

**You added or changed a reader path (parsing foreign XML).**
Treat all input as hostile (ADR-0004): no unbounded allocation, no entity expansion,
inflation bounded by output counted, unrecognized tokens dropped and never cast with
`as`. Add a **fixture-backed corpus case** for any real-world file shape you learn
about (`test/corpus/fixtures/<case>/…`), then `pnpm run corpus`. A round-trip case
(write → read) is the strongest reader proof.

**You just wrote an assertion that guards an invariant. Prove it can fail.**
A test that cannot fail is worse than no test: it reports a guarantee nobody is holding. Break the
thing on purpose, watch the assertion fire, restore. This has caught real theatre more than once.
An archive-length check meant to pin two writers to the same compression level passed happily at
level 9, because below roughly 200 rows every level compresses a fixture identically; a
compile-time exhaustiveness proof is only a proof once you have added a field and seen it named in
the error. Applies to every mechanism in this repo that exists to catch a future mistake: the facet
registries, the entry/layering gates, the size budgets. Cheap, and it is the difference between a
check and a comment.

**You are about to claim something is faster, or that it stops blocking the event loop.**
Measure, and distrust the first number. A bad measurement will talk you out of a correct change.
Three traps, all hit in one sitting while sizing `writeXlsxAsync`:
- **One process per case.** Running the baseline and the candidate in the same process loads the
  second with the first's GC pressure. On a ~42 MB payload that alone made the faster path look
  slower. Pass the case in on `argv` and run the script twice.
- **Never hand-roll a `setInterval` watcher for loop blocking.** It reports timer coalescing as
  blocking, and it reports `max = 0` when the loop is blocked so hard the callback never fires
  *once*, so the worst case reads as flawless. `perf_hooks.monitorEventLoopDelay` measures the
  actual thing.
- **Responsiveness and throughput are two claims.** Moving work to a worker can leave wall-clock
  untouched while cutting the longest stall from seconds to milliseconds. Say which one you
  measured; a change often buys one and not the other.

Probes go in `.tmp/`. Put the numbers in the commit or the ADR. The next agent should not have to
re-derive them to know whether the trade still holds.

**You are fixing a bug.**
Test-first. Write an implementation-blind corpus case that reproduces it
(`write-corpus-case` skill), set its `baseline` to what the code does *today*, watch it
fail, then fix until `pnpm run corpus` is green. We never fix the same bug twice.

**You need a cross-part or Excel-quirk invariant seeded, and the only ground truth is what Excel Desktop does.**
On a Windows host with Excel installed, don't do it by hand: run the Excel-oracle harness.
Write a probe (`tools/excel-oracle/probes/<invariant>.json`: a cell spec plus the cells to
observe), then `node tools/excel-oracle/run.ts <probe.json> --out test/corpus/fixtures/excel-oracle/<invariant>.json`.
It opens the file headless over COM, reads formula/value per cell, re-saves to reveal the
geometry Excel considers canonical, and writes an auditable observation sidecar. This
**seeds** the invariant only. Then **lock** it with a Tier-2 seam fact that runs in CI and
a case carrying `provenance: {source: 'excel-desktop-verification', ref: '<sidecar>'}`. The
harness is a probe, not a test: it needs Windows+Excel+`pwsh`, self-guards to a loud refusal
without them, and **never** runs in CI (`pnpm run corpus` must not depend on Excel).

If the invariant is *geometry*, such as whether an over-limit `ht`/`width` is clamped, quantized
or honoured, use the sibling probe instead, which takes a workbook you already wrote rather than a
probe spec:
`pwsh -NoProfile -File tools/excel-oracle/read-geometry.ps1 -Path <file.xlsx> [-Rows n] [-Cols n] [-NoResave]`.
It reports per-row `RowHeight`, per-column `ColumnWidth` and the sheet's `StandardHeight`/
`StandardWidth`, and re-saves a copy beside the input so you can diff the `ht`/`width` Excel itself
writes. Reading a value back is the only thing that separates a clamp from a passthrough. See
`docs/knowledge/specs/grid-geometry-limits-are-excels-not-the-schemas.md` for what it found.

Both probes answer *state-observable* questions on *one Excel build* only. See
[ADR 0013](./decisions/0013-excel-desktop-as-automatable-tier3-oracle.md)
for what is and isn't scriptable and the five standing pitfalls.

**Someone reports a workbook *looks* wrong in Excel: text missing, colours not the ones authored.**
Neither probe above can answer this, because both are state-observable and painting is not state.
Run the control before you touch the writer. Re-save the file through Excel itself
(`pwsh -NoProfile -File tools/excel-oracle/observe.ps1 -Path <wb.xlsx>`) and have the reporter try
the Excel-authored copy. **If the fault survives, nothing this library emits is in the causal path**,
and a "fix" to the writer would be a guess that outlives the report. Two such reports are already
settled this way: `docs/knowledge/specs/frozen-pane-header-ink-is-an-excel-repaint-fault.md` (a
frozen header row whose ink goes missing until clicked, which reproduces in Excel's own re-save) and
`docs/knowledge/specs/dark-mode-repaints-authored-cell-colors.md` (Dark Mode overrides every
encoding of an authored colour). Read both before opening a rendering investigation; when a genuinely
new one needs pixels, the interactive tier is the `excel-gui-automation` skill, and sampling the
rendered pixels beats describing a screenshot.

**You are unsure how an OOXML element / attribute / enum / child-ordering should look.**
Do not guess. The format is full of surprises. In order:
1. Ask the vendored **`ooxml-lookup`** skill. It holds the ECMA-376 graph as a local
   SQLite database and answers the four-hop question (element → type → base type →
   attribute group → facets) in one call, which is the join you would otherwise do by
   hand across several XSD files:

   ```bash
   S=.claude/skills/ooxml-lookup/scripts/ooxml.mjs
   node $S element x:c          # what is it: kind, type, namespace, profiles
   node $S children x:c         # legal children in schema order, with cardinality
   node $S attributes x:c       # attributes, inherited and attributeGroups expanded
   node $S values ST_CellType   # the legal value space: enum, facets, patterns, unions
   ```

   Write prefixes the way you already write them: `x:`, `c:`, `a:`, `r:`, `s:`, `v:` all
   resolve. Answers come back canonicalised (`x:c` replies `sml:c`) because `x` is also
   VML's excel namespace, and a bare name returns every match rather than guessing. This is
   **read-only reference**, not a validator; see the note below. Read
   `.claude/skills/ooxml-lookup/SKILL.md` for the rest, including direct SQL when the
   subcommands do not fit.
2. Query the **microsoft-learn MCP** (`microsoft_docs_search` / `microsoft_docs_fetch`)
   for Excel's *real-world deviations* from the standard, which is the prose the schema can't
   encode. This is enabled for the project (ADR-0007); if a run says the server isn't
   available, enable `microsoft-learn` for the project.
3. Check `docs/knowledge/specs/` for a note we already wrote on the same corner.

**You changed the build/emit path (`tsconfig.build.json`, import specifiers, a runtime reference type-stripping tolerates).**
The dev/test loop runs *stripped* `src/` `.ts`; consumers run *`tsc`-emitted* `dist/` JS, and those two artifacts can diverge. `pnpm run build && pnpm run corpus:dist` runs the full behavioral corpus against the emitted JS (`CORPUS_TARGET=dist`), not just the `smoke:dist` round-trip. CI's `build` workflow does this on every PR; run it locally when you touch anything emit-shaped.

**You added, removed or moved a public export.**
Symbols live in exactly one entry barrel under `src/entries/`, and `src/index.ts` is `export *`
over seven of the eight, so adding a name in two places does not conflict, it makes the name
*vanish* from the root specifier with no error anywhere. `node scripts/check-entries.ts` (already
in `verify --full`) is what catches that, along with an entry `package.json` forgot to publish.
Then run `pnpm run docs`. The reference is generated from the root barrel, so a symbol missing
from the diff is a symbol that fell out of the union. Error classes go in `src/entries/errors.ts`
and nowhere else (ADR-0023).

**You imported a Node built-in, or reached for `process` or `Buffer`, inside `src/`.**
`node scripts/check-browser-safe.ts` (also in `verify --full`) walks the module graph from all
seven browser-facing entries and fails on either. A bundler resolves imports rather than call
graphs, so one such import in a module a browser never calls is still enough to break a browser
build (ADR-0040). If the code genuinely needs Node, it is published from `src/entries/node.ts`
and named in `src/entries/node-unavailable.ts`, the stub a `browser` condition resolves to;
if it does not, write the platform version (`crypto.getRandomValues`, `TextEncoder`,
`src/sha512.ts`) instead.

**You changed what a module imports, and it crossed a directory.**
`node scripts/check-layering.ts` proves the graph's direction still holds. If the import turned a
`import type` into a value import, also run `pnpm run build && pnpm run size`: the per-entry
closures are the only measurement that notices a codec joining an entry's graph, since the package
total does not move when a boundary is crossed, only when something new is written.

Read the **module count** beside the kilobytes. A one-line helper hoisted into the wrong module is
how a whole path arrives: putting a relationship-type predicate in `io/opc/read-opc.ts` and calling
it from the writer added five modules and 8 KB to `/node`, because the writer then loaded the reader
and its bounded inflater to test a string suffix. A leaf importing nothing is the fix, and the count
is what says the problem was structural rather than the code being big.

**You added a type, or a field whose type is new, to something a consumer can reach.**
`node scripts/check-public-types.ts` (in the `invariants` gate) walks out of every entry barrel and
fails on a type it reaches that no entry exports: a consumer who can hold the value and cannot write
its type. Publish it from the entry that owns its concern, or decline with `@unpublished` and the
reason in its doc comment, which the run counts so the declines stay visible. Nothing else sees this:
the emitted `.d.ts` typechecks either way, and `docs:check` regenerates from the barrel, so it sees
only what the barrel already lists. If the type is a closed token union, publish its `tokenSet` guard
too; that is the standing rule and `src/entries/core.ts`'s header states it.

**You added an `AssertNever` exhaustiveness proof beside a table.**
Assert it in `src/type-tests/facet-tables.type-test.ts` as well.
`node scripts/check-facet-register.ts` requires it. The alias already fails at its own declaration,
so the gate is not about proving the table: it is about the register's claim to be the one place you
read to learn which tables carry a guarantee. It had drifted to six of eleven, and reading it told you
`RowProperties` and `PageSetup` had no proof, which was false.

**You are adding a reader for something a package part already carries.**
Look for a pass before you write a scan. `parseXmlPasses` (`src/xml/xml-read.ts`) delivers one parse
to several readers, and both the worksheet part and the workbook part are read that way: a reader
written as a `SaxPass` joins the existing parse, a reader written as a function over the part's text
adds a whole scan of it. That is not a micro-optimisation. The worksheet part is the largest in a
package and was being read five times over, which spent 45% of a large file's read on four scans that
matched no element; the workbook part had six, of which four matched nothing. The same question
applies to a binary part: a `RecordReader` is constructed per record, so anything it does eagerly is
paid once for every record in the part, including the ~700 BIFF12 types this library does not model.

**You are extending a shared enumeration, facet table, or record-type list.**
The `AssertNever` proof beside the table covers omission from *the table*. It says nothing about a
consumer that re-enumerates the same set beside it, and that is how these have actually drifted:
`grep` the new member's siblings across `src/` before you trust the build. A table whose members a
binary codec *indexes* is the sharp case, because the index list is a second enumeration by
construction; see "A table is only a single source of truth if the other copy is derived from it"
in `docs/architecture.md` for the three shapes this takes and what each one is owed.

**You are about to finish a turn / open a PR.**
The Stop hook covers every gate `verify --full` runs. For full CI parity add the one it
cannot: `pnpm run test:ooxml`. CI runs all three workflows (`build`, `corpus`,
`ooxml-validation`) regardless, so the oracle is always enforced before merge even when
you can't run it locally.

## The schema oracle, and what to do without it

`OpenXmlValidator`, reached through the `ooxml-validate` package (the `validate:ooxml` /
`test:ooxml` scripts, ADR-0002), is the **single authoritative** schema/semantic check.
The same oracle at the same Open XML SDK version serves `ts-pptx`, so the two projects
enforce one rule set rather than two that drifted apart.

There is no .NET toolchain to install: the package downloads a prebuilt, self-contained
binary on its first call, verifies its checksum and build provenance, and caches it under
`~/.cache/ooxml-validate`, so the only requirement beyond a dev install is network
access, once. Startup dominates the cost (~0.4 s, then ~9 ms per additional package), so
validate several files in one call.

Exit codes: `0` = every input clean, `1` = validation/package errors found, `2` = the tool
could not run. Every input appears in the report with an explicit `valid` flag, so a
missing entry is a broken contract rather than a clean file. Known, tracked errors are
baselined in `test/ooxml-validation/allowed-errors.json`; a *new* error fails the gate and
a *stale* baseline fails it too, so keep that file honest when you fix or introduce a
diagnostic.

If the oracle cannot be obtained (offline, say), do **not** reach for a second validator.
The vendored schema graph is deliberately **not** wired into an `xmllint`-style path
(ADR-0002, ADR-0034): a naive XSD-only pass gives false alarms and false confidence. It
can't do the semantic checks or validate the OPC parts, and the Transitional schemas are
subtly permissive. Instead: rely on `pnpm run corpus` (well-formedness plus structure)
locally, query `ooxml-lookup` and the Learn MCP to reason about correctness, and let CI's
`ooxml-validation` workflow run the authoritative oracle on your PR.

When the oracle *has* run and you have a diagnostic, `ooxml-lookup` answers the question
that follows it. Hand `explain` the diagnostic as JSON, meaning the `id`, `description`,
`xpath` and `partUri` an `ooxml-validate` report already gives you, and it returns what
*would* have been legal at that position:

```bash
node .claude/skills/ooxml-lookup/scripts/ooxml.mjs explain \
  '{"id":"Sch_UndeclaredAttribute","description":"The '"'"'foo'"'"' attribute is not declared.","xpath":"/x:workbook[1]/x:sheets[1]/x:sheet[1]","partUri":"/xl/workbook.xml"}'
```
