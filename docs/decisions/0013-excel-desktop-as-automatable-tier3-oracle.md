# ADR 0013: Excel Desktop is an automatable Tier-3 oracle for state-observable behavior

**Status:** Accepted (2026-07-21) · Phase 4 · amends [ADR 0012](./0012-three-tiers-of-correctness-evidence.md) (three tiers of evidence)

## Context

[ADR 0012](./0012-three-tiers-of-correctness-evidence.md) named Tier 3, *what Excel
Desktop actually does with a file*, as the only ground truth for cross-part invariants
the spec omits, and described it as "**not cheaply automatable; reached by hand and
captured as provenance.**" That framing made every Tier-3 seed a manual, folkloric step:
open the file, squint at the cells, remember what you saw. The open shared-formula `ref`
hazard in ADR 0012 sat unresolved precisely because "someone with Excel Desktop" had to
do that by hand and nobody had.

That "by hand" premise turns out to be too pessimistic. On a Windows host with Excel
installed, Excel Desktop is COM-automatable **headless** (`Visible=$false`,
`DisplayAlerts=$false`, clean quit). A script can open a package, read `.HasFormula` /
`.Formula` / `.Value2` per cell, re-save a copy and diff the geometry Excel itself wrote,
and capture `Application.Version` / `Build`, which are the exact observations a Tier-3 seed
needs, without a human watching. We proved this end-to-end on the very hazard ADR 0012 left
open (see Decision), so this ADR promotes Tier 3 from "manual folklore" to "scriptable probe,
with a sharply drawn boundary."

## Decision

**On a Windows+Excel host, Tier 3 is scriptable for state-observable behaviors and is
run through a dedicated probe harness. But it is never a test, never runs in CI, and
never replaces the Tier-2 seam fact that locks a behavior.**

### The harness

`tools/excel-oracle/` is a standalone probe tool, outside the corpus run path:

- `emit-probe.ts` is a narrow, strictly-typed spec-to-`.xlsx` builder, a deliberate small
  re-expression of the corpus adapter's `buildFrom` rather than a reuse of that `any`-typed
  corpus-internal machinery.
- `observe.ps1` is the COM driver. It owns **every** safety guardrail (below) and emits one
  JSON observation blob: `{version, build, openThrew, repaired, cells:[{address, hasFormula,
  formula, value}], resaved…}`.
- `read-geometry.ps1` is a sibling COM driver answering the *geometry* question instead of the
  cell one: per-row `RowHeight`, per-column `ColumnWidth`, the sheet's `StandardHeight`/
  `StandardWidth`, and the same optional re-save. It takes an existing workbook path rather than a
  probe spec, since the fixtures worth asking this of are ones a writer already produced, so it is
  run directly and carries its own Excel-present guard. Same guardrails, verbatim.
- `run.ts` orchestrates emit, observe and collect (canonical-ref readback happens in Node
  via `fflate`), stamps the observation with `probeSpecRef` and the authored `verdict`, and
  **self-guards**: it refuses to run, loudly and non-zero, if `pwsh` or a registered Excel
  COM server is absent, so on a non-Excel host it degrades with a clear message rather than
  silently emitting empty facts.

One command: `node tools/excel-oracle/run.ts <probe.json>` gives observation JSON out, and
`--xlsx <file.xlsx> --cells …` does the same for a package that already exists rather than one
emitted from a spec (added 2026-09-11, after a run of reader questions about shapes the writer never
produces had each needed a one-off COM script). For
geometry, `pwsh -NoProfile -File tools/excel-oracle/read-geometry.ps1 -Path <file.xlsx>` gives
observation JSON out.

### What is and isn't scriptable

**Scriptable (state-observable):** auto-fill and materialization (did Excel add cells?),
the canonical geometry Excel re-saves, per-cell value/formula readback, whether `Open`
threw, and whether the workbook name carries `[Repaired]`.

**Not scriptable (interactive-only):** the modal *repair dialog experience* itself.
`DisplayAlerts=$false` **suppresses** that modal, which is what keeps the agent from
deadlocking, so the harness can *detect that a repair happened* (Open throws, the name
carries `[Repaired]`, re-saved content diverges) but cannot reproduce what a human clicking
through the dialog would see. **Automation-open is not interactive-open.** Every observation
records which class it is (`openClass` field).

### The seed-once / lock-in-CI split is unchanged (ADR 0012)

Excel is **Windows/machine-bound and never runs in CI.** A harness observation is captured
**once** as recorded provenance that *seeds* a case; a **Tier-2 seam fact is what locks it**
and runs in CI. `node test/corpus/run.ts` must never depend on Excel being installed. The
harness lowers the *cost* of a Tier-3 seed; it does not move Tier 3 into the CI gate.

### The five standing pitfalls (the contract for using the harness)

1. **Automation is not interactive.** Suppressed modals mean detectable-but-not-reproducible
   repair. Record `openClass`; never claim the interactive experience from an automation run.
2. **Environment-bound.** The observation is only as portable as the Excel build that made
   it. Record `version` and `build` in every sidecar; a different build may behave differently.
3. **A stray modal deadlocks the agent forever.** Mandatory guardrails, all owned by
   `observe.ps1`: `DisplayAlerts=$false`, `AutomationSecurity=msoAutomationSecurityForceDisable`,
   `AskToUpdateLinks=$false`, a wall-clock watchdog (`Wait-Job -Timeout`) that force-kills a
   hung `EXCEL.EXE`, and guaranteed teardown (`Quit`, `ReleaseComObject`, `GC`, and an orphan
   sweep scoped to PIDs this run spawned).
4. **Provenance is the deliverable.** The point of a run is a durable, auditable sidecar
   (`test/corpus/fixtures/excel-oracle/<invariant>.json`), not console output. It carries
   `{excel:{version,build}, capturedAt, probeSpecRef, cells, resave, verdict}`.
5. **Don't over-claim.** The harness answers *state-observable* questions on *this build*.
   It is not a conformance oracle (that is Tier 2, ADR 0002) and it is not proof of what
   every consumer does (LibreOffice materializes shared-`ref` interiors where Excel does not).

### Seeding provenance convention

A case seeded from a harness run carries
`provenance: {source: 'excel-desktop-verification', ref: '<sidecar path>'}`, so the audit
trail points from the CI-run seam fact back to the one-time Excel observation that justified
it. The sidecar in turn carries `probeSpecRef` back to the probe that produced it.

## The seed this ADR was proven on

The open hazard in ADR 0012, where `planSharedFormulas` emits `ref="B1:D5"` for a non-contiguous
group (master `B1`, clones `B2` and `D5`), covering empty interior cells, was seeded through the
harness (sidecar: `test/corpus/fixtures/excel-oracle/shared-formula-sparse-ref.json`, Excel
16.0 build 20131). The verdict inverts the presumed fix direction:

**Excel treats the shared-formula `ref` as a bounding-box hint, not an assertion that every
enclosed cell is a clone.** It opened the package **without repair**, did **not materialize**
the empty interior cells, and **re-saved a byte-structurally identical group** (same
`ref="B1:D5"`, same `si="0"`, the same two clones). ts-xlsx's output is already Excel's own
canonical form, so the ADR 0012 candidate fixes (split into contiguous runs, or degrade clones
to standalone `<f>`) would make ts-xlsx **diverge** from Excel and are wrong for this geometry.
This is now *locked* by the Tier-2 seam fact
`shared-formula-sparse-ref-matches-excel-canonical.case.ts`, which asserts our emitted geometry
(one master, `ref="B1:D5"`, exactly the two authored clones as slaves) matches Excel's canonical
re-save and runs in CI without re-opening Excel. ADR 0012's hazard is closed.

## Consequences

- **Positive:** Tier-3 seeds drop from "a human opens Excel and remembers" to one
  reproducible command that yields an auditable sidecar. The open hazard that motivated
  ADR 0012 is now seeded (verdict: benign), not folklore.
- **Positive:** the capability is discoverable. A playbook dispatch row points to it and
  this ADR states its limits, so a fresh agent finds both the tool and its boundary.
- **Neutral / unchanged:** Tier 3 is still irreducible and still never in CI. The harness
  makes the *seed* cheap; the **seam fact is still what keeps a behavior locked**. On a
  non-Windows or non-Excel host the harness self-guards to a clear refusal, so the seed simply
  cannot be taken there, which is honest, not a regression.
- **Negative:** a second machine dependency (Windows plus Excel plus `pwsh`) now exists for
  *seeding*. It is quarantined to `tools/` and gated only by typecheck and lint, never by the
  corpus or a CI job, so it cannot make the corpus depend on Excel.
