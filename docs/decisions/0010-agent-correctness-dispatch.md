# ADR 0010: Make the correctness net easy for agents to dispatch

**Status:** Accepted (2026-07-20) · Phase 4 · builds on [ADR 0002](./0002-ooxml-validation-oracle.md) (validation oracle) and [ADR 0007](./0007-spec-reference-vendored-schemas-and-learn-mcp.md) (spec reference plus Learn MCP)

> **One premise below expired on 2026-08-15.** Everything here that reasons from "the
> oracle needs .NET", meaning the friction it names, the "no-.NET answer", and the
> Stop-hook exclusion's stated reason, was true of the repo-owned tool ADR 0002 specified.
> [ADR 0033](./0033-the-ooxml-oracle-is-a-shared-package.md) replaced it with a
> downloaded binary, so there is no .NET requirement to route around. The decisions
> themselves stand: the dispatch table, the skill, the corpus-in-the-hook gate, and
> emphatically **the declined second XSD validator**, whose reasoning never
> depended on .NET being awkward.

## Context

The project's machine-checkable safety net is deep: the corpus (well-formedness, package
structure, regression), the Microsoft `OpenXmlValidator` oracle, vendored ECMA-376
XSDs, the microsoft-learn MCP, and the `docs/knowledge/specs/` notes. But the knowledge
of _which check to run when_ was spread across CLAUDE.md, `docs/architecture.md`, four
ADRs, two skill files, and two READMEs. An agent mid-task had to reconstruct that
decision tree every time, and the highest-value tool, the schema oracle, had no skill
wrapping it and the most operational friction (.NET requirement, emit-a-file-first,
exit-code and baseline semantics). Two smaller leaks compounded it: the turn-boundary
Stop hook ran only typecheck plus test:src, so a turn could end "green" while regressing
the corpus; and the microsoft-learn MCP was enabled only in the git-ignored
`settings.local.json`, so a fresh clone or agent didn't get it.

The gap was **dispatch and discoverability, not capability.** The fix is to surface the
existing net, not to add new validators.

## Decision

### A dispatch table agents actually reach for

`docs/agent-correctness-playbook.md` maps _what you are doing_ → _the check that proves
it correct_ → _the exact command_, cheapest-to-most-authoritative. CLAUDE.md §3 points
to it from the "Verify, don't assume" rule, and the closing See-also lists it beside
`architecture.md` and the ADRs.

### A skill for the oracle

`.claude/skills/validate-ooxml/` wraps `OpenXmlValidator` end to end: emit a repro
workbook from the `src` public API, run `pnpm run validate:ooxml`, read exit codes
(`0`/`1`/`2`), and keep `test/ooxml-validation/allowed-errors.json` honest. Skills
advertise themselves by description, so this is the primary streamlining win. It joins
`write-corpus-case` and `harvest-triage` as the third agent-facing workflow.

### Corpus joins the turn-boundary gate

`.claude/settings.json` Stop hook now runs typecheck, test:src and **corpus** when `src/`
is dirty (timeout 120 s raised to 180 s; measured ~23 s combined, 671 green, 0
regressions). Corpus defaults to the `rewrite` adapter, so it genuinely exercises `src`.
This closes the "finished green but regressed" gap. The .NET oracle is deliberately
**not** added to the hook, because it needs .NET and is slower; CI's `ooxml-validation`
workflow enforces it on every PR regardless.

### The MCP enablement is now shared

`enabledMcpjsonServers: ["microsoft-learn"]` moves into the committed
`.claude/settings.json`, so every clone and agent gets the ADR-0007 spec-grounding MCP
(subject to the usual first-use approval prompt) instead of relying on a git-ignored
local file.

### Declined: a second, dep-light XSD validator

The obvious "make it reachable without .NET" move, wiring the vendored XSDs into an
`xmllint --schema` path, was **considered and rejected.** It directly contradicts the
standing decision in `schemas/README.md` and ADR-0002: the XSDs are read-only reference,
and a second validation path is a liability, not a convenience. XSD-only validation
can't run the semantic checks or validate the OPC parts, and the Transitional schemas
are subtly permissive, so it produces both false alarms and false confidence, and two
oracles that disagree is worse than one authoritative oracle plus an honest "not
available here." The no-.NET answer is therefore documented, not automated: rely on
corpus locally, read the XSDs and Learn MCP to reason, and let CI run the oracle. Reversing
this would mean overturning ADR-0002's single-oracle stance, so do not wire a second
validator absent that explicit decision.

## Consequences

- **Positive:** the whole net is legible from one page; the schema oracle has a
  self-advertising skill; the corpus can't silently regress across a turn; the spec MCP
  is on by default. No source code changed and no new runtime or dev dependency was added.
- **Neutral:** the Stop hook is ~15 s slower when `src/` is dirty, acceptable against
  its 180 s budget, and it only fires at turn boundaries, not per edit (the per-edit gate
  was already retired for false-alarm noise; corpus at turn boundary does not reintroduce
  that).
- **Revisit the XSD-validator decision only if** ADR-0002's single-oracle stance is
  itself revisited. Do not re-litigate absent that.
