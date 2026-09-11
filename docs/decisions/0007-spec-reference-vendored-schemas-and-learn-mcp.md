# ADR 0007: Spec reference, vendored OOXML schemas plus Microsoft Learn MCP

**Status:** Accepted (2026-07-19) · Phase 4

## Context

Correctness for this library is defined by an external standard: ECMA-376
(Office Open XML) plus the ways Microsoft Excel deviates from it in practice.
Agents implementing parser/writer paths repeatedly need the authoritative shape
of the format, meaning element structure, attribute types, enumerations, child ordering,
and the "how does Excel _really_ do this" deltas. We already have:

- a hand-authored behavior-spec corpus (`docs/knowledge/specs/`, ~150 notes) from
  the Phase 1 harvest, and
- an independent conformance oracle, Microsoft's `OpenXmlValidator` wrapped in a
  repo-owned .NET tool ([ADR-0002](./0002-ooxml-validation-oracle.md)).

What was missing was **machine-readable ground truth to read from** while
implementing, and a low-friction way to reach Microsoft's Open-Specification
prose. The question was whether MCP servers help here, and which.

## Decision

Two complementary moves, split by the nature of the knowledge.

### 1. Vendor the ECMA-376 Transitional XSD set (static, so in-repo rather than an MCP)

The base standard is a _fixed, versioned document_. The most durable,
agent-friendly form is a pinned local copy, not a live network lookup: offline,
greppable, deterministic, version-locked, and immune to service outages, which is exactly
the self-contained, "verify don't assume" posture of CLAUDE.md.

`schemas/ooxml-transitional/` holds the complete 26-file Transitional schema set
(ECMA-376 Part 4, 5th ed., Dec 2016) **verbatim**. Transitional is what Excel
emits, so it is the set that matters for real `.xlsx`. The full set is vendored
rather than pruned to the spreadsheet closure, so the `<xsd:import>` graph stays
resolvable and provenance stays clean; `schemas/PROVENANCE.md` records source,
hashes, and licensing (ECMA free-availability plus Microsoft OSP, unmodified).

These are **reference, not a second validator**. Validation stays with the
ADR-0002 oracle. They are repo-only: the `package.json` `files` allowlist
(`dist`, `LICENSE`, `README.md`) keeps them out of the published package.

### 2. Wire the Microsoft Learn MCP (dynamic prose, so an MCP earns its place)

Microsoft's Open Specifications ([MS-XLSX], [MS-OI29500], [MS-CFB], …) carry the
implementer deltas. They live on `learn.microsoft.com` and change over time. The
official **Microsoft Learn MCP server** (`https://learn.microsoft.com/api/mcp`,
remote HTTP, no auth) gives grounded semantic retrieval over them, which beats
raw web search on citation fidelity. It is project-scoped in `.mcp.json`.

## Rejected

- **A validation MCP, or re-using XSDs as a validator.** Validation is solved
  (ADR-0002); a second path is redundant and would drift.
- **Context7 and generic docs MCPs.** They index framework/library docs, not the OOXML
  standard. No value here.
- **GitHub MCP.** The harvest is complete; `gh` via Bash covers residual needs.
- **An ECMA-376 MCP.** None exists; the standard is a static PDF plus XSD, correctly
  handled by vendoring.

## Consequences

Agents get local, pinned schema ground truth to read while implementing, plus
grounded access to Excel's open-spec deltas. The MCP is a development-time,
agent-side tool, so it adds **no** runtime or published-package dependency. The
vendored schemas must be re-extracted, never hand-edited, if ever corrected, to
preserve provenance. Contributors who do not enable the MCP simply lose the
grounded-search convenience; nothing in the build or tests depends on it.
