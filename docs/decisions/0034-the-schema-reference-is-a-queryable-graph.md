# ADR 0034: The schema reference is a queryable graph, not a vendored XSD dump

**Status:** Accepted (2026-08-15) · **supersedes part 1 of** [ADR 0007](./0007-spec-reference-vendored-schemas-and-learn-mcp.md),
whose part 2 (Microsoft Learn MCP for the evolving prose) is unchanged.

## Context

[ADR 0007](./0007-spec-reference-vendored-schemas-and-learn-mcp.md) got the posture right
and the artifact wrong.

The posture: the base standard is a _fixed, versioned document_, so the durable
agent-friendly form is a pinned local copy, offline, deterministic, version-locked, immune
to service outages. That reasoning has not weakened. What has changed is the one line in
its **Rejected** list:

> **An ECMA-376 MCP.** None exists; the standard is a static PDF plus XSD, correctly handled
> by vendoring.

One exists now. `ooxml-ai-tooling` ingests the ECMA-376 XSDs into a SQLite graph and serves
it through two surfaces, an MCP server and an agent skill, both fully local, both with no
account and no network. The premise that forced the choice is gone.

Which matters, because vendoring the raw XSDs has a cost ADR 0007 counted as a benefit.
"Greppable" is the word it used. In practice greppable is the failure mode: `rg tblPr
schemas/` returns a matching line with no declaration around it, so the next move is `Read`
on a file of several hundred KB, and the question an agent actually has (_what may go
inside `x:c`, in what order, with what attributes_) is a four-hop join across element,
type, base type, `attributeGroup` and `simpleType` facets, spread over several files, done
by hand, in context, every time. It works. It is enormous, and it is silently wrong when the
agent stops one hop early at an unresolved base type.

A binary `.db` has no cheap wrong path. `rg` says `binary file matches` and stops, `Read`
refuses, and the only way in is the designed one: a CLI that answers the four-hop question
in a single call and returns JSON.

## Decision

Delete `schemas/`. Vendor the `ooxml-lookup` skill into `.claude/skills/ooxml-lookup/`
instead, alongside the three skills already there.

```
.claude/skills/ooxml-lookup/
  SKILL.md
  scripts/*.mjs           the query CLI, zero runtime dependencies
  scripts/data/ooxml.db   the ECMA-376 graph, Transitional and Strict
  .clawhub/origin.json    the registry version and fingerprint it came from
```

`node .claude/skills/ooxml-lookup/scripts/ooxml.mjs children sml:c` and friends. Node 24+,
which this repo already requires; nothing to install, nothing to configure, and no network
at any point. Every command prints JSON on stdout.

This is the **same posture ADR 0007 chose**, a pinned artifact in the tree, applied to a
better artifact. It is _reference_, not a validator: conformance validation stays with the
independent `OpenXmlValidator` oracle ([ADR 0002](./0002-ooxml-validation-oracle.md), as
mechanised by [ADR 0033](./0033-the-ooxml-oracle-is-a-shared-package.md)), and the upstream
project draws that boundary in its own README. Do not wire the graph into a second
validation path; that rejection is as live as it was in ADR 0007.

Repo-only, like the XSDs before it: the `package.json` `files` allowlist (`dist`, `LICENSE`,
`README.md`) keeps it out of the published package.

## Rejected

- **Keep the XSDs as well.** Two copies of the same ground truth, one of which is the cheap
  wrong path. The point of the change is closing that path, not adding an alternative to it.
- **The MCP server (`mcp-server-ooxml`) instead of the skill.** Equivalent answers, and it
  is what `../ts-pptx` now uses. Rejected here because it puts the schema behind a server
  the agent has to have loaded (`.mcp.json` plus `enabledMcpjsonServers`, with a silent
  degradation to "no answer" when it is not) whereas a vendored skill is present for
  anyone who clones the repo, which is exactly the property ADR 0007 was protecting. Two
  consumers on two surfaces also means both surfaces get real use.
- **Install from ClawHub without committing** (`clawhub install @shbernal/ooxml-lookup`,
  gitignored). Zero repo weight and a one-line version bump, but it makes the capability a
  per-contributor setup step and reintroduces the dependency on a service being reachable,
  the two things ADR 0007 vendored the XSDs to avoid.

## Consequences

**Bigger in the tree, smaller in context.** 2.3 MB of committed skill against 976 KB of
XSD, and the `.db` is a binary blob that will not delta-compress across rebuilds. That is
the trade: git history pays once per schema-graph correction so that every agent turn stops
paying to hand-join XSDs.

**Updating is manual and deliberate.** No lockfile and no bot watches this. Re-vendor with
`clawhub install @shbernal/ooxml-lookup --version <v>` and copy the result over;
`.clawhub/origin.json` records the version and fingerprint currently in the tree.
`ooxml-lookup` is pre-1.0 and its answers may change shape without a deprecation period, and
[its changelog](https://github.com/shbernal/ooxml-ai-tooling/blob/main/CHANGELOG.md) is
where that is announced.

**Prefixes: write them the way you already write them.** `x:c` and `c:ser` resolve, as do
`w:`, `a:`, `p:`, `s:`, `m:`, `r:` and `v:`. Answers come back in the graph's canonical
spelling, so `x:c` replies `sml:c`, because `x` is _also_ VML's excel namespace and printing
both as `x:` would render two different namespaces identically. A bare name returns every
match rather than guessing, which is often the faster way in.

This needed an upstream fix. Until `ooxml-lookup` 0.0.4 the graph recorded only prefixes
_observed_ in the XSDs, and nothing binds one to SpreadsheetML's own namespace, so `x:`
resolved to VML and `c:` to nothing at all, which also meant `explain` could not read a
spreadsheet diagnostic, since `ooxml-validate` writes those as `/x:worksheet[1]/…`. Vendor
0.0.4 or later; **0.0.3 on ClawHub is broken and ships no database at all.**

**Strict is available now**, which the vendored set deliberately excluded.
`ooxml.mjs diff <qname>` answers "will this still be valid in Strict" for a name. Nothing in
this repo targets Strict, since Excel emits Transitional, so this is upside, not a new
obligation.

**Two breadcrumbs now dangle.** [ADR 0010](./0010-agent-correctness-dispatch.md) and
[ADR 0033](./0033-the-ooxml-oracle-is-a-shared-package.md) cite `schemas/README.md` for the
"read-only reference, never a second validator" rule. Their bodies are left alone, per the
retraction convention in [the index](./README.md): the rule they cite survives verbatim,
one paragraph up in this record, and it is the rule rather than the file that mattered.
