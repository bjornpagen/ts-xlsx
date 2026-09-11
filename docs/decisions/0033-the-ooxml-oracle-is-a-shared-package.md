# ADR 0033: The OOXML oracle is a shared package, not a repo-owned .NET tool

**Status:** Accepted (2026-08-15) · **supersedes the mechanism of** [ADR 0002](./0002-ooxml-validation-oracle.md),
whose stance (Microsoft's `OpenXmlValidator` is the single authoritative conformance oracle, at
`Microsoft365`, as development-only tooling) is unchanged.

## Context

ADR 0002 gave this repo a `tools/ooxml-validator/` .NET console tool around
`OpenXmlValidator`, pinned to `DocumentFormat.OpenXml` 3.5.1 with a committed NuGet lockfile.
It worked. Two things it could not fix on its own:

- **`ts-pptx` had one too, and it was not the same one.** That project wrapped a third-party
  binary pinned to `DocumentFormat.OpenXml` **3.2.0**, reported clean files by _omitting_ them,
  and exited `0` no matter what happened. Two sibling projects were validating OOXML against
  **different rule sets** while both describing it as "Microsoft's validator". Nothing in either
  repo could notice that, because neither could see the other's pin.
- **Building it needed the .NET 10 SDK.** That is a real barrier for contributors and an
  `actions/setup-dotnet` step in CI, for a tool whose _output_ is all anyone wants.

`ooxml-validate` was built to remove the first problem: one oracle, one pin, one report
contract, distributed as a prebuilt binary. Its own code descends from this repo's
`Program.cs`, so the generalisation went upstream rather than the other way around.

## Decision

Depend on **`ooxml-validate`** as a dev dependency and delete the repo-owned tool.

That package now owns: the .NET oracle and its `net10.0` plus `DocumentFormat.OpenXml 3.5.1` plus
`packages.lock.json` pin; binary distribution (self-contained builds published per platform,
fetched on first use, checksum- and provenance-verified, cached in `~/.cache/ooxml-validate`);
the `Microsoft365` conformance pin; batching; and the CI gate that turns an unobtainable
oracle into a hard failure instead of a silent skip.

This repo keeps what is actually about _this_ project: `test/ooxml-validation/run.ts`, which
emits the buffered and both streaming writers' real output plus the negative controls, and
`allowed-errors.json`, the frozen baseline it holds them against.

Unchanged from ADR 0002 and worth restating because it is easy to assume otherwise: the
conformance target is `Microsoft365`; this is development and CI tooling only, and the
published package gains no dependency on any of it; the vendored XSDs are still not a
validator ([ADR 0007](./0007-spec-reference-vendored-schemas-and-learn-mcp.md),
`schemas/README.md`).

## Consequences

**No .NET on this repo's critical path.** `global.json`, `actions/setup-dotnet` and the NuGet
lockfile are gone. Anyone with a dev install can now run the authoritative oracle, which was
previously the one gate a contributor could be locked out of.

**The baseline did not move.** This repo was already on 3.5.1 at `Microsoft365`, so adopting
the shared oracle changed no diagnostic: `allowed-errors.json` is byte-identical and still
empty. Had it moved, that would have been a bug in the shared oracle rather than something to
baseline here.

**The pin left the building, and that is the trade.** An SDK bump in `ooxml-validate` now moves
this repo's baseline from a repository this one does not gate. That is the deliberate cost of
having one rule set instead of two, and it is not unmanaged: `ooxml-validate` carries its own
fixture corpus and a committed snapshot of the diagnostics they produce, so a bump PR there has
to show the delta in its own diff before it can land. What this repo gives up is the ability to
pin _differently_, which is exactly the ability that caused the divergence.

**Renovate's subject changes.** [ADR 0027](./0027-dependencies-are-updated-by-a-bot-and-ci-is-the-reviewer.md)
listed the NuGet lockfile as a thing to watch here; there is no longer one. The oracle now
arrives as an ordinary npm dev dependency, and its npm version _is_ its binary's version, so a
bump of that one number moves both halves at once.

**What the wrapper script was quietly doing.** `scripts/ooxml-validator.ts` shifted off the
`--` that pnpm forwards to a script, so deleting it broke `pnpm run validate:ooxml --
file.xlsx` while leaving the plain spelling fine. Re-adding a wrapper for one character would
have put a per-repo patch back on top of the shared CLI's contract, the exact shape of the
problem this ADR is about, so the fix went upstream: `ooxml-validate` 0.0.3 treats a bare
`--` as end-of-options, and that is the floor this repo depends on. The instance is trivial;
the rule it illustrates is not. What a shared tool should do belongs in the shared tool, even
when the local patch is smaller.
