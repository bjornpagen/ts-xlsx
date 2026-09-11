# ADR 0027: Dependencies are updated by a bot, and CI is the reviewer

**Status:** Accepted (2026-07-30) · maintenance slice, after the 1.0.2 release path settled ·
**amended** 2026-08-15 by [ADR 0033](./0033-the-ooxml-oracle-is-a-shared-package.md)

> **What 0033 changed here.** The oracle stopped being a `tools/ooxml-validator` .NET
> project and became the `ooxml-validate` npm devDependency, so there is no NuGet
> lockfile left to watch (the "Watch for" note below now has no subject), and the
> stop-and-wait class named "the .NET validator and SDK" is enforced differently: the
> devDependency automerge grant would otherwise sweep the oracle up with the linters, so
> `renovate.json5` names it and switches automerge back off. The rule is unchanged and the
> reason is the one stated below, that the thing that changed is the judge.

## Context

A large part of why this fork exists is that upstream's dependency tree rotted in place,
not because anyone decided to keep old packages, but because nobody decided anything for
two years and the default of a dependency is to age. This repository is three months old
and already carries a runtime dependency, five devDependencies that _are_ the quality
gates, a .NET oracle, a pinned package manager, and four workflows' worth of GitHub
Actions. Nothing about being young protects it from the same failure.

The obvious counter, "we'll update them when we touch them", is precisely the policy
upstream had. It fails silently, and the moment it has failed is not observable.

The complication is that this project is deliberately human-minimal. A bot that opens
fifteen pull requests a month solves the rot and replaces it with a review queue nobody
drains, which is the same problem wearing a hat. So the question is not whether to
automate the updates; it is **who reviews them**, given that the answer cannot usually be
"a person".

## Decision

**Renovate runs on this repository, and the gate set is the reviewer.** Configuration is
`.github/renovate.json5`, in JSON5 so the reasoning lives next to the setting it explains.

1. **An update that survives the gates is merged without a human.** `verify --full`, the
   regression corpus on both `src/` and the emitted artifact, the size budget and the OOXML
   oracle all run on every pull request. That is the same evidence a human reviewer of a
   version bump would ask for and the only evidence they could actually check. This applies
   to **non-major devDependencies**, **GitHub Actions minor/patch/digest**, and **weekly
   lockfile maintenance**.

2. **Four classes stop and wait**, because CI cannot judge them:
   - **`dependencies`.** The runtime tree is what consumers install. A green corpus proves
     it works here, not that the version range we publish is one we meant to publish.
   - **Every major**, by Renovate's default. A major is a claim that something changed
     enough to break, and the gates confirming it did not is an argument, not a decision.
   - **The .NET validator and SDK.** A green run after bumping the oracle proves less than
     one after bumping the library, because the thing that changed is the judge.
   - **The Node floor**, meaning `engines`, `.nvmrc` and `@types/node` majors, held at
     _dependency-dashboard approval_ rather than opened as a PR. The supported floor is set
     by ADR-0001 and is stated in three places at once; a bot may raise the subject, not the
     floor, and never in one place out of three.

3. **A three-day cooldown applies to everything**, waived for vulnerability fixes. The
   window in which a compromised or yanked release is caught is the first days after it is
   published; automerge without a cooldown lands one within minutes of it existing. The
   asymmetry is deliberate, because a known CVE is a worse risk than a bad release.

4. **GitHub Actions are pinned to commit digests** (`helpers:pinGitHubActionDigests`), and
   the digests are then automerged. An action tag is mutable code running with the
   workflow's identity, and in `publish.yml` that identity is one npm will exchange for a
   token that can publish this package (ADR-0026). Pinning is only sustainable because the
   same bot keeps the pins current: pinned-and-stale is worse than unpinned.

5. **Alerts come from OSV directly** (`osvVulnerabilityAlerts`), not from GitHub's
   Dependabot alerts, which are a repository setting, invisible in a diff and silently
   switchable, the same objection ADR-0026 raised against click-ops environments.

## Consequences

- **Positive:** the tree stays current by default rather than by remembering; the failure
  mode of an update is a red check on a pull request instead of a discovery two years
  later; the actions that run with publish rights are immutable references; and the
  decisions that are genuinely ours (the runtime floor, what consumers install, what the
  oracle is) are the short list a human actually reads.
- **Negative / deferred:** automerge makes the gate set load-bearing in a new way. A gap in
  coverage that a human reviewer might have caught by reading a changelog now merges
  silently, so the honest reading of this ADR is that it raises the cost of a weak test
  rather than lowering the cost of a dependency. It also grants a bot write access to
  `master`; the blast radius is bounded by ADR-0026, since publishing needs a `v*` tag and a
  published release, neither of which Renovate can supply. That bound got tighter to state
  and thinner in fact on 2026-08-05, when the environment's required reviewer was removed:
  a merge to `master` still publishes nothing, but nothing downstream of a tag asks a human
  any more.
- **Watch for:** NuGet lockfile updates. `tools/ooxml-validator` restores in locked mode, so
  a `.csproj` bump whose `packages.lock.json` was not regenerated fails the build. That
  failure is loud and lands on a PR nobody has merged, which is the acceptable version of
  this going wrong; if it recurs, the rule is to disable the `nuget` manager rather than to
  relax locked mode.
- **Revisit when:** the automerged set produces a regression that reached `master`, or the
  review queue for the stop-and-wait classes stops being drained. Either one means the line
  is drawn in the wrong place.
