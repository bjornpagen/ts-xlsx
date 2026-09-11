# ADR 0026: Releasing is a GitHub release, and npm follows with no credential

**Status:** Accepted (2026-07-29) · release slice, immediately after 1.0.0

## Context

ADR-0015 settled _what_ a version is called and left publishing as "a separate,
deliberate, human-triggered action". 1.0.0 was then published by hand: `npm publish`
from a maintainer's terminal, with a one-time password typed at the prompt. That works
exactly once per person who has the password, and it produces a package the registry
cannot vouch for, because npm serves the bytes and can say nothing about which commit or
which tree produced them. It also means every release is a chance to publish a dirty tree, a
wrong version, or a build nobody gated.

The obvious fix, a workflow holding an npm automation token, trades one problem for
another. A long-lived credential in repository settings is exfiltratable by any workflow
change that can read secrets, and it is invisible: nothing in a diff shows who can publish.

## Decision

1. **Publishing a GitHub release publishes to npm.** `.github/workflows/publish.yml`
   triggers on `release: published`. Cutting a release is therefore the whole ritual;
   there is no second command to remember and no terminal that must be someone's.

2. **No npm credential exists in this repository.** Authentication is npm **trusted
   publishing**: the job proves its identity with a GitHub OIDC token (`id-token: write`),
   and npm exchanges it for a short-lived one _only_ when the repository, workflow
   filename and environment match the trusted publisher configured on the package. The
   same exchange emits a provenance attestation, so the registry can state which commit
   and which run built what it serves. There is nothing to leak and nothing to rotate.

3. **The publish job runs in the `npm-publish` environment, and that environment is
   code.** `.github/workflows/environment.yml` declares it, allowing deployments only from
   `v*` **tags** and no required reviewer, so the gate guarding a publish is reviewable in a
   diff instead of being a settings page nobody reads, and the workflow fails when the live
   environment disagrees with the declaration.

   It _provisioned_ the environment through 1.1.0 and now only checks it (**amended
   2026-08-08**). Writing an environment is an administration endpoint, and `administration`
   is not among the permissions a workflow may request for `GITHUB_TOKEN`, so provisioning
   needed a PAT, `ENV_ADMIN_TOKEN`, which was never created; the workflow spent every day
   of its life failing on its own first step. Reading an environment needs nothing but read
   access to the repository, so the check runs on the built-in token. See the consequences
   for what detection buys that provisioning did not.

   The environment carried a required reviewer through 1.0.3 and no longer does
   (**amended 2026-08-05**). On a single-maintainer project it was never a second pair of
   eyes: the only login that could approve was the one that had just cut the release, so it
   asked the releaser to confirm a decision they had made one command earlier. Consent that
   only the decider can give is a delay, not a control. What it genuinely stopped was an
   _accidental_ publish; see the consequences below for what now carries that.

4. **The workflow enumerates no gates.** `npm publish` runs `prepublishOnly`, which is
   build, `verify --full`, `smoke:dist` and the size budgets. That is the single definition of
   "everything" that ADR-0022 established and corpus.yml already refuses to duplicate.

5. **Three refusals are wired in**, because a release is the one action that cannot be
   taken back: a tag whose name disagrees with the version its own commit declares fails
   the run; a version already on the registry is a no-op rather than a red X, so re-running
   a workflow is safe; and a pre-release tag takes the `next` dist-tag, never `latest`.

## Consequences

- **Positive:** a release is one action and nothing else; the published package carries
  provenance; no credential exists to steal; and who may publish is a file, not a memory.
- **The blast radius of a `v*` tag is now the whole gate.** With no reviewer, pushing a tag
  and publishing a release is the entire path to the registry, so anything holding push
  access (a leaked token, an agent given more rope than intended, a mistyped tag on the
  wrong commit) publishes unattended, and a version number cannot be taken back. Two
  things still stand between a mistake and npm: the tag-name/version agreement check, which
  refuses a tag whose commit declares a different version, and `prepublishOnly`, which
  refuses a tree that fails its gates. Neither catches a _correct_ publish of something
  nobody meant to release. Rehearse (`dry_run`) when a release is in any way unusual;
  that habit is the replacement for the click, and unlike the click it can actually fail.
- **Negative / deferred:** the trusted publisher must be configured once on npmjs.com
  before CI can publish at all, and it pins three names that live outside this repository's
  review: the repository, the workflow _filename_, and the _environment_. Renaming
  `publish.yml` or the `npm-publish` environment breaks publishing until npm is updated to
  match, and npm reports any of the three disagreeing identically.

  **The environment is now checked rather than provisioned, and the repository holds no
  credential at all** (2026-08-08). `ENV_ADMIN_TOKEN` was never created, so the reviewer
  removal of 2026-08-05 was applied by hand and the file and the live environment agreed by
  coincidence rather than by construction, which is the one property this ADR had bought.
  Restoring the secret would have bought it back at the price of the _other_ property, since
  a PAT with Administration rights is a long-lived credential sitting beside a publish path
  that deliberately has none, and it is the strictly more dangerous of the two: it can rewrite
  the gate that decides who may publish. So the write was dropped instead. `environment.yml`
  now reads the environment on the built-in token and fails when it differs from what the file
  declares, on a weekly schedule, on any edit to the declaration, and on demand.

  What that gives up is one-click reconciliation: a drift is now fixed by hand under Settings
  → Environments. What it keeps is the half that was load-bearing, since the failure this
  guards against is a settings edit nobody sees, and a check that goes red names it. What it
  gains is a check that can actually run, which the provisioning version never could.

  Proving it could fail found a defect in it, which is the argument for the practice. Run
  against a throwaway environment configured wrong on purpose, it went red, with
  `gh: Not Found (HTTP 404)` and none of its own output. GitHub serves the
  deployment-branch-policies endpoint only while custom policies are switched _on_; with them
  off it 404s rather than returning an empty list. So the drift that matters most, the tag
  rule removed, was the one case reported as an unattributed HTTP error. The switch is now
  read first and "off" is spelled as the empty list it means.

- **Verified, not assumed.** The verification corrected the design twice and the
  npm-side configuration once, each against a mistake of ours.

  `npm publish --dry-run` is not local. It builds the tarball and then asks the registry,
  which refuses a version it already serves, so a rehearsal reaches the publish step only
  for a version that is not out yet.

  `setup-node`'s `registry-url` is **required**. It sets `NODE_AUTH_TOKEN` to the literal
  placeholder `XXXXX-XXXXX-XXXXX-XXXXX`, from which we inferred that npm would send that
  placeholder instead of exchanging an OIDC token, and dropped the option. That inference
  was wrong, and 1.0.1 failed `ENEEDAUTH` proving it: `registry-url` writes the .npmrc
  entry naming the registry this publish authenticates against, and without it npm does not
  begin the exchange at all. The placeholder is genuinely unused, because that variable is
  read when _installing_ private dependencies, not when publishing. The lesson generalises: a
  plausible reading of an observed value is not evidence, and the release path is one place
  where the difference is a burned version number.

  The publisher's **environment** field was then wrong on npm, and being wrong looked like
  nothing. It named `environment.yml`, the workflow that _provisions_ the deployment
  environment, where it had to name the environment itself, `npm-publish`. npm answers a
  rejected identity with a 404, which reads as "no such package" and says nothing about
  which of the three claims failed to match, so 1.0.2 failed against a workflow that was
  already correct and the two commits after its tag sharpened the error message rather than
  the cause. The general rule that falls out: a publish that 404s on a package that
  demonstrably exists is an identity mismatch, and the thing to read is the trusted
  publisher on npmjs.com, not this repository.

  The rehearsal had been unable to catch any of this, which is the part worth keeping.
  `npm publish --dry-run` treats a rejected identity as a warning and exits `0`, so the
  rehearsal passed _because_ it was a rehearsal, and the misconfiguration reached a real
  release twice. A check that cannot fail on the thing it is checking is decoration, and it
  reads as reassurance, which is worse than having no check at all. The job now reads the
  verbose log and fails when the exchange was rejected, or when npm never attempted one.
  Both markers were taken from the failing run's own output and re-run against it, rather
  than written from what the log was assumed to say, which is the same discipline the two
  corrections above had to be learned through.

- **Revisit when:** npm changes the trusted-publishing contract, or a second package ships
  from this repository (the environment and the publisher config are both single-package
  shaped today).
