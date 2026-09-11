# ADR 0015: Package name, SemVer, and the first published version

**Status:** Accepted (2026-07-21) · Phase 4 publishing-readiness slice

## Context

ADR-0001's addendum flagged one open item from the Phase 4 build slice: "the
definitive package **name** remains the one human decision deferred to the rebrand
slice." `package.json` has said `@shbernal/ts-xlsx` since that slice, and
`prepublishOnly` (build, full test, `smoke:dist`, `size`) already gates a real
publish. But nothing had confirmed the name was _final_ rather than a placeholder,
and nothing had stated a versioning policy for the version field, which has sat at
the placeholder `0.0.0-dev` throughout the rewrite. The library is not being
published yet; this ADR settles the policy so it is ready the moment it is.

## Decision

1. **The package name is final: `@shbernal/ts-xlsx`.** This closes the human
   decision ADR-0001 deferred. No further rebrand slice is planned.

2. **The project follows [SemVer](https://semver.org) from the first publish
   onward.** Once published, a breaking change to the public API (anything the
   generated [`docs/api/`](../api/README.md) reference or the barrel `src/index.ts`
   exports) requires a major bump; additive, backward-compatible API is a minor
   bump; a fix with no API change is a patch. This is a natural fit for
   CLAUDE.md §1's "breaking changes are welcome, not tolerated" stance, because SemVer
   is how that stance stays honest to consumers instead of surprising them.

3. **The first published version is `1.0.0`, not a `0.x` series.** SemVer's own
   spec (semver.org #4) treats `0.x` as "anything may change at any time" and
   `1.0.0` as the point a public API is declared stable. This library does not have
   an unstable-exploration phase to signal. Everything landed already clears
   the CLAUDE.md §2 bar (strict types, lint-clean, unit plus regression corpus tested,
   docs generated from types) before it merges to `master`. An extended `0.x`
   period would understate that bar, not reflect it. `1.0.0` is therefore the
   correct first tag, not an aspirational one to grow into.

4. **`package.json`'s `version` field stays `0.0.0-dev` until the actual publish.**
   It is a placeholder that signals "not yet released", mirrored by the absence of
   a `CHANGELOG.md` `## [1.0.0]` heading (see below). The first `npm publish` is
   the moment that changes it to `1.0.0` directly; nothing bumps it before then.
   Publishing itself remains a separate, deliberate, human-triggered action, and this
   ADR does not authorize it.

5. **`CHANGELOG.md` follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)**,
   with an `## [Unreleased]` section that accumulates entries as work lands, cut
   into `## [1.0.0]` at publish time. History prior to this ADR is not backfilled
   into it. That history is fully available in `git log`, and the ADR series is the
   durable record of _why_; the changelog exists to tell a consumer _what changed
   between releases_, a concern that starts now, not retroactively.

## Consequences

- **Positive:** the two open human decisions from ADR-0001 (name, and now version
  policy) are both closed; a future publish is a mechanical `npm publish` after
  flipping `version` and cutting the changelog section, not a decision point;
  consumers get standard, predictable SemVer guarantees from day one with no `0.x`
  churn to live through first.
- **Negative / deferred:** none. This is policy, not a behavior change, and
  publishing itself is still not happening yet.
- **Revisit when:** the project is actually ready to publish (at which point this
  ADR's decision 4 is executed, not reopened) or a concrete need for a pre-1.0
  `0.x` signal emerges (none is anticipated).

---

## Addendum (2026-07-29): decision 4 executed, and the tag namespace reclaimed

The project released. Decision 4 was executed rather than reopened: `package.json`'s
`version` went from `0.0.0-dev` straight to `1.0.0`, `CHANGELOG.md`'s `## [Unreleased]`
section was cut into `## [1.0.0]`, and the release was published on GitHub. **npm publish
did not happen in the same step.** The GitHub release is the artifact today, and
`npm publish` remains the separate, human-triggered action this ADR always said it was,
because it needs credentials and a `@shbernal` scope that only the human can grant.

One thing this ADR did not anticipate: the fork inherited ExcelJS's ~190 git tags
(`v0.0.1` … `v4.4.1`) on both the local clone and `origin`, so the `v1.0.0` name decision 3
chose was **already taken**, by an ExcelJS commit that is not in our history. The
inherited tags were deleted from `origin` and locally rather than versioning around them.
Deleting is the choice consistent with CLAUDE.md §1's clean break: those tags mark releases
of a different library, they collide with every version we will ship between 1.0.0 and
4.4.1, and keeping them would leave the tag list unable to say whose release a `v2.0.0`
names. Nothing is lost, because `exceljs/exceljs` still holds every one of them, and
`git fetch upstream --tags` restores them if we ever want them back.
