# Declaration of Principles

> This file is the constitution of the project. Every agent — human or AI — reads
> it before touching the code. `AGENTS.md` is a symlink to this file so that all
> runtimes converge on the same rules. If a decision is not covered here, decide
> in the _spirit_ of this document, act, and record what you did.

---

## 0. What this is

This is a **hard fork** of ExcelJS with the explicit goal of becoming a **fully
independent, modern, TypeScript-first library** for reading and writing
spreadsheet (OOXML / `.xlsx`, and adjacent) documents.

We forked because upstream `exceljs/exceljs` is effectively unmaintained: no merge
to `master` since January 2024, no release since v4.4.0 (October 2023), ~650 open
issues and ~140 open PRs rotting on the vine — while the package still serves tens
of millions of downloads a month. There is enormous accumulated value trapped in
that backlog and enormous accumulated debt in that codebase. We intend to **extract
the value and discard the debt.**

We are not a compatibility fork. We do not exist to keep other people's imports
working. We exist to be the _best_ spreadsheet library for a world where most code
is written and maintained by autonomous agents alongside humans.

---

## 1. Non-negotiable stances

1. **This is a clean break.** We aim for full independence from upstream. The harvest is
   complete (see `docs/architecture.md`) and we no longer track upstream.

2. **No backwards-compatibility guarantee.** Ever. The old API is a _reference_,
   not a contract. If a better shape exists, we take it.

3. **Breaking changes are welcome, not tolerated.** A breaking change that makes
   the library clearer, safer, or faster is a _good day_. We version honestly
   (SemVer major bumps are cheap and expected during the rebuild) and document the
   break — but we never contort the design to avoid one.

4. **Nothing is sacred.** No comment, no naming convention, no file layout, no
   abstraction survives on the grounds that "it was already there." Legacy status
   is not a reason to keep something. It is, if anything, a reason to re-examine it.

5. **We do not keep legacy code.** Code is either _useful and modern_ or it is
   _deleted_. There is no "leave it, it works" tier. If it works but is ugly,
   untyped, untested, or unclear — it is not done, it is debt, and debt gets paid
   or removed.

6. **AI-first, human-minimal.** This project is optimized so that autonomous agents
   can do the overwhelming majority of the work with high confidence and low human
   supervision. The human's time is the scarcest resource; spend it only where
   judgment is genuinely irreducible. **Do not clobber the human with questions.**
   Make the reasonable call, act, and note the assumption so it can be reverted.

---

## 2. Quality is the moat

Because humans intervene rarely, **the machine-checkable safety net is the primary
guarantor of correctness.** We do not trust vibes; we trust green checks. The bar
is deliberately high and non-negotiable.

Every change must satisfy, with **zero** exceptions merged to the main branch:

- **Strict TypeScript.** `strict: true` plus `noUncheckedIndexedAccess`,
  `exactOptionalPropertyTypes`, `noImplicitOverride`, and friends. `any` is a code
  smell that must be justified in a comment or eliminated. Prefer `unknown` +
  narrowing. Public API is fully and precisely typed — the types _are_ the docs.
- **Linting & formatting clean.** oxlint for rules, oxfmt for layout — two tools
  from one toolchain, neither of them doing the other's job. Type-aware rules come
  from tsgolint, so the linter sees what the typechecker sees. No warnings.
  Formatting is never debated; it is automated. A rule this project will not adopt
  is turned off _in the config, with its finding count and the reason_ — a decline
  on the record is a decision; a rule left unmentioned is an accident.
- **Unit tests.** Every behavior is covered. New behavior ships with tests in the
  same change. Bugs are fixed _test-first_: a failing reproduction, then the fix.
- **Regression corpus.** Every real-world file/edge case we learn about (especially
  from the harvested upstream issues) becomes a permanent fixture-backed test. We
  never fix the same bug twice.
- **Type-level tests** for the public API surface (`expectTypeOf` / `tsd`-style)
  so refactors can't silently degrade the developer experience.
- **Static analysis & supply-chain hygiene.** Dependency audit is part of CI and is
  _expected to stay green_ — a large part of why we forked was upstream's rotting
  transitive dependencies. We keep the dependency tree small, modern, and clean.
- **Coverage is watched, not gamed.** Coverage thresholds are a floor, not a
  target; the goal is _meaningful_ coverage of behavior, not line percentage.

**Definition of Done:** typed, linted, tested (unit + regression where relevant),
documented at the API surface, dependency-clean, and green in CI. Anything short of
that is not done — it is in progress.

---

## 3. How agents work here

- **Bias to action.** When you have enough information to make a sound decision,
  make it and proceed. Do not ask permission for things this document already
  authorizes (deleting legacy, breaking APIs, modernizing tooling, rewriting bad
  comments).
- **Small, coherent, reviewable changes.** Each change does one thing, is fully
  green, and leaves the tree in a better state than it found it. No half-migrations
  landed on main.
- **Leave a trail, not a mess.** Record non-obvious decisions and assumptions in the
  relevant doc (`docs/architecture.md`, ADRs under `docs/decisions/`, or the change
  description) so the next agent — or the human — can audit and reverse them.
- **Verify, don't assume.** Prefer reproductions and tests over reasoning about what
  "should" work. The OOXML format is full of surprises; the corpus is how we tame
  them. `docs/agent-correctness-playbook.md` is the dispatch table — _what you are
  doing_ → _the check that proves it correct_ → _the exact command_ — so you never
  have to rebuild that decision tree.
- **Scratch goes in `.tmp/`.** Throwaway probes, dumps, generated workbooks, screenshots —
  everything regenerable — lives in the repo's `.tmp/` (also exported as `$SCRATCH`, and as
  `$TMPDIR` for shells that would otherwise find it unset). Never the repo root, never an
  absolute system temp path: in-repo keeps it inspectable and already git-ignored, and
  relative keeps the command short enough to reuse. A probe earns its place only while its
  question is open: once the answer is locked somewhere durable (a test, a corpus case, a
  fixture's recorded verdict, an ADR), delete the probe in the same change that lands the
  answer. Scratch that outlives the work it served is not a record, it is a haystack, and a
  `.tmp/` nobody can read is no longer inspectable, which was the whole reason it lives here.
- **Commit messages go through a file, never through a shell.** Write the message with
  your file-writing tool, then `git commit -F <file>`. Only a one-line
  `git commit -m "subject"` may be typed inline. This is not style: an agent on Windows
  has both a POSIX shell and PowerShell available, and the two disagree on multi-line
  quoting (`<<'EOF'` against the PowerShell here-string). The wrong dialect does not
  error — the delimiter is passed through as text and lands in the message. Two commits
  in this history open with a bare `@` for exactly that reason. Routing the message
  through a file removes the shell from the path, so there is no dialect left to get
  wrong. The `no-shell-quoting-leak` rule in `lefthook.yml` is the backstop, not the
  place to learn this: when it fires you have already burned the commit attempt.
- **When you must ask, batch it and make it count.** Only escalate to the human for
  decisions that are genuinely theirs: irreversible/outward-facing actions
  (publishing, naming/branding, license), or a true fork-in-the-road where the
  options are strategically divergent and you cannot pick a clearly-better one.
- **Security- and correctness-first, always.** A spreadsheet library parses
  untrusted input. Treat every parser path as hostile-input-facing. No unbounded
  allocation, no zip-bomb naïveté, no eval-shaped surprises.

## 4. Code & comment standards

- **Comments explain _why_, never _what_.** The code says what. If a comment
  restates the code, delete it. If the code needs a comment to be understood,
  first try to make the code not need it.
- **Delete dead code and stale comments on sight.** Do not preserve them "just in
  case." Git remembers.
- **Names carry meaning.** Precise, honest names over short or clever ones. The
  public API is a UX surface for both humans and agents — optimize for
  discoverability and for correct use being the easy path.
- **No cargo cult.** Every dependency, config line, build step, and abstraction
  earns its place or is removed. Simplicity is a feature.
- **Modern idioms only.** ESM, `async`/`await`, immutable-by-default, narrow
  interfaces. No callbacks-as-API, no `var`, no CommonJS-isms in source.

---

## 5. The one rule that overrides convenience

> If keeping something would make the library worse but easier, and removing it
> would make the library better but harder — **we do the harder, better thing.**

That is the whole point of the fork.

---

_See `docs/architecture.md` for how these principles are realized — the corpus contract,
the module layout, and the working agreements; `docs/agent-correctness-playbook.md` for
the situation → check → command dispatch table; and `docs/decisions/` for the ADRs._
