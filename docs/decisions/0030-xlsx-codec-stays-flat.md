# ADR 0030: `src/io/xlsx/` stays flat; the read/write directory split is rejected

**Status:** Accepted (2026-08-08) · module layout · refines the layering rules of [ADR 0029](./0029-toolchain-standup.md)-era tooling (`scripts/check-layering.ts`)

## Context

`src/io/xlsx/` holds ~32 non-test modules in one directory, the largest flat directory in the tree.
It also spelled the read/write split three different ways: `styles.ts`/`read-styles.ts` (prefix),
`pivot.ts`/`pivot-read.ts` (suffix), `shared-strings.ts`/`shared-strings-read.ts` (suffix), while
seven other features handled both directions in a single unprefixed file. That reads like drift, and
the obvious remedy is `xlsx/read/` plus `xlsx/write/` plus `xlsx/shared/`, with a `check-layering.ts`
rule forbidding a read module from importing a write module. The rule would turn a naming convention
into a mechanism, which is what this project prefers (`check-layering.ts` says so in its own header).

We classified every module by reachability from the two entry pairs (`read.ts`/`read-rows.ts` and
`write.ts`/`write-stream.ts`) instead of by its name. The result: 10 read-only, 9 write-only, and
**13 reachable from both**.

## Decision

**The directory stays flat.** The three-way split is rejected. Two findings drove it:

1. **The "shared" bucket is mostly not shared infrastructure. It is seven feature modules that
   carry both directions.** `comments.ts`, `tables.ts`, `images.ts`, `hyperlinks.ts`,
   `data-validation.ts`, `conditional-formatting.ts` and `threaded-comments.ts` each export a
   `parseX` beside an `xXml`. That co-location is load-bearing: the two halves share one feature's
   element vocabulary and must agree with each other, and a round-trip case is the assertion that
   they do. A read/write split would cut each in half and file the two functions that have to stay
   in step under different directories, a strictly worse arrangement reached in the name of tidiness.

2. **The directory names would lie.** `conditional-formatting.ts` imports the `StyleRegistry` from
   `styles.ts`, and `read.ts` imports `conditional-formatting.ts`. So the read pipeline depends on
   the write-side registry transitively no matter which directory anything sits in. Directories named
   `read/` and `write/` would advertise a separation the graph does not have.

**The invariant we do keep** is narrower and true: no read-pipeline module imports a write-pipeline
module _directly_. Making that hold required extracting `parseColor` and `colorAttrs` into
`color-xml.ts`, because reading and writing `<color>` are one concern with two directions, and
`parseColor` had been sitting in the write-side style table, which made `read-styles.ts` and
`read-worksheet.ts` import the writer to decode a colour. That was the only such import in the codec.

**We do not gate that invariant**, and this is the part most likely to be revisited, so here is why:
a gate for it cannot be written honestly today.

- `check-layering.ts` matches `forbidden` entries as directory prefixes (`target.startsWith(layer + '/')`),
  so a file-level rule silently never fires. Its `RULES.find` also stops at the first matching rule,
  so per-file rules cannot stack.
- A rule that _derives_ the write pipeline from reachability defeats itself. The moment a read module
  imports a write module, that module becomes reachable from the read roots and therefore classifies
  as **shared**, which is exactly the label that makes the check pass. Verified, not assumed.
- A declared list of write-pipeline modules would work and would go stale in silence, since a new
  write module simply would not be on it. Silent staleness is the failure mode these checkers exist
  to prevent; adding one that has it would cost more credibility than the rule is worth.

So the convention is documented in `docs/architecture.md`, which now describes the three kinds of
module and why the tidy-up is wrong, and enforced by review.

**Enforced by review means enforced by nobody in particular, and it has already slipped once.**
`rich-runs.ts` was read-only and unprefixed for as long as it existed, and is `read-rich-runs.ts`
since 2026-09-04. The reason a checker is still the wrong answer is above and unchanged; the reason
to say this here is that a convention a reader assumes is gated is a convention they stop checking.

## Consequences

- **Good:** the seven bidirectional feature modules stay whole. The one real read-to-write dependency
  is gone. The naming is uniform: `pivot-read.ts` became `read-pivot.ts` and
  `shared-strings-read.ts` became `read-shared-strings.ts`, so `read-` is the prefix everywhere.
- **Good:** `<color>` serialisation and parsing now sit together, where the round-trip agreement they
  owe each other is visible in one file.
- **Bad:** the read/write direction rests on a documented convention, not a gate, which is the weaker
  of the two kinds of guarantee this project trusts. A reader borrowing a helper out of a write module
  would pass CI. `docs/architecture.md` names `color-xml.ts` as the precedent for where such a helper
  goes.
- **Neutral:** the directory stays large. Size alone was never the problem; the mixed-concern module
  was, and that one is fixed.

## Revisit if

- `check-layering.ts` grows file-level `forbidden` matching _and_ multi-rule stacking, at which point
  the direction rule becomes expressible without a stale-prone list. That is the cheapest path to
  gating this.
- A feature module's two halves stop sharing anything but a filename, which would make splitting that
  module (not the directory) correct.
