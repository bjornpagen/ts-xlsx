# ADR 0012: Three tiers of correctness evidence; round-trip proves consistency, not conformance

**Status:** Accepted (2026-07-21) · Phase 4 · sharpens [ADR 0002](./0002-ooxml-validation-oracle.md) (validation oracle) and [ADR 0010](./0010-agent-correctness-dispatch.md) (correctness dispatch)

## Context

Most of the corpus writes a workbook with our writer and reads it back with our
reader: `roundtripWorkbook` and the bulk of the bespoke capabilities. ADR 0002
already noted the hazard in one line ("a writer and reader can agree on malformed
OOXML") and stood up the `OpenXmlValidator` oracle in response. The table-header bug
(`9579f152`) showed the hazard is sharper than that line implies, and that even the
oracle does not close it.

A table declared its column names only in `xl/tables/table1.xml` and left the header
row missing from the worksheet grid. Our reader takes column names from the table
part, so **write→read was a perfect fixed point** and every round-trip case stayed
green. The structural facts (`inspectPackage`) were green too, because they were
partitioned by package part and each part was internally valid. And the schema oracle
was green, because each part conforms to ECMA-376 in isolation. The rule Excel
enforces, that the header cells must exist and carry the column names, is a _cross-part
behavioral_ rule the written spec does not state. The only thing that rejected the
file was Excel Desktop.

The lesson is not "add one more fact." It is that a green corpus was systematically
misread as evidence of conformance when it was mostly evidence of self-consistency,
and that the gap between the two is where "files that do not open in Excel" lives.

## Decision

Treat correctness evidence as **three tiers of strictly increasing authority and
cost**, and never let a lower tier stand in for a claim only a higher tier can make.

### Tier 1: self-consistency (`R(W(x)) ≈ x`)

Write with our writer, read with our reader, assert the model survives. This is a
**fixed-point property of our own code.** It proves the encoder and decoder share a
model; it proves nothing about whether that model matches Excel.

- **Catches:** _unilateral_ bugs, where one half is malformed and the other half strict
  about it (a dropped field, an unparseable emission).
- **Structurally blind to:** _correlated_ bugs, where writer and reader are wrong in
  compensating directions or ignore the same field. These are not a rare tail: writer
  and reader for a feature are typically authored together from one mental model, so
  their blind spots are correlated _by construction_. The table bug is the canonical
  instance.
- **Cheapest, highest volume, weakest.** Sufficient for **intra-model** invariants (a
  value survives, a style does not bleed) and nothing more.

### Tier 2: spec-conformance (independent oracle)

Microsoft's `OpenXmlValidator` (ADR 0002), plus the structural `inspectPackage` facts.
An **independent implementation** breaks the write/read correlation. It is not our
code, so it cannot share our blind spot.

- **Catches:** everything the written standard (ECMA-376) states, meaning schema validity,
  OPC part structure, enums, child-element ordering, per-part semantic rules.
- **Ceiling:** it enforces what the spec _says_, not what Excel _does_. Excel's real
  acceptance rules are a **superset** of the spec, and many are cross-part invariants
  the spec never spells out. The table bug passed the oracle.
- Required for any **single-part conformance** claim; round-trip is not enough there.

### Tier 3: Excel behavior (the application itself)

What Excel Desktop actually does with the file. The **only** ground truth for Excel's
quirks and for cross-part invariants the spec omits. On a Windows+Excel host it is
**scriptable for state-observable behavior** via the Excel-oracle harness
([ADR 0013](./0013-excel-desktop-as-automatable-tier3-oracle.md)), covering auto-fill,
canonical re-save, and value/formula readback, though the interactive repair-dialog
experience is not. It is captured as provenance
(`{source: 'excel-desktop-verification'}`) either way.

- The table invariant was **discovered** here and could only be discovered here.
- Required for any **cross-part correspondence** or known-Excel-quirk claim.

### The rule

**Match the tier to the invariant, and record which tier witnessed it.**

| Invariant class                                            | Lowest sufficient tier                       | Example                                         |
| ---------------------------------------------------------- | -------------------------------------------- | ----------------------------------------------- |
| Intra-model (value/style survives a round-trip)            | Tier 1                                       | a number keeps its type; a font does not bleed  |
| Single-part conformance (well-formed, ordered, valid enum) | Tier 2                                       | `CT_Worksheet` child order; valid `patternType` |
| Cross-part correspondence / Excel quirk                    | Tier 3 (seed) plus a Tier-2 seam fact (lock) | table columns against header cells              |

A cross-part invariant is **seeded** by one Tier-3 verification (prove Excel enforces
it) and then **locked** by a corpus fact whose _shape is the relationship itself_, so
regressions are caught in CI without re-opening Excel. The header fix did exactly this:
one Excel-Desktop verification, then two new seam facts (`cellText`, `columnNames`)
that phrase the correspondence.

### Corollary: the corpus is partitioned by part, and that is a blind spot

`inspectPackage` describes each package part well and the relationships _between_ parts
barely. Any invariant spanning two parts is unstateable until a fact is deliberately
built to cross the seam. An audit of the current vocabulary (2026-07-21) found these
**unstated cross-part seams** Excel is known or expected to enforce:

- ~~`<mergeCell ref>` against covered cells, which must be empty (no `mergeCells` fact).~~
  **Closed** by the `mergeCleanReport` fact and
  `merged-range-opens-without-repair-prompt.case.ts`.
- `<tablePart r:id>` in a sheet → rel → table part → content-type override (tables carry
  no owning sheet or rel id).
- Drawing/image anchor → `r:embed` → drawing rel → media part → content-type (no
  drawing-XML parsing or media enumeration).
- `<dataValidation sqref>` against the cells it claims (no `dataValidation` fact).
- Hyperlink as external rel against internal location (no hyperlink fact).
- `<sheet r:id>` → rel → existing worksheet part (captured but only checked part to declaration).
- Cell `s=` index → `styles.xml` xf/font/fill (`cellText` drops `s`; only the default font
  is inspected).
- ~~Shared-formula master against slaves (`formulas` holds `<f>` text only, not `t="shared" si= ref=`).~~
  **Closed** by the `sharedFormulas` fact and
  `shared-formula-master-slave-geometry-structural.case.ts`, which lock the master/slave
  `si`/`ref` geometry structurally. The `ref` for a _non-contiguous_ clone set
  is also **closed**, verified benign against Excel (ADR 0013) and locked by
  `shared-formula-sparse-ref-matches-excel-canonical.case.ts` (see hazard below).
- Comment → VML shape → `legacyDrawing` rel.

**Structural ceiling:** `worksheetRels` reads only `sheet1.xml.rels`, and tables and drawings
are not tied to an owning sheet. Every _multi-sheet_ cross-part chain and every
"reference-in-sheet → rel → target part" resolution is unstateable until that is lifted.
It is the precondition for several rows above, so it is the highest-leverage fix.

These are not scheduled here; they are the backlog this ADR makes visible. Close each
the same way: Tier-3 seed, then a seam fact that locks it.

### Tier-3 hazard (closed, verdict benign, locked in CI): shared-formula `ref` over-covers a non-contiguous clone set

> **Seeded 2026-07-21 via the Excel-oracle harness ([ADR 0013](./0013-excel-desktop-as-automatable-tier3-oracle.md)); locked the same day.**
> **Verdict: BENIGN. The presumed fix direction was wrong.** Excel treats the shared-formula
> `ref` as a _bounding-box hint_, not an instruction to materialize the interior. On Excel 16.0
> build 20131 it opened the `ref="B1:D5"` package **without repair**, did **not** auto-fill the
> empty interior cells, and **re-saved a byte-structurally identical group** (same `ref="B1:D5"`,
> same `si="0"`, the same two clones). ts-xlsx's output _is_ Excel's own canonical form, so the
> two candidate fixes below (split into contiguous runs, or degrade clones) would make ts-xlsx
> **diverge from Excel**. Sidecar: `test/corpus/fixtures/excel-oracle/shared-formula-sparse-ref.json`.
> **Locked (Phase 4):** `shared-formula-sparse-ref-matches-excel-canonical.case.ts` is the Tier-2
> seam fact. It asserts the emitted geometry is one master with `ref="B1:D5"` and exactly the two
> authored clones as slaves (the interior is never materialized), matching Excel's canonical re-save,
> and runs in CI without re-opening Excel. The discovery narrative below is kept for context; the seed
> has answered its open question and the seam fact keeps it answered.

Discovered while closing the shared-formula seam (2026-07-21). `planSharedFormulas`
(`src/io/xlsx/worksheet-xml.ts`) computes a master's `ref` as the **bounding rectangle**
of the master plus all its clones. For the common fills, down a column or across a row,
that rectangle _is_ the clone set, and the emitted geometry is exact. But when the clones
are non-contiguous, the rectangle covers cells that were never cloned:

```
master B1, clones B2 + D5  →  <f t="shared" ref="B1:D5" si="0">…</f>
```

`B1:D5` is fifteen cells; only three (`B1`, `B2`, `D5`) carry a `<c>`. The other twelve are
absent from the sheet. The hazard is that the shared-formula `ref` is, to some consumers, an
_instruction to materialize the formula across the whole rectangle_: **LibreOffice** auto-fills
every cell in `ref` with the translated formula, so those twelve empty cells would silently
gain a formula the caller never wrote. This is the same shape as the merge-repair bug, where the
writer emits geometry a consuming app interprets more aggressively than we intended, which is
why it is called out rather than left implicit.

The two candidate fixes this framing implied were **split the group into maximal contiguous runs**
(each its own master, each needing the formula **translated** to its new anchor, which is
R1C1/relative-offset machinery we do not yet have) and **degrade non-fill clones to standalone
`<f>` cells**. Both are real work, and crucially the seed proved neither is _correct_: Excel keeps
the group whole and the over-wide `ref` intact, so either would make ts-xlsx diverge from the
application it must match. The right move was to seed the Tier-3 question before touching the
writer, and the seed's answer was "the writer is already right." That is why the resolution is a
seam fact, not a code change.

**How the seed was taken.** The tier is irreducible, since there is no local oracle for it (the
schema validator passed the analogous table bug), so the probe file (master `B1`, clones `B2` and
`D5`, `ref="B1:D5"`) was emitted, opened headless in Excel Desktop, and re-saved via the
Excel-oracle harness ([ADR 0013](./0013-excel-desktop-as-automatable-tier3-oracle.md)): probe
`tools/excel-oracle/probes/shared-formula-sparse-ref.json` produced sidecar
`test/corpus/fixtures/excel-oracle/shared-formula-sparse-ref.json`. Excel _ignored_ the empty interior
(no auto-fill) and re-saved the identical group, and the verdict box above records the result. The seam
fact `shared-formula-sparse-ref-matches-excel-canonical.case.ts` then locked that `ref`-to-clone-set
geometry so CI catches any regression without re-opening Excel.

## Consequences

- **Positive:** "the corpus is green" is no longer mistaken for "Excel accepts this."
  The playbook's defense-in-depth table now carries the epistemic distinction, not just
  a cost ordering. New cross-part invariants have a named recipe (seed plus seam fact) and a
  visible backlog.
- **Neutral:** cross-part correctness still costs a manual Excel verification to seed.
  That cost is real and intended, because Tier 3 is irreducible and the seam fact is what
  keeps it a one-time cost per invariant.
- **Does not change** the single-oracle stance (ADR 0010): Tier 2 remains exactly one
  independent oracle. This ADR adds no validator; it names what each existing tier can and
  cannot witness.
