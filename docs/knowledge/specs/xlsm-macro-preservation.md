# Preserve VBA on round-trip of macro-enabled workbooks (.xlsm)

Cluster: formats

## Scenario

A user maintains a macro-enabled spreadsheet (.xlsm) with VBA automation, for example a report
template whose embedded macro exports to PDF. They want to open it, edit worksheet data
programmatically, and save it back with the macros intact so it still runs when reopened in Excel.
Today, reading and rewriting such a workbook silently strips the VBA project, yielding a plain .xlsx
that has lost its macros. The related, explicitly out-of-scope wish is to _invoke_ macros from
Node, which needs a live Excel or automation host and is not a document-tool concern.

> Spec note, not a corpus case: capturing this needs a real .xlsm fixture and a feature that does not
> yet exist, so there is no assertable behavior to lock today. The durable value is the preservation
> contract and its packaging details.

## Desired behavior

- On reading an .xlsm, the library **retains the VBA project part** (`vbaProject.bin`) and its
  associated relationships and content-type declarations for **byte-faithful round-trip**, so the
  preserved bytes stay the sole thing the writer emits.
- On writing, if the source carried a VBA project, or the caller explicitly requests macro-enabled
  output, the package is emitted with the correct macro-enabled content type and the
  `vbaProject.bin` part **re-embedded byte-for-byte**, so a round-trip that only edits cell values
  yields a file Excel still treats as macro-enabled with functioning macros.
- Callers can **read** the macro source through a lazy, read-only typed view
  (`Workbook.vbaProject`), derived from the preserved bytes without a write-back path; see
  ADR 0016. Reading never perturbs what the writer emits.
- Callers can **attach, replace, copy, or remove** a macro project at the bytes level through
  `Workbook.vbaProjectBytes`, a get/set accessor pair over the raw `vbaProject.bin`; see ADR 0017.
  Setting bytes makes the written package macro-enabled and embeds them verbatim, validated
  fail-closed so a malformed blob is rejected and never half-applied, and setting `undefined` demotes
  the workbook to a plain package. Replacing or removing drops a now-stale signature over the old
  bytes.
- Callers can **author or edit a macro project's module source**, but this needs genuinely compiled,
  source-matched p-code that **only a real Excel can produce**, so it is done by the offline
  `tools/vba-compiler` build tool via VBIDE, not in the shipped library; see ADR 0019. It emits either
  a `vbaProject.bin`, attached via `Workbook.vbaProjectBytes`, or a whole edited `.xlsm` in place, for
  document and designer code-behind. **The earlier pure-TS "author/edit from source" surface
  (`writeVbaProject`/`setVbaProject`, `editVbaModuleSources`/`setVbaModuleSource`/`editXlsxVbaModuleSource`,
  `addVbaModule`) has been removed:** it emitted modules with no or zeroed p-code and reset
  `_VBA_PROJECT` to an "unmatchable version cookie" on the theory that Excel recompiles from source on
  open. Excel does **not**, since a module runs the compiled p-code it ships, so those files either
  threw "Invalid data format" or silently ran stale code (ADR 0019 retracts ADRs 0017 §2.3c and 0018).
  The prior "Verified against real Excel" claims never exercised the compile and load step, because
  macros were force-disabled or Enable Content was never clicked.
- Callers can still apply **pure-TS structural edits** that never touch a module's p-code:
  `removeVbaModule` and `addVbaReference`, with `Workbook.removeVbaModule` and `addVbaReference` at
  model level and `editXlsxVbaRemoveModule` and `editXlsxVbaAddReference` at package level, swapping
  only `xl/vbaProject.bin` and leaving every other part byte-for-byte. These splice the original
  `.bin`, editing only the `dir` stream and, for a removal, `PROJECT` and `PROJECTwm`, and leaving
  every module stream and `_VBA_PROJECT` untouched, because the `dir` stream is authoritative for the
  module and reference list. Verified with `execute-verdict.ps1`, which opens with macros enabled and
  runs a surviving macro.
- A signed VBA project's `vbaProjectSignature` part is preserved alongside. It is a sibling part
  reached from `xl/_rels/vbaProject.bin.rels`, so the closure walk carries it through, and its
  distinct `.bin` content type is re-declared per-part rather than collapsed into the `vbaProject`
  default, which a single extension `<Default>` would otherwise do, locked by `preserved-parts.test.ts`.
  Editing the _project itself_ legitimately invalidates the signature; editing unrelated cells does
  not (see open questions). Every project-mutating surface drops the now-stale signature.
- **Out of scope:** executing or running macros, which needs a live host (ADR 0013), and authoring or
  editing module _source_ inside the shipped library, which needs real compiled p-code and is done by
  the offline `tools/vba-compiler` instead (ADR 0019). In scope in pure TS: attaching and replacing a
  whole `.bin` via `Workbook.vbaProjectBytes`, and structural edits that don't touch p-code, removing a
  module and adding a reference (ADRs 0018/0019). The library is a document tool, not a VBA interpreter.

## Prior art

OOXML defines a distinct macro-enabled content type for the workbook part and packages the VBA as a
binary `vbaProject.bin` referenced by a workbook relationship; the ZIP container is otherwise
identical to .xlsx. Many OOXML libraries treat the VBA project as an opaque blob they copy through
without parsing. Macro-enabled templates (.xltm) are the template analog.

## Open questions

- API surface: automatic preservation whenever a VBA part is detected on read, or gated behind an
  explicit flag when writing?
- Output-format choice (.xlsx against .xlsm): infer from source, from filename extension, or an
  explicit option?
- ~~Expose the VBA project bytes to callers, or only pass them through opaquely?~~ **Decided
  (ADR 0016):** exposed as a lazy, read-only typed view (`Workbook.vbaProject`), derived from the
  preserved bytes with no write-back path.
- ~~Let callers author macros, or only preserve what a read produced?~~ **Decided (ADRs 0017–0019):**
  authoring is supported, but **source authoring and editing need a real Excel**, since only it emits
  runnable p-code, so it lives in the offline `tools/vba-compiler`. The pure-TS from-source surface
  tried in ADRs 0017/0018 was retracted (ADR 0019) because Excel does not recompile from source on
  open. In the shipped library, attach-blob (`Workbook.vbaProjectBytes`) and pure-TS structural edits
  (`removeVbaModule`, `addVbaReference`, and their wrappers) are shipped, each fail-closed.
- **Signature on a project edit is dropped**: editing the project invalidates any signature over
  it, so every project-mutating surface discards the stale `vbaProjectSignature` (part, relationship,
  content-type override) rather than leave it advertising a broken signature. Editing unrelated cells
  still preserves it, and the part round-trips correctly typed. ~~An `isSigned` accessor is still
  sourceable from the preserved closure once a consumer needs it.~~ **Built (ADR 0021):**
  `Workbook.vbaProjectSigned` (boolean) and `Workbook.vbaProjectSignatures` (raw bytes plus generation)
  read the presence of a signature over the preserved closure, a **read** with no cryptographic
  verification, so the drop behavior above is now observable in-memory, since a replaced project reads
  unsigned. Detection keys off the relationship Type's final segment, so all three generations (legacy
  2006, agile 2014, V3 2020) are recognised regardless of the year the URI carries, and the closure walk
  already preserves each part. CMS and PKCS#7 parsing, cert-chain validation, and signer identity stay
  out of scope.
- ~~Parse macro/toolbar-referenced `customUI`, or leave it opaque?~~ **Audited and fixed:** the ribbon
  parts (`customUI/customUI.xml`, `customUI14.xml`) hang off the _package root_ `_rels/.rels`, not the
  workbook rels, so the workbook-closure net never reached them and the writer, which regenerates the
  root rels from the model, dropped them. Now captured as package-root preserved references and
  re-declared verbatim on write, locked by `preserved-parts.test.ts`. **A typed read view now sits on
  top** (ADR 0020): `Workbook.customUI` parses the `<ribbon>` subtree (tabs, groups, controls and
  callback names) of each part, lazily and memoised, keying dialect off the `<customUI>` namespace. It
  is strictly additive, since preservation stays the sole emission authority and the bytes still
  round-trip opaquely, and the reader just projects a view over them exactly as `Workbook.vbaProject`
  does. Backstage, QAT, contextualTabs, commands, and ribbon _authoring_ stay deferred for want of a
  consumer. Fixing this surfaced a latent wrong fact: the round-trip fixture had used the customUI14
  _namespace_ as its relationship _type_, corrected to the real
  `.../office/2007/relationships/ui/extensibility`.

Related: `roundtrip-preserves-unmodeled-package-parts`; ADR 0016 (VBA read view, authoring deferred),
ADR 0017 (VBA authoring in scope; attach-blob plus from-scratch synthesis), ADR 0018 (editing an existing
module's source by splice), ADR 0020 (customUI ribbon read view; authoring deferred).
