# ADR 0018: Editing an existing macro's source is done by splicing the original `.bin`, not re-synthesizing it

**Status:** **Retracted (2026-07-24) by [ADR 0019](0019-vba-authoring-needs-real-pcode-recompile-cookie-retracted.md).**
The edit-source-by-splice mechanism (write source at `MODULEOFFSET 0`, zero the offset, reset
`_VBA_PROJECT` to the recompile cookie) does not work: Excel runs a module's compiled p-code, not its
source, so a spliced module either fails to load or silently runs stale code. Its "clean" GUI verdict
never clicked Enable Content. Editing an existing module's source now lives in the offline
`tools/vba-compiler` (in-place mode). Retained here for historical context. · Originally Accepted
(2026-07-23) · extended ADR 0017 with the _edit-existing_ case · amends nothing in ADR 0016's read view.

## Context

ADR 0017 §2.3 shipped authoring _from scratch_: `Workbook.setVbaProject({modules})` synthesizes a fresh
`vbaProject.bin` from a module list. By construction that project is **reference-free and host-default**.
`writeVbaProject` emits no `REFERENCE*` records and a fixed host block, and it **rejects document and
designer modules** because their host linkage (the `ThisWorkbook`/`Sheet1` code-behind wiring) cannot be
synthesized without the host.

That leaves a real, common need unmet: take an _existing_ `.xlsm`, change **one module's source**, often
a document code-behind like `ThisWorkbook`, and re-emit it with the project's **references, host-extender
info, and every other module preserved exactly**. From-scratch authoring cannot do this: rebuilding the
project from a source list would drop everything the source list does not model (references, host linkage,
untouched modules' p-code). The whole value here is _preservation_, and re-synthesis is preservation's
opposite.

## Decision

1. **Splice, don't re-synthesize.** `editVbaModuleSources(bin, edits)` (internal, in
   `src/vba/project-editor.ts`; `edits` is a `Map<moduleName, newSource>`) operates on the _original_
   `.bin` bytes and keeps every CFB stream verbatim except:
   - the edited module's source stream, replaced with `compressContainer(encode(newSource))` at
     MODULEOFFSET 0, no p-code;
   - the `dir` stream, decompressed, with **only** the edited modules' `MODULEOFFSET` records zeroed,
     then recompressed;
   - `_VBA_PROJECT`, replaced with the 7-byte recompile-from-source cookie (the same unmatchable header
     §2.3c uses), so Excel recompiles the edited modules on open.

   References, `PROJECTLibFlags`, constants, `PROJECT`/`PROJECTwm` text, host-extender info, and every
   _other_ module survive **because we never touch the streams that hold them**. Preservation by
   not-destroying is strictly higher fidelity than model-and-re-emit. Parse-first, fail-closed: a
   malformed original raises `VbaParseError` _before_ any mutation; an unknown module or an unrepresentable
   character raises `VbaAuthorError`.

2. **This amends nothing in the read view.** ADR 0016's `VbaProject` stays a source-only, read-only
   projection. The editor is a bytes-to-bytes transform that never models references or host info into the
   read view, so it cannot silently drop what it does not model, which is the failure mode that sank the
   re-synthesis alternative. That alternative was rejected because it would require modeling every
   [MS-OVBA] `REFERENCE*` and host record in the reader _and_ re-emitting it in the writer: more code and
   more break risk for strictly less fidelity.

3. **Document and designer modules are editable, which is the standout advantage.** Because host linkage
   already lives in the preserved `dir`/`PROJECT` streams, the splice _inherits_ it rather than
   synthesizing it. This is exactly the module class ADR 0017 §2.3c rejects for from-scratch authoring.
   Editing a `ThisWorkbook` code-behind in a real workbook is verified to open clean in Excel (below).

4. **Two public entry points, differing in fidelity.**
   - **`Workbook.setVbaModuleSource(name, source)`** (model level) reads `vbaProjectBytes`, splices, and
     routes the result back through the `vbaProjectBytes` **setter**, inheriting fail-closed validation,
     preserved-ref rebuild, signature-drop, and macro content-type with zero writer changes. It throws if
     no project is attached. The map-shaped primitive `editVbaModuleSources` is also barrel-exported for
     the batch/functional path.
   - **`editXlsxVbaModuleSource(xlsx, name, source)` / `editXlsxVbaModuleSources(xlsx, edits)`** (package
     level, `src/io/xlsx/edit-vba.ts`) unzip the package, locate `xl/vbaProject.bin` via the reader's
     OPC resolution (`_rels/.rels` → officeDocument → workbook `.rels` → `vbaProject` rel), splice, drop
     any stale signature (part, rels entry, and content-type override, every `vbaProjectSignature*`
     flavour), then re-zip. **Only `xl/vbaProject.bin` changes; every other part is byte-for-byte.**

5. **The package-level path is the highest-fidelity way to edit an existing `.xlsm`, and it exists because
   the model path can perturb strict parts.** The model round-trip (`readXlsx`, then
   `setVbaModuleSource`, then `writeXlsx`) re-serializes the _whole_ package from the parsed model, so it
   preserves only what the model captures. On a rich real-world workbook that round-trip perturbs parts
   Excel is strict about and Excel prompts to repair on open, and this is **not** the VBA edit's fault: a
   control round-trip with **no VBA edit at all** (`readXlsx` then `writeXlsx`) repairs identically. It is a
   **pre-existing, VBA-independent `writeXlsx` whole-package fidelity gap**, where our writer re-emits a
   simplified package and the minimal corpus round-trip passes precisely because it is minimal. The
   package-preserving path sidesteps it entirely by never re-authoring anything but the macro project.
   **Guidance:** to edit a macro in a real file whose non-macro content must survive exactly, use
   `editXlsxVbaModuleSource(s)`. `setVbaModuleSource` plus `writeXlsx` remains correct for model-built or
   minimal workbooks and until the general writer reaches whole-package parity. This is captured in the
   function JSDoc.

6. **Editing the project drops a stale signature.** A signature validates the _old_ project bytes; the
   instant those change it is invalid. Both paths drop it (the model path inherits ADR 0017 §2.1's
   signature-drop closure; the package path removes the signature part, its relationship, and its
   content-type override directly) so the package advertises _no_ signature rather than a broken one.

## Verification

Real Excel (`excel-gui-automation` open-verdict probe, Excel 16.0, interactive session; recorded facts,
not CI) on a genuine 10-module Excel-authored workbook with 7 document code-behinds, a 45 KB procedural
`JsonConverter`, and real references:

- splicing a **procedural** module into the original package (swapping only `vbaProject.bin`) is **clean**;
- splicing a **document** module (`ThisWorkbook`/`Contacts`) into the original package is **clean**,
  screenshot-confirmed (normal title, macros-disabled security bar, no repair). This is the
  from-scratch-impossible case, opening clean and recompiling.
- `editXlsxVbaModuleSource` editing a document module end-to-end is **clean**.

The `_VBA_PROJECT` recompile cookie interacts fine with a real p-code'd project, retiring the risk that a
recompile-all header would confuse a preserved p-code project. CI locks the parse round-trip (unit tests)
and a security-cluster corpus case
(`test/corpus/cases/xlsm-vba-edit-module-source-preserves-references.case.ts`) that drives the public path
and asserts a `REFERENCEREGISTERED` record survives verbatim, a document module edits in place, an
untouched code-page-1251 module stays byte-identical, and `_VBA_PROJECT` resets to the cookie.

## Consequences

- **Positive:** editing an existing macro's source is now in scope, including document and designer
  code-behinds, which from-scratch authoring cannot touch, with references, host info, and untouched
  modules preserved. `editXlsxVbaModuleSource(s)` makes the natural functional path Excel-clean on rich
  real files today, without waiting on the general writer.
- **Scope unchanged elsewhere:** the read view is untouched; adding or removing modules or references is
  _not_ in this slice (edit-existing-source only); executing macros remains permanently out of scope
  (ADR 0013).
- **Flagged separately, now diagnosed and largely closed:** the general `writeXlsx` whole-package
  repair-prompt on rich real workbooks (surfaced by the no-edit control probe) was a broader
  writer-fidelity investigation, independent of this feature. Diagnosed since: a referential-integrity
  sweep of the re-emitted package found the _sole_ structural break to be a **dropped external-workbook
  link**, where the writer dropped `xl/externalLinks/*` and the workbook's `<externalReferences>` while
  keeping the `[1]…` formulas that resolve through it, dangling the reference. Fixed by carrying external
  links through the workbook preserved-reference net (external targets and the `<externalReferences>`
  ordering now round-trip; corpus case `external-workbook-link-survives-roundtrip`). The remaining
  differences on such a file (tables, theme and doc-props re-emitted in reduced-but-valid form; calcChain
  regenerated by Excel) are not repair causes. A confirming Excel re-probe of the model path is the
  outstanding step (a probe, not CI, per ADR 0013). The package-preserving path remains the
  highest-fidelity choice for VBA edits regardless. See
  `docs/knowledge/specs/external-workbook-reference-formulas.md` and
  `docs/knowledge/specs/excel-repair-on-open-structural-constraints.md`.
- **Revisit when:** a consumer needs to add or remove modules or references (a new slice modeling those
  records), or when the general writer reaches whole-package parity, at which point the model path matches
  the package path on fidelity and the guidance in decision 5 can relax.

Related: ADR 0017 (VBA authoring in scope; from-scratch synthesis), ADR 0016 (VBA read view, unchanged
here), ADR 0013 (Excel as a test oracle, never a runtime),
`docs/knowledge/specs/xlsm-macro-preservation.md`.
