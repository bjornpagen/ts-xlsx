# ADR 0020: The customUI ribbon is readable through a typed view; authoring stays deferred

**Status:** Accepted (2026-07-24) · customUI read slice

## Context

A macro-enabled workbook can customise the Office ribbon through one or two parts hung off the
**package root** `_rels/.rels`: `customUI/customUI.xml` (Office 2007 RibbonX, namespace
`.../office/2006/01/customui`) and `customUI/customUI14.xml` (Office 2010+, namespace
`.../office/2009/07/customui`, which also adds backstage, QAT and commands). Until now the library
handled both **only** opaquely. Read captured each as a `PreservedRootReference` (bytes, content type,
relationship closure) and the writer re-declared it verbatim in the regenerated root rels, so a ribbon
survived a load/edit/save intact but **nothing could look inside it** through the public API.
`docs/knowledge/specs/xlsm-macro-preservation.md` named "parse the ribbon XML into a model, or leave it
opaque?" as a deferred question: round-trip fidelity was solved, but no consumer had forced a reader.

This mirrors exactly the position VBA was in before ADR 0016. Preservation proved the bytes were there;
a typed read view was a bounded, self-contained slice that did not need a downstream caller to justify
existing, the same way `Workbook.vbaProject` did not. CLAUDE.md §3 ("bias to action… make the
reasonable call") authorises building it now.

The schema was pinned against the Microsoft-published references (MS-CUSTOMUI / MS-CUSTOMUI2 via the
Microsoft Learn docs), **not** guessed. Doing so surfaced a latent wrong fact in the existing fixture
(see Consequences).

## Decision

1. **The ribbon is exposed as a read-only, typed _view_, not a new source of truth.**
   `Workbook.customUI: readonly CustomUiDocument[]` parses each preserved-root `customUI` part
   **lazily** and memoises the result. Each `CustomUiDocument` carries its `dialect` (`'2007'` or
   `'2010'`) and the parsed `Ribbon` (`tabs → groups → controls`). Barrel-exported from `src/customui`
   alongside `parseCustomUi`, `CustomUiParseError`, the control/tab/group types, and the namespace and
   relationship-type constants.

2. **Preservation stays the sole emission authority.** There is **no** write path from `CustomUiDocument`
   back to bytes. Editing a workbook re-emits the original `customUI` XML unchanged, so the ribbon read
   view is strictly additive and cannot regress ribbon preservation, because only one representation is
   ever serialised. Same invariant as the VBA read view (ADR 0016 §2).

3. **Dialect is keyed off the XML namespace, not the relationship type.** The `<customUI>` root
   namespace is the authoritative signal; the OPC relationship type is only used to _discover_ which
   root references are ribbon parts (`isCustomUiRelType`, matching the `/ui/extensibility` suffix that
   both real rel types share, `.../2006/relationships/ui/extensibility` for 2007 and, confusingly,
   `.../2007/relationships/ui/extensibility` for 2010). This was a deliberate robustness call:
   the relationship type is frequently mis-copied (see Consequences), the namespace is not.

4. **Scope is the `<ribbon>` subtree only.** Tabs, groups and controls, plus each control's callback
   names (`onAction` is the whole reason a macro workbook ships a ribbon). A document's `<commands>`,
   `<backstage>`, `<contextMenus>`, and the ribbon's `qat` and `contextualTabs` are **not** parsed. They
   still round-trip byte-for-byte, they are simply not surfaced. `CustomUiDocument.ribbon` is `undefined`
   for a document that customises only those. The shape leaves room to add `backstage` and `qat` later
   without breaking callers.

5. **The control model is a discriminated union keyed by element name, over a shared base.**
   `RibbonControl.kind` is the closed set of RibbonX control element local-names, with `'unknown'` as
   the fallback so an unrecognised element is surfaced rather than dropped. The common attributes
   (`id` / `idQ` / `idMso` / `label` / `onAction`) are lifted onto the typed view; container controls
   carry `children`; and the **full raw attribute map is preserved on every control**, so nothing is
   lost. _Exhaustive per-control attribute typing is deliberately deferred._ Pinning the exact schema of
   ~15 control types for a no-consumer v1 is a lot of API for a lot of risk, and the raw map already
   covers the many `get*` dynamic callbacks and layout hints the typed fields don't. When a consumer
   needs a specific control's exotic attributes typed, that's an additive refinement.

6. **The parser is treated as hostile-input-facing (CLAUDE.md §3).** It builds on the entity-safe
   `xmlEvents` scanner (no DTD or entity expansion), caps nesting depth so a deeply-nested part cannot
   overflow the recursive walk, and fails closed with `CustomUiParseError` on malformed XML, an
   unbalanced tree, a missing `<customUI>` root, or an unrecognised namespace, never a crash or a
   half-built tree. A ribbon-free workbook yields an empty array, never an error.

7. **Authoring stays out of scope.** Reading only. Mutating a ribbon back into valid RibbonX with correct
   `idQ` and callback wiring is a separate, larger effort with its own consumer-need question, the same
   posture ADR 0016 held for VBA authoring in its first slice.

## Consequences

- **Positive:** `workbook.customUI[*].ribbon?.tabs[*].groups[*].controls[*].onAction` reads the ribbon
  and its macro callbacks through a precisely-typed view, with **zero new dependencies** (it reuses the
  reader's `xmlEvents` SAX scanner and fflate's UTF-8 decode) and no risk to the preservation guarantee.
- **Corrected on the way in, with a latent wrong fact fixed.** The `preserved-parts.test.ts` round-trip
  fixture declared the `customUI14` relationship as `Type=".../office/2009/07/customui"`, the schema
  _namespace_ copy-pasted where the relationship _type_ belongs. The round-trip test passed anyway
  because preservation is verbatim, faithfully carrying garbage in and garbage out, so the error was
  invisible until a reader depended on the type. It is corrected to the real
  `.../office/2007/relationships/ui/extensibility` so the fixture represents a genuine Office file, and
  the reader keys dialect off the namespace to be robust to exactly this class of mistake in the wild.
- **Negative / deferred:** callers can _read_ the ribbon tree but not _create or edit_ it, and
  backstage, QAT, contextualTabs and commands remain opaque, though they still round-trip. Forward map,
  each gated on a forcing consumer: parse the customUI14 backstage and QAT surface; type a specific
  control's full attribute set; and the large one, ribbon authoring (model to valid RibbonX bytes).
- **Revisit when:** a concrete consumer needs backstage data, a fully-typed control, or ribbon authoring.
  No new ADR unless the shape of the decision changes.
