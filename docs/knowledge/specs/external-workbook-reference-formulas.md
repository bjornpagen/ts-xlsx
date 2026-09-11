# External-workbook reference formulas need package machinery, not a magic string

Cluster: formulas

## Scenario

A user wants a cell formula that pulls a value from a different workbook file, say `Sheet1!A1` of a
sibling workbook `测试.xlsx` in the same folder. They set the cell formula to the literal text
`=[测试.xlsx]Sheet1!A1`. When opened in Excel or WPS, neither resolves the reference: Excel interprets
the bracketed name against its own default document path, and WPS only resolves its own
`=./[file.xlsx]Sheet1!A1` dialect. The reference does nothing because the library passed the string
through verbatim without generating any of the package machinery an external reference actually
requires.

> Spec note, not a corpus case: this is a capability the library does not model at all
> (cross-workbook references), not a reproducible bug with an implementation-blind assertion against
> today's code. It records the OOXML mechanism and the API design questions for Phase 3.

## Desired behavior

The library should author, and round-trip, formulas that reference cells or ranges in _other_ workbook
files, such that Excel resolves them on open. In `.xlsx` an external reference is **not** the literal
`[filename.xlsx]Sheet1!A1` text inside `<f>`. It requires dedicated parts:

- An **externalLink part** (`xl/externalLinks/externalLink1.xml`) declaring the target workbook, its
  sheet names, and a cached snapshot of the referenced values, so the file shows a value when the
  target is offline.
- A **relationship** from that part to the target workbook, `TargetMode="External"`, whose `Target`
  (`测试.xlsx`) is where the relative or absolute path actually lives. The bracketed name in the
  formula is only an index, not the path.
- A workbook-level `<externalReferences>/<externalReference r:id="…"/>` entry plus rel wiring,
  assigning each external workbook a 1-based index.
- The cell formula then uses that **index** in brackets: `[1]Sheet1!A1`, not the filename. Writing the
  raw filename in brackets produces a formula Excel cannot bind, which is the reported "does nothing".

## Already implemented: faithful round-trip preservation

_Authoring_ a new external reference is still unbuilt, but _preserving_ one that a read file already
carries is done. An `externalLink` part is captured through the workbook-level preserved-reference net,
the same machinery that carries pivot and slicer caches and the VBA project, so a read-then-write
round-trip re-emits, intact:

- the `externalLink` part(s) and their content-type overrides;
- each link's own `TargetMode="External"` relationship, the pointer to the source workbook, which
  required teaching the preserved-part closure to retain external relationships verbatim, since it
  previously dropped every external target; and
- the workbook's `<externalReferences>` block, re-emitted in its **original order** so every `[n]` a
  formula or defined name resolves an external cell through still points at the same linked workbook.

This closes a whole-package writer-fidelity gap: before, a no-op load-then-save dropped the link while
keeping the `[n]`-using formulas, dangling the reference and prompting an Excel repair on open. Locked by
the `external-workbook-link-survives-roundtrip` corpus case. This is byte-preservation, not a structured
model, and the reference is not yet surfaced as an inspectable value (see the Read-side open question).

## Open questions

- **Public API shape:** a first-class cell value kind carrying `{path, sheet, cellOrRange,
cachedValue}`, structured and in the spirit of a typed API, against parsing `[name]Sheet!ref` magic
  strings, which is fragile and collides with defined-name and table syntax. Prefer the structured
  value.
- **Relative and absolute paths and `TargetMode="External"`:** let the user control the relationship
  `Target` so relative sibling paths work, and document that Excel resolves relative to the host file's
  folder (legacy fell back to the user's Documents folder).
- **Cached values:** write a plausible cached result so the reference displays before recalculation,
  and decide the policy when unknown.
- **Read side:** round-tripping faithfully is **done** (see _Already implemented_ above, where the link,
  its external target, and the `<externalReferences>` ordering all survive). Still open: parsing those
  parts into _structured external-reference values_ the caller can inspect and edit, rather than
  preserving them as opaque bytes.
- **Scope:** whole external workbooks, named ranges in external workbooks, and DDE/OLE links are
  distinct sub-features; start with cell and range references to another `.xlsx` by path.

Related: `defined-name-formula-expression`, `internal-hyperlink-target-portability`,
`formula-recalculation-expectations`, `cross-sheet-reference-preserved-in-formula-and-validation`,
`formula-cell-value-type-minimal-required-fields`, `excel-repair-on-open-structural-constraints`
(a dropped-but-referenced part is a repair cause), `xlsm-macro-preservation` (the same preserved-part
net).
