# ADR 0025: The workbook default font is declared, not assumed

**Status:** Accepted (2026-07-29) · Styles

## Context

The styles part's font id 0 is the workbook's default font: the face every cell that names no font
of its own renders in, **empty cells included**, and the Maximum Digit Width every character-unit
`<col width>` is expressed in. `<fonts>` always has a first entry, so there is no such thing as
declining to declare one; a writer that splices in a constant is _declaring a default_, not omitting
one.

The writer spliced in a constant. `DEFAULT_FONT_BODY` (`Calibri 11`, `theme="1"`, `family="2"`,
`scheme="minor"`) went into every package with no workbook input at all. Two failures followed, and
the second was the one nobody had noticed.

**A themed workbook rendered its unstyled cells in the wrong face.** `setTheme({fonts: {minor:
'Aptos'}})` wrote the theme part correctly and could not reach a cell. Font 0 went on claiming
`scheme="minor"`, meaning _I am the theme's body face_, while naming Calibri outright, and Excel
resolves the explicit name. The two parts of one package contradicted each other while each stayed
individually well-formed, which is why no validator caught it. Working around it meant setting
`font` on every column and naming the face in every rich-text run.

**A file's own default font did not survive being read and written back.** Reading a package whose
font 0 was Aptos Narrow and writing it _unmodified_ replaced the declared default with Calibri and
re-added Aptos Narrow as a redundant custom entry. Populated cells kept the face through that entry,
so the damage hid: only empty cells, and any consumer reading font 0 as "the workbook default",
saw Calibri. The layout meaning of every column width went with it. That directly violated a spec
note already in the tree (`default-font-must-not-be-assumed-for-column-widths`): _"The library must
not unconditionally inject its own assumed default font ahead of, or in place of, the default font a
file already declares."_

## Ground truth

Before designing anything, we read what Excel itself writes. The corpus already held twenty-odd
Excel-authored packages, so this cost nothing:

| fixture                                                         | theme `<a:minorFont><a:latin>` | font 0                                                                        |
| --------------------------------------------------------------- | ------------------------------ | ----------------------------------------------------------------------------- |
| `formula-string-result-under-date-format-roundtrip/source.xlsx` | `Aptos Narrow`                 | `<sz 11><color theme="1"><name Aptos Narrow><family 2><scheme minor>`         |
| `builtin-cjk-date-numfmt-ids…/source.xlsx`                      | `Calibri`                      | `<name 等线><family 2><scheme minor>`                                         |
| `vector-shape-drawing…/sample.xlsx`                             | none                           | `<sz 12><color theme="1"><name NotoSansCJKjp-Regular><family 2>`, no `scheme` |
| `column-width-and-pagesetup…/sample.xlsx`                       | none                           | `<sz 8><name Arial>`, no scheme, no family, no colour                         |
| `fill-border-color…/source.xlsx`                                | none                           | `<sz 10><color indexed="8"><name Helvetica Neue>`, no scheme                  |

Three facts, all load-bearing:

1. When font 0 **is** the theme's body face, Excel names the resolved face outright _and_ keeps
   `scheme="minor"`. So deriving the name from the theme is what reproduces Excel's own output.
2. When it is **not**, Excel writes no `<scheme>` at all. `scheme="minor"` is therefore a _claim_,
   and the first failure above is precisely that claim sitting beside a name contradicting it.
3. The face is **script-resolved, not latin-resolved**. The CJK file's theme names `Calibri` as its
   latin body face and Excel still wrote `等线` as font 0, having picked the `script="Hans"` entry
   inside the same `<a:minorFont>`. We do not model script resolution and are not going to.

## Decision

### The model carries three members, not one

```ts
get declaredDefaultFont(): Font | undefined   // what the source package stated
setDefaultFont(font: Font): void              // author one, merging (like setTheme)
get defaultFont(): Font                       // the resolved, complete result
```

They answer three different questions. `declaredDefaultFont` is the round-trip view and is
`undefined` when a file declared nothing, deliberately not a fabricated Calibri, because claiming a
file said something it did not is how the second failure happened. `defaultFont` is what an unstyled
cell actually renders in, is never `undefined`, and is what the writer emits; it mirrors
`themeColors`/`themeFonts`, which likewise fall back to the Office default rather than returning
nothing.

### The precedence chain, merging at every hop

```
explicit cell font > row/column default > authored defaultFont > authored theme body face
                                        > the source file's font 0 > source theme body face > Calibri 11
```

Merging: `setDefaultFont({size: 14})` keeps the resolved face and changes only the size, and calling
it twice accumulates, the shape `setTheme` already has and the shape every other style level in
the model already has.

The two authored levels outrank the file because authoring is an explicit act. Between them
`setDefaultFont` wins on the face, because it names font 0 outright while `setTheme` names it only
by implication.

_Authored theme body face above the file's own font 0_ is the hop worth justifying. Without it,
read → `setTheme({fonts: {minor}})` → write would leave the reported symptom in place for every
read-modify-write, which is the majority path. With it, restyling a workbook's body face reaches its
unstyled cells the same way it reaches its `scheme="minor"` ones.

_The file's font 0 above any derivation_ is fact 3 above. With **nothing** authored, the declared
font 0 passes through verbatim rather than being re-derived, or a CJK workbook would lose `等线` to
`Calibri` on every save.

### `scheme` and `family` are derived, never copied

`scheme` is `minor` exactly when the resolved face equals `themeFonts.minor`, and absent otherwise
(fact 2). `family` travels with it: both describe _the theme's body face_, so both hold while the
resolved face still is that face and are dropped when a caller names another. Either may be stated
outright by the caller, in which case the caller's word stands.

The exception is the verbatim pass-through above: with nothing authored, a file's `scheme="minor"`
beside `等线` rides through untouched. It is the producer's own script resolution, not a
contradiction we may "fix".

### Size and colour are completed, never merely merged

However partial the authored `Font`, the emitted font 0 always states a size and a colour. A font 0
stating neither is the "missing default font" foreign readers (Apple Numbers, and Excel in some
cases) warn about on open, because empty cells fall back to a default the file never properly
defines. We do **not** extend that to `family` and `scheme`: Excel itself writes bare font 0 entries
(`<sz 8><name Arial>`), and inventing metadata would be a guess about a face we cannot classify.

### The declared default also interns to id 0

The reader flattens font 0 onto every xf that names it, so a cell carrying only a fill, whose xf
still says `fontId="0"`, arrives with the file's default face as a concrete `cell.font`. That face
is an artefact of reading, not an authored intent: in the source file the cell said nothing about
its font. So `#internFont` treats both the resolved default _and_ the declared one as id 0. A cell
that merely inherited follows a newly authored default instead of stranding itself on the old face
through a custom entry while its unstyled neighbours move.

This is also what makes the round-trip _narrower_ rather than wider: the duplicate entry the second
failure produced now collapses.

## Consequences

**Emitted bytes change** for any workbook with an authored theme font, or read from a package whose
font 0 was not Calibri. That is the fix, not a side effect. A plain `new Workbook()` is unaffected:
its resolution lands on exactly the old constant, and a test holds the two together so they cannot
drift.

**The theme is not rewritten.** `setDefaultFont` writes font 0 and nothing else; the theme reaches
font 0 one-way, by derivation. Coupling them, so Excel's UI would label the face a theme font,
would have been higher fidelity to what Excel produces when a _user_ changes the body face, but it
makes one public API silently mutate another's state, and the derivation already removes the
symptom.

**No worksheet-level override.** OOXML has no per-sheet default font, so it could only be sugar that
stamps every column. The workbook level is where the correctness bug was; the sugar can be designed
on its own merits later. The spec note stays open for it.

**Column widths stay opaque.** Changing font 0 changes what a character-unit width _means_ without
changing the unit, and read→write stays faithful as long as we neither reinterpret nor recompute
them, which we do not. Pixel-accurate layout still needs a font-metric table and is still out of
scope. Recorded here so this is not later "fixed" by someone who spots the coupling.

**A cell that explicitly names the current default face collapses to id 0.** This is unobservable in
the source file's own terms, but it means a caller cannot pin a cell to today's default face and have
it stay behind when the default moves. Pinning is expressible by naming a face that differs in any
facet; wanting to pin to _exactly_ the default is not a use case we have seen.

## Alternatives rejected

**Stop flattening font 0 onto xfs on read.** It would have made the intern-as-id-0 rule unnecessary,
but `cell.font` would stop being a concrete face for any styled cell, which is a worse API and a
documented behaviour with tests behind it.

**Keep the constant and add the default font as a separate emission step.** Two places would then
decide what font 0 is, and `#internFont`'s dedup would have to know about both. One resolved value,
computed in the model and handed to the registry at construction, is the whole reason the duplicate
entry disappears.

## References

- `docs/knowledge/specs/default-font-workbook-worksheet-level.md`
- `docs/knowledge/specs/default-font-must-not-be-assumed-for-column-widths.md`
- `test/corpus/cases/workbook-default-font-is-declared-not-assumed.case.ts`
