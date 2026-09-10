# Workbook

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `AddImageOptions`

<sub>interface</sub>

A picture registered on the workbook, ready to be anchored to a worksheet.

```ts
interface AddImageOptions {
  /** The image bytes. */
  readonly buffer: Uint8Array;
  /** The file kind: `"png"`, `"jpeg"`/`"jpg"`, `"gif"`, … A leading dot or a URL query string is
   * tolerated and stripped; omit it entirely to infer the kind from the bytes' magic number. */
  readonly extension?: string;
}
```

---

### `AddWorksheetOptions`

<sub>interface</sub>

```ts
interface AddWorksheetOptions {
  readonly state?: WorksheetState['state'];
}
```

---

### `DEFAULT_WORKBOOK_VIEW`

<sub>const</sub>

The window geometry a workbook starts from: the values desktop Excel writes for its own default
window.

A default is emitted rather than the element left out because Excel writes `<bookViews>` into every
file it saves and consumers lay panes out against that rect. With no view at all the frozen-pane
split is computed against an uninitialised window, and the frozen region can stay unpainted until
some later event forces a relayout.

```ts
const DEFAULT_WORKBOOK_VIEW: { readonly x: -110; readonly y: -110; readonly width: 19420; readonly height: 12220; readonly activeTab: 0; }
```

---

### `DefinedName`

<sub>interface</sub>

A named reference in the workbook, one of the entries Excel surfaces in its Name Manager. A name
maps an identifier to a formula (`refersTo`), most often a cell range like `Sheet1!$A$1:$B$2` but
possibly any formula. A name is global to the workbook unless it names a sheet in `scope`,
which restricts it to that sheet and lets another sheet reuse the same name independently.

```ts
interface DefinedName {
  /** The name as typed in a formula, e.g. `"TaxRate"`. Built-in names carry an `_xlnm.` prefix. */
  readonly name: string;
  /** The formula the name resolves to, e.g. `"Sheet1!$A$1:$B$2"`. */
  readonly refersTo: string;
  /** The sheet the name is scoped to; omit for a workbook-global name. */
  readonly scope?: string;
  /** A human note shown beside the name in Excel's Name Manager. */
  readonly comment?: string;
  /** Hide the name from the Name Manager UI without removing it. */
  readonly hidden?: boolean;
}
```

---

### `PreservedTheme`

<sub>interface</sub>

The workbook's theme part, captured verbatim from a source package: the `<clrScheme>`,
`<fontScheme>`, and `<fmtScheme>` every `theme="n"` colour reference and every `scheme="major|minor"`
font in the file resolves against.

Held opaquely, like [`Workbook.restoreDifferentialStyles`](./workbook.md#workbook)'s `<dxf>` fragments: the model does
not interpret the theme, it only refuses to destroy it. `entryPath` is where the source package held
the part (reached through the workbook's `.../theme` relationship, so not necessarily
`xl/theme/theme1.xml`), and `parts` is the transitive closure it reaches, the entry included. The
closure matters because a theme can carry its own relationships: a picture used as a themed fill is
wired by an `r:embed` into the theme's rels part, and re-emitting the theme without it would leave
that reference dangling.

```ts
interface PreservedTheme {
  readonly entryPath: string;
  readonly parts: readonly PreservedPart[];
}
```

---

### `PreservedWorkbookReference`

<sub>interface</sub>

A workbook-level reference to package content the model does not model: a pivot cache
(`pivotCacheDefinition`) or a slicer cache (`slicerCache`), preserved verbatim across a round-trip
instead of being dropped. `relType` is the workbook relationship Type URI to re-emit; `entryPath`
is the part it points at; `parts` is the transitive closure that reference reaches (the entry
included). `pivotCacheId` carries the `<pivotCache cacheId>` a pivot cache is registered under in
the workbook's `<pivotCaches>`, so the wiring a pivot table resolves its cache through is re-emitted
too; it is absent for a slicer cache, which the workbook lists in an extension block instead.
`externalReferenceIndex` is the 0-based position of an `externalLink` within the workbook's
`<externalReferences>`, the `[n]` a formula or defined name resolves an external cell through, so
the block is re-emitted in the original order and every `[n]` still points at the same linked
workbook; it is absent for a pivot/slicer cache.

```ts
interface PreservedWorkbookReference {
  readonly relType: string;
  readonly entryPath: string;
  readonly parts: readonly PreservedPart[];
  readonly pivotCacheId?: string;
  readonly externalReferenceIndex?: number;
}
```

---

### `Workbook`

<sub>class</sub>

```ts
class Workbook {
  readonly properties: WorkbookProperties = {};
  readonly view: WorkbookView = {...DEFAULT_WORKBOOK_VIEW};
  fullCalcOnLoad = false;
  dateEpoch: DateEpoch = 1900;
  codeName?: string;
  protection: WorkbookProtection | undefined = undefined;
  get worksheets(): readonly Worksheet[];
  get activeTabIndex(): number;
  get preservedReferences(): readonly PreservedWorkbookReference[];
  get preservedRootReferences(): readonly PreservedRootReference[];
  get customUI(): readonly CustomUiDocument[];
  get vbaProject(): VbaProject | undefined;
  get vbaProjectBytes(): Uint8Array | undefined;
  set vbaProjectBytes(bytes: Uint8Array | undefined);
  get vbaProjectSigned(): boolean;
  get vbaProjectSignatures(): readonly VbaProjectSignature[];
  removeVbaModule(name: string): void;
  addVbaReference(ref: VbaLibraryReference): void;
  get differentialStyles(): readonly string[];
  get indexedColors(): readonly string[];
  get mruColors(): readonly string[];
  get tableStyles(): TableStyleTable;
  addTableStyle(style: TableStyle): void;
  get customTableStyles(): readonly TableStyle[];
  get themePart(): PreservedTheme | undefined;
  setTheme(overrides: ThemeOverrides): void;
  get themeColors(): ThemeColorScheme;
  get themeFonts(): ThemeFontScheme;
  get declaredDefaultFont(): Font | undefined;
  setDefaultFont(font: Font): void;
  get defaultFont(): Font;
  get themeOverrides(): ThemeOverrides | undefined;
  resolveColor(color: Color): string | undefined;
  get namedStyles(): readonly NamedCellStyle[];
  addPerson(person: Person): void;
  get persons(): readonly Person[];
  getPerson(id: string): Person | undefined;
  addImage(options: AddImageOptions): number;
  get media(): readonly WorkbookImage[];
  getImage(id: number): WorkbookImage | undefined;
  exportImages(sheet: Worksheet): WorksheetImages;
  importImages(sheet: Worksheet, images: WorksheetImages): void;
  get definedNames(): readonly DefinedName[];
  defineName(definedName: DefinedName): void;
  addWorksheet(name: string, options: AddWorksheetOptions = {}): Worksheet;
  getWorksheet(nameOrId: string | number): Worksheet | undefined;
  requireWorksheet(nameOrId: string | number): Worksheet;
}
```

**Members**

#### `Workbook.view`

```ts
readonly view: WorkbookView = {...DEFAULT_WORKBOOK_VIEW};
```

The workbook's window state: position, size, and the selected sheet. Always present (see
[`DEFAULT_WORKBOOK_VIEW`](./workbook.md#defaultworkbookview) for why it is defaulted rather than left unset) and always written.
Reading a file replaces it with that file's saved geometry, so a round-trip restores the window
the author left rather than stamping ours over it.

#### `Workbook.fullCalcOnLoad`

```ts
fullCalcOnLoad = false;
```

Ask consuming spreadsheet apps to recalculate every formula when the file is opened, rather than
trusting the cached results stored with each formula cell. Set this when the producer cannot
compute formula results itself. This is the OOXML `fullCalcOnLoad` flag, off by default, so a
workbook whose cached results are authoritative stays unmarked.

#### `Workbook.dateEpoch`

```ts
dateEpoch: DateEpoch = 1900;
```

Which date system this workbook's serials count in: `1900` (the Windows default) or `1904` (the
`date1904` flag of `<workbookPr>`, Excel for Macintosh's original). It governs every conversion
between a `Date` and the number a cell actually stores, in both directions, so setting it after
cells hold dates changes what those cells mean rather than converting them.

Read from the file and written back, because dropping it is not a cosmetic loss: the serials stay
as they were and the consumer re-reads them under the other system, so the workbook silently
changes meaning by four years and a day.

#### `Workbook.codeName`

```ts
codeName?: string;
```

The workbook's VBA identity (`<workbookPr codeName>`), the name a macro means by `ThisWorkbook`.
Undefined for a workbook with no VBA project, which is what Excel writes for one.

Preserved rather than modeled: nothing here reads it, but a `.xlsm` whose code name is dropped on
a round trip has had the binding between its macros and its document cut.

#### `Workbook.protection`

```ts
protection: WorkbookProtection | undefined = undefined;
```

Workbook-level structure/window protection: the OOXML `<workbookProtection>` element. Absent by
default (an unprotected workbook). Set it to lock the workbook shell, or leave it as read from a
file so a protected workbook stays locked across a passthrough save rather than being silently
unlocked. Distinct from a worksheet's own `protect()`, which guards a single sheet's cells.

#### `Workbook.worksheets`

```ts
get worksheets(): readonly Worksheet[];
```

The worksheets in insertion order.

#### `Workbook.activeTabIndex`

```ts
get activeTabIndex(): number;
```

The 0-based index of the active sheet: [`WorkbookView.activeTab`](./workbook.md#workbookview) resolved against the sheets
that actually exist. Exactly one sheet is always active: an out-of-range tab (a caller's stale
index, or a file whose sheet was removed after the view was saved) falls back to the first sheet
rather than to none, because a package where no sheet is selected gives the consumer no view to
initialise on open.

#### `Workbook.preservedReferences`

```ts
get preservedReferences(): readonly PreservedWorkbookReference[];
```

The workbook-level preserved references, in the order they were read.

#### `Workbook.preservedRootReferences`

```ts
get preservedRootReferences(): readonly PreservedRootReference[];
```

The package-root preserved references, in the order they were read.

#### `Workbook.customUI`

```ts
get customUI(): readonly CustomUiDocument[];
```

The ribbon customisations decoded from this workbook's `customUI` parts: `customUI.xml` (Office
2007) and/or `customUI14.xml` (Office 2010+), in the order their root relationships were read. Each
[`CustomUiDocument`](./customui-ribbon.md#customuidocument) is tagged with its dialect and exposes the parsed `<ribbon>` tree. Empty
for a workbook that customises no ribbon.

This is a **read-only view** over parts the writer already round-trips verbatim. Mutating the
returned objects changes nothing on write; the original `customUI` XML is re-emitted byte-for-byte
regardless. Parsed lazily on first access and memoised.

**Throws:** [`CustomUiParseError`](./customui-errors.md#customuiparseerror) if a `customUI` part is present but its XML is malformed.

#### `Workbook.vbaProject`

```ts
get vbaProject(): VbaProject | undefined;
```

The VBA project decoded from this workbook's preserved `vbaProject.bin`, or `undefined` for a
workbook with no macros. This is a **read-only view** over the bytes the writer already round-trips
verbatim. Mutating the returned object changes nothing on write; the original macro blob is
re-emitted byte-for-byte regardless. Parsed lazily on first access and memoised.

**Throws:** [`VbaParseError`](./vba-errors.md#vbaparseerror) if a macro project is present but its `vbaProject.bin` is malformed.

#### `Workbook.vbaProjectBytes`

```ts
get vbaProjectBytes(): Uint8Array | undefined;
set vbaProjectBytes(bytes: Uint8Array | undefined);
```

The raw `vbaProject.bin` bytes attached to this workbook (the exact macro blob the writer will
embed), or `undefined` for a workbook with no macros. The getter returns a defensive copy, so
mutating it changes nothing on write.

Assigning bytes attaches (or replaces) the macro project: the written package becomes
macro-enabled and re-embeds these bytes verbatim. The bytes must be a well-formed VBA container
(a CFB holding a `dir` stream); a malformed blob is rejected with [`VbaParseError`](./vba-errors.md#vbaparseerror) rather
than written out to produce a package Excel would flag for repair. This is the attach-blob path:
copy a project between workbooks with `dst.vbaProjectBytes = src.vbaProjectBytes`, or import a
`.bin` produced by another tool. Assigning `undefined` removes the project, reverting the workbook
to a plain (non-macro) package.

Replacing or removing the project also drops any digital signature the previous blob carried,
because a signature over the old bytes cannot validate new ones, so the result never advertises
a broken signature.

#### `Workbook.vbaProjectSigned`

```ts
get vbaProjectSigned(): boolean;
```

Whether this workbook's VBA project carries a digital signature: `true` if any signature part is
attached, `false` for an unsigned project or a workbook with no macros.

This reflects the **presence** of a signature blob, not its cryptographic validity: the library
neither parses the PKCS#7/CMS structure nor validates the certificate chain or signer. A `true`
here means "a signature is attached," never "this signature is valid." Replacing or editing the
project drops its signatures (a signature over the old bytes cannot validate new ones), so this
reads `false` again after [`vbaProjectBytes`](./workbook.md#workbookvbaprojectbytes), [`removeVbaModule`](./workbook.md#workbookremovevbamodule), or
[`addVbaReference`](./workbook.md#workbookaddvbareference) mutates the project. See [`vbaProjectSignatures`](./workbook.md#workbookvbaprojectsignatures) for the raw bytes and
which generation(s) are present.

#### `Workbook.vbaProjectSignatures`

```ts
get vbaProjectSignatures(): readonly VbaProjectSignature[];
```

The digital signatures attached to this workbook's VBA project, in the order their relationships
are wired off `vbaProject.bin`. Up to three generations (legacy, agile, V3) can coexist over the
same project bytes. Empty for an unsigned project or a workbook with no macros.

Each entry's `bytes` are the raw signature blob passed through verbatim; the library does not parse
or verify them (see [`vbaProjectSigned`](./workbook.md#workbookvbaprojectsigned) on presence-vs-validity). Hand a blob to an external
verifier if you need cryptographic validation; that is deliberately out of this library's scope.

#### `Workbook.removeVbaModule`

```ts
removeVbaModule(name: string): void;
```

Remove a standard module from this workbook's existing macro project, in place: a structural splice
that leaves every remaining module's compiled p-code untouched (see [`removeVbaModule`](./workbook.md#workbookremovevbamodule)).
Replacing the project also drops a stale signature, as [`vbaProjectBytes`](./workbook.md#workbookvbaprojectbytes) does.

Only `procedural` and `class` modules can be removed this way. See [`removeVbaModule`](./workbook.md#workbookremovevbamodule) for why.
To author or edit module *source* (which needs real compiled p-code), use the offline
`tools/vba-compiler`, then attach its output via [`vbaProjectBytes`](./workbook.md#workbookvbaprojectbytes).

**Throws:** [`VbaAuthorError`](./vba-errors.md#vbaauthorerror) if the workbook has no macro project, or `name` is not in the project,
or names a `document`/`designer` module.
**Throws:** [`VbaParseError`](./vba-errors.md#vbaparseerror) if the attached `vbaProject.bin` is malformed.

#### `Workbook.addVbaReference`

```ts
addVbaReference(ref: VbaLibraryReference): void;
```

Add a registered (COM type-library) reference to this workbook's existing macro project, in place.
Every existing module, reference, and host-info record rides through unchanged (see
[`addVbaReference`](./workbook.md#workbookaddvbareference)). Replacing the project also drops a stale signature, as
[`vbaProjectBytes`](./workbook.md#workbookvbaprojectbytes) does.

**Throws:** [`VbaAuthorError`](./vba-errors.md#vbaauthorerror) if the workbook has no macro project, or any field of `ref` is invalid
(see [`VbaLibraryReference`](./vba-project-editor.md#vbalibraryreference)).
**Throws:** [`VbaParseError`](./vba-errors.md#vbaparseerror) if the attached `vbaProject.bin` is malformed.

#### `Workbook.differentialStyles`

```ts
get differentialStyles(): readonly string[];
```

The preserved differential-style (`<dxfs>`) fragments, in index order.

#### `Workbook.indexedColors`

```ts
get indexedColors(): readonly string[];
```

The preserved custom indexed-color palette, in index order; empty when the default palette rules.

#### `Workbook.mruColors`

```ts
get mruColors(): readonly string[];
```

The preserved most-recently-used colour swatches, in order; empty when the file declared none.

#### `Workbook.tableStyles`

```ts
get tableStyles(): TableStyleTable;
```

The preserved `<tableStyles>` block; `styles` is empty when the file declared no custom style.

#### `Workbook.addTableStyle`

```ts
addTableStyle(style: TableStyle): void;
```

Register a custom table style: a named look a table applies to itself by putting that name in
[`TableStyleInfo.name`](./tables.md#tablestyleinfo), exactly as it would name one of Excel's built-in gallery styles.

```ts
workbook.addTableStyle({
  name: 'Harbour',
  elements: {
    wholeTable: {border: {top: {style: 'thin'}, bottom: {style: 'thin'}}},
    headerRow: {font: {bold: true, color: {argb: 'FFFFFFFF'}},
                fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FFBB2649'}}},
    firstRowStripe: {fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FFF6E7EB'}}},
  },
});
sheet.addTable({name: 'Cargo', ref: 'A1:B3', columns, style: {name: 'Harbour'}});
```

Each element's formatting is interned into the workbook's shared differential-style table, so two
elements, or a conditional-formatting rule, that paint the same way share one entry.

Registering a name a source file already defined **overrides** that definition rather than adding
a second one beside it.

**Throws:** [`AuthoringError`](./errors.md#authoringerror) if the name is empty, or an element carries a `size` outside the four
stripe types. See `checkTableStyle` for why those are refused here rather than silently
dropped.
**Throws:** `RangeError` if a `size` is not a positive integer.

#### `Workbook.customTableStyles`

```ts
get customTableStyles(): readonly TableStyle[];
```

The table styles authored on this workbook, in registration order.

#### `Workbook.themePart`

```ts
get themePart(): PreservedTheme | undefined;
```

The preserved theme part, or undefined when the workbook rides the library's default theme.

#### `Workbook.setTheme`

```ts
setTheme(overrides: ThemeOverrides): void;
```

Author the workbook's theme: any subset of the twelve colour-scheme slots, and either of the two
typefaces. Merges into what the workbook already has, so branding one accent leaves the other
eleven slots alone, and calling it twice accumulates.

This is the workbook-wide palette. A cell that names a colour as `theme="4"`, which is what Excel
writes whenever a user picks from the theme row of the colour picker, follows `accent1` here, so
one call restyles every such cell, chart and table style at once. Colours are `RRGGBB`; a leading
`#` and an 8-hex ARGB are both accepted and reduced, and anything else throws rather than writing
a value Excel silently renders as flat black.

What it does **not** touch: the theme's format scheme, meaning the gradient, line and effect
styles that give a theme its texture. Those ride through from the source theme (or the library's
default) untouched, because nobody hand-authors gradient stops from a spreadsheet API and
regenerating them would replace a designer's work with the Office default. For the same reason a
slot left unauthored keeps the source's own encoding, including the `<a:sysClr>` form Excel uses
for `dk1`/`lt1` so they follow the viewer's window colours.

**Throws:** `SyntaxError` if a colour is not 6 or 8 hexadecimal digits, with or without a leading `#`.
Every colour is read before any is stored, so a refused call changes nothing.

#### `Workbook.themeColors`

```ts
get themeColors(): ThemeColorScheme;
```

The colour scheme every `theme="n"` reference in this workbook resolves against: anything
[`setTheme`](./workbook.md#workbooksettheme) authored, over the preserved theme's `<a:clrScheme>`, over the Office default.

Note the slot *order*: `theme="0"` is `lt1` and `theme="1"` is `dk1`, which is not the order the
slots appear in the theme part. See [`THEME_COLOR_SLOTS`](./theme.md#themecolorslots).

Every slot reads back as six uppercase hexadecimal digits, whichever spelling [`setTheme`](./workbook.md#workbooksettheme) was
given or the theme part used.

#### `Workbook.themeFonts`

```ts
get themeFonts(): ThemeFontScheme;
```

The theme's major (heading) and minor (body) typefaces, authored values over the source's.

#### `Workbook.declaredDefaultFont`

```ts
get declaredDefaultFont(): Font | undefined;
```

The default font as the source package declared it: font id 0 of its styles part, the face every
cell that names no font of its own renders in. `undefined` for a workbook authored from scratch or
read from a package carrying no styles part: nothing was declared, and the library does not
fabricate a declaration on the file's behalf.

This is the *round-trip* surface. [`defaultFont`](./workbook.md#workbookdefaultfont) is what the workbook actually renders in,
which is this once anything has been authored over it.

#### `Workbook.setDefaultFont`

```ts
setDefaultFont(font: Font): void;
```

Author the workbook's default font: the face, size and colour every cell with no font of its own
renders in, **empty cells included**. Merges into whatever the workbook already had, so
`setDefaultFont({size: 14})` keeps the resolved face and changes only the size, and calling it
twice accumulates. This is the one knob that reaches a cell no row or column default can: an
untouched cell in an unformatted column.

It writes the styles part's font 0 and **nothing else**. In particular it does not rewrite the
theme's body typeface. The dependency runs the other way: with no default font authored, font 0
follows [`themeFonts`](./workbook.md#workbookthemefonts)'s minor face, so `setTheme({fonts: {minor}})` already reaches every
unstyled cell and needs no second call here. See [`defaultFont`](./workbook.md#workbookdefaultfont) for the full chain.

**Throws:** `RangeError` if `size` is not a positive finite number, or `name` is empty. Both
produce a styles part Excel renders from some other font without ever reporting why. Native
rather than [`AuthoringError`](./errors.md#authoringerror): one argument out of range is what `RangeError` is for.

#### `Workbook.defaultFont`

```ts
get defaultFont(): Font;
```

The font every cell that names none of its own renders in, resolved and complete: what the writer
emits as font id 0. Never `undefined`: a workbook always renders in *some* face, and the chain
below always reaches one.

```
authored default font  >  authored theme body face  >  the source file's font 0  >  theme body face
```

The two authored levels outrank the file because authoring is an explicit act; between them
[`setDefaultFont`](./workbook.md#workbooksetdefaultfont) wins on the face because it names font 0 outright while
[`setTheme`](./workbook.md#workbooksettheme) names it only by implication. With **nothing** authored the file's own font 0
passes through verbatim. That is deliberate, because a producer resolves that face by script and
we do not: Excel writes `等线` as font 0 under a theme whose latin body face is `Calibri`, and
re-deriving would silently rewrite it.

`family` and `scheme` describe the *theme's* body face, so they are carried exactly while the
resolved face still is that face and dropped when a caller names another, which is also what
Excel writes: a font 0 naming a non-theme face carries no `<scheme>` at all. Either may be stated
outright, in which case the caller's word stands.

#### `Workbook.themeOverrides`

```ts
get themeOverrides(): ThemeOverrides | undefined;
```

The colour slots and typefaces [`setTheme`](./workbook.md#workbooksettheme) authored on this workbook, or `undefined` when
none were, in which case the source theme (or the writer's default) rides through untouched.

A serializer composes these *over* the existing part rather than generating one from scratch, so
a preserved theme keeps its format scheme, its unauthored slots' exact encoding, and the
relationships it carries. Its colours are already reduced to the six uppercase hexadecimal digits
a slot holds.

#### `Workbook.resolveColor`

```ts
resolveColor(color: Color): string | undefined;
```

Resolve a colour reference to a concrete 8-hex ARGB string, or `undefined` when it does not
resolve to a fixed colour: an `auto` colour, one of the two system indexed colours, or a theme
slot this workbook's scheme does not declare.

This is a *derived* view, not a rewrite: the [`Color`](./styles.md#color) stays exactly as its file encoded it,
so a round-trip re-emits `theme="4" tint="0.4"` rather than a literal ARGB. Resolving into the
model would sever every cell's link to the theme, so recolouring the workbook would stop working,
and would inflate the styles table with one distinct colour per shade.

A `theme` reference resolves through [`themeColors`](./workbook.md#workbookthemecolors); an `indexed` one through the workbook's
custom `<indexedColors>` palette when it declares one, else the built-in legacy palette. A `tint`
is applied last.

#### `Workbook.namedStyles`

```ts
get namedStyles(): readonly NamedCellStyle[];
```

The named cell styles, in index order (index 0 is Normal); empty when only the default exists.

#### `Workbook.addPerson`

```ts
addPerson(person: Person): void;
```

Register an identity a threaded comment can name: an author, or someone `@mentioned` in a message.
A message reaches it by [`Comment.personId`](./comment-thread.md#comment), a mention by [`Mention.personId`](./comment-thread.md#mention).

Keyed by [`Person.id`](./comment-thread.md#person) alone, so registering the same id twice replaces the entry rather than
adding a second: the id is the identity. Registering the same human twice under *different* ids is
legitimate and is what Excel itself does. See `restorePersons`. The id is normalised to the
brace-wrapped upper-case GUID form the format requires, so a `crypto.randomUUID()` is accepted as-is.

**Throws:** `SyntaxError` if the id is not a GUID.

#### `Workbook.persons`

```ts
get persons(): readonly Person[];
```

The registered threaded-comment identities, in the order they were read. That order carries no
meaning, because Excel re-sorts the registry by person id when it saves, so nothing may depend
on it.

#### `Workbook.getPerson`

```ts
getPerson(id: string): Person | undefined;
```

Look up a registered identity by its [`Person.id`](./comment-thread.md#person), or `undefined` if the registry has none.

#### `Workbook.addImage`

```ts
addImage(options: AddImageOptions): number;
```

Register a picture on the workbook and return its numeric id. Pass the id to
[`Worksheet.addImage`](./worksheet.md#worksheetaddimage) to anchor the picture to a sheet; the same id may be anchored on any
number of sheets and positions, and the bytes are still stored only once.

#### `Workbook.media`

```ts
get media(): readonly WorkbookImage[];
```

The registered images, indexed by the id [`addImage`](./workbook.md#workbookaddimage) returned.

#### `Workbook.getImage`

```ts
getImage(id: number): WorkbookImage | undefined;
```

Look up a registered image by its id, or `undefined` if no image carries that id.

#### `Workbook.exportImages`

```ts
exportImages(sheet: Worksheet): WorksheetImages;
```

Every picture `sheet` shows, resolved out of this workbook's media registry into the
workbook-independent form [`importImages`](./workbook.md#workbookimportimages) consumes. `sheet` must belong to this workbook:
that is whose registry its image ids index.

This is the attached-part half of a sheet copy, and it is deliberately a separate call from
[`Worksheet.model`](./worksheet.md#worksheetmodel): a model is a serialisable value, an image is bytes on the workbook, and
ADR-0005 keeps them apart. Carrying a sheet whole is therefore the two of them together:

```ts
destination.model = source.model;
destinationWorkbook.importImages(destination, sourceWorkbook.exportImages(source));
```

The exported pictures share the registry's byte arrays rather than copying them: the library
never mutates image bytes, and copying every picture would double the memory of an image-heavy
workbook to defend against a mutation nothing performs.

**Throws:** [`AuthoringError`](./errors.md#authoringerror) if the sheet anchors an image id this workbook has not registered,
which is what a sheet from *another* workbook looks like from here. Emitting a package with a
drawing pointing at media that was never registered is the silently-broken-image failure this
refuses to start.

#### `Workbook.importImages`

```ts
importImages(sheet: Worksheet, images: WorksheetImages): void;
```

Show `images` on `sheet`, a worksheet of this workbook, registering each picture's bytes here
and re-anchoring it against the id they land on. The counterpart to [`exportImages`](./workbook.md#workbookexportimages), and
the affordance that lets a picture cross workbooks at all: an [`AnchoredImage`](./images.md#anchoredimage)'s `imageId`
indexes one workbook's registry and means nothing in the next, so a raw anchor moved between
workbooks points at media that does not exist there.

The sheet's existing pictures are replaced, not appended to, so `importImages` is a transfer
rather than an accumulation: the same direction [`Worksheet.model`](./worksheet.md#worksheetmodel) assignment goes, and
what makes re-importing a sheet's own export leave it unchanged. An import whose `background` is
absent clears the destination's background for the same reason.

Registration is content-addressed: a picture whose bytes are already here is re-used at its
existing id rather than stored twice, so importing the same sheet repeatedly, or twenty sheets
sharing one logo, costs one media part.

#### `Workbook.definedNames`

```ts
get definedNames(): readonly DefinedName[];
```

The workbook's defined names, in the order they were registered.

#### `Workbook.defineName`

```ts
defineName(definedName: DefinedName): void;
```

Register a defined name on the workbook.

**Throws:** [`AuthoringError`](./errors.md#authoringerror) if the name is empty, or if a [`DefinedName.scope`](./workbook.md#definedname) is given that names no
existing worksheet. A scoped name must target a sheet that is already part of the workbook.

#### `Workbook.addWorksheet`

```ts
addWorksheet(name: string, options: AddWorksheetOptions = {}): Worksheet;
```

Create a worksheet and append it to the workbook.

**Throws:** [`AuthoringError`](./errors.md#authoringerror) if the name is empty, too long, contains a forbidden character,
or collides (case-insensitively) with an existing sheet.

#### `Workbook.getWorksheet`

```ts
getWorksheet(nameOrId: string | number): Worksheet | undefined;
```

Look up a worksheet by name (case-insensitive) or by numeric id.

#### `Workbook.requireWorksheet`

```ts
requireWorksheet(nameOrId: string | number): Worksheet;
```

[`getWorksheet`](./workbook.md#workbookgetworksheet), for a caller who knows the sheet is there: the miss throws instead of
returning `undefined`, and the message names every sheet the workbook does have.

The partial lookup is the right primitive for asking *whether* a sheet exists, and the wrong
one for reaching a sheet a template is expected to carry: `undefined` flows on into a `?.`
chain and fails several steps later with nothing left to say about which name was missing.
That listing is the whole point: a lookup miss is a typo, a stale template or a renamed tab,
and all three are answered by seeing the real names.

**Throws:** [`AuthoringError`](./errors.md#authoringerror) if no worksheet has that name (case-insensitive) or numeric id.

---

### `WorkbookProperties`

<sub>interface</sub>

Document-level metadata: what Excel's File ▸ Info panel shows. Mostly the package's core
properties (`docProps/core.xml`); `company` is the exception and lives in the extended part,
because that is where OOXML puts it. One interface either way: which part a field lands in is
the format's business, not the caller's.

```ts
interface WorkbookProperties {
  /** The document's title (`dc:title`), as Excel's File ▸ Info shows it. */
  title?: string;
  creator?: string;
  lastModifiedBy?: string;
  created?: Date;
  modified?: Date;
  /**
   * `Company` in the extended properties (`docProps/app.xml`), not the core ones. It is the only
   * field here that does not live beside the others, because OOXML puts it in the other part.
   */
  company?: string;
}
```

---

### `WorkbookView`

<sub>interface</sub>

The workbook's saved window state: OOXML's `<workbookView>`, the single entry of `<bookViews>`.

This is the rect a consumer restores the document window to, and the layout every pane geometry is
computed against: a frozen split is positioned within it. `activeTab` names the sheet whose tab is
selected on open.

The position and size are in twips (1/20 of a point), Excel's window unit. A slightly negative
`x`/`y` is normal and is what Excel itself writes: a maximised window's frame sits just outside the
work area.

```ts
interface WorkbookView {
  /** Left edge of the document window, in twips. */
  x: number;
  /** Top edge of the document window, in twips. */
  y: number;
  /** Window width, in twips. */
  width: number;
  /** Window height, in twips. */
  height: number;
  /** 0-based index into {@link Workbook.worksheets} of the sheet selected on open. */
  activeTab: number;
  /** Window visibility; omit for a normally visible window. */
  visibility?: Visibility;
  /** Whether the document window opens minimised; omit for a restored window. */
  minimized?: boolean;
}
```
