// The workbook: the top of the model and the entry point of the public API.
//
// It owns its worksheets and the document-level properties. Sheet identity follows
// Excel's rules: names are unique case-insensitively, bounded in length, and free of
// the characters Excel forbids. An invalid book therefore cannot be constructed in the
// first place, rather than failing only at write time.

import {type CustomUiDocument, isCustomUiRelType, parseCustomUi} from '../customui/index.ts';
import {AuthoringError, quoted} from '../errors.ts';
import {
  // Imported for the `{@link}` targets in the accessor docs below: the doc comments explain the
  // structural splices by pointing at the functions that perform them. tsc counts a `{@link}` as a
  // use and oxlint does not, so `noUnusedLocals` is satisfied and the linter still has to be told.
  // Upstream at oxc-project/oxc#11639, open since 2025-06-12 (a fix was drafted in #13989 and not
  // merged); do not file another. Nor is this oxlint diverging from the linter it replaced:
  // typescript-eslint reports the same and closed its own reports as not-planned, referring people
  // to `jsdoc/no-undefined-types`. `noUnusedLocals` is the only checker in this toolchain that
  // resolves a doc link. When #11639 lands, `--report-unused-disable-directives` retires these four.
  // oxlint-disable-next-line eslint/no-unused-vars
  addVbaReference,
  // oxlint-disable-next-line eslint/no-unused-vars
  removeVbaModule,
  type VbaLibraryReference,
  type VbaProject,
  type VbaProjectSignature,
} from '../vba/index.ts';
import {commentThreadGuid, type Person} from './comment-thread.ts';
import type {DateEpoch} from './date.ts';
import {type SheetSplice, spliceFormula} from './formula-references.ts';
import {pictureProperties, type WorkbookImage, type WorksheetImages} from './image.ts';
import {INTERNAL} from './internal.ts';
import {INVALID_SHEET_NAME_CHARS, MAX_SHEET_NAME_LENGTH} from './limits.ts';
import type {PreservedPart, PreservedRootReference} from './preserved.ts';
import type {Color, Font} from './style.ts';
import {checkTableStyle, type TableStyle} from './table-style.ts';
import {
  OFFICE_BODY_FACE,
  // Imported for a `{@link}` target, as the vba block above is.
  // oxlint-disable-next-line eslint/no-unused-vars
  THEME_COLOR_SLOTS,
  type ThemeColorScheme,
  type ThemeFontScheme,
  type ThemeOverrides,
} from './theme.ts';
import {WorkbookMedia} from './workbook-media.ts';
import type {WorkbookProtection} from './workbook-protection.ts';
import {type NamedCellStyle, type TableStyleTable, WorkbookStyleTables} from './workbook-styles.ts';
import {type DeclaredThemeSchemes, WorkbookTheme} from './workbook-theme.ts';
import {WorkbookVbaProject} from './workbook-vba.ts';
import {
  refuseImagesBesideKeptDrawing,
  type Visibility,
  Worksheet,
  type WorksheetState,
} from './worksheet.ts';

/**
 * A workbook-level reference to package content the model does not model: a pivot cache
 * (`pivotCacheDefinition`) or a slicer cache (`slicerCache`), preserved verbatim across a round-trip
 * instead of being dropped. `relType` is the workbook relationship Type URI to re-emit; `entryPath`
 * is the part it points at; `parts` is the transitive closure that reference reaches (the entry
 * included). `pivotCacheId` carries the `<pivotCache cacheId>` a pivot cache is registered under in
 * the workbook's `<pivotCaches>`, so the wiring a pivot table resolves its cache through is re-emitted
 * too; it is absent for a slicer cache, which the workbook lists in an extension block instead.
 * `externalReferenceIndex` is the 0-based position of an `externalLink` within the workbook's
 * `<externalReferences>`, the `[n]` a formula or defined name resolves an external cell through, so
 * the block is re-emitted in the original order and every `[n]` still points at the same linked
 * workbook; it is absent for a pivot/slicer cache.
 */
export interface PreservedWorkbookReference {
  readonly relType: string;
  readonly entryPath: string;
  readonly parts: readonly PreservedPart[];
  readonly pivotCacheId?: string;
  readonly externalReferenceIndex?: number;
}

/**
 * The workbook's theme part, captured verbatim from a source package: the `<clrScheme>`,
 * `<fontScheme>`, and `<fmtScheme>` every `theme="n"` colour reference and every `scheme="major|minor"`
 * font in the file resolves against.
 *
 * Held opaquely, like {@link Workbook.restoreDifferentialStyles}'s `<dxf>` fragments: the model does
 * not interpret the theme, it only refuses to destroy it. `entryPath` is where the source package held
 * the part (reached through the workbook's `.../theme` relationship, so not necessarily
 * `xl/theme/theme1.xml`), and `parts` is the transitive closure it reaches, the entry included. The
 * closure matters because a theme can carry its own relationships: a picture used as a themed fill is
 * wired by an `r:embed` into the theme's rels part, and re-emitting the theme without it would leave
 * that reference dangling.
 */
export interface PreservedTheme {
  readonly entryPath: string;
  readonly parts: readonly PreservedPart[];
}

/**
 * The workbook's saved window state: OOXML's `<workbookView>`, the single entry of `<bookViews>`.
 *
 * This is the rect a consumer restores the document window to, and the layout every pane geometry is
 * computed against: a frozen split is positioned within it. `activeTab` names the sheet whose tab is
 * selected on open.
 *
 * The position and size are in twips (1/20 of a point), Excel's window unit. A slightly negative
 * `x`/`y` is normal and is what Excel itself writes: a maximised window's frame sits just outside the
 * work area.
 */
export interface WorkbookView {
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

/**
 * The window geometry a workbook starts from: the values desktop Excel writes for its own default
 * window.
 *
 * A default is emitted rather than the element left out because Excel writes `<bookViews>` into every
 * file it saves and consumers lay panes out against that rect. With no view at all the frozen-pane
 * split is computed against an uninitialised window, and the frozen region can stay unpainted until
 * some later event forces a relayout.
 */
export const DEFAULT_WORKBOOK_VIEW = {
  x: -110,
  y: -110,
  width: 19420,
  height: 12220,
  activeTab: 0,
} as const satisfies WorkbookView;

/**
 * Document-level metadata: what Excel's File ▸ Info panel shows. Mostly the package's core
 * properties (`docProps/core.xml`); `company` is the exception and lives in the extended part,
 * because that is where OOXML puts it. One interface either way: which part a field lands in is
 * the format's business, not the caller's.
 */
export interface WorkbookProperties {
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

/**
 * A named reference in the workbook, one of the entries Excel surfaces in its Name Manager. A name
 * maps an identifier to a formula (`refersTo`), most often a cell range like `Sheet1!$A$1:$B$2` but
 * possibly any formula. A name is global to the workbook unless it names a sheet in {@link scope},
 * which restricts it to that sheet and lets another sheet reuse the same name independently.
 */
export interface DefinedName {
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

export interface AddWorksheetOptions {
  readonly state?: WorksheetState['state'];
}

/** A picture registered on the workbook, ready to be anchored to a worksheet. */
export interface AddImageOptions {
  /** The image bytes. */
  readonly buffer: Uint8Array;
  /** The file kind: `"png"`, `"jpeg"`/`"jpg"`, `"gif"`, … A leading dot or a URL query string is
   * tolerated and stripped; omit it entirely to infer the kind from the bytes' magic number. */
  readonly extension?: string;
}

export class Workbook {
  readonly properties: WorkbookProperties = {};

  /**
   * The workbook's window state: position, size, and the selected sheet. Always present (see
   * {@link DEFAULT_WORKBOOK_VIEW} for why it is defaulted rather than left unset) and always written.
   * Reading a file replaces it with that file's saved geometry, so a round-trip restores the window
   * the author left rather than stamping ours over it.
   */
  readonly view: WorkbookView = {...DEFAULT_WORKBOOK_VIEW};

  /**
   * Ask consuming spreadsheet apps to recalculate every formula when the file is opened, rather than
   * trusting the cached results stored with each formula cell. Set this when the producer cannot
   * compute formula results itself. This is the OOXML `fullCalcOnLoad` flag, off by default, so a
   * workbook whose cached results are authoritative stays unmarked.
   */
  fullCalcOnLoad = false;

  /**
   * Which date system this workbook's serials count in: `1900` (the Windows default) or `1904` (the
   * `date1904` flag of `<workbookPr>`, Excel for Macintosh's original). It governs every conversion
   * between a `Date` and the number a cell actually stores, in both directions, so setting it after
   * cells hold dates changes what those cells mean rather than converting them.
   *
   * Read from the file and written back, because dropping it is not a cosmetic loss: the serials stay
   * as they were and the consumer re-reads them under the other system, so the workbook silently
   * changes meaning by four years and a day.
   */
  dateEpoch: DateEpoch = 1900;

  /**
   * The workbook's VBA identity (`<workbookPr codeName>`), the name a macro means by `ThisWorkbook`.
   * Undefined for a workbook with no VBA project, which is what Excel writes for one.
   *
   * Preserved rather than modeled: nothing here reads it, but a `.xlsm` whose code name is dropped on
   * a round trip has had the binding between its macros and its document cut.
   */
  codeName?: string;

  /**
   * Workbook-level structure/window protection: the OOXML `<workbookProtection>` element. Absent by
   * default (an unprotected workbook). Set it to lock the workbook shell, or leave it as read from a
   * file so a protected workbook stays locked across a passthrough save rather than being silently
   * unlocked. Distinct from a worksheet's own `protect()`, which guards a single sheet's cells.
   */
  protection: WorkbookProtection | undefined = undefined;

  readonly #worksheets: Worksheet[] = [];
  #nextSheetId = 1;

  // The picture registry and its content index. See WorkbookMedia.
  readonly #media = new WorkbookMedia();

  readonly #definedNames: DefinedName[] = [];

  // The style-tables slice: the `<dxfs>`, the named cell styles, the two colour lists and the
  // table-style definitions, all preserved rather than interpreted because each is the target of an
  // index held elsewhere in the file. See `workbook-styles.ts`.
  readonly #styles = new WorkbookStyleTables();

  // The theme slice: the preserved part, the decoded scheme and its cache, and what a caller
  // authored over them. The indexed palette is read off the styles slice on demand: it is styles
  // state, not theme state, and only colour resolution wants both. See `workbook-theme.ts`.
  readonly #theme = new WorkbookTheme(() => this.#styles.indexedPalette());

  // Workbook-level references to package content the model does not interpret (pivot caches, slicer
  // caches), captured verbatim on read so a round-trip re-emits them rather than dropping the pivots
  // and slicers they back. Empty for a workbook authored from scratch.
  readonly #preservedReferences: PreservedWorkbookReference[] = [];

  // Every row and column splice made to a sheet of this workbook, in order. A preserved part spells
  // references to the sheets in bytes the model does not read, so the writer replays these over it
  // (see `io/xlsx/preserved-splices.ts`). Replaying the whole list over the bytes as they were read is
  // sound because preserved parts only enter a workbook as it is read, before any splice can be made,
  // and a sheet's name, which each entry is keyed by, never changes.
  readonly #splices: SheetSplice[] = [];

  // The threaded-comment identity registry (`xl/persons/person.xml`), the workbook-level table every
  // message resolves its author through and every @mention its target. Keyed by person id. See
  // `restorePersons` for why nothing else will do. Empty for a workbook with no threaded comments.
  readonly #persons = new Map<string, Person>();

  /** The worksheets in insertion order. */
  get worksheets(): readonly Worksheet[] {
    return this.#worksheets;
  }

  /**
   * The 0-based index of the active sheet: {@link WorkbookView.activeTab} resolved against the sheets
   * that actually exist. Exactly one sheet is always active: an out-of-range tab (a caller's stale
   * index, or a file whose sheet was removed after the view was saved) falls back to the first sheet
   * rather than to none, because a package where no sheet is selected gives the consumer no view to
   * initialise on open.
   */
  get activeTabIndex(): number {
    const {activeTab} = this.view;
    return Number.isInteger(activeTab) && activeTab >= 0 && activeTab < this.#worksheets.length
      ? activeTab
      : 0;
  }

  /** The workbook-level preserved references, in the order they were read. */
  get preservedReferences(): readonly PreservedWorkbookReference[] {
    return this.#preservedReferences;
  }

  // Package-root references to unmodeled content wired from `_rels/.rels` (the ribbon customUI parts,
  // custom document properties, a thumbnail), captured verbatim on read so a round-trip re-declares
  // them in the regenerated root rels rather than dropping them. Empty for a workbook authored from
  // scratch.
  readonly #preservedRootReferences: PreservedRootReference[] = [];

  /** The package-root preserved references, in the order they were read. */
  get preservedRootReferences(): readonly PreservedRootReference[] {
    return this.#preservedRootReferences;
  }

  // Lazily-parsed customUI ribbon view. `#customUiParsed` distinguishes "not yet parsed" from a genuine
  // "no ribbon parts" (an empty array), so a ribbon-free workbook is not re-scanned on every access.
  #customUiParsed = false;
  #customUI: readonly CustomUiDocument[] = [];

  /**
   * The ribbon customisations decoded from this workbook's `customUI` parts: `customUI.xml` (Office
   * 2007) and/or `customUI14.xml` (Office 2010+), in the order their root relationships were read. Each
   * {@link CustomUiDocument} is tagged with its dialect and exposes the parsed `<ribbon>` tree. Empty
   * for a workbook that customises no ribbon.
   *
   * This is a **read-only view** over parts the writer already round-trips verbatim. Mutating the
   * returned objects changes nothing on write; the original `customUI` XML is re-emitted byte-for-byte
   * regardless. Parsed lazily on first access and memoised.
   *
   * @throws {CustomUiParseError} if a `customUI` part is present but its XML is malformed.
   */
  get customUI(): readonly CustomUiDocument[] {
    if (!this.#customUiParsed) {
      this.#customUI = this.#preservedRootReferences
        .filter((ref) => isCustomUiRelType(ref.relType))
        .flatMap((ref) => {
          const bytes = ref.parts.find((part) => part.path === ref.entryPath)?.bytes;
          return bytes === undefined ? [] : [parseCustomUi(bytes)];
        });
      this.#customUiParsed = true;
    }
    return this.#customUI;
  }

  // The macro-project slice, held by reference to the preserved list so an attach or replace is
  // visible to the writer. The accessors below are the public surface; see `workbook-vba.ts`.
  readonly #vba = new WorkbookVbaProject(this.#preservedReferences);

  /**
   * The VBA project decoded from this workbook's preserved `vbaProject.bin`, or `undefined` for a
   * workbook with no macros. This is a **read-only view** over the bytes the writer already round-trips
   * verbatim. Mutating the returned object changes nothing on write; the original macro blob is
   * re-emitted byte-for-byte regardless. Parsed lazily on first access and memoised.
   *
   * @throws {VbaParseError} if a macro project is present but its `vbaProject.bin` is malformed.
   */
  get vbaProject(): VbaProject | undefined {
    return this.#vba.project;
  }

  /**
   * The raw `vbaProject.bin` bytes attached to this workbook (the exact macro blob the writer will
   * embed), or `undefined` for a workbook with no macros. The getter returns a defensive copy, so
   * mutating it changes nothing on write.
   *
   * Assigning bytes attaches (or replaces) the macro project: the written package becomes
   * macro-enabled and re-embeds these bytes verbatim. The bytes must be a well-formed VBA container
   * (a CFB holding a `dir` stream); a malformed blob is rejected with {@link VbaParseError} rather
   * than written out to produce a package Excel would flag for repair. This is the attach-blob path:
   * copy a project between workbooks with `dst.vbaProjectBytes = src.vbaProjectBytes`, or import a
   * `.bin` produced by another tool. Assigning `undefined` removes the project, reverting the workbook
   * to a plain (non-macro) package.
   *
   * Replacing or removing the project also drops any digital signature the previous blob carried,
   * because a signature over the old bytes cannot validate new ones, so the result never advertises
   * a broken signature.
   */
  get vbaProjectBytes(): Uint8Array | undefined {
    return this.#vba.bytes;
  }

  set vbaProjectBytes(bytes: Uint8Array | undefined) {
    this.#vba.bytes = bytes;
  }

  /**
   * Whether this workbook's VBA project carries a digital signature: `true` if any signature part is
   * attached, `false` for an unsigned project or a workbook with no macros.
   *
   * This reflects the **presence** of a signature blob, not its cryptographic validity: the library
   * neither parses the PKCS#7/CMS structure nor validates the certificate chain or signer. A `true`
   * here means "a signature is attached," never "this signature is valid." Replacing or editing the
   * project drops its signatures (a signature over the old bytes cannot validate new ones), so this
   * reads `false` again after {@link vbaProjectBytes}, {@link removeVbaModule}, or
   * {@link addVbaReference} mutates the project. See {@link vbaProjectSignatures} for the raw bytes and
   * which generation(s) are present.
   */
  get vbaProjectSigned(): boolean {
    return this.#vba.signatures.length > 0;
  }

  /**
   * The digital signatures attached to this workbook's VBA project, in the order their relationships
   * are wired off `vbaProject.bin`. Up to three generations (legacy, agile, V3) can coexist over the
   * same project bytes. Empty for an unsigned project or a workbook with no macros.
   *
   * Each entry's `bytes` are the raw signature blob passed through verbatim; the library does not parse
   * or verify them (see {@link vbaProjectSigned} on presence-vs-validity). Hand a blob to an external
   * verifier if you need cryptographic validation; that is deliberately out of this library's scope.
   */
  get vbaProjectSignatures(): readonly VbaProjectSignature[] {
    return this.#vba.signatures;
  }

  /**
   * Remove a standard module from this workbook's existing macro project, in place: a structural splice
   * that leaves every remaining module's compiled p-code untouched (see {@link removeVbaModule}).
   * Replacing the project also drops a stale signature, as {@link vbaProjectBytes} does.
   *
   * Only `procedural` and `class` modules can be removed this way. See {@link removeVbaModule} for why.
   * To author or edit module *source* (which needs real compiled p-code), use the offline
   * `tools/vba-compiler`, then attach its output via {@link vbaProjectBytes}.
   *
   * @throws {VbaAuthorError} if the workbook has no macro project, or `name` is not in the project,
   *   or names a `document`/`designer` module.
   * @throws {VbaParseError} if the attached `vbaProject.bin` is malformed.
   */
  removeVbaModule(name: string): void {
    this.#vba.removeModule(name);
  }

  /**
   * Add a registered (COM type-library) reference to this workbook's existing macro project, in place.
   * Every existing module, reference, and host-info record rides through unchanged (see
   * {@link addVbaReference}). Replacing the project also drops a stale signature, as
   * {@link vbaProjectBytes} does.
   *
   * @throws {VbaAuthorError} if the workbook has no macro project, or any field of `ref` is invalid
   *   (see {@link VbaLibraryReference}).
   * @throws {VbaParseError} if the attached `vbaProject.bin` is malformed.
   */
  addVbaReference(ref: VbaLibraryReference): void {
    this.#vba.addReference(ref);
  }

  /** The preserved differential-style (`<dxfs>`) fragments, in index order. */
  get differentialStyles(): readonly string[] {
    return this.#styles.differentialStyles;
  }

  /** The preserved custom indexed-color palette, in index order; empty when the default palette rules. */
  get indexedColors(): readonly string[] {
    return this.#styles.indexedColors;
  }

  /** The preserved most-recently-used colour swatches, in order; empty when the file declared none. */
  get mruColors(): readonly string[] {
    return this.#styles.mruColors;
  }

  /** The preserved `<tableStyles>` block; `styles` is empty when the file declared no custom style. */
  get tableStyles(): TableStyleTable {
    return this.#styles.tableStyles;
  }

  /**
   * Register a custom table style: a named look a table applies to itself by putting that name in
   * {@link TableStyleInfo.name}, exactly as it would name one of Excel's built-in gallery styles.
   *
   * ```ts
   * workbook.addTableStyle({
   *   name: 'Harbour',
   *   elements: {
   *     wholeTable: {border: {top: {style: 'thin'}, bottom: {style: 'thin'}}},
   *     headerRow: {font: {bold: true, color: {argb: 'FFFFFFFF'}},
   *                 fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FFBB2649'}}},
   *     firstRowStripe: {fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FFF6E7EB'}}},
   *   },
   * });
   * sheet.addTable({name: 'Cargo', ref: 'A1:B3', columns, style: {name: 'Harbour'}});
   * ```
   *
   * Each element's formatting is interned into the workbook's shared differential-style table, so two
   * elements, or a conditional-formatting rule, that paint the same way share one entry.
   *
   * Registering a name a source file already defined **overrides** that definition rather than adding
   * a second one beside it.
   *
   * @throws {AuthoringError} if the name is empty, or an element carries a `size` outside the four
   *   stripe types. See {@link checkTableStyle} for why those are refused here rather than silently
   *   dropped.
   * @throws {RangeError} if a `size` is not a positive integer.
   */
  addTableStyle(style: TableStyle): void {
    checkTableStyle(style);
    this.#styles.addCustomTableStyle(style);
  }

  /** The table styles authored on this workbook, in registration order. */
  get customTableStyles(): readonly TableStyle[] {
    return this.#styles.customTableStyles;
  }

  /** The preserved theme part, or undefined when the workbook rides the library's default theme. */
  get themePart(): PreservedTheme | undefined {
    return this.#theme.part;
  }

  /**
   * Author the workbook's theme: any subset of the twelve colour-scheme slots, and either of the two
   * typefaces. Merges into what the workbook already has, so branding one accent leaves the other
   * eleven slots alone, and calling it twice accumulates.
   *
   * This is the workbook-wide palette. A cell that names a colour as `theme="4"`, which is what Excel
   * writes whenever a user picks from the theme row of the colour picker, follows `accent1` here, so
   * one call restyles every such cell, chart and table style at once. Colours are `RRGGBB`; a leading
   * `#` and an 8-hex ARGB are both accepted and reduced, and anything else throws rather than writing
   * a value Excel silently renders as flat black.
   *
   * What it does **not** touch: the theme's format scheme, meaning the gradient, line and effect
   * styles that give a theme its texture. Those ride through from the source theme (or the library's
   * default) untouched, because nobody hand-authors gradient stops from a spreadsheet API and
   * regenerating them would replace a designer's work with the Office default. For the same reason a
   * slot left unauthored keeps the source's own encoding, including the `<a:sysClr>` form Excel uses
   * for `dk1`/`lt1` so they follow the viewer's window colours.
   *
   * @throws {SyntaxError} if a colour is not 6 or 8 hexadecimal digits, with or without a leading `#`.
   *   Every colour is read before any is stored, so a refused call changes nothing.
   */
  setTheme(overrides: ThemeOverrides): void {
    this.#theme.author(overrides);
  }

  /**
   * The colour scheme every `theme="n"` reference in this workbook resolves against: anything
   * {@link setTheme} authored, over the preserved theme's `<a:clrScheme>`, over the Office default.
   *
   * Note the slot *order*: `theme="0"` is `lt1` and `theme="1"` is `dk1`, which is not the order the
   * slots appear in the theme part. See {@link THEME_COLOR_SLOTS}.
   *
   * Every slot reads back as six uppercase hexadecimal digits, whichever spelling {@link setTheme} was
   * given or the theme part used.
   */
  get themeColors(): ThemeColorScheme {
    return this.#theme.colors;
  }

  /** The theme's major (heading) and minor (body) typefaces, authored values over the source's. */
  get themeFonts(): ThemeFontScheme {
    return this.#theme.fonts;
  }

  // Font id 0 exactly as a source file declared it, before anything was authored over it. Held apart
  // from the authored layer because the two answer different questions and only their separation makes
  // the round-trip faithful: a file's font 0 must ride through untouched, while an authored one must
  // win. Undefined for a workbook built from scratch, or read from a package with no styles part.
  #declaredDefaultFont: Font | undefined;

  // The facets {@link setDefaultFont} has accumulated, merged in call order. Undefined until a caller
  // authors one, which is what tells {@link defaultFont} it may re-derive rather than pass through.
  #authoredDefaultFont: Font | undefined;

  /**
   * The default font as the source package declared it: font id 0 of its styles part, the face every
   * cell that names no font of its own renders in. `undefined` for a workbook authored from scratch or
   * read from a package carrying no styles part: nothing was declared, and the library does not
   * fabricate a declaration on the file's behalf.
   *
   * This is the *round-trip* surface. {@link defaultFont} is what the workbook actually renders in,
   * which is this once anything has been authored over it.
   */
  get declaredDefaultFont(): Font | undefined {
    return this.#declaredDefaultFont;
  }

  /**
   * Author the workbook's default font: the face, size and colour every cell with no font of its own
   * renders in, **empty cells included**. Merges into whatever the workbook already had, so
   * `setDefaultFont({size: 14})` keeps the resolved face and changes only the size, and calling it
   * twice accumulates. This is the one knob that reaches a cell no row or column default can: an
   * untouched cell in an unformatted column.
   *
   * It writes the styles part's font 0 and **nothing else**. In particular it does not rewrite the
   * theme's body typeface. The dependency runs the other way: with no default font authored, font 0
   * follows {@link themeFonts}'s minor face, so `setTheme({fonts: {minor}})` already reaches every
   * unstyled cell and needs no second call here. See {@link defaultFont} for the full chain.
   *
   * @throws {RangeError} if `size` is not a positive finite number, or `name` is empty. Both
   *   produce a styles part Excel renders from some other font without ever reporting why. Native
   *   rather than {@link AuthoringError}: one argument out of range is what `RangeError` is for.
   */
  setDefaultFont(font: Font): void {
    if (font.size !== undefined && !(Number.isFinite(font.size) && font.size > 0)) {
      throw new RangeError(`default font size must be a positive number, not ${font.size}`);
    }
    if (font.name !== undefined && font.name === '') {
      throw new RangeError('default font name cannot be empty');
    }
    this.#authoredDefaultFont = {...this.#authoredDefaultFont, ...font};
  }

  /**
   * The font every cell that names none of its own renders in, resolved and complete: what the writer
   * emits as font id 0. Never `undefined`: a workbook always renders in *some* face, and the chain
   * below always reaches one.
   *
   * ```
   * authored default font  >  authored theme body face  >  the source file's font 0  >  theme body face
   * ```
   *
   * The two authored levels outrank the file because authoring is an explicit act; between them
   * {@link setDefaultFont} wins on the face because it names font 0 outright while
   * {@link setTheme} names it only by implication. With **nothing** authored the file's own font 0
   * passes through verbatim. That is deliberate, because a producer resolves that face by script and
   * we do not: Excel writes `等线` as font 0 under a theme whose latin body face is `Calibri`, and
   * re-deriving would silently rewrite it.
   *
   * `family` and `scheme` describe the *theme's* body face, so they are carried exactly while the
   * resolved face still is that face and dropped when a caller names another, which is also what
   * Excel writes: a font 0 naming a non-theme face carries no `<scheme>` at all. Either may be stated
   * outright, in which case the caller's word stands.
   */
  get defaultFont(): Font {
    const declared = this.#declaredDefaultFont;
    const authored = this.#authoredDefaultFont;
    const authoredFace = this.#theme.authoredFonts.minor;
    if (declared !== undefined && authored === undefined && authoredFace === undefined) {
      return declared;
    }
    const bodyFace = this.themeFonts.minor ?? OFFICE_BODY_FACE;
    const face = authored?.name ?? authoredFace ?? declared?.name ?? bodyFace;
    // Size and colour are completed rather than merely merged: a font 0 that states neither is the
    // "missing default font" foreign readers warn about, so the emitted entry always carries both.
    const font: {-readonly [K in keyof Font]?: Font[K]} = {
      size: 11,
      color: {theme: 1},
      ...declared,
      ...authored,
      name: face,
    };
    const followsTheme = face === bodyFace;
    if (authored?.family === undefined) {
      if (followsTheme) font.family = 2;
      else delete font.family;
    }
    if (authored?.scheme === undefined) {
      if (followsTheme) font.scheme = 'minor';
      else delete font.scheme;
    }
    return font;
  }

  /**
   * The colour slots and typefaces {@link setTheme} authored on this workbook, or `undefined` when
   * none were, in which case the source theme (or the writer's default) rides through untouched.
   *
   * A serializer composes these *over* the existing part rather than generating one from scratch, so
   * a preserved theme keeps its format scheme, its unauthored slots' exact encoding, and the
   * relationships it carries. Its colours are already reduced to the six uppercase hexadecimal digits
   * a slot holds.
   */
  get themeOverrides(): ThemeOverrides | undefined {
    return this.#theme.overrides;
  }

  /**
   * Resolve a colour reference to a concrete 8-hex ARGB string, or `undefined` when it does not
   * resolve to a fixed colour: an `auto` colour, one of the two system indexed colours, or a theme
   * slot this workbook's scheme does not declare.
   *
   * This is a *derived* view, not a rewrite: the {@link Color} stays exactly as its file encoded it,
   * so a round-trip re-emits `theme="4" tint="0.4"` rather than a literal ARGB. Resolving into the
   * model would sever every cell's link to the theme, so recolouring the workbook would stop working,
   * and would inflate the styles table with one distinct colour per shade.
   *
   * A `theme` reference resolves through {@link themeColors}; an `indexed` one through the workbook's
   * custom `<indexedColors>` palette when it declares one, else the built-in legacy palette. A `tint`
   * is applied last.
   */
  resolveColor(color: Color): string | undefined {
    return this.#theme.resolveColor(color);
  }

  /** The named cell styles, in index order (index 0 is Normal); empty when only the default exists. */
  get namedStyles(): readonly NamedCellStyle[] {
    return this.#styles.namedStyles;
  }

  /**
   * Register an identity a threaded comment can name: an author, or someone `@mentioned` in a message.
   * A message reaches it by {@link Comment.personId}, a mention by {@link Mention.personId}.
   *
   * Keyed by {@link Person.id} alone, so registering the same id twice replaces the entry rather than
   * adding a second: the id is the identity. Registering the same human twice under *different* ids is
   * legitimate and is what Excel itself does. See {@link restorePersons}. The id is normalised to the
   * brace-wrapped upper-case GUID form the format requires, so a `crypto.randomUUID()` is accepted as-is.
   *
   * @throws {SyntaxError} if the id is not a GUID.
   */
  addPerson(person: Person): void {
    const id = commentThreadGuid(person.id, 'a person id');
    this.#persons.set(id, {...person, id});
  }

  /**
   * The registered threaded-comment identities, in the order they were read. That order carries no
   * meaning, because Excel re-sorts the registry by person id when it saves, so nothing may depend
   * on it.
   */
  get persons(): readonly Person[] {
    return [...this.#persons.values()];
  }

  /** Look up a registered identity by its {@link Person.id}, or `undefined` if the registry has none. */
  getPerson(id: string): Person | undefined {
    return this.#persons.get(id);
  }

  /**
   * Register a picture on the workbook and return its numeric id. Pass the id to
   * {@link Worksheet.addImage} to anchor the picture to a sheet; the same id may be anchored on any
   * number of sheets and positions, and the bytes are still stored only once.
   */
  addImage(options: AddImageOptions): number {
    return this.#media.register(options.extension, options.buffer);
  }

  /** The registered images, indexed by the id {@link addImage} returned. */
  get media(): readonly WorkbookImage[] {
    return this.#media.all;
  }

  /** Look up a registered image by its id, or `undefined` if no image carries that id. */
  getImage(id: number): WorkbookImage | undefined {
    return this.#media.get(id);
  }

  /**
   * Every picture `sheet` shows, resolved out of this workbook's media registry into the
   * workbook-independent form {@link importImages} consumes. `sheet` must belong to this workbook:
   * that is whose registry its image ids index.
   *
   * This is the attached-part half of a sheet copy, and it is deliberately a separate call from
   * {@link Worksheet.model}: a model is a serialisable value, an image is bytes on the workbook, and
   * ADR-0005 keeps them apart. Carrying a sheet whole is therefore the two of them together:
   *
   * ```ts
   * destination.model = source.model;
   * destinationWorkbook.importImages(destination, sourceWorkbook.exportImages(source));
   * ```
   *
   * The exported pictures share the registry's byte arrays rather than copying them: the library
   * never mutates image bytes, and copying every picture would double the memory of an image-heavy
   * workbook to defend against a mutation nothing performs.
   *
   * @throws {AuthoringError} if the sheet anchors an image id this workbook has not registered,
   *   which is what a sheet from *another* workbook looks like from here. Emitting a package with a
   *   drawing pointing at media that was never registered is the silently-broken-image failure this
   *   refuses to start.
   */
  exportImages(sheet: Worksheet): WorksheetImages {
    const background = sheet.backgroundImageId;
    return {
      anchored: sheet.images.map((picture) => ({
        image: this.#media.require(picture.imageId, sheet.name),
        anchor: picture.anchor,
        ...pictureProperties(picture),
      })),
      background:
        background === undefined ? undefined : this.#media.require(background, sheet.name),
    };
  }

  /**
   * Show `images` on `sheet`, a worksheet of this workbook, registering each picture's bytes here
   * and re-anchoring it against the id they land on. The counterpart to {@link exportImages}, and
   * the affordance that lets a picture cross workbooks at all: an {@link AnchoredImage}'s `imageId`
   * indexes one workbook's registry and means nothing in the next, so a raw anchor moved between
   * workbooks points at media that does not exist there.
   *
   * The sheet's existing pictures are replaced, not appended to, so `importImages` is a transfer
   * rather than an accumulation: the same direction {@link Worksheet.model} assignment goes, and
   * what makes re-importing a sheet's own export leave it unchanged. An import whose `background` is
   * absent clears the destination's background for the same reason.
   *
   * Registration is content-addressed: a picture whose bytes are already here is re-used at its
   * existing id rather than stored twice, so importing the same sheet repeatedly, or twenty sheets
   * sharing one logo, costs one media part.
   */
  importImages(sheet: Worksheet, images: WorksheetImages): void {
    // Asked before anything is cleared: `addImageAnchor` would refuse too, but only after this had
    // already removed the sheet's background, leaving half an import applied.
    if (images.anchored.length > 0) refuseImagesBesideKeptDrawing(sheet);
    for (const id of new Set(sheet.images.map((image) => image.imageId))) sheet.removeImage(id);
    sheet.removeBackgroundImage();
    for (const picture of images.anchored) {
      sheet.addImageAnchor(
        this.#media.registerExisting(picture.image),
        picture.anchor,
        pictureProperties(picture),
      );
    }
    if (images.background !== undefined) {
      sheet.addBackgroundImage(this.#media.registerExisting(images.background));
    }
  }

  /** The workbook's defined names, in the order they were registered. */
  get definedNames(): readonly DefinedName[] {
    return this.#definedNames;
  }

  /**
   * Register a defined name on the workbook.
   *
   * @throws {AuthoringError} if the name is empty, or if a {@link DefinedName.scope} is given that names no
   *   existing worksheet. A scoped name must target a sheet that is already part of the workbook.
   */
  defineName(definedName: DefinedName): void {
    if (definedName.name.length === 0) {
      throw new AuthoringError('a defined name cannot be empty');
    }
    if (definedName.scope !== undefined && this.getWorksheet(definedName.scope) === undefined) {
      throw new AuthoringError(
        `defined name ${quoted(definedName.name)} is scoped to unknown worksheet ${quoted(definedName.scope)}`,
      );
    }
    this.#definedNames.push(definedName);
  }

  /**
   * Create a worksheet and append it to the workbook.
   *
   * @throws {AuthoringError} if the name is empty, too long, contains a forbidden character,
   *   or collides (case-insensitively) with an existing sheet.
   */
  addWorksheet(name: string, options: AddWorksheetOptions = {}): Worksheet {
    this.#assertValidSheetName(name);
    const sheet = new Worksheet(name, this.#nextSheetId++, options.state ?? 'visible');
    sheet[INTERNAL].setFormulaHost((edit) => {
      this.#spliceFormulasBeyond(sheet, edit);
    });
    this.#worksheets.push(sheet);
    return sheet;
  }

  // A splice of one sheet moves the references to it everywhere else a formula lives: every other sheet,
  // each defined name, and the preserved parts the writer replays the journal over. A sheet-scoped name reads an unqualified reference against its own sheet, as
  // a formula on that sheet does; a workbook-scoped one has no sheet of its own.
  #spliceFormulasBeyond(spliced: Worksheet, edit: SheetSplice): void {
    this.#splices.push(edit);
    for (const sheet of this.#worksheets) {
      if (sheet !== spliced) sheet[INTERNAL].spliceFormulas(edit);
    }
    for (const [index, definedName] of this.#definedNames.entries()) {
      const refersTo = spliceFormula(definedName.refersTo, definedName.scope, edit);
      if (refersTo !== definedName.refersTo) this.#definedNames[index] = {...definedName, refersTo};
    }
  }

  /** Look up a worksheet by name (case-insensitive) or by numeric id. */
  getWorksheet(nameOrId: string | number): Worksheet | undefined {
    if (typeof nameOrId === 'number') {
      return this.#worksheets.find((sheet) => sheet.id === nameOrId);
    }
    const target = nameOrId.toLowerCase();
    return this.#worksheets.find((sheet) => sheet.name.toLowerCase() === target);
  }

  /**
   * {@link getWorksheet}, for a caller who knows the sheet is there: the miss throws instead of
   * returning `undefined`, and the message names every sheet the workbook does have.
   *
   * The partial lookup is the right primitive for asking *whether* a sheet exists, and the wrong
   * one for reaching a sheet a template is expected to carry: `undefined` flows on into a `?.`
   * chain and fails several steps later with nothing left to say about which name was missing.
   * That listing is the whole point: a lookup miss is a typo, a stale template or a renamed tab,
   * and all three are answered by seeing the real names.
   *
   * @throws {AuthoringError} if no worksheet has that name (case-insensitive) or numeric id.
   */
  requireWorksheet(nameOrId: string | number): Worksheet {
    const sheet = this.getWorksheet(nameOrId);
    if (sheet !== undefined) return sheet;
    const wanted = typeof nameOrId === 'number' ? `id ${nameOrId}` : quoted(nameOrId);
    if (this.#worksheets.length === 0) {
      throw new AuthoringError(`no worksheet ${wanted}: this workbook has no worksheets`);
    }
    const have = this.#worksheets.map((sheet) => quoted(sheet.name)).join(', ');
    throw new AuthoringError(`no worksheet ${wanted}; this workbook has ${have}`);
  }

  #assertValidSheetName(name: string): void {
    if (name.length === 0) {
      throw new AuthoringError('worksheet name cannot be empty');
    }
    if (name.length > MAX_SHEET_NAME_LENGTH) {
      throw new AuthoringError(
        `worksheet name ${quoted(name)} exceeds the ${MAX_SHEET_NAME_LENGTH}-character limit`,
      );
    }
    if (INVALID_SHEET_NAME_CHARS.test(name)) {
      throw new AuthoringError(
        `worksheet name ${quoted(name)} contains a character Excel forbids (* ? : \\ / [ ])`,
      );
    }
    if (name.startsWith("'") || name.endsWith("'")) {
      throw new AuthoringError(
        `worksheet name ${quoted(name)} cannot start or end with an apostrophe`,
      );
    }
    if (this.getWorksheet(name) !== undefined) {
      throw new AuthoringError(
        `a worksheet named ${quoted(name)} already exists (names are case-insensitive)`,
      );
    }
  }

  /**
   * The codec's channel into this workbook. See `core/internal.ts` for why these are not public
   * methods. Declared last so every private field it closes over is already in scope.
   */
  readonly [INTERNAL]: WorkbookInternals = {
    addPreservedReference: (reference) => {
      this.#preservedReferences.push(reference);
    },
    addPreservedRootReference: (reference) => {
      this.#preservedRootReferences.push(reference);
    },
    restoreDifferentialStyles: (fragments) => {
      this.#styles.restoreDifferentialStyles(fragments);
    },
    adoptDifferentialStyle: (fragment) => this.#styles.adoptDifferentialStyle(fragment),
    restoreIndexedColors: (fragments) => {
      this.#styles.restoreIndexedColors(fragments);
    },
    restoreMruColors: (fragments) => {
      this.#styles.restoreMruColors(fragments);
    },
    restoreTableStyles: (table) => {
      this.#styles.restoreTableStyles(table);
    },
    restoreThemePart: (theme, declared) => {
      this.#theme.restorePart(theme, declared);
    },
    restoreNamedStyles: (styles) => {
      this.#styles.restoreNamedStyles(styles);
    },
    restoreDefaultFont: (font) => {
      this.#declaredDefaultFont = font;
    },
    restorePersons: (persons) => {
      this.#persons.clear();
      for (const person of persons) this.#persons.set(person.id, person);
    },
    splices: () => this.#splices,
  };
}

/**
 * What a codec may do to a `Workbook` that an author may not: reinstate state read from a file, in
 * the finished form the file stated it. Reached as `workbook[INTERNAL]`; see `core/internal.ts`.
 *
 * Every operation replaces what it restores rather than merging, because a reader states a table
 * whole: a half-restored `<dxfs>` would leave existing `dxfId` references pointing into a mix of
 * two files.
 */
export interface WorkbookInternals {
  /** Record a workbook-level preserved reference (a pivot or slicer cache) read from a file. */
  addPreservedReference(reference: PreservedWorkbookReference): void;

  /** Record a package-root preserved reference (a customUI ribbon part, custom props) read from a file. */
  addPreservedRootReference(reference: PreservedRootReference): void;

  /**
   * Reinstate the differential-style (`<dxfs>`) table read from a file: the deserialization
   * counterpart the writer re-emits verbatim. Each entry is one `<dxf>…</dxf>` fragment, preserved as
   * opaque XML so a conditional-formatting rule's `dxfId` (an index into this table) stays valid on
   * re-write.
   */
  restoreDifferentialStyles(fragments: readonly string[]): void;

  /**
   * The index of a `<dxf>` fragment a sheet carried inline, found among the differential styles or
   * appended after them, so a conditional format read from the extension form holds a `dxfId` like any
   * other rule.
   */
  adoptDifferentialStyle(fragment: string): number;

  /**
   * Reinstate the custom indexed-color palette (`<colors><indexedColors>`) read from a file, each
   * entry a verbatim `<rgbColor rgb="…"/>` fragment, so a colour referenced by `indexed="…"` keeps
   * its intended RGB on re-write instead of the palette being dropped and the colour shifting to a
   * default-palette entry.
   */
  restoreIndexedColors(fragments: readonly string[]): void;

  /**
   * Reinstate the most-recently-used colour swatches (`<colors><mruColors>`) read from a file, each
   * entry a verbatim `<color rgb="…"/>` fragment. This is the "Recent Colors" row a spreadsheet
   * application offers, which is the author's own working set rather than anything the model
   * interprets.
   */
  restoreMruColors(fragments: readonly string[]): void;

  /**
   * Reinstate the custom table-style definitions (`<tableStyles>`) read from a file (see
   * {@link TableStyleTable}), so a table whose `styleName` names a custom style still resolves to a
   * real definition on re-write instead of dangling, and the file's nominated default table/pivot
   * styles survive.
   */
  restoreTableStyles(table: TableStyleTable): void;

  /**
   * Reinstate the theme part read from a file: opaque preserved XML plus the closure of parts it
   * reaches (see {@link PreservedTheme}), so a workbook's colour and font schemes survive a re-write
   * instead of being replaced by the library's default Office theme. Passing `undefined` drops back
   * to that default.
   */
  restoreThemePart(theme: PreservedTheme | undefined, declared: DeclaredThemeSchemes): void;

  /**
   * Reinstate the named cell styles (`cellStyleXfs`/`cellStyles`) read from a file, index for index,
   * so a cell's link to a named style (its `xfId`) stays valid on re-write. Index 0 is the Normal
   * default.
   */
  restoreNamedStyles(styles: readonly NamedCellStyle[]): void;

  /**
   * Reinstate font id 0 as a file declared it: the face its unstyled cells render in, and the metric
   * its column widths are expressed in character units of. Restored rather than assumed because the
   * library must not inject its own default ahead of one a file already states: doing so replaces the
   * declared face on every empty cell and silently changes what every `<col width>` means.
   *
   * {@link Workbook.setDefaultFont} is the authoring verb; this is only the file's own word.
   */
  restoreDefaultFont(font: Font | undefined): void;

  /**
   * Reinstate the threaded-comment identity registry (`xl/persons/person.xml`) read from a file: the
   * authors and mentioned people a comment thread's messages point at.
   *
   * Entries are keyed by {@link Person.id} and by nothing else. A single human legitimately owns
   * several entries: Excel interns a *mentioned* identity as its own `providerId="PeoplePicker"` entry
   * beside that person's `providerId="AD"` authoring entry (same `displayName`, same `userId`, a
   * different id) and points the mention at the new one. Collapsing entries by name or `userId` would
   * merge those two and silently re-point every mention at the wrong identity.
   *
   * {@link Workbook.addPerson} is the authoring verb.
   */
  restorePersons(persons: readonly Person[]): void;

  /**
   * Every row and column splice made to the workbook's sheets, in the order they were made, for the
   * writer to replay over the references a preserved part spells.
   */
  splices(): readonly SheetSplice[];
}
