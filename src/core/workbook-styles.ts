// The workbook's style tables, lifted off `Workbook` into the slice they actually are.
//
// Six fields that are all one thing: a table read out of `styles.xml` and handed back to the writer
// mostly untouched. `<dxfs>`, the named cell styles, the custom `<indexedColors>` palette, the
// `<mruColors>` swatches and the `<tableStyles>` block are all preserved rather than interpreted,
// because each is the target of an index held somewhere else in the file. Dropping one does not lose
// a look, it leaves a `dxfId`, an `xfId`, an `indexed="…"` or a `tableStyleInfo/@name` pointing at
// nothing. Beside them sits the one authoring verb that writes into the same territory,
// `addCustomTableStyle`, so a caller's own definitions and a file's are read back through the same
// object rather than from two halves of `Workbook`.
//
// `style.ts` holds the style *model* and `io/xlsx/styles.ts` the writer's interning tables; what
// lives here is the per-workbook state between them, and the four shapes that state is made of.
//
// Those four sat in `style.ts` until they were noticed for what they are: a named cell style, a
// differential style and the `<tableStyles>` block are not cell style at all, they are entries in
// workbook style *tables*, and keeping them beside `Fill` meant every module wanting a colour also
// pulled in the table-styles preservation vocabulary. They compose `CellStyle`, which is why they
// looked at home there; composing a type is not being one.
//
// The edge to the rest of the model is one, and it is `indexedPalette`. Resolving an `indexed="…"`
// colour needs this palette, but that is a colour-resolution concern rather than a styles one, so
// `WorkbookTheme` reaches it as a narrow accessor and the state stays here with the `<indexedColors>`
// element the writer emits from it. That is the same boundary `workbook-theme.ts` describes from its
// side; the accessor now closes over this slice instead of over `Workbook` itself.
//
// The doc comments for the public surface stay on `Workbook`'s accessors, which is what the API
// reference is generated from and what a consumer reads. See `docs/architecture.md`.

import {replaceContents} from './containers.ts';
import type {CellStyle} from './style.ts';
import type {TableStyle} from './table-style.ts';

/**
 * A named cell style: the OOXML `cellStyleXfs`/`cellStyles` layer. A spreadsheet applies a built-in
 * or custom style (e.g. "Normal", "Accent1") whose visual facets live in this shared, named layer
 * rather than on each cell's direct format; a cell links to it and inherits any facet the direct
 * format leaves unset. The facets are a cell's own (see {@link CellStyle}); `name` is the style's
 * display name and `builtinId` its Excel gallery index when it is a built-in style.
 */
export type NamedCellStyle = Readonly<CellStyle> & {
  readonly name?: string;
  readonly builtinId?: number;
};

/**
 * A differential style (OOXML CT_Dxf): formatting laid *over* whatever a cell already carries. Only
 * the facets present override; the rest of the cell's own style shows through. It carries the subset
 * of the cell-style facets (see {@link CellStyle}) a `<dxf>` can express: font, number format, fill,
 * and border.
 *
 * Differential styles live in one workbook-level table (`<dxfs>`) that several features index into:
 * a conditional-formatting rule's highlight format, and a table style's per-element formatting
 * (`<tableStyleElement dxfId="…">`). They are interned and shared, so two features asking for the
 * same formatting land on one entry.
 *
 * Not every facet reaches every consumer. As a **table style element**, Excel applies only the font,
 * fill, and border: its own object model exposes `Font`, `Interior`, and `Borders` on a table style
 * element and nothing for a number format, so a `numFmt` set here is carried faithfully through a
 * round-trip but has no visible effect. The type is left whole rather than split, because the same
 * value is legitimately reused across both consumers and narrowing it would only move the surprise.
 */
export type DifferentialStyle = Pick<CellStyle, 'font' | 'numFmt' | 'fill' | 'border'>;

/**
 * The `<tableStyles>` block of a styles part: the custom table/pivot style definitions a file
 * declares, and the two gallery names it nominates as the default for a new table and a new pivot.
 *
 * Each entry of {@link styles} is one `<tableStyle>…</tableStyle>` fragment kept verbatim, for the
 * same reason a `<dxf>` is: a `tableStyleElement`'s `dxfId` indexes the differential-style table,
 * which the writer re-emits **at its original indices**, so the references stay valid without
 * reparsing anything. That index-stability is load-bearing: renumbering the dxf table would
 * silently re-point every preserved table style at a different format.
 *
 * The two default names are ordinary strings, not fragments: they are re-escaped on write, so they
 * are held decoded.
 */
export interface TableStyleTable {
  readonly styles: readonly TableStyleDefinition[];
  readonly defaultTableStyle?: string | undefined;
  readonly defaultPivotStyle?: string | undefined;
  /**
   * The namespace prefixes the verbatim {@link styles} fragments use, mapped to their URI and to
   * whether the source marked the prefix ignorable (`mc:Ignorable`).
   *
   * Carrying a fragment verbatim carries its *prefixes* too. Excel stamps a revision id
   * (`xr9:uid="{…}"`) on every `<tableStyle>` it writes, so a fragment re-emitted under a
   * `<styleSheet>` that declares only the default namespace is not namespace-well-formed, and no
   * consumer can parse the part at all, which is a far louder failure than the dropped table style
   * this preservation exists to prevent. The writer re-declares each prefix on `<styleSheet>` and
   * re-states the ignorable ones, exactly as the source did.
   */
  readonly namespaces?: readonly TableStyleNamespace[];
}

/**
 * One preserved `<tableStyle>` definition: its bytes, and the `name` a table's style info matches it by.
 *
 * The name travels beside the fragment rather than being read back out of it. An authored style
 * replaces the preserved definition it shares a name with, and a pattern pulling `name="…"` out of
 * the bytes did not see `name='…'`, which is the same attribute in single quotes: the writer emitted
 * both definitions under one name, the duplicate {@link Workbook.addTableStyle} promises cannot happen.
 * The reader takes the name from the scanner that read the element, decoded, whatever its quoting.
 */
export interface TableStyleDefinition {
  /** The style's `name`, decoded; empty for a fragment that declared none. */
  readonly name: string;
  /** The `<tableStyle>…</tableStyle>` element exactly as the file spelled it. */
  readonly xml: string;
}

/** One namespace declaration a preserved `<tableStyle>` fragment depends on. */
export interface TableStyleNamespace {
  readonly prefix: string;
  readonly uri: string;
  /** Whether the source listed this prefix in the stylesheet's `mc:Ignorable`. */
  readonly ignorable: boolean;
}

/**
 * The preserved style tables of a workbook: the `<dxfs>` fragments, the named cell styles, the two
 * colour lists, and the table-style definitions a file declared plus the ones a caller authored.
 *
 * Every restore replaces what it restores rather than merging, because a reader states a table
 * whole: a half-restored `<dxfs>` would leave existing `dxfId` references pointing into a mix of two
 * files.
 */
export class WorkbookStyleTables {
  // Differential styles (`<dxfs>`) are a workbook-level table in styles.xml that conditional
  // formatting references by index. The library models the classic scale rules directly but preserves
  // the dxf table as opaque XML fragments, so a rule that references a dxfId (a highlight fill, a
  // custom number format) keeps a valid target across a read/write cycle instead of dangling.
  readonly #differentialStyles: string[] = [];

  // Named cell styles (`cellStyleXfs`/`cellStyles` in styles.xml): the shared, named formatting layer
  // a cell links to by index. Preserved so a cell whose fill/font/… lives only in a named style keeps
  // that style, and the link, across a round-trip. Empty when a file declares nothing beyond the
  // default Normal style, in which case the writer emits just that default.
  readonly #namedStyles: NamedCellStyle[] = [];

  // A custom indexed-color palette (`<colors><indexedColors>` in styles.xml) read from a file, each
  // entry a verbatim `<rgbColor rgb="…"/>` fragment. Preserved so an `indexed="…"` colour reference
  // keeps its intended RGB across a round-trip instead of resolving to a different default-palette
  // entry. Empty for a workbook that never overrode the palette.
  readonly #indexedColors: string[] = [];

  // The most-recently-used colour swatches (`<colors><mruColors>` in styles.xml), each a verbatim
  // `<color rgb="…"/>` fragment. The author's own working set of colours; dropping it on a re-write
  // quietly resets a habit. Empty for a workbook that never picked a custom colour.
  readonly #mruColors: string[] = [];

  // The custom table-style definitions (`<tableStyles>` in styles.xml), each `<tableStyle>` kept
  // verbatim, plus the gallery names the file nominates as the default for a new table and pivot. A
  // table's `tableStyleInfo/@name` can name one of these definitions, so dropping the block leaves
  // that reference dangling and the table renders unstyled.
  #tableStyles: TableStyleTable = {styles: []};

  // Table styles authored on this workbook, keyed by name so registering the same name twice replaces
  // rather than duplicates: two definitions sharing a name leave a table's reference ambiguous.
  readonly #customTableStyles = new Map<string, TableStyle>();

  get differentialStyles(): readonly string[] {
    return this.#differentialStyles;
  }

  get namedStyles(): readonly NamedCellStyle[] {
    return this.#namedStyles;
  }

  get indexedColors(): readonly string[] {
    return this.#indexedColors;
  }

  get mruColors(): readonly string[] {
    return this.#mruColors;
  }

  get tableStyles(): TableStyleTable {
    return this.#tableStyles;
  }

  get customTableStyles(): readonly TableStyle[] {
    return [...this.#customTableStyles.values()];
  }

  // Validation lives on `Workbook.addTableStyle`, the authoring surface a caller's mistake has to
  // surface at; by here the style is already known good.
  addCustomTableStyle(style: TableStyle): void {
    this.#customTableStyles.set(style.name, style);
  }

  // The custom palette as plain ARGB strings. `#indexedColors` holds verbatim `<rgbColor rgb="…"/>`
  // fragments, the form the writer re-emits, so the value is read out here rather than stored twice
  // in two shapes that could drift.
  indexedPalette(): readonly string[] {
    return this.#indexedColors.map((fragment) => /\brgb="([^"]*)"/.exec(fragment)?.[1] ?? '');
  }

  // The reader's channel; each is reached through one `Workbook[INTERNAL]` restore.
  restoreDifferentialStyles(fragments: readonly string[]): void {
    replaceContents(this.#differentialStyles, fragments);
  }

  // A differential style a sheet carries inline rather than by index takes the index of an identical
  // entry, or a new one at the end, which leaves every existing index where it was. Excel writes an
  // extension-form rule's style into the table as well, so the lookup usually finds it there.
  adoptDifferentialStyle(fragment: string): number {
    const existing = this.#differentialStyles.indexOf(fragment);
    return existing === -1 ? this.#differentialStyles.push(fragment) - 1 : existing;
  }

  restoreNamedStyles(styles: readonly NamedCellStyle[]): void {
    replaceContents(this.#namedStyles, styles);
  }

  restoreIndexedColors(fragments: readonly string[]): void {
    replaceContents(this.#indexedColors, fragments);
  }

  restoreMruColors(fragments: readonly string[]): void {
    replaceContents(this.#mruColors, fragments);
  }

  restoreTableStyles(table: TableStyleTable): void {
    this.#tableStyles = table;
  }
}
