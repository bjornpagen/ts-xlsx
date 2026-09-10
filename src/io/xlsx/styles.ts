// The write-side style table: the interned `<numFmts>`/`<fills>`/`<cellXfs>` backing styles.xml.
//
// The elements themselves are next door in `style-elements.ts`. The line is exactly the one this
// module is named for: nothing there touches a `StyleRegistry`, and this calls in there for every
// fragment it interns.
//
// OOXML styles are a *shared* table referenced by index: a cell (or a formatted row/column)
// names a `<cellXfs>` entry via its `s` attribute, and that entry names a fill by id and a
// number format by id. Identical styles must collapse to one entry, both to produce
// well-formed OOXML and to keep write cost bounded on large, lightly-formatted sheets (the
// historical performance cliff came from re-serialising a distinct style per cell). The
// registry interns each distinct fill, number format, and xf, handing back a stable index.
//
// Fills, number formats, fonts, borders, alignment, and protection are modelled today.
// Fills/fonts/borders are shared sub-tables the xf names by id, whereas alignment and protection
// are child elements *of* the xf, so each is interned into the xf signature directly rather than
// into its own id table, and an aligned/protected xf carries them as body children in that order.
// An unstyled cell/row/column resolves to xf 0.

import {type Border, type Fill, type Font} from '../../core/style.ts';
import {TABLE_STYLE_ELEMENT_TYPES, type TableStyle} from '../../core/table-style.ts';
import type {
  DifferentialStyle,
  NamedCellStyle,
  TableStyleTable,
} from '../../core/workbook-styles.ts';
import {escapeAttr, escapeFormatCode, numAttr, XML_DECLARATION} from '../../xml/xml.ts';
import type {XfStyle} from '../style/xf-style.ts';
import {fontXml} from './font-xml.ts';
import {MARKUP_COMPATIBILITY_NS, SPREADSHEETML_NS} from './namespaces.ts';
import {
  alignmentAttrs,
  borderXml,
  cellStyleTag,
  type CellFormat,
  CUSTOM_NUMFMT_BASE,
  DEFAULT_BORDER,
  DEFAULT_FONT_BODY,
  DEFAULT_FORMAT,
  dxfXml,
  fillSignature,
  formatSignature,
  isDefaultFormat,
  patternFillXml,
  protectionAttrs,
  xfXml,
} from './style-elements.ts';

// Excel reserves fill ids 0 and 1 for the "none" and "gray125" patterns it always emits;
// custom fills are numbered from 2 so a foreign reader's built-in assumptions still hold.
const RESERVED_FILL_COUNT = 2;

// Font id 0 is the always-present default font: the workbook's, not an assumed one. Custom fonts
// are numbered from 1.
const RESERVED_FONT_COUNT = 1;

// Border id 0 is the always-present empty border (every edge absent); custom borders from 1.
const RESERVED_BORDER_COUNT = 1;

/** What a {@link StyleRegistry} needs from its workbook before any style is interned.
 */
export interface StyleRegistryOptions {
  /**
   * The font every cell naming none of its own renders in, emitted as id 0. Take it from
   * {@link Workbook.defaultFont}, which resolves the authored/declared/theme chain and guarantees a
   * complete entry. Omitted, the registry falls back to the Office default.
   */
  readonly defaultFont?: Font;

  /**
   * The default font the *source package* declared, when a workbook was read
   * ({@link Workbook.declaredDefaultFont}). It also interns to id 0, which is what lets an authored
   * default font reach cells the reader had already resolved.
   *
   * The reader flattens font 0 onto every xf that names it, so a cell that merely inherited the
   * file's default arrives carrying it as a concrete face. That face is an artefact of reading, not
   * an authored intent (in the source file the cell said nothing about its font) so it must
   * collapse back to id 0 and follow the new default rather than pin itself to the old one through a
   * custom entry.
   */
  readonly declaredDefaultFont?: Font;
}

/**
 * One de-duplicating table of serialised style fragments: the entries in id order, and the key each
 * was first interned under. Fills, fonts, borders, number formats and differential styles are all
 * this same shape, differing only in the id their first entry is numbered from and in what they use
 * as a key, so stating the array/map invariant once leaves one place it can break rather than five.
 *
 * The key is not always the entry. A fill is keyed by a signature over the model and stores the XML
 * that signature serialises to; a font is keyed by its body and stores that body wrapped in
 * `<font>`. Where the two coincide, the caller passes the same string twice.
 */
class InternTable {
  readonly #entries: string[] = [];
  readonly #idByKey = new Map<string, number>();
  readonly #base: number;

  constructor(base: number) {
    this.#base = base;
  }

  get entries(): readonly string[] {
    return this.#entries;
  }

  get size(): number {
    return this.#entries.length;
  }

  /** The id for `key`, appending `entry` at the next id on first sight. */
  intern(key: string, entry: string): number {
    const existing = this.#idByKey.get(key);
    if (existing !== undefined) return existing;
    const id = this.#base + this.#entries.length;
    this.#entries.push(entry);
    this.#idByKey.set(key, id);
    return id;
  }

  /**
   * Append a preserved entry at the next id, keeping it verbatim and *not* deduping it away.
   *
   * A seeded entry still becomes reusable: an authored entry identical to it interns to this id
   * rather than appending a copy. It does not displace an earlier entry already holding the key,
   * which is what keeps the first of two identical seeds the one everything resolves to.
   */
  seed(entry: string): number {
    const id = this.#base + this.#entries.length;
    this.#entries.push(entry);
    if (!this.#idByKey.has(entry)) this.#idByKey.set(entry, id);
    return id;
  }
}

/**
 * The interned style table one written package is built against: every distinct fill, number format,
 * font, border and xf, each handed back as the stable index a cell's `s` attribute names.
 *
 * One workbook, one registry. The streaming writer is why that is worth saying: it interns a row's
 * styles at the moment the row is flushed, long before `xl/styles.xml` is emitted, so the ids already
 * baked into those rows are only correct if the very same registry emits the part.
 */
export class StyleRegistry {
  // The `<font>` body emitted as id 0.
  readonly #defaultFontBody: string;

  // Every serialised font body that means "id 0": the emitted default, plus the one a source file
  // declared. See {@link StyleRegistryOptions.declaredDefaultFont} for why the second belongs here.
  readonly #font0Bodies: ReadonlySet<string>;

  constructor(options: StyleRegistryOptions = {}) {
    this.#defaultFontBody =
      options.defaultFont === undefined ? DEFAULT_FONT_BODY : fontXml(options.defaultFont);
    this.#font0Bodies = new Set(
      options.declaredDefaultFont === undefined
        ? [this.#defaultFontBody]
        : [this.#defaultFontBody, fontXml(options.declaredDefaultFont)],
    );
  }

  // Custom fill xml fragments, keyed by a signature over the fill they serialise.
  readonly #fills = new InternTable(RESERVED_FILL_COUNT);

  // Custom number-format codes, each its own key.
  readonly #numFmts = new InternTable(CUSTOM_NUMFMT_BASE);

  // Custom font elements, keyed by the body they wrap.
  readonly #fonts = new InternTable(RESERVED_FONT_COUNT);

  // Custom border xml fragments, each its own key.
  readonly #borders = new InternTable(RESERVED_BORDER_COUNT);

  // xf 0 is the default (no fill/font/border/alignment/protection, General format); further entries append as styles appear.
  readonly #formats: CellFormat[] = [DEFAULT_FORMAT];
  readonly #xfIndexBySignature = new Map<string, number>();

  // The named-style layer (`<cellStyleXfs>` / `<cellStyles>`): the base formats a cell's `xfId` links
  // into, and the names that label them. Index 0 is always Normal. A file with named styles seeds this
  // in place of the default via {@link seedNamedStyles}; otherwise the default alone is emitted.
  readonly #cellStyleXfs: CellFormat[] = [DEFAULT_FORMAT];
  readonly #cellStyleNames: {name: string; builtinId?: number; xfId: number}[] = [
    {name: 'Normal', builtinId: 0, xfId: 0},
  ];

  // Differential styles (`<dxfs>`) that conditional formatting references by index. Fragments read
  // from a file are seeded first and kept verbatim so a foreign rule's dxfId stays valid; a style
  // authored on a rule is serialised and appended after them, dedup'd by its fragment.
  readonly #dxfs = new InternTable(0);

  // A custom indexed-color palette (`<colors><indexedColors>`) read from a file, each entry a verbatim
  // `<rgbColor rgb="…"/>`. Preserved and re-emitted unchanged so cells/fonts/borders that reference a
  // colour by `indexed="…"` keep their intended RGB; dropping it would silently resolve every indexed
  // colour to a different default-palette entry. Empty for a workbook that never overrode the palette.
  readonly #indexedColors: string[] = [];

  // The most-recently-used colour swatches (`<colors><mruColors>`) read from a file, each entry a
  // verbatim `<color rgb="…"/>` fragment. Purely a UI convenience, the palette Excel offers under
  // "Recent Colors", but it is the author's own working set, so dropping it on a re-write quietly
  // resets a habit. Empty for a workbook that never picked a custom colour.
  readonly #mruColors: string[] = [];

  // The custom table/pivot style definitions (`<tableStyles>`) read from a file, kept verbatim, plus
  // the gallery names it nominates as defaults. A table's `tableStyleInfo/@name` can point at one of
  // these definitions, so dropping the block leaves that reference dangling and the table renders
  // unstyled. Excel writes the container (with both default attributes) into essentially every file
  // even when it declares no custom style at all.
  #tableStyles: TableStyleTable = {styles: []};

  // Table styles authored on the workbook, serialised on registration and keyed by name so a second
  // definition of the same name replaces the first, as does one that overrides a preserved
  // definition, since two `<tableStyle>` elements sharing a name leave a table's reference ambiguous.
  readonly #authoredTableStyles = new Map<string, string>();

  /**
   * The `<cellXfs>` index for a composed cell/row/column style. A style with no facet needs
   * no entry and resolves to the default xf 0, so its owner emits no `s` attribute at all.
   */
  styleId(style: XfStyle): number {
    const format = this.#composeFormat(style, style.xfId ?? 0);
    // An all-default format that links to no named style needs no entry and resolves to xf 0, so its
    // owner emits no `s` attribute. A non-zero xfId is itself information (the cell inherits a named
    // style) so it forces a real entry even when the direct facets are empty.
    if (isDefaultFormat(format)) return 0;

    const signature = formatSignature(format);
    let index = this.#xfIndexBySignature.get(signature);
    if (index === undefined) {
      index = this.#formats.length;
      this.#formats.push(format);
      this.#xfIndexBySignature.set(signature, index);
    }
    return index;
  }

  // Compose a style's facets into an interned {@link CellFormat}, interning each fill/font/border/
  // number-format into its shared sub-table. Shared by the cell-format path ({@link styleId}) and the
  // named-style path ({@link seedNamedStyles}), which differ only in which table the result lands in.
  #composeFormat(style: XfStyle, xfId: number): CellFormat {
    // A `none` pattern is the reserved fill 0; a gradient is always a real, interned fill.
    const paints =
      style.fill !== undefined && (style.fill.type === 'gradient' || style.fill.pattern !== 'none');
    const fillId = paints ? this.#internFill(style.fill) : 0;
    // A number format is a format-code *string*; a caller that assigns a structured object (e.g. a
    // parsed `{id, formatCode}` copied from another cell) must not have it stringified into the styles
    // part as `formatCode="[object Object]"`, which Excel reports as a corrupt package. A non-string
    // format is dropped to the General format rather than corrupting the file.
    const numFmtId =
      typeof style.numFmt === 'string' && style.numFmt !== ''
        ? this.#internNumFmt(style.numFmt)
        : 0;
    const fontId = style.font ? this.#internFont(style.font) : 0;
    const borderId = style.border ? this.#internBorder(style.border) : 0;
    const alignment = style.alignment ? alignmentAttrs(style.alignment) : '';
    const protection = style.protection ? protectionAttrs(style.protection) : '';
    const quotePrefix = style.quotePrefix === true;
    return {fillId, numFmtId, fontId, borderId, alignment, protection, quotePrefix, xfId};
  }

  /**
   * Seed the named cell styles (`<cellStyleXfs>`/`<cellStyles>`) read from a file, in place of the
   * lone default, interning each style's facets into the shared sub-tables so its `fillId`/`fontId`/…
   * references stay valid against the rebuilt tables. Index 0 stays Normal. A cell's `xfId` indexes
   * this table, so it must be seeded before any {@link styleId} that carries an `xfId`.
   */
  seedNamedStyles(styles: readonly NamedCellStyle[]): void {
    if (styles.length === 0) return;
    this.#cellStyleXfs.length = 0;
    this.#cellStyleNames.length = 0;
    styles.forEach((style, index) => {
      this.#cellStyleXfs.push(this.#composeFormat(style, 0));
      const entry: {name: string; builtinId?: number; xfId: number} = {
        name: style.name ?? `Style ${index}`,
        xfId: index,
      };
      if (style.builtinId !== undefined) entry.builtinId = style.builtinId;
      this.#cellStyleNames.push(entry);
    });
  }

  /**
   * Seed the differential-style table with fragments read from a file, keeping each `<dxf>…</dxf>`
   * verbatim and at its original index. Call once before any {@link differentialStyleId}; authored
   * styles append after these.
   *
   * **Index stability is a contract, not an implementation detail.** A `dxfId` is an index into this
   * one table, and more than one preserved construct resolves through it: a conditional-formatting
   * rule's `dxfId`, and every `<tableStyleElement dxfId="…">` inside a preserved `<tableStyle>` (see
   * {@link seedTableStyles}). Those constructs are carried as opaque XML precisely *because* the
   * indices they name do not move. Renumbering, reordering, or de-duplicating the seeded entries
   * would silently re-point every one of them at a different format: a change no schema check and no
   * round-trip of our own can catch, because the file stays perfectly valid and merely renders wrong.
   */
  seedDifferentialStyles(fragments: readonly string[]): void {
    for (const fragment of fragments) this.#dxfs.seed(fragment);
  }

  /**
   * Seed the custom indexed-color palette (`<colors><indexedColors>`) read from a file, each entry a
   * verbatim `<rgbColor rgb="…"/>` fragment. Re-emitting it unchanged is what keeps an `indexed="…"`
   * colour reference resolving to the RGB the source intended. An empty list leaves the workbook on
   * the default palette and emits no `<colors>` element.
   */
  seedIndexedColors(fragments: readonly string[]): void {
    this.#indexedColors.length = 0;
    this.#indexedColors.push(...fragments);
  }

  /**
   * Seed the most-recently-used colour swatches (`<colors><mruColors>`) read from a file, each entry a
   * verbatim `<color rgb="…"/>` fragment. An empty list emits no `<mruColors>` element.
   */
  seedMruColors(fragments: readonly string[]): void {
    this.#mruColors.length = 0;
    this.#mruColors.push(...fragments);
  }

  /**
   * Seed the custom table-style definitions (`<tableStyles>`) read from a file, each `<tableStyle>`
   * kept verbatim so a table's `tableStyleInfo/@name` still resolves and each element's `dxfId` still
   * indexes the differential-style table {@link seedDifferentialStyles} preserves at its original
   * indices. Replaces any block already held.
   */
  seedTableStyles(table: TableStyleTable): void {
    this.#tableStyles = table;
  }

  /**
   * Serialise a table style authored on the workbook, interning each element's formatting into the
   * differential-style table and emitting the `dxfId` that reaches it. Call after
   * {@link seedTableStyles}, whose preserved definitions these append after.
   *
   * A definition here **replaces** a preserved one of the same name. Two `<tableStyle>` elements
   * sharing a name is ambiguous, since a table's `tableStyleInfo/@name` would reach whichever a
   * consumer happened to index first, so authoring a name the source already used is read as overriding it,
   * which is what asking for it means.
   */
  addTableStyle(style: TableStyle): void {
    const elements = TABLE_STYLE_ELEMENT_TYPES.flatMap((type) => {
      const element = style.elements[type];
      if (element === undefined) return [];
      // `size` defaults to 1, so it is written only when a band is genuinely wider than one row.
      // Through `numAttr`, because `TableStyleElement.size` is public and unvalidated and the
      // attribute is an `xsd:unsignedInt`: interpolated directly, a `NaN` reached the part as the
      // four letters, which is a package Excel reports as damaged rather than a value it ignores.
      const size = element.size === 1 ? '' : numAttr('size', element.size);
      return [
        `<tableStyleElement type="${type}"${size} dxfId="${this.differentialStyleId(element)}"/>`,
      ];
    });
    // `pivot`/`table` default to true, so each is written only when the caller opts a style out.
    const flags =
      (style.pivot === false ? ' pivot="0"' : '') + (style.table === false ? ' table="0"' : '');
    this.#authoredTableStyles.set(
      style.name,
      `<tableStyle name="${escapeAttr(style.name)}"${flags} count="${elements.length}">` +
        `${elements.join('')}</tableStyle>`,
    );
  }

  /**
   * Intern an authored differential style, returning the `<dxfs>` index that references it: a
   * conditional-formatting rule's `dxfId`, or a table style element's. Identical styles collapse to
   * one entry, whichever feature asked for them, so a highlight rule and a table style's header row
   * painted the same way share a single `<dxf>`.
   *
   * Authored entries append after the seeded ones ({@link seedDifferentialStyles}), which is what
   * keeps every preserved reference pointing where it did.
   */
  differentialStyleId(style: DifferentialStyle): number {
    const fragment = dxfXml(style);
    return this.#dxfs.intern(fragment, fragment);
  }

  #internFill(fill: Fill): number {
    return this.#fills.intern(fillSignature(fill), patternFillXml(fill, {solidBgFallback: true}));
  }

  #internNumFmt(code: string): number {
    return this.#numFmts.intern(code, code);
  }

  // A font whose partial carries no facet that differs from the default contributes nothing
  // and maps to font id 0; otherwise its serialised form is interned and dedup'd like a fill.
  #internFont(font: Font): number {
    const xml = fontXml(font);
    if (xml === '' || this.#font0Bodies.has(xml)) return 0;
    return this.#fonts.intern(xml, `<font>${xml}</font>`);
  }

  // A border that overrides no edge serialises to the empty default border and maps to id 0;
  // otherwise its serialised form is interned and dedup'd like a fill or font.
  #internBorder(border: Border): number {
    const xml = borderXml(border);
    return xml === DEFAULT_BORDER ? 0 : this.#borders.intern(xml, xml);
  }

  /** Serialise the accumulated table into a complete, valid styles.xml part. */
  toXml(): string {
    const fillCount = RESERVED_FILL_COUNT + this.#fills.size;
    const fills =
      '<fill><patternFill patternType="none"/></fill>' +
      '<fill><patternFill patternType="gray125"/></fill>' +
      this.#fills.entries.join('');
    const cellXfs = this.#formats.map((format) => xfXml(format, format.xfId)).join('');
    const cellStyleXfs = this.#cellStyleXfs.map((format) => xfXml(format, null)).join('');
    const cellStyles = this.#cellStyleNames.map(cellStyleTag).join('');
    const fontCount = RESERVED_FONT_COUNT + this.#fonts.size;
    const fonts = `<font>${this.#defaultFontBody}</font>${this.#fonts.entries.join('')}`;
    const borderCount = RESERVED_BORDER_COUNT + this.#borders.size;
    const borders = DEFAULT_BORDER + this.#borders.entries.join('');
    return (
      XML_DECLARATION +
      `<styleSheet xmlns="${SPREADSHEETML_NS}"${this.#foreignNamespaceAttrs()}>` +
      this.#numFmtsXml() +
      `<fonts count="${fontCount}">${fonts}</fonts>` +
      `<fills count="${fillCount}">${fills}</fills>` +
      `<borders count="${borderCount}">${borders}</borders>` +
      `<cellStyleXfs count="${this.#cellStyleXfs.length}">${cellStyleXfs}</cellStyleXfs>` +
      `<cellXfs count="${this.#formats.length}">${cellXfs}</cellXfs>` +
      `<cellStyles count="${this.#cellStyleNames.length}">${cellStyles}</cellStyles>` +
      this.#dxfsXml() +
      this.#tableStylesXml() +
      this.#colorsXml() +
      '</styleSheet>'
    );
  }

  // The `xmlns:…` declarations a preserved `<tableStyle>` fragment depends on, plus the
  // markup-compatibility attributes that tell a consumer to ignore what it does not understand. A
  // workbook carrying no such fragment emits nothing, so the ordinary stylesheet root is unchanged.
  //
  // This is the cost of verbatim preservation: a fragment carries its prefixes with it, and a prefix
  // no ancestor declares makes the whole part unparseable, a much louder failure than the dropped
  // table style the preservation exists to prevent. See {@link TableStyleTable.namespaces}.
  #foreignNamespaceAttrs(): string {
    const namespaces = this.#tableStyles.namespaces ?? [];
    if (namespaces.length === 0) return '';
    const declarations = namespaces.map((ns) => ` xmlns:${ns.prefix}="${escapeAttr(ns.uri)}"`);
    const ignorable = namespaces.filter((ns) => ns.ignorable).map((ns) => ns.prefix);
    const compatibility =
      ignorable.length === 0
        ? ''
        : ` xmlns:mc="${MARKUP_COMPATIBILITY_NS}" mc:Ignorable="${escapeAttr(ignorable.join(' '))}"`;
    return declarations.join('') + compatibility;
  }

  // <tableStyles> sits between <dxfs> and <colors> in CT_Stylesheet's child sequence. It is emitted
  // only when the workbook has something to say there (a preserved or authored style definition, or
  // a nominated default) so a workbook that authors none leaves every table on the built-in gallery
  // and writes nothing. `count` counts the definitions, not the attributes, so a container that only
  // nominates defaults (the shape Excel writes into nearly every file) is self-closing with count="0".
  //
  // Preserved definitions come first and authored ones after, except that an authored style replaces
  // the preserved definition it shares a name with. See {@link addTableStyle}.
  #tableStylesXml(): string {
    const {defaultTableStyle, defaultPivotStyle} = this.#tableStyles;
    const authored = this.#authoredTableStyles;
    const preserved = this.#tableStyles.styles
      .filter((style) => !authored.has(style.name))
      .map((style) => style.xml);
    const styles = [...preserved, ...authored.values()];
    if (styles.length === 0 && defaultTableStyle === undefined && defaultPivotStyle === undefined) {
      return '';
    }
    const attrs =
      ` count="${styles.length}"` +
      (defaultTableStyle === undefined
        ? ''
        : ` defaultTableStyle="${escapeAttr(defaultTableStyle)}"`) +
      (defaultPivotStyle === undefined
        ? ''
        : ` defaultPivotStyle="${escapeAttr(defaultPivotStyle)}"`);
    if (styles.length === 0) return `<tableStyles${attrs}/>`;
    return `<tableStyles${attrs}>${styles.join('')}</tableStyles>`;
  }

  // <colors> is the last modelled child of <styleSheet>, after <dxfs> and <tableStyles>. It holds the
  // custom <indexedColors> palette and then the <mruColors> swatch list, in that CT_Colors order. It
  // is emitted only when a file carried one of them, so an ordinary workbook stays on the built-in
  // indexed colours and writes no <colors> element.
  #colorsXml(): string {
    const indexed =
      this.#indexedColors.length === 0
        ? ''
        : `<indexedColors>${this.#indexedColors.join('')}</indexedColors>`;
    const mru =
      this.#mruColors.length === 0 ? '' : `<mruColors>${this.#mruColors.join('')}</mruColors>`;
    if (indexed === '' && mru === '') return '';
    return `<colors>${indexed}${mru}</colors>`;
  }

  // <dxfs> holds the differential styles conditional formatting references by index. An empty table
  // is still emitted as a self-closing count="0" element, the shape Excel writes; a populated one
  // lists the seeded (foreign) fragments first, then any authored styles, preserving every index.
  #dxfsXml(): string {
    if (this.#dxfs.size === 0) return '<dxfs count="0"/>';
    return `<dxfs count="${this.#dxfs.size}">${this.#dxfs.entries.join('')}</dxfs>`;
  }

  // <numFmts> is the first child of <styleSheet> and is omitted entirely when no custom
  // format was used, matching how Excel writes an all-built-in workbook.
  #numFmtsXml(): string {
    if (this.#numFmts.size === 0) return '';
    const entries = this.#numFmts.entries
      .map(
        (code, i) =>
          `<numFmt numFmtId="${CUSTOM_NUMFMT_BASE + i}" formatCode="${escapeFormatCode(code)}"/>`,
      )
      .join('');
    return `<numFmts count="${this.#numFmts.size}">${entries}</numFmts>`;
  }
}
