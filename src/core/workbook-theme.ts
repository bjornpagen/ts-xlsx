// The workbook's theme, lifted off `Workbook` into the slice it actually is.
//
// `theme.ts` holds the theme *model*: the slot order, the defaults, the value vocabulary. What lives
// here is the per-workbook state around it: the preserved part, the schemes decoded out of it, and
// the overrides a caller authored. Folding this into `theme.ts` would mix a pure value module with
// mutable per-instance state, which is how `theme.ts` grows the same legibility problem `Workbook`
// had.
//
// Nothing here touches the part's *text*. The codec decodes the schemes as it reads the part and
// hands both over together, and composes the authored overrides back onto the part as it writes; the
// model holds values on either side of that. See `io/xlsx/theme-xml.ts`.
//
// The boundary that matters, and the reason this slice is not as clean as the VBA one: colour
// resolution needs the workbook's custom `<indexedColors>` palette as well as the theme scheme, and
// that palette is *also* the writer's source for the `<indexedColors>` element and is populated by
// the reader. So it lives with the other style tables (`workbook-styles.ts`) and reaches here as a
// narrow accessor over them. Moving it in would drag that whole table along and turn a theme overlay
// into a colour-and-styles overlay, which is not a slice.
//
// The doc comments for the public surface stay on `Workbook`'s accessors, which is what the API
// reference is generated from and what a consumer reads.

import {resolveColor} from './color-resolution.ts';
import type {Color} from './style.ts';
import {
  DEFAULT_THEME_COLOR_SCHEME,
  DEFAULT_THEME_FONTS,
  normalizeThemeColor,
  type ThemeColorScheme,
  type ThemeColorSlot,
  type ThemeFontScheme,
  type ThemeOverrides,
} from './theme.ts';
import type {PreservedTheme} from './workbook.ts';

/** The schemes a theme part declares, as the codec that read it decoded them. */
export interface DeclaredThemeSchemes {
  readonly colors: ThemeColorScheme;
  readonly fonts: ThemeFontScheme;
}

/**
 * The theme slice of a workbook: the preserved part, the colour scheme and typefaces every
 * `theme="n"` reference resolves against, and the overrides {@link Workbook.setTheme} authors over
 * them.
 */
export class WorkbookTheme {
  // The workbook's custom indexed palette, read on demand rather than held: it belongs to the styles
  // slice, not to the theme, and only colour resolution needs it.
  readonly #indexedPalette: () => readonly string[];

  // The theme part read from a file, kept verbatim with the closure of parts it reaches. The writer
  // emits its own default theme for a workbook that has none, so without this a branded theme would be
  // overwritten by that default and every `theme="n"` colour in the file would silently re-render.
  // Undefined for a workbook authored from scratch, or read from a package declaring no theme.
  #part: PreservedTheme | undefined;

  // The schemes the preserved part declares, decoded by the reader that restored it. Empty for a
  // workbook authored from scratch, and empty for a part declaring nothing this reader understands;
  // either way the getters below fall back to the Office defaults.
  #declared: DeclaredThemeSchemes = {colors: {}, fonts: {}};

  // Colour slots and typefaces the caller authored, merged over whatever the workbook already had.
  #authored: {colors: {-readonly [K in ThemeColorSlot]?: string}; fonts: ThemeFontScheme} = {
    colors: {},
    fonts: {},
  };

  constructor(indexedPalette: () => readonly string[]) {
    this.#indexedPalette = indexedPalette;
  }

  get part(): PreservedTheme | undefined {
    return this.#part;
  }

  // The reader's channel: the part, and the schemes the reader decoded out of it.
  restorePart(theme: PreservedTheme | undefined, declared: DeclaredThemeSchemes): void {
    this.#part = theme;
    this.#declared = declared;
  }

  author(overrides: ThemeOverrides): void {
    // Normalised eagerly rather than at write time: a colour the writer would reject surfaces far
    // from the call that supplied it, and a slot stored as spelled reached `themeColors` and colour
    // resolution raw, where `#FF0000` resolved to `FF#FF0000`. Every colour is read before any is
    // stored, so a refused call leaves the scheme as it was.
    const colors = Object.fromEntries(
      Object.entries(overrides.colors ?? {}).map(([slot, value]) => [
        slot,
        normalizeThemeColor(value),
      ]),
    );
    Object.assign(this.#authored.colors, colors);
    this.#authored.fonts = {...this.#authored.fonts, ...overrides.fonts};
  }

  // A theme that declares no scheme (or none the reader decodes) falls back to the Office default
  // rather than resolving nothing: the file still renders against *some* scheme, and the default is
  // the one the writer would have shipped.
  get colors(): ThemeColorScheme {
    const {colors} = this.#declared;
    const base = Object.keys(colors).length === 0 ? DEFAULT_THEME_COLOR_SCHEME : colors;
    return {...base, ...this.#authored.colors};
  }

  get fonts(): ThemeFontScheme {
    const {fonts} = this.#declared;
    const base = Object.keys(fonts).length === 0 ? DEFAULT_THEME_FONTS : fonts;
    return {...base, ...this.#authored.fonts};
  }

  // The typefaces a caller authored, unmerged. `Workbook.defaultFont` needs to know whether the body
  // face was named outright or merely inherited, which the resolved {@link fonts} cannot say.
  get authoredFonts(): ThemeFontScheme {
    return this.#authored.fonts;
  }

  // What the writer composes onto the part it is about to emit, or `undefined` when nothing was
  // authored and the part rides through verbatim.
  get overrides(): ThemeOverrides | undefined {
    const {colors, fonts} = this.#authored;
    if (Object.keys(colors).length === 0 && Object.keys(fonts).length === 0) return undefined;
    return {colors, fonts};
  }

  resolveColor(color: Color): string | undefined {
    return resolveColor(color, {theme: this.colors, indexed: this.#indexedPalette()});
  }
}
