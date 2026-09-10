// Turning an OOXML colour *reference* into a concrete colour.
//
// A `<color>` in a spreadsheet states its value in one of three ways, and two of them are references
// that mean nothing on their own: `indexed="n"` points into a 64-entry legacy palette carried over
// from the BIFF era, and `theme="n"` points into the workbook theme's colour scheme. Either may
// additionally carry a `tint`, which lightens or darkens whatever it resolved to.
//
// Resolution is deliberately a *derived* view. Nothing here writes back into the model: a `Color`
// keeps the encoding its file used, so a round-trip re-emits `theme="4" tint="0.4"` rather than
// rewriting every cell as a literal ARGB, which would bloat the styles table, break the link to the
// theme (recolouring the workbook would stop working), and change what the file means.

import {hex} from '../hex.ts';
import {type Color, parseArgb} from './style.ts';
import {
  DEFAULT_THEME_COLOR_SCHEME,
  isThemeRgb,
  THEME_COLOR_SLOTS,
  type ThemeColorScheme,
} from './theme.ts';

/**
 * The built-in indexed colour palette (ECMA-376 §18.8.27), by index. Entries 0–7 duplicate 8–15, a
 * redundancy the spec preserves for backwards compatibility with the legacy formats this palette came
 * from, and the table is only 64 long: indices 64 and 65 are the *system* foreground and background,
 * which name whatever the operating system's window colours are and therefore have no fixed value at
 * all (see {@link SYSTEM_INDEXED_COLORS}).
 *
 * The spec writes each entry with a leading `00`. That byte is not an alpha channel, since a palette
 * of fully transparent colours would be absurd. It is an artefact of the 32-bit colour records these
 * values were lifted from, which is why {@link resolveColor} returns them fully opaque.
 */
// Laid out eight per row, matching how the spec tabulates the palette: the 0-7 / 8-15
// duplication and the 16-entry banding are legible here and invisible one-per-line.
// oxfmt-ignore
export const DEFAULT_INDEXED_COLORS: readonly string[] = [
  '00000000', '00FFFFFF', '00FF0000', '0000FF00', '000000FF', '00FFFF00', '00FF00FF', '0000FFFF',
  '00000000', '00FFFFFF', '00FF0000', '0000FF00', '000000FF', '00FFFF00', '00FF00FF', '0000FFFF',
  '00800000', '00008000', '00000080', '00808000', '00800080', '00008080', '00C0C0C0', '00808080',
  '009999FF', '00993366', '00FFFFCC', '00CCFFFF', '00660066', '00FF8080', '000066CC', '00CCCCFF',
  '00000080', '00FF00FF', '00FFFF00', '0000FFFF', '00800080', '00800000', '00008080', '000000FF',
  '0000CCFF', '00CCFFFF', '00CCFFCC', '00FFFF99', '0099CCFF', '00FF99CC', '00CC99FF', '00FFCC99',
  '003366FF', '0033CCCC', '0099CC00', '00FFCC00', '00FF9900', '00FF6600', '00666699', '00969696',
  '00003366', '00339966', '00003300', '00333300', '00993300', '00993366', '00333399', '00333333',
];

/**
 * The two indices that are not colours: 64 is the system foreground and 65 the system background.
 * They resolve to whatever the viewing system's window colours are, so this library reports them
 * unresolved rather than inventing black and white: a caller that wants to paint them must decide
 * for itself what "automatic" means in its context. `indexed="64"` in particular is extremely common:
 * it is the placeholder every solid fill Excel writes carries as its background colour.
 */
export const SYSTEM_INDEXED_COLORS: ReadonlySet<number> = new Set([64, 65]);

/** What a colour reference is resolved against: the workbook's theme scheme and indexed palette. */
export interface ColorResolutionContext {
  /** The theme's colour scheme; the Office default when the workbook carries no theme of its own. */
  readonly theme?: ThemeColorScheme | undefined;
  /**
   * The workbook's custom indexed palette, by index, each entry an ARGB string. Empty or absent means
   * the workbook rides {@link DEFAULT_INDEXED_COLORS}. A custom palette replaces the built-in one
   * wholesale, which is what `<indexedColors>` means, so a short custom palette leaves the indices
   * past its end unresolved rather than falling through to the built-in entry.
   */
  readonly indexed?: readonly string[] | undefined;
}

/**
 * Resolve a colour reference to a concrete 8-hex ARGB string, or `undefined` when it cannot be
 * resolved: an `auto` colour, a system indexed colour, a theme slot the workbook's scheme does not
 * declare, or an index past the end of a custom palette.
 *
 * Precedence follows what the encodings mean: an explicit `argb` is already concrete and wins; then
 * `theme`, then `indexed`. A `tint` applies to whatever the base resolved to.
 *
 * Alpha: a resolved `theme`/`indexed` colour comes back fully opaque, because neither the theme scheme
 * nor the palette carries a meaningful alpha (see {@link DEFAULT_INDEXED_COLORS}). An explicit `argb`
 * keeps the alpha the file stated.
 */
export function resolveColor(
  color: Color,
  context: ColorResolutionContext = {},
): string | undefined {
  const base = resolveBase(color, context);
  if (base === undefined) return undefined;
  const {tint} = color;
  if (tint === undefined || !Number.isFinite(tint) || tint === 0) return base;
  return applyTint(base, tint);
}

// Resolution is a read path over foreign data, so a value that does not parse resolves to nothing
// rather than to a half-parsed colour; the writer's counterpart throws, because there a malformed
// value is a caller's bug. Uppercasing here is what settles the casing the three encodings disagree
// on: a resolved colour is one concrete ARGB however it was spelled.
function resolveArgb(value: string): string | undefined {
  return parseArgb(value)?.toUpperCase();
}

function resolveBase(color: Color, context: ColorResolutionContext): string | undefined {
  if (color.argb !== undefined) return resolveArgb(color.argb);
  if (color.theme !== undefined) {
    const slot = THEME_COLOR_SLOTS[color.theme];
    if (slot === undefined) return undefined;
    const value = (context.theme ?? DEFAULT_THEME_COLOR_SCHEME)[slot];
    // A caller-built scheme reaches here unchecked, so the slot is read by the theme reader's grammar
    // rather than trusted: prefixing the opaque alpha onto `#FF0000` or an 8-hex value made `FF#FF0000`
    // and a ten-digit colour.
    return value !== undefined && isThemeRgb(value) ? `FF${value.toUpperCase()}` : undefined;
  }
  if (color.indexed !== undefined) {
    if (SYSTEM_INDEXED_COLORS.has(color.indexed)) return undefined;
    const custom = context.indexed;
    const entry =
      custom !== undefined && custom.length > 0
        ? custom[color.indexed]
        : DEFAULT_INDEXED_COLORS[color.indexed];
    if (entry === undefined) return undefined;
    // The palette's leading `00` is not an alpha channel, so the entry is re-opacified rather than
    // carried through (see {@link DEFAULT_INDEXED_COLORS}).
    const parsed = resolveArgb(entry);
    return parsed === undefined ? undefined : `FF${parsed.slice(2)}`;
  }
  return undefined;
}

/**
 * Apply a `tint` to a concrete ARGB colour: convert to HLS, shift the luminance, convert back.
 *
 * The shift is ECMA-376's (§18.3.1.15): `-1` darkens to black, `+1` lightens to white, `0` is a
 * no-op. The spec states the formula but not the rounding either side of the RGB↔HLS conversion, so
 * an implementation cannot be checked against the prose alone. This one lands within 2/255 per
 * channel of what Excel Desktop itself renders, measured over three hues × twelve tints and recorded
 * in `test/corpus/fixtures/excel-oracle/theme-color-tint-luminance.json`. Closing that last gap would
 * mean reproducing one Excel build's fixed-point HLS rounding, and the difference is not visible.
 */
export function applyTint(argb: string, tint: number): string {
  const alpha = argb.slice(0, 2);
  const r = Number.parseInt(argb.slice(2, 4), 16) / 255;
  const g = Number.parseInt(argb.slice(4, 6), 16) / 255;
  const b = Number.parseInt(argb.slice(6, 8), 16) / 255;

  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const lightness = (max + min) / 2;
  let saturation = 0;
  let hue = 0;
  if (max !== min) {
    const delta = max - min;
    saturation = lightness > 0.5 ? delta / (2 - max - min) : delta / (max + min);
    if (max === r) hue = (g - b) / delta + (g < b ? 6 : 0);
    else if (max === g) hue = (b - r) / delta + 2;
    else hue = (r - g) / delta + 4;
    hue /= 6;
  }

  // Lum' = Lum * (1 + tint) when darkening; Lum' = Lum * (1 - tint) + tint when lightening. The
  // spec writes the second as `Lum * (1 - tint) + (HLSMAX - HLSMAX * (1 - tint))`, which is the same
  // thing once the HLSMAX scale is divided out.
  const shifted = tint < 0 ? lightness * (1 + tint) : lightness * (1 - tint) + tint;

  return alpha + hlsToRgbHex(hue, saturation, clamp01(shifted));
}

function clamp01(value: number): number {
  return value < 0 ? 0 : value > 1 ? 1 : value;
}

function hlsToRgbHex(hue: number, saturation: number, lightness: number): string {
  if (saturation === 0) {
    const grey = channelHex(lightness);
    return grey + grey + grey;
  }
  const q =
    lightness < 0.5
      ? lightness * (1 + saturation)
      : lightness + saturation - lightness * saturation;
  const p = 2 * lightness - q;
  return (
    channelHex(hueToChannel(p, q, hue + 1 / 3)) +
    channelHex(hueToChannel(p, q, hue)) +
    channelHex(hueToChannel(p, q, hue - 1 / 3))
  );
}

function hueToChannel(p: number, q: number, offset: number): number {
  let t = offset;
  if (t < 0) t += 1;
  if (t > 1) t -= 1;
  if (t < 1 / 6) return p + (q - p) * 6 * t;
  if (t < 1 / 2) return q;
  if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
  return p;
}

function channelHex(value: number): string {
  return hex(Math.round(clamp01(value) * 255), 2);
}
