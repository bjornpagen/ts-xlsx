// The `<color>` element, both directions.
//
// OOXML spells a colour the same way wherever one appears (`<color>`, `<fgColor>`, `<bgColor>`,
// `<tabColor>`, a differential format's colour) as some combination of `rgb`, `theme`, `tint` and
// `indexed` attributes. Writing that tuple and reading it back are one concern with two directions,
// and they have to agree: `parseColor` must accept exactly what `colorAttrs` emits, or a round-trip
// loses a colour.
//
// They lived in `styles.ts`, the write-side style *table*, which made the style-table reader and the
// worksheet reader import from the writer to decode a colour: the only place in this codec where the
// read pipeline reached into the write pipeline. Nothing about decoding `<color>` belongs to the
// interning table; it just happened to be where the first caller was.

import {type Color, parseArgb} from '../../core/style.ts';
import {quoted} from '../../errors.ts';
import {numFinite, numInteger} from '../../xml/xml-attrs.ts';
import {numberText} from '../../xml/xml.ts';

// The write side of the ARGB grammar `parseArgb` states: this is the single choke point through
// which every fill/font/border/tab colour flows on its way into the file, so a value that does not
// parse is a programming error at the API surface and throws with the offending value rather than
// writing corrupt XML. Excel does not report a malformed `rgb`; it silently renders flat black,
// which is why this is loud and the reader's counterpart is silent.
function normalizeArgb(argb: string): string {
  const rgb = parseArgb(argb);
  if (rgb === undefined) {
    throw new SyntaxError(
      `invalid ARGB colour ${quoted(argb)}: expected 6 or 8 hexadecimal digits`,
    );
  }
  return rgb;
}

/** Serialise a {@link Color} as the attribute list a `<color>`-shaped element carries. */
export function colorAttrs(color: Color): string {
  const parts: string[] = [];
  if (color.argb !== undefined) parts.push(`rgb="${normalizeArgb(color.argb)}"`);
  if (color.theme !== undefined) parts.push(`theme="${numberText(color.theme)}"`);
  if (color.tint !== undefined) parts.push(`tint="${numberText(color.tint)}"`);
  if (color.indexed !== undefined) parts.push(`indexed="${numberText(color.indexed)}"`);
  return parts.join(' ');
}

// The read counterpart of colorAttrs: decode a `<color>`/`<fgColor>`/… element's attributes.
// theme/indexed must be integers and tint a finite number; a malformed foreign attribute is dropped
// rather than propagated as NaN, so a downstream colorAttrs never emits `theme="NaN"`.
export function parseColor(attrs: {readonly [k: string]: string}): Color {
  const color: {argb?: string; theme?: number; tint?: number; indexed?: number} = {};
  if (attrs.rgb !== undefined) color.argb = attrs.rgb;
  const theme = numInteger(attrs.theme, 0);
  if (theme !== undefined) color.theme = theme;
  const tint = numFinite(attrs.tint);
  if (tint !== undefined) color.tint = tint;
  const indexed = numInteger(attrs.indexed, 0);
  if (indexed !== undefined) color.indexed = indexed;
  return color;
}
