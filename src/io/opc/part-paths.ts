// The path algebra of an OPC package: a part's extension, where its relationships live, and how one
// part names another. Pure string work over package-absolute paths, with no notion of what any part
// contains, so the XML and BIFF12 codecs, and the readers and writers within each, share it.

// Where the writer always puts the theme part. The workbook's theme relationship and the package's
// content-type override both name this path unconditionally, so a theme preserved from a source
// package that called its part something else is re-emitted here rather than at its original name.
export const THEME_PART_PATH = 'xl/theme/theme1.xml';

/**
 * The extension of a part path (`xl/media/image1.JPEG` → `jpeg`), or `''` when it carries none.
 *
 * Lower-cased here rather than by each caller. OPC extensions are case-insensitive, so an extension is
 * only ever used as a key or compared against a literal, and every caller but one folded the case
 * itself; the one that did not handed a raw `JPEG` to `addImage`, which folds later, so the same media
 * type could reach `[Content_Types].xml` as two `<Default Extension>` entries. Folding once at the
 * source is what makes the missed fold unrepresentable rather than merely unlikely.
 */
export function extensionOf(partPath: string): string {
  const dot = partPath.lastIndexOf('.');
  const slash = partPath.lastIndexOf('/');
  return dot > slash ? partPath.slice(dot + 1).toLowerCase() : '';
}

// The relationships part path for `dir/name.ext` → `dir/_rels/name.ext.rels`.
export function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  const dir = slash === -1 ? '' : partPath.slice(0, slash + 1);
  const base = slash === -1 ? partPath : partPath.slice(slash + 1);
  return `${dir}_rels/${base}.rels`;
}

// The two directions of a relationship target, kept side by side because they are inverses: the
// writer names a part from the part that references it, and the reader turns that name back into a
// package path. A change to either that is not made to the other is a round trip that no longer
// round-trips, which is the property `part-paths.test.ts` pins across the pair.

// A relationship target expressed relative to the part that carries it: the `..` hops out of the
// referencing part's directory up to the common ancestor, then down to the target. Both paths are
// package-absolute (`xl/drawings/preservedP1.vml` → `xl/media/preservedP2.jpeg` → `../media/preservedP2.jpeg`).
export function relativePartPath(fromPath: string, toPath: string): string {
  const fromDir = fromPath.split('/').slice(0, -1);
  const toSegments = toPath.split('/');
  let common = 0;
  while (
    common < fromDir.length &&
    common < toSegments.length - 1 &&
    fromDir[common] === toSegments[common]
  ) {
    common++;
  }
  const up = fromDir.length - common;
  return [...Array<string>(up).fill('..'), ...toSegments.slice(common)].join('/');
}

// The inverse: resolve a relationship target (relative to the referencing part's directory, or
// absolute from the package root) into a package part path, collapsing `.`/`..` segments. Unlike
// {@link relativePartPath}, whose inputs the writer produced, this one reads a Target that came
// verbatim out of an untrusted package, so every OPC-legal shape a well-formed writer never emits
// still has to land on a bounded path. An absolute target goes through the same segment walk as a
// relative one: `/xl/../xl/worksheets/sheet1.xml` and `//xl/styles.xml` name real parts, and stripping
// only the leading slash left a key no part answers to.
export function resolveRelativePart(basePart: string, target: string): string {
  const joined = target.startsWith('/')
    ? target
    : `${basePart.slice(0, basePart.lastIndexOf('/') + 1)}${target}`;
  const out: string[] = [];
  for (const segment of joined.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') out.pop();
    else out.push(segment);
  }
  return out.join('/');
}
