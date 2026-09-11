// Anchored images on the wire: the `xl/drawings/drawing{n}.xml` part (a DrawingML two-cell anchor
// per image) and the reader that turns a drawing back into anchors. The drawing's relationships to
// the `xl/media/` bytes are recorded by the writer as it plans each image's embed id. The image bytes
// themselves are opaque here: the writer copies them verbatim into a media part and the reader hands
// them back untouched.

import {
  type AnchorPoint,
  type Extent,
  type ImageAnchor,
  type ImageCrop,
  type ImageEditAs,
  isImageEditAs,
  isOneCellAnchor,
  type PictureProperties,
} from '../../core/image.ts';
import {assertWritableNumber, AuthoringError} from '../../errors.ts';
import {enumToken, numFinite, numInteger} from '../../xml/xml-attrs.ts';
import {parseXml, TextCapture} from '../../xml/xml-read.ts';
import {localName} from '../../xml/xml-scan.ts';
import {
  checkedToken,
  escapeAttr,
  intAttr,
  numAttr,
  numberText,
  textAttr,
  XML_DECLARATION,
} from '../../xml/xml.ts';
import {relAttr, RELATIONSHIPS_NS} from '../opc/namespaces.ts';
import {DRAWINGML_NS, XDR_NS} from './namespaces.ts';

// The content type Excel expects for each image kind, keyed by lower-case extension. An unlisted
// extension falls back to `image/<ext>`, which is what a well-behaved consumer infers anyway.
// A Map, not an object literal: the extension comes off a media part's name in the package, and an
// object would answer `constructor` with a function that then stringifies into a content-type
// attribute.
const IMAGE_CONTENT_TYPES: ReadonlyMap<string, string> = new Map([
  ['png', 'image/png'],
  ['jpg', 'image/jpeg'],
  ['jpeg', 'image/jpeg'],
  ['gif', 'image/gif'],
  ['bmp', 'image/bmp'],
  ['tif', 'image/tiff'],
  ['tiff', 'image/tiff'],
  ['emf', 'image/x-emf'],
  ['wmf', 'image/x-wmf'],
  ['svg', 'image/svg+xml'],
]);

/** The content type for a media part's `<Default Extension>` entry in `[Content_Types].xml`. */
export function imageContentType(extension: string): string {
  const ext = extension.toLowerCase();
  return IMAGE_CONTENT_TYPES.get(ext) ?? `image/${ext}`;
}

/** One image placed in a drawing: where it sits, the drawing-local relationship id that ties it to its
 * media bytes, and the picture's own properties. */
export interface DrawingImage {
  readonly anchor: ImageAnchor;
  /** The `r:embed` id referencing this image's entry in the drawing's own `.rels`. */
  readonly embedId: string;
  readonly properties: PictureProperties;
  /** The `r:id` of the hyperlink relationship in the drawing's own `.rels`, when the picture is a link. */
  readonly hyperlinkId?: string;
}

/** The `xl/drawings/drawing{n}.xml` part: one anchor per image, two-cell or one-cell by its shape. */
export function drawingXml(images: readonly DrawingImage[]): string {
  const anchors = images.map((image, i) => anchorXml(image, i + 1)).join('');
  return (
    XML_DECLARATION +
    `<xdr:wsDr xmlns:xdr="${XDR_NS}" xmlns:a="${DRAWINGML_NS}" xmlns:r="${RELATIONSHIPS_NS}">` +
    anchors +
    '</xdr:wsDr>'
  );
}

function anchorXml(image: DrawingImage, id: number): string {
  const {anchor} = image;
  const pic = picXml(image, id, anchor.rotation);
  return isOneCellAnchor(anchor)
    ? oneCellAnchorXml(anchor.from, anchor.ext, pic)
    : twoCellAnchorXml(anchor.from, anchor.to, anchor.editAs ?? 'oneCell', pic);
}

// A picture anchored between two grid points. The geometry lives entirely in <xdr:from>/<xdr:to>, so
// the picture carries no absolute <a:xfrm>: a zeroed one would override the anchor and collapse the
// image to nothing in strict viewers (LibreOffice), while a non-zero one would fight the anchor. A
// rotation is the one transform kept: it can't be derived from the anchor, so it rides a rot-only xfrm.
function twoCellAnchorXml(
  from: AnchorPoint,
  to: AnchorPoint,
  editAs: ImageEditAs,
  pic: string,
): string {
  return (
    `<xdr:twoCellAnchor editAs="${checkedToken(editAs, isImageEditAs, 'image anchor edit mode')}">` +
    `<xdr:from>${anchorPointXml(from)}</xdr:from>` +
    `<xdr:to>${anchorPointXml(to)}</xdr:to>` +
    pic +
    '<xdr:clientData/>' +
    '</xdr:twoCellAnchor>'
  );
}

// A picture pinned at one grid point with a fixed EMU extent. editAs is a two-cell-only attribute and
// the schema forbids it here, so a one-cell anchor never carries one.
function oneCellAnchorXml(from: AnchorPoint, ext: Extent, pic: string): string {
  return (
    '<xdr:oneCellAnchor>' +
    `<xdr:from>${anchorPointXml(from)}</xdr:from>` +
    `<xdr:ext${numAttr('cx', ext.cx)}${numAttr('cy', ext.cy)}/>` +
    pic +
    '<xdr:clientData/>' +
    '</xdr:oneCellAnchor>'
  );
}

// The picture itself, spelled as Excel 16.0 spells what a picture says about itself: alternative text and a
// title as `descr` and `title` on `cNvPr`, a link as an `a:hlinkClick` inside it naming a hyperlink
// relationship of the drawing, and a crop as an `a:srcRect` between the blip and its stretch.
function picXml(image: DrawingImage, id: number, rotation: number | undefined): string {
  const {embedId, properties, hyperlinkId} = image;
  const xfrm = rotation !== undefined ? `<a:xfrm${numAttr('rot', rotation)}/>` : '';
  const attributes =
    textAttr('descr', properties.description) + textAttr('title', properties.title);
  const link =
    hyperlinkId === undefined
      ? ''
      : `<a:hlinkClick r:id="${escapeAttr(hyperlinkId)}"` +
        `${textAttr('tooltip', properties.hyperlink?.tooltip)}/>`;
  const cNvPr =
    link === ''
      ? `<xdr:cNvPr id="${id}" name="Picture ${id}"${attributes}/>`
      : `<xdr:cNvPr id="${id}" name="Picture ${id}"${attributes}>${link}</xdr:cNvPr>`;
  return (
    '<xdr:pic>' +
    `<xdr:nvPicPr>${cNvPr}` +
    '<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>' +
    `<xdr:blipFill><a:blip r:embed="${embedId}"/>` +
    srcRectXml(properties.crop) +
    '<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
    `<xdr:spPr>${xfrm}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>` +
    '</xdr:pic>'
  );
}

// A crop edge is a fraction of the picture in the model and an `xsd:int` in thousandths of a percent on
// the wire, where `100000` is the whole picture. An edge that crops nothing is omitted, as Excel omits it.
const CROP_UNITS = 100_000;
const INT_MAX = 2_147_483_647;

function srcRectXml(crop: ImageCrop | undefined): string {
  if (crop === undefined) return '';
  const edge = (name: string, fraction: number | undefined): string => {
    if (fraction === undefined) return '';
    assertWritableNumber(fraction);
    const units = Math.round(fraction * CROP_UNITS);
    if (units > INT_MAX) {
      throw new AuthoringError(`cannot write the picture crop ${name}="${units}": past xsd:int`);
    }
    return units === 0 ? '' : intAttr(name, units, -INT_MAX - 1);
  };
  const attributes =
    edge('l', crop.left) + edge('t', crop.top) + edge('r', crop.right) + edge('b', crop.bottom);
  return attributes === '' ? '' : `<a:srcRect${attributes}/>`;
}

// Read a crop edge in either spelling `ST_Percentage` allows: thousandths of a percent (`9999`), or a
// percentage with its sign (`10%`).
function cropEdge(value: string | undefined): number | undefined {
  if (value === undefined) return undefined;
  const percent = value.endsWith('%') ? numFinite(value.slice(0, -1)) : undefined;
  const units = percent === undefined ? numInteger(value) : Math.round(percent * 1000);
  // Bounded to `xsd:int`, so a crop read from a file is always one the writer can put back.
  if (units === undefined || units === 0 || units > INT_MAX || units < -INT_MAX - 1) {
    return undefined;
  }
  return units / CROP_UNITS;
}

// The grid point's four numbers are all author-reachable through `addImageAnchor`, and each is an
// `xsd:int` or an EMU offset with no spelling for a non-finite value, so they are refused on the same
// terms as the extent above. That they are elements rather than attributes changes nothing.
function anchorPointXml(point: AnchorPoint): string {
  return (
    `<xdr:col>${numberText(point.col)}</xdr:col>` +
    `<xdr:colOff>${numberText(point.colOff ?? 0)}</xdr:colOff>` +
    `<xdr:row>${numberText(point.row)}</xdr:row>` +
    `<xdr:rowOff>${numberText(point.rowOff ?? 0)}</xdr:rowOff>`
  );
}

/** An image anchor parsed from a drawing part, with the `r:embed` id that names its media. A two-cell
 * anchor carries `to` and its `editAs`, the schema default filled in; a one-cell anchor carries `ext`
 * instead. */
export interface ParsedImageAnchor {
  readonly from: AnchorPoint;
  readonly to?: AnchorPoint;
  readonly ext?: Extent;
  readonly editAs?: ImageEditAs;
  readonly rotation?: number;
  readonly embed: string;
  readonly description?: string;
  readonly title?: string;
  readonly crop?: ImageCrop;
  /** The `r:id` of the picture's `a:hlinkClick`, for the caller to resolve against the drawing's rels. */
  readonly hyperlinkId?: string;
  readonly tooltip?: string;
}

type PointDraft = {col: number; row: number; colOff: number; rowOff: number};

function blankPoint(): PointDraft {
  return {col: 0, row: 0, colOff: 0, rowOff: 0};
}

/** A drawing part as the reader sees it: the picture anchors it could model, and whether those anchors
 * are the whole of it. */
export interface ParsedDrawing {
  readonly anchors: readonly ParsedImageAnchor[];
  /**
   * Whether every anchor is a two-cell or one-cell anchor holding exactly one embedded picture. When it
   * is not, writing the drawing back from `anchors` would leave the rest out, so the reader keeps the
   * whole part byte for byte instead.
   */
  readonly fullyModeled: boolean;
}

// What a drawing can hold that the image model cannot write back: a chart (`graphicFrame`), a shape or
// text box (`sp`), a connector (`cxnSp`), a group (`grpSp`), ink (`contentPart`), an anchor pinned to
// the page rather than the grid (`absoluteAnchor`), and a markup-compatibility wrapper
// (`AlternateContent`), whose branches spell one object twice.
const UNMODELED_CONTENT: ReadonlySet<string> = new Set([
  'graphicFrame',
  'sp',
  'cxnSp',
  'grpSp',
  'contentPart',
  'absoluteAnchor',
  'AlternateContent',
]);

/**
 * Parse a drawing part into its picture anchors (`<xdr:twoCellAnchor>` and `<xdr:oneCellAnchor>`), and
 * decide in the same scan whether they are the whole drawing. It is fully modeled only when every
 * anchor holds exactly one picture whose bytes are embedded; anything else in it, including a picture
 * that links its image rather than embedding it, has no place in the image model.
 *
 * This used to be two scans: one listing the content the model skips, and this one. The list missed
 * an absolute anchor and a linked picture, so a drawing holding either counted as modeled and lost them
 * on write. What an anchor holds is now judged by what the parse found in it, one embedded picture or
 * not, so no kind of anchor content can slip past unnamed; the list left names only the anchor kinds
 * and wrappers this parse does not otherwise open an anchor for.
 */
export function parseDrawing(xml: string): ParsedDrawing {
  const anchors: ParsedImageAnchor[] = [];
  let fullyModeled = true;
  let from: PointDraft | null = null;
  let to: PointDraft | null = null;
  let ext: Extent | undefined;
  let editAs: ImageEditAs | undefined;
  let rotation: number | undefined;
  let embed: string | undefined;
  let pictures = 0;
  // What the picture says about itself, gathered from its `cNvPr`, `hlinkClick` and `srcRect`.
  let properties: {
    description?: string;
    title?: string;
    crop?: ImageCrop;
    hyperlinkId?: string;
    tooltip?: string;
  } = {};
  // The point (<xdr:from> or <xdr:to>) whose coordinate children are currently streaming in.
  let target: PointDraft | null = null;
  // Depth inside <xdr:pic>, so the anchor-level <xdr:ext> is not confused with the <a:ext> nested in
  // a picture's spPr transform (both have local name "ext").
  let picDepth = 0;
  // The coordinate child currently streaming in, so its text lands on the right field.
  const coord = new TextCapture(COORDINATES);

  parseXml(xml, {
    onOpen(name, attrs, selfClosing, scope) {
      const local = localName(name);
      if (UNMODELED_CONTENT.has(local)) {
        fullyModeled = false;
      } else if (local === 'twoCellAnchor' || local === 'oneCellAnchor') {
        from = blankPoint();
        to = local === 'twoCellAnchor' ? blankPoint() : null;
        ext = undefined;
        rotation = undefined;
        embed = undefined;
        pictures = 0;
        properties = {};
        // `twoCell` is the schema default, so it is what a file omitting the attribute means, and what
        // one spelling it with a token outside the enumeration is read as. Left undefined, it was
        // written back with this library's authoring default, `oneCell`, and the picture stopped
        // resizing with its cells.
        editAs =
          local === 'twoCellAnchor'
            ? (enumToken(attrs.editAs, isImageEditAs) ?? 'twoCell')
            : undefined;
      } else if (local === 'pic') {
        picDepth++;
        pictures++;
      } else if (local === 'cNvPr' && picDepth > 0) {
        if (attrs.descr !== undefined && attrs.descr !== '') properties.description = attrs.descr;
        if (attrs.title !== undefined && attrs.title !== '') properties.title = attrs.title;
      } else if (local === 'hlinkClick' && picDepth > 0) {
        const id = relAttr(scope, attrs, 'id');
        if (id !== undefined && id !== '') {
          properties.hyperlinkId = id;
          if (attrs.tooltip !== undefined) properties.tooltip = attrs.tooltip;
        }
      } else if (local === 'srcRect' && picDepth > 0) {
        const crop: {-readonly [K in keyof ImageCrop]?: number} = {};
        const left = cropEdge(attrs.l);
        const top = cropEdge(attrs.t);
        const right = cropEdge(attrs.r);
        const bottom = cropEdge(attrs.b);
        if (left !== undefined) crop.left = left;
        if (top !== undefined) crop.top = top;
        if (right !== undefined) crop.right = right;
        if (bottom !== undefined) crop.bottom = bottom;
        if (Object.keys(crop).length > 0) properties.crop = crop;
      } else if (local === 'xfrm' && picDepth > 0) {
        // The picture's own rotation: the one spPr transform that can't be derived from the anchor.
        const rot = numFinite(attrs.rot);
        if (rot !== undefined && rot !== 0) rotation = rot;
      } else if (local === 'from') {
        target = from;
      } else if (local === 'to') {
        target = to;
      } else if (local === 'ext' && picDepth === 0) {
        const cx = numFinite(attrs.cx, 0);
        const cy = numFinite(attrs.cy, 0);
        if (cx !== undefined && cy !== undefined) ext = {cx, cy};
      } else if (local === 'blip') {
        // Resolved by namespace, so a drawing binding the relationships namespace to any prefix is
        // read. This site used to hedge with `attrs['r:embed'] ?? attrs.embed`, which was the one
        // place the problem had been noticed and the hedge caught only the unprefixed spelling, which
        // is not in the namespace at all.
        const value = relAttr(scope, attrs, 'embed');
        if (value !== undefined) embed = value;
        // A linked picture names its image through a relationship the model holds no field for.
        if (relAttr(scope, attrs, 'link') !== undefined) fullyModeled = false;
      } else if (target !== null) {
        coord.open(local, selfClosing);
      }
    },
    onText(chunk) {
      coord.text(chunk);
    },
    onClose(name) {
      const local = localName(name);
      const text = coord.close(local);
      if (text !== undefined) {
        const value = numFinite(text);
        if (target !== null && value !== undefined) setCoordinate(target, local, value);
      } else if (local === 'from' || local === 'to') {
        target = null;
      } else if (local === 'pic') {
        picDepth--;
      } else if (local === 'twoCellAnchor' || local === 'oneCellAnchor') {
        const rot = rotation !== undefined ? {rotation} : {};
        if (from === null || embed === undefined || pictures !== 1) {
          fullyModeled = false;
        } else if (to !== null) {
          const mode = editAs !== undefined ? {editAs} : {};
          anchors.push({from: {...from}, to: {...to}, ...mode, ...rot, embed, ...properties});
        } else if (ext !== undefined) {
          anchors.push({from: {...from}, ext, ...rot, embed, ...properties});
        } else {
          fullyModeled = false;
        }
        from = null;
        to = null;
        ext = undefined;
      }
    },
  });
  return {anchors, fullyModeled};
}

const COORDINATES = new Set<string>(['col', 'colOff', 'row', 'rowOff']);

function setCoordinate(point: PointDraft, coord: string, value: number): void {
  if (coord === 'col') point.col = value;
  else if (coord === 'colOff') point.colOff = value;
  else if (coord === 'row') point.row = value;
  else point.rowOff = value;
}
