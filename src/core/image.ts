// Anchored images: the model for a picture pinned to a worksheet's grid.
//
// The image *bytes* live once on the workbook (a small media registry, addressed by a numeric id);
// a worksheet then anchors that image to a rectangle of cells. Storing the bytes centrally means the
// same picture used on two sheets (a logo in a header band, say) is one media part, not two.

import {hex} from '../hex.ts';
import {sha512} from '../sha512.ts';
import {tokenSet} from '../token-set.ts';

/** A point in the drawing grid: a 0-based column and row, plus an EMU offset into that cell.
 * The offsets default to zero, pinning the point to the cell's top-left corner. */
export interface AnchorPoint {
  /** 0-based column index (column A is 0). */
  readonly col: number;
  /** 0-based row index (row 1 is 0). */
  readonly row: number;
  /** Horizontal offset into the cell, in EMUs (914400 per inch). Defaults to 0. */
  readonly colOff?: number;
  /** Vertical offset into the cell, in EMUs. Defaults to 0. */
  readonly rowOff?: number;
}

/** EMUs per pixel at Excel's notional 96 DPI (914400 EMU/inch ÷ 96 px/inch). The conversion is
 * DPI-independent by construction: a pixel extent is a fixed physical size regardless of screen. */
export const PX_TO_EMU = 9525;

/** How a two-cell-anchored image tracks edits to the cells it spans. `twoCell` moves and resizes with
 * them; `oneCell` moves but keeps its size; `absolute` is pinned to the page and does neither. The
 * schema default, which a file omitting the attribute means, is `twoCell`; an image authored here
 * without one is written as `oneCell`, which is this library's own default. */
export type ImageEditAs = 'oneCell' | 'twoCell' | 'absolute';

/** Narrow a raw `<xdr:twoCellAnchor editAs>` token to a known {@link ImageEditAs}. */
export const isImageEditAs = tokenSet<ImageEditAs>({oneCell: true, twoCell: true, absolute: true});

/** A fixed image size in EMUs: the extent of a one-cell anchor, which pixel dimensions convert into
 * via {@link PX_TO_EMU}. */
export interface Extent {
  readonly cx: number;
  readonly cy: number;
}

/** A two-cell anchor: the image's top-left (`from`) and bottom-right (`to`) grid points. The image
 * fills the rectangle between them and reflows as the intervening rows/columns resize; `editAs`
 * selects how strictly it follows. */
export interface TwoCellAnchor {
  readonly from: AnchorPoint;
  readonly to: AnchorPoint;
  readonly editAs?: ImageEditAs;
  /** Clockwise rotation in 1/60000 of a degree (`2700000` = 45°), preserved from a loaded file. */
  readonly rotation?: number;
}

/** A one-cell anchor: a single top-left grid point plus a fixed extent. The image keeps its size as
 * the grid resizes, moving only with its anchor cell. `editAs` is a two-cell-only attribute and has
 * no place here. */
export interface OneCellAnchor {
  readonly from: AnchorPoint;
  readonly ext: Extent;
  /** Clockwise rotation in 1/60000 of a degree (`2700000` = 45°), preserved from a loaded file. */
  readonly rotation?: number;
}

/** Where an image sits on the grid: a rectangle between two cells, or a point plus a fixed extent. */
export type ImageAnchor = TwoCellAnchor | OneCellAnchor;

/** Narrow an anchor to its one-cell (fixed-extent) form; the complement is {@link TwoCellAnchor}. */
export function isOneCellAnchor(anchor: ImageAnchor): anchor is OneCellAnchor {
  return 'ext' in anchor;
}

// Sub-cell anchor geometry. Excel measures a column in characters of the default font (~7 px each at
// 96 DPI) and a row in points (1/72 inch); a column or row that sets no size falls back to Excel's
// own defaults. These constants live here, beside the anchor model they serve, rather than in the
// Worksheet that merely supplies the per-column/row sizes.
const CHAR_WIDTH_PX = 7;
const EMU_PER_POINT = 12700;
const DEFAULT_COL_WIDTH_CHARS = 8.43;
const DEFAULT_ROW_HEIGHT_POINTS = 15;

/** A column's width in characters of the default font, or `undefined` to take Excel's default. */
export type ColumnWidthLookup = (col: number) => number | undefined;
/** A row's height in points, or `undefined` to take Excel's default. */
export type RowHeightLookup = (row: number) => number | undefined;

/** Resolve a possibly-fractional anchor point to the cell it floors to plus a sub-cell EMU offset
 * scaled by that cell's real width/height, so `col: 3.5` lands halfway across column 3 regardless of
 * the column's size. An already-integer point keeps a zero offset (unless one was given). The two
 * lookups supply each column/row's size; a size they leave `undefined` falls back to Excel's default. */
export function resolveAnchorPoint(
  point: AnchorPoint,
  columnWidth: ColumnWidthLookup,
  rowHeight: RowHeightLookup,
): AnchorPoint {
  const col = Math.floor(point.col);
  const row = Math.floor(point.row);
  const colWidthEmu = Math.round(
    (columnWidth(col) ?? DEFAULT_COL_WIDTH_CHARS) * CHAR_WIDTH_PX * PX_TO_EMU,
  );
  const rowHeightEmu = Math.round((rowHeight(row) ?? DEFAULT_ROW_HEIGHT_POINTS) * EMU_PER_POINT);
  const colOff = (point.colOff ?? 0) + Math.round((point.col - col) * colWidthEmu);
  const rowOff = (point.rowOff ?? 0) + Math.round((point.row - row) * rowHeightEmu);
  return {col, row, colOff, rowOff};
}

/** An image pinned to a worksheet: which workbook media it shows (`imageId`) and where. */
export interface AnchoredImage {
  /** Index into the workbook's media registry (the id {@link Workbook.addImage} returned). */
  readonly imageId: number;
  readonly anchor: ImageAnchor;
}

/** A picture's bytes and its file kind, as held in the workbook's media registry. */
export interface WorkbookImage {
  /** Lower-case file extension without a dot: `"png"`, `"jpeg"`, `"gif"`. Drives the media part's
   * name and content type. */
  readonly extension: string;
  readonly data: Uint8Array;
}

// Leading magic bytes for the raster formats a spreadsheet embeds, most-specific first. Used to infer
// an extension when the caller supplies none, so a package never declares an `image/undefined` type.
const IMAGE_MAGIC: ReadonlyArray<{readonly ext: string; readonly sig: readonly number[]}> = [
  {ext: 'png', sig: [0x89, 0x50, 0x4e, 0x47]},
  {ext: 'jpeg', sig: [0xff, 0xd8, 0xff]},
  {ext: 'gif', sig: [0x47, 0x49, 0x46]},
  {ext: 'bmp', sig: [0x42, 0x4d]},
  {ext: 'tiff', sig: [0x49, 0x49, 0x2a, 0x00]},
  {ext: 'tiff', sig: [0x4d, 0x4d, 0x00, 0x2a]},
];

function sniffImageExtension(data: Uint8Array): string {
  for (const {ext, sig} of IMAGE_MAGIC) {
    if (sig.every((b, i) => data[i] === b)) return ext;
  }
  // An unrecognised blob still needs a valid media name and content type; png keeps the package
  // well-formed rather than emitting a `<Default>` with no or a bogus extension.
  return 'png';
}

/** Reduce a caller-supplied extension to the bare, lower-case alphanumeric token OOXML expects for a
 * media part's name and `<Default Extension>`. A leading dot (`".png"`), a URL query string
 * (`"png?alt=media"`), or any other separator a real-world filename/URL drags in is stripped to the
 * leading run of alphanumerics; a missing or all-punctuation hint falls back to sniffing the bytes'
 * magic number, so the package is always well-formed. */
export function normalizeImageExtension(extension: string | undefined, data: Uint8Array): string {
  if (typeof extension === 'string') {
    const token = extension.toLowerCase().match(/[a-z0-9]+/)?.[0];
    if (token !== undefined) return token;
  }
  return sniffImageExtension(data);
}

/**
 * An anchored image in workbook-independent form: the picture's own bytes rather than a media id
 * into one particular workbook's registry.
 *
 * An {@link AnchoredImage} means nothing away from that registry: its `imageId` is an index, and
 * the same index names a different picture (or none) in the next workbook. Attaching the picture
 * itself is what lets an anchor cross that boundary, which is why the transfer form carries bytes
 * where the stored form carries an id.
 */
export interface PortableImage {
  readonly image: WorkbookImage;
  readonly anchor: ImageAnchor;
}

/**
 * Every picture a worksheet shows, in the workbook-independent form of {@link PortableImage}: the
 * images anchored to the grid, in the order they were added, and the background tiled behind it.
 * {@link Workbook.exportImages} produces one, {@link Workbook.importImages} applies one.
 */
export interface WorksheetImages {
  readonly anchored: readonly PortableImage[];
  readonly background: WorkbookImage | undefined;
}

/**
 * A picture's content identity as a map key: its kind, its length, and a digest of its bytes.
 *
 * **Content-addressed, not reference-addressed.** Two byte-identical pictures are one picture,
 * however they reached the registry. That is what keeps repeated imports from growing the media list
 * without bound, so the same logo carried onto twenty sheets registers once, and it is why identity
 * cannot be object identity: a picture arriving from another workbook would never satisfy that.
 *
 * **A key rather than a scan**, which is the part that changed. The comparison used to walk the whole
 * media list byte by byte per candidate, and importing a sheet's pictures asks it once per anchored
 * image, so merging a workbook of fifty distinct megabyte images compared bytes fifty times over
 * fifty candidates. The workbook indexes each picture as it is registered instead.
 *
 * SHA-512 rather than a cheap checksum because a collision here silently substitutes one picture for
 * another; the length is in the key as well, so a collision would have to match that too.
 */
export function imageContentKey(image: WorkbookImage): string {
  let digest = '';
  for (const byte of sha512(image.data)) digest += hex(byte, 2);
  return `${image.extension}:${image.data.length}:${digest}`;
}
