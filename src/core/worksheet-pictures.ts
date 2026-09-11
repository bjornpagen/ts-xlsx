// The pictures on a worksheet, lifted off `Worksheet` into the slice they are.
//
// Named for the pictures rather than for the images because `WorksheetImages` is already taken, by
// the public value type `Workbook.exportImages` hands back: a sheet's pictures in a
// workbook-independent form. Two things in `core/` under one name would be a standing invitation to
// import the wrong one.
//
// Two kinds, and they are not the same thing: an *anchored* image is pinned to a grid rectangle or
// a grid point and moves with the cells under it, while the *background* tiles behind the whole
// sheet and is anchored to nothing. Both are held here because both are one opaque workbook image
// id that this class never dereferences: what an id names lives on `Workbook`, and the media
// registry is what resolves it.
//
// That opacity is what makes this a slice. Its only edge outside itself is `resolveAnchorPoint`,
// which needs to know how wide a column is and how tall a row is to turn a grid point into the
// model's own EMU units; those two arrive as accessors, the way `WorkbookTheme` takes the indexed
// palette it does not own. See the delegation rule in `docs/architecture.md`: the public accessors
// stay on `Worksheet` with their doc comments and become one-line calls into this.

import {replaceContents} from './containers.ts';
import {
  type AnchoredImage,
  type AnchorPoint,
  type Extent,
  type ImageAnchor,
  type ImageEditAs,
  type PictureProperties,
  pictureProperties,
  PX_TO_EMU,
  resolveAnchorPoint,
  type TwoCellAnchor,
} from './image.ts';

/** The two shapes {@link WorksheetPictures.add} accepts, in caller-facing pixel units. */
export type PixelAnchor =
  | {readonly tl: AnchorPoint; readonly br: AnchorPoint; readonly editAs?: ImageEditAs}
  | {readonly tl: AnchorPoint; readonly ext: {readonly width: number; readonly height: number}};

/** How a sheet reports the sizes an anchor point resolves against; `undefined` means "the default". */
export interface AnchorMetrics {
  columnWidth(col: number): number | undefined;
  rowHeight(row: number): number | undefined;
}

export class WorksheetPictures {
  readonly #metrics: AnchorMetrics;

  // Mutable and handed out whole to `GridEdits`, which re-pins every anchor in place on a splice.
  // A copy would leave the sheet's images behind the grid they sit on.
  readonly #anchors: AnchoredImage[] = [];

  // A sheet background is a single workbook image tiled behind the grid, distinct from an anchored
  // drawing: it has no anchor and rides its own worksheet relationship, not a drawing part.
  #backgroundImageId: number | undefined;

  constructor(metrics: AnchorMetrics) {
    this.#metrics = metrics;
  }

  /** The anchored images, in the order they were added. */
  get anchors(): AnchoredImage[] {
    return this.#anchors;
  }

  get backgroundImageId(): number | undefined {
    return this.#backgroundImageId;
  }

  // Bind the pure anchor geometry to this sheet's per-column/row sizes; a size a column or row does
  // not set defers to the sheet default, then (inside resolveAnchorPoint) to Excel's own default.
  add(imageId: number, anchor: PixelAnchor, properties: PictureProperties = {}): void {
    const picture = pictureProperties(properties);
    const columnWidth = (col: number): number | undefined => this.#metrics.columnWidth(col);
    const rowHeight = (row: number): number | undefined => this.#metrics.rowHeight(row);
    if ('ext' in anchor) {
      const ext: Extent = {
        cx: Math.round(anchor.ext.width * PX_TO_EMU),
        cy: Math.round(anchor.ext.height * PX_TO_EMU),
      };
      const from = resolveAnchorPoint(anchor.tl, columnWidth, rowHeight);
      this.#anchors.push({imageId, anchor: {from, ext}, ...picture});
      return;
    }
    const from = resolveAnchorPoint(anchor.tl, columnWidth, rowHeight);
    const to = resolveAnchorPoint(anchor.br, columnWidth, rowHeight);
    const twoCell: TwoCellAnchor =
      anchor.editAs !== undefined ? {from, to, editAs: anchor.editAs} : {from, to};
    this.#anchors.push({imageId, anchor: twoCell, ...picture});
  }

  addAnchor(imageId: number, anchor: ImageAnchor, properties: PictureProperties = {}): void {
    this.#anchors.push({imageId, anchor, ...pictureProperties(properties)});
  }

  remove(imageId: number): void {
    replaceContents(
      this.#anchors,
      this.#anchors.filter((image) => image.imageId !== imageId),
    );
  }

  setBackground(imageId: number | undefined): void {
    this.#backgroundImageId = imageId;
  }
}
