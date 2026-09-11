# Worksheet Pictures

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `PixelAnchor`

<sub>type</sub>

The two shapes `WorksheetPictures.add` accepts, in caller-facing pixel units.

```ts
type PixelAnchor =
  | {readonly tl: AnchorPoint; readonly br: AnchorPoint; readonly editAs?: ImageEditAs}
  | {readonly tl: AnchorPoint; readonly ext: {readonly width: number; readonly height: number}};
```
