# Dark Mode repaints authored cell colours, and no encoding of a colour opts out

Cluster: styles

## Scenario

Excel Desktop's Dark Mode (View → Switch Modes) does not merely darken the application chrome. It
repaints the grid, including colours the package states explicitly. A deck authored to a brand
palette, white header text on a brand fill, does not render in those colours for a reader who has
the mode on.

The tempting inference is that this is a defect of _how_ the colour was written, and therefore that
some other encoding escapes it: a near-white instead of pure white, a theme slot instead of a
literal ARGB, an indexed palette entry instead of either. It is not. All five encodings of white
this library can emit land on the same repainted value.

## Measured behavior

Excel Desktop (Microsoft 365, version 16.0, build 20228.0, Windows). One frozen, autofiltered
header row; each column authors bold 9pt white a different way over the same `FF009EE0` fill, plus
three controls. Captured light, toggled to dark, captured again, both sampled from the rendered
pixels of the header band.

| Header font as authored                    | Fill as authored | Rendered light         | Rendered dark          |
| ------------------------------------------ | ---------------- | ---------------------- | ---------------------- |
| `argb FFFFFFFF`                            | `FF009EE0`       | `#FFFFFF` on `#009EE0` | `#262626` on `#0095D7` |
| `argb FFFEFEFE`                            | `FF009EE0`       | `#FEFEFE` on `#009EE0` | `#272727` on `#0095D7` |
| `theme 0` (lt1)                            | `FF009EE0`       | `#FFFFFF` on `#009EE0` | `#262626` on `#0095D7` |
| `theme 1, tint 1` (dk1 lightened to white) | `FF009EE0`       | `#FFFFFF` on `#009EE0` | `#262626` on `#0095D7` |
| `indexed 1` (palette white)                | `FF009EE0`       | `#FFFFFF` on `#009EE0` | `#262626` on `#0095D7` |
| _automatic_ (no colour stated)             | `FF009EE0`       | `#000000` on `#009EE0` | `#F6F6F6` on `#0095D7` |
| `argb FFFFFFFF`                            | `FF003C64`       | `#FFFFFF` on `#003C64` | `#262626` on `#B6D2FB` |
| `argb FF262626`                            | `FFFFFFFF`       | `#262626` on `#FFFFFF` | `#DCDCDC` on `#262626` |

Three facts fall out.

**The font inverts unconditionally.** Every encoding of white, whether literal, near-literal, theme
slot, tinted theme slot, or indexed, renders `#262626`. There is no way to state white and keep it.
The inversion is on the resolved colour, not on how it was spelled, which is why substituting one
encoding for another achieves nothing.

**The fill inverts too, but only when it is near an extreme.** `FF003C64` (dark navy) becomes
`#B6D2FB`; `FFFFFFFF` becomes `#262626`. A mid-luminance fill barely moves: `FF009EE0` renders
`#0095D7`, a dimming rather than an inversion. So fill and font do _not_ invert in step, and the
pairing that suffers is a stated font over a mid-luminance fill.

**An automatic font follows the mode, which is the point.** Where no colour is stated, black
becomes `#F6F6F6`, correct for a dark ground. Dark Mode is built for cells that state nothing;
what it does to cells that state something is the compromise.

## The cost is fidelity, not legibility

Worth stating plainly, because the reflex is to treat this as an accessibility regression. Contrast
survives the repaint, measured as WCAG ratios over the two brand pairings above:

| Pairing                       | Light   | Dark   |
| ----------------------------- | ------- | ------ |
| White on brand cyan `#009EE0` | 3.01:1  | 4.53:1 |
| White on brand navy `#003C64` | 11.47:1 | 9.80:1 |

The cyan pairing gets _better_; the navy one stays comfortably above threshold. Excel is choosing a
readable ink for the ground it has decided on. What it discards is the design: a deck built to a
palette renders in Excel's interpretation of that palette, and a reader in dark mode is not seeing
the colours anyone chose.

The mode's dimming of a mid-luminance fill is small enough (`#009EE0` to `#0095D7`) that a fill
authored _light_ with white text on it, a combination that is already fragile, is where this
would genuinely collapse. Not measured here; flagged as the shape to watch.

## What follows for the writer

**Nothing.** Dark Mode is a view setting on the reader's installation, not a property of the
package, and there is no colour encoding that survives it. The table above is the whole search
space this library can address. An agent that finds a deck "rendering in the wrong colours" should
check the reader's Switch Modes state before touching a style.

The only advice that helps is to the reader: toggle Dark Mode off (View → Switch Modes) to see a
workbook in its authored colours. Nothing needs to change in the file to make that work.

## Provenance

`source: excel-desktop-verification`, the tier ADR-0013 describes. One Excel build, one host.

The rendering tier is required here: headless COM paints nothing, and Excel exposes no COM
property for the mode. Probed, `Application`, `Window`, `Workbook` and `Worksheet` carry no
dark or appearance member, so the toggle is reachable only through the ribbon, KeyTip `M1` on the View
tab. Taken by driving a visible Excel through the `excel-gui-automation` skill, then sampling the
header band's pixels light and dark. A recorded observation, not reproducible from the repo; the
probe workbook was authored by a scratch script under `.tmp/`.

Related: `theme-color-index-order`, `default-font-must-not-be-assumed-for-column-widths`,
`frozen-pane-header-ink-is-an-excel-repaint-fault`.
