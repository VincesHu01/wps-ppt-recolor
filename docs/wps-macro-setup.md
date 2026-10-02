# Installing the WPS macro

## Install (once)

1. Open a presentation in WPS Presentation → **Developer** tab.
   If the tab is missing: Settings → General and Save → enable "Show Developer tab".
   Personal editions may ask you to install the macro component first.
2. **Developer → Macro Editor** → right-click the project tree → **Insert → Module**.
3. Paste the whole of [`scripts/wps_macro.js`](../scripts/wps_macro.js) into the code window.
4. Save.

To make it a button: File → Options → Quick Access Toolbar → add `recolorSelection` from the
macros category. The script also exports a localized alias `PPT配色自适应` if you prefer a name
that reads naturally in the toolbar.

## Entry points

| Function | What it does |
| --- | --- |
| `recolorSelection()` | Recolors the selection; if nothing is selected, the last inserted object |
| `recolorLastInserted()` | Always the last inserted object |
| `recolorWholeSlide()` | Everything on the current slide |
| `showSlidePalette()` | Reports the extracted palette and changes nothing |

Run `showSlidePalette()` first. It reports the primary, the color order, the neutral range and the
theme accents, so you can confirm the sampling before anything is modified.

## Configuration

Top of the script:

```js
var CFG = { depth: 0.85, useThemeFirst: false, keepWhiteBlack: true };
```

| Field | Effect |
| --- | --- |
| `depth` | Lightness derivation strength. 0.6 is gentle, 1.1 is pronounced |
| `useThemeFirst` | When true, the primary comes from theme `accent1` rather than the most used slide color |
| `keepWhiteBlack` | Preserve pure white and pure black |

## Version compatibility

WPS editions do not all expose the same JSAPI surface, so every optional property is wrapped and
degrades quietly:

| Feature | Primary path | Fallback |
| --- | --- | --- |
| Gradients | `Fill.GradientColorStops`, stop by stop | Two-color gradient via `ForeColor` / `BackColor` |
| Theme colors | `Slide.ThemeColorScheme` | `Slide.ColorScheme`, then scanning layout shapes |
| Charts | `Shape.Chart.SeriesCollection` | Skipped |
| SmartArt | Not addressable through JSAPI | Use the Python engine, which rewrites `diagrams/colorsN.xml` |
| Tables | `Shape.Table.Cell(r, c).Shape.Fill` | Skipped |

The result dialog states how many colors changed and how many were skipped. A non-zero skip count
means that edition could not reach those properties and the Python engine should handle it.

## Channel order for RGB

WPS follows the VBA convention: the `RGB` property is `R + G * 256 + B * 65536`, so red is the low
byte. If a build returns colors with red and blue swapped, swap the bodies of `vbaR` and `vbaB`.

## Troubleshooting

| Symptom | Fix |
| --- | --- |
| "Cannot get the current slide" | Switch to Normal view, not Slide Sorter or Slide Show |
| Dialog reports 0 changes | The object is bitmaps or theme references; use the Python engine |
| Colors come out wrong | Run `showSlidePalette()` first; if sampling is off, use `--exclude` or `useThemeFirst` |
| Function missing from the macro list | Module not saved, or macro security is blocking macros |
| Undo | `Ctrl+Z`, every step is an undoable edit |
