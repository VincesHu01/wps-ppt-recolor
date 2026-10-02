# Color derivation rules

`scripts/recolor.py` and `scripts/wps_macro.js` implement the same rules, so both channels
produce the same result.

## 1. Sampling

Two tiers, current slide first:

1. **Current slide.** All shapes except the ones being recolored and anything passed to
   `--exclude`.
2. **Backup.** When the slide yields fewer than two families, sampling widens to the slide layout
   (`--scope layout`) or every slide (`--scope deck`), then to the theme's `accent1..6`.

Weights are `role x area`:

| Role | Weight |
| --- | --- |
| Fill | 1.0 |
| Gradient stop | 0.75 |
| Line | 0.35 |
| Text | 0.25 |

Area contributes as `(area / 1e12) ** 0.45`. The sublinear exponent stops one full-bleed
background from drowning out every content color.

### Borrowing a theme accent

Accents are **not** taken in `accent1, accent2, ...` order. Candidates are scored by their minimum
hue distance to the families already found, sorted ascending, and accepted only if that distance
is above 25 degrees (distinct enough to be a separate family) and at most `--harmony-max`
(90 degrees by default).

That last rule exists because of a real failure: a deck using only deep blue and teal had
`accent2 = orange` in its theme file, and the naive version cheerfully introduced orange, producing
a new mismatch in the middle of fixing one. If nothing is harmonious enough, the extra family is
derived from the primary instead.

## 2. Clustering

- **Neutrals are detected by absolute RGB chroma, not HSL saturation.** This is the single most
  important detail in the whole tool. `#EAF1F9` — an extremely common light panel background — has
  an HSL saturation of 0.56, so an `S < 0.12` test classifies it as a vivid color and it then
  dominates the weighted mean, pulling the primary somewhere it does not belong. The test used is
  `(max - min) / 255 < 0.12`, or luminance above 0.945, or below 0.075.
- Chroma is binned by hue at 12 degrees, adjacent bins merge when their circular distance is under
  26 degrees.
- Family hue is a weighted circular mean (vector sum, then `atan2`).

Families sort by weight descending. On ties — several equally sized blocks is a common layout —
the family closest to the theme's `accent1` wins, so the result does not depend on traversal order.

## 3. Mapping

Source families are sorted by weight and mapped positionally onto the target families. If there
are more source families than targets, later families reuse the target list and are pushed along
the lightness axis by +/- 0.16 per cycle so two families do not collide.

Inside one family:

```
new_L = target_L + (member_L - family_weighted_mean_L) * depth
```

The family's weighted mean lightness is anchored to the target's lightness, and each member moves
by its own deviation scaled by `depth`. This preserves order — the brightest member stays the
brightest — and, unlike remapping into a fixed interval, it does not drift the family's overall
brightness or invert a dark outline into a light one.

Saturation follows as `target_S * (0.55 + 0.45 * member_S / family_mean_S)`, so muted members stay
muted relative to vivid ones.

Colors carrying `lumMod`, `tint` or `shade` modifiers get a new hue and saturation but keep their
modifiers. Re-lightening an already lightened color is how palettes end up entirely white.

## 4. Gradients

- All stops of a single gradient are force-merged into one family (`merge_gradient_clusters`).
  Skipping this splits warm gradients into two families and returns two unrelated hues.
- After mapping, if the end-to-end lightness gap is below `--gradient-spread`, the stops are spread
  to that gap while preserving their order. Source gradients are frequently almost flat, and a
  faithful mapping of a flat gradient produces something that no longer reads as a gradient.

## 5. Neutrals

Mapped linearly between the slide's own dark and light neutrals:

```
new_L = dark_L + t * (light_L - dark_L)
```

Hue is taken from the primary and saturation is capped at 0.02 to 0.12, so greys carry a hint of
the deck's color instead of reading as dead grey. Pure white (`L > 0.965`) and pure black
(`L < 0.045`) are preserved by default.

## 6. Write-back scope

Color nodes are rewritten; nothing else is. Covers solid fills, every gradient stop, `a:ln`
outlines, `a:highlight`, run properties (`a:rPr`, `a:defRPr`, `a:endParaRPr`), effect colors such
as glow and outer shadow, image `a:duotone`, group children at any depth, table cells, chart parts
and SmartArt color parts.

Color spellings are normalized to `a:srgbClr`, with `prstClr`, `sysClr`, `scrgbClr` and `hslClr`
converted to RGB first. Alpha and lightness modifier children are preserved.
