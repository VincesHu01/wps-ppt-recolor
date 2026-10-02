# Sample deck

Both files here are synthetic. No real data, no client names, no real figures.

- `sample_before.pptx` — a generic quarterly review, palette `#1B3A5C` and `#2E8B8B`, with a
  vendor diagram pasted onto slide 2 and a vendor chart onto slide 3, both in clashing colors.
- `sample_after.pptx` — the same deck after the engine ran on slides 2 and 3.

## Rebuild

```bash
python3 make_sample.py            # python-pptx is only needed to build the sample
cp sample_before.pptx sample_after.pptx
python3 ../scripts/recolor.py sample_after.pptx --slide 2 --target last --out sample_after.pptx
python3 ../scripts/recolor.py sample_after.pptx --slide 3 --target last --out sample_after.pptx
```

The engine itself has no third-party dependencies.

## Why the diagram is a group

`make_sample.py` wraps the vendor shapes into a real `p:grpSp` rather than leaving them as separate
shapes. That mirrors how a pasted template normally arrives, and it is what makes the default
`--target last` select the whole thing at once. Left ungrouped, only the last shape would be
recolored and the rest would pollute the palette sampling.
