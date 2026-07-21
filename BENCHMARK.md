# Golden-sample benchmark

This benchmark records the first end-to-end reconstruction without publishing the private reference slide, source deck, or extracted artwork.

The recorded golden sample validates the native PowerPoint rendering, comparison, protected-picture, and iterative correction layers introduced before v0.3. The new v0.3 image-only bootstrap is validated structurally and must complete its Windows golden-sample run before it is described as fully automatic.

## Environment

- 64-bit Windows desktop environment
- Desktop PowerPoint with COM automation
- PowerShell bridge with UTF-8 plan transport
- One 16:9 slide rendered by desktop PowerPoint at 1600 × 900
- Four specified artwork elements required to remain picture objects

## Results

| Stage | Pixel similarity | Global SSIM | Edge similarity | Composite score |
|---|---:|---:|---:|---:|
| Initial editable conversion | 0.9449 | 0.4267 | 0.9737 | 0.7175 |
| Global font-scale correction | 0.9560 | 0.4950 | 0.9777 | 0.7529 |
| Targeted text geometry and color correction | 0.9666 | 0.6790 | 0.9824 | 0.8404 |
| Editable rounded panels and shadows | 0.9708 | 0.7670 | 0.9829 | 0.8815 |

The final pixel similarity exceeded the requested 90% visual-fidelity threshold. The stricter composite score remained below 0.90 because it also penalizes glyph outlines and native font antialiasing differences.

## Integrity checks

- Source presentation hash recorded before every modification pass.
- All edits written to copied presentations.
- Four protected picture objects retained their identity, type, bounds, rotation, and crop geometry.
- No protected address was accepted by a modifying operation.
- Chinese text rendered without garbling.
- Font sweep tested multiple candidate CJK families; the original system sans-serif rendering remained the best whole-slide result.

## Interpretation

Pixel similarity, structural editability, and protected-object integrity are separate acceptance dimensions. A deck should not be flattened merely to raise SSIM. When exact font rasterization is unavailable, report the metric plateau and keep editable text.
