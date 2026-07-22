# Public image-only reconstruction demo

This directory contains a real second test of the image-only workflow. The repository owner supplied the reference image, the reconstructed PowerPoint file, the desktop PowerPoint render, and the workflow recording for public demonstration.

## Files

- `reference.png`: the only visual input used for reconstruction.
- `reconstructed.png`: the reconstructed slide exported by desktop PowerPoint.
- `editable-demo.pptx`: the editable one-slide PowerPoint result.
- `difference-heatmap.png`: a generated visualization of pixel differences.
- `workflow.gif`: a compressed 30-second recording of the real workflow.
- `demo-report.json`: machine-readable comparison and editability summary.

## Measured result

Image comparison was performed with the repository's `compare_slide_images.py` script using `contain` alignment on a 1672 × 940 canvas.

| Metric | Value |
| --- | ---: |
| Pixel similarity | `0.903267` |
| Composite visual score | `0.868828` |
| Global SSIM | `0.804273` |
| Edge similarity | `0.953807` |
| Mean absolute error | `0.096733` |

Read-only OOXML inspection with `inspect_pptx.py` found:

| Structure | Count |
| --- | ---: |
| Slides | `1` |
| Total slide objects | `358` |
| Shapes (`sp`) | `223` |
| Connectors (`cxnSp`) | `130` |
| Pictures (`pic`) | `5` |
| Text objects | `101` |
| Text characters | `693` |
| Potential overflow objects | `0` |

The PowerPoint slide size is 13.3333 × 7.5 inches (16:9). The detected fonts are Microsoft YaHei and Segoe UI Symbol.

## Reproduce the checks

From the repository root:

```powershell
python .\plugins\ppt-visual-reconstructor\scripts\compare_slide_images.py `
  --reference .\docs\demo\reference.png `
  --candidate .\docs\demo\reconstructed.png `
  --output-dir .\demo-comparison

python .\plugins\ppt-visual-reconstructor\scripts\inspect_pptx.py `
  .\docs\demo\editable-demo.pptx `
  --output .\demo-inspection.json
```

This demo is intentionally reported separately from the first golden sample in the project benchmark. It demonstrates that the public image-only path clears the 90% pixel-similarity target on a dense Chinese technical slide while retaining PowerPoint editability.
