---
name: reconstruct
description: Recreate a reference PNG or JPEG slide as a high-fidelity editable PowerPoint presentation using SlideTwin tools.
---

# SlideTwin Reconstruction

Use SlideTwin to transform a reference slide image into an editable PowerPoint slide.

## Object policy

- Rebuild normal text as native PowerPoint text.
- Rebuild simple geometry as native PowerPoint shapes.
- Preserve photographs, artistic lettering and complex illustrations as separate images.
- Never flatten the complete slide into one image.

## Workflow

1. Analyze the uploaded reference image.
2. Create a semantic scene plan JSON.
3. Validate the scene plan.
4. Prepare image assets and SVG intermediates.
5. Build the editable PowerPoint slide.
6. Render with desktop PowerPoint.
7. Compare the render with the reference.
8. Apply correction plans when needed.

## Output

Return:

- editable PPTX path
- rendered comparison image
- similarity metrics
- remaining limitations

Do not modify original files in place.
