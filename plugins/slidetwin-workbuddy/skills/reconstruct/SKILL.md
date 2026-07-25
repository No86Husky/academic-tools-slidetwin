---
name: reconstruct
description: Recreate a reference PNG or JPEG slide as a high-fidelity editable PowerPoint presentation. Use when the user supplies a slide screenshot or reference image and wants editable text, shapes, connectors, and separately preserved complex imagery.
---

# SlideTwin Reconstruction

Recreate one reference slide image as an editable visual twin in PowerPoint.

## Environment requirements

- Windows 10 or Windows 11, 64-bit
- Desktop Microsoft PowerPoint with COM automation available
- Node.js
- Python 3.10 or newer
- Pillow and NumPy

## Object policy

Classify every visible element into one of these categories:

1. Native PowerPoint text
2. Native PowerPoint shape or connector
3. Editable SVG object
4. Protected raster picture

Use native PowerPoint text for normal wording. Use native shapes for cards, rectangles, circles, lines, arrows, and other simple geometry. Preserve photographs, artistic lettering, detailed illustrations, and visually complex regions as separate cropped pictures.

Never flatten the entire slide into one full-slide image.

## Required workflow

1. Resolve the reference PNG or JPEG to an absolute local path.
2. Create a dedicated output directory under the current workspace.
3. Call `ppt_environment_status`.
4. If desktop PowerPoint has not been verified on this machine, call `ppt_probe_powerpoint` once.
5. Inspect the reference image at full resolution and transcribe all readable text.
6. Write a complete UTF-8 semantic scene-plan JSON that follows the bundled SlideTwin scene format.
7. Call `ppt_prepare_image_scene` to validate the plan, crop protected image regions, and create the SVG preview.
8. Correct the scene plan when validation reports a specific error.
9. Call `ppt_build_editable_slide`.
10. Call `ppt_compare_slide` using the native PowerPoint PNG render.
11. Review the metrics, overlay, and difference heatmap.
12. When meaningful correctable differences remain, inspect the PPTX, write an explicit JSON correction plan, and call `ppt_apply_correction_plan`.
13. Compare again. Run no more than three correction passes unless the user explicitly requests additional passes.
14. Stop when the quality target is reached or improvement plateaus.
15. Return the final editable PPTX path, native render path, comparison report, editable coverage, and documented limitations.

## Quality policy

- Target at least 90% pixel similarity when reference quality and available fonts permit.
- Treat the native desktop PowerPoint render as the rendering authority.
- Report visual similarity and editable coverage separately.
- Prefer editability over artificially improving the score by flattening content.
- Do not claim that a threshold was reached unless the comparison report confirms it.
- Document unreadable text, unavailable fonts, low-resolution source regions, or a measured quality plateau.

## Safety

- Never modify an original presentation in place.
- Write every build and correction pass to a new output directory.
- Keep protected pictures unchanged unless the user explicitly changes their protected status.
- Do not execute unrelated commands.
