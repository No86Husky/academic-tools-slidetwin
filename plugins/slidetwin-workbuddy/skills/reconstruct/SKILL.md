---
name: reconstruct
description: Recreate a reference PNG or JPEG slide as a high-fidelity editable PowerPoint presentation using the same SlideTwin runtime, object policy, native PowerPoint renderer, comparison metrics, and correction loop as the Codex integration.
user-invocable: true
---

# SlideTwin Reconstruction — WorkBuddy parity mode

Recreate one reference slide image as an editable visual twin in PowerPoint.

User-provided reference path or additional instructions:

$ARGUMENTS

## Non-negotiable policy

- Use native PowerPoint text for readable titles, body copy, labels, numbers, and captions.
- Use native PowerPoint shapes and connectors for simple geometry.
- Use one SVG object only for a genuinely vector-like icon that is not usefully represented by native primitives.
- Preserve photographs, artistic lettering, detailed illustrations, textures, and visually complex regions as separate cropped pictures.
- Never flatten the entire slide into one visible image or hide a full-slide raster inside an SVG.
- Treat desktop PowerPoint's native PNG export as the rendering authority.

## Required workflow

1. Resolve the reference image to an absolute local path and create a dedicated output directory.
2. Call `ppt_environment_status`.
3. Call `ppt_probe_powerpoint` before the first native build unless the same installation already has a successful probe.
4. Inspect the reference at full resolution and transcribe every readable character exactly.
5. Write a complete UTF-8 scene plan using the canonical SlideTwin schema used by the shared `ppt-visual-reconstructor` runtime. Protected crops must use `source.type = "reference_crop"` and `source.box_px`.
6. Call `ppt_prepare_image_scene` and fix every explicit validation error.
7. Do not use Grep, text search, or SVG XML inspection as a substitute for visual review. `semantic-preview.svg` is an interchange artifact, not the final fidelity authority.
8. Call `ppt_build_editable_slide`, then visually inspect the native PowerPoint PNG render.
9. Call `ppt_compare_slide` after every build or correction pass. Actually open and review the candidate render, aligned reference, overlay, and difference heatmap; do not rely on the JSON metrics alone.
10. Call `ppt_inspect_powerpoint` before writing a correction plan. Check text, fonts, object bounds, clipping, overflow, wrapping, overlap, grouping, and z-order.
11. Correct the largest localized differences first: global geometry, picture crop, text transcription and wrapping, typography, then decoration.
12. Call `ppt_apply_correction_plan`, render again, and compare again.
13. Stop only when the quality gate passes or two consecutive correction passes produce no meaningful improvement.

## PASS requirements

A result is PASS only when all conditions hold:

- text-content accuracy is 100%;
- protected-picture integrity is 100%;
- no visible garbling;
- no unintended clipping or overflow;
- `pixel_similarity >= 0.90`;
- `composite_visual_score >= 0.90` when the same fonts and PowerPoint renderer are available;
- visual similarity and editable coverage are reported separately;
- at least 95% of detected text and simple geometry are native PowerPoint objects unless every exception is documented.

Never claim that 90% was reached unless `comparison.json` confirms it.

## PLATEAU and FAIL

- Return PLATEAU when the gate is unmet and two consecutive passes improve both pixel similarity and composite score by less than 0.003, or a documented font/asset/rendering limitation prevents further editable improvement.
- Return FAIL when the runtime, PowerPoint probe, build, render, comparison, anti-flattening policy, protected imagery, or final PPTX integrity fails.
- Only PASS may be described as meeting the SlideTwin acceptance threshold.

## Deliverables

Return the final editable PPTX, native PowerPoint render, resolved scene plan, semantic SVG, protected assets, comparison JSON, overlay, heatmap, PowerPoint inspection report, correction plans, editability report, limitations, and exactly one status: PASS, PLATEAU, or FAIL.
