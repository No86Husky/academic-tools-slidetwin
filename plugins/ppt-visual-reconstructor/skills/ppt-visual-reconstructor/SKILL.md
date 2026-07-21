---
name: ppt-visual-reconstructor
description: Create an editable PowerPoint slide from only a reference PNG/JPEG image, or repair an existing SVG/PPTX reconstruction. Use when Codex must decompose a flattened slide image into native text and shapes, preserve photographs or artwork as separate raster pictures, emit a semantic SVG intermediate, render with desktop PowerPoint, iteratively reach at least 90% pixel similarity, and report editability separately from visual fidelity.
---

# PPT Visual Reconstructor

Treat desktop PowerPoint's own export as the rendering authority. Do not declare visual fidelity from a third-party renderer alone.

## Primary promise

A high-resolution reference PNG or JPEG is the only required user input. Do not require the user to first create an SVG or draft PPTX.

When the user supplies only an image, Codex must perform the image-only workflow below. An existing SVG or PPTX is an optional accelerator and activates the repair workflow after the first inventory step.

## Inputs

Collect:

- Required: one reference PNG or JPEG at the highest available resolution.
- Optional: an existing PPTX or SVG when repairing a prior conversion.
- Optional: original photos, logos, illustrations, fonts, slide size, or a list of protected image regions.

Infer the aspect ratio, slide size, object policy, and likely fonts when the optional inputs are absent. Ask the user only when text is unreadable or a region's required editability is genuinely ambiguous. By default, automatically keep photographs, artistic lettering, and complex textured illustrations as separate picture objects.

If original assets are unavailable, identify regions that cannot be recovered cleanly from the flattened reference and state the editability tradeoff.

## Object policy

- Rebuild titles, body copy, labels, numbers, and captions as native text boxes.
- Rebuild simple fills, rules, cards, circles, and arrows as native PowerPoint shapes.
- Keep protected artwork and photographs as image objects. Never trace them into PowerPoint geometry.
- Keep complex icons as single SVG objects unless native-shape editing materially benefits the user.
- Preserve charts as native charts only when source data is available; otherwise preserve them as images and report the limitation.
- Assign stable, semantic object names before iterative correction.

## Image-only workflow

1. Inspect the reference at full resolution. Transcribe all visible text and detect the slide canvas, layout regions, object boundaries, colors, typography, photographs, artwork, icons, and decoration.
2. Classify every visible element as `native_text`, `native_shape`, `svg_object`, or `raster_picture` according to the object policy.
3. Read [references/scene-format.md](references/scene-format.md) and write an image-only scene plan in the reference-image pixel coordinate system. One semantic paragraph or label must remain one text box.
4. Run `scripts/prepare_image_scene.py` to validate the plan, crop protected picture regions from the reference, and create `semantic-preview.svg`. Never use a visible full-slide raster element to pass the metric.
5. Inspect the semantic SVG against the reference. Correct missing elements, wrong colors, incorrect bounds, or text transcription before creating the PPTX.
6. On Windows, run `scripts/build_powerpoint_from_scene_v1.ps1` to create a new presentation from a blank slide. The builder creates native PowerPoint text, shapes, and lines directly and inserts only the classified complex regions as pictures.
7. Export the new slide with desktop PowerPoint, run `scripts/compare_slide_images.py`, and generate the first visual metrics, overlays, and difference heatmap.
8. Inspect the created PPTX, generate a correction plan, and use `scripts/apply_powerpoint_plan_v4.ps1` for typography, color, visibility, movement, and editable shape refinement.
9. Repeat native rendering, comparison, and targeted correction until the quality gate is met or two consecutive passes produce no meaningful improvement.
10. Deliver the editable PPTX, semantic SVG, protected image assets, final render, scene plan, correction plans, and quality/editability report.

The semantic SVG is an inspection and interchange artifact. Do not mechanically convert every SVG glyph and path into PowerPoint geometry: that turns text into outlines and makes the result less editable. The final builder must use the scene semantics to create native PowerPoint objects wherever possible.

## Existing SVG/PPTX repair workflow

1. Inspect the reference image, SVG, and PPTX structure.
2. Record the slide aspect ratio, fonts, object counts, grouped objects, custom geometries, images, and overflow risks.
3. Produce an element manifest with object type, bounds, style, editability target, and asset source.
4. Build or repair the slide with native PowerPoint objects according to the object policy.
5. Export the slide with desktop PowerPoint at the reference resolution or a proportional resolution.
6. Align the exported image with the reference and compute layout, text, color, imagery, and editability checks.
7. Correct the largest localized differences first: global geometry, image crop, text wrapping, typography, then decorative details.
8. Repeat until the quality threshold is met or two consecutive iterations make no meaningful improvement.
9. Deliver the editable PPTX, final rendered preview, and a concise quality report.

Read [references/windows-bridge.md](references/windows-bridge.md) before asking a Windows user to run desktop PowerPoint automation. Read [references/scene-format.md](references/scene-format.md) before creating a slide from only an image. Read [references/plan-format.md](references/plan-format.md) before generating or reviewing a JSON correction plan.

## Hard requirements

- Preserve reference text content exactly; do not tolerate garbled characters.
- Accept a single PNG/JPEG as a complete starting input in image-only mode.
- Do not flatten the whole slide to improve the visual score.
- Do not claim that a raster image embedded inside an SVG is an editable reconstruction.
- Do not convert protected images into shapes.
- Do not stretch the reference to a different aspect ratio without explicit approval.
- Flag missing fonts and font substitution before layout correction.
- Flag any object outside the slide canvas.
- Inspect unexpected overlap, clipping, line wrapping, transparency, grouping, and z-order.
- Keep the source files unchanged and save repaired output as a new file.

## Initial quality gate

Require all of the following for an accepted result:

- Text content accuracy: 100%.
- Protected-image preservation: 100%.
- No unintended overflow or clipping.
- No visible garbling.
- Pixel similarity: at least 0.90 when a supported comparison backend is available.
- Composite visual score: target at least 0.90 when the same fonts and rendering engine are available. If native font rasterization plateaus below this target, report the score, region-level findings, and visual-review result instead of flattening editable content.
- Editable coverage reported separately from visual similarity.
- At least 95% of detected text and simple geometric elements represented as native PowerPoint objects unless the report documents a justified exception.

## Bundled tools

- Run `scripts/prepare_image_scene.py` to validate an image-only scene plan, extract complex regions as separate PNG assets, and produce a semantic SVG preview.
- Run `scripts/build_powerpoint_from_scene_v1.ps1` on Windows to create a new editable PPTX from the resolved scene plan and export its first native render.
- Run `scripts/inspect_pptx.py` to create a read-only OOXML inventory without requiring PowerPoint.
- Run `scripts/probe_powerpoint_v4.ps1` on Windows to verify that desktop PowerPoint COM automation, Chinese text, PPTX saving, and PNG export work.
- Run `scripts/inspect_powerpoint_v1.ps1` on Windows to read native object properties and export every slide with desktop PowerPoint without modifying the source deck.
- Run `scripts/compare_slide_images.py` to align the reference to the PowerPoint render, compute baseline visual metrics, and create overlay and difference images.
- Run `scripts/apply_powerpoint_plan_v4.ps1` on Windows to copy a source deck, apply an explicit JSON modification plan (font scaling/replacement/spacing/color, targeted movement, visibility, and editable rounded-rectangle creation), verify protected objects, save the copy, and render the result.
- Run `scripts/probe_fonts_v1.ps1` on Windows when font-family substitution may explain width or glyph differences.
- Run `scripts/sweep_powerpoint_fonts_v1.ps1` on a visually aligned reconstruction to render installed Chinese font-family variants in one batch; compare regions separately before selecting fonts for the final plan.

If the desktop PowerPoint probe has not passed, limit work to structural inspection and clearly label visual validation as pending.

## Image-only output contract

Return:

- `reconstructed-from-image.pptx`
- `semantic-preview.svg`
- `scene-plan.resolved.json`
- `assets/` containing only regions intentionally preserved as pictures
- `reconstructed-from-image.png`
- final comparison JSON, overlay, and heatmap
- a concise report containing visual metrics, text accuracy, native-object counts, protected-picture integrity, and editable coverage
