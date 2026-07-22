---
name: ppt-visual-reconstructor
description: Compatibility alias for SlideTwin. Create an editable PowerPoint slide from only a reference PNG/JPEG image, or repair an existing SVG/PPTX reconstruction. Use when Codex must decompose a flattened slide image into native text and shapes, preserve photographs or artwork as separate raster pictures, emit a semantic SVG intermediate, render with desktop PowerPoint, iteratively reach at least 90% pixel similarity, and report editability separately from visual fidelity.
---

# SlideTwin (legacy `$ppt-visual-reconstructor` alias)

This Skill preserves the original invocation used by existing installations and workflow videos. For new prompts, prefer `$slidetwin`; both names follow the same reconstruction workflow.

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

## Use the installed tools first

When the plugin-provided MCP tools are callable, use them instead of asking the user to copy PowerShell commands:

1. Call `ppt_environment_status` once in a new environment.
2. On Windows, call `ppt_probe_powerpoint` before the first native build unless a successful probe from the same installation is already available.
3. Create all working files in one absolute output directory.
4. Call `ppt_prepare_image_scene`, inspect its semantic SVG and report, then call `ppt_build_editable_slide`.
5. Call `ppt_compare_slide` after every native render. Call `ppt_inspect_powerpoint` before writing a correction plan and `ppt_apply_correction_plan` after the plan is reviewed.

Use the bundled scripts directly only when the corresponding MCP tool is unavailable. Native PPTX creation and repair require Windows desktop PowerPoint; do not simulate approval with a third-party renderer on another operating system.

## Image-only workflow

1. Inspect the reference at full resolution. Transcribe all visible text and detect the slide canvas, layout regions, object boundaries, colors, typography, photographs, artwork, icons, and decoration.
2. Classify every visible element as `native_text`, `native_shape`, `svg_object`, or `raster_picture` according to the object policy.
3. Read [references/scene-format.md](references/scene-format.md) and write an image-only scene plan in the reference-image pixel coordinate system. One semantic paragraph or label must remain one text box.
4. Call `ppt_prepare_image_scene` to validate the plan, crop protected picture regions from the reference, and create `semantic-preview.svg`. Never use a visible full-slide raster element to pass the metric.
5. Inspect the semantic SVG against the reference. Correct missing elements, wrong colors, incorrect bounds, or text transcription before creating the PPTX.
6. On Windows, call `ppt_build_editable_slide` to create a new presentation from a blank slide. The builder creates native PowerPoint text, shapes, and lines directly and inserts only the classified complex regions as pictures.
7. Call `ppt_compare_slide` on the new native render to generate the first visual metrics, overlay, and difference heatmap.
8. Call `ppt_inspect_powerpoint`, generate a correction plan, and call `ppt_apply_correction_plan` for typography, color, visibility, movement, and editable shape refinement.
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

- `ppt_environment_status`: verify the MCP runtime without launching PowerPoint.
- `ppt_probe_powerpoint`: verify desktop PowerPoint COM automation, Chinese text, PPTX saving, and PNG export.
- `ppt_prepare_image_scene`: validate a scene plan, extract protected PNG assets, and produce the semantic SVG.
- `ppt_build_editable_slide`: create the editable PPTX and its first native PowerPoint render.
- `ppt_inspect_powerpoint`: inspect native objects and export slides through PowerPoint without modifying the source deck.
- `ppt_inspect_pptx_structure`: create a cross-platform read-only OOXML inventory.
- `ppt_compare_slide`: compute visual metrics and create the overlay and heatmap.
- `ppt_apply_correction_plan`: copy a deck, apply an explicit JSON plan, verify protected pictures, save the copy, and render it.

Every MCP tool has a same-purpose script in `scripts/` for manual fallback and development.

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
