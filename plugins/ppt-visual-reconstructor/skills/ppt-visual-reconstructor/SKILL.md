---
name: ppt-visual-reconstructor
description: Reconstruct, repair, and visually refine editable PowerPoint slides from reference PNG/JPEG images, SVG conversions, or existing PPTX drafts. Use when Codex must preserve selected artwork as raster images, rebuild other elements as editable PowerPoint text and shapes, compare a PowerPoint-rendered slide against a reference image, fix SVG-to-shape text/layout errors, or report visual fidelity and editability.
---

# PPT Visual Reconstructor

Treat desktop PowerPoint's own export as the rendering authority. Do not declare visual fidelity from a third-party renderer alone.

## Inputs

Collect:

- A reference PNG or JPEG at the highest available resolution.
- An existing PPTX or SVG when repairing a prior conversion.
- Original photos, logos, or illustrations when available.
- A list of elements that must remain raster images.
- The target slide size and required fonts.

If original assets are unavailable, identify regions that cannot be recovered cleanly from the flattened reference and state the editability tradeoff.

## Object policy

- Rebuild titles, body copy, labels, numbers, and captions as native text boxes.
- Rebuild simple fills, rules, cards, circles, and arrows as native PowerPoint shapes.
- Keep protected artwork and photographs as image objects. Never trace them into PowerPoint geometry.
- Keep complex icons as single SVG objects unless native-shape editing materially benefits the user.
- Preserve charts as native charts only when source data is available; otherwise preserve them as images and report the limitation.
- Assign stable, semantic object names before iterative correction.

## Workflow

1. Inspect the reference image, SVG, and PPTX structure.
2. Record the slide aspect ratio, fonts, object counts, grouped objects, custom geometries, images, and overflow risks.
3. Produce an element manifest with object type, bounds, style, editability target, and asset source.
4. Build or repair the slide with native PowerPoint objects according to the object policy.
5. Export the slide with desktop PowerPoint at the reference resolution or a proportional resolution.
6. Align the exported image with the reference and compute layout, text, color, imagery, and editability checks.
7. Correct the largest localized differences first: global geometry, image crop, text wrapping, typography, then decorative details.
8. Repeat until the quality threshold is met or two consecutive iterations make no meaningful improvement.
9. Deliver the editable PPTX, final rendered preview, and a concise quality report.

Read [references/windows-bridge.md](references/windows-bridge.md) before asking a Windows user to run desktop PowerPoint automation. Read [references/plan-format.md](references/plan-format.md) before generating or reviewing a JSON correction plan.

## Hard requirements

- Preserve reference text content exactly; do not tolerate garbled characters.
- Do not flatten the whole slide to improve the visual score.
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

## Bundled tools

- Run `scripts/inspect_pptx.py` to create a read-only OOXML inventory without requiring PowerPoint.
- Run `scripts/probe_powerpoint_v4.ps1` on Windows to verify that desktop PowerPoint COM automation, Chinese text, PPTX saving, and PNG export work.
- Run `scripts/inspect_powerpoint_v1.ps1` on Windows to read native object properties and export every slide with desktop PowerPoint without modifying the source deck.
- Run `scripts/compare_slide_images.py` to align the reference to the PowerPoint render, compute baseline visual metrics, and create overlay and difference images.
- Run `scripts/apply_powerpoint_plan_v4.ps1` on Windows to copy a source deck, apply an explicit JSON modification plan (font scaling/replacement/spacing/color, targeted movement, visibility, and editable rounded-rectangle creation), verify protected objects, save the copy, and render the result.
- Run `scripts/probe_fonts_v1.ps1` on Windows when font-family substitution may explain width or glyph differences.
- Run `scripts/sweep_powerpoint_fonts_v1.ps1` on a visually aligned reconstruction to render installed Chinese font-family variants in one batch; compare regions separately before selecting fonts for the final plan.

If the desktop PowerPoint probe has not passed, limit work to structural inspection and clearly label visual validation as pending.
