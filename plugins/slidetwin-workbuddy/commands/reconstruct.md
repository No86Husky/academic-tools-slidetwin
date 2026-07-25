---
description: Recreate a reference slide image as an editable PowerPoint
argument-hint: "[reference image path or additional instructions]"
---

Use the bundled SlideTwin reconstruction skill and the `slidetwin-tools` MCP server.

Reference input or additional instructions:

$ARGUMENTS

Follow the complete image-only reconstruction workflow:

1. Resolve the reference PNG or JPEG to an absolute local path.
2. Check the SlideTwin runtime with `ppt_environment_status`.
3. Run `ppt_probe_powerpoint` if this Windows PowerPoint installation has not been verified.
4. Analyze the image at full resolution and write a UTF-8 semantic scene-plan JSON.
5. Preserve photographs, artistic lettering, and complex illustrations as separate raster pictures.
6. Rebuild ordinary text, cards, lines, arrows, and simple geometry as editable PowerPoint objects.
7. Run `ppt_prepare_image_scene`, `ppt_build_editable_slide`, and `ppt_compare_slide`.
8. When meaningful correctable differences remain, write an explicit correction plan and run `ppt_apply_correction_plan`.
9. Compare again, stopping when the target is reached or improvement plateaus.
10. Return the editable PPTX path, native render, comparison report, editable coverage, and any limitations.

Never flatten the entire slide into one full-slide image merely to improve similarity.
