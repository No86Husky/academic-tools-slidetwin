# Windows PowerPoint bridge

Use these scripts only on Windows with desktop PowerPoint installed. All scripts preserve the input deck and write results to a new or temporary directory.

## Recommended sequence

When the plugin MCP tools are available, call them in this order and do not ask the user to copy commands:

```text
ppt_environment_status
ppt_probe_powerpoint
ppt_prepare_image_scene
ppt_build_editable_slide
ppt_compare_slide
ppt_inspect_powerpoint
ppt_apply_correction_plan
ppt_compare_slide
```

All tool paths must be absolute. The output directory may be new; the tool creates it. The scripts below are the manual fallback when MCP tools are unavailable.

1. Verify COM automation and Chinese text:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\probe_powerpoint_v4.ps1
```

2. For a single-image input, let Codex write `scene-plan.json`, then create the semantic SVG, protected image assets, editable PPTX, and first render:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_powerpoint_from_scene_v1.ps1 `
  -ReferenceImagePath .\reference.png `
  -ScenePlanPath .\scene-plan.json `
  -OutputDirectory .\image-only-output
```

3. Inspect native objects and produce a PowerPoint render:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\inspect_powerpoint_v1.ps1 `
  -PresentationPath .\source.pptx
```

4. Compare the first native render with the reference:

```powershell
python .\scripts\compare_slide_images.py `
  --reference .\reference.png `
  --candidate .\image-only-output\reconstructed-from-image.png `
  --output-dir .\image-only-output\comparison
```

5. Probe installed fonts only when font substitution is plausible:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\probe_fonts_v1.ps1
```

6. Apply a reviewed correction plan:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\apply_powerpoint_plan_v4.ps1 `
  -PresentationPath .\source.pptx `
  -PlanPath .\plan.json `
  -OutputDirectory .\output
```

7. If glyph shape remains the largest difference after geometry is aligned, run a font sweep on the latest reconstruction:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\sweep_powerpoint_fonts_v1.ps1 `
  -PresentationPath .\output\reconstructed.pptx `
  -OutputDirectory .\font-sweep
```

Compare font variants by region. Do not choose a new font based on one title if it makes card or body text worse.

## PowerShell copying rule

Copy only the command text. Do not copy the visible `PS C:\...>` prompt, PowerShell banner, prior error output, or explanatory prose into the console.

Every named parameter requires a space before its value:

```powershell
-PresentationPath "C:\path\source.pptx"
```

not:

```powershell
-PresentationPathC:\path\source.pptx
```

## Output contract

The image-only builder writes:

```text
image-only-output/
├── reconstructed-from-image.pptx
├── reconstructed-from-image.png
├── semantic-preview.svg
├── scene-plan.resolved.json
├── scene-preparation-result.json
├── image-only-build-result.json
└── assets/
    └── protected-region.png
```

The correction-plan runner writes:

```text
output/
├── reconstructed.pptx
├── application-result.json
└── renders/
    ├── slide-1.png
    └── ...
```

Treat `application-result.json` as required evidence. Confirm `success`, `protected_objects_unchanged`, the source hash, and the operation log before accepting the presentation.

## Troubleshooting

- A blank PowerPoint window during automation does not necessarily indicate failure; check the PowerShell stage log.
- Close an existing output presentation before overwriting the same output directory.
- Windows PowerShell 5.1 may corrupt non-ASCII script source without a BOM. Bundled PowerShell scripts are ASCII-only and construct Unicode test text from code points.
- PowerPoint uses numeric `MsoTriState` values (`-1` and `0`) for several COM properties; do not replace them with PowerShell booleans.
