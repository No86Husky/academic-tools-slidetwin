# Windows PowerPoint bridge

Use these scripts only on Windows with desktop PowerPoint installed. All scripts preserve the input deck and write results to a new or temporary directory.

## Recommended sequence

1. Verify COM automation and Chinese text:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\probe_powerpoint_v4.ps1
```

2. Inspect native objects and produce a PowerPoint render:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\inspect_powerpoint_v1.ps1 `
  -PresentationPath .\source.pptx
```

3. Probe installed fonts only when font substitution is plausible:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\probe_fonts_v1.ps1
```

4. Apply a reviewed correction plan:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\apply_powerpoint_plan_v4.ps1 `
  -PresentationPath .\source.pptx `
  -PlanPath .\plan.json `
  -OutputDirectory .\output
```

5. If glyph shape remains the largest difference after geometry is aligned, run a font sweep on the latest reconstruction:

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

The plan runner writes:

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
