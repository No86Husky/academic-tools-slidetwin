# WorkBuddy integration test plan

Run these tests on 64-bit Windows with desktop PowerPoint installed.

## A. Static plugin validation

```powershell
codebuddy plugin validate .\plugins\slidetwin-workbuddy
python .\tools\validate_workbuddy_plugin.py
```

Expected: both commands exit successfully.

## B. Local plugin loading

```powershell
$env:SLIDETWIN_RUNTIME_ROOT = (Resolve-Path .\plugins\ppt-visual-reconstructor).Path
codebuddy --plugin-dir .\plugins\slidetwin-workbuddy --debug
```

Inside CodeBuddy Code:

```text
/reload-plugins
```

Expected:

- One SlideTwin plugin is loaded.
- One plugin MCP server is reported.
- `/slidetwin:reconstruct` is available.
- The existing Codex plugin files are not loaded or changed by this test.

## C. MCP discovery

Ask the agent to call `ppt_environment_status`.

Expected:

- `success` is true.
- The server reports all required Python and PowerShell scripts.
- `plugin_root` points to the shared `plugins/ppt-visual-reconstructor` runtime, not the WorkBuddy cache directory.

## D. PowerPoint probe

Call `ppt_probe_powerpoint` with a new absolute output directory.

Expected:

- A PowerPoint file is created.
- Chinese text and a shape are created without encoding corruption.
- A native PNG render is exported.
- `powerpoint-probe.json` reports success.

## E. Image-only reconstruction

Upload a known reference PNG and run:

```text
/slidetwin:reconstruct
```

Expected:

1. The model writes a UTF-8 scene-plan JSON.
2. `ppt_prepare_image_scene` succeeds.
3. `ppt_build_editable_slide` creates `reconstructed-from-image.pptx`.
4. Desktop PowerPoint exports a PNG render.
5. `ppt_compare_slide` writes metrics, overlay, and difference heatmap.
6. The slide contains native text and shapes rather than one full-slide image.
7. Complex imagery remains in separate protected picture objects.

## F. Correction pass

Use a reference for which the first pass has visible, correctable differences.

Expected:

- The agent creates an explicit correction-plan JSON.
- `ppt_apply_correction_plan` writes a new presentation rather than modifying the source.
- A second comparison report is produced.
- Protected pictures remain unchanged.

## G. Marketplace installation

```powershell
codebuddy plugin marketplace add . --name slidetwin-tools
codebuddy plugin install slidetwin@slidetwin-tools --scope user
codebuddy plugin list --json
```

Expected:

- The marketplace is recognized.
- The plugin is installed at user scope.
- No dependency errors are reported.
- The cached plugin can still find the shared runtime through `SLIDETWIN_RUNTIME_ROOT`.

## H. Regression protection

Run the existing Codex repository validation and, where available, the existing Codex plugin tests.

Expected:

- Existing Codex marketplace metadata is unchanged.
- Existing `.codex-plugin/plugin.json`, Codex Skills, and Codex `.mcp.json` are unchanged.
- `$slidetwin` remains functional in Codex.

## Acceptance matrix

| ID | Test | Pass condition |
| --- | --- | --- |
| W01 | WorkBuddy manifest | `codebuddy plugin validate` passes |
| W02 | Marketplace | `slidetwin@slidetwin-tools` is discoverable |
| W03 | MCP startup | one MCP server loads without dependency error |
| W04 | Tool discovery | all eight existing SlideTwin tools are visible |
| W05 | PowerPoint probe | PPTX and PNG are generated |
| W06 | Command discovery | `/slidetwin:reconstruct` is available |
| W07 | Scene preparation | scene plan validates and assets are produced |
| W08 | Editable build | PPTX opens and is not flattened |
| W09 | Native render | PowerPoint PNG export succeeds |
| W10 | Comparison | metrics, overlay, and heatmap are produced |
| W11 | Correction | one safe correction pass succeeds |
| W12 | Codex regression | original Codex integration remains unchanged |
