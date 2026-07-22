# PPT Visual Reconstructor

PPT Visual Reconstructor is a Codex plugin and Windows companion toolkit for creating a high-fidelity editable PowerPoint slide from only a reference image.

Upload one PNG or JPEG. The plugin decomposes it into semantic text, simple shapes, lines, icons, and complex image regions; creates an SVG intermediate; builds a new editable PPTX with desktop PowerPoint; and iteratively compares the native render with the reference. Selected artwork and photographs remain separate raster images while text and simple graphics become native editable objects. Existing SVG and PPTX files are optional accelerators, not required inputs.

> 中文概述：使用者只需上传一张PPT参考图片，插件自动生成语义SVG中间稿和可编辑PPT，再使用桌面版PowerPoint的真实渲染结果进行迭代校正，目标像素还原度不低于90%。

## Status

Version `0.4.0` is an installable image-only public preview. It adds:

- A unique `ppt-visual-tools` Codex marketplace, avoiding collisions with a generic `personal` marketplace.
- A Windows one-command installer that can install Git, Node.js, Python, and Codex CLI when explicitly requested.
- A zero-dependency local MCP server that exposes the workflow as eight callable Codex tools.
- Direct status, PowerPoint probe, image-scene preparation, editable-slide build, native inspection, OOXML inspection, render comparison, and correction-plan tools.
- Protocol and manifest tests in GitHub Actions.

The v0.3 reconstruction layer already provides:

- Single-image input without requiring a source SVG or draft PPTX.
- Semantic scene decomposition into native text, native shapes, SVG objects, and protected raster pictures.
- Reference-region extraction for photographs, artistic lettering, and complex illustrations.
- Semantic SVG generation without pretending that one embedded full-slide bitmap is editable.
- Blank-presentation construction through native PowerPoint objects.
- Editable-coverage reporting alongside visual similarity.

The v0.2 correction and validation layer has already verified:

- 64-bit Windows and desktop PowerPoint COM automation.
- Chinese text without encoding corruption through the PowerShell bridge.
- Recursive inspection of grouped PowerPoint objects.
- Non-destructive JSON correction plans.
- Protected-picture geometry verification.
- Native PowerPoint PNG rendering and deterministic visual comparison.
- Font-family calibration across installed CJK fonts.
- Editable rounded-card and shadow reconstruction.

On the first golden sample, the workflow improved the strict composite score from `0.7175` to `0.8815` and reached `0.9708` pixel similarity while keeping four specified artwork objects as pictures. The strict score deliberately penalizes font rasterization and antialiasing differences; report both metrics instead of treating them as interchangeable.

## How it works

1. Inspect the uploaded reference image at full resolution and transcribe its text.
2. Classify each visible element as native text, native shape, SVG object, or protected raster picture.
3. Write a semantic scene plan and generate a semantic SVG preview plus cropped protected-image assets.
4. Create a new blank PowerPoint slide and rebuild the scene with native objects.
5. Export the slide through desktop PowerPoint and align it with the reference.
6. Measure visual differences and create an explicit correction plan.
7. Correct typography, geometry, colors, crop, shadows, and layering.
8. Verify protected objects, render again, and repeat until the quality gate is met or the remaining mismatch is documented.

The SVG is an intermediate inspection artifact. Text is rebuilt as native PowerPoint text rather than converted into glyph outlines, because outline conversion looks editable but prevents normal text editing.

Desktop PowerPoint is the rendering authority. LibreOffice or third-party previews are not used to approve final fidelity.

## Supported environment

- Windows 10 or Windows 11, 64-bit
- A 64-bit desktop PowerPoint installation with COM automation enabled
- Windows PowerShell 5.1 or PowerShell 7+
- Python 3.10+ for OOXML inspection and image comparison
- Python packages: `numpy` and `Pillow`
- Node.js 20+ and Codex CLI for direct plugin/MCP use

The PowerPoint COM scripts must run on a Windows machine with desktop PowerPoint installed. Structural PPTX inspection and image comparison are cross-platform.

## Local quick start

Clone or download this repository, then validate the PowerPoint bridge:

```powershell
powershell -ExecutionPolicy Bypass -File .\plugins\ppt-visual-reconstructor\scripts\probe_powerpoint_v4.ps1
```

For the image-only path, let Codex create a scene plan using the uploaded reference and then run:

```powershell
powershell -ExecutionPolicy Bypass -File .\plugins\ppt-visual-reconstructor\scripts\build_powerpoint_from_scene_v1.ps1 `
  -ReferenceImagePath .\reference.png `
  -ScenePlanPath .\scene-plan.json `
  -OutputDirectory .\image-only-output
```

The command creates the protected image assets, semantic SVG, editable PPTX, native PowerPoint render, and build report.

Inspect a presentation and export native slide renders:

```powershell
powershell -ExecutionPolicy Bypass -File .\plugins\ppt-visual-reconstructor\scripts\inspect_powerpoint_v1.ps1 `
  -PresentationPath .\example.pptx
```

Compare a PowerPoint render with a reference image:

```powershell
python .\plugins\ppt-visual-reconstructor\scripts\compare_slide_images.py `
  --reference .\reference.png `
  --candidate .\slide-1.png `
  --output-dir .\comparison
```

Apply a correction plan to a new presentation:

```powershell
powershell -ExecutionPolicy Bypass -File .\plugins\ppt-visual-reconstructor\scripts\apply_powerpoint_plan_v4.ps1 `
  -PresentationPath .\example.pptx `
  -PlanPath .\examples\preserve-pictures-and-scale-text.json `
  -OutputDirectory .\reconstructed-output
```

See [the image-only scene format](plugins/ppt-visual-reconstructor/skills/ppt-visual-reconstructor/references/scene-format.md), [the Windows bridge reference](plugins/ppt-visual-reconstructor/skills/ppt-visual-reconstructor/references/windows-bridge.md), and [the correction-plan format](plugins/ppt-visual-reconstructor/skills/ppt-visual-reconstructor/references/plan-format.md) for details.

## Codex plugin layout

```text
.agents/plugins/marketplace.json
plugins/ppt-visual-reconstructor/
├── .codex-plugin/plugin.json
├── .mcp.json
├── scripts/
└── skills/ppt-visual-reconstructor/
    ├── SKILL.md
    ├── agents/openai.yaml
    └── references/
```

## Install from GitHub

### One command on Windows 11

Open PowerShell and copy only the command below. It downloads the public installer, installs missing prerequisites with `winget`, registers the marketplace, installs the plugin, and runs the PowerPoint bridge probe:

```powershell
$p = Join-Path $env:TEMP "install-ppt-visual-reconstructor.ps1"; Invoke-WebRequest "https://raw.githubusercontent.com/No86Husky/academic-tools/main/install.ps1" -OutFile $p; powershell.exe -NoProfile -ExecutionPolicy Bypass -File $p -InstallPrerequisites -RunPowerPointProbe
```

The prerequisite switch authorizes installation of Git, Node.js LTS, Python 3.12, Codex CLI, `numpy`, and `Pillow`. Omit `-InstallPrerequisites` when those commands already exist. You may download and inspect `install.ps1` before running it.

### Manual install

If Git and Codex CLI are already available:

```powershell
git clone https://github.com/No86Husky/academic-tools.git "$HOME\academic-tools"
codex plugin marketplace add "$HOME\academic-tools"
codex plugin add ppt-visual-reconstructor@ppt-visual-tools
```

Start a new Codex thread after installation so the skill metadata is reloaded. Upload one slide image and use:

```text
Use $ppt-visual-reconstructor to create an editable PowerPoint slide from only this image. Automatically keep photographs, artistic lettering, and complex illustrations as separate pictures; rebuild all text and simple geometry as native PowerPoint objects; and iterate toward at least 90% pixel similarity.
```

Codex should call the installed tools directly. The expected tool sequence is:

```text
ppt_environment_status -> ppt_probe_powerpoint -> ppt_prepare_image_scene
-> ppt_build_editable_slide -> ppt_compare_slide -> correction loop
```

You should not normally need to run the underlying PowerShell scripts by hand.

## Safety and editability

- Source presentations are copied before modification.
- Correction plans are explicit JSON, reviewable before execution.
- Protected picture addresses cannot be targeted by supported modification operations.
- Protected pictures are checked for unchanged type, identity, geometry, rotation, and crop settings.
- The workflow never flattens the whole slide merely to improve a score.
- Visual similarity and editable coverage are reported separately.

## Current limitations

- Image-only scene plans are generated by Codex's visual reasoning; the local scripts do not yet bundle a separate deterministic OCR and segmentation engine.
- The first image-only builder handles one slide per case. Batch orchestration is planned after the single-slide quality gate is stable.
- Complex freeform vector paths may remain one SVG object or raster picture instead of becoming hundreds of fragile PowerPoint primitives.
- PowerPoint shape addresses are index-based and can change after users manually add, remove, or regroup objects.
- Exact font rasterization can differ even when font family, size, bounds, and spacing match.
- A low-resolution reference, unreadable text, or unavailable font can prevent the 90% target; the plugin must report the plateau instead of flattening the slide.
- Protected-picture verification does not yet hash the embedded media bytes.
- PowerPoint for macOS and web PowerPoint are not supported by the COM bridge.

## Development

Run repository and plugin validation:

```bash
python tools/validate_repository.py
python /root/.codex/skills/.system/plugin-creator/scripts/validate_plugin.py plugins/ppt-visual-reconstructor
python /root/.codex/skills/oai/skill-creator/scripts/quick_validate.py plugins/ppt-visual-reconstructor/skills/ppt-visual-reconstructor
```

The last two commands are available in Codex development environments. The first command uses only the Python standard library and is suitable for GitHub Actions.

See [CONTRIBUTING.md](CONTRIBUTING.md) for contribution rules.

## License

MIT
