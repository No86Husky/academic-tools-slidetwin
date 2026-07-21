# PPT Visual Reconstructor

PPT Visual Reconstructor is a Codex plugin and Windows companion toolkit for repairing editable PowerPoint slides against reference images.

It is designed for the workflow where a visually strong slide already exists as a PNG/JPEG or SVG, but the editable PowerPoint conversion has incorrect fonts, spacing, wrapping, shapes, shadows, or layering. Selected artwork and photographs can remain raster images while text and simple graphics stay native and editable.

> 中文概述：把参考图片或 SVG 转换结果修复成高还原度、可编辑的 PowerPoint，并使用桌面版 PowerPoint 的真实渲染结果进行迭代校正。

## Status

Version `0.2.0` is a tested public preview. The golden-sample workflow has verified:

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

1. Inspect the reference image, source SVG, and draft PPTX.
2. Mark photographs or artwork that must remain picture objects.
3. Inventory native PowerPoint objects and render the deck through desktop PowerPoint.
4. Align the render with the reference and measure visual differences.
5. Apply an explicit correction plan to a copy of the deck.
6. Verify protected objects, render again, and repeat until the quality gate is met or the remaining mismatch is documented.

Desktop PowerPoint is the rendering authority. LibreOffice or third-party previews are not used to approve final fidelity.

## Supported environment

- Windows 10 or Windows 11, 64-bit
- A 64-bit desktop PowerPoint installation with COM automation enabled
- Windows PowerShell 5.1 or PowerShell 7+
- Python 3.10+ for OOXML inspection and image comparison
- Python packages: `numpy` and `Pillow`

The PowerPoint COM scripts must run on a Windows machine with desktop PowerPoint installed. Structural PPTX inspection and image comparison are cross-platform.

## Local quick start

Clone or download this repository, then validate the PowerPoint bridge:

```powershell
powershell -ExecutionPolicy Bypass -File .\plugins\ppt-visual-reconstructor\scripts\probe_powerpoint_v4.ps1
```

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

See [the Windows bridge reference](plugins/ppt-visual-reconstructor/skills/ppt-visual-reconstructor/references/windows-bridge.md) and [the correction-plan format](plugins/ppt-visual-reconstructor/skills/ppt-visual-reconstructor/references/plan-format.md) for details.

## Codex plugin layout

```text
.agents/plugins/marketplace.json
plugins/ppt-visual-reconstructor/
├── .codex-plugin/plugin.json
├── scripts/
└── skills/ppt-visual-reconstructor/
    ├── SKILL.md
    ├── agents/openai.yaml
    └── references/
```

## Install from GitHub

Clone the public repository:

```bash
git clone https://github.com/No86Husky/academic-tools.git
cd academic-tools
```

Add the cloned repository root as a local Codex marketplace, then install the plugin:

```bash
codex plugin marketplace add <absolute-path-to-academic-tools>
codex plugin add ppt-visual-reconstructor@personal
```

Start a new Codex thread after installation so the skill metadata is reloaded. Invoke it explicitly with `$ppt-visual-reconstructor`, or ask Codex to reconstruct or repair an editable PowerPoint slide from a reference image.

## Safety and editability

- Source presentations are copied before modification.
- Correction plans are explicit JSON, reviewable before execution.
- Protected picture addresses cannot be targeted by supported modification operations.
- Protected pictures are checked for unchanged type, identity, geometry, rotation, and crop settings.
- The workflow never flattens the whole slide merely to improve a score.
- Visual similarity and editable coverage are reported separately.

## Current limitations

- Visual correction plans are generated iteratively; a single fully automatic reconstruction command is planned for a later release.
- PowerPoint shape addresses are index-based and can change after users manually add, remove, or regroup objects.
- Exact font rasterization can differ even when font family, size, bounds, and spacing match.
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
