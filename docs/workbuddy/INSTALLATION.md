# Install SlideTwin for WorkBuddy / CodeBuddy Code

This integration is separate from the existing Codex plugin. It reuses the tested SlideTwin PowerPoint runtime without modifying or reinstalling the Codex integration.

## Requirements

- Windows 10 or Windows 11, 64-bit
- Desktop Microsoft PowerPoint, 64-bit
- PowerShell 5.1 or PowerShell 7+
- Git
- Node.js 18.20 or newer
- Python 3.10 or newer
- CodeBuddy Code

## Preview one-command installation

Open PowerShell and run the installer from a local checkout of the `feat/workbuddy-plugin` branch:

```powershell
powershell -ExecutionPolicy Bypass -File .\install-workbuddy.ps1 `
  -InstallPrerequisites `
  -RunPowerPointProbe
```

The installer:

1. Installs missing Git, Node.js, Python, and CodeBuddy Code only when `-InstallPrerequisites` is supplied.
2. Clones or updates the SlideTwin repository under `~/.slidetwin/academic-tools-slidetwin`.
3. Installs Python image-processing dependencies.
4. Records `SLIDETWIN_RUNTIME_ROOT` so the cached WorkBuddy plugin can call the existing SlideTwin runtime.
5. Registers the local marketplace as `slidetwin-tools`.
6. Installs `slidetwin@slidetwin-tools` at user scope.
7. Optionally verifies desktop PowerPoint COM automation.

The installer does not register, remove, update, or otherwise modify the Codex plugin.

## Manual development installation

From the repository root:

```powershell
$env:SLIDETWIN_RUNTIME_ROOT = (Resolve-Path .\plugins\ppt-visual-reconstructor).Path
codebuddy plugin validate .\plugins\slidetwin-workbuddy
codebuddy --plugin-dir .\plugins\slidetwin-workbuddy
```

Inside CodeBuddy Code, run:

```text
/reload-plugins
/slidetwin:reconstruct
```

## Manual marketplace installation

From CodeBuddy Code or the CLI:

```text
/plugin marketplace add No86Husky/academic-tools-slidetwin
/plugin install slidetwin@slidetwin-tools
/reload-plugins
```

During preview development, the GitHub default branch may not yet contain the WorkBuddy integration. Use the one-command installer or add a local checkout of `feat/workbuddy-plugin` until the feature branch is merged.

## First use

Upload one PNG or JPEG reference slide and run:

```text
/slidetwin:reconstruct
```

You can add instructions such as:

```text
/slidetwin:reconstruct Preserve the artistic title and the photograph as separate images. Rebuild all other text and simple geometry as editable PowerPoint objects.
```

## Model usage

The selected WorkBuddy/CodeBuddy model performs visual understanding and writes the semantic scene plan. SlideTwin performs local PowerPoint construction, native rendering, image comparison, and correction application. The plugin does not use the repository owner's Codex allowance or API credentials.
