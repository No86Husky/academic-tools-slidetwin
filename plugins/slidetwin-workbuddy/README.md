# SlideTwin for WorkBuddy / CodeBuddy Code

This directory is a standalone WorkBuddy-compatible integration for SlideTwin.

The original Codex plugin remains unchanged. The WorkBuddy plugin has its own manifest, Skill, MCP configuration, launcher, installer, documentation, and validation workflow.

## Architecture

```text
WorkBuddy or CodeBuddy model
        ↓ visual understanding and scene planning
WorkBuddy SlideTwin Skill
        ↓ MCP calls
WorkBuddy runtime launcher
        ↓ delegates to the tested shared runtime
Existing SlideTwin PowerPoint scripts
        ↓
Desktop PowerPoint COM + editable PPTX + native render + comparison
```

The launcher uses `SLIDETWIN_RUNTIME_ROOT` to locate the existing `plugins/ppt-visual-reconstructor` runtime. It also adapts Codex-specific MCP descriptions before returning the tool list to WorkBuddy. It does not alter the shared runtime files.

## User command

```text
/slidetwin:reconstruct
```

Additional instructions can be passed after the command, for example:

```text
/slidetwin:reconstruct Preserve the artistic title and photograph as pictures. Rebuild everything else as editable objects.
```

## Installation

Use the repository-level `install-workbuddy.ps1` installer. See `docs/workbuddy/INSTALLATION.md` for requirements, preview installation, manual testing, and marketplace installation.

## Validation

```powershell
python .\tools\validate_workbuddy_plugin.py
codebuddy plugin validate .\plugins\slidetwin-workbuddy
```

See `docs/workbuddy/TESTING.md` for the complete acceptance matrix.

## Current status

Version `0.6.0-preview` provides:

- independent WorkBuddy marketplace and plugin manifests;
- a user-invocable reconstruction Skill;
- a cache-safe MCP launcher;
- reuse of the existing tested SlideTwin runtime;
- a WorkBuddy-specific one-command installer;
- static and PowerShell syntax validation in GitHub Actions;
- explicit regression protection for the original Codex integration.
