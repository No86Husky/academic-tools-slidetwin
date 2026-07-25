# Install SlideTwin for WorkBuddy desktop

This integration is separate from the existing Codex plugin. It reuses the tested local SlideTwin PowerPoint runtime without modifying, reinstalling, or removing the Codex integration.

## Requirements

- Windows 10 or Windows 11, 64-bit
- WorkBuddy desktop
- Desktop Microsoft PowerPoint, 64-bit
- PowerShell 5.1 or PowerShell 7+
- Git
- Node.js 18.20 or newer
- Python 3.10 or newer

CodeBuddy Code is **not required** for the WorkBuddy desktop installation.

## Preview setup

Download or clone the `feat/workbuddy-plugin` branch, open PowerShell in the repository directory, and run:

```powershell
powershell -ExecutionPolicy Bypass -File .\install-workbuddy.ps1 `
  -InstallPrerequisites `
  -RunPowerPointProbe
```

The installer:

1. Installs missing Git, Node.js, and Python only when `-InstallPrerequisites` is supplied.
2. Clones or updates SlideTwin under `~/.slidetwin/academic-tools-slidetwin`.
3. Installs the Python image-processing dependencies.
4. records `SLIDETWIN_RUNTIME_ROOT` for the local PowerPoint runtime.
5. Safely backs up and merges the `slidetwin-tools` server into `~/.workbuddy/mcp.json`.
6. Builds `dist/slidetwin-workbuddy-skill.zip`.
7. Optionally verifies desktop PowerPoint COM automation.
8. Opens File Explorer with the uploadable Skill ZIP selected.

It does not install CodeBuddy Code and does not change the original Codex plugin.

## Upload the Skill in WorkBuddy

1. Open WorkBuddy.
2. Open **技能**.
3. Click **添加技能**.
4. Choose **上传技能**.
5. Upload:

```text
~\.slidetwin\academic-tools-slidetwin\dist\slidetwin-workbuddy-skill.zip
```

WorkBuddy supports importing a local Skill package and configures the uploaded Skill automatically.

## Confirm the MCP connection

The installer writes the user-level MCP configuration to:

```text
~\.workbuddy\mcp.json
```

In WorkBuddy, open **连接器** or the MCP configuration page and confirm that:

```text
slidetwin-tools
```

shows a green/connected state.

If WorkBuddy was open while the installer ran, close and reopen WorkBuddy before checking the MCP state.

## First use

Create a new task, upload one PNG or JPEG reference slide, enable the SlideTwin Skill, and ask:

```text
请使用 SlideTwin 将这张图片重建为可编辑 PowerPoint。
复杂插图和艺术字保留为独立图片，普通文字、色块、框、线条和箭头转换为可编辑对象。
```

The selected WorkBuddy model performs image understanding and writes the semantic scene plan. SlideTwin performs local PowerPoint construction, native rendering, comparison, and correction. It does not use the repository owner's Codex allowance or API credentials.

## Manual MCP configuration

When automatic MCP configuration is skipped, paste a configuration shaped like this into WorkBuddy's MCP editor, replacing the example paths with the actual local paths:

```json
{
  "mcpServers": {
    "slidetwin-tools": {
      "type": "stdio",
      "command": "C:\\Program Files\\nodejs\\node.exe",
      "args": [
        "C:\\Users\\YOUR_NAME\\.slidetwin\\academic-tools-slidetwin\\plugins\\slidetwin-workbuddy\\scripts\\launch-mcp.mjs"
      ],
      "env": {
        "SLIDETWIN_RUNTIME_ROOT": "C:\\Users\\YOUR_NAME\\.slidetwin\\academic-tools-slidetwin\\plugins\\ppt-visual-reconstructor"
      },
      "description": "SlideTwin local editable-PowerPoint reconstruction tools"
    }
  }
}
```

## Preview limitation

This branch has passed repository, Node, JSON, MCP-forwarding, and PowerShell parser checks. A real WorkBuddy + desktop PowerPoint reconstruction still needs to be completed on a user's Windows machine before the preview is merged into `main`.
