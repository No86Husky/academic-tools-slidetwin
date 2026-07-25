#!/usr/bin/env python3
"""Validate the standalone WorkBuddy integrations without third-party packages."""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
import zipfile

ROOT = Path(__file__).resolve().parents[1]
PLUGIN = ROOT / "plugins" / "slidetwin-workbuddy"
MARKETPLACE = ROOT / ".codebuddy-plugin" / "marketplace.json"
DESKTOP_SKILL = ROOT / "workbuddy-desktop" / "slidetwin"


def fail(message: str) -> None:
    raise AssertionError(message)


def read_json(path: Path) -> dict:
    if not path.is_file():
        fail(f"Missing JSON file: {path.relative_to(ROOT)}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        fail(f"Invalid JSON in {path.relative_to(ROOT)}: {exc}")


def require_file(relative: str) -> Path:
    path = ROOT / relative
    if not path.is_file():
        fail(f"Missing required file: {relative}")
    return path


def validate_marketplace() -> None:
    data = read_json(MARKETPLACE)
    if data.get("name") != "slidetwin-tools":
        fail("WorkBuddy marketplace name must be slidetwin-tools")
    plugins = data.get("plugins")
    if not isinstance(plugins, list) or len(plugins) != 1:
        fail("WorkBuddy marketplace must expose exactly one standalone plugin")
    entry = plugins[0]
    if entry.get("name") != "slidetwin":
        fail("Marketplace plugin name must be slidetwin")
    if entry.get("source") != "./plugins/slidetwin-workbuddy":
        fail("Marketplace source must point only to the WorkBuddy plugin directory")


def validate_manifest() -> None:
    manifest = read_json(PLUGIN / ".codebuddy-plugin" / "plugin.json")
    if manifest.get("name") != "slidetwin":
        fail("WorkBuddy plugin manifest name must be slidetwin")
    if not str(manifest.get("version", "")).startswith("0.6.0"):
        fail("WorkBuddy preview manifest must use the 0.6.0 version line")
    if "codex" in json.dumps(manifest, ensure_ascii=False).lower():
        fail("WorkBuddy manifest must not advertise itself as a Codex plugin")


def validate_mcp() -> None:
    config = read_json(PLUGIN / ".mcp.json")
    servers = config.get("mcpServers")
    if not isinstance(servers, dict) or "slidetwin-tools" not in servers:
        fail("WorkBuddy MCP configuration must define slidetwin-tools")
    server = servers["slidetwin-tools"]
    if server.get("command") != "node":
        fail("WorkBuddy MCP server must launch with node")
    args = server.get("args")
    expected = "${CODEBUDDY_PLUGIN_ROOT}/scripts/launch-mcp.mjs"
    if not isinstance(args, list) or args != [expected]:
        fail(f"WorkBuddy MCP args must be exactly [{expected!r}]")


def validate_text_assets() -> None:
    skill = require_file("plugins/slidetwin-workbuddy/skills/reconstruct/SKILL.md")
    launcher = require_file("plugins/slidetwin-workbuddy/scripts/launch-mcp.mjs")
    installer = require_file("install-workbuddy.ps1")
    require_file("docs/workbuddy/INSTALLATION.md")
    require_file("docs/workbuddy/TESTING.md")
    require_file("tools/build_workbuddy_skill.py")

    skill_text = skill.read_text(encoding="utf-8")
    if "Never flatten the entire slide" not in skill_text:
        fail("WorkBuddy plugin Skill must preserve the anti-flattening policy")
    if "ppt_environment_status" not in skill_text or "ppt_compare_slide" not in skill_text:
        fail("WorkBuddy plugin Skill must state the required MCP workflow")
    if "$ARGUMENTS" not in skill_text:
        fail("WorkBuddy plugin Skill must forward user arguments")
    if "SLIDETWIN_RUNTIME_ROOT" not in launcher.read_text(encoding="utf-8"):
        fail("WorkBuddy launcher must support the shared runtime environment variable")

    installer_text = installer.read_text(encoding="utf-8")
    if "@tencent-ai/codebuddy-code" in installer_text:
        fail("WorkBuddy desktop installer must not install CodeBuddy Code")
    if '".workbuddy\\mcp.json"' not in installer_text:
        fail("WorkBuddy desktop installer must configure the user-level MCP file")
    if "build_workbuddy_skill.py" not in installer_text:
        fail("WorkBuddy desktop installer must build the uploadable Skill package")


def validate_desktop_skill_package() -> None:
    skill_md = DESKTOP_SKILL / "SKILL.md"
    manifest = DESKTOP_SKILL / "manifest.yaml"
    reference = DESKTOP_SKILL / "references" / "scene-plan-format.md"
    for path in (skill_md, manifest, reference):
        if not path.is_file():
            fail(f"Missing WorkBuddy desktop Skill source: {path.relative_to(ROOT)}")

    skill_text = skill_md.read_text(encoding="utf-8")
    for required in (
        "name: slidetwin",
        "ppt_environment_status",
        "ppt_build_editable_slide",
        "ppt_compare_slide",
        "不得为了提高相似度而把整页参考图作为一张全页背景图片",
    ):
        if required not in skill_text:
            fail(f"WorkBuddy desktop Skill is missing required content: {required}")

    manifest_text = manifest.read_text(encoding="utf-8")
    for required in ("name: slidetwin", "version:", "description:", "category:", "author:"):
        if required not in manifest_text:
            fail(f"WorkBuddy desktop manifest is missing: {required}")

    with tempfile.TemporaryDirectory() as temp_dir:
        output = Path(temp_dir) / "slidetwin-workbuddy-skill.zip"
        result = subprocess.run(
            [sys.executable, str(ROOT / "tools" / "build_workbuddy_skill.py"), "--output", str(output)],
            cwd=ROOT,
            text=True,
            capture_output=True,
            check=False,
        )
        if result.returncode != 0:
            fail(f"Desktop Skill build failed:\n{result.stdout}\n{result.stderr}")
        if not output.is_file():
            fail("Desktop Skill build did not create a ZIP")
        with zipfile.ZipFile(output, "r") as archive:
            names = set(archive.namelist())
            required_names = {
                "slidetwin/SKILL.md",
                "slidetwin/manifest.yaml",
                "slidetwin/references/scene-plan-format.md",
            }
            if not required_names.issubset(names):
                fail(f"Desktop Skill ZIP is missing entries: {sorted(required_names - names)}")
            if archive.testzip() is not None:
                fail("Desktop Skill ZIP integrity check failed")


def validate_codex_is_separate() -> None:
    require_file("plugins/ppt-visual-reconstructor/.codex-plugin/plugin.json")
    require_file("plugins/ppt-visual-reconstructor/.mcp.json")
    require_file("plugins/ppt-visual-reconstructor/scripts/mcp-server.mjs")
    if PLUGIN.resolve() == (ROOT / "plugins" / "ppt-visual-reconstructor").resolve():
        fail("WorkBuddy and Codex plugin roots must be different")


def validate_javascript() -> None:
    node = shutil.which("node")
    if not node:
        print("SKIP: node is not available; JavaScript syntax check was not run.")
        return
    launcher = PLUGIN / "scripts" / "launch-mcp.mjs"
    result = subprocess.run(
        [node, "--check", str(launcher)],
        cwd=ROOT,
        text=True,
        capture_output=True,
        check=False,
    )
    if result.returncode != 0:
        fail(f"Node syntax check failed:\n{result.stdout}\n{result.stderr}")


def main() -> int:
    checks = [
        validate_marketplace,
        validate_manifest,
        validate_mcp,
        validate_text_assets,
        validate_desktop_skill_package,
        validate_codex_is_separate,
        validate_javascript,
    ]
    for check in checks:
        check()
        print(f"PASS: {check.__name__}")
    print("WorkBuddy integration validation passed.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except AssertionError as exc:
        print(f"FAIL: {exc}", file=sys.stderr)
        raise SystemExit(1)
