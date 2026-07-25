#!/usr/bin/env python3
"""Validate the standalone WorkBuddy integrations without third-party packages."""

from __future__ import annotations

import hashlib
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
CANONICAL_SKILL_ROOT = (
    ROOT / "plugins" / "ppt-visual-reconstructor" / "skills" / "ppt-visual-reconstructor"
)
CANONICAL_ZIP_FILES = {
    "slidetwin/references/shared-workflow.md": CANONICAL_SKILL_ROOT / "SKILL.md",
    "slidetwin/references/scene-format.md": CANONICAL_SKILL_ROOT / "references" / "scene-format.md",
    "slidetwin/references/plan-format.md": CANONICAL_SKILL_ROOT / "references" / "plan-format.md",
    "slidetwin/references/windows-bridge.md": CANONICAL_SKILL_ROOT / "references" / "windows-bridge.md",
}


def fail(message: str) -> None:
    raise AssertionError(message)


def sha256_bytes(content: bytes) -> str:
    return hashlib.sha256(content).hexdigest()


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
    if not str(manifest.get("version", "")).startswith("0.7.0"):
        fail("WorkBuddy parity preview manifest must use the 0.7.0 version line")
    description = str(manifest.get("description", "")).lower()
    for required in ("quality", "powerpoint", "comparison"):
        if required not in description:
            fail(f"WorkBuddy parity manifest description is missing: {required}")


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
    plugin_skill = require_file("plugins/slidetwin-workbuddy/skills/reconstruct/SKILL.md")
    launcher = require_file("plugins/slidetwin-workbuddy/scripts/launch-mcp.mjs")
    installer = require_file("install-workbuddy.ps1")
    require_file("install-workbuddy-offline.ps1")
    require_file("docs/workbuddy/INSTALLATION.md")
    require_file("docs/workbuddy/TESTING.md")
    require_file("tools/build_workbuddy_skill.py")
    require_file("tools/build_workbuddy_offline_bundle.py")

    plugin_skill_text = plugin_skill.read_text(encoding="utf-8")
    for required in (
        "Never flatten the entire slide",
        "ppt_environment_status",
        "ppt_compare_slide",
        "ppt_inspect_powerpoint",
        "PASS requirements",
        "PLATEAU and FAIL",
        "$ARGUMENTS",
        "Do not use Grep",
    ):
        if required not in plugin_skill_text:
            fail(f"WorkBuddy plugin Skill is missing parity content: {required}")

    launcher_text = launcher.read_text(encoding="utf-8")
    for required in (
        "SLIDETWIN_RUNTIME_ROOT",
        "Never use Grep",
        "open and visually inspect",
        "Do not rely on metric JSON alone",
        "Stop only at PASS",
    ):
        if required not in launcher_text:
            fail(f"WorkBuddy launcher is missing parity guidance: {required}")

    installer_text = installer.read_text(encoding="utf-8")
    if "@tencent-ai/codebuddy-code" in installer_text:
        fail("WorkBuddy desktop installer must not install CodeBuddy Code")
    if '".workbuddy\\mcp.json"' not in installer_text:
        fail("WorkBuddy desktop installer must configure the user-level MCP file")
    if "build_workbuddy_skill.py" not in installer_text:
        fail("WorkBuddy desktop installer must build the uploadable Skill package")


def validate_desktop_skill_source() -> None:
    skill_md = DESKTOP_SKILL / "SKILL.md"
    manifest = DESKTOP_SKILL / "manifest.yaml"
    quality_gate = DESKTOP_SKILL / "references" / "quality-gate.md"
    deprecated_pointer = DESKTOP_SKILL / "references" / "scene-plan-format.md"
    for path in (skill_md, manifest, quality_gate, deprecated_pointer):
        if not path.is_file():
            fail(f"Missing WorkBuddy desktop Skill source: {path.relative_to(ROOT)}")

    skill_text = skill_md.read_text(encoding="utf-8")
    for required in (
        "name: slidetwin",
        "references/shared-workflow.md",
        "references/scene-format.md",
        "references/quality-gate.md",
        "ppt_environment_status",
        "ppt_build_editable_slide",
        "ppt_compare_slide",
        "ppt_inspect_powerpoint",
        "不得使用 Grep",
        "PASS",
        "PLATEAU",
        "FAIL",
        "未经比较报告确认，不得声称达到 90%",
    ):
        if required not in skill_text:
            fail(f"WorkBuddy desktop Skill is missing parity content: {required}")

    quality_text = quality_gate.read_text(encoding="utf-8")
    for required in (
        "metrics.pixel_similarity >= 0.90",
        "metrics.composite_visual_score >= 0.90",
        "Text-content accuracy is `1.00`",
        "Protected-picture integrity is `1.00`",
        "two consecutive correction passes",
        "less than `0.003`",
        "PASS — quality gate satisfied",
    ):
        if required not in quality_text:
            fail(f"WorkBuddy quality gate is missing: {required}")

    pointer_text = deprecated_pointer.read_text(encoding="utf-8")
    if "Do not use this file" not in pointer_text or "source.type" not in pointer_text:
        fail("Deprecated WorkBuddy scene format must point clearly to the canonical schema")

    manifest_text = manifest.read_text(encoding="utf-8")
    for required in ("name: slidetwin", "version: 0.7.0", "description:", "category:", "author:"):
        if required not in manifest_text:
            fail(f"WorkBuddy desktop manifest is missing: {required}")


def validate_desktop_skill_package() -> None:
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
                "slidetwin/references/shared-workflow.md",
                "slidetwin/references/scene-format.md",
                "slidetwin/references/plan-format.md",
                "slidetwin/references/windows-bridge.md",
                "slidetwin/references/quality-gate.md",
                "slidetwin/references/canonical-provenance.json",
            }
            missing = required_names - names
            if missing:
                fail(f"Desktop Skill ZIP is missing entries: {sorted(missing)}")
            if "slidetwin/references/scene-plan-format.md" in names:
                fail("Desktop Skill ZIP contains the deprecated divergent scene format")
            if archive.testzip() is not None:
                fail("Desktop Skill ZIP integrity check failed")

            provenance = json.loads(
                archive.read("slidetwin/references/canonical-provenance.json").decode("utf-8")
            )
            provenance_files = provenance.get("files")
            if not isinstance(provenance_files, dict):
                fail("Canonical provenance file is invalid")

            for archive_name, source_path in CANONICAL_ZIP_FILES.items():
                expected = source_path.read_bytes()
                actual = archive.read(archive_name)
                if actual != expected:
                    fail(f"WorkBuddy package drifted from canonical Codex file: {archive_name}")
                record = provenance_files.get(archive_name)
                if not isinstance(record, dict):
                    fail(f"Missing provenance record: {archive_name}")
                if record.get("sha256") != sha256_bytes(expected):
                    fail(f"Incorrect canonical hash in provenance: {archive_name}")


def validate_codex_is_separate_but_shared() -> None:
    require_file("plugins/ppt-visual-reconstructor/.codex-plugin/plugin.json")
    require_file("plugins/ppt-visual-reconstructor/.mcp.json")
    require_file("plugins/ppt-visual-reconstructor/scripts/mcp-server.mjs")
    if PLUGIN.resolve() == (ROOT / "plugins" / "ppt-visual-reconstructor").resolve():
        fail("WorkBuddy and Codex plugin roots must remain separate")
    launcher_text = (PLUGIN / "scripts" / "launch-mcp.mjs").read_text(encoding="utf-8")
    if 'path.resolve(pluginRoot, "..", "ppt-visual-reconstructor")' not in launcher_text:
        fail("WorkBuddy launcher must delegate to the shared Codex-tested runtime")


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
        validate_desktop_skill_source,
        validate_desktop_skill_package,
        validate_codex_is_separate_but_shared,
        validate_javascript,
    ]
    for check in checks:
        check()
        print(f"PASS: {check.__name__}")
    print("WorkBuddy Codex-parity integration validation passed.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except AssertionError as exc:
        print(f"FAIL: {exc}", file=sys.stderr)
        raise SystemExit(1)
