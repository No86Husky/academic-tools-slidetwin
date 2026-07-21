#!/usr/bin/env python3
"""Validate repository structure without non-standard Python dependencies."""

from __future__ import annotations

import ast
import json
import re
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
PLUGIN = ROOT / "plugins" / "ppt-visual-reconstructor"
MANIFEST = PLUGIN / ".codex-plugin" / "plugin.json"
SKILL = PLUGIN / "skills" / "ppt-visual-reconstructor" / "SKILL.md"
MARKETPLACE = ROOT / ".agents" / "plugins" / "marketplace.json"
SEMVER = re.compile(r"^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-[0-9A-Za-z.-]+)?(?:\+[0-9A-Za-z.-]+)?$")


def fail(message: str) -> None:
    raise AssertionError(message)


def load_json(path: Path) -> object:
    with path.open(encoding="utf-8-sig") as handle:
        return json.load(handle)


def main() -> int:
    required = [MANIFEST, SKILL, MARKETPLACE, ROOT / "LICENSE", ROOT / "README.md"]
    for path in required:
        if not path.is_file():
            fail(f"Missing required file: {path.relative_to(ROOT)}")

    manifest = load_json(MANIFEST)
    if not isinstance(manifest, dict):
        fail("plugin.json must contain an object")
    if manifest.get("name") != PLUGIN.name:
        fail("Plugin folder and manifest name must match")
    if not SEMVER.match(str(manifest.get("version", ""))):
        fail("Plugin version must be semantic versioning")
    if manifest.get("license") != "MIT":
        fail("Public preview must declare the MIT license")
    for key in ("description", "author", "skills", "interface"):
        if key not in manifest:
            fail(f"plugin.json is missing {key}")

    marketplace = load_json(MARKETPLACE)
    entries = marketplace.get("plugins", []) if isinstance(marketplace, dict) else []
    matching = [entry for entry in entries if entry.get("name") == PLUGIN.name]
    if len(matching) != 1:
        fail("Marketplace must contain exactly one plugin entry")
    if matching[0].get("source", {}).get("path") != "./plugins/ppt-visual-reconstructor":
        fail("Marketplace source path is incorrect")

    skill_text = SKILL.read_text(encoding="utf-8")
    if not skill_text.startswith("---\n"):
        fail("SKILL.md must start with YAML frontmatter")
    frontmatter = skill_text.split("---", 2)[1]
    if "name: ppt-visual-reconstructor" not in frontmatter:
        fail("SKILL.md name does not match the plugin skill")
    if "description:" not in frontmatter:
        fail("SKILL.md must include a description")

    for path in PLUGIN.rglob("*.py"):
        ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    for path in PLUGIN.rglob("*.ps1"):
        try:
            path.read_text(encoding="ascii")
        except UnicodeDecodeError as error:
            fail(f"PowerShell script is not ASCII-only: {path.relative_to(ROOT)} ({error})")

    text_files = [
        *ROOT.glob("*.md"),
        *PLUGIN.rglob("*.md"),
        MANIFEST,
        MARKETPLACE,
    ]
    for path in text_files:
        text = path.read_text(encoding="utf-8-sig")
        if "[TODO:" in text:
            fail(f"Unresolved placeholder in {path.relative_to(ROOT)}")

    print("Repository validation passed")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except AssertionError as error:
        print(f"Validation failed: {error}", file=sys.stderr)
        raise SystemExit(1)
