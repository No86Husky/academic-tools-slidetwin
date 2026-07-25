#!/usr/bin/env python3
"""Build the uploadable WorkBuddy desktop SlideTwin Skill ZIP."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import zipfile

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "workbuddy-desktop" / "slidetwin"
CANONICAL_SKILL_ROOT = (
    ROOT / "plugins" / "ppt-visual-reconstructor" / "skills" / "ppt-visual-reconstructor"
)
DEFAULT_OUTPUT = ROOT / "dist" / "slidetwin-workbuddy-skill.zip"
FIXED_TIMESTAMP = (2026, 7, 25, 0, 0, 0)

CANONICAL_FILES = {
    Path("references/shared-workflow.md"): CANONICAL_SKILL_ROOT / "SKILL.md",
    Path("references/scene-format.md"): CANONICAL_SKILL_ROOT / "references" / "scene-format.md",
    Path("references/plan-format.md"): CANONICAL_SKILL_ROOT / "references" / "plan-format.md",
    Path("references/windows-bridge.md"): CANONICAL_SKILL_ROOT / "references" / "windows-bridge.md",
}
DEPRECATED_SOURCE_FILES = {Path("references/scene-plan-format.md")}


def iter_source_files():
    for path in sorted(SOURCE.rglob("*")):
        if not path.is_file():
            continue
        relative = path.relative_to(SOURCE)
        if relative in DEPRECATED_SOURCE_FILES:
            continue
        yield path, relative


def sha256_bytes(content: bytes) -> str:
    return hashlib.sha256(content).hexdigest()


def write_entry(archive: zipfile.ZipFile, archive_name: Path, content: bytes) -> None:
    info = zipfile.ZipInfo(str(archive_name).replace(os.sep, "/"), FIXED_TIMESTAMP)
    info.compress_type = zipfile.ZIP_DEFLATED
    info.external_attr = 0o644 << 16
    archive.writestr(info, content)


def build(output: Path) -> Path:
    required = [SOURCE / "SKILL.md", SOURCE / "manifest.yaml"] + list(CANONICAL_FILES.values())
    missing = [str(path.relative_to(ROOT)) for path in required if not path.is_file()]
    if missing:
        raise FileNotFoundError(f"Missing required Skill files: {', '.join(missing)}")

    output = output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()

    provenance: dict[str, dict[str, str]] = {}

    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for path, relative in iter_source_files():
            content = path.read_bytes()
            archive_name = Path("slidetwin") / relative
            write_entry(archive, archive_name, content)
            provenance[str(archive_name).replace(os.sep, "/")] = {
                "source": str(path.relative_to(ROOT)).replace(os.sep, "/"),
                "sha256": sha256_bytes(content),
            }

        for relative, source in CANONICAL_FILES.items():
            content = source.read_bytes()
            archive_name = Path("slidetwin") / relative
            write_entry(archive, archive_name, content)
            provenance[str(archive_name).replace(os.sep, "/")] = {
                "source": str(source.relative_to(ROOT)).replace(os.sep, "/"),
                "sha256": sha256_bytes(content),
            }

        provenance_name = Path("slidetwin") / "references" / "canonical-provenance.json"
        provenance_content = (
            json.dumps(
                {
                    "schema_version": "1.0",
                    "purpose": "Prove that the WorkBuddy Skill bundles the canonical Codex SlideTwin workflow and formats.",
                    "files": provenance,
                },
                ensure_ascii=False,
                indent=2,
                sort_keys=True,
            )
            + "\n"
        ).encode("utf-8")
        write_entry(archive, provenance_name, provenance_content)

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
        if not required_names.issubset(names):
            missing_names = sorted(required_names - names)
            raise RuntimeError(f"Generated ZIP is missing required WorkBuddy Skill files: {missing_names}")
        if "slidetwin/references/scene-plan-format.md" in names:
            raise RuntimeError("Generated ZIP still contains the deprecated divergent scene-plan-format.md")
        bad = archive.testzip()
        if bad:
            raise RuntimeError(f"Generated ZIP contains a corrupt entry: {bad}")

    return output


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()
    result = build(args.output)
    print(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
