#!/usr/bin/env python3
"""Build the uploadable WorkBuddy desktop SlideTwin Skill ZIP."""

from __future__ import annotations

import argparse
import os
from pathlib import Path
import zipfile

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "workbuddy-desktop" / "slidetwin"
DEFAULT_OUTPUT = ROOT / "dist" / "slidetwin-workbuddy-skill.zip"
FIXED_TIMESTAMP = (2026, 7, 25, 0, 0, 0)


def iter_files(source: Path):
    for path in sorted(source.rglob("*")):
        if path.is_file():
            yield path


def build(output: Path) -> Path:
    required = [SOURCE / "SKILL.md", SOURCE / "manifest.yaml"]
    missing = [str(path.relative_to(ROOT)) for path in required if not path.is_file()]
    if missing:
        raise FileNotFoundError(f"Missing required Skill files: {', '.join(missing)}")

    output = output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()

    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for path in iter_files(SOURCE):
            relative = path.relative_to(SOURCE)
            archive_name = Path("slidetwin") / relative
            info = zipfile.ZipInfo(str(archive_name).replace(os.sep, "/"), FIXED_TIMESTAMP)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            archive.writestr(info, path.read_bytes())

    with zipfile.ZipFile(output, "r") as archive:
        names = set(archive.namelist())
        required_names = {"slidetwin/SKILL.md", "slidetwin/manifest.yaml"}
        if not required_names.issubset(names):
            raise RuntimeError("Generated ZIP is missing required WorkBuddy Skill files")
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
