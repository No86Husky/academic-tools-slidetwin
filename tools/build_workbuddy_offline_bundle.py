#!/usr/bin/env python3
"""Build a self-contained offline SlideTwin runtime bundle for WorkBuddy."""

from __future__ import annotations

import argparse
import os
from pathlib import Path
import zipfile

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUTPUT = ROOT / "dist" / "slidetwin-workbuddy-offline.zip"
FIXED_TIMESTAMP = (2026, 7, 25, 0, 0, 0)

FILES = [
    Path("install-workbuddy-offline.ps1"),
    Path("requirements.txt"),
    Path("tools/build_workbuddy_skill.py"),
]
DIRECTORIES = [
    Path("plugins/ppt-visual-reconstructor"),
    Path("plugins/slidetwin-workbuddy"),
    Path("workbuddy-desktop/slidetwin"),
]


def iter_payload_files():
    for relative in FILES:
        path = ROOT / relative
        if not path.is_file():
            raise FileNotFoundError(f"Missing required file: {relative}")
        yield path, relative

    for directory in DIRECTORIES:
        root = ROOT / directory
        if not root.is_dir():
            raise FileNotFoundError(f"Missing required directory: {directory}")
        for path in sorted(root.rglob("*")):
            if path.is_file():
                yield path, path.relative_to(ROOT)


def write_entry(archive: zipfile.ZipFile, source: Path, archive_name: Path) -> None:
    info = zipfile.ZipInfo(str(archive_name).replace(os.sep, "/"), FIXED_TIMESTAMP)
    info.compress_type = zipfile.ZIP_DEFLATED
    info.external_attr = 0o644 << 16
    archive.writestr(info, source.read_bytes())


def build(output: Path) -> Path:
    output = output.resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.exists():
        output.unlink()

    offline_installer = ROOT / "install-workbuddy-offline.ps1"
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        write_entry(archive, offline_installer, Path("install-workbuddy-offline.ps1"))
        for source, relative in iter_payload_files():
            if relative == Path("install-workbuddy-offline.ps1"):
                continue
            write_entry(archive, source, Path("payload") / relative)

    with zipfile.ZipFile(output, "r") as archive:
        names = set(archive.namelist())
        required = {
            "install-workbuddy-offline.ps1",
            "payload/requirements.txt",
            "payload/plugins/ppt-visual-reconstructor/scripts/mcp-server.mjs",
            "payload/plugins/slidetwin-workbuddy/scripts/launch-mcp.mjs",
        }
        missing = required - names
        if missing:
            raise RuntimeError(f"Offline bundle is missing: {sorted(missing)}")
        corrupt = archive.testzip()
        if corrupt:
            raise RuntimeError(f"Offline bundle contains a corrupt entry: {corrupt}")

    return output


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()
    print(build(args.output))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
