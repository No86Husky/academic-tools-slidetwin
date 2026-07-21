#!/usr/bin/env python3
"""Create a read-only structural inventory of a PPTX file using stdlib only."""

from __future__ import annotations

import argparse
import json
from collections import Counter
from pathlib import Path
from typing import Any
from zipfile import ZipFile
import xml.etree.ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
EMU_PER_INCH = 914_400


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def get_name(element: ET.Element, kind: str) -> str:
    paths = {
        "sp": "./p:nvSpPr/p:cNvPr",
        "pic": "./p:nvPicPr/p:cNvPr",
        "grpSp": "./p:nvGrpSpPr/p:cNvPr",
        "cxnSp": "./p:nvCxnSpPr/p:cNvPr",
        "graphicFrame": "./p:nvGraphicFramePr/p:cNvPr",
    }
    path = paths.get(kind)
    node = element.find(path, NS) if path else None
    return node.get("name", "") if node is not None else ""


def get_bounds(element: ET.Element, kind: str) -> dict[str, int] | None:
    if kind == "grpSp":
        xfrm = element.find("./p:grpSpPr/a:xfrm", NS)
    elif kind in {"sp", "pic", "cxnSp"}:
        xfrm = element.find("./p:spPr/a:xfrm", NS)
    else:
        xfrm = element.find("./p:xfrm", NS)
    if xfrm is None:
        return None
    off = xfrm.find("./a:off", NS)
    ext = xfrm.find("./a:ext", NS)
    if off is None or ext is None:
        return None
    return {
        "x": int(off.get("x", "0")),
        "y": int(off.get("y", "0")),
        "width": int(ext.get("cx", "0")),
        "height": int(ext.get("cy", "0")),
    }


def slide_relationships(zf: ZipFile, slide_number: int) -> dict[str, str]:
    rel_path = f"ppt/slides/_rels/slide{slide_number}.xml.rels"
    if rel_path not in zf.namelist():
        return {}
    root = ET.fromstring(zf.read(rel_path))
    return {node.get("Id", ""): node.get("Target", "") for node in root}


def inspect_slide(
    zf: ZipFile,
    slide_number: int,
    slide_width: int,
    slide_height: int,
) -> dict[str, Any]:
    root = ET.fromstring(zf.read(f"ppt/slides/slide{slide_number}.xml"))
    relationships = slide_relationships(zf, slide_number)
    records: list[dict[str, Any]] = []
    font_counter: Counter[str] = Counter()
    size_counter: Counter[int] = Counter()
    text_characters = 0
    overflow: list[str] = []

    for element in root.iter():
        kind = local_name(element.tag)
        if kind not in {"sp", "pic", "grpSp", "cxnSp", "graphicFrame"}:
            continue
        name = get_name(element, kind)
        text = "".join(node.text or "" for node in element.findall(".//a:t", NS))
        text_characters += len(text)
        bounds = get_bounds(element, kind)
        geometry = None
        if kind == "sp":
            if element.find("./p:spPr/a:custGeom", NS) is not None:
                geometry = "custom"
            else:
                preset = element.find("./p:spPr/a:prstGeom", NS)
                geometry = preset.get("prst") if preset is not None else None

        fonts: set[str] = set()
        sizes: set[int] = set()
        for props in element.findall(".//a:rPr", NS) + element.findall(".//a:defRPr", NS):
            size_raw = props.get("sz")
            if size_raw:
                size = int(size_raw)
                sizes.add(size)
                size_counter[size] += 1
            for font in (
                props.findall("./a:latin", NS)
                + props.findall("./a:ea", NS)
                + props.findall("./a:cs", NS)
            ):
                typeface = font.get("typeface")
                if typeface:
                    fonts.add(typeface)
                    font_counter[typeface] += 1

        image_target = None
        if kind == "pic":
            blip = element.find(".//a:blip", NS)
            if blip is not None:
                rel_id = blip.get(f"{{{NS['r']}}}embed", "")
                image_target = relationships.get(rel_id)

        if bounds and kind != "grpSp":
            if (
                bounds["x"] < 0
                or bounds["y"] < 0
                or bounds["x"] + bounds["width"] > slide_width
                or bounds["y"] + bounds["height"] > slide_height
            ):
                overflow.append(name or f"unnamed-{kind}")

        records.append(
            {
                "type": kind,
                "name": name,
                "bounds_emu": bounds,
                "geometry": geometry,
                "text": text,
                "fonts": sorted(fonts),
                "font_sizes_hundredth_pt": sorted(sizes),
                "image_target": image_target,
            }
        )

    type_counts = Counter(record["type"] for record in records)
    return {
        "slide_number": slide_number,
        "summary": {
            "objects": len(records),
            "object_types": dict(sorted(type_counts.items())),
            "text_objects": sum(bool(record["text"]) for record in records),
            "text_characters": text_characters,
            "custom_geometries": sum(record["geometry"] == "custom" for record in records),
            "images": type_counts.get("pic", 0),
            "groups": type_counts.get("grpSp", 0),
            "fonts": dict(font_counter.most_common()),
            "font_sizes_hundredth_pt": {
                str(key): value for key, value in size_counter.most_common()
            },
            "potential_overflow_objects": sorted(set(overflow)),
        },
        "objects": records,
    }


def inspect_pptx(path: Path) -> dict[str, Any]:
    with ZipFile(path) as zf:
        presentation = ET.fromstring(zf.read("ppt/presentation.xml"))
        size = presentation.find("p:sldSz", NS)
        if size is None:
            raise ValueError("ppt/presentation.xml does not contain p:sldSz")
        width = int(size.get("cx", "0"))
        height = int(size.get("cy", "0"))
        slide_paths = sorted(
            (name for name in zf.namelist() if name.startswith("ppt/slides/slide") and name.endswith(".xml")),
            key=lambda name: int(Path(name).stem.removeprefix("slide")),
        )
        slides = [
            inspect_slide(zf, index, width, height)
            for index in range(1, len(slide_paths) + 1)
        ]
    return {
        "file": str(path.resolve()),
        "slide_size": {
            "width_emu": width,
            "height_emu": height,
            "width_inches": round(width / EMU_PER_INCH, 4),
            "height_inches": round(height / EMU_PER_INCH, 4),
            "aspect_ratio": round(width / height, 6) if height else None,
        },
        "slide_count": len(slides),
        "slides": slides,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", type=Path, help="PPTX file to inspect")
    parser.add_argument("--output", type=Path, help="Optional JSON output path")
    args = parser.parse_args()
    report = inspect_pptx(args.pptx)
    payload = json.dumps(report, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(payload + "\n", encoding="utf-8")
    else:
        print(payload)


if __name__ == "__main__":
    main()
