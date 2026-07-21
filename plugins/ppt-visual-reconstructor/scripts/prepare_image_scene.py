#!/usr/bin/env python3
"""Prepare protected image crops and a semantic SVG from an image-only scene plan."""

from __future__ import annotations

import argparse
import base64
import copy
import html
import json
import math
import re
from pathlib import Path
from typing import Any

import numpy as np
from PIL import Image


SUPPORTED_KINDS = {"text", "shape", "line", "image"}
SUPPORTED_SHAPES = {
    "rectangle",
    "rounded_rectangle",
    "ellipse",
    "triangle",
    "diamond",
    "chevron",
}
HEX_COLOR = re.compile(r"^#[0-9A-Fa-f]{6}$")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Validate an image-only scene plan, crop protected imagery, and write a semantic SVG preview."
    )
    parser.add_argument("--reference", required=True, type=Path)
    parser.add_argument("--plan", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    return parser.parse_args()


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def load_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8-sig") as handle:
        value = json.load(handle)
    require(isinstance(value, dict), "Scene plan root must be a JSON object")
    return value


def validate_color(value: str, field: str) -> str:
    require(isinstance(value, str) and HEX_COLOR.match(value) is not None, f"{field} must be #RRGGBB")
    return value.upper()


def clamp01(value: Any, default: float = 1.0) -> float:
    if value is None:
        return default
    number = float(value)
    return max(0.0, min(1.0, number))


def parse_box(value: Any, field: str) -> list[float]:
    require(isinstance(value, list) and len(value) == 4, f"{field} must contain four numbers")
    box = [float(item) for item in value]
    require(all(math.isfinite(item) for item in box), f"{field} contains a non-finite number")
    require(box[2] > 0 and box[3] > 0, f"{field} width and height must be positive")
    return box


def safe_id(value: str) -> str:
    normalized = re.sub(r"[^a-zA-Z0-9._-]+", "-", value).strip("-.")
    return normalized or "element"


def color_to_rgb(value: str) -> np.ndarray:
    value = validate_color(value, "background color")
    return np.array([int(value[1:3], 16), int(value[3:5], 16), int(value[5:7], 16)], dtype=np.float32)


def infer_corner_color(image: Image.Image) -> str:
    rgb = np.asarray(image.convert("RGB"), dtype=np.uint8)
    samples = np.stack([rgb[0, 0], rgb[0, -1], rgb[-1, 0], rgb[-1, -1]], axis=0)
    median = np.median(samples, axis=0).round().astype(np.uint8)
    return f"#{median[0]:02X}{median[1]:02X}{median[2]:02X}"


def make_background_transparent(
    image: Image.Image,
    background_color: str | None,
    tolerance: float,
    feather: float,
) -> Image.Image:
    rgba = np.asarray(image.convert("RGBA"), dtype=np.uint8).copy()
    background = color_to_rgb(background_color or infer_corner_color(image))
    distance = np.linalg.norm(rgba[:, :, :3].astype(np.float32) - background[None, None, :], axis=2)
    feather = max(1.0, feather)
    alpha_factor = np.clip((distance - tolerance) / feather, 0.0, 1.0)
    rgba[:, :, 3] = np.minimum(rgba[:, :, 3], np.round(alpha_factor * 255).astype(np.uint8))
    return Image.fromarray(rgba, mode="RGBA")


def resolve_source_file(path_value: str, plan_path: Path) -> Path:
    source = Path(path_value)
    if not source.is_absolute():
        source = (plan_path.parent / source).resolve()
    require(source.is_file(), f"Image source file does not exist: {source}")
    return source


def crop_reference_asset(
    reference: Image.Image,
    element: dict[str, Any],
    output_path: Path,
) -> None:
    source = element["source"]
    box = parse_box(source.get("box_px", element.get("bounds_px")), f"{element['id']}.source.box_px")
    left, top, width, height = box
    crop_box = (
        max(0, int(round(left))),
        max(0, int(round(top))),
        min(reference.width, int(round(left + width))),
        min(reference.height, int(round(top + height))),
    )
    require(crop_box[2] > crop_box[0] and crop_box[3] > crop_box[1], f"Empty crop for {element['id']}")
    cropped = reference.crop(crop_box)
    if bool(source.get("make_background_transparent", False)):
        cropped = make_background_transparent(
            cropped,
            source.get("background_color"),
            float(source.get("tolerance", 18)),
            float(source.get("feather", 18)),
        )
    cropped.save(output_path)


def mime_type_for(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".svg":
        return "image/svg+xml"
    if suffix in {".jpg", ".jpeg"}:
        return "image/jpeg"
    return "image/png"


def data_uri(path: Path) -> str:
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime_type_for(path)};base64,{encoded}"


def svg_style_color(style: dict[str, Any] | None, default: str = "none") -> tuple[str, float]:
    if not style:
        return default, 1.0
    color = style.get("color", default)
    if color != "none":
        color = validate_color(color, "style.color")
    return color, clamp01(style.get("opacity"), 1.0)


def shape_svg(element: dict[str, Any]) -> str:
    x, y, width, height = parse_box(element["bounds_px"], f"{element['id']}.bounds_px")
    fill_color, fill_opacity = svg_style_color(element.get("fill"), "none")
    stroke_color, stroke_opacity = svg_style_color(element.get("stroke"), "none")
    stroke_width = float((element.get("stroke") or {}).get("width_pt", 0))
    common = (
        f'fill="{fill_color}" fill-opacity="{fill_opacity:.4f}" '
        f'stroke="{stroke_color}" stroke-opacity="{stroke_opacity:.4f}" '
        f'stroke-width="{stroke_width:.4f}" opacity="{clamp01(element.get("opacity")):.4f}"'
    )
    shape_type = element.get("shape_type", "rectangle")
    if shape_type == "rounded_rectangle":
        radius = float(element.get("corner_radius_px", min(width, height) * 0.06))
        return f'<rect x="{x}" y="{y}" width="{width}" height="{height}" rx="{radius}" {common}/>'
    if shape_type == "ellipse":
        return f'<ellipse cx="{x + width / 2}" cy="{y + height / 2}" rx="{width / 2}" ry="{height / 2}" {common}/>'
    if shape_type == "triangle":
        points = f"{x + width / 2},{y} {x + width},{y + height} {x},{y + height}"
        return f'<polygon points="{points}" {common}/>'
    if shape_type == "diamond":
        points = f"{x + width / 2},{y} {x + width},{y + height / 2} {x + width / 2},{y + height} {x},{y + height / 2}"
        return f'<polygon points="{points}" {common}/>'
    if shape_type == "chevron":
        notch = width * 0.28
        points = (
            f"{x},{y} {x + width - notch},{y} {x + width},{y + height / 2} "
            f"{x + width - notch},{y + height} {x},{y + height} {x + notch},{y + height / 2}"
        )
        return f'<polygon points="{points}" {common}/>'
    return f'<rect x="{x}" y="{y}" width="{width}" height="{height}" {common}/>'


def line_svg(element: dict[str, Any]) -> str:
    points = element.get("points_px")
    require(isinstance(points, list) and len(points) == 4, f"{element['id']}.points_px must contain four numbers")
    x1, y1, x2, y2 = [float(item) for item in points]
    stroke = element.get("stroke") or {}
    color, opacity = svg_style_color(stroke, "#000000")
    width = float(stroke.get("width_pt", 1))
    return (
        f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" '
        f'stroke="{color}" stroke-opacity="{opacity:.4f}" stroke-width="{width:.4f}"/>'
    )


def text_svg(element: dict[str, Any], px_per_pt: float) -> str:
    x, y, width, height = parse_box(element["bounds_px"], f"{element['id']}.bounds_px")
    font = element.get("font") or {}
    paragraph = element.get("paragraph") or {}
    size_px = float(font.get("size_pt", 18)) * px_per_pt
    family = html.escape(str(font.get("family", "Microsoft YaHei")), quote=True)
    color = validate_color(str(font.get("color", "#000000")), f"{element['id']}.font.color")
    weight = "700" if bool(font.get("bold", False)) else "400"
    style = "italic" if bool(font.get("italic", False)) else "normal"
    margin_left = float(paragraph.get("margin_left_pt", 0)) * px_per_pt
    margin_top = float(paragraph.get("margin_top_pt", 0)) * px_per_pt
    align = str(paragraph.get("align", "left"))
    if align == "center":
        text_x = x + width / 2
        anchor = "middle"
    elif align == "right":
        text_x = x + width - margin_left
        anchor = "end"
    else:
        text_x = x + margin_left
        anchor = "start"
    lines = str(element.get("text", "")).splitlines() or [""]
    line_height = float(paragraph.get("line_height", 1.2)) * size_px
    baseline = y + margin_top + size_px
    tspans = []
    for index, line in enumerate(lines):
        dy = 0 if index == 0 else line_height
        tspans.append(f'<tspan x="{text_x}" dy="{dy}">{html.escape(line)}</tspan>')
    spacing = float(font.get("spacing_pt", 0)) * px_per_pt
    return (
        f'<text x="{text_x}" y="{baseline}" text-anchor="{anchor}" '
        f'font-family="{family}" font-size="{size_px:.4f}" font-weight="{weight}" '
        f'font-style="{style}" fill="{color}" letter-spacing="{spacing:.4f}" '
        f'opacity="{clamp01(element.get("opacity")):.4f}">{"".join(tspans)}</text>'
    )


def image_svg(element: dict[str, Any]) -> str:
    x, y, width, height = parse_box(element["bounds_px"], f"{element['id']}.bounds_px")
    asset_path = Path(element["source"]["asset_path"])
    preserve = str(element.get("preserve_aspect", "none"))
    preserve_value = "xMidYMid meet" if preserve == "meet" else "xMidYMid slice" if preserve == "slice" else "none"
    return (
        f'<image x="{x}" y="{y}" width="{width}" height="{height}" '
        f'href="{data_uri(asset_path)}" preserveAspectRatio="{preserve_value}" '
        f'opacity="{clamp01(element.get("opacity")):.4f}"/>'
    )


def render_svg(plan: dict[str, Any], output_path: Path) -> None:
    slide = plan["slide"]
    width_px = float(slide["width_px"])
    height_px = float(slide["height_px"])
    width_pt = float(slide["width_pt"])
    px_per_pt = width_px / width_pt
    background = validate_color(str(slide.get("background", "#FFFFFF")), "slide.background")
    fragments = [f'<rect x="0" y="0" width="{width_px}" height="{height_px}" fill="{background}"/>']
    for element in sorted(plan["elements"], key=lambda item: (int(item.get("z", 0)), item["id"])):
        if bool(element.get("exclude_from_final", False)) or bool(element.get("guide_only", False)):
            continue
        kind = element["kind"]
        if kind == "shape":
            fragments.append(shape_svg(element))
        elif kind == "line":
            fragments.append(line_svg(element))
        elif kind == "text":
            fragments.append(text_svg(element, px_per_pt))
        elif kind == "image":
            fragments.append(image_svg(element))
    document = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width_px}" height="{height_px}" '
        f'viewBox="0 0 {width_px} {height_px}">\n'
        f'{chr(10).join(fragments)}\n</svg>\n'
    )
    output_path.write_text(document, encoding="utf-8")


def validate_and_resolve(
    plan: dict[str, Any],
    plan_path: Path,
    reference_path: Path,
    reference: Image.Image,
    assets_dir: Path,
) -> tuple[dict[str, Any], list[str]]:
    resolved = copy.deepcopy(plan)
    require(str(resolved.get("schema_version")) == "1.0", "schema_version must be 1.0")
    slide = resolved.get("slide")
    require(isinstance(slide, dict), "slide is required")
    width_px = int(round(float(slide.get("width_px", reference.width))))
    height_px = int(round(float(slide.get("height_px", reference.height))))
    require(width_px == reference.width and height_px == reference.height, "Plan pixel canvas must match the reference image")
    slide["width_px"] = width_px
    slide["height_px"] = height_px
    slide.setdefault("width_pt", 960)
    slide.setdefault("height_pt", float(slide["width_pt"]) * height_px / width_px)
    slide["background"] = validate_color(str(slide.get("background", "#FFFFFF")), "slide.background")

    elements = resolved.get("elements")
    require(isinstance(elements, list), "elements must be an array")
    seen: set[str] = set()
    warnings: list[str] = []
    for index, element in enumerate(elements):
        require(isinstance(element, dict), f"elements[{index}] must be an object")
        element_id = str(element.get("id", "")).strip()
        require(element_id != "", f"elements[{index}].id is required")
        require(element_id not in seen, f"Duplicate element id: {element_id}")
        seen.add(element_id)
        kind = str(element.get("kind", ""))
        require(kind in SUPPORTED_KINDS, f"Unsupported kind for {element_id}: {kind}")
        element.setdefault("name", element_id)
        element.setdefault("z", index)
        element.setdefault("opacity", 1)
        if kind != "line":
            bounds = parse_box(element.get("bounds_px"), f"{element_id}.bounds_px")
            x, y, width, height = bounds
            if x < 0 or y < 0 or x + width > width_px or y + height > height_px:
                warnings.append(f"{element_id} extends outside the slide canvas")
            element["bounds_px"] = bounds
        if kind == "shape":
            shape_type = str(element.get("shape_type", "rectangle"))
            require(shape_type in SUPPORTED_SHAPES, f"Unsupported shape_type for {element_id}: {shape_type}")
        if kind == "text":
            require("text" in element, f"{element_id}.text is required")
        if kind == "image":
            source = element.get("source")
            require(isinstance(source, dict), f"{element_id}.source is required")
            source_type = str(source.get("type", ""))
            if source_type == "reference_crop":
                bounds = parse_box(source.get("box_px", element["bounds_px"]), f"{element_id}.source.box_px")
                area_ratio = (bounds[2] * bounds[3]) / (width_px * height_px)
                require(
                    area_ratio < 0.95 or bool(element.get("guide_only", False)),
                    f"{element_id} covers nearly the full slide; full-slide raster output is not allowed",
                )
                asset_path = assets_dir / f"{safe_id(element_id)}.png"
                crop_reference_asset(reference, element, asset_path)
            elif source_type == "file":
                asset_path = resolve_source_file(str(source.get("path", "")), plan_path)
            else:
                raise ValueError(f"Unsupported source.type for {element_id}: {source_type}")
            source["asset_path"] = str(asset_path.resolve())

    resolved["reference_image"] = str(reference_path.resolve())
    return resolved, warnings


def main() -> int:
    args = parse_args()
    reference_path = args.reference.resolve()
    plan_path = args.plan.resolve()
    output_dir = args.output_dir.resolve()
    require(reference_path.is_file(), f"Reference image does not exist: {reference_path}")
    require(plan_path.is_file(), f"Scene plan does not exist: {plan_path}")
    output_dir.mkdir(parents=True, exist_ok=True)
    assets_dir = output_dir / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)

    with Image.open(reference_path) as source_image:
        reference = source_image.convert("RGBA")
    plan = load_json(plan_path)
    resolved, warnings = validate_and_resolve(plan, plan_path, reference_path, reference, assets_dir)

    resolved_path = output_dir / "scene-plan.resolved.json"
    resolved_path.write_text(json.dumps(resolved, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    svg_path = output_dir / "semantic-preview.svg"
    render_svg(resolved, svg_path)

    counts: dict[str, int] = {}
    final_elements = [
        element
        for element in resolved["elements"]
        if not bool(element.get("exclude_from_final", False)) and not bool(element.get("guide_only", False))
    ]
    editable_count = 0
    editable_area = 0.0
    visible_area = 0.0
    for element in final_elements:
        classification = str(element.get("classification", element["kind"]))
        counts[classification] = counts.get(classification, 0) + 1
        editable = bool(element.get("editable", classification in {"native_text", "native_shape"}))
        editable_count += int(editable)
        if element["kind"] != "line":
            _, _, width, height = parse_box(element["bounds_px"], f"{element['id']}.bounds_px")
            area = width * height
            visible_area += area
            if editable:
                editable_area += area
    result = {
        "success": True,
        "reference_image": str(reference_path),
        "resolved_plan": str(resolved_path),
        "semantic_svg": str(svg_path),
        "assets_directory": str(assets_dir),
        "element_count": len(final_elements),
        "editable_element_count": editable_count,
        "editable_element_ratio": round(editable_count / len(final_elements), 6) if final_elements else 1.0,
        "editable_area_ratio_estimate": round(editable_area / visible_area, 6) if visible_area else 1.0,
        "classification_counts": counts,
        "warnings": warnings,
    }
    result_path = output_dir / "scene-preparation-result.json"
    result_path.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:  # noqa: BLE001
        print(json.dumps({"success": False, "error": str(exc), "error_type": type(exc).__name__}, ensure_ascii=False, indent=2))
        raise SystemExit(1)
