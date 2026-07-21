#!/usr/bin/env python3
"""Compare a reference slide image with a PowerPoint-rendered candidate."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

import numpy as np
from PIL import Image


def fit_reference(reference: Image.Image, size: tuple[int, int], mode: str) -> Image.Image:
    target_width, target_height = size
    reference = reference.convert("RGB")
    if mode == "stretch":
        return reference.resize(size, Image.Resampling.LANCZOS)

    source_ratio = reference.width / reference.height
    target_ratio = target_width / target_height
    if mode == "contain":
        if source_ratio > target_ratio:
            width = target_width
            height = round(width / source_ratio)
        else:
            height = target_height
            width = round(height * source_ratio)
        resized = reference.resize((width, height), Image.Resampling.LANCZOS)
        canvas = Image.new("RGB", size, "white")
        canvas.paste(resized, ((target_width - width) // 2, (target_height - height) // 2))
        return canvas

    if mode == "cover":
        if source_ratio > target_ratio:
            height = target_height
            width = round(height * source_ratio)
        else:
            width = target_width
            height = round(width / source_ratio)
        resized = reference.resize((width, height), Image.Resampling.LANCZOS)
        left = (width - target_width) // 2
        top = (height - target_height) // 2
        return resized.crop((left, top, left + target_width, top + target_height))

    raise ValueError(f"Unsupported fit mode: {mode}")


def global_ssim(first: np.ndarray, second: np.ndarray) -> float:
    first_gray = first.mean(axis=2)
    second_gray = second.mean(axis=2)
    mean_first = float(first_gray.mean())
    mean_second = float(second_gray.mean())
    variance_first = float(first_gray.var())
    variance_second = float(second_gray.var())
    covariance = float(((first_gray - mean_first) * (second_gray - mean_second)).mean())
    c1 = (0.01 * 255) ** 2
    c2 = (0.03 * 255) ** 2
    numerator = (2 * mean_first * mean_second + c1) * (2 * covariance + c2)
    denominator = (mean_first**2 + mean_second**2 + c1) * (
        variance_first + variance_second + c2
    )
    return numerator / denominator if denominator else 1.0


def edge_map(image: np.ndarray) -> np.ndarray:
    gray = image.mean(axis=2)
    dx = np.zeros_like(gray)
    dy = np.zeros_like(gray)
    dx[:, 1:] = np.abs(gray[:, 1:] - gray[:, :-1])
    dy[1:, :] = np.abs(gray[1:, :] - gray[:-1, :])
    return np.sqrt(dx * dx + dy * dy) / np.sqrt(2 * 255 * 255)


def difference_bounds(mask: np.ndarray) -> dict[str, int] | None:
    ys, xs = np.where(mask)
    if xs.size == 0:
        return None
    return {
        "left": int(xs.min()),
        "top": int(ys.min()),
        "right": int(xs.max()) + 1,
        "bottom": int(ys.max()) + 1,
        "width": int(xs.max() - xs.min() + 1),
        "height": int(ys.max() - ys.min() + 1),
    }


def compare(reference_path: Path, candidate_path: Path, output_dir: Path, mode: str) -> dict:
    candidate_image = Image.open(candidate_path).convert("RGB")
    reference_image = fit_reference(Image.open(reference_path), candidate_image.size, mode)
    reference = np.asarray(reference_image, dtype=np.float32)
    candidate = np.asarray(candidate_image, dtype=np.float32)
    absolute = np.abs(reference - candidate)
    normalized = absolute / 255.0

    mae = float(normalized.mean())
    rmse = float(np.sqrt(np.mean(normalized * normalized)))
    pixel_similarity = max(0.0, 1.0 - mae)
    ssim = float(global_ssim(reference, candidate))
    reference_edges = edge_map(reference)
    candidate_edges = edge_map(candidate)
    edge_similarity = max(0.0, 1.0 - float(np.abs(reference_edges - candidate_edges).mean()))
    visual_score = (
        0.45 * max(0.0, min(1.0, ssim))
        + 0.35 * pixel_similarity
        + 0.20 * edge_similarity
    )

    max_channel_difference = absolute.max(axis=2)
    threshold_mask = max_channel_difference >= 30
    different_pixel_ratio = float(threshold_mask.mean())

    output_dir.mkdir(parents=True, exist_ok=True)
    aligned_reference_path = output_dir / "aligned-reference.png"
    overlay_path = output_dir / "overlay.png"
    heatmap_path = output_dir / "difference-heatmap.png"
    report_path = output_dir / "comparison.json"

    reference_image.save(aligned_reference_path)
    Image.blend(reference_image, candidate_image, 0.5).save(overlay_path)

    magnitude = np.clip(max_channel_difference / 96.0, 0.0, 1.0)
    heatmap = np.empty_like(candidate, dtype=np.uint8)
    heatmap[:, :, 0] = (255 * magnitude).astype(np.uint8)
    heatmap[:, :, 1] = (40 * (1.0 - magnitude)).astype(np.uint8)
    heatmap[:, :, 2] = (40 * (1.0 - magnitude)).astype(np.uint8)
    Image.fromarray(heatmap, mode="RGB").save(heatmap_path)

    report = {
        "reference_path": str(reference_path.resolve()),
        "candidate_path": str(candidate_path.resolve()),
        "fit_mode": mode,
        "canvas": {"width": candidate_image.width, "height": candidate_image.height},
        "metrics": {
            "mean_absolute_error": round(mae, 6),
            "root_mean_square_error": round(rmse, 6),
            "pixel_similarity": round(pixel_similarity, 6),
            "global_ssim": round(ssim, 6),
            "edge_similarity": round(edge_similarity, 6),
            "composite_visual_score": round(visual_score, 6),
            "different_pixel_ratio_at_30": round(different_pixel_ratio, 6),
        },
        "difference_bounds_px": difference_bounds(threshold_mask),
        "outputs": {
            "aligned_reference": str(aligned_reference_path.resolve()),
            "overlay": str(overlay_path.resolve()),
            "difference_heatmap": str(heatmap_path.resolve()),
        },
    }
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--reference", type=Path, required=True)
    parser.add_argument("--candidate", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--fit", choices=("contain", "cover", "stretch"), default="contain")
    args = parser.parse_args()
    report = compare(args.reference, args.candidate, args.output_dir, args.fit)
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
