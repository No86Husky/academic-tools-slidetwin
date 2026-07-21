from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET
from pathlib import Path

from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "plugins" / "ppt-visual-reconstructor" / "scripts" / "prepare_image_scene.py"


class PrepareImageSceneTests(unittest.TestCase):
    def make_reference(self, directory: Path) -> Path:
        path = directory / "reference.png"
        image = Image.new("RGB", (160, 90), "white")
        draw = ImageDraw.Draw(image)
        draw.rectangle((100, 20, 139, 59), fill="#FF5A16")
        image.save(path)
        return path

    def make_plan(self, directory: Path, full_slide_picture: bool = False) -> Path:
        crop = [0, 0, 160, 90] if full_slide_picture else [100, 20, 40, 40]
        plan = {
            "schema_version": "1.0",
            "reference_image": "reference.png",
            "slide": {
                "width_px": 160,
                "height_px": 90,
                "width_pt": 960,
                "height_pt": 540,
                "background": "#FFFFFF",
            },
            "elements": [
                {
                    "id": "title",
                    "kind": "text",
                    "classification": "native_text",
                    "editable": True,
                    "bounds_px": [10, 5, 80, 15],
                    "z": 2,
                    "text": "Test title",
                    "font": {"family": "Arial", "size_pt": 18, "color": "#111111"},
                    "paragraph": {"align": "left", "vertical_align": "top"},
                },
                {
                    "id": "card",
                    "kind": "shape",
                    "classification": "native_shape",
                    "editable": True,
                    "shape_type": "rounded_rectangle",
                    "bounds_px": [8, 25, 80, 50],
                    "z": 1,
                    "fill": {"color": "#FFFFFF", "opacity": 1},
                    "stroke": {"color": "#E7E7E7", "opacity": 1, "width_pt": 0.5},
                },
                {
                    "id": "artwork",
                    "kind": "image",
                    "classification": "raster_picture",
                    "editable": False,
                    "preserve_as_image": True,
                    "bounds_px": crop,
                    "z": 3,
                    "source": {"type": "reference_crop", "box_px": crop},
                },
            ],
        }
        path = directory / "scene-plan.json"
        path.write_text(json.dumps(plan), encoding="utf-8")
        return path

    def run_prepare(self, reference: Path, plan: Path, output: Path) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            [
                sys.executable,
                str(SCRIPT),
                "--reference",
                str(reference),
                "--plan",
                str(plan),
                "--output-dir",
                str(output),
            ],
            check=False,
            capture_output=True,
            text=True,
        )

    def test_creates_resolved_plan_svg_and_crop(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary)
            reference = self.make_reference(directory)
            plan = self.make_plan(directory)
            output = directory / "output"
            result = self.run_prepare(reference, plan, output)
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
            self.assertTrue((output / "scene-plan.resolved.json").is_file())
            self.assertTrue((output / "semantic-preview.svg").is_file())
            self.assertTrue((output / "assets" / "artwork.png").is_file())
            ET.parse(output / "semantic-preview.svg")
            result_payload = json.loads((output / "scene-preparation-result.json").read_text(encoding="utf-8"))
            self.assertEqual(result_payload["element_count"], 3)
            self.assertEqual(result_payload["editable_element_count"], 2)
            self.assertAlmostEqual(result_payload["editable_element_ratio"], 2 / 3, places=5)
            with Image.open(output / "assets" / "artwork.png") as artwork:
                self.assertEqual(artwork.size, (40, 40))

    def test_rejects_visible_full_slide_raster(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            directory = Path(temporary)
            reference = self.make_reference(directory)
            plan = self.make_plan(directory, full_slide_picture=True)
            output = directory / "output"
            result = self.run_prepare(reference, plan, output)
            self.assertNotEqual(result.returncode, 0)
            self.assertIn("full-slide raster output is not allowed", result.stdout)


if __name__ == "__main__":
    unittest.main()
