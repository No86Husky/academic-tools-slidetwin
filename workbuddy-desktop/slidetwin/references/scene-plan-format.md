# Deprecated compatibility pointer

Do not use this file as a scene-plan specification.

The WorkBuddy Skill ZIP is built with the canonical Codex SlideTwin scene specification at:

```text
references/scene-format.md
```

The canonical file is copied directly from:

```text
plugins/ppt-visual-reconstructor/skills/ppt-visual-reconstructor/references/scene-format.md
```

This compatibility pointer remains only so older test scripts and previously documented paths fail clearly instead of silently using a divergent schema.

In particular, protected reference crops must use the canonical structure:

```json
{
  "kind": "image",
  "classification": "raster_picture",
  "source": {
    "type": "reference_crop",
    "box_px": [1050, 160, 420, 520]
  }
}
```

Do not use the obsolete `source.crop_from_reference` form.
