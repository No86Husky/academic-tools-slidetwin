# SlideTwin scene plan format

The host model writes a UTF-8 JSON file that describes one reference slide in pixel coordinates.

## Minimal document

```json
{
  "schema_version": "1.0",
  "slide": {
    "width_px": 1600,
    "height_px": 900,
    "width_pt": 960,
    "height_pt": 540,
    "background": "#FFFFFF"
  },
  "elements": []
}
```

## Common element fields

Every element should contain:

```json
{
  "id": "unique-id",
  "name": "human-readable-name",
  "kind": "text",
  "classification": "native_text",
  "z": 10,
  "bounds_px": [100, 80, 600, 90],
  "rotation": 0
}
```

`bounds_px` uses `[left, top, width, height]` in reference-image pixels.

## Text element

```json
{
  "id": "title",
  "name": "Main title",
  "kind": "text",
  "classification": "native_text",
  "z": 20,
  "bounds_px": [120, 70, 900, 90],
  "text": "Example title",
  "font": {
    "family": "Microsoft YaHei",
    "size_pt": 30,
    "bold": true,
    "italic": false,
    "color": "#17365D",
    "spacing_pt": 0
  },
  "paragraph": {
    "align": "left",
    "vertical_align": "middle",
    "word_wrap": true,
    "margin_left_pt": 0,
    "margin_right_pt": 0,
    "margin_top_pt": 0,
    "margin_bottom_pt": 0
  }
}
```

## Shape element

```json
{
  "id": "card-1",
  "name": "Card 1",
  "kind": "shape",
  "classification": "native_shape",
  "shape_type": "rounded_rectangle",
  "z": 5,
  "bounds_px": [100, 220, 420, 250],
  "fill": {"color": "#F4F7FB", "opacity": 1},
  "stroke": {"color": "#9EB6CE", "width_pt": 1, "opacity": 1},
  "shadow": null
}
```

Supported simple shape types include `rectangle`, `rounded_rectangle`, `ellipse`, `triangle`, `diamond`, and `chevron`.

## Line element

```json
{
  "id": "arrow-1",
  "name": "Arrow 1",
  "kind": "line",
  "classification": "native_shape",
  "z": 8,
  "points_px": [520, 340, 700, 340],
  "stroke": {
    "color": "#4F81BD",
    "width_pt": 2,
    "opacity": 1,
    "dash_style": "solid",
    "end_arrow": true
  }
}
```

## Protected image element

```json
{
  "id": "illustration-1",
  "name": "Protected illustration",
  "kind": "image",
  "classification": "raster_picture",
  "z": 15,
  "bounds_px": [1050, 160, 420, 520],
  "source": {
    "crop_from_reference": [1050, 160, 420, 520]
  }
}
```

Use protected images only for photographs, artistic lettering, complex illustrations, textures, or detail-rich regions. Do not crop the whole slide as one protected image.

## General rules

- Use unique IDs.
- Preserve reading order and visual layering through `z`.
- Use `#RRGGBB` colors.
- Keep all coordinates within the reference canvas.
- Mark guide-only objects with `guide_only: true` and objects excluded from the final slide with `exclude_from_final: true`.
