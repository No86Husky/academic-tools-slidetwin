# Image-only scene plan

The image-only workflow uses one semantic scene plan as the bridge between a flattened reference image, an SVG preview, and a native editable PowerPoint slide.

## Coordinate system

- Express all element bounds in reference-image pixels.
- Set `slide.width_px` and `slide.height_px` to the normalized reference dimensions.
- Set `slide.width_pt` and `slide.height_pt` to the PowerPoint page size. A 16:9 slide normally uses `960 × 540` points.
- The builders scale pixel coordinates into points. Do not mix coordinate systems inside one plan.

## Minimal document

```json
{
  "schema_version": "1.0",
  "reference_image": "reference.png",
  "slide": {
    "width_px": 1600,
    "height_px": 900,
    "width_pt": 960,
    "height_pt": 540,
    "background": "#FFFFFF"
  },
  "quality_gate": {
    "pixel_similarity": 0.9,
    "text_accuracy": 1.0,
    "protected_picture_integrity": 1.0
  },
  "elements": []
}
```

## Common element fields

Every element requires:

- `id`: stable kebab-case identifier.
- `kind`: `text`, `shape`, `line`, or `image`.
- `bounds_px`: `[x, y, width, height]` for every kind except `line`.
- `z`: integer stacking order; lower values are created first.
- `classification`: one of `native_text`, `native_shape`, `svg_object`, or `raster_picture`.
- `editable`: whether the final PowerPoint object is expected to remain editable.

Optional common fields:

- `name`: semantic PowerPoint object name.
- `rotation`: clockwise degrees.
- `opacity`: from `0` to `1`.
- `notes`: reconstruction rationale or ambiguity.

## Text

```json
{
  "id": "slide-title",
  "kind": "text",
  "classification": "native_text",
  "editable": true,
  "bounds_px": [96, 30, 940, 76],
  "z": 30,
  "text": "标题文字",
  "font": {
    "family": "Microsoft YaHei",
    "size_pt": 30,
    "bold": true,
    "italic": false,
    "color": "#111111",
    "spacing_pt": 0
  },
  "paragraph": {
    "align": "left",
    "vertical_align": "middle",
    "margin_left_pt": 0,
    "margin_right_pt": 0,
    "margin_top_pt": 0,
    "margin_bottom_pt": 0,
    "word_wrap": false
  }
}
```

Use one semantic text box for one paragraph or label. Do not split one sentence into one box per character or per word merely to imitate spacing.

## Native shapes

```json
{
  "id": "card-1",
  "kind": "shape",
  "classification": "native_shape",
  "editable": true,
  "shape_type": "rounded_rectangle",
  "bounds_px": [80, 420, 340, 340],
  "z": 10,
  "fill": {"color": "#FFFFFF", "opacity": 1},
  "stroke": {"color": "#E9E9E9", "opacity": 1, "width_pt": 0.6},
  "shadow": {
    "color": "#000000",
    "opacity": 0.08,
    "blur_pt": 6,
    "offset_x_pt": 0,
    "offset_y_pt": 3
  }
}
```

Supported `shape_type` values in the first image-only builder are `rectangle`, `rounded_rectangle`, `ellipse`, `triangle`, `diamond`, and `chevron`.

## Lines

```json
{
  "id": "accent-rule",
  "kind": "line",
  "classification": "native_shape",
  "editable": true,
  "points_px": [108, 548, 138, 548],
  "z": 20,
  "stroke": {"color": "#FF5A16", "opacity": 1, "width_pt": 1.2}
}
```

## Pictures preserved from the reference

```json
{
  "id": "calligraphy-art",
  "kind": "image",
  "classification": "raster_picture",
  "editable": false,
  "preserve_as_image": true,
  "bounds_px": [84, 205, 300, 210],
  "z": 15,
  "source": {
    "type": "reference_crop",
    "box_px": [84, 205, 300, 210]
  }
}
```

The preparation tool crops these regions from the reference and creates separate PNG assets. Use this policy for photographs, brush lettering, complex textures, and illustrations that cannot be reconstructed semantically.

An existing local asset may instead use:

```json
"source": {"type": "file", "path": "assets/logo.png"}
```

For a complex vector icon that should remain one scalable PowerPoint graphic, use `kind: "image"`, `classification: "svg_object"`, and a local `.svg` file source. The first builder inserts it as one SVG graphic; do not misreport it as native PowerPoint geometry.

## Classification rules

1. Titles, paragraphs, labels, numbers, and captions become `native_text`.
2. Cards, circles, rules, arrows, and flat-color geometric decorations become `native_shape`.
3. Simple vector icons may remain one `svg_object` when converting them into many native primitives would reduce usability.
4. Photographs, artistic lettering, gradients with texture, and complex illustrations become `raster_picture`.
5. Never use the full reference image as a visible slide-sized background merely to pass the visual metric. A full-slide reference may be retained only as an excluded or hidden guide during reconstruction.

## Image-only acceptance report

Report at least:

- Pixel similarity and composite visual score.
- Native text count and text-content accuracy.
- Native shape count.
- SVG-object count.
- Protected raster-picture count and integrity result.
- Estimated editable coverage by semantic element count and by visible area.
- Remaining mismatched regions and the reason they were not further flattened.
