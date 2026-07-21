# Correction-plan format

Plans are UTF-8 JSON files consumed by `apply_powerpoint_plan_v4.ps1`. The runner copies the source deck, applies operations in order, verifies protected objects, saves the result, and renders every slide through PowerPoint.

## Top-level structure

```json
{
  "version": 1,
  "description": "Explain the intent and expected input deck.",
  "protected_addresses": ["1/11", "1/12"],
  "operations": []
}
```

An address is slash-delimited: `slide/top-level-shape/group-child/...`. For example, `1/15/9` means slide 1, top-level shape 15, child 9.

Addresses are based on the inspected input deck. Regrouping or inserting shapes can invalidate a later plan; inspect again after structural edits.

## Text operations

Scale every text object on a slide:

```json
{"op": "scale_all_text", "slide_number": 1, "factor": 0.8, "exclude_addresses": []}
```

Change every text object on a slide to one font:

```json
{"op": "set_all_text_font", "slide_number": 1, "font_name": "Noto Sans SC", "exclude_addresses": []}
```

Target one text object:

```json
{"op": "scale_font", "address": "1/6", "factor": 0.95}
{"op": "set_font_size", "address": "1/6", "font_size_pt": 30.0}
{"op": "set_font_name", "address": "1/6", "font_name": "Microsoft YaHei UI"}
{"op": "set_font_spacing", "address": "1/6", "spacing_pt": 1.2}
{"op": "set_font_color", "address": "1/6", "color": "#111111"}
```

Spacing is expressed in points and may be negative. Avoid large negative spacing that makes glyphs overlap even when the bounding box matches.

## Geometry and visibility

Move an existing object by a delta in points:

```json
{"op": "move_shape", "address": "1/6", "delta_left_pt": 2.4, "delta_top_pt": -1.2}
```

Hide or reveal an object:

```json
{"op": "set_shape_visible", "address": "1/15/1", "visible": false}
```

Add an editable rounded rectangle:

```json
{
  "op": "add_rounded_rectangle",
  "slide_number": 1,
  "name": "Card Panel",
  "left_pt": 48.0,
  "top_pt": 250.0,
  "width_pt": 205.0,
  "height_pt": 206.0,
  "radius_adjustment": 0.06,
  "fill_color": "#FFFFFF",
  "fill_transparency": 0.0,
  "line_visible": true,
  "line_color": "#F1F1F1",
  "line_transparency": 0.0,
  "line_weight_pt": 0.7,
  "z_order": 17
}
```

Use `radius_adjustment` near `0.5` for a capsule and a smaller value such as `0.04`–`0.08` for a card. Layer multiple translucent rectangles when the reference uses soft stacked shadows.

## Protected objects

Every supported modifying operation rejects an address listed in `protected_addresses`. Use protection for artwork, photographs, logos, or any object that must remain a picture.

The current fingerprint covers identity, object type, bounds, rotation, and crop geometry. Embedded media-byte hashing is planned for a later release.

## Plan discipline

- State whether a plan expects the original deck or the result of a previous plan.
- Apply operations in deterministic order: global text scale, targeted typography, colors, movement, then structural shapes.
- Keep source files unchanged.
- Review `application-result.json` after every run.
- Stop after two iterations without meaningful score improvement and report the remaining renderer or editability tradeoff.
