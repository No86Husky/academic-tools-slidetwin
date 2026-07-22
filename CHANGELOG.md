# Changelog

## 0.5.0

- Renamed the public product to SlideTwin to avoid confusion with similarly named projects.
- Added `$slidetwin` as the primary Skill invocation.
- Retained `$ppt-visual-reconstructor`, the plugin ID, and marketplace ID as compatibility identifiers for existing installations and demo videos.
- Added the public image-only demo, editable PPTX, visual comparison, difference heatmap, workflow GIF, and reproducible metrics.
- Updated the MCP identity, installers, validation, and documentation for the SlideTwin release.

## 0.4.0

- Added the unique `ppt-visual-tools` Codex marketplace.
- Added a Windows installer with explicit prerequisite installation and optional PowerPoint probing.
- Added a zero-dependency Node.js MCP server with eight reconstruction and validation tools.
- Made MCP tools the primary Skill workflow while preserving direct scripts as fallbacks.
- Added MCP protocol, manifest, marketplace, and installer validation to CI.

## 0.3.0

- Made a single reference PNG or JPEG the only required input.
- Added a semantic scene-plan format for native text, shapes, lines, SVG objects, and protected raster regions.
- Added protected-region extraction and semantic SVG preview generation.
- Added blank-presentation reconstruction through desktop PowerPoint COM.
- Added an explicit image-only output contract and editable-coverage acceptance gate.
- Kept the v0.2 native render, comparison, protected-object, typography, and iterative correction workflow.

## 0.2.0

- Verified Windows 11 x64 and PowerPoint 2024 x64 automation.
- Added recursive native PowerPoint inspection and rendering.
- Added reference-image alignment, overlays, heatmaps, and visual metrics.
- Added non-destructive JSON correction plans with protected-object checks.
- Added text size, family, spacing, color, movement, visibility, and rounded-rectangle operations.
- Added installed-font probing and batch font calibration.
- Added Windows and correction-plan references.

## 0.1.0

- Initial plugin and skill scaffold.
- Added cross-platform OOXML structure inspection.
- Added PowerPoint COM environment probe.
