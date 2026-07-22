---
name: slidetwin
description: Primary SlideTwin workflow for recreating a high-fidelity editable PowerPoint slide from one reference PNG/JPEG, or repairing an existing SVG/PPTX reconstruction. Use when Codex must preserve complex imagery as pictures, rebuild text and simple geometry as native PowerPoint objects, render with desktop PowerPoint, and iterate toward at least 90% pixel similarity.
---

# SlideTwin

SlideTwin creates an editable visual twin of a reference slide.

Before taking reconstruction actions, read [the complete shared workflow](../ppt-visual-reconstructor/SKILL.md) and follow all of its object policy, installed-tool sequence, quality gates, safety rules, references, and output contract. The shared file retains its historical folder name solely so the original `$ppt-visual-reconstructor` invocation used in existing prompts and videos remains compatible.

Use `$slidetwin` as the primary invocation for new work. Do not treat the compatibility alias as a different reconstruction method.
