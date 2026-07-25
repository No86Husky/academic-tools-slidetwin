# SlideTwin quality gate for WorkBuddy

This file converts the Codex SlideTwin acceptance policy into explicit terminal states for WorkBuddy.

## Required evidence

A reconstruction cannot be classified until all of the following evidence exists:

- an editable PPTX generated or corrected by desktop PowerPoint;
- the native PowerPoint PNG render used as the candidate image;
- `comparison.json` from `ppt_compare_slide`;
- `overlay.png` and `difference-heatmap.png`;
- a PowerPoint inspection report;
- an editability report or reliable native-object counts.

The semantic SVG alone is never sufficient evidence.

## PASS

Classify the result as `PASS` only when every hard condition below is satisfied:

1. `metrics.pixel_similarity >= 0.90`.
2. `metrics.composite_visual_score >= 0.90` when the same fonts and the desktop PowerPoint renderer are available. If native font rasterization is the only documented source of a lower composite score, the result may not be called PASS without an explicit human visual review.
3. Text-content accuracy is `1.00` for all readable reference text.
4. Protected-picture integrity is `1.00`.
5. There is no visible garbling.
6. There is no unintended clipping or overflow.
7. No object that should be visible lies outside the slide canvas.
8. The whole reference slide has not been flattened into one visible picture or one raster image embedded inside an SVG.
9. At least 95% of detected text and simple geometric elements are represented by native PowerPoint objects, unless each exception is documented.
10. Visual similarity and editable coverage are reported separately.

Only PASS may be described as “达到 SlideTwin 质量门槛” or “达到 90% 验收要求”.

## PLATEAU

Classify the result as `PLATEAU` when:

- one or more PASS conditions remain unmet; and
- two consecutive correction passes produce no meaningful visual or structural improvement; or
- the remaining difference is caused by unavailable fonts, unreadable source content, unavailable original assets, or a rendering limitation that cannot be corrected without violating editability.

For reproducible automation, treat a pass as having no meaningful metric improvement when both:

- `pixel_similarity` improves by less than `0.003`; and
- `composite_visual_score` improves by less than `0.003`.

This numeric plateau rule is a WorkBuddy orchestration supplement. It does not relax any PASS requirement.

A PLATEAU result must report:

- the actual metrics;
- every failed PASS condition;
- the largest remaining mismatch regions;
- the reason further correction stopped;
- whether additional user assets or fonts could improve the result.

A PLATEAU result must not claim that 90% acceptance was reached.

## FAIL

Classify the result as `FAIL` when any of the following occurs:

- the MCP runtime or required executable is unavailable;
- the PowerPoint COM probe fails;
- the PPTX cannot be created, opened, saved, or rendered;
- comparison evidence is missing or invalid;
- readable text is garbled or materially mistranscribed and cannot be corrected;
- protected pictures are corrupted or replaced;
- the output is a prohibited full-slide flattening;
- the final PPTX is missing or unusable.

FAIL must be reported directly. Do not return a draft file as though it were a completed reconstruction.

## Mandatory final statement

The final response must include exactly one status label:

- `PASS — quality gate satisfied`
- `PLATEAU — best current editable reconstruction; formal gate not satisfied`
- `FAIL — reconstruction not safely completed`

The statement must cite the measured metrics and not infer them from appearance alone.
