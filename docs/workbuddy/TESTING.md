# WorkBuddy integration and Codex-parity test plan

Run these tests on 64-bit Windows with desktop PowerPoint installed.

The target is not merely “the tool runs.” The target is that WorkBuddy executes the same canonical SlideTwin workflow and reaches the same acceptance class as the Codex integration on a controlled reference set.

## A. Static validation

```powershell
python .\tools\validate_workbuddy_plugin.py
```

Expected:

- the WorkBuddy Skill ZIP builds successfully;
- the ZIP contains the canonical Codex shared workflow, scene format, correction-plan format, and Windows bridge instructions;
- the packaged canonical files are byte-identical to their Codex sources;
- the deprecated divergent `scene-plan-format.md` is not packaged;
- Node syntax validation passes;
- the original Codex runtime remains separate and unchanged.

## B. WorkBuddy desktop loading

Upload `dist/slidetwin-workbuddy-skill.zip` through WorkBuddy’s Add Skill interface and restart WorkBuddy after MCP installation.

Expected:

- SlideTwin appears in installed Skills and is enabled;
- `slidetwin-tools` appears under Connectors/MCP;
- `ppt_environment_status` succeeds;
- the MCP tool descriptions include the WorkBuddy visual-review and quality-gate requirements.

## C. PowerPoint probe

Call `ppt_probe_powerpoint` with a new absolute output directory.

Expected:

- a PPTX is created;
- Chinese text and a shape are created without encoding corruption;
- a native PNG render is exported;
- `powerpoint-probe.json` reports success.

## D. Workflow-parity trace

Run one known reference slide and retain the complete tool-call trace.

Required order:

1. `ppt_environment_status`;
2. `ppt_probe_powerpoint` when not already verified;
3. `ppt_prepare_image_scene`;
4. `ppt_build_editable_slide`;
5. `ppt_compare_slide`;
6. `ppt_inspect_powerpoint` before a correction plan;
7. `ppt_apply_correction_plan` when the gate is unmet and differences are correctable;
8. `ppt_compare_slide` again after every correction.

Fail the test when WorkBuddy stops after reading the SVG source, uses Grep as a substitute for visual review, omits the native PowerPoint PNG comparison, or reports completion without comparison evidence.

## E. Visual-evidence review

After each build or correction, confirm that WorkBuddy opens or otherwise visually reviews:

- the native PowerPoint candidate PNG;
- `aligned-reference.png`;
- `overlay.png`;
- `difference-heatmap.png`.

Reading only `comparison.json`, SVG XML, file names, or object manifests does not pass this test.

## F. PASS quality gate

A formal PASS requires all of the following:

- text-content accuracy: 100%;
- protected-picture integrity: 100%;
- no visible garbling;
- no unintended clipping or overflow;
- `pixel_similarity >= 0.90`;
- `composite_visual_score >= 0.90` when the same fonts and PowerPoint renderer are available;
- at least 95% of detected text and simple geometry represented as native PowerPoint objects, unless every exception is documented;
- visual similarity and editable coverage reported separately;
- no full-slide flattening or full-slide raster hidden inside SVG.

Only this state may be labeled:

```text
PASS — quality gate satisfied
```

## G. PLATEAU behavior

Use a slide with a known font or source-asset limitation.

Expected:

- WorkBuddy performs targeted correction rather than immediately stopping;
- after two consecutive correction passes where both pixel similarity and composite score improve by less than 0.003, WorkBuddy may stop as PLATEAU;
- the final response lists actual metrics, failed PASS conditions, remaining mismatch regions, and the limiting reason;
- the response does not claim that the 90% gate was reached.

Required label:

```text
PLATEAU — best current editable reconstruction; formal gate not satisfied
```

## H. FAIL behavior

Induce one controlled failure, such as an unavailable PowerPoint bridge or invalid protected-image source.

Expected:

- WorkBuddy stops without presenting a draft as completed;
- the failure condition and missing evidence are reported;
- the original user file is unchanged.

Required label:

```text
FAIL — reconstruction not safely completed
```

## I. Codex–WorkBuddy controlled comparison

Use the same source image, same Windows computer, same fonts, same render width, and same SlideTwin runtime.

Retain for both hosts:

- scene plan and resolved plan;
- protected assets;
- PPTX;
- native render;
- comparison JSON;
- overlay and heatmap;
- inspection report;
- correction plans;
- final editable coverage.

Compare:

- text accuracy;
- element count and classifications;
- object bounds and font choices;
- pixel similarity;
- composite visual score;
- editable coverage;
- number of correction passes;
- final PASS/PLATEAU/FAIL class.

The initial parity target is:

- the same terminal quality class as Codex;
- WorkBuddy pixel similarity no more than 0.02 below the Codex baseline;
- WorkBuddy composite visual score no more than 0.02 below the Codex baseline;
- editable coverage within 5 percentage points of the Codex baseline;
- identical text accuracy and protected-picture integrity.

These comparative tolerances are development targets; they do not relax the formal PASS gate.

## J. Golden reference set

Maintain 10–20 test slides covering:

- text-and-shape layouts;
- long Chinese text;
- multi-column layouts;
- photographs and complex illustrations;
- artistic lettering;
- icons and vector decorations;
- gradients and textured regions;
- tight alignment and clipping risks.

Each golden case should include the accepted Codex artifacts and expected quality class. Do not merge the WorkBuddy PR as production-equivalent until the full set is run and documented.

## Acceptance matrix

| ID | Test | Pass condition |
| --- | --- | --- |
| W01 | Skill build | deterministic ZIP builds and passes integrity check |
| W02 | Canonical workflow | shared workflow is byte-identical to Codex source |
| W03 | Canonical schemas | scene, correction-plan, and Windows references are identical |
| W04 | MCP startup | `slidetwin-tools` connects without dependency error |
| W05 | PowerPoint probe | PPTX and PNG are generated correctly |
| W06 | Workflow trace | mandatory tool order is present |
| W07 | No Grep substitution | SVG/XML search is not used as visual evidence |
| W08 | Native build | PPTX opens and is not flattened |
| W09 | Visual review | render, aligned reference, overlay, and heatmap are inspected |
| W10 | PPT inspection | inspection occurs before correction |
| W11 | Correction loop | every correction is followed by another comparison |
| W12 | PASS semantics | PASS is used only when every hard gate passes |
| W13 | PLATEAU semantics | unmet gate and two non-improving passes are reported honestly |
| W14 | FAIL semantics | unsafe or incomplete reconstruction is not presented as complete |
| W15 | Codex parity | controlled benchmark remains within defined tolerances |
| W16 | Codex regression | original Codex integration remains functional |
