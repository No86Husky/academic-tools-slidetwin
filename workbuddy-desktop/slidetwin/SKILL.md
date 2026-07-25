---
name: slidetwin
description: 将单页PPT参考图片高保真重建为可编辑PowerPoint。适用于用户上传PNG、JPG或幻灯片截图，并要求文字、简单图形和连线可编辑，复杂插图或艺术字作为独立图片保留的任务。
---

# SlideTwin：WorkBuddy 质量等价工作流

根据用户上传的单页参考图片，调用已连接的 `slidetwin-tools` MCP，在 Windows 桌面版 PowerPoint 中生成一页高保真、可编辑的 PPTX。

用户补充要求：

$ARGUMENTS

## 唯一权威工作流

在执行任何重建操作前，必须完整读取并遵循：

1. `references/shared-workflow.md`：由 Codex 版完整 SlideTwin 工作流原样打包而来；其中出现的 “Codex” 在本 Skill 中均指当前 WorkBuddy 宿主模型。
2. `references/scene-format.md`：唯一合法的 scene-plan 格式。
3. `references/plan-format.md`：correction plan 格式。
4. `references/windows-bridge.md`：Windows PowerPoint 自动化要求。
5. `references/quality-gate.md`：PASS、PLATEAU、FAIL 的机器化验收规则。

本文件只负责 WorkBuddy 入口和宿主适配。对象政策、工具顺序、质量门槛、安全规则和输出契约以这些共享文件为准，不得自行简化。

## 使用前检查

1. 必须能够访问用户上传的 PNG 或 JPEG，并取得绝对本地路径。
2. 首先调用 `ppt_environment_status`。
3. 若当前安装尚未通过 PowerPoint COM 探针，则调用 `ppt_probe_powerpoint`。
4. 每次任务建立独立的绝对输出目录，不覆盖用户原文件或上一轮结果。
5. MCP、PowerPoint、参考图片或必要字体检查失败时，停止并返回 `FAIL`，不得假装已完成。

## 强制执行顺序

1. 以原始分辨率检查参考图，完整转录所有可读文字。
2. 将每个可见元素分类为 `native_text`、`native_shape`、`svg_object` 或 `raster_picture`。
3. 按 `references/scene-format.md` 写出 UTF-8 `scene-plan.json`。所有坐标统一使用参考图像素坐标。
4. 调用 `ppt_prepare_image_scene`，修复所有明确的 schema、越界、重复 ID、图片来源和分类错误。
5. `semantic-preview.svg` 只是结构交换产物，不是最终质量依据。不得使用 Grep、文本搜索或阅读 SVG XML 来代替视觉检查。
6. WorkBuddy 无法直接视觉查看 SVG 时，不要停留在 SVG 源码；立即调用 `ppt_build_editable_slide`，以 PowerPoint 原生导出的 PNG 作为视觉检查对象。
7. 调用 `ppt_compare_slide`。必须实际查看以下图像，而不只是读取 JSON 数字：
   - PowerPoint 原生渲染 PNG；
   - `aligned-reference.png`；
   - `overlay.png`；
   - `difference-heatmap.png`。
8. 调用 `ppt_inspect_powerpoint` 检查原生对象、文字、字体、越界、裁切、换行、重叠和 z-order。
9. 若未通过质量门槛，按“全局几何 → 图片裁切 → 文字内容与换行 → 字体字号 → 局部装饰”的顺序写出显式 correction plan，并调用 `ppt_apply_correction_plan`。
10. 每次校正后必须再次调用 `ppt_compare_slide`，并重新查看渲染图、overlay 和 heatmap。
11. 达到 `PASS`，或连续两轮没有有意义的改进而进入 `PLATEAU` 后，才允许结束。

## 对象政策

- 标题、正文、标签、数字和说明文字必须优先使用 PowerPoint 原生文本框。
- 矩形、圆形、卡片、色块、直线、箭头和简单几何必须优先使用 PowerPoint 原生形状或连接线。
- 简单矢量图标可作为一个 SVG 图形对象保留，但不得声称其内部路径等同于 PowerPoint 原生形状。
- 照片、复杂插画、艺术字、纹理和细节密集区域作为独立、可移动的受保护图片保留。
- 严禁把整页参考图作为可见的全页背景或嵌入 SVG 来虚假提高相似度。

## 结束状态

### PASS

只有 `references/quality-gate.md` 中全部硬性条件通过，且 `comparison.json` 明确确认视觉门槛后，才可声称“达到 SlideTwin 质量门槛”。

### PLATEAU

未达到全部门槛，但连续两轮校正没有有意义的改进时可以停止。必须明确说明未通过正式验收，并报告实际指标、失败项目、差异区域及原因。

### FAIL

环境、构建、渲染、比较、文件完整性或反扁平化要求失败时返回 FAIL，不交付伪完成结果。

## 最终输出

至少返回：

- 最终可编辑 PPTX；
- PowerPoint 原生渲染 PNG；
- `scene-plan.resolved.json`；
- `semantic-preview.svg`；
- 受保护图片素材目录；
- 最终 `comparison.json`；
- `overlay.png` 与 `difference-heatmap.png`；
- PowerPoint 结构检查报告；
- correction plans；
- PASS / PLATEAU / FAIL 状态；
- 视觉相似度、文字准确率、受保护图片完整性和可编辑覆盖率。

未经比较报告确认，不得声称达到 90%。