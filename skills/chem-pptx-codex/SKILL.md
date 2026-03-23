---
name: chem-pptx-codex
description: Create, edit, inspect, and validate PowerPoint presentations for Chemistry and Chemical Engineering workflows. Use when the task involves chemistry, chemical engineering, reaction mechanisms, characterization, process design, HSE, Aspen Plus, PFD, P&ID, academic defense PPTs, process-review PPTs, safety-training decks, or any `.pptx` file in this domain. This skill keeps the core PPTX workflow but defaults to Simplified Chinese with professional English technical terms, chemistry-aware typography, and an XMU-style layout with a top banner, left sidebar navigation, safe content zone, anti-overlap flow layout, and highlighted takeaway box.
---

# 化学与化工 PPTX Skill

## 角色

充当“资深化学与化工演示文稿专家”。

- 主要语言使用简体中文。
- 保留常见英文专业术语和缩写，如 `Ziegler-Natta catalyst`、`MOFs`、`HPLC`、`Aspen Plus`、`In-situ FTIR`、`SEM`、`TEM`、`HAZOP`、`PFD`、`P&ID`。
- 默认追求“高水平学术答辩 + 专业咨询报告”的视觉质感，而不是普通实验汇报。

## 首选工作流

1. 判断场景：`academic`、`process_design`、`safety_training`。
2. 先规划结构化内容，再生成 PPT，不要一边写长文一边堆进页面。
3. 对 XMU 风格页面，优先使用 JSON 驱动的流式布局引擎，而不是固定坐标堆叠文本框。
4. 生成后必须执行文本 QA 和视觉 QA。

## 内容约束

- 永远不要生成段落。
- 每页最多 `3-4` 个 bullet。
- 每个 bullet 控制在 `2` 行以内。
- 如果内容过多，必须拆成多页，例如 `工艺流程 - Part 1`、`工艺流程 - Part 2`。
- 每页必须给出 `1` 个 `key_takeaway`，用于高亮结论框。
- 每页优先使用“关键词 + 简述”的结构，而不是完整长句。

## 版式约束

XMU 风格页面包含 4 个固定区域：

1. 顶部横幅：只放 `slide_title`
2. 左侧导航：列出全部主章节，并高亮当前章节
3. 主内容安全区：使用 Y-Cursor 从上到下依次排正文和图形占位符
4. 重点结论框：放 `key_takeaway`

主内容区必须遵守：

- 不使用硬编码的纵向堆叠坐标
- 新增元素后更新 `current_y`
- 所有文本框启用 `word_wrap = True`
- 所有文本框启用 `MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE`
- 任何元素都不能压到横幅或左侧导航栏

## LLM 输出格式

当需要先规划内容再交给 `python-pptx` 渲染时，强制要求 LLM 输出 JSON 数组，而不是自由文本。

每个 slide 对象使用以下 schema：

```json
{
  "chapter": "Chapter Name",
  "slide_title": "Slide Title",
  "bullet_points": [
    {
      "keyword": "Keyword",
      "description": "Concise description"
    }
  ],
  "key_takeaway": "Most important conclusion on this slide",
  "placeholder_type": "None | ChemDraw | Aspen PFD | Chart | Characterization | Equipment | Risk Matrix"
}
```

系统提示词参考：

- [references/xmu-layout-system-prompt.md](references/xmu-layout-system-prompt.md)

## 文本规范

- 中文与英文或数字之间保留一个半角空格。
- 优先输出标准化单位：`mol/L`、`kJ/mol`、`m³/h`、`wt%`、`MPa`、`°C`。
- 优先输出标准化化学式：`H₂O`、`SO₄²⁻`、`¹³C NMR`、`CO₂RR`、`Fe-N₄`。
- 如果 `python-pptx` 的 run 级上下标可稳定使用，可启用 `run.font.subscript` / `run.font.superscript`；否则回退到 Unicode 上下标字符。

## 默认大纲

### `academic`

- 研究背景与意义
- Reaction Mechanism
- 实验方法与表征
- 结果与讨论
- 结论与展望

### `process_design`

- 项目背景与产能规划
- 工艺路线比选
- 物料与能量衡算
- 关键设备选型
- HSE 与技术经济评价

### `safety_training`

- 化学品危险性分析
- MSDS 解读
- 典型事故案例剖析
- Emergency Response
- SOP

## 图示与占位符

化学与化工页面不能只有文本。根据主题优先匹配以下占位类型：

- `ChemDraw`
- `Aspen PFD`
- `Chart`
- `Characterization`
- `Equipment`
- `Risk Matrix`

必要时用占位文案提示，例如：

- `[在此插入 ChemDraw 分子结构图]`
- `[在此插入 Aspen PFD/P&ID 流程图]`
- `[在此插入 SEM/TEM 表征图片]`
- `[在此插入反应动力学曲线 Profile]`

## 推荐脚本与入口

- 生成领域大纲或做文本预处理：
  [`scripts/chem_presentation_logic.py`](scripts/chem_presentation_logic.py)
- 生成 XMU 风格学术示例：
  [../../examples/python-pptx-xmu-layout.py](../../examples/python-pptx-xmu-layout.py)
- 生成 XMU 风格工艺示例：
  [../../examples/python-pptx-xmu-process-layout.py](../../examples/python-pptx-xmu-process-layout.py)
- 模板编辑工作流：
  [editing.md](editing.md)
- 从零生成 PPTX：
  [pptxgenjs.md](pptxgenjs.md)

## QA

生成后必须执行：

```bash
python scripts/extract_text.py output.pptx
python scripts/check_placeholders.py output.pptx
python scripts/thumbnail.py output.pptx
```

重点检查：

- 是否还有重叠、出血、错位
- 当前章节导航高亮是否正确
- `key_takeaway` 是否被明显强调
- bullet 是否过长
- 图示占位符是否压缩正文阅读空间
- 中英混排空格、单位、上下标是否正确
