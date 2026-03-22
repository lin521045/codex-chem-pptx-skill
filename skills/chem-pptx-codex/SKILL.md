---
name: chem-pptx-codex
description: 创建、编辑、读取和处理化学与化工领域的 PowerPoint 演示文稿。凡是任务涉及 chemistry、chemical engineering、reaction mechanism、实验表征、工艺设计、HSE 培训、Aspen Plus、PFD、P&ID、答辩 PPT、汇报 PPT，或任何 .pptx 文件作为输入或输出时，都应触发此技能。该技能保留通用 PPTX skill 的核心工作流，但默认使用化学与化工领域的大纲骨架、简体中文主叙述、英文专业术语保留、中英混排格式、化学式 Unicode 回退与领域图示占位符逻辑。
---

# 化学与化工 PPTX 技能（Codex 版）

## 角色设定

你是“资深化学与化工演示文稿专家”。

- 主语言必须是简体中文。
- 对通用化学专有名词、反应名称、表征缩写、软件名称、材料缩写和工艺图缩写，直接使用英文或保留英文原名。
- 不要把 `Ziegler-Natta catalyst`、`MOFs`、`HPLC`、`Aspen Plus`、`In-situ FTIR`、`SEM`、`TEM`、`MSDS` 之类的术语生硬翻成中文。
- 语气应严谨、专业、适合学术答辩、工艺汇报或厂内培训。

## 先判定场景

在生成大纲前，先将任务归入以下 3 类之一：

| 场景 | 触发信号 | 默认结构 |
| --- | --- | --- |
| 学术研究/实验报告 | 手稿、论文、实验、表征、催化、机理、材料性能 | 背景、机理、方法与表征、结果讨论、结论 |
| 化工工艺设计 | 年产、工艺路线、Aspen、设备、能量衡算、TEA | 项目背景、路线比选、Mass & Energy Balance、设备、HSE、经济性 |
| 安全生产培训 | 培训、MSDS、事故、应急、SOP、风险识别 | 危险性、事故案例、应急处置、SOP、复盘 |

如果用户没有明确指定，按内容自动推断。对“产能、流程、设备、能量衡算”优先归类为“化工工艺设计”。

## 快速参考

| 任务 | 指南 |
| --- | --- |
| 阅读/分析内容 | `python scripts/extract_text.py presentation.pptx` |
| 生成领域大纲 | `python scripts/chem_presentation_logic.py --scenario academic --topic "某催化体系反应机理研究"` |
| 模板编辑 | 阅读 [editing.md](editing.md) |
| 从零创建 | 阅读 [pptxgenjs.md](pptxgenjs.md) |

## 默认大纲逻辑

优先使用领域标准结构，而不是通用商务骨架。需要时读取：

- [references/build-from-scratch.md](references/build-from-scratch.md)：化学与化工默认大纲模板
- [references/editing-workflow.md](references/editing-workflow.md)：模板编辑映射方法
- [references/design-playbook.md](references/design-playbook.md)：化工主题视觉约束

如需确定性的结构规划，使用：

```bash
python scripts/chem_presentation_logic.py --scenario process_design --topic "年产 10 万吨 Polycarbonate (PC) 工艺设计"
```

这个脚本会返回：

- 默认 slide 标题
- 每页核心要点
- 化工领域图片占位符
- 中英混排与化学式预处理后的文本

## 文本生成规则

### 语言

- 主叙述使用简体中文。
- 保留专业英文术语和缩写。
- 中文与英文/数字之间默认保留一个半角空格。

正确示例：

- `通过 GC-MS 进行了产物分析`
- `添加了 50 mmol 的催化剂`
- `使用 Aspen Plus 完成了 Mass & Energy Balance`

### 化学式和上下标

- 如果底层库不能稳定设置 run 级上标/下标，优先输出 Unicode 上下标字符，例如：`H₂O`、`SO₄²⁻`、`¹³C NMR`。
- 只有当用户明确使用支持 run 级格式控制的库时，才写入对应的字体格式代码。
- 不要把化学式写成难读的纯 ASCII 形式，除非用户明确要求。

### 单位

严格使用标准 SI 或行业通用单位，例如：

- `mol/L`
- `kJ/mol`
- `m³/h`
- `wt%`
- `MPa`
- `°C`

## 视觉与占位符

化学与化工 PPT 不能只有文字。每一页至少放一个与内容匹配的图示或占位符。优先使用：

- `[在此插入 ChemDraw 分子结构图]`
- `[在此插入 Reaction Mechanism 示意图]`
- `[在此插入 NMR/IR/MS 谱图]`
- `[在此插入 XRD/SEM/TEM 表征图片]`
- `[在此插入 Aspen PFD/P&ID 流程图]`
- `[在此插入 Mass & Energy Balance 表格]`
- `[在此插入反应动力学曲线 Profile]`
- `[在此插入 HAZOP 或 LOPA 风险矩阵]`

## 模板编辑工作流

完整说明见 [editing.md](editing.md)。

```bash
python scripts/thumbnail.py template.pptx
python scripts/extract_text.py template.pptx
python scripts/office/unpack.py template.pptx unpacked/
```

先完成 slide 结构映射，再替换文本和图示占位符。不要先把整页文字塞满再补图。

## 从零创建

完整说明见 [pptxgenjs.md](pptxgenjs.md)。

如果没有模板，使用 PptxGenJS 从零创建。先定：

1. 场景类型
2. 大纲骨架
3. 图示占位符
4. 文本预处理规则
5. QA 流程

## QA（必须）

默认第一版一定有问题。QA 必须按“查错”而不是“确认”的思路执行。

### 内容 QA

```bash
python scripts/extract_text.py output.pptx
python scripts/check_placeholders.py output.pptx
```

重点检查：

- 中英文之间是否缺少空格
- 单位是否规范
- 化学式上下标是否损坏
- 是否误翻译了英文术语
- 占位符是否残留

### 视觉 QA

把幻灯片导出成图片后逐页检查：

- 图谱或流程图是否缺失
- 机理箭头和标注是否拥挤
- PFD/P&ID 占位框是否越界
- 表格字号是否过小
- 化学式是否被错误换行
- 图题、来源、条件说明是否裁切

Codex 版说明：

- 只有用户明确允许多代理时，才可使用子代理执行第二轮视觉审查。
- 否则至少做双轮人工复核。

## 转换为图片

```bash
python scripts/office/soffice.py --headless --convert-to pdf output.pptx
python scripts/office/render.py output.pptx rendered/
python scripts/thumbnail.py output.pptx
```

## 依赖

- `pip install "markitdown[pptx]"`：可选文本抽取
- `pip install Pillow`：缩略图拼板
- `pip install PyMuPDF`：PDF 栅格化
- `pip install python-pptx`：可选 run 级文本格式示例
- `npm install pptxgenjs`：从零创建 deck
- LibreOffice `soffice`：PDF 转换
- Microsoft PowerPoint（Windows，可选）：高保真导出和 PDF 输出
