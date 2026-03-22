# chem-pptx-codex

[![CI](https://github.com/lin521045/codex-chem-pptx-skill/actions/workflows/validate.yml/badge.svg)](https://github.com/lin521045/codex-chem-pptx-skill/actions/workflows/validate.yml)
[![Docs](https://img.shields.io/badge/docs-GitHub%20Pages-0A66C2)](https://lin521045.github.io/codex-chem-pptx-skill/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

```bash
npx skills add https://github.com/lin521045/codex-chem-pptx-skill --skill chem-pptx-codex
```

## 摘要

这是一个专门面向“化学与化工”场景的 Codex PPT skill。它保留通用 `pptx` skill 的核心架构和工具链，但把 System Prompt、默认大纲、文本预处理和视觉占位逻辑全部重构为化学与化工领域版本。

- 角色设定为“资深化学与化工演示文稿专家”，主语言为简体中文，并自然混用专业英文术语、缩写与专有名词
- 内置 3 套默认大纲骨架：学术研究/实验报告、化工工艺设计、安全生产培训
- 强制执行中英混排空格、SI 单位规范、化学式 Unicode 上下标回退和领域图示占位符策略
- 保留 PPTX 读取、解包、重打包、渲染、缩略图、占位符检查等原始工作流

## Skill 入口

- 技能目录：[`skills/chem-pptx-codex`](skills/chem-pptx-codex)
- 主提示文件：[`skills/chem-pptx-codex/SKILL.md`](skills/chem-pptx-codex/SKILL.md)
- 逻辑辅助脚本：[`skills/chem-pptx-codex/scripts/chem_presentation_logic.py`](skills/chem-pptx-codex/scripts/chem_presentation_logic.py)
- 高端 `python-pptx` 布局示例：[`examples/python-pptx-xmu-layout.py`](examples/python-pptx-xmu-layout.py)

## 快速参考

| 任务 | 指南 |
| --- | --- |
| 阅读/分析现有 `.pptx` | `python skills/chem-pptx-codex/scripts/extract_text.py presentation.pptx` |
| 规划化工领域大纲 | `python skills/chem-pptx-codex/scripts/chem_presentation_logic.py --scenario process_design --topic "年产 10 万吨 Polycarbonate (PC) 工艺设计"` |
| 编辑模板 deck | 阅读 [`skills/chem-pptx-codex/editing.md`](skills/chem-pptx-codex/editing.md) |
| 从零创建化工 deck | 阅读 [`skills/chem-pptx-codex/pptxgenjs.md`](skills/chem-pptx-codex/pptxgenjs.md) |

## 这个仓库改了什么

### 1. Role / System Prompt

Skill 明确要求 Codex 扮演“资深化学与化工演示文稿专家”，默认输出简体中文，但下列内容优先保留英文：

- 通用化学术语与缩写：MOFs、GC-MS、HPLC、BET、XRD、DFT、TEA
- 反应或催化体系名称：Ziegler-Natta catalyst、Fischer-Tropsch synthesis
- 软件和流程图名称：Aspen Plus、Aspen HYSYS、PFD、P&ID
- 表征和原位测试：In-situ FTIR、¹³C NMR、SEM、TEM

### 2. Domain Outline Templates

默认大纲骨架覆盖：

- 学术研究/实验报告：研究背景与意义、Reaction Mechanism、实验方法与表征、结果与讨论、结论与展望
- 化工工艺设计：项目背景与产能规划、工艺路线比选、Mass & Energy Balance、关键设备选型、HSE、Techno-Economic Analysis
- 安全生产培训：化学品危险性分析、MSDS 解读、事故案例剖析、Emergency Response、SOP

详细骨架见 [`skills/chem-pptx-codex/references/build-from-scratch.md`](skills/chem-pptx-codex/references/build-from-scratch.md)。

### 3. 文本预处理

在向幻灯片填充文本前，优先执行：

- 中英文和阿拉伯数字之间自动插入半角空格
- `mol / L`、`kJ / mol`、`m3/h`、`wt %` 等常见表达统一到标准格式
- 对化学式、同位素和带电离子尽量输出 Unicode 上下标
- 保留英文专有名词，不做生硬直译

预处理脚本见 [`skills/chem-pptx-codex/scripts/chem_presentation_logic.py`](skills/chem-pptx-codex/scripts/chem_presentation_logic.py)。

### 4. 视觉占位符

每页幻灯片默认至少有一个化工领域专用图示占位符，例如：

- `[在此插入 ChemDraw 分子结构图]`
- `[在此插入 Aspen PFD/P&ID 流程图]`
- `[在此插入 SEM/TEM 表征图片]`
- `[在此插入反应动力学曲线 Profile]`

### 5. 高级版式系统

新增了一套 `python-pptx` 的高端科研答辩布局：

- 顶部横幅：统一使用厦门大学蓝 `#003F88`
- 左侧导航栏：动态遍历整套章节并高亮当前章
- 内容安全区：正文与图示严格限制在右侧 Safe Zone
- 字体与正文颜色：白色横幅标题 + 深灰正文，风格更接近高端咨询报告和正式学术答辩

相关文件：

- [`examples/python-pptx-xmu-layout.py`](examples/python-pptx-xmu-layout.py)
- [`skills/chem-pptx-codex/references/xmu-layout-system-prompt.md`](skills/chem-pptx-codex/references/xmu-layout-system-prompt.md)

## 示例

- PptxGenJS 示例：[`examples/starter-pptxgenjs.js`](examples/starter-pptxgenjs.js)
- `python-pptx` 运行级字体示例：[`examples/python-pptx-chemical-formatting.py`](examples/python-pptx-chemical-formatting.py)
- `python-pptx` 厦大蓝布局示例：[`examples/python-pptx-xmu-layout.py`](examples/python-pptx-xmu-layout.py)
- `python-pptx` 厦大蓝工艺设计示例：[`examples/python-pptx-xmu-process-layout.py`](examples/python-pptx-xmu-process-layout.py)

`starter-pptxgenjs.js` 演示了如何生成一份“年产 10 万吨 Polycarbonate (PC) 工艺设计”主题的中英混排演示文稿。

## 依赖

- `pip install "markitdown[pptx]"`：可选文本抽取
- `pip install Pillow`：缩略图拼板
- `pip install PyMuPDF`：PDF 栅格化
- `pip install python-pptx`：可选 Python 侧文本格式控制示例
- `npm install pptxgenjs`：从零创建 deck
- LibreOffice `soffice`：PDF 转换
- Microsoft PowerPoint（Windows，可选）：高保真渲染和导出

## 仓库入口

- 仓库：[github.com/lin521045/codex-chem-pptx-skill](https://github.com/lin521045/codex-chem-pptx-skill)
- 文档页：[lin521045.github.io/codex-chem-pptx-skill](https://lin521045.github.io/codex-chem-pptx-skill/)
