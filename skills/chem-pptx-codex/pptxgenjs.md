# PptxGenJS 使用指南

## 何时使用

当没有合适模板，或者你需要完全控制“化工工艺图页、谱图页、设备页、HSE 页”的版式时，使用 PptxGenJS 从零生成。

官方文档：

- [PptxGenJS 首页](https://gitbrent.github.io/PptxGenJS/)
- [Introduction](https://gitbrent.github.io/PptxGenJS/docs/introduction/)

## 先做这 4 步

1. 判定场景：`academic` / `process_design` / `safety_training`
2. 用 `chem_presentation_logic.py` 生成默认大纲
3. 对标题、正文、单位、化学式执行文本预处理
4. 为每页分配图示占位符，再落版

## 文本填充前的规则

### 中英混排

中文与英文缩写、软件名、数字之间保留半角空格，例如：

- `通过 GC-MS 进行了产物分析`
- `Aspen Plus 模拟得到 18.4 t/h 的物料流率`
- `反应温度控制在 240 °C`

### 化学式

PptxGenJS 本身更适合直接写入 Unicode 结果，因此推荐先把：

- `H2O` 处理为 `H₂O`
- `SO4^2-` 处理为 `SO₄²⁻`
- `13C NMR` 处理为 `¹³C NMR`

### 占位符

化工页面不要只放文字。优先插入或保留：

- ChemDraw 分子结构图
- Reaction Mechanism 示意图
- NMR / IR / MS / XRD 图谱
- Aspen PFD / P&ID 流程图
- 设备剖面图或选型表
- 风险矩阵、事故树、应急流程图

## 推荐组件

- `addSectionHeader(slide, title, subtitle)`
- `addProcessCard(slide, data)`
- `addSpectrumPlaceholder(slide, label)`
- `addFlowDiagramPlaceholder(slide, label)`
- `addMetricTable(slide, rows)`

## 常见坑

- 颜色值不要带 `#`
- 不要把机理步骤、条件、结论塞成一段大段文字
- 不要忽略中文和英文之间的空格
- 不要把 `m3/h`、`kJ / mol`、`wt %` 原样放进成稿
- 不要只写“插图”这种泛化占位，应该写成具体的化工图示占位

## 示例

完整示例见：

- [`../../examples/starter-pptxgenjs.js`](../../examples/starter-pptxgenjs.js)
- [`../../examples/python-pptx-chemical-formatting.py`](../../examples/python-pptx-chemical-formatting.py)
