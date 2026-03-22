# 编辑演示文稿

## 基于模板的化工工作流

当用户提供实验室模板、学院答辩模板、工厂汇报模板或历史工艺包 deck 时，优先走模板编辑路径。

1. 先分析现有模板：

```bash
python scripts/thumbnail.py template.pptx
python scripts/extract_text.py template.pptx
```

如本地安装了 `markitdown[pptx]`，也可额外运行：

```bash
python -m markitdown template.pptx
```

2. 建立“内容块 -> 模板页”的映射，不要急着直接替换文案。

化工类页面优先映射为：

- 背景与意义页
- 分子结构 / 机理示意页
- 表征结果页
- 流程路线比选页
- Mass & Energy Balance 表格页
- 关键设备选型页
- HSE / 风险矩阵页
- 结论与展望页

3. 解包：

```bash
python scripts/office/unpack.py template.pptx unpacked/
```

4. 先做结构调整：

- 删除不适用的商务页或空白页
- 复制适合承载图谱、流程图、表格的 slide
- 把“标题 + bullet”页拆成图文页、表格页、机理页或流程页

5. 再替换内容：

- 全部文本先过一遍中英混排和单位规范
- 把化学式统一为可显示的 Unicode 形式
- 每页至少保留一个领域图示占位符
- 如果模板原有 4 张图框而实际只有 3 张，不要只清空最后一张，直接删除整组无用元素

6. 清理：

```bash
python scripts/clean.py unpacked/
```

7. 重打包：

```bash
python scripts/office/pack.py unpacked/ output.pptx
```

8. 执行 QA。

## 常用脚本

| 脚本 | 作用 |
| --- | --- |
| `scripts/office/unpack.py` | 解包 PPTX |
| `scripts/add_slide.py` | 复制 slide |
| `scripts/clean.py` | 清理孤儿 slide 和关系 |
| `scripts/office/pack.py` | 重打包 PPTX |
| `scripts/thumbnail.py` | 生成缩略总览 |
| `scripts/office/validate.py` | 校验解包目录结构 |
| `scripts/chem_presentation_logic.py` | 生成化工大纲、占位符和文本预处理结果 |

## 常见坑

- 将 `Aspen Plus`、`GC-MS`、`HSE` 等术语错误翻译成不自然中文
- 中文和英文缩写贴在一起，没有空格
- `H2O`、`SO4^2-`、`13C NMR` 直接原样进入成稿，导致视觉不专业
- 表征图、PFD 或 P&ID 页只有文字说明，没有图示占位
- 单页塞进过多机理步骤或设备参数，导致字号过小
