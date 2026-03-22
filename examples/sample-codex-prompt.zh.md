# 示例 Codex Prompt

下面是一段可以直接复制给 Codex 的中文示例 prompt。

```text
请使用 $chem-pptx-codex 帮我制作一份 8 到 10 页的化工工艺设计 PPT。

主题：年产 10 万吨 Polycarbonate (PC) 工艺设计
受众：化工学院答辩委员会与企业工程评审老师

目标：
说明为什么选择非光气熔融酯交换路线、关键的 Mass & Energy Balance 假设、主要设备选型依据，以及 HSE 与 Techno-Economic Analysis 的核心结论。

风格要求：
- 主语言为简体中文
- 保留 Polycarbonate (PC)、Aspen Plus、Mass & Energy Balance、HSE、Techno-Economic Analysis 等英文术语
- 中英文混排时自动保留半角空格
- 每页都要有专业图示占位，不要做纯文字页

请按以下方式完成：
1. 先按“化工工艺设计”场景给出页结构建议。
2. 再给出配色和字体方向。
3. 如果没有模板，就用 PptxGenJS 从零创建 starter deck。
4. 每页补上具体的化工图示占位符。
5. 生成后做内容 QA 和视觉 QA。
6. 最后说明哪些地方还需要我补充 Aspen 流程图、设备参数或经济性数据。
```
