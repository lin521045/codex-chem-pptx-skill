# 使用说明

## 安装

```bash
npx skills add https://github.com/lin521045/codex-chem-pptx-skill --skill chem-pptx-codex
```

## 典型提示词

- 用 `chem-pptx-codex` 根据我的实验手稿生成一套中文答辩 PPT，保留 ¹H NMR、GC-MS、DFT 等英文术语
- 帮我做一份“年产 10 万吨 Polycarbonate (PC) 工艺设计”PPT，默认采用化工工艺设计大纲
- 把这个 `.pptx` 模板改成 HSE 安全培训版本，并补上 MSDS、Emergency Response、SOP 页面

## 建议工作流

1. 先判断场景属于学术研究、工艺设计还是安全培训。
2. 用 `chem_presentation_logic.py` 规划默认大纲和图示占位符。
3. 如果要做高端答辩版式，优先走 `python-pptx-xmu-layout.py` 这类自定义布局，而不是默认模板。
4. 再走模板编辑或 PptxGenJS 从零生成流程。
5. 最后执行文本 QA、占位符 QA 和视觉 QA。
