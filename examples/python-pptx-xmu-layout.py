#!/usr/bin/env python3
"""Generate an academic XMU-style chemistry deck with the flow layout engine."""

from __future__ import annotations

import json
from pathlib import Path

from xmu_flow_layout_engine import build_presentation, load_slide_specs


ACADEMIC_SAMPLE_JSON = json.dumps(
    [
        {
            "chapter": "研究背景与意义",
            "slide_title": "MOF 衍生 Fe-N₄ 催化位的研究背景",
            "bullet_points": [
                {"keyword": "研究动机", "description": "CO₂RR 为 C₁ 化学品低碳转化提供新路径"},
                {"keyword": "现有瓶颈", "description": "非贵金属体系常受选择性与稳定性限制"},
                {"keyword": "本课题切入", "description": "聚焦 MOF 衍生 Fe-N₄ 位点的结构-性能关系"},
            ],
            "key_takeaway": "高分散 Fe-N₄ 位点是同时兼顾活性与选择性的关键切入点。",
            "placeholder_type": "Chart",
        },
        {
            "chapter": "Reaction Mechanism",
            "slide_title": "Reaction Mechanism 假设与验证思路",
            "bullet_points": [
                {"keyword": "活性位假设", "description": "Fe-N₄ 位点优先稳定 *COOH 中间体"},
                {"keyword": "验证手段", "description": "结合 In-situ FTIR 与 DFT 识别速率控制步骤"},
                {"keyword": "对照策略", "description": "通过 Fe 团簇样与空白样排除非目标位点贡献"},
            ],
            "key_takeaway": "机制页只回答一个问题：为什么 Fe-N₄ 会更优。",
            "placeholder_type": "ChemDraw",
        },
        {
            "chapter": "实验方法与表征",
            "slide_title": "实验方法与表征链条",
            "bullet_points": [
                {"keyword": "样品制备", "description": "采用 ZIF-8 限域热解构建 Fe-N₄ 分散位点"},
                {"keyword": "表征组合", "description": "用 XRD、XPS、HAADF-STEM、XANES 交叉验证"},
                {"keyword": "数据质量", "description": "统一重复次数、误差条和碳平衡口径"},
            ],
            "key_takeaway": "表征设计必须服务于位点确认，而不是简单堆谱图。",
            "placeholder_type": "Characterization",
        },
        {
            "chapter": "结果与讨论",
            "slide_title": "结果与讨论：性能与证据闭环",
            "bullet_points": [
                {"keyword": "性能表现", "description": "在 -0.72 V vs. RHE 下 FE_CO 达到 94%，j_CO 同步提升"},
                {"keyword": "结构证据", "description": "HAADF-STEM 未见明显 Fe 团簇，支持单原子分散"},
                {"keyword": "机理证据", "description": "In-situ FTIR 捕获 *COOH 与 *CO 的演化信号"},
            ],
            "key_takeaway": "最强说服力来自“性能指标 + 位点证据 + 机理证据”的同页闭环。",
            "placeholder_type": "Chart",
        },
        {
            "chapter": "结论与展望",
            "slide_title": "结论与展望",
            "bullet_points": [
                {"keyword": "核心结论", "description": "MOF 衍生策略成功构建高分散 Fe-N₄ 活性位"},
                {"keyword": "方法意义", "description": "建立了位点结构与 CO₂RR 选择性间的定量联系"},
                {"keyword": "下一步", "description": "延伸到 MEA 条件并结合 operando XAS 深化验证"},
            ],
            "key_takeaway": "结论页要明确回答做成了什么、为什么可信、下一步做什么。",
            "placeholder_type": "None",
        },
    ],
    ensure_ascii=False,
    indent=2,
)


def main() -> int:
    slides = load_slide_specs(ACADEMIC_SAMPLE_JSON)
    output = Path(__file__).resolve().parents[1] / "xmu-chem-academic-defense-sample.pptx"
    build_presentation(slides, output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
