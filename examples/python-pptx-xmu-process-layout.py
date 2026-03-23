#!/usr/bin/env python3
"""Generate a process-design XMU-style chemistry deck with the flow layout engine."""

from __future__ import annotations

import json
from pathlib import Path

from xmu_flow_layout_engine import build_presentation, load_slide_specs


PROCESS_SAMPLE_JSON = json.dumps(
    [
        {
            "chapter": "项目背景与产能规划",
            "slide_title": "年产 10 万吨 Polycarbonate (PC) 项目背景",
            "bullet_points": [
                {"keyword": "建设目标", "description": "面向光学级与工程塑料级 Polycarbonate (PC) 市场"},
                {"keyword": "产能基准", "description": "设计产能 10 万吨/年，开工时数按 8000 h/a 计"},
                {"keyword": "边界条件", "description": "与上游 DPC 供给及 Phenol 回收系统协同设计"},
            ],
            "key_takeaway": "第一页必须先锁定产品定位、产能边界和系统协同关系。",
            "placeholder_type": "Chart",
        },
        {
            "chapter": "工艺路线比选",
            "slide_title": "工艺路线比选与推荐方案",
            "bullet_points": [
                {"keyword": "候选路线", "description": "对比光气法与非光气熔融酯交换法的成熟度与风险"},
                {"keyword": "筛选逻辑", "description": "重点比较 HSE 压力、能耗水平和副产物处理难度"},
                {"keyword": "推荐结论", "description": "选择 DPC 与 BPA 熔融酯交换路线作为基准方案"},
            ],
            "key_takeaway": "路线页的价值在于解释“为什么选它”，而不是罗列路线名称。",
            "placeholder_type": "Aspen PFD",
        },
        {
            "chapter": "物料与能量衡算",
            "slide_title": "物料衡算 (Mass Balance) 与能量衡算 (Energy Balance)",
            "bullet_points": [
                {"keyword": "主流股", "description": "PC 产品约 12.5 t/h，BPA 与 DPC 进料构成主负荷"},
                {"keyword": "回收流", "description": "Phenol 回收流直接影响塔系负荷与公用工程需求"},
                {"keyword": "热负荷", "description": "反应段、脱挥段和精馏段是主要用热集中区"},
            ],
            "key_takeaway": "衡算页只展示影响设备尺寸和公用工程的关键数字。",
            "placeholder_type": "Aspen PFD",
        },
        {
            "chapter": "关键设备选型",
            "slide_title": "关键设备选型",
            "bullet_points": [
                {"keyword": "反应器", "description": "优先保证高黏熔体体系的传热与停留时间分布"},
                {"keyword": "分离设备", "description": "脱挥器与塔器需兼顾真空稳定性和 Phenol 腐蚀环境"},
                {"keyword": "连续运行", "description": "关键转动设备采用 1 用 1 备提升装置可靠性"},
            ],
            "key_takeaway": "设备选型必须体现工艺约束、材质约束和连续运行策略。",
            "placeholder_type": "Equipment",
        },
        {
            "chapter": "HSE 与技术经济评价",
            "slide_title": "HSE 与技术经济评价 (Techno-Economic Analysis)",
            "bullet_points": [
                {"keyword": "风险识别", "description": "聚焦高温熔体泄漏、真空失效与 Phenol 暴露场景"},
                {"keyword": "经济抓手", "description": "Phenol 回收率与热集成效率是 OPEX 敏感项"},
                {"keyword": "综合判断", "description": "在优化回收与公用工程后具备继续深化设计基础"},
            ],
            "key_takeaway": "没有 HSE 和 TEA 闭环，工艺设计页就缺少决策价值。",
            "placeholder_type": "Risk Matrix",
        },
    ],
    ensure_ascii=False,
    indent=2,
)


def main() -> int:
    slides = load_slide_specs(PROCESS_SAMPLE_JSON)
    output = Path(__file__).resolve().parents[1] / "xmu-chem-process-design-sample.pptx"
    build_presentation(slides, output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
