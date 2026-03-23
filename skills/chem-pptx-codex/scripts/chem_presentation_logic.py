#!/usr/bin/env python3
"""Generate chemistry presentation plans with JSON slide schemas."""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import asdict, dataclass


SUBSCRIPT_MAP = str.maketrans("0123456789+-()", "₀₁₂₃₄₅₆₇₈₉₊₋₍₎")
SUPERSCRIPT_MAP = str.maketrans("0123456789+-()", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁽⁾")
CJK = r"\u4e00-\u9fff"
ASCII_BLOCK = r"A-Za-z0-9"
FORMULA_RE = re.compile(r"(?<![A-Za-z])((?:\d+)?(?:[A-Z][a-z]?\d*)+(?:\^[0-9]*[+-]|[+-])?)(?![A-Za-z])")
ISOTOPE_RE = re.compile(r"\b(\d{1,3})([A-Z][a-z]?)(?=\s*(?:NMR|MS|label|标记))")


@dataclass
class BulletPoint:
    keyword: str
    description: str


@dataclass
class SlidePlan:
    chapter: str
    slide_title: str
    bullet_points: list[BulletPoint]
    key_takeaway: str
    placeholder_type: str


def insert_cn_en_spacing(text: str) -> str:
    text = re.sub(fr"([{CJK}])([{ASCII_BLOCK}])", r"\1 \2", text)
    text = re.sub(fr"([{ASCII_BLOCK}])([{CJK}])", r"\1 \2", text)
    return re.sub(r"\s{2,}", " ", text).strip()


def normalize_units(text: str) -> str:
    replacements = {
        "mol / L": "mol/L",
        "kJ / mol": "kJ/mol",
        "kg / h": "kg/h",
        "t / h": "t/h",
        "m3/h": "m³/h",
        "m3 / h": "m³/h",
        "Nm3/h": "Nm³/h",
        "wt %": "wt%",
        "vol %": "vol%",
        "° C": "°C",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    return re.sub(r"(\d)(mmol|mol|g|kg|t|wt%|MPa|kPa|bar|°C|K|h|min|s|t/h|kg/h|m³/h)\b", r"\1 \2", text)


def to_subscript(value: str) -> str:
    return value.translate(SUBSCRIPT_MAP)


def to_superscript(value: str) -> str:
    return value.translate(SUPERSCRIPT_MAP)


def format_formula_token(token: str) -> str:
    isotope = ""
    core = token
    match = re.match(r"^(\d+)([A-Z].*)$", token)
    if match and re.search(r"\d", match.group(2)):
        isotope = to_superscript(match.group(1))
        core = match.group(2)

    charge = ""
    if "^" in core:
        core, charge = core.split("^", 1)
    elif core.endswith(("+", "-")) and re.search(r"\d", core):
        core, charge = core[:-1], core[-1]

    core = re.sub(r"([A-Za-z\)])(\d+)", lambda m: m.group(1) + to_subscript(m.group(2)), core)
    return isotope + core + (to_superscript(charge) if charge else "")


def normalize_chemistry_unicode(text: str) -> str:
    text = ISOTOPE_RE.sub(lambda m: f"{to_superscript(m.group(1))}{m.group(2)}", text)

    def repl(match: re.Match[str]) -> str:
        token = match.group(1)
        if not re.search(r"[\d\^+-]", token):
            return token
        return format_formula_token(token)

    return FORMULA_RE.sub(repl, text)


def normalize_text(text: str) -> str:
    return normalize_chemistry_unicode(normalize_units(insert_cn_en_spacing(text)))


def shorten_text(text: str, max_chars: int) -> str:
    if len(text) <= max_chars:
        return text
    return text[: max_chars - 1].rstrip(" ，,;；:：") + "…"


def compact_bullet(keyword: str, description: str) -> BulletPoint:
    return BulletPoint(
        keyword=shorten_text(normalize_text(keyword), 22),
        description=shorten_text(normalize_text(description), 46),
    )


def choose_placeholder_type(title: str, scenario: str) -> str:
    title_lower = title.lower()
    if "mechanism" in title_lower or "机理" in title:
        return "ChemDraw"
    if any(key in title_lower for key in ("nmr", "ir", "ms", "xrd", "xps", "ftir")) or "表征" in title:
        return "Characterization"
    if any(key in title for key in ("流程", "路线", "衡算")) or "pfd" in title_lower or "p&id" in title_lower:
        return "Aspen PFD"
    if "设备" in title or "equipment" in title_lower:
        return "Equipment"
    if "hse" in title_lower or "安全" in title or "应急" in title:
        return "Risk Matrix"
    if scenario == "academic":
        return "Chart"
    if scenario == "process_design":
        return "Aspen PFD"
    return "Risk Matrix"


def build_outline(topic: str, scenario: str) -> list[SlidePlan]:
    topic = normalize_text(topic)
    templates: dict[str, list[SlidePlan]] = {
        "academic": [
            SlidePlan(
                chapter="研究背景与意义",
                slide_title="研究背景与意义",
                bullet_points=[
                    compact_bullet("研究问题", f"{topic} 对低碳转化或高值利用具有现实意义"),
                    compact_bullet("行业痛点", "现有体系常受活性、选择性或稳定性限制"),
                    compact_bullet("本研究目标", "构建清晰的结构-性能关系并给出机制解释"),
                ],
                key_takeaway="研究价值在于把性能提升与位点机理直接关联。",
                placeholder_type="Chart",
            ),
            SlidePlan(
                chapter="Reaction Mechanism",
                slide_title="Reaction Mechanism 假设",
                bullet_points=[
                    compact_bullet("活性位", "提出关键位点或中间体的形成与转化路径"),
                    compact_bullet("核心假设", "利用 In-situ FTIR 或 DFT 验证关键吸附步骤"),
                    compact_bullet("验证重点", "比较速率控制步骤与目标产物选择性来源"),
                ],
                key_takeaway="机理假设必须能被原位表征和对照实验同时支持。",
                placeholder_type="ChemDraw",
            ),
            SlidePlan(
                chapter="实验方法与表征",
                slide_title="实验方法与表征",
                bullet_points=[
                    compact_bullet("实验设计", "明确反应条件、对照组与重复性策略"),
                    compact_bullet("表征方案", "组合 NMR、IR、MS、XRD、XPS 或 SEM/TEM"),
                    compact_bullet("数据质量", "确保碳平衡、误差条和检测限口径一致"),
                ],
                key_takeaway="表征链条需要与机制问题一一对应，而不是材料堆砌。",
                placeholder_type="Characterization",
            ),
            SlidePlan(
                chapter="结果与讨论",
                slide_title="结果与讨论",
                bullet_points=[
                    compact_bullet("性能指标", "突出收率、选择性、FE 或稳定性窗口"),
                    compact_bullet("证据闭环", "把谱图、显微图和动力学结果合并解释"),
                    compact_bullet("对比结论", "与文献或基准样形成可量化差异"),
                ],
                key_takeaway="结果页只保留最强证据，避免把所有数据压进一页。",
                placeholder_type="Chart",
            ),
            SlidePlan(
                chapter="结论与展望",
                slide_title="结论与展望",
                bullet_points=[
                    compact_bullet("核心结论", "总结位点结构、性能提升与机制证据"),
                    compact_bullet("局限性", "指出样品稳定性或外推场景的边界"),
                    compact_bullet("下一步", "延伸到 operando 表征或放大验证"),
                ],
                key_takeaway="结论页要回答做成了什么、为什么可信、下一步做什么。",
                placeholder_type="None",
            ),
        ],
        "process_design": [
            SlidePlan(
                chapter="项目背景与产能规划",
                slide_title="项目背景与产能规划",
                bullet_points=[
                    compact_bullet("建设目标", f"围绕 {topic} 明确产品定位与设计边界"),
                    compact_bullet("产能假设", "给出年产量、开工时数与负荷率基准"),
                    compact_bullet("边界条件", "明确原料、公用工程和副产物去向"),
                ],
                key_takeaway="第一性问题是产能边界，而不是先堆设备清单。",
                placeholder_type="Chart",
            ),
            SlidePlan(
                chapter="工艺路线比选",
                slide_title="工艺路线比选",
                bullet_points=[
                    compact_bullet("候选路线", "对比原料可得性、成熟度与放大风险"),
                    compact_bullet("决策依据", "从 HSE、能耗和副产物流向筛选方案"),
                    compact_bullet("推荐路线", "给出最终路线及其不选项的关键原因"),
                ],
                key_takeaway="路线比选页必须明确为什么选它，而不是只列方案名称。",
                placeholder_type="Aspen PFD",
            ),
            SlidePlan(
                chapter="物料与能量衡算",
                slide_title="物料与能量衡算",
                bullet_points=[
                    compact_bullet("主流股", "锁定关键进料、回收流和产品流量级"),
                    compact_bullet("热负荷", "标出反应段、分离段和回收段的热需求"),
                    compact_bullet("操作窗口", "给出温压范围与关键控制点"),
                ],
                key_takeaway="衡算页只呈现影响设备和公用工程的关键数字。",
                placeholder_type="Aspen PFD",
            ),
            SlidePlan(
                chapter="关键设备选型",
                slide_title="关键设备选型",
                bullet_points=[
                    compact_bullet("核心设备", "反应器、塔器和换热器的功能定位清晰"),
                    compact_bullet("材质约束", "兼顾腐蚀、黏度和高温操作窗口"),
                    compact_bullet("冗余设计", "对连续运行设备保留检修与备用策略"),
                ],
                key_takeaway="设备选型要体现工艺约束，而不是只给设备名称。",
                placeholder_type="Equipment",
            ),
            SlidePlan(
                chapter="HSE 与技术经济评价",
                slide_title="HSE 与技术经济评价",
                bullet_points=[
                    compact_bullet("风险识别", "聚焦高温、高压、有毒或真空失效场景"),
                    compact_bullet("经济抓手", "区分 CAPEX、OPEX 与回收率敏感项"),
                    compact_bullet("综合结论", "形成是否继续深化设计的判断依据"),
                ],
                key_takeaway="没有 HSE 和 TEA 的闭环，工艺设计就不完整。",
                placeholder_type="Risk Matrix",
            ),
        ],
        "safety_training": [
            SlidePlan(
                chapter="化学品危险性分析",
                slide_title="化学品危险性分析",
                bullet_points=[
                    compact_bullet("危险介质", f"明确 {topic} 相关介质的毒性、腐蚀性和燃爆性"),
                    compact_bullet("暴露途径", "突出吸入、接触与误混引发的后果"),
                    compact_bullet("高风险环节", "锁定储运、取样和切换操作节点"),
                ],
                key_takeaway="危险性分析要落到具体工况，而不是停留在概念层面。",
                placeholder_type="Risk Matrix",
            ),
            SlidePlan(
                chapter="MSDS 解读",
                slide_title="MSDS 解读",
                bullet_points=[
                    compact_bullet("关键参数", "提取闪点、爆炸极限、暴露限值和稳定性"),
                    compact_bullet("防护要求", "对应 PPE、储运条件和禁配物"),
                    compact_bullet("操作提示", "强调现场最容易忽略的限制条件"),
                ],
                key_takeaway="MSDS 页要转化成能执行的现场动作。",
                placeholder_type="None",
            ),
            SlidePlan(
                chapter="典型事故案例剖析",
                slide_title="典型事故案例剖析",
                bullet_points=[
                    compact_bullet("事故经过", "还原关键时间点和失控环节"),
                    compact_bullet("根本原因", "区分直接原因、管理失效和设计缺陷"),
                    compact_bullet("教训迁移", "形成对本装置可复用的防范动作"),
                ],
                key_takeaway="案例价值不在复述事故，而在形成可迁移的控制措施。",
                placeholder_type="Chart",
            ),
            SlidePlan(
                chapter="Emergency Response",
                slide_title="Emergency Response",
                bullet_points=[
                    compact_bullet("响应触发", "定义报警、隔离和升级汇报的门槛"),
                    compact_bullet("处置流程", "说明泄漏、着火和人员暴露的动作顺序"),
                    compact_bullet("资源配置", "匹配洗消、堵漏和疏散资源"),
                ],
                key_takeaway="应急页的重点是动作顺序清晰，而不是术语齐全。",
                placeholder_type="Risk Matrix",
            ),
            SlidePlan(
                chapter="SOP",
                slide_title="SOP 与复盘",
                bullet_points=[
                    compact_bullet("标准步骤", "列出最关键的操作前、中、后检查点"),
                    compact_bullet("控制点", "强调许可、联锁和交接班确认"),
                    compact_bullet("持续改进", "把演练和事件复盘写回 SOP"),
                ],
                key_takeaway="培训闭环一定要落回 SOP 和考核机制。",
                placeholder_type="None",
            ),
        ],
    }
    return templates[scenario]


def split_dense_slides(slides: list[SlidePlan], max_bullets: int = 4) -> list[SlidePlan]:
    expanded: list[SlidePlan] = []
    for slide in slides:
        if len(slide.bullet_points) <= max_bullets:
            expanded.append(slide)
            continue
        chunks = [slide.bullet_points[i : i + max_bullets] for i in range(0, len(slide.bullet_points), max_bullets)]
        for index, chunk in enumerate(chunks, start=1):
            suffix = f" - Part {index}" if len(chunks) > 1 else ""
            expanded.append(
                SlidePlan(
                    chapter=slide.chapter,
                    slide_title=f"{slide.slide_title}{suffix}",
                    bullet_points=chunk,
                    key_takeaway=slide.key_takeaway,
                    placeholder_type=slide.placeholder_type,
                )
            )
    return expanded


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(description="Generate chemistry presentation JSON plans.")
    parser.add_argument("--scenario", choices=["academic", "process_design", "safety_training"], required=True)
    parser.add_argument("--topic", required=True)
    parser.add_argument("--text", help="Optional text to normalize")
    parser.add_argument("--json", action="store_true", help="Print slide JSON only")
    args = parser.parse_args()

    slides = split_dense_slides(build_outline(args.topic, args.scenario))
    normalized_text = normalize_text(args.text) if args.text else None

    if args.json:
        print(json.dumps([asdict(slide) for slide in slides], ensure_ascii=False, indent=2))
        return 0

    print(f"场景: {args.scenario}")
    print(f"主题: {normalize_text(args.topic)}")
    if normalized_text:
        print(f"预处理文本: {normalized_text}")
    for index, slide in enumerate(slides, start=1):
        print(f"\n{index}. {slide.slide_title} [{slide.chapter}]")
        for bullet in slide.bullet_points:
            print(f"- {bullet.keyword}: {bullet.description}")
        print(f"- key_takeaway: {slide.key_takeaway}")
        print(f"- placeholder_type: {slide.placeholder_type}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
