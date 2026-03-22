#!/usr/bin/env python3
"""Domain-specific outline and text preprocessing helpers for chemistry decks."""

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
class SlidePlan:
    title: str
    bullets: list[str]
    placeholder: str


def insert_cn_en_spacing(text: str) -> str:
    text = re.sub(fr"([{CJK}])([{ASCII_BLOCK}])", r"\1 \2", text)
    text = re.sub(fr"([{ASCII_BLOCK}])([{CJK}])", r"\1 \2", text)
    return text


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
    text = re.sub(r"(\d)(mmol|mol|g|kg|t|wt%|MPa|kPa|bar|°C|K|h|min|s)\b", r"\1 \2", text)
    return text


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
    if charge:
        charge = to_superscript(charge)
    return isotope + core + charge


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


def choose_placeholder(title: str, scenario: str) -> str:
    title_lower = title.lower()
    if "mechanism" in title_lower or "机理" in title:
        return "[在此插入 Reaction Mechanism 示意图]"
    if any(key in title_lower for key in ("nmr", "ir", "ms", "xrd")) or "表征" in title:
        return "[在此插入 NMR/IR/MS/XRD 谱图]"
    if any(key in title for key in ("流程", "路线", "衡算")) or "pfd" in title_lower or "p&id" in title_lower:
        return "[在此插入 Aspen PFD/P&ID 流程图]"
    if "设备" in title or "equipment" in title_lower:
        return "[在此插入关键设备选型表或设备剖面图]"
    if "hse" in title_lower or "安全" in title or "应急" in title:
        return "[在此插入 HAZOP/LOPA 风险矩阵或应急流程图]"
    if scenario == "academic":
        return "[在此插入 ChemDraw 分子结构图或表征图片]"
    if scenario == "process_design":
        return "[在此插入工艺流程图或 Mass & Energy Balance 表格]"
    return "[在此插入事故案例图示或 SOP 流程图]"


def build_outline(topic: str, scenario: str) -> list[SlidePlan]:
    templates: dict[str, list[tuple[str, list[str]]]] = {
        "academic": [
            ("研究背景与意义", [f"{topic} 的研究动机", "行业或学术痛点", "本研究的目标与贡献"]),
            ("Reaction Mechanism 假设", ["关键中间体与路径设想", "催化位或活性中心说明", "需要验证的核心问题"]),
            ("实验方法与表征", ["反应条件设计", "NMR/IR/MS 等表征方案", "数据采集与重复性"]),
            ("结果与讨论", ["活性、选择性或收率结果", "谱图或显微图证据", "与文献或对照组比较"]),
            ("结论与展望", ["核心发现总结", "方法学局限性", "下一步研究方向"]),
        ],
        "process_design": [
            ("项目背景与产能规划", [f"{topic} 的建设背景", "产品指标与边界条件", "设计产能与运行假设"]),
            ("工艺路线比选", ["候选路线优缺点", "原料与公用工程适配性", "最终路线选择依据"]),
            ("Mass & Energy Balance", ["主要流股组成", "能量负荷与回收机会", "关键操作参数"]),
            ("关键设备选型", ["反应器与分离设备方案", "材质、操作窗口与冗余设计", "设备尺寸或负荷校核"]),
            ("HSE 与 Techno-Economic Analysis", ["危险源识别与控制措施", "CAPEX/OPEX 估算", "经济性结论与敏感性"]),
        ],
        "safety_training": [
            ("化学品危险性分析", [f"{topic} 相关介质危险特性", "暴露途径与后果", "高风险场景识别"]),
            ("MSDS 解读", ["关键理化参数", "储运与个体防护要求", "禁配与稳定性"]),
            ("典型事故案例剖析", ["事故经过", "直接原因与根本原因", "教训与防范措施"]),
            ("Emergency Response", ["报警与隔离流程", "泄漏/着火处置步骤", "升级汇报机制"]),
            ("SOP 与复盘", ["标准操作步骤", "关键控制点", "考核与持续改进"]),
        ],
    }
    return [
        SlidePlan(title=title, bullets=[normalize_text(item) for item in bullets], placeholder=choose_placeholder(title, scenario))
        for title, bullets in templates[scenario]
    ]


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser(description="Generate chemistry deck outline helpers.")
    parser.add_argument("--scenario", choices=["academic", "process_design", "safety_training"], required=True)
    parser.add_argument("--topic", required=True)
    parser.add_argument("--text", help="Optional text to normalize")
    parser.add_argument("--json", action="store_true", help="Print JSON instead of markdown-like text")
    args = parser.parse_args()

    outline = build_outline(args.topic, args.scenario)
    normalized = normalize_text(args.text) if args.text else None

    payload = {
        "scenario": args.scenario,
        "topic": normalize_text(args.topic),
        "normalized_text": normalized,
        "slides": [asdict(item) for item in outline],
    }

    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0

    print(f"场景: {payload['scenario']}")
    print(f"主题: {payload['topic']}")
    if normalized:
        print(f"预处理文本: {normalized}")
    for index, slide in enumerate(payload["slides"], start=1):
        print(f"\n{index}. {slide['title']}")
        for bullet in slide["bullets"]:
            print(f"- {bullet}")
        print(f"- {slide['placeholder']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
