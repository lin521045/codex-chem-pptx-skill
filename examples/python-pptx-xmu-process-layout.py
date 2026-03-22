#!/usr/bin/env python3
"""Generate a process-design deck with XMU-style top banner and left sidebar."""

from __future__ import annotations

import re
from dataclasses import dataclass

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt


@dataclass
class SlideSpec:
    chapter_name: str
    slide_title: str
    bullet_points: list[str]
    visual_placeholder: str


XMU_BLUE = RGBColor(0x00, 0x3F, 0x88)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BG = RGBColor(0xF8, 0xF9, 0xFA)
TEXT = RGBColor(0x33, 0x33, 0x33)
DIM = RGBColor(0x88, 0x88, 0x88)
SIDEBAR_BG = RGBColor(0xF1, 0xF4, 0xF8)
LINE = RGBColor(0xD9, 0xE2, 0xEC)
PLACEHOLDER_BG = RGBColor(0xEE, 0xF4, 0xFB)

SLIDE_W = 13.333
SLIDE_H = 7.5
BANNER_H = 0.9
SIDEBAR_W = 2.7
SAFE_X = 3.35
SAFE_Y = 1.5
SAFE_W = 9.35
SAFE_H = 5.6

SUBSCRIPT_MAP = str.maketrans("0123456789+-()", "₀₁₂₃₄₅₆₇₈₉₊₋₍₎")
SUPERSCRIPT_MAP = str.maketrans("0123456789+-()", "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁽⁾")


def insert_cn_en_spacing(text: str) -> str:
    text = re.sub(r"([\u4e00-\u9fff])([A-Za-z0-9])", r"\1 \2", text)
    text = re.sub(r"([A-Za-z0-9])([\u4e00-\u9fff])", r"\1 \2", text)
    return text


def normalize_units(text: str) -> str:
    replacements = {
        "mol / L": "mol/L",
        "kJ / mol": "kJ/mol",
        "m3/h": "m³/h",
        "m3 / h": "m³/h",
        "wt %": "wt%",
        "° C": "°C",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    text = re.sub(r"(\d)(mol|kg|t|MPa|bar|°C|h|wt%|kmol/h|t/h|m³/h)\b", r"\1 \2", text)
    return text


def to_subscript(value: str) -> str:
    return value.translate(SUBSCRIPT_MAP)


def to_superscript(value: str) -> str:
    return value.translate(SUPERSCRIPT_MAP)


def normalize_chem_text(text: str) -> str:
    text = insert_cn_en_spacing(text)
    text = normalize_units(text)
    text = re.sub(r"\b(\d{1,3})([A-Z][a-z]?)(?=\s*(NMR|MS))", lambda m: f"{to_superscript(m.group(1))}{m.group(2)}", text)
    text = re.sub(r"([A-Za-z\)])(\d+)", lambda m: m.group(1) + to_subscript(m.group(2)), text)
    text = re.sub(r"\^([0-9+-]+)", lambda m: to_superscript(m.group(1)), text)
    return text


def add_banner(slide, title: str) -> None:
    banner = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(0),
        Inches(0),
        Inches(SLIDE_W),
        Inches(BANNER_H),
    )
    banner.fill.solid()
    banner.fill.fore_color.rgb = XMU_BLUE
    banner.line.fill.background()

    title_box = slide.shapes.add_textbox(Inches(0.55), Inches(0.14), Inches(SLIDE_W - 1.1), Inches(0.56))
    frame = title_box.text_frame
    frame.clear()
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = frame.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = normalize_chem_text(title)
    run.font.name = "Arial"
    run.font.size = Pt(23)
    run.font.bold = True
    run.font.color.rgb = WHITE


def add_sidebar(slide, chapters: list[str], active_index: int) -> None:
    sidebar = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(0),
        Inches(BANNER_H),
        Inches(SIDEBAR_W),
        Inches(SLIDE_H - BANNER_H),
    )
    sidebar.fill.solid()
    sidebar.fill.fore_color.rgb = SIDEBAR_BG
    sidebar.line.fill.background()

    label_box = slide.shapes.add_textbox(Inches(0.28), Inches(1.08), Inches(2.0), Inches(0.25))
    label_frame = label_box.text_frame
    label_frame.clear()
    p = label_frame.paragraphs[0]
    run = p.add_run()
    run.text = "PROCESS NAVIGATION"
    run.font.name = "Arial"
    run.font.size = Pt(9)
    run.font.bold = True
    run.font.color.rgb = DIM

    nav_top = 1.45
    nav_step = 0.82
    for idx, chapter in enumerate(chapters):
        y = nav_top + idx * nav_step
        active = idx == active_index
        if active:
            highlight = slide.shapes.add_shape(
                MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
                Inches(0.18),
                Inches(y - 0.05),
                Inches(2.28),
                Inches(0.52),
            )
            highlight.fill.solid()
            highlight.fill.fore_color.rgb = XMU_BLUE
            highlight.line.fill.background()

        box = slide.shapes.add_textbox(Inches(0.34), Inches(y), Inches(2.0), Inches(0.34))
        frame = box.text_frame
        frame.clear()
        p = frame.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        run = p.add_run()
        run.text = chapter
        run.font.name = "Arial"
        run.font.size = Pt(13 if active else 11.5)
        run.font.bold = active
        run.font.color.rgb = WHITE if active else DIM


def add_safe_zone(slide) -> None:
    box = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(SAFE_X),
        Inches(SAFE_Y),
        Inches(SAFE_W),
        Inches(SAFE_H),
    )
    box.fill.solid()
    box.fill.fore_color.rgb = WHITE
    box.line.color.rgb = LINE
    box.line.width = Pt(1)


def add_body_and_placeholder(slide, spec: SlideSpec) -> None:
    left_x = SAFE_X + 0.35
    top_y = SAFE_Y + 0.35
    content_w = SAFE_W * 0.56
    placeholder_x = SAFE_X + SAFE_W * 0.63
    placeholder_w = SAFE_W * 0.30

    body_box = slide.shapes.add_textbox(Inches(left_x), Inches(top_y), Inches(content_w), Inches(4.6))
    frame = body_box.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.vertical_anchor = MSO_ANCHOR.TOP

    for idx, bullet in enumerate(spec.bullet_points):
        p = frame.paragraphs[0] if idx == 0 else frame.add_paragraph()
        p.level = 0
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(8)
        p.bullet = True
        run = p.add_run()
        run.text = normalize_chem_text(bullet)
        run.font.name = "Arial"
        run.font.size = Pt(16)
        run.font.color.rgb = TEXT

    placeholder = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        Inches(placeholder_x),
        Inches(top_y),
        Inches(placeholder_w),
        Inches(4.2),
    )
    placeholder.fill.solid()
    placeholder.fill.fore_color.rgb = PLACEHOLDER_BG
    placeholder.line.color.rgb = XMU_BLUE
    placeholder.line.width = Pt(1.5)

    placeholder_box = slide.shapes.add_textbox(
        Inches(placeholder_x + 0.18),
        Inches(top_y + 0.2),
        Inches(placeholder_w - 0.36),
        Inches(3.8),
    )
    placeholder_frame = placeholder_box.text_frame
    placeholder_frame.clear()
    placeholder_frame.word_wrap = True
    placeholder_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = placeholder_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = normalize_chem_text(spec.visual_placeholder)
    run.font.name = "Arial"
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = XMU_BLUE


def create_content_slide(prs: Presentation, chapters: list[str], active_index: int, spec: SlideSpec):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = BG
    add_banner(slide, spec.slide_title)
    add_sidebar(slide, chapters, active_index)
    add_safe_zone(slide)
    add_body_and_placeholder(slide, spec)
    return slide


def build_process_specs() -> list[SlideSpec]:
    return [
        SlideSpec(
            chapter_name="项目背景与产能规划",
            slide_title="项目背景与产能规划",
            bullet_points=[
                "目标产品为 Polycarbonate (PC)，设计产能为 10 万吨/年",
                "路线选择需兼顾高分子量控制、Phenol 回收与 HSE 风险",
                "装置按 8000 h/a 运行，并预留与上游 DPC 系统热集成接口",
            ],
            visual_placeholder="[在此插入产品应用图或项目边界示意图]",
        ),
        SlideSpec(
            chapter_name="工艺路线比选",
            slide_title="工艺路线比选与推荐方案",
            bullet_points=[
                "比较光气法与非光气熔融酯交换法的安全性、成熟度与环保压力",
                "非光气路线更符合当前政策环境，并具备更好的副产 Phenol 回收潜力",
                "推荐采用 DPC 与 BPA 熔融酯交换路线作为基准方案",
            ],
            visual_placeholder="[在此插入 Aspen PFD / 路线评分矩阵]",
        ),
        SlideSpec(
            chapter_name="物料与能量衡算",
            slide_title="物料衡算 (Mass Balance) 与能量衡算 (Energy Balance)",
            bullet_points=[
                "PC 目标产出约 12.5 t/h，BPA 进料约 9.8 t/h，DPC 进料约 10.6 t/h",
                "回收 Phenol 流量约 4.2 t/h，需重点考虑精馏系统负荷",
                "主要热负荷集中在反应段、脱挥段与塔系再沸器，可通过热油系统集成",
            ],
            visual_placeholder="[在此插入 Mass Balance 表格与 Energy Balance 热负荷图]",
        ),
        SlideSpec(
            chapter_name="关键设备选型",
            slide_title="关键设备选型",
            bullet_points=[
                "反应器需适配高黏熔体体系，并强化传质与停留时间分布控制",
                "脱挥器与精馏塔材质需兼顾高温与 Phenol 腐蚀环境",
                "关键转动设备按 1 用 1 备配置，提升装置连续运行稳定性",
            ],
            visual_placeholder="[在此插入关键设备选型表或设备剖面图]",
        ),
        SlideSpec(
            chapter_name="HSE 与技术经济评价",
            slide_title="HSE 与技术经济评价 (Techno-Economic Analysis)",
            bullet_points=[
                "非光气路线显著降低光气相关重大毒害风险，但需关注高温熔体泄漏与真空失效",
                "建议对反应段、脱挥段和精馏塔系统开展 HAZOP / LOPA 分析",
                "在高 Phenol 回收率与公用工程优化条件下，项目具备良好经济可行性",
            ],
            visual_placeholder="[在此插入 HAZOP 风险矩阵或 CAPEX / OPEX 摘要图]",
        ),
    ]


def main() -> int:
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)

    specs = build_process_specs()
    chapters = [item.chapter_name for item in specs]

    for index, spec in enumerate(specs):
        create_content_slide(prs, chapters, index, spec)

    prs.save("xmu-chem-process-design-sample.pptx")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
