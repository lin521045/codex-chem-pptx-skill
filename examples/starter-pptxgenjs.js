const PptxGenJS = require("pptxgenjs");

const SUBSCRIPT_MAP = { "0": "₀", "1": "₁", "2": "₂", "3": "₃", "4": "₄", "5": "₅", "6": "₆", "7": "₇", "8": "₈", "9": "₉", "+": "₊", "-": "₋" };
const SUPERSCRIPT_MAP = { "0": "⁰", "1": "¹", "2": "²", "3": "³", "4": "⁴", "5": "⁵", "6": "⁶", "7": "⁷", "8": "⁸", "9": "⁹", "+": "⁺", "-": "⁻" };

function convertByMap(value, table) {
  return value.split("").map((char) => table[char] || char).join("");
}

function insertCnEnSpacing(text) {
  return text
    .replace(/([\u4e00-\u9fff])([A-Za-z0-9])/g, "$1 $2")
    .replace(/([A-Za-z0-9])([\u4e00-\u9fff])/g, "$1 $2");
}

function normalizeUnits(text) {
  return text
    .replace(/mol \/ L/g, "mol/L")
    .replace(/kJ \/ mol/g, "kJ/mol")
    .replace(/m3\/h/g, "m³/h")
    .replace(/wt %/g, "wt%")
    .replace(/° C/g, "°C")
    .replace(/(\d)(mmol|mol|kg|t|MPa|bar|°C|h)\b/g, "$1 $2");
}

function normalizeChemistryUnicode(text) {
  let output = text.replace(/\b(\d{1,3})([A-Z][a-z]?)(?=\s*(NMR|MS))/g, (_, digits, element) => {
    return `${convertByMap(digits, SUPERSCRIPT_MAP)}${element}`;
  });

  output = output.replace(/([A-Za-z\)])(\d+)/g, (_, prefix, digits) => {
    return `${prefix}${convertByMap(digits, SUBSCRIPT_MAP)}`;
  });

  output = output.replace(/\^([0-9+-]+)/g, (_, charge) => convertByMap(charge, SUPERSCRIPT_MAP));
  return output;
}

function formatChemText(text) {
  return normalizeChemistryUnicode(normalizeUnits(insertCnEnSpacing(text)));
}

function addSectionHeader(slide, title, subtitle, dark = false) {
  slide.background = { color: dark ? "16423C" : "F7FAF7" };
  slide.addText(title, {
    x: 0.65,
    y: 0.45,
    w: 8.9,
    h: 0.55,
    fontFace: "Cambria",
    fontSize: 24,
    bold: true,
    color: dark ? "FFFFFF" : "16423C",
    margin: 0,
  });
  slide.addText(subtitle, {
    x: 0.68,
    y: 1.0,
    w: 8.4,
    h: 0.35,
    fontFace: "Calibri",
    fontSize: 10.5,
    color: dark ? "D9EEE8" : "5A6A66",
    italic: true,
    margin: 0,
  });
}

function addBulletCard(slide, title, bullets, x, y, w, h, accent) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.08,
    fill: { color: "FFFFFF" },
    line: { color: "D7E3DD", width: 1 },
    shadow: { type: "outer", color: "000000", blur: 2, angle: 45, offset: 1, opacity: 0.08 },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x,
    y,
    w: 0.12,
    h,
    fill: { color: accent },
    line: { color: accent, width: 0 },
  });
  slide.addText(title, {
    x: x + 0.28,
    y: y + 0.2,
    w: w - 0.4,
    h: 0.3,
    fontFace: "Cambria",
    fontSize: 15,
    bold: true,
    color: "17312A",
    margin: 0,
  });

  const lines = [];
  bullets.forEach((bullet, index) => {
    lines.push({ text: `• ${formatChemText(bullet)}`, options: { breakLine: index < bullets.length - 1 } });
  });

  slide.addText(lines, {
    x: x + 0.3,
    y: y + 0.58,
    w: w - 0.45,
    h: h - 0.82,
    fontFace: "Calibri",
    fontSize: 11.5,
    color: "44555A",
    margin: 0,
    valign: "top",
  });
}

function addPlaceholder(slide, text, x, y, w, h) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.04,
    line: { color: "3D8D7A", width: 1.5, dash: "dash" },
    fill: { color: "EAF5F1" },
  });
  slide.addText(text, {
    x: x + 0.24,
    y: y + 0.24,
    w: w - 0.48,
    h: h - 0.48,
    fontFace: "Calibri",
    fontSize: 13,
    color: "16423C",
    align: "center",
    valign: "mid",
    bold: true,
    margin: 0,
  });
}

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_16x9";
pptx.author = "Codex Chemistry PPTX Skill";
pptx.company = "lin521045";
pptx.subject = "年产 10 万吨 Polycarbonate (PC) 工艺设计";
pptx.title = "年产 10 万吨 Polycarbonate (PC) 工艺设计";
pptx.lang = "zh-CN";

async function main() {
  const cover = pptx.addSlide();
  addSectionHeader(cover, "年产 10 万吨 Polycarbonate (PC) 工艺设计", "Chemistry & Chemical Engineering starter deck for Codex", true);
  cover.addText(formatChemText("采用非光气熔融酯交换路线，重点展示路线比选、Mass & Energy Balance、关键设备选型和 HSE 结论。"), {
    x: 0.72,
    y: 1.8,
    w: 6.2,
    h: 0.8,
    fontFace: "Calibri",
    fontSize: 18,
    color: "FFFFFF",
    margin: 0,
  });
  addPlaceholder(cover, "[在此插入 Aspen PFD/P&ID 总流程图]", 7.0, 1.35, 2.5, 2.8);

  const bg = pptx.addSlide();
  addSectionHeader(bg, "项目背景与产能规划", "Project Basis & Capacity Planning");
  addBulletCard(bg, "边界条件", [
    "目标产品为 Bisphenol A 基 Polycarbonate (PC)",
    "设计产能为 10 万吨/年，按 8000 h/a 计",
    "产品定位为光学级与工程塑料级双规格",
  ], 0.65, 1.55, 4.2, 2.05, "16423C");
  addBulletCard(bg, "工艺目标", [
    "优先采用非光气路线，降低光气相关 HSE 风险",
    "兼顾高分子量控制与副产 Phenol 回收",
    "预留与 DPC 上游装置的热集成接口",
  ], 0.65, 3.82, 4.2, 2.0, "3D8D7A");
  addPlaceholder(bg, "[在此插入项目区位 / 产品应用示意图]", 5.15, 1.55, 4.2, 4.27);

  const route = pptx.addSlide();
  addSectionHeader(route, "工艺路线比选", "Route Screening");
  addBulletCard(route, "候选路线 A", [
    "光气法成熟度高，但 HSE 与环保压力大",
    "装置隔离和尾气治理要求显著增加",
  ], 0.65, 1.7, 2.7, 1.8, "C84C05");
  addBulletCard(route, "候选路线 B", [
    "非光气熔融酯交换路线更适合当前政策环境",
    "副产 Phenol 可回收，流程集成潜力更好",
  ], 3.55, 1.7, 2.7, 1.8, "16423C");
  addBulletCard(route, "选择结论", [
    "最终选用 DPC 与 BPA 的熔融酯交换路线",
    "以 Techno-Economic Analysis 与 HSE 综合评分作为依据",
  ], 6.45, 1.7, 2.9, 1.8, "3D8D7A");
  addPlaceholder(route, "[在此插入路线比选流程示意图或评分矩阵]", 0.65, 3.9, 8.7, 1.7);

  const meb = pptx.addSlide();
  addSectionHeader(meb, "Mass & Energy Balance", "Key Streams and Utility Loads");
  addBulletCard(meb, "关键流股", [
    "PC 目标产出约 12.5 t/h",
    "BPA 进料约 9.8 t/h，DPC 进料约 10.6 t/h",
    "回收 Phenol 流量约 4.2 t/h",
  ], 0.65, 1.65, 4.15, 2.0, "1F3C88");
  addBulletCard(meb, "关键能耗", [
    "反应段操作温度控制在 240 °C",
    "真空脱挥系统是主要电耗节点",
    "建议配置热油系统与精馏塔热集成回路",
  ], 0.65, 3.9, 4.15, 2.0, "44576A");
  addPlaceholder(meb, "[在此插入 Mass & Energy Balance 表格]", 5.05, 1.65, 4.25, 1.95);
  addPlaceholder(meb, "[在此插入反应动力学曲线 Profile]", 5.05, 3.9, 4.25, 1.95);

  const equip = pptx.addSlide();
  addSectionHeader(equip, "关键设备选型与 HSE", "Equipment Selection, HSE & TEA");
  addBulletCard(equip, "关键设备", [
    "熔融缩聚反应器采用高黏体系强化传质设计",
    "脱挥器与精馏塔需兼顾高温与 Phenol 腐蚀环境",
    "关键转动设备设置 1 用 1 备",
  ], 0.65, 1.55, 4.2, 2.0, "16423C");
  addBulletCard(equip, "HSE / TEA 结论", [
    "避免光气后，重大毒害风险显著下降",
    "建议对高温熔体泄漏与真空失效开展 HAZOP",
    "在 Phenol 高回收率假设下项目具备经济可行性",
  ], 0.65, 3.8, 4.2, 2.0, "C84C05");
  addPlaceholder(equip, "[在此插入关键设备选型表]", 5.1, 1.55, 4.2, 1.95);
  addPlaceholder(equip, "[在此插入 HAZOP 风险矩阵或 CAPEX/OPEX 摘要图]", 5.1, 3.8, 4.2, 1.95);

  await pptx.writeFile({ fileName: "polycarbonate-process-design-starter.pptx" });
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
