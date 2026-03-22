const PptxGenJS = require("pptxgenjs");

function spaceCnEn(text) {
  return text
    .replace(/([\u4e00-\u9fff])([A-Za-z0-9])/g, "$1 $2")
    .replace(/([A-Za-z0-9])([\u4e00-\u9fff])/g, "$1 $2");
}

function chemText(text) {
  return spaceCnEn(text)
    .replace(/CO2/g, "CO₂")
    .replace(/Fe-N4/g, "Fe-N₄")
    .replace(/13C NMR/g, "¹³C NMR")
    .replace(/H2O/g, "H₂O")
    .replace(/SO4\^2-/g, "SO₄²⁻")
    .replace(/(\d)(mmol|mol|mA|V|h|wt%|nm|eV|cm-1|cm\^-1)\b/g, "$1 $2")
    .replace(/cm-1/g, "cm⁻¹");
}

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_16x9";
pptx.author = "Codex Chemistry PPTX Skill";
pptx.company = "lin521045";
pptx.subject = "Academic chemistry research sample";
pptx.title = "MOF 衍生 Fe-N-C 单原子催化剂用于 CO₂ 电还原";
pptx.lang = "zh-CN";

function addHeader(slide, title, subtitle, dark = false) {
  slide.background = { color: dark ? "102A43" : "F7FAFC" };
  slide.addText(title, {
    x: 0.6,
    y: 0.42,
    w: 8.8,
    h: 0.5,
    fontFace: "Cambria",
    fontSize: 24,
    bold: true,
    color: dark ? "FFFFFF" : "102A43",
    margin: 0,
  });
  slide.addText(subtitle, {
    x: 0.63,
    y: 0.98,
    w: 8.6,
    h: 0.28,
    fontFace: "Calibri",
    fontSize: 10.5,
    color: dark ? "D9E8F5" : "556B7B",
    italic: true,
    margin: 0,
  });
}

function addCard(slide, cfg) {
  const { x, y, w, h, title, bullets, accent = "0F766E" } = cfg;
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.04,
    fill: { color: "FFFFFF" },
    line: { color: "D7E3EA", width: 1 },
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
    x: x + 0.26,
    y: y + 0.18,
    w: w - 0.38,
    h: 0.3,
    fontFace: "Cambria",
    fontSize: 15,
    bold: true,
    color: "17324D",
    margin: 0,
  });

  const rich = [];
  bullets.forEach((bullet, index) => {
    rich.push({ text: `• ${chemText(bullet)}`, options: { breakLine: index < bullets.length - 1 } });
  });

  slide.addText(rich, {
    x: x + 0.28,
    y: y + 0.56,
    w: w - 0.44,
    h: h - 0.75,
    fontFace: "Calibri",
    fontSize: 11.5,
    color: "425466",
    margin: 0,
    valign: "top",
  });
}

function addPlaceholder(slide, text, x, y, w, h, fill = "E8F7F5", line = "0F766E") {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.03,
    fill: { color: fill },
    line: { color: line, width: 1.4, dash: "dash" },
  });
  slide.addText(text, {
    x: x + 0.18,
    y: y + 0.18,
    w: w - 0.36,
    h: h - 0.36,
    fontFace: "Calibri",
    fontSize: 13,
    bold: true,
    color: "134E4A",
    align: "center",
    valign: "mid",
    margin: 0,
  });
}

function addTag(slide, text, x, y, w, fill) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h: 0.34,
    rectRadius: 0.08,
    fill: { color: fill },
    line: { color: fill, width: 0 },
  });
  slide.addText(text, {
    x,
    y: y + 0.03,
    w,
    h: 0.24,
    fontFace: "Calibri",
    fontSize: 10.5,
    bold: true,
    color: "FFFFFF",
    align: "center",
    margin: 0,
  });
}

async function main() {
  const cover = pptx.addSlide();
  addHeader(cover, "MOF 衍生 Fe-N-C 单原子催化剂用于 CO₂ 电还原", "Academic chemistry research sample for $chem-pptx-codex", true);
  cover.addText(chemText("围绕 Fe-N₄ 活性位构建、原位表征与 Reaction Mechanism 推断，展示一套科研风格的中文答辩 / 汇报页面。"), {
    x: 0.68,
    y: 1.82,
    w: 5.9,
    h: 0.9,
    fontFace: "Calibri",
    fontSize: 18,
    color: "FFFFFF",
    margin: 0,
  });
  addTag(cover, "CO₂RR", 0.7, 3.12, 1.0, "0F766E");
  addTag(cover, "Single-Atom", 1.82, 3.12, 1.32, "2563EB");
  addTag(cover, "In-situ FTIR", 3.28, 3.12, 1.42, "7C3AED");
  addPlaceholder(cover, "[在此插入 MOF 衍生催化剂结构示意图]", 7.02, 1.3, 2.45, 3.15, "DDEFFB", "2563EB");

  const bg = pptx.addSlide();
  addHeader(bg, "研究背景与意义", "Why Fe-N-C for CO₂ electroreduction?");
  addCard(bg, {
    x: 0.62,
    y: 1.55,
    w: 4.1,
    h: 2.02,
    title: "问题提出",
    accent: "0F766E",
    bullets: [
      "CO₂ 电还原可实现 C1 化学品的低碳转化",
      "贵金属催化剂成本高，非贵金属体系更具应用潜力",
      "Fe-N-C 在 CO 选择性方面表现突出，但活性位争议仍然存在",
    ],
  });
  addCard(bg, {
    x: 0.62,
    y: 3.82,
    w: 4.1,
    h: 1.88,
    title: "本研究目标",
    accent: "2563EB",
    bullets: [
      "构建高密度 Fe-N₄ 位点",
      "联用 XPS、XANES 和 In-situ FTIR 解析结构演变",
      "建立性能与位点结构之间的关联",
    ],
  });
  addPlaceholder(bg, "[在此插入文献综述图或 CO₂RR 路径示意图]", 5.0, 1.55, 4.28, 4.15);

  const strategy = pptx.addSlide();
  addHeader(strategy, "催化剂设计与 Reaction Mechanism 假设", "From MOF precursor to Fe-N₄ active site");
  addPlaceholder(strategy, "[在此插入 ChemDraw 分子结构图 / 合成路线图]", 0.68, 1.52, 4.12, 2.1);
  addCard(strategy, {
    x: 5.02,
    y: 1.52,
    w: 4.26,
    h: 2.1,
    title: "设计策略",
    accent: "7C3AED",
    bullets: [
      "以 ZIF-8 为前驱体限域 Fe 物种",
      "热解后形成分散的 Fe-N₄ 单原子位点",
      "通过孔结构调控提升 CO₂ 传质与局部 pH 环境",
    ],
  });
  addCard(strategy, {
    x: 0.68,
    y: 3.9,
    w: 8.6,
    h: 1.72,
    title: "Reaction Mechanism 假设",
    accent: "0F766E",
    bullets: [
      "Fe-N₄ 位点优先吸附 *COOH 中间体并降低形成能垒",
      "适中的 *CO 脱附能力有利于高 FECO 输出",
      "In-situ FTIR 预计可观察到 *COOH 与 *CO 特征吸收信号",
    ],
  });

  const methods = pptx.addSlide();
  addHeader(methods, "实验方法与表征", "Synthesis, electrochemistry and characterization");
  addCard(methods, {
    x: 0.65,
    y: 1.52,
    w: 2.75,
    h: 3.9,
    title: "样品制备",
    accent: "2563EB",
    bullets: [
      "Fe 盐与 2-methylimidazole 协同组装得到 Fe@ZIF",
      "900 °C 热解并酸洗去除纳米颗粒",
      "对比样为 N-C 与 Fe nanoparticle/C",
    ],
  });
  addCard(methods, {
    x: 3.58,
    y: 1.52,
    w: 2.75,
    h: 3.9,
    title: "电化学测试",
    accent: "0F766E",
    bullets: [
      "H-type 电解池，0.5 mol/L KHCO₃ 电解液",
      "电位窗口为 -0.4 V 至 -1.0 V vs. RHE",
      "产物通过 GC 与 ¹H NMR 定量分析",
    ],
  });
  addPlaceholder(methods, "[在此插入 XRD / XPS / SEM / TEM 表征总览图]", 6.5, 1.52, 2.82, 1.78, "EEF6FF", "2563EB");
  addPlaceholder(methods, "[在此插入电化学测试装置与 In-situ FTIR 示意图]", 6.5, 3.64, 2.82, 1.78, "ECFDF5", "0F766E");

  const results = pptx.addSlide();
  addHeader(results, "结果与讨论", "Catalytic activity, selectivity and structural evidence");
  addPlaceholder(results, "[在此插入 FECO 与 jCO 随电位变化曲线]", 0.7, 1.55, 4.0, 2.15, "EEF6FF", "2563EB");
  addPlaceholder(results, "[在此插入 SEM/TEM 与 HAADF-STEM 图片]", 5.0, 1.55, 4.25, 2.15, "ECFDF5", "0F766E");
  addCard(results, {
    x: 0.7,
    y: 3.98,
    w: 4.0,
    h: 1.62,
    title: "性能结果",
    accent: "2563EB",
    bullets: [
      "在 -0.72 V vs. RHE 下 FECO 达到 94%",
      "jCO 为 28.6 mA/cm²，稳定运行 20 h 性能衰减有限",
      "与 N-C 对照样相比，活性与选择性均显著提升",
    ],
  });
  addCard(results, {
    x: 5.0,
    y: 3.98,
    w: 4.25,
    h: 1.62,
    title: "结构证据",
    accent: "0F766E",
    bullets: [
      "HAADF-STEM 未观察到明显 Fe 团簇",
      "XANES / EXAFS 支持 Fe-N₄ 配位结构",
      "In-situ FTIR 捕获 *COOH 与线性吸附 *CO 信号",
    ],
  });

  const conclusion = pptx.addSlide();
  addHeader(conclusion, "结论与展望", "Takeaways and next steps");
  addCard(conclusion, {
    x: 0.68,
    y: 1.56,
    w: 4.18,
    h: 3.95,
    title: "核心结论",
    accent: "0F766E",
    bullets: [
      "MOF 衍生策略可有效构建高分散 Fe-N₄ 单原子位点",
      "Fe-N₄ 位点对 CO₂RR 中 *COOH 形成具有促进作用",
      "结合多尺度表征可较清晰地建立位点结构 - 性能关系",
    ],
  });
  addCard(conclusion, {
    x: 5.08,
    y: 1.56,
    w: 4.2,
    h: 1.88,
    title: "局限性",
    accent: "2563EB",
    bullets: [
      "当前仅完成 H-type 电解池验证",
      "原位谱学时间分辨率仍有限",
      "缺少更高电流密度条件下的数据",
    ],
  });
  addPlaceholder(conclusion, "[在此插入未来工作路线图：MEA、DFT 与 operando XAS]", 5.08, 3.72, 4.2, 1.8, "F4F1FF", "7C3AED");

  await pptx.writeFile({ fileName: "academic-chemistry-research-sample.pptx" });
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
