import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const workspaceDir = "C:/Users/s1325/Documents/小龙虾制成塑料袋";
const skillDir = "C:/Users/s1325/.codex/plugins/cache/openai-primary-runtime/presentations/26.909.12148/skills/presentations";
const buildDir = path.join(workspaceDir, "ppt_build");
const stagingDir = path.join(workspaceDir, ".codex-finalizer");
const outDir = path.join(workspaceDir, "ppt_output");
const finalPath = path.join(outDir, "小龙虾壳制环保塑料袋_从原料到化学反应_v2.pptx");
const imgDir = workspaceDir;
const font = "Microsoft YaHei";

await fs.mkdir(buildDir, { recursive: true });
await fs.mkdir(stagingDir, { recursive: true });
await fs.mkdir(outDir, { recursive: true });

const ppt = Presentation.create({ slideSize: { width: 1280, height: 720 } });

const C = {
  ink: "#15232E",
  navy: "#0C2D3D",
  ocean: "#167A8A",
  teal: "#1EA7A1",
  orange: "#D96C3B",
  coral: "#F19A62",
  cream: "#FFF8EE",
  sand: "#F4E9D7",
  paper: "#F7FAF8",
  muted: "#617079",
  line: "#D8E2E3",
  red: "#B84E4E",
  green: "#2E7D5A",
  white: "#FFFFFF",
};

const noLine = { fill: "none", width: 0 };
const thinLine = { style: "solid", fill: C.line, width: 1 };

function addText(slide, text, x, y, w, h, opts = {}) {
  const s = slide.shapes.add({
    geometry: opts.geometry || "textbox",
    position: { left: x, top: y, width: w, height: h },
    fill: opts.fill ?? "none",
    line: opts.line ?? noLine,
    borderRadius: opts.radius,
  });
  s.text = text;
  s.text.style = {
    typeface: opts.font || font,
    fontSize: opts.size || 24,
    bold: opts.bold || false,
    italic: opts.italic || false,
    color: opts.color || C.ink,
    alignment: opts.align || "left",
    verticalAlignment: opts.valign || "top",
    autoFit: opts.autoFit || "shrinkText",
    wrap: "square",
    insets: opts.insets || { left: 0, right: 0, top: 0, bottom: 0 },
  };
  return s;
}

function addRect(slide, x, y, w, h, fill, radius = 18, line = noLine) {
  return slide.shapes.add({ geometry: "roundRect", position: { left: x, top: y, width: w, height: h }, fill, line, borderRadius: radius });
}

function addLine(slide, x1, y1, x2, y2, color = C.line, width = 2) {
  return slide.shapes.add({ geometry: "line", position: { left: x1, top: y1, width: x2 - x1, height: y2 - y1 }, line: { style: "solid", fill: color, width }, fill: "none" });
}

function addArrow(slide, x, y, w, h, fill = C.orange) {
  return slide.shapes.add({ geometry: "rightArrow", position: { left: x, top: y, width: w, height: h }, fill, line: noLine });
}

async function bytes(file) {
  return new Uint8Array(await fs.readFile(path.join(imgDir, file)));
}

const coverImg = await bytes("video_contact_sheet.jpg");
const rawImg = await bytes("video_0_250.jpg");
const bagImg = await bytes("video_250_500.jpg");

function addHeader(slide, title, kicker = "小龙虾壳制环保塑料袋") {
  addText(slide, kicker.toUpperCase(), 64, 34, 360, 24, { size: 14, bold: true, color: C.ocean });
  addText(slide, title, 64, 66, 1120, 58, { size: 38, bold: true, color: C.navy });
  addLine(slide, 64, 132, 1216, 132, C.line, 1);
}

function addFooter(slide, n, text = "讲解材料 · 研究底稿与开放论文整理") {
  addText(slide, text, 64, 686, 780, 18, { size: 12, color: C.muted });
  addText(slide, String(n).padStart(2, "0"), 1160, 680, 56, 26, { size: 16, bold: true, color: C.ocean, align: "right" });
}

// 1 Cover
{
  const s = ppt.slides.add();
  s.background.fill = C.navy;
  s.images.add({ blob: coverImg, contentType: "image/jpeg", alt: "小龙虾壳处理、薄膜和袋子演示画面", fit: "cover", position: { left: 0, top: 0, width: 1280, height: 720 } });
  s.shapes.add({ geometry: "rect", position: { left: 0, top: 0, width: 1280, height: 720 }, fill: "#0C2D3D/76", line: noLine });
  addText(s, "从小龙虾壳\n到环保塑料袋", 72, 170, 720, 170, { size: 56, bold: true, color: C.white });
  addText(s, "一场面向零基础观众的材料化学导览", 76, 370, 640, 46, { size: 24, color: "#E9F4F2" });
  addText(s, "原料 → 反应 → 成膜 → 袋子 → 证据边界", 76, 450, 700, 34, { size: 18, color: C.coral, bold: true });
  addText(s, "基于项目视频、研究底稿与开放获取论文", 76, 630, 620, 22, { size: 14, color: "#D8E7E8" });
  s.speakerNotes.textFrame.setText("封面画面来自项目视频抽帧。全套内容依据研究底稿.md、项目视频以及 6 篇开放获取论文整理。\n提示观众：今天讲的是材料路线和化学原理，也会讲清楚哪些结论尚未被验证。");
}

// 2 raw material
{
  const s = ppt.slides.add();
  s.background.fill = C.cream;
  addHeader(s, "一只虾壳，为什么有机会变成薄膜？");
  s.images.add({ blob: rawImg, contentType: "image/jpeg", alt: "项目视频中的小龙虾壳与粉碎处理", fit: "cover", position: { left: 720, top: 168, width: 496, height: 432 }, geometry: "roundRect", borderRadius: 20 });
  addText(s, "虾壳不是单一材料", 68, 182, 560, 44, { size: 30, bold: true, color: C.navy });
  addText(s, "它是一种天然复合结构：矿物、蛋白质和几丁质彼此交织。", 68, 238, 560, 62, { size: 24, color: C.ink });
  const items = [
    ["碳酸钙", "让外壳硬，但会妨碍成膜", C.orange],
    ["蛋白质", "需要先去除，减少杂质", C.teal],
    ["几丁质", "真正的高分子骨架", C.ocean],
  ];
  items.forEach((it, i) => {
    const y = 340 + i * 82;
    addRect(s, 72, y, 22, 22, it[2], 11);
    addText(s, it[0], 110, y - 4, 140, 30, { size: 22, bold: true, color: C.navy });
    addText(s, it[1], 270, y - 4, 360, 36, { size: 20, color: C.muted });
  });
  addText(s, "关键转换", 72, 590, 100, 24, { size: 16, bold: true, color: C.orange });
  addText(s, "把“硬壳”拆成可溶解、可浇铸的高分子。", 184, 588, 500, 28, { size: 20, bold: true, color: C.navy });
  addFooter(s, 2);
  s.speakerNotes.textFrame.setText("来源：研究底稿.md 第 1、4 节；项目视频约 0:00–1:20。\n说明：天然壳的主要目标组分是几丁质，壳聚糖通常由几丁质脱乙酰制得。视频中把壳聚糖作为目标物的说法需要用这一层化学关系补充解释。");
}

// 3 process map
{
  const s = ppt.slides.add();
  s.background.fill = C.paper;
  addHeader(s, "完整路线：从壳到袋，至少要过六道关");
  const steps = [
    ["01", "清洗\n干燥\n粉碎", C.sand],
    ["02", "酸处理\n去矿物", "#FBE4D6"],
    ["03", "碱处理\n去蛋白", "#E2F2EF"],
    ["04", "浓碱\n脱乙酰", "#DDECF5"],
    ["05", "醋酸溶解\n加入甘油", "#F6E8C9"],
    ["06", "浇铸干燥\n裁切封合", "#E6F0E4"],
  ];
  steps.forEach((st, i) => {
    const x = 64 + i * 190;
    addRect(s, x, 230, 150, 176, st[2], 22, { style: "solid", fill: "#FFFFFF/60", width: 1 });
    addText(s, st[0], x + 18, 248, 70, 24, { size: 14, bold: true, color: C.orange });
    addText(s, st[1], x + 18, 292, 114, 76, { size: 25, bold: true, color: C.navy, valign: "middle" });
    if (i < steps.length - 1) addArrow(s, x + 158, 304, 28, 24, C.orange);
  });
  addText(s, "化学上最关键的三次变化", 70, 458, 360, 28, { size: 18, bold: true, color: C.ocean });
  addText(s, "去矿物：让无机硬壳离开\n脱乙酰：把几丁质变成可溶的壳聚糖\n成膜：让分散的高分子重新排成连续薄层", 70, 500, 860, 110, { size: 22, color: C.ink });
  addRect(s, 950, 474, 256, 122, C.navy, 18);
  addText(s, "袋子只是最后一步\n前面的材料纯度和分子结构\n决定了它能不能撑住", 976, 493, 205, 90, { size: 20, bold: true, color: C.white, align: "center", valign: "middle" });
  addFooter(s, 3);
  s.speakerNotes.textFrame.setText("来源：研究底稿.md 第 4 节及项目视频。\n提示：视频实际展示了多次酸/碱处理、溶液浇铸和袋子演示，但没有给出足以安全复现的完整条件。本页用流程关系帮助观众建立整体地图，不等于实验 SOP。");
}

// 4 decalcification reaction
{
  const s = ppt.slides.add();
  s.background.fill = C.cream;
  addHeader(s, "反应一：酸把碳酸钙从壳里“洗掉”");
  addText(s, "壳中的矿物主要以碳酸钙存在。加入酸后，会出现肉眼可见的气泡。", 72, 172, 1060, 38, { size: 24, color: C.ink });
  addRect(s, 80, 258, 1120, 132, "#FBE4D6", 24);
  addText(s, "CaCO₃  +  2H⁺   →   Ca²⁺  +  CO₂↑  +  H₂O", 110, 292, 1060, 62, { size: 34, bold: true, color: C.navy, align: "center", valign: "middle" });
  addText(s, "这是一个“酸与碳酸盐反应”的典型例子。", 110, 408, 540, 30, { size: 20, bold: true, color: C.orange });
  addText(s, "气泡来自二氧化碳\n矿物转成可溶的钙盐\n固体骨架变得更轻、更柔软", 110, 452, 430, 120, { size: 23, color: C.ink });
  addRect(s, 690, 430, 470, 150, C.navy, 20);
  addText(s, "为什么必须先去矿？", 720, 452, 410, 28, { size: 20, bold: true, color: C.coral });
  addText(s, "因为无机矿物不溶于后续的醋酸成膜体系，会让膜粗糙、脆，甚至难以连续成形。", 720, 490, 400, 72, { size: 20, color: C.white });
  addFooter(s, 4);
  s.speakerNotes.textFrame.setText("来源：研究底稿.md 第 4、7 节；01_Zullo_2026_Crayfish_Films.pdf；02_Hromis_2026_Crayfish_Extraction.pdf。\n化学式为碳酸盐被酸化的净离子示意式。论文使用盐酸等酸性体系进行脱矿；具体浓度和操作仅作文献事实，不作为家庭复现实验建议。");
}

// 5 deacetylation
{
  const s = ppt.slides.add();
  s.background.fill = C.paper;
  addHeader(s, "反应二：脱乙酰让几丁质变成壳聚糖");
  addText(s, "几丁质与壳聚糖的差别，藏在高分子链上的“侧基”里。", 72, 172, 1040, 34, { size: 24, color: C.ink });
  addRect(s, 80, 246, 480, 164, "#E4F0F4", 24);
  addText(s, "几丁质", 112, 272, 150, 34, { size: 28, bold: true, color: C.ocean });
  addText(s, "—NHCOCH₃", 112, 326, 240, 44, { size: 34, bold: true, color: C.navy });
  addText(s, "乙酰基较多\n在水中几乎不溶", 370, 300, 140, 60, { size: 19, color: C.muted });
  addArrow(s, 586, 296, 76, 48, C.orange);
  addText(s, "浓碱 + 加热\n脱去一部分乙酰基", 552, 360, 190, 56, { size: 16, bold: true, color: C.orange, align: "center" });
  addRect(s, 710, 246, 480, 164, "#E5F2E8", 24);
  addText(s, "壳聚糖", 742, 272, 150, 34, { size: 28, bold: true, color: C.green });
  addText(s, "—NH₂", 742, 326, 180, 44, { size: 34, bold: true, color: C.navy });
  addText(s, "氨基增多\n可被稀酸质子化并溶解", 980, 300, 170, 60, { size: 19, color: C.muted });
  addRect(s, 140, 468, 1000, 92, C.navy, 18);
  addText(s, "示意式：—NHCOCH₃ + OH⁻  →  —NH₂ + CH₃COO⁻ + H₂O", 182, 494, 920, 38, { size: 27, bold: true, color: C.white, align: "center" });
  addText(s, "真实体系是高分子链上的“部分脱乙酰”，脱乙酰度会影响溶解性、强度和成膜表现。", 102, 594, 1040, 32, { size: 20, color: C.muted, align: "center" });
  addFooter(s, 5);
  s.speakerNotes.textFrame.setText("来源：研究底稿.md 第 4、8 节；01_Zullo_2026_Crayfish_Films.pdf；02_Hromis_2026_Crayfish_Extraction.pdf；03_Kadak_2023_Crayfish_Chitosan.xml。\n说明：本页用结构片段和净变化示意解释脱乙酰，实际反应发生在多糖链上且不会 100% 完全转化。论文中的脱乙酰步骤使用高浓度氢氧化钠和高温，属于实验室高风险操作。");
}

// 6 film formation
{
  const s = ppt.slides.add();
  s.background.fill = C.cream;
  addHeader(s, "反应三：壳聚糖溶液怎样变成一张膜？");
  addText(s, "成膜不是“把粉末压薄”，而是让高分子先溶解，再在干燥时重新排成连续网络。", 72, 170, 1100, 42, { size: 24, color: C.ink });
  addRect(s, 72, 246, 520, 128, "#FBE4D6", 20);
  addText(s, "壳聚糖 + 稀醋酸", 108, 270, 400, 34, { size: 28, bold: true, color: C.navy, align: "center" });
  addText(s, "—NH₂ + H⁺  ⇌  —NH₃⁺", 120, 322, 390, 36, { size: 27, bold: true, color: C.orange, align: "center" });
  addText(s, "氨基被质子化后，聚合物更容易分散在水相里。", 92, 394, 470, 54, { size: 20, color: C.muted, align: "center" });
  addArrow(s, 610, 290, 62, 44, C.teal);
  addRect(s, 706, 246, 480, 128, "#E4F0F4", 20);
  addText(s, "加入甘油", 742, 270, 390, 34, { size: 28, bold: true, color: C.navy, align: "center" });
  addText(s, "增塑剂：让链段更容易移动", 742, 322, 390, 30, { size: 21, color: C.ocean, align: "center" });
  addText(s, "浇铸 → 脱泡 → 干燥", 118, 492, 420, 34, { size: 28, bold: true, color: C.navy });
  addLine(s, 126, 560, 560, 560, C.orange, 4);
  addRect(s, 126, 540, 18, 40, C.orange, 8);
  addRect(s, 332, 540, 18, 40, C.teal, 8);
  addRect(s, 538, 540, 18, 40, C.ocean, 8);
  addText(s, "液体", 94, 602, 80, 24, { size: 18, color: C.muted, align: "center" });
  addText(s, "薄层", 300, 602, 80, 24, { size: 18, color: C.muted, align: "center" });
  addText(s, "连续膜", 506, 602, 100, 24, { size: 18, color: C.muted, align: "center" });
  addRect(s, 706, 448, 480, 170, C.navy, 20);
  addText(s, "薄膜的两面性", 742, 470, 400, 28, { size: 22, bold: true, color: C.coral, align: "center" });
  addText(s, "链间氢键让它有强度\n甘油让它不那么脆\n但水会重新破坏链间作用", 760, 520, 370, 90, { size: 23, color: C.white, align: "center" });
  addFooter(s, 6);
  s.speakerNotes.textFrame.setText("来源：研究底稿.md 第 3、6 节；04_Chitosan_Starch_2022.pdf；05_CornStarch_Chitosan_2021.pdf。\n示意：壳聚糖氨基在酸性水相中质子化，之后随水分蒸发形成连续膜。甘油常作为增塑剂，能提高柔韧性，但可能牺牲强度并增加吸水。");
}

// 7 from film to bag
{
  const s = ppt.slides.add();
  s.background.fill = C.paper;
  addHeader(s, "从薄膜到袋子：真正难的是“用起来”");
  s.images.add({ blob: bagImg, contentType: "image/jpeg", alt: "项目视频中的袋膜和成品袋演示", fit: "cover", position: { left: 770, top: 170, width: 446, height: 430 }, geometry: "roundRect", borderRadius: 20 });
  const stages = [
    ["膜", "厚度均匀\n不易脆裂", C.sand],
    ["裁切", "尺寸一致\n边缘不撕裂", "#FBE4D6"],
    ["连接", "封边可靠\n提手不脱落", "#E4F0F4"],
  ];
  stages.forEach((it, i) => {
    const y = 220 + i * 120;
    addRect(s, 72, y, 148, 82, it[2], 18);
    addText(s, it[0], 88, y + 14, 110, 28, { size: 24, bold: true, color: C.navy, align: "center" });
    addText(s, it[1], 252, y + 8, 380, 52, { size: 21, color: C.ink });
    if (i < stages.length - 1) addArrow(s, 130, y + 90, 32, 22, C.orange);
  });
  addRect(s, 72, 580, 630, 54, C.navy, 16);
  addText(s, "袋子是否合格，要看干态、湿态、接缝和提手。", 96, 592, 580, 28, { size: 21, bold: true, color: C.white, align: "center" });
  addFooter(s, 7);
  s.speakerNotes.textFrame.setText("来源：研究底稿.md 第 1、2 节；项目视频约 4:40–5:30、7:40 以后。\n项目视频展示了可成形的袋膜与手工演示，但没有提供购物袋所需的静载、反复提举、提手撕裂、接缝强度和湿态数据。图中“合格”是测试维度，不是项目已通过认证的结论。");
}

// 8 property mechanism
{
  const s = ppt.slides.add();
  s.background.fill = C.cream;
  addHeader(s, "材料性能的核心矛盾：强度、柔韧性和耐水性");
  addText(s, "同一张膜里，分子间作用力和水分子一直在“拉扯”。", 72, 170, 1040, 34, { size: 24, color: C.ink });
  addRect(s, 70, 240, 340, 240, "#E4F0F4", 24);
  addText(s, "高分子链靠近", 100, 270, 280, 34, { size: 26, bold: true, color: C.ocean, align: "center" });
  addText(s, "氢键\n静电作用\n链缠结", 120, 338, 240, 100, { size: 28, bold: true, color: C.navy, align: "center", valign: "middle" });
  addText(s, "强度 ↑", 130, 448, 220, 28, { size: 22, bold: true, color: C.green, align: "center" });
  addArrow(s, 454, 330, 62, 48, C.orange);
  addRect(s, 552, 240, 340, 240, "#FBE4D6", 24);
  addText(s, "加入甘油", 582, 270, 280, 34, { size: 26, bold: true, color: C.orange, align: "center" });
  addText(s, "链段更容易移动\n伸长率 ↑\n强度可能下降", 582, 338, 280, 108, { size: 27, bold: true, color: C.navy, align: "center", valign: "middle" });
  addArrow(s, 934, 330, 62, 48, C.teal);
  addRect(s, 1032, 240, 174, 240, "#E5F2E8", 24);
  addText(s, "遇水", 1050, 270, 138, 34, { size: 26, bold: true, color: C.green, align: "center" });
  addText(s, "吸水\n溶胀\n变软", 1050, 344, 138, 102, { size: 28, bold: true, color: C.navy, align: "center", valign: "middle" });
  addRect(s, 170, 550, 950, 72, C.navy, 18);
  addText(s, "所以“能成膜”不等于“能装湿东西”，更不等于“能替代所有塑料袋”。", 204, 571, 880, 32, { size: 23, bold: true, color: C.white, align: "center" });
  addFooter(s, 8);
  s.speakerNotes.textFrame.setText("来源：研究底稿.md 第 2、3、8 节；04_Chitosan_Starch_2022.pdf；05_CornStarch_Chitosan_2021.pdf；06_Chitosan_Cellulose_2024.pdf。\n论文和底稿共同指出：甘油常提高柔韧性但降低强度；壳聚糖/淀粉复合膜吸水明显；纤维素可作为增强方向。具体数值因配方与测试条件而异。");
}

// 9 evidence boundary
{
  const s = ppt.slides.add();
  s.background.fill = C.paper;
  addHeader(s, "视频证明了什么，还没有证明什么？");
  s.images.add({ blob: coverImg, contentType: "image/jpeg", alt: "项目视频抽帧联系图", fit: "cover", position: { left: 70, top: 180, width: 520, height: 410 }, geometry: "roundRect", borderRadius: 20, crop: { left: 0.05, top: 0.12, right: 0.05, bottom: 0.08 } });
  addRect(s, 650, 180, 255, 410, "#E5F2E8", 22);
  addText(s, "视频展示", 680, 208, 190, 30, { size: 24, bold: true, color: C.green, align: "center" });
  addText(s, "壳的清洗、粉碎和酸碱处理\n\n白色固体与商品壳聚糖对照\n\n薄膜浇铸、干燥和手工拉扯\n\n袋子可以被拿起来展示", 680, 266, 190, 276, { size: 19, color: C.ink, align: "center", valign: "middle" });
  addRect(s, 930, 180, 280, 410, "#FBE4D6", 22);
  addText(s, "尚未证明", 960, 208, 220, 30, { size: 24, bold: true, color: C.red, align: "center" });
  addText(s, "稳定的提取收率和纯度\n\n干湿态拉伸、接缝和提手强度\n\n标准化生物降解或堆肥结果\n\n食品接触、迁移和卫生合规", 960, 266, 220, 276, { size: 19, color: C.ink, align: "center", valign: "middle" });
  addText(s, "把“看起来像袋子”与“可以作为商品使用”分开，是材料研发最重要的判断。", 100, 620, 1060, 30, { size: 21, bold: true, color: C.navy, align: "center" });
  addFooter(s, 9);
  s.speakerNotes.textFrame.setText("来源：研究底稿.md 第 1、2 节；项目视频全片抽帧核对。\n本页是证据边界说明：视频能够证明演示样品出现了相应操作和形态，但不能替代材料表征、力学测试、降解测试或食品接触合规测试。");
}

// 10 safety and environment
{
  const s = ppt.slides.add();
  s.background.fill = C.navy;
  addText(s, "安全与环保，要一起算", 64, 44, 760, 56, { size: 40, bold: true, color: C.white });
  addText(s, "原料来自废弃物，不会自动让整条路线低风险、低影响。", 66, 112, 900, 34, { size: 22, color: "#D8E7E8" });
  addRect(s, 70, 210, 530, 340, "#183E4D", 24);
  addText(s, "实验安全", 102, 244, 460, 32, { size: 28, bold: true, color: C.coral });
  addText(s, "浓酸：腐蚀、放热、产生气泡\n热浓碱：深部灼伤，不能在家尝试\n废液：酸碱洗涤液不能直接倒下水道\n成品：未经检测，不接触食品或饮水", 104, 310, 460, 180, { size: 23, color: C.white });
  addRect(s, 670, 210, 540, 340, "#234B43", 24);
  addText(s, "环境判断", 702, 244, 480, 32, { size: 28, bold: true, color: "#A8E0C2" });
  addText(s, "看原料：是否利用了餐厨废弃物\n看过程：酸碱、水、热能和添加剂用了多少\n看去向：最终在什么条件下分解\n看替代：是否真的减少了整体环境负担", 704, 310, 470, 180, { size: 23, color: C.white });
  addRect(s, 170, 590, 940, 54, "#F19A62", 16);
  addText(s, "“可降解”是一个需要测试条件的结论，不是材料名字的同义词。", 196, 603, 888, 26, { size: 21, bold: true, color: C.navy, align: "center" });
  addFooter(s, 10, "讲解材料 · 安全边界来自研究底稿与论文条件核对");
  s.speakerNotes.textFrame.setText("来源：研究底稿.md 第 7 节。\n强调：高浓度氢氧化钠脱乙酰属于最高危步骤之一；本项目当前建议路线是先采购规格明确的壳聚糖做低危膜样，不在家庭环境复现虾壳提取。成品膜暂按非食品接触材料处理。");
}

// 11 roadmap
{
  const s = ppt.slides.add();
  s.background.fill = C.cream;
  addHeader(s, "最现实的研发顺序：先验证材料，再追溯原料");
  const phases = [
    ["阶段 1", "先买规格明确的壳聚糖", "用稀醋酸 + 甘油做可重复膜样\n先找到“成膜窗口”", C.ocean],
    ["阶段 2", "再做袋膜性能", "记录厚度、湿态、接缝和提手\n一次只改变一个变量", C.orange],
    ["阶段 3", "最后做虾壳提取", "在有通风柜、耐腐蚀设备和废液流程的实验室\n验证收率、脱乙酰度和批次一致性", C.green],
  ];
  phases.forEach((p, i) => {
    const x = 72 + i * 388;
    addRect(s, x, 220, 340, 260, "#FFFFFF", 24, { style: "solid", fill: C.line, width: 1 });
    addRect(s, x, 220, 340, 14, p[3], 7);
    addText(s, p[0], x + 26, 260, 100, 26, { size: 16, bold: true, color: p[3] });
    addText(s, p[1], x + 26, 302, 286, 64, { size: 27, bold: true, color: C.navy });
    addText(s, p[2], x + 26, 392, 286, 70, { size: 20, color: C.muted });
  });
  addText(s, "今天应该记住的三句话", 72, 548, 320, 28, { size: 18, bold: true, color: C.ocean });
  addText(s, "1  虾壳先被拆成几丁质，再转成可成膜的壳聚糖。\n2  袋子的难点是湿态强度、接缝和一致性。\n3  环保与安全都要用测试和完整流程来证明。", 72, 590, 1050, 82, { size: 23, bold: true, color: C.navy });
  addFooter(s, 11);
  s.speakerNotes.textFrame.setText("来源：研究底稿.md 第 3、6、7 节。\n这是当前底稿给出的最短材料路线：先用采购壳聚糖建立可重复膜样，再进入袋膜性能，最后才投入自提壳聚糖的设备与试剂。");
}

// 12 sources
{
  const s = ppt.slides.add();
  s.background.fill = C.paper;
  addHeader(s, "资料来源与本次讲解的边界");
  addText(s, "项目资料", 72, 180, 300, 30, { size: 22, bold: true, color: C.ocean });
  addText(s, "• 研究底稿.md（2026-09-30）\n• 如何用小龙虾制造可降解塑料袋.mp4\n• 项目视频抽帧：\n  video_0_250.jpg、video_250_500.jpg\n  video_contact_sheet.jpg", 72, 224, 600, 142, { size: 18, color: C.ink });
  addText(s, "开放论文", 72, 394, 300, 30, { size: 22, bold: true, color: C.ocean });
  addText(s, "01 Zullo 2026 · Molecules · 10.3390/molecules31162819\n02 Hromiš 2026 · Gels · 10.3390/gels12080664\n03 Kadak 2023 · International Journal of Biology · 10.30498/ijb.2023.323958.3253\n04 Chitosan–Starch 2022 · Polymers · 10.3390/polym14020278\n05 Cornstarch–Chitosan 2021 · Polymers · 10.3390/polym13244431\n06 Chitosan–Cellulose 2024 · Polymers · 10.3390/polym16050568", 72, 438, 650, 156, { size: 16, color: C.ink });
  addRect(s, 790, 198, 390, 350, C.navy, 24);
  addText(s, "读者提示", 824, 236, 320, 30, { size: 24, bold: true, color: C.coral, align: "center" });
  addText(s, "这套PPT用于科普讲解。\n\n它解释化学逻辑与研发判断，\n不替代实验室 SOP、SDS、\n材料检测或食品接触法规。\n\n任何“可降解”“环保”“安全”\n都应对应具体测试条件。", 824, 288, 320, 210, { size: 22, color: C.white, align: "center", valign: "middle" });
  addText(s, "谢谢", 920, 604, 220, 48, { size: 32, bold: true, color: C.orange, align: "right" });
  addFooter(s, 12, "资料页 · 项目文件与开放获取论文");
  s.speakerNotes.textFrame.setText("来源清单与边界说明。论文文件保存在工作区 papers/ 目录；项目底稿明确指出直接制膜研究的原料物种、配方和测试条件不能直接照搬为商品袋方案。");
}

const draftPath = path.join(buildDir, "draft.pptx");
await (await PresentationFile.exportPptx(ppt)).save(draftPath);

const { finalizePresentation } = await import(pathToFileURL(path.join(skillDir, "container_tools/artifact_tool_utils.mjs")).href);
const result = await finalizePresentation({
  explicitTotalSlideCount: 12,
  requiredNativeTableOwnerSlides: [],
  requiredNativeChartOwnerSlides: [],
  workspaceDir,
  candidatePath: draftPath,
  finalPath,
  pythonExecutable: "C:/Users/s1325/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe",
  integrityValidatorPath: path.join(skillDir, "container_tools/inspect_presentation_package_integrity.py"),
  layoutValidatorPath: path.join(skillDir, "container_tools/inspect_presentation_layout_geometry.py"),
  layoutArgs: ["--expected-slide-size-emu", "12192000,6858000", "--validate-heading-fit"],
  requiredNativeTableOwnerSlides: [],
  fontPolicy: { basis: "design", families: [font] },
  verifyArtifactToolImport: true,
  receiptPath: path.join(stagingDir, "deck.validation.v2.json"),
});
console.log(JSON.stringify({ finalPath, result }, null, 2));
