const PptxGenJS = require("pptxgenjs");

const pptx = new PptxGenJS();
pptx.defineLayout({ name: "WIDE", width: "13.333", height: "7.5" });
pptx.layout = "WIDE";

// --- Color Palette ---
const C = {
  primary: "1a1a2e",
  accent: "c62828",
  light: "f5f5f5",
  white: "ffffff",
  text: "333333",
  gray: "888888",
  tableHeader: "1a1a2e",
  tableAlt: "f0f0f5",
  chart1: "1a1a2e",
  chart2: "c62828",
  chart3: "455a64",
  chart4: "e0a800",
};

// --- Master slide: white bg, thin bottom line ---
pptx.defineSlideMaster({
  title: "MASTER",
  background: { color: C.white },
  objects: [
    {
      line: { color: "cccccc", x: 0.8, y: 7.2, w: 11.7, h: 0 },
    },
    {
      text: "中国摩托车行业行研报告 · 2026",
      options: {
        x: "70%",
        y: 7.05,
        w: "27%",
        h: 0.3,
        fontSize: 8,
        color: C.gray,
        align: "right",
      },
    },
  ],
});

// ========== SLIDE 1: 封面 ==========
{
  const slide = pptx.addSlide();
  slide.background = { color: C.primary };
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: "100%",
    h: "100%",
    fill: { color: C.primary },
  });
  // Accent bar
  slide.addShape(pptx.ShapeType.rect, {
    x: 1.2,
    y: 2.4,
    w: 0.06,
    h: 2.0,
    fill: { color: C.accent },
  });
  slide.addText("中国摩托车行业\n行研报告", {
    x: 1.6,
    y: 2.3,
    w: 8,
    h: 2.2,
    fontSize: 42,
    bold: true,
    color: C.white,
    fontFace: "Microsoft YaHei",
    lineSpacing: 52,
  });
  slide.addText("2025年行业回顾 · 竞争格局 · 趋势展望", {
    x: 1.6,
    y: 4.7,
    w: 8,
    h: 0.6,
    fontSize: 16,
    color: "bbbbbb",
    fontFace: "Microsoft YaHei",
  });
  slide.addText("数据来源：中国摩托车商会 · 海关总署 · 上市公司年报 · 中商情报网\n报告日期：2026年6月", {
    x: 1.6,
    y: 6.2,
    w: 8,
    h: 0.8,
    fontSize: 10,
    color: C.gray,
    fontFace: "Microsoft YaHei",
    lineSpacing: 16,
  });
}

// ========== SLIDE 2: 目录 ==========
{
  const slide = pptx.addSlide();
  slide.addText("目录", { x: 0.8, y: 0.5, w: 4, h: 0.8, fontSize: 28, bold: true, color: C.primary });
  const toc = [
    ["01", "行业概览", "产销规模与关键指标"],
    ["02", "市场结构", "燃油·电动·大排量三分天下"],
    ["03", "品牌竞争格局", "TOP10 销量排名与集中度"],
    ["04", "大排量市场", "增长最快的赛道"],
    ["05", "电动摩托车", "绿源反超雅迪，极核崛起"],
    ["06", "上市公司表现", "春风·隆鑫·钱江业绩对比"],
    ["07", "出口市场", "从制造出海到品牌出海"],
    ["08", "趋势与展望", "高端化·电动化·精品化"],
  ];
  toc.forEach((item, i) => {
    const y = 1.8 + i * 0.65;
    slide.addText(item[0], { x: 1.2, y: y, w: 0.8, h: 0.5, fontSize: 20, bold: true, color: C.accent });
    slide.addText(item[1], { x: 2.2, y: y, w: 5, h: 0.5, fontSize: 16, bold: true, color: C.text });
    slide.addText(item[2], { x: 2.2, y: y + 0.28, w: 6, h: 0.35, fontSize: 11, color: C.gray });
  });
}

// ========== SLIDE 3: 行业概览 ==========
{
  const slide = pptx.addSlide();
  slide.addText("01 行业概览", { x: 0.8, y: 0.5, w: 6, h: 0.8, fontSize: 24, bold: true, color: C.primary });
  slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.2, w: 1.2, h: 0.04, fill: { color: C.accent } });

  // Key metrics - 4 big numbers
  const metrics = [
    ["2,210.93万辆", "2025年产量", "+10.69%"],
    ["2,196.77万辆", "2025年销量", "+10.25%"],
    ["1,336.57万辆", "整车出口量", "+21.33%"],
    ["88.5亿美元", "出口金额", "+26.78%"],
  ];
  metrics.forEach((m, i) => {
    const x = 0.8 + i * 3.1;
    slide.addShape(pptx.ShapeType.rect, { x: x, y: 1.6, w: 2.8, h: 1.6, fill: { color: C.light } });
    slide.addText(m[0], { x: x + 0.15, y: 1.7, w: 2.5, h: 0.7, fontSize: 22, bold: true, color: C.accent });
    slide.addText(m[1], { x: x + 0.15, y: 2.35, w: 2.5, h: 0.4, fontSize: 11, color: C.text });
    slide.addText(m[2], { x: x + 0.15, y: 2.7, w: 2.5, h: 0.35, fontSize: 11, color: C.gray });
  });

  // Bottom section: domestic vs export
  const rows = [
    ["", "2025年", "同比变化", "说明"],
    ["国内销售", "860.2万辆", "-3.45%", "内销承压，通勤需求下滑"],
    ["整车出口", "1,336.57万辆", "+21.33%", "出口占比超60%，核心增长引擎"],
    ["规上企业营收", "1,462亿元", "+12.80%", "前11月数据，利润增37.73%"],
    ["全产业链规模", "≈万亿元", "—", "含摩旅、装备、保险等衍生消费"],
  ];
  slide.addTable(rows, {
    x: 0.8,
    y: 3.6,
    w: 11.7,
    h: 3.0,
    fontSize: 11,
    fontFace: "Microsoft YaHei",
    border: { type: "none" },
    colW: [2.5, 3.0, 2.8, 3.4],
    rowH: [0.45, 0.5, 0.5, 0.5, 0.5],
    color: C.text,
    autoPage: false,
  });
  // Style header row
  rows.forEach((_, ri) => {
    if (ri === 0) {
      slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 3.6, w: 11.7, h: 0.45, fill: { color: C.tableHeader } });
      slide.addText("指标", { x: 0.95, y: 3.62, w: 2.5, h: 0.4, fontSize: 10, bold: true, color: C.white });
      slide.addText("2025年", { x: 3.35, y: 3.62, w: 3.0, h: 0.4, fontSize: 10, bold: true, color: C.white });
      slide.addText("同比变化", { x: 6.35, y: 3.62, w: 2.8, h: 0.4, fontSize: 10, bold: true, color: C.white });
      slide.addText("说明", { x: 9.15, y: 3.62, w: 3.4, h: 0.4, fontSize: 10, bold: true, color: C.white });
    }
  });
}

// ========== SLIDE 4: 市场结构 ==========
{
  const slide = pptx.addSlide();
  slide.addText("02 市场结构", { x: 0.8, y: 0.5, w: 6, h: 0.8, fontSize: 24, bold: true, color: C.primary });
  slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.2, w: 1.2, h: 0.04, fill: { color: C.accent } });

  // Left: pie-like breakdown
  const segs = [
    ["燃油摩托车", "1,846万辆", "84.0%", "通勤+大排量休闲"],
    ["电动摩托车", "351万辆", "16.0%", "增长+1.14%"],
  ];
  segs.forEach((s, i) => {
    const y = 1.6 + i * 1.5;
    slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: y, w: 5.5, h: 1.2, fill: { color: C.light } });
    slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: y, w: 0.06, h: 1.2, fill: { color: i === 0 ? C.chart1 : C.chart3 } });
    slide.addText(s[0], { x: 1.1, y: y + 0.05, w: 3, h: 0.4, fontSize: 15, bold: true, color: C.text });
    slide.addText(s[1] + " · 占比 " + s[2], { x: 1.1, y: y + 0.45, w: 4, h: 0.35, fontSize: 12, color: C.accent, bold: true });
    slide.addText(s[3], { x: 1.1, y: y + 0.8, w: 4, h: 0.3, fontSize: 10, color: C.gray });
  });

  // Right: 大排量 callout
  slide.addShape(pptx.ShapeType.rect, { x: 6.8, y: 1.6, w: 5.7, h: 1.2, fill: { color: "fff3e0" } });
  slide.addShape(pptx.ShapeType.rect, { x: 6.8, y: 1.6, w: 0.06, h: 1.2, fill: { color: C.chart4 } });
  slide.addText("250cc+ 大排量休闲娱乐摩托", { x: 7.1, y: 1.65, w: 5.2, h: 0.4, fontSize: 14, bold: true, color: C.text });
  slide.addText("产销 95.37万辆 / 95.23万辆", { x: 7.1, y: 2.1, w: 5.2, h: 0.35, fontSize: 13, bold: true, color: C.accent });
  slide.addText("同比 +23.3% / +25.87%（增速远超行业均值）", { x: 7.1, y: 2.5, w: 5.2, h: 0.25, fontSize: 10, color: C.gray });

  // Bottom table: segment breakdown
  const rows2 = [
    ["排量段", "销量(万辆)", "同比", "特征"],
    ["250-400cc", "54.1", "+18.5%", "入门级，国产主力区间"],
    ["400-500cc", "11.7", "-9.4%", "中间段承压"],
    ["500-800cc", "26.8", "+42.3%", "增速最快，品牌向上核心"],
    ["800cc以上", "2.6", "+15.4%", "以进口为主"],
  ];
  slide.addTable(rows2, {
    x: 0.8,
    y: 4.3,
    w: 11.7,
    h: 2.5,
    fontSize: 11,
    border: { type: "none" },
    colW: [2.5, 3.0, 2.8, 3.4],
    rowH: [0.4, 0.45, 0.45, 0.45, 0.45],
    color: C.text,
  });
  slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 4.3, w: 11.7, h: 0.4, fill: { color: C.tableHeader } });
  ["排量段", "销量(万辆)", "同比", "特征"].forEach((h, hi) => {
    slide.addText(h, { x: 0.95 + hi * 2.8, y: 4.32, w: 2.5, h: 0.35, fontSize: 10, bold: true, color: C.white });
  });
}

// ========== SLIDE 5: 品牌竞争格局 ==========
{
  const slide = pptx.addSlide();
  slide.addText("03 品牌竞争格局", { x: 0.8, y: 0.5, w: 6, h: 0.8, fontSize: 24, bold: true, color: C.primary });
  slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.2, w: 1.2, h: 0.04, fill: { color: C.accent } });

  const brands = [
    ["1", "大长江（豪爵）", "268.67", "连续23年销冠"],
    ["2", "隆鑫（隆鑫·无极）", "157.83", "出口量行业第一"],
    ["3", "宗申（宗申·赛科龙）", "113.92", "含三轮车"],
    ["4", "广东大冶（升仕·启典）", "110.91", "双品牌高低通吃"],
    ["5", "新大洲本田", "100.31", "合资首破百万"],
    ["6", "广州豪进", "73.66", "—"],
    ["7", "重庆银翔", "72.91", "三轮车大户"],
    ["8", "洛阳北方（大阳）", "61.02", "内销占比59%"],
    ["9", "五羊-本田", "59.48", "内销占比74%"],
    ["10", "江门珠峰", "52.91", "—"],
  ];

  // Simple horizontal bars
  const maxW = 268.67;
  brands.forEach((b, i) => {
    const y = 1.6 + i * 0.52;
    const barW = (b[2] / maxW) * 7.5;
    slide.addText(b[0], { x: 0.8, y: y, w: 0.4, h: 0.4, fontSize: 10, bold: true, color: C.gray });
    slide.addText(b[1], { x: 1.3, y: y, w: 2.4, h: 0.4, fontSize: 10, color: C.text });
    slide.addShape(pptx.ShapeType.rect, { x: 3.8, y: y + 0.08, w: barW, h: 0.25, fill: { color: i < 3 ? C.accent : C.chart3 } });
    slide.addText(b[2] + "万辆", { x: 3.8 + barW + 0.15, y: y, w: 1.2, h: 0.4, fontSize: 9, bold: true, color: C.text });
    slide.addText(b[3], { x: 10.5, y: y, w: 2.5, h: 0.4, fontSize: 9, color: C.gray });
  });

  // Market concentration note
  slide.addText("行业集中度：CR3 ≈ 24.6%  ·  CR10 ≈ 48.8%", {
    x: 0.8, y: 6.9, w: 6, h: 0.35, fontSize: 11, bold: true, color: C.text,
  });
}

// ========== SLIDE 6: 大排量市场 ==========
{
  const slide = pptx.addSlide();
  slide.addText("04 大排量市场（≥250cc）", { x: 0.8, y: 0.5, w: 7, h: 0.8, fontSize: 24, bold: true, color: C.primary });
  slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.2, w: 1.2, h: 0.04, fill: { color: C.accent } });

  slide.addText("2025年：产销 95.37万辆 / 95.23万辆  ·  同比 +23.3% / +25.87%  ·  内销41.9万 / 出口53.3万", {
    x: 0.8, y: 1.4, w: 12, h: 0.4, fontSize: 11, color: C.gray,
  });

  // Left: Top5 brands
  slide.addText("大排量品牌 TOP5", { x: 0.8, y: 2.0, w: 4, h: 0.5, fontSize: 14, bold: true, color: C.text });
  const bigBikes = [
    ["春风动力", "11.22万辆", "450SR 单车销冠"],
    ["无极 VOGE", "7.02万辆", "CU525/DS525X"],
    ["钱江 QJMOTOR", "6.48万辆", "排量覆盖50-1200cc"],
    ["奔达", "6.32万辆", "巡航车市占率30%"],
    ["本田", "3.55万辆", "合资品牌偏保守"],
  ];
  bigBikes.forEach((b, i) => {
    const y = 2.5 + i * 0.65;
    slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: y, w: 5.3, h: 0.55, fill: { color: i % 2 === 0 ? C.light : C.white } });
    slide.addText(b[0], { x: 1.0, y: y + 0.05, w: 2.0, h: 0.45, fontSize: 12, bold: true, color: C.text });
    slide.addText(b[1], { x: 3.0, y: y + 0.05, w: 1.8, h: 0.45, fontSize: 12, bold: true, color: C.accent });
    slide.addText(b[2], { x: 5.0, y: y + 0.1, w: 1.8, h: 0.35, fontSize: 9, color: C.gray });
    if (i === 0) {
      slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: y, w: 0.05, h: 0.55, fill: { color: C.accent } });
    }
  });

  // Right: segment share
  slide.addText("细分品类市占率（示例）", { x: 7.0, y: 2.0, w: 5, h: 0.5, fontSize: 14, bold: true, color: C.text });
  // 仿赛车
  slide.addText("仿赛车市场", { x: 7.0, y: 2.6, w: 3, h: 0.35, fontSize: 11, bold: true, color: C.text });
  slide.addText("春风44% · 钱江24% · 铃木15% · 珠峰14%", { x: 7.0, y: 2.9, w: 6, h: 0.3, fontSize: 10, color: C.gray });
  // 巡航车
  slide.addText("巡航车市场", { x: 7.0, y: 3.4, w: 3, h: 0.35, fontSize: 11, bold: true, color: C.text });
  slide.addText("奔达30% · 钱江21% · 无极20% · 春风17% · 豪爵8%", { x: 7.0, y: 3.7, w: 6, h: 0.3, fontSize: 10, color: C.gray });

  // Key insight
  slide.addShape(pptx.ShapeType.rect, { x: 7.0, y: 4.5, w: 5.7, h: 2.2, fill: { color: C.light } });
  slide.addText("核心洞察", { x: 7.3, y: 4.65, w: 5, h: 0.4, fontSize: 13, bold: true, color: C.accent });
  slide.addText(
    "• 500-800cc 增速+42.3%，是最热赛道\n• 大排量头部品牌与总销量排名差异显著\n  豪爵总销第一，但大排未进前五\n• 出口占比56%，且增速(+48.5%)远超内销(+5.4%)\n• 春风450SR单车型3.7万辆，国产250cc+销冠\n• 张雪机车2024年成立，当年产值达7.5亿元",
    { x: 7.3, y: 5.1, w: 5.2, h: 1.5, fontSize: 10, color: C.text, lineSpacing: 16 }
  );
}

// ========== SLIDE 7: 电动摩托车 ==========
{
  const slide = pptx.addSlide();
  slide.addText("05 电动摩托车", { x: 0.8, y: 0.5, w: 6, h: 0.8, fontSize: 24, bold: true, color: C.primary });
  slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.2, w: 1.2, h: 0.04, fill: { color: C.accent } });

  slide.addText("2025年：产销 361万辆 / 351万辆  ·  2026年1-5月：149万辆，+5.2%  ·  3月占国内摩托车总销量 43.69%", {
    x: 0.8, y: 1.4, w: 12, h: 0.4, fontSize: 11, color: C.gray,
  });

  // 2025 vs 2026 comparison
  const data2025 = [
    ["排名", "品牌", "2025全年"],
    ["1", "雅迪", "92.11万辆"],
    ["2", "绿源", "59.15万辆"],
    ["3", "宗申", "46.51万辆"],
  ];
  const data2026 = [
    ["排名", "品牌", "2026年1-5月"],
    ["1", "绿源 ▲", "30.65万辆"],
    ["2", "雅迪 ▼", "25.78万辆"],
    ["3", "春风极核 ▲", "21.42万辆"],
    ["4", "新日", "16.89万辆"],
    ["5", "宗申", "16.27万辆"],
  ];

  slide.addText("2025年 TOP3", { x: 0.8, y: 2.0, w: 3, h: 0.4, fontSize: 13, bold: true, color: C.text });
  slide.addTable(data2025, {
    x: 0.8, y: 2.5, w: 5.2, h: 2.2, fontSize: 10, border: { type: "none" }, colW: [0.8, 2.0, 2.4],
  });
  slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 2.5, w: 5.2, h: 0.4, fill: { color: C.tableHeader } });
  ["排名", "品牌", "2025全年"].forEach((h, hi) => {
    slide.addText(h, { x: 0.95 + hi * 2.0, y: 2.52, w: 2.0, h: 0.35, fontSize: 10, bold: true, color: C.white });
  });

  slide.addText("2026年1-5月 TOP5", { x: 7.3, y: 2.0, w: 4, h: 0.4, fontSize: 13, bold: true, color: C.text });
  slide.addTable(data2026, {
    x: 7.3, y: 2.5, w: 5.2, h: 2.6, fontSize: 10, border: { type: "none" }, colW: [0.8, 2.2, 2.2],
  });
  slide.addShape(pptx.ShapeType.rect, { x: 7.3, y: 2.5, w: 5.2, h: 0.4, fill: { color: C.tableHeader } });
  ["排名", "品牌", "2026年1-5月"].forEach((h, hi) => {
    slide.addText(h, { x: 7.45 + hi * 2.0, y: 2.52, w: 2.0, h: 0.35, fontSize: 10, bold: true, color: C.white });
  });

  // Key insight box
  slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 5.2, w: 11.7, h: 1.7, fill: { color: C.light } });
  slide.addText("格局之变", { x: 1.1, y: 5.3, w: 3, h: 0.4, fontSize: 14, bold: true, color: C.accent });
  slide.addText(
    "● 绿源首次反超雅迪登顶电摩冠军 — 2026年1-5月领先近5万辆\n● 春风极核(ZEEHO)成最大黑马 — 从零到行业第三仅用两年，代表燃油摩企电动化转型成功\n● 行业格局从「雅迪独大」转向「绿源·雅迪·极核三强争霸」\n● 新国标2025年底实施，加速行业洗牌 — TOP10市占率已达93.4%",
    { x: 1.1, y: 5.7, w: 11.2, h: 1.1, fontSize: 10, color: C.text, lineSpacing: 18 }
  );
}

// ========== SLIDE 8: 上市公司表现 ==========
{
  const slide = pptx.addSlide();
  slide.addText("06 上市公司表现", { x: 0.8, y: 0.5, w: 6, h: 0.8, fontSize: 24, bold: true, color: C.primary });
  slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.2, w: 1.2, h: 0.04, fill: { color: C.accent } });

  // Three company cards
  const companies = [
    {
      name: "春风动力 603129",
      rank: "行业第1",
      revenue: "197.46亿",
      profit: "17.34亿",
      color: C.accent,
      items: ["全地形车 96.08亿 (48.7%)", "燃油摩托 64.71亿 (32.8%)", "电动两轮车 19.12亿 (9.7%)"],
    },
    {
      name: "隆鑫通用 603766",
      rank: "行业第2",
      revenue: "191.35亿",
      profit: "16.48亿",
      color: C.chart3,
      items: ["摩托车及发动机 141.27亿 (73.8%)", "无极系列收入 39.54亿 (+25.4%)", "通用机械 43.55亿 (22.8%)"],
    },
    {
      name: "钱江摩托 000913",
      rank: "行业第4",
      revenue: "≈59亿(估)",
      profit: "≈4.3亿(估)",
      color: "888888",
      items: ["背靠吉利科技集团", "贝纳利+QJMOTOR双品牌", "正推进精品化转型"],
    },
  ];
  companies.forEach((c, i) => {
    const x = 0.8 + i * 4.0;
    slide.addShape(pptx.ShapeType.rect, { x: x, y: 1.6, w: 3.7, h: 4.6, fill: { color: C.light } });
    slide.addShape(pptx.ShapeType.rect, { x: x, y: 1.6, w: 3.7, h: 0.06, fill: { color: c.color } });
    slide.addText(c.name, { x: x + 0.2, y: 1.8, w: 3.3, h: 0.5, fontSize: 14, bold: true, color: C.text });
    slide.addText(c.rank, { x: x + 0.2, y: 2.3, w: 3.3, h: 0.35, fontSize: 10, color: C.gray });
    slide.addText(c.revenue, { x: x + 0.2, y: 2.85, w: 3.3, h: 0.55, fontSize: 26, bold: true, color: c.color });
    slide.addText("营收", { x: x + 0.2, y: 3.3, w: 1.5, h: 0.3, fontSize: 10, color: C.gray });
    slide.addText(c.profit, { x: x + 1.8, y: 2.85, w: 1.7, h: 0.55, fontSize: 26, bold: true, color: c.color });
    slide.addText("净利润", { x: x + 1.8, y: 3.3, w: 1.5, h: 0.3, fontSize: 10, color: C.gray });
    c.items.forEach((item, j) => {
      slide.addText("• " + item, { x: x + 0.2, y: 3.8 + j * 0.4, w: 3.3, h: 0.35, fontSize: 9.5, color: C.text });
    });
  });

  slide.addText(
    "数据来源：各公司2025年年报、三季报。钱江摩托全年数据基于前三季度推算，完整年报尚未发布。",
    { x: 0.8, y: 6.6, w: 12, h: 0.3, fontSize: 8, color: C.gray }
  );
}

// ========== SLIDE 9: 出口市场 ==========
{
  const slide = pptx.addSlide();
  slide.addText("07 出口市场", { x: 0.8, y: 0.5, w: 6, h: 0.8, fontSize: 24, bold: true, color: C.primary });
  slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.2, w: 1.2, h: 0.04, fill: { color: C.accent } });

  // Big numbers
  const exMetrics = [
    ["1,336.57万辆", "2025年整车出口量", "+21.33%"],
    ["88.5亿美元", "出口金额", "+26.78%"],
    ["808万辆", "2026年1-2月出口", "+32.5%"],
    ["53.3万辆", "大排量出口", "+48.5%"],
  ];
  exMetrics.forEach((m, i) => {
    const x = 0.8 + i * 3.1;
    slide.addShape(pptx.ShapeType.rect, { x: x, y: 1.6, w: 2.8, h: 1.4, fill: { color: C.light } });
    slide.addText(m[0], { x: x + 0.15, y: 1.7, w: 2.5, h: 0.55, fontSize: 20, bold: true, color: C.accent });
    slide.addText(m[1], { x: x + 0.15, y: 2.25, w: 2.5, h: 0.35, fontSize: 10, color: C.text });
    slide.addText(m[2], { x: x + 0.15, y: 2.6, w: 2.5, h: 0.3, fontSize: 10, color: C.gray });
  });

  // Key text
  slide.addText("出口占比已超总销量 60%，成为行业增长主引擎", {
    x: 0.8, y: 3.4, w: 8, h: 0.4, fontSize: 13, bold: true, color: C.text,
  });

  // Trend points
  slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 4.0, w: 5.7, h: 2.8, fill: { color: C.light } });
  slide.addText("出口趋势", { x: 1.1, y: 4.1, w: 5, h: 0.4, fontSize: 13, bold: true, color: C.accent });
  slide.addText(
    "● 从\"通路车出口\"到\"品牌出海\"：大排量玩乐车型\n  在欧美渗透率从2020年趋近零升至10%+\n● 春风、隆鑫、钱江等自主品牌加速海外渠道建设\n● 2026年大排量出口预计 71.8万辆 (+33%)\n● 摩托赛事成为品牌国际化的新名片\n  （如张雪机车参与国际赛事提升品牌声量）",
    { x: 1.1, y: 4.55, w: 5.2, h: 2.1, fontSize: 10, color: C.text, lineSpacing: 18 }
  );

  // Right side
  slide.addShape(pptx.ShapeType.rect, { x: 7.0, y: 4.0, w: 5.7, h: 2.8, fill: { color: C.light } });
  slide.addText("内销挑战", { x: 7.3, y: 4.1, w: 5, h: 0.4, fontSize: 13, bold: true, color: C.chart3 });
  slide.addText(
    "● 国内燃油摩托销售 514.49万辆，同比 -6.18%\n● 通勤类小排量需求持续萎缩\n● \"禁限摩\"政策在部分城市仍制约消费释放\n● 消费升级趋势明确：大排量内销 +5.4%\n● 低端产能过剩，价格战压力犹存\n● 中国摩托车商会2026年6月发布倡议书\n  呼吁抵制内卷、高质量发展",
    { x: 7.3, y: 4.55, w: 5.2, h: 2.1, fontSize: 10, color: C.text, lineSpacing: 18 }
  );
}

// ========== SLIDE 10: 趋势与展望 ==========
{
  const slide = pptx.addSlide();
  slide.addText("08 趋势与展望", { x: 0.8, y: 0.5, w: 6, h: 0.8, fontSize: 24, bold: true, color: C.primary });
  slide.addShape(pptx.ShapeType.rect, { x: 0.8, y: 1.2, w: 1.2, h: 0.04, fill: { color: C.accent } });

  const trends = [
    {
      title: "高端化",
      desc: "500-800cc排量增速+42.3%，自主品牌产品力持续提升，CR3在中大排市场集中度达46.6%，品牌溢价能力增强。",
    },
    {
      title: "电动化加速",
      desc: "新国标落地推动行业洗牌，电摩占国内销量比例突破40%。绿源反超雅迪、春风极核异军突起，竞争格局重塑。",
    },
    {
      title: "全球化升级",
      desc: "从\"制造出海\"走向\"品牌出海\"，2026年大排量出口预计+33%。赛事营销、海外建厂成为品牌国际化新路径。",
    },
    {
      title: "生态扩容",
      desc: "全产业链规模迈向万亿元。骑行装备市场400亿（+20-30%）、摩旅、赛事、社群等衍生消费持续扩大产业边界。",
    },
    {
      title: "产业出清",
      desc: "行业从价格战转向价值战，低端产能加速出清。摩托车商会呼吁行业自律，利润增速(+37.7%)远超营收增速(+12.8%)。",
    },
    {
      title: "跨界入局",
      desc: "华为、小鹏等科技企业进入摩托车赛道，智能化（车联网、智能驾驶辅助）成为产品差异化新方向。",
    },
  ];

  trends.forEach((t, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.8 + col * 4.0;
    const y = 1.6 + row * 2.8;
    slide.addShape(pptx.ShapeType.rect, { x: x, y: y, w: 3.7, h: 2.5, fill: { color: C.light } });
    slide.addShape(pptx.ShapeType.rect, { x: x, y: y, w: 0.06, h: 2.5, fill: { color: C.accent } });
    slide.addText(t.title, { x: x + 0.3, y: y + 0.2, w: 3.2, h: 0.5, fontSize: 16, bold: true, color: C.text });
    slide.addText(t.desc, { x: x + 0.3, y: y + 0.8, w: 3.2, h: 1.5, fontSize: 10, color: C.text, lineSpacing: 20 });
  });
}

// ========== SLIDE 11: 结束页 ==========
{
  const slide = pptx.addSlide();
  slide.background = { color: C.primary };
  slide.addText("谢谢", {
    x: 0, y: 2.5, w: "100%", h: 1.5,
    fontSize: 48, bold: true, color: C.white,
    align: "center",
  });
  slide.addText("中国摩托车行业行研报告 · 2026", {
    x: 0, y: 4.2, w: "100%", h: 0.6,
    fontSize: 14, color: "bbbbbb", align: "center",
  });
  slide.addText("数据来源：中国摩托车商会 · 海关总署 · 中商情报网 · 上市公司年报\n声明：本报告仅供参考，不构成投资建议", {
    x: 0, y: 5.8, w: "100%", h: 0.6,
    fontSize: 10, color: C.gray, align: "center", lineSpacing: 16,
  });
}

// --- Save ---
const outPath = "D:\\ClaudeCode projects\\2026.6.25\\中国摩托车行业行研报告_2026.pptx";
pptx.writeFile({ fileName: outPath }).then(() => {
  console.log("PPT saved to: " + outPath);
});
