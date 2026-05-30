import fs from "node:fs";
import path from "node:path";
import pptxgen from "pptxgenjs";
import {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";

const OUT = process.cwd();
const IMG_DIR = path.join(OUT, "images");
fs.mkdirSync(IMG_DIR, { recursive: true });

const C = {
  green: "0B6B43",
  green2: "138A5B",
  mint: "EAF6EF",
  mint2: "F5FBF7",
  dark: "1F2933",
  gray: "667085",
  light: "F3F6F4",
  line: "D9E6DE",
  amber: "C89211",
  red: "C2410C",
  blue: "2563EB",
  white: "FFFFFF",
};

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "理赔智能体项目组";
pptx.company = "客户事业部";
pptx.subject = "理赔智能体项目成果汇报";
pptx.title = "理赔智能体项目成果汇报";
pptx.lang = "zh-CN";
pptx.theme = {
  headFontFace: "PingFang SC",
  bodyFontFace: "PingFang SC",
  lang: "zh-CN",
};
pptx.defineLayout({ name: "CUSTOM", width: 13.333, height: 7.5 });
pptx.layout = "CUSTOM";
pptx.defineSlideMaster({
  title: "MASTER",
  background: { color: C.white },
  objects: [
    { line: { x: 0.45, y: 7.08, w: 12.45, h: 0, line: { color: C.line, width: 0.7 } } },
  ],
  slideNumber: {
    x: 12.35,
    y: 7.13,
    color: C.gray,
    fontFace: "PingFang SC",
    fontSize: 7,
  },
});

const slides = [
  {
    section: "封面",
    title: "理赔智能体项目成果汇报",
    headline: "从智能问答走向理赔服务执行入口",
    note: "面向客户事业部高层 | 健康险理赔智能服务 / 数生理赔智能服务",
  },
  {
    section: "整体数据情况",
    title: "理赔智能服务已形成可持续运营的数据看板体系",
    headline: "用统一指标衡量规模、体验与质量，支撑后续精细化运营",
    bullets: ["会话量、覆盖率、解决率、满意度", "转人工率、响应时长、问题闭环效率", "本页预留真实数据图表区域"],
    chart: "整体经营看板：趋势折线 + 核心指标卡片",
  },
  {
    section: "健康险理赔智能服务",
    title: "质量指标是健康险智能服务运营的核心抓手",
    headline: "一级差错率与整体差错率将作为运营闭环的主监控指标",
    bullets: ["预留一级差错率和整体差错率趋势图", "结合抽检样本追踪问题归因", "用上线后复测验证优化效果"],
    chart: "双折线趋势图 + 质量运营闭环",
  },
  {
    section: "健康险理赔智能服务",
    title: "稳定的日常运营机制，是质量持续提升的关键",
    headline: "月 6000+ AI 会话，约 10% 抽检，持续推进 badcase 复返与能力优化",
    bullets: ["线上 AI 会话抽检：月会话量 6000+，抽检比例约 10%", "badcase 复返与优化：持续发现问题、定位原因、提升能力", "形成“抽检-归因-优化-验证”的运营飞轮"],
    chart: "运营飞轮 + 抽检漏斗",
    metrics: ["6000+|月 AI 会话量", "10%|会话抽检比例", "闭环|Badcase 复返优化"],
  },
  {
    section: "健康险理赔智能服务",
    title: "当前瓶颈集中在复杂意图、场景覆盖和多模态识别质量",
    headline: "问题类型清晰，后续可围绕识别、覆盖、接口联动分层优化",
    bullets: ["意图识别：客户描述不清、多意图并发、上下文不完整", "场景覆盖：操作类场景、增值服务咨询、工单接口对接", "其他问题：OCR 识别错误、ASR 语音识别偏差"],
    chart: "问题优先级矩阵：影响程度 x 解决优先级",
  },
  {
    section: "健康险理赔智能服务",
    title: "健康险将从 RAG 问答架构升级为 Agent 服务架构",
    headline: "从“检索回答”升级为“理解、决策、执行、复盘”",
    bullets: ["架构方向：RAG 到 Agent", "优化重点：意图识别、多意图处理、操作场景接入、接口联动", "目标：提升复杂问题处理能力，降低人工运营成本"],
    chart: "架构演进图：RAG → Agent",
  },
  {
    section: "数生理赔智能服务",
    title: "数生理赔智能服务已具备规模化运营基础",
    headline: "覆盖情况、回复准确率、响应时长将构成下一阶段规模化评估框架",
    bullets: ["预留覆盖情况、回复准确率、响应时长图表", "建议同时展示产品覆盖数、场景覆盖率与转人工率", "通过持续监控识别体验瓶颈"],
    chart: "三联图：覆盖率 / 准确率 / 响应时长",
  },
  {
    section: "数生理赔智能服务",
    title: "通过 Skill 能力，智能体可从“回答问题”升级为“完成操作”",
    headline: "操作类场景突破使理赔智能服务不再局限于咨询解答",
    bullets: ["操作类场景通过 Skill 实现标准化执行", "智能体可承接引导、查询、办理、触发流程", "客户获得更完整的自助理赔体验"],
    chart: "能力跃迁图：解答型 → 助理型 → 执行型",
  },
  {
    section: "数生理赔智能服务",
    title: "下一阶段目标覆盖 90% 理赔场景，支撑数生近千款产品",
    headline: "以高覆盖率和产品规模化接入为核心增长目标",
    bullets: ["AI 能力提升：覆盖 90% 的理赔场景", "产品覆盖：除 3C 外，逐步覆盖数生近千款产品", "场景能力：咨询、查询、操作、工单、转人工协同"],
    chart: "三阶段路线图 + 场景覆盖热力图",
    metrics: ["90%|目标理赔场景覆盖", "近千款|目标产品覆盖", "全链路|咨询/查询/操作协同"],
  },
  {
    section: "渠道赋能",
    title: "理赔智能体将成为多渠道统一调用的智能服务能力",
    headline: "服务大厅已对接，App 对接中，能力将沉淀为统一理赔智能底座",
    bullets: ["服务大厅已对接：理赔场景调用", "App 对接中：后续固定入口、固定人群调用理赔智能体", "多渠道复用同一套理赔智能服务能力"],
    chart: "Hub-and-Spoke 渠道赋能架构图",
  },
  {
    section: "数生理赔智能服务",
    title: "平台能力和响应时效是规模化运营的主要约束",
    headline: "需补齐执行日志、结构化输出与链路前置能力",
    bullets: ["灵犀平台：未打印 Agent 执行日志，难以问题复盘", "灵犀平台：不支持自定义结构化输出，影响整体场景运营", "响应时效：情绪/敏感词识别前置到 App，调用前确定保单/赔案"],
    chart: "问题-方案映射图 + 当前链路 vs 优化后链路",
  },
];

function addText(slide, text, x, y, w, h, opts = {}) {
  slide.addText(text, {
    x,
    y,
    w,
    h,
    margin: 0,
    breakLine: false,
    fit: "shrink",
    fontFace: "PingFang SC",
    color: opts.color || C.dark,
    fontSize: opts.size || 14,
    bold: !!opts.bold,
    valign: opts.valign || "mid",
    align: opts.align || "left",
    ...opts,
  });
}

function addHeader(slide, section, title, headline) {
  addText(slide, section, 0.55, 0.28, 3.2, 0.22, { size: 7.5, color: C.green2, bold: true, charSpace: 0.6 });
  addText(slide, title, 0.55, 0.58, 8.9, 0.52, { size: 21, bold: true, color: C.dark });
  slide.addShape(pptx.ShapeType.line, { x: 0.55, y: 1.2, w: 1.15, h: 0, line: { color: C.green, width: 2 } });
  if (headline) addText(slide, headline, 0.55, 1.34, 8.9, 0.35, { size: 11.2, color: C.gray });
}

function addTag(slide, text, x, y, w, color = C.green) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h: 0.34,
    rectRadius: 0.08,
    fill: { color: color === C.green ? C.mint : "FFF7ED" },
    line: { color, width: 0.8 },
  });
  addText(slide, text, x + 0.11, y + 0.065, w - 0.22, 0.16, { size: 7.6, color, bold: true, align: "center" });
}

function addMetric(slide, raw, x, y, w) {
  const [num, label] = raw.split("|");
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h: 0.9,
    rectRadius: 0.08,
    fill: { color: C.mint2 },
    line: { color: C.line, width: 0.8 },
  });
  addText(slide, num, x + 0.15, y + 0.14, w - 0.3, 0.28, { size: 22, bold: true, color: C.green, align: "center" });
  addText(slide, label, x + 0.15, y + 0.52, w - 0.3, 0.18, { size: 8, color: C.gray, align: "center" });
}

function addChartPlaceholder(slide, title, x, y, w, h) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    rectRadius: 0.08,
    fill: { color: C.mint2 },
    line: { color: C.line, width: 1, dash: "dash" },
  });
  addText(slide, "图表占位", x + 0.25, y + 0.22, 1.2, 0.2, { size: 8, bold: true, color: C.green });
  addText(slide, title, x + 0.25, y + 0.58, w - 0.5, 0.3, { size: 12.8, bold: true, color: C.dark });
  for (let i = 0; i < 4; i++) {
    const barW = (w - 1.1) / 5;
    slide.addShape(pptx.ShapeType.rect, {
      x: x + 0.42 + i * (barW + 0.22),
      y: y + h - 0.55 - i * 0.22,
      w: barW,
      h: 0.38 + i * 0.22,
      fill: { color: i % 2 ? C.green2 : C.green, transparency: 16 },
      line: { color: i % 2 ? C.green2 : C.green },
    });
  }
}

function addBullets(slide, bullets, x, y, w) {
  bullets.forEach((b, idx) => {
    slide.addShape(pptx.ShapeType.ellipse, {
      x,
      y: y + idx * 0.48 + 0.08,
      w: 0.08,
      h: 0.08,
      fill: { color: C.green },
      line: { color: C.green },
    });
    addText(slide, b, x + 0.22, y + idx * 0.48, w - 0.22, 0.22, { size: 10.2, color: C.dark });
  });
}

function addFooterNote(slide, text) {
  addText(slide, text, 0.55, 6.85, 9, 0.16, { size: 6.8, color: C.gray });
}

function drawFlow(slide, x, y, w, h) {
  const nodes = [
    ["用户输入", x + 0.15, y + 0.12],
    ["AI服务边界判断", x + 2.25, y + 0.12],
    ["会话/查询/操作", x + 4.55, y + 0.12],
    ["答案检查", x + 6.95, y + 0.12],
    ["输出/转人工", x + 9.0, y + 0.12],
  ];
  nodes.forEach(([t, nx, ny], i) => {
    slide.addShape(i === 1 ? pptx.ShapeType.hexagon : pptx.ShapeType.roundRect, {
      x: nx,
      y: ny,
      w: 1.7,
      h: 0.72,
      rectRadius: 0.07,
      fill: { color: i === 1 ? "F7FBF8" : C.white },
      line: { color: i === 4 ? C.red : C.green, width: 1.1 },
    });
    addText(slide, t, nx + 0.12, ny + 0.25, 1.46, 0.14, { size: 8.2, bold: true, color: i === 4 ? C.red : C.green, align: "center" });
    if (i < nodes.length - 1) {
      slide.addShape(pptx.ShapeType.line, {
        x: nx + 1.72,
        y: ny + 0.36,
        w: 0.52,
        h: 0,
        line: { color: C.green, width: 1.1, beginArrowType: "none", endArrowType: "triangle" },
      });
    }
  });
}

function drawWheel(slide, x, y) {
  const items = [
    ["抽检", x + 1.2, y],
    ["归因", x + 2.55, y + 0.95],
    ["优化", x + 1.2, y + 1.9],
    ["验证", x - 0.15, y + 0.95],
  ];
  items.forEach(([t, nx, ny]) => {
    slide.addShape(pptx.ShapeType.ellipse, {
      x: nx,
      y: ny,
      w: 0.78,
      h: 0.78,
      fill: { color: C.mint },
      line: { color: C.green, width: 1 },
    });
    addText(slide, t, nx + 0.1, ny + 0.31, 0.58, 0.13, { size: 8, color: C.green, bold: true, align: "center" });
  });
  slide.addShape(pptx.ShapeType.arc, { x: x + 0.32, y: y + 0.25, w: 2.75, h: 2.15, line: { color: C.green2, width: 1.4, beginArrowType: "none", endArrowType: "triangle" }, adjustPoint: 0.28 });
}

function drawArchitecture(slide, x, y) {
  const steps = [
    ["RAG问答", "检索知识\n生成回答"],
    ["Agent框架", "理解意图\n规划路径"],
    ["服务执行", "调用工具\n接口联动"],
    ["运营复盘", "日志追踪\n持续优化"],
  ];
  steps.forEach(([t, s], i) => {
    const nx = x + i * 2.22;
    slide.addShape(pptx.ShapeType.roundRect, {
      x: nx,
      y,
      w: 1.72,
      h: 1.04,
      rectRadius: 0.08,
      fill: { color: i === 0 ? "F8FAFC" : C.mint2 },
      line: { color: i === 0 ? "CBD5E1" : C.green, width: 1 },
    });
    addText(slide, t, nx + 0.16, y + 0.18, 1.4, 0.14, { size: 8.8, bold: true, color: i === 0 ? C.gray : C.green, align: "center" });
    addText(slide, s, nx + 0.16, y + 0.45, 1.4, 0.34, { size: 7.2, color: C.gray, align: "center", breakLine: true });
    if (i < steps.length - 1) {
      slide.addShape(pptx.ShapeType.line, { x: nx + 1.8, y: y + 0.52, w: 0.38, h: 0, line: { color: C.green2, width: 1.1, endArrowType: "triangle" } });
    }
  });
}

function drawChannelHub(slide, x, y) {
  slide.addShape(pptx.ShapeType.ellipse, {
    x: x + 2.4,
    y: y + 0.65,
    w: 1.45,
    h: 1.45,
    fill: { color: C.green },
    line: { color: C.green },
  });
  addText(slide, "理赔\n智能体", x + 2.67, y + 1.08, 0.92, 0.35, { size: 10, bold: true, color: C.white, align: "center", breakLine: true });
  const ns = [
    ["服务大厅\n已对接", x + 0.1, y + 0.2, C.green],
    ["App\n对接中", x + 4.75, y + 0.2, C.amber],
    ["客服渠道\n规划中", x + 0.1, y + 2.3, C.gray],
    ["运营后台\n规划中", x + 4.75, y + 2.3, C.gray],
  ];
  ns.forEach(([t, nx, ny, color]) => {
    slide.addShape(pptx.ShapeType.roundRect, { x: nx, y: ny, w: 1.45, h: 0.72, rectRadius: 0.07, fill: { color: C.white }, line: { color, width: 1 } });
    addText(slide, t, nx + 0.15, ny + 0.18, 1.15, 0.25, { size: 8, color, bold: true, align: "center", breakLine: true });
    slide.addShape(pptx.ShapeType.line, { x: nx + (nx < x + 2.4 ? 1.5 : -0.05), y: ny + 0.36, w: nx < x + 2.4 ? 0.95 : -0.95, h: y + 1.38 - (ny + 0.36), line: { color, width: 0.9, dash: "dash", endArrowType: "triangle" } });
  });
}

function saveSvg(name, svg) {
  const file = path.join(IMG_DIR, name);
  fs.writeFileSync(file, svg, "utf8");
  return file;
}

saveSvg(
  "agent-flow.svg",
  `<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="520" viewBox="0 0 1200 520"><rect width="1200" height="520" fill="#F5FBF7"/><g font-family="Arial,'PingFang SC',sans-serif" font-size="28" font-weight="700"><rect x="70" y="210" rx="18" width="180" height="80" fill="#fff" stroke="#0B6B43" stroke-width="3"/><text x="160" y="260" text-anchor="middle" fill="#0B6B43">用户输入</text><polygon points="385,170 535,250 385,330 235,250" fill="#fff" stroke="#0B6B43" stroke-width="3"/><text x="385" y="245" text-anchor="middle" fill="#0B6B43">边界判断</text><rect x="570" y="210" rx="18" width="190" height="80" fill="#fff" stroke="#0B6B43" stroke-width="3"/><text x="665" y="260" text-anchor="middle" fill="#0B6B43">服务执行</text><rect x="840" y="210" rx="18" width="190" height="80" fill="#fff" stroke="#0B6B43" stroke-width="3"/><text x="935" y="260" text-anchor="middle" fill="#0B6B43">输出答案</text></g><g stroke="#0B6B43" stroke-width="4" marker-end="url(#a)"><defs><marker id="a" markerWidth="10" markerHeight="8" refX="10" refY="4" orient="auto"><path d="M0,0 L10,4 L0,8 Z" fill="#0B6B43"/></marker></defs><line x1="250" y1="250" x2="285" y2="250"/><line x1="535" y1="250" x2="570" y2="250"/><line x1="760" y1="250" x2="840" y2="250"/></g></svg>`
);
saveSvg(
  "channel-hub.svg",
  `<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="700" viewBox="0 0 1200 700"><rect width="1200" height="700" fill="#fff"/><circle cx="600" cy="350" r="110" fill="#0B6B43"/><text x="600" y="340" fill="#fff" text-anchor="middle" font-size="34" font-weight="700" font-family="Arial,'PingFang SC'">理赔智能体</text><text x="600" y="382" fill="#EAF6EF" text-anchor="middle" font-size="22" font-family="Arial,'PingFang SC'">统一能力底座</text><g font-family="Arial,'PingFang SC'" font-size="28" font-weight="700" fill="#0B6B43" stroke="#0B6B43" stroke-width="3"><rect x="100" y="100" rx="18" width="210" height="90" fill="#F5FBF7"/><text x="205" y="155" text-anchor="middle">服务大厅</text><rect x="890" y="100" rx="18" width="210" height="90" fill="#FFF7ED" stroke="#C89211"/><text x="995" y="155" text-anchor="middle" fill="#C89211">App对接中</text><rect x="100" y="510" rx="18" width="210" height="90" fill="#fff" stroke="#94A3B8"/><text x="205" y="565" text-anchor="middle" fill="#667085">客服渠道</text><rect x="890" y="510" rx="18" width="210" height="90" fill="#fff" stroke="#94A3B8"/><text x="995" y="565" text-anchor="middle" fill="#667085">运营后台</text></g><g stroke="#0B6B43" stroke-width="4" stroke-dasharray="10 8"><line x1="310" y1="145" x2="500" y2="300"/><line x1="890" y1="145" x2="700" y2="300"/><line x1="310" y1="555" x2="500" y2="400"/><line x1="890" y1="555" x2="700" y2="400"/></g></svg>`
);

function cover() {
  const slide = pptx.addSlide("MASTER");
  slide.background = { color: C.mint2 };
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: C.mint2 }, line: { color: C.mint2 } });
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.18, h: 7.5, fill: { color: C.green }, line: { color: C.green } });
  addTag(slide, "项目成果汇报", 0.72, 0.78, 1.25);
  addText(slide, slides[0].title, 0.72, 1.55, 6.6, 0.72, { size: 32, bold: true, color: C.dark });
  addText(slide, slides[0].headline, 0.74, 2.42, 6.4, 0.35, { size: 18, bold: true, color: C.green });
  addText(slide, slides[0].note, 0.76, 3.0, 6.2, 0.25, { size: 10.8, color: C.gray });
  drawFlow(slide, 0.75, 4.3, 10, 1.3);
  slide.addShape(pptx.ShapeType.arc, { x: 7.9, y: 0.75, w: 3.9, h: 3.9, line: { color: C.green, transparency: 62, width: 1.4 }, adjustPoint: 0.24 });
  slide.addShape(pptx.ShapeType.arc, { x: 8.6, y: 1.4, w: 3.4, h: 3.4, line: { color: C.green2, transparency: 72, width: 1.1 }, adjustPoint: 0.4 });
  addText(slide, "2026", 11.35, 6.75, 0.8, 0.18, { size: 8, color: C.gray, align: "right" });
}

function contentSlide(item, idx) {
  const slide = pptx.addSlide("MASTER");
  addHeader(slide, item.section, item.title, item.headline);
  if (item.metrics) {
    item.metrics.forEach((m, i) => addMetric(slide, m, 0.55 + i * 2.0, 1.92, 1.7));
    addBullets(slide, item.bullets, 0.65, 3.18, 5.6);
  } else {
    addBullets(slide, item.bullets, 0.65, 1.96, 5.4);
  }

  if (idx === 3) {
    drawWheel(slide, 7.3, 2.65);
    addChartPlaceholder(slide, "抽检漏斗：全部会话 → 样本 → badcase → 优化上线", 7.02, 4.95, 4.85, 1.22);
  } else if (idx === 5) {
    drawArchitecture(slide, 6.55, 2.55);
    addChartPlaceholder(slide, "能力分层：感知层 / 理解层 / 决策层 / 执行层 / 运营层", 6.55, 4.25, 5.1, 1.55);
  } else if (idx === 9) {
    drawChannelHub(slide, 6.15, 2.0);
  } else if ([1, 2, 6].includes(idx)) {
    addChartPlaceholder(slide, item.chart, 6.35, 1.9, 5.65, 4.2);
  } else if (idx === 4) {
    const cards = ["意图识别", "场景覆盖", "多模态识别"];
    cards.forEach((c, i) => {
      const x = 6.28 + i * 1.9;
      slide.addShape(pptx.ShapeType.roundRect, { x, y: 2.05, w: 1.62, h: 2.6, rectRadius: 0.08, fill: { color: i === 0 ? "FFF7ED" : C.mint2 }, line: { color: i === 0 ? C.amber : C.green, width: 1 } });
      addText(slide, c, x + 0.14, 2.28, 1.34, 0.18, { size: 9.3, bold: true, color: i === 0 ? C.amber : C.green, align: "center" });
      addText(slide, i === 0 ? "描述不清\n多意图\n上下文缺失" : i === 1 ? "操作类\n增值服务\n工单接口" : "OCR错误\nASR偏差\n图片语音", x + 0.2, 2.83, 1.22, 0.75, { size: 7.7, color: C.gray, align: "center", breakLine: true });
    });
    addChartPlaceholder(slide, item.chart, 6.28, 5.05, 5.6, 1.1);
  } else if (idx === 7) {
    drawArchitecture(slide, 6.45, 2.2);
    addText(slide, "传统问答：解释规则", 6.55, 3.75, 2.0, 0.18, { size: 9, color: C.gray });
    addText(slide, "Agent + Skill：调用能力完成操作", 8.55, 3.75, 2.8, 0.18, { size: 9, color: C.green, bold: true });
    addChartPlaceholder(slide, "操作闭环：诉求 → 识别 → Skill调用 → 结果反馈", 6.45, 4.55, 5.3, 1.22);
  } else if (idx === 8) {
    item.metrics.forEach((m, i) => addMetric(slide, m, 6.45 + i * 1.72, 2.0, 1.45));
    addChartPlaceholder(slide, "路线图：近期能力补齐 / 中期产品扩展 / 长期渠道协同", 6.45, 3.45, 5.3, 2.15);
  } else if (idx === 10) {
    const pairs = [
      ["执行日志缺失", "补齐 Agent 日志\n支持复盘追踪"],
      ["结构化输出不足", "自定义字段输出\n提升运营效率"],
      ["响应链路偏长", "识别与保单确认前置\n减少等待"],
    ];
    pairs.forEach(([a, b], i) => {
      const y = 2.0 + i * 1.25;
      slide.addShape(pptx.ShapeType.roundRect, { x: 6.25, y, w: 2.0, h: 0.72, rectRadius: 0.07, fill: { color: "FFF7ED" }, line: { color: C.amber, width: 1 } });
      addText(slide, a, 6.43, y + 0.26, 1.64, 0.13, { size: 8.4, bold: true, color: C.amber, align: "center" });
      slide.addShape(pptx.ShapeType.line, { x: 8.38, y: y + 0.36, w: 0.55, h: 0, line: { color: C.green, width: 1.1, endArrowType: "triangle" } });
      slide.addShape(pptx.ShapeType.roundRect, { x: 9.05, y, w: 2.45, h: 0.72, rectRadius: 0.07, fill: { color: C.mint2 }, line: { color: C.green, width: 1 } });
      addText(slide, b, 9.24, y + 0.17, 2.05, 0.28, { size: 7.7, bold: true, color: C.green, align: "center", breakLine: true });
    });
  } else {
    addChartPlaceholder(slide, item.chart, 6.35, 2.1, 5.4, 3.7);
  }
  addFooterNote(slide, `图表建议：${item.chart}`);
}

cover();
slides.slice(1).forEach((s, i) => contentSlide(s, i + 1));

const docSections = [
  ["项目背景与目标", "理赔智能体项目面向客户事业部核心理赔服务场景，目标是提升服务效率、改善用户体验，并推动智能客服从咨询解答向服务执行升级。"],
  ["项目总体成果概览", "项目已形成健康险与数生理赔两条服务主线，并围绕会话量、覆盖率、准确率、响应时长、转人工率等指标建立持续运营基础。"],
  ["健康险理赔智能服务成果", "健康险部分重点关注质量稳定性，通过一级差错率、整体差错率和抽检结果持续衡量服务质量。"],
  ["健康险日常运营机制", "当前月线上 AI 会话量 6000+，抽检比例约 10%。通过 badcase 复返和优化形成抽检、归因、优化、验证的运营闭环。"],
  ["健康险当前问题与后续架构升级计划", "主要问题包括客户描述不清、多意图处理、操作类场景覆盖、增值服务咨询覆盖、工单接口对接，以及 OCR/ASR 识别误差。后续建议从 RAG 问答架构升级为 Agent 服务架构。"],
  ["数生理赔智能服务成果", "数生理赔部分将围绕覆盖情况、回复准确率和响应时长建立数据看板，支撑规模化运营与体验优化。"],
  ["数生操作类场景突破", "通过 Skill 能力，理赔智能体可以从回答问题升级为完成操作，承接引导、查询、办理、触发流程等操作类场景。"],
  ["多渠道赋能进展", "服务大厅已完成理赔场景对接，App 对接中。后续固定入口和固定人群将调用理赔智能体，实现多渠道能力复用。"],
  ["当前问题与优化建议", "灵犀平台当前存在 Agent 执行日志不足、自定义结构化输出不支持等限制。响应时效方面，建议将情绪/敏感词识别、保单/赔案确认等模块前置化。"],
  ["下一阶段计划与资源诉求", "下一阶段目标是覆盖 90% 的理赔场景，并逐步支持除 3C 外的数生近千款产品，同时推进健康险架构升级和多渠道能力沉淀。"],
];

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after ?? 180 },
    heading: opts.heading,
    alignment: opts.align,
    children: [
      new TextRun({
        text,
        bold: !!opts.bold,
        size: opts.size || 22,
        color: opts.color || C.dark,
        font: "PingFang SC",
      }),
    ],
  });
}

function cell(text, fill = C.white, color = C.dark, bold = false) {
  return new TableCell({
    width: { size: 50, type: WidthType.PERCENTAGE },
    shading: { type: ShadingType.CLEAR, color: "auto", fill },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: C.line },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: C.line },
      left: { style: BorderStyle.SINGLE, size: 1, color: C.line },
      right: { style: BorderStyle.SINGLE, size: 1, color: C.line },
    },
    children: [p(text, { size: 20, color, bold, after: 80 })],
  });
}

const doc = new Document({
  creator: "理赔智能体项目组",
  title: "理赔智能体项目成果汇报",
  description: "面向客户事业部高层的项目成果汇报材料",
  sections: [
    {
      properties: {},
      children: [
        p("理赔智能体项目成果汇报", { heading: HeadingLevel.TITLE, size: 36, color: C.green, bold: true, after: 260 }),
        p("从智能问答走向理赔服务执行入口", { size: 26, color: C.dark, bold: true, after: 260 }),
        p("面向客户事业部高层 | 健康险理赔智能服务 / 数生理赔智能服务", { size: 20, color: C.gray, after: 420 }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [
            new TableRow({ children: [cell("汇报重点", C.green, C.white, true), cell("核心表达", C.green, C.white, true)] }),
            new TableRow({ children: [cell("效率提升", C.mint2, C.green, true), cell("通过抽检闭环、Skill 执行、多渠道复用提升理赔服务效率")] }),
            new TableRow({ children: [cell("用户体验", C.mint2, C.green, true), cell("从咨询解答扩展到查询、办理和流程触发，减少用户等待与重复描述")] }),
            new TableRow({ children: [cell("智能客服亮点", C.mint2, C.green, true), cell("Agent 架构升级、操作类场景突破、服务大厅与 App 渠道赋能")] }),
          ],
        }),
        ...docSections.flatMap(([title, body], idx) => [
          p(`${idx + 1}. ${title}`, { heading: HeadingLevel.HEADING_1, size: 26, color: C.green, bold: true, after: 160 }),
          p(body, { size: 22, color: C.dark, after: 220 }),
        ]),
        p("附：PPT 页面清单", { heading: HeadingLevel.HEADING_1, size: 26, color: C.green, bold: true, after: 160 }),
        ...slides.map((s, idx) => p(`${idx + 1}. ${s.title || s.section}：${s.headline || s.note || ""}`, { size: 20, color: C.dark, after: 100 })),
      ],
    },
  ],
});

await pptx.writeFile({ fileName: path.join(OUT, "report.pptx") });
const buffer = await Packer.toBuffer(doc);
fs.writeFileSync(path.join(OUT, "report.docx"), buffer);

console.log("Generated files:");
console.log("- report.pptx");
console.log("- report.docx");
console.log("- source.js");
console.log("- images/");
