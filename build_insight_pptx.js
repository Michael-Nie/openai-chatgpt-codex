// Build the OpenAI Codex × Frontier INSIGHT deck via pptxgenjs (skill workflow).
// Design: dark/light sandwich, icons-in-circles motif, tinted cards + shadows,
// native charts. NO header bars / edge accent stripes (per skill guidance).
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const FA = require("react-icons/fa");

// ---------- palette (content-informed, semantic) ----------
const INK="0E1729", INK2="1B2A47", PAPER="FFFFFF", EMER="10A37F", TEAL="0E7C86",
      AMBER="C2410C", RED="B91C1C", SLATE="64748B", INKTX="1E293B", ICE="CADCFC",
      CLOUD="F2F6F6", PALE_E="E6F5EF", PALE_T="E2F0EF", PALE_A="FBEDE3", MIST="DCE5E6";

const F="Arial";
const W=13.33, H=7.5;

// ---------- icon rasterization ----------
function svgOf(Icon, color, size=256){
  return ReactDOMServer.renderToStaticMarkup(React.createElement(Icon,{color,size:String(size)}));
}
async function png(Icon, color="#FFFFFF"){
  const buf = await sharp(Buffer.from(svgOf(Icon,color))).png().toBuffer();
  return "image/png;base64,"+buf.toString("base64");
}
const ICONS = {};
async function loadIcons(){
  const need = {
    layer:FA.FaLayerGroup, code:FA.FaCode, ent:FA.FaSitemap, robot:FA.FaRobot,
    chip:FA.FaMicrochip, wrench:FA.FaWrench, mobile:FA.FaMobileAlt, cogs:FA.FaCogs,
    shield:FA.FaShieldAlt, clock:FA.FaClock, brain:FA.FaBrain, db:FA.FaDatabase,
    terminal:FA.FaTerminal, chart:FA.FaChartLine, lock:FA.FaLock, users:FA.FaUsers,
    plug:FA.FaPlug, balance:FA.FaBalanceScale, crown:FA.FaCrown, dollar:FA.FaDollarSign,
    warn:FA.FaExclamationTriangle, check:FA.FaCheckCircle, times:FA.FaTimesCircle,
    sync:FA.FaSyncAlt, compass:FA.FaCompass, bolt:FA.FaBolt, arrow:FA.FaArrowRight,
    building:FA.FaBuilding, rocket:FA.FaRocket,
  };
  for(const k in need){ ICONS[k] = await png(need[k], "#FFFFFF"); }
}

// ---------- helpers ----------
const sh = () => ({ type:"outer", color:"0E1729", blur:7, offset:3, angle:90, opacity:0.13 });
const shUp = () => ({ type:"outer", color:"0E1729", blur:7, offset:3, angle:270, opacity:0.13 });

function iconCircle(slide, key, x, y, d, circ){
  slide.addShape("ellipse", { x, y, w:d, h:d, fill:{color:circ}, line:{type:"none"}, shadow: sh() });
  const p = d*0.27;
  slide.addImage({ data: ICONS[key], x:x+p, y:y+p, w:d-2*p, h:d-2*p });
}
function rrect(slide, x,y,w,h, fill, opt={}){
  slide.addShape("roundRect", { x,y,w,h, rectRadius:0.09, fill:{color:fill},
    line: opt.line||{type:"none"}, shadow: opt.shadow!==false ? sh() : undefined });
}
function head(slide, kicker, title, color=EMER){
  slide.addText(kicker, { x:0.62, y:0.44, w:12, h:0.3, fontFace:F, fontSize:12, color, bold:true, charSpacing:3, margin:0 });
  slide.addText(title,  { x:0.62, y:0.74, w:12.1, h:0.72, fontFace:F, fontSize:27, color:INK, bold:true, margin:0 });
}
function foot(slide, n, note){
  if(note) slide.addText(note, { x:0.62, y:7.08, w:10.4, h:0.3, fontFace:F, fontSize:8.5, color:SLATE, margin:0 });
  slide.addText(String(n), { x:12.5, y:7.08, w:0.55, h:0.3, fontFace:F, fontSize:9, color:SLATE, align:"right", margin:0 });
}
function tag(slide, x, y, text, color){
  slide.addText(text, { x, y, w:0.9, h:0.26, fontFace:F, fontSize:9, italic:true, color, align:"right", margin:0 });
}

(async () => {
  await loadIcons();
  const pres = new pptxgen();
  pres.defineLayout({ name:"WIDE", width:W, height:H });
  pres.layout = "WIDE";
  pres.author = "Insight Research";
  pres.title = "OpenAI Codex × Frontier 洞察";

  // ============================================================ 1 · TITLE (dark)
  let s = pres.addSlide(); s.background = { color: INK };
  s.addText("深度洞察  ·  DEEP INSIGHT", { x:0.9, y:1.15, w:11, h:0.4, fontFace:F, fontSize:14, color:EMER, bold:true, charSpacing:4, margin:0 });
  s.addText("OpenAI Codex × Frontier", { x:0.9, y:1.6, w:11.6, h:1.1, fontFace:F, fontSize:50, color:PAPER, bold:true, margin:0 });
  s.addText("纵深的编码 Agent  ×  横向的企业 Agent 平台 —— 同一战略的一体两面",
    { x:0.9, y:2.78, w:11.6, h:0.5, fontFace:F, fontSize:17, color:ICE, margin:0 });
  // two motif cards
  const cy=3.7, cw=5.6, chh=1.95;
  rrect(s, 0.9, cy, cw, chh, INK2);
  iconCircle(s, "code", 1.25, cy+0.32, 0.92, EMER);
  s.addText("Codex(纵深)", { x:2.45, y:cy+0.34, w:3.9, h:0.4, fontFace:F, fontSize:18, color:PAPER, bold:true, margin:0 });
  s.addText("把『软件工程』一个垂直做到极致:模型 × harness × 沙箱三位一体的自主编码 coworker。",
    { x:2.45, y:cy+0.78, w:3.0, h:1.0, fontFace:F, fontSize:11.5, color:ICE, margin:0 });
  rrect(s, 6.83, cy, cw, chh, INK2);
  iconCircle(s, "ent", 7.18, cy+0.32, 0.92, TEAL);
  s.addText("Frontier(横向)", { x:8.38, y:cy+0.34, w:3.9, h:0.4, fontFace:F, fontSize:18, color:PAPER, bold:true, margin:0 });
  s.addText("把 agent 推广到全企业,补齐上下文/治理/评估的 AI coworker 控制平面。",
    { x:8.38, y:cy+0.78, w:3.0, h:1.0, fontFace:F, fontSize:11.5, color:ICE, margin:0 });
  s.addText("×", { x:6.18, y:cy+0.55, w:0.8, h:0.8, fontFace:F, fontSize:40, color:EMER, bold:true, align:"center", margin:0 });
  s.addText("2026 年 6 月   ·   5 角度并行检索 + 对抗式核验(2/3 证伪即剔除)   ·   可信度全程分层标注",
    { x:0.9, y:6.05, w:11.6, h:0.35, fontFace:F, fontSize:11, color:SLATE, margin:0 });

  // ============================================================ 2 · EXEC SUMMARY (light)
  s = pres.addSlide(); s.background={color:PAPER};
  head(s, "EXECUTIVE SUMMARY", "执行摘要:七条核心洞察");
  const exec = [
    ["rocket", EMER, "一个战略,两条战线", "GPT-5.3-Codex 与 Frontier 同日(2026-02-05)发布,互为纵深与横向。"],
    ["code", EMER, "Codex 凶猛迭代", "codex-1 → 5-Codex → 5.1-Max(24h+)→ 5.3 → GPT-5.5;基准持续刷新。"],
    ["ent", TEAL, "Frontier=控制平面", "四组件(上下文/执行/评估/治理),像管员工一样管 agent。"],
    ["dollar", AMBER, "企业收入引擎", "企业收入 >40%、目标年底~50%,Frontier 是『载体』(自报)。"],
    ["crown", TEAL, "格局两层叠加", "编码工具层(Copilot/Cursor/Claude Code)+ 企业平台层(Frontier vs 在位者)。"],
    ["warn", AMBER, "三重『Frontier』撞名", "OpenAI Frontier ≠ Microsoft Frontier Suite ≠ OpenAI『frontier risk』。"],
    ["sync", RED, "对前作的关键更正", "Frontier Alliances/Deployment Co/Codex Labs 实为真实;仅 Codex 2.0 系列判虚构。"],
  ];
  // two columns: 4 left, 3 right
  const colX=[0.62, 6.86], rowY=1.78, rh=1.28;
  exec.forEach((e,i)=>{
    const col = i<4?0:1, idx = i<4?i:i-4;
    const x=colX[col], y=rowY+idx*rh;
    iconCircle(s, e[0], x, y, 0.7, e[1]);
    s.addText(e[2], { x:x+0.92, y:y-0.02, w:5.0, h:0.34, fontFace:F, fontSize:14.5, color:INK, bold:true, margin:0 });
    s.addText(e[3], { x:x+0.92, y:y+0.32, w:5.2, h:0.72, fontFace:F, fontSize:11, color:INKTX, margin:0 });
  });
  foot(s, 2, "〔强/中〕详见配套大纲与研究报告 · 中/弱项多为厂商自报或单源");

  // ============================================================ 3 · AGENT STACK (light)
  s = pres.addSlide(); s.background={color:PAPER};
  head(s, "STRATEGIC MAP", "OpenAI 的 Agent 栈:Codex 与 Frontier 的位置");
  const bands = [
    ["chip", "①  模型层 · Models", "GPT-5.x 家族,含 GPT-5.3-Codex 推理引擎", SLATE, CLOUD],
    ["wrench", "②  构建层 · Build", "AgentKit / Agents SDK(Agent Builder · Connector Registry · ChatKit)", SLATE, CLOUD],
    ["code", "③  专用执行 · Codex", "软件工程垂直的自主编码 coworker —— 模型×harness×沙箱(纵深)", EMER, PALE_E],
    ["ent", "④  编排治理 · Frontier", "全企业 AI coworker 控制平面:上下文 + 执行 + 评估 + 治理(横向)", TEAL, PALE_T],
    ["mobile", "⑤  分发层 · ChatGPT 超级应用", "整合 ChatGPT + Codex + agentic browsing + 合作伙伴应用", SLATE, CLOUD],
  ];
  let by=1.72; const bh2=0.82, bgap=0.13;
  bands.forEach((b,i)=>{
    const hl = (b[3]===EMER||b[3]===TEAL);
    rrect(s, 0.62, by, 12.1, bh2, b[4]);
    iconCircle(s, b[0], 0.86, by+0.13, 0.56, b[3]);
    s.addText(b[1], { x:1.6, y:by+0.06, w:3.8, h:0.7, fontFace:F, fontSize:15, color:hl?b[3]:INK, bold:true, valign:"middle", margin:0 });
    s.addText(b[2], { x:5.5, y:by+0.06, w:7.0, h:0.7, fontFace:F, fontSize:12, color:INKTX, valign:"middle", margin:0 });
    by += bh2+bgap;
  });
  s.addText([
    { text:"Codex = 纵深", options:{ bold:true, color:EMER } },
    { text:"(一个垂直做到极致) · ", options:{ color:INKTX } },
    { text:"Frontier = 横向", options:{ bold:true, color:TEAL } },
    { text:"(推广到全企业 + 治理) · 同日推进 2026-02-05", options:{ color:INKTX } },
  ], { x:0.62, y:by+0.06, w:12.1, h:0.4, fontFace:F, fontSize:12.5, align:"center", margin:0 });
  foot(s, 3, "栈为本报告分析框架(综合官方表述 + 媒体框定)");

  // ============================================================ 4 · CODEX LINEAGE + CHART (light)
  s = pres.addSlide(); s.background={color:PAPER};
  head(s, "CODEX · LINEAGE", "Codex:三代复用,凶猛迭代");
  // left facts
  rrect(s, 0.62, 1.78, 5.7, 1.55, CLOUD);
  iconCircle(s, "code", 0.92, 2.04, 0.62, EMER);
  s.addText("身份澄清(2021 ≠ 2025)", { x:1.72, y:1.96, w:4.4, h:0.34, fontFace:F, fontSize:14, color:INK, bold:true, margin:0 });
  s.addText("2021 初代(驱动 Copilot,API 2023-03 关停)≠ 2025 重启的自主软件工程 agent(codex-1 ← o3 RL 后训练)。",
    { x:1.72, y:2.3, w:4.45, h:0.95, fontFace:F, fontSize:11, color:INKTX, margin:0 });
  rrect(s, 0.62, 3.5, 5.7, 1.55, PALE_A);
  iconCircle(s, "sync", 0.92, 3.76, 0.62, AMBER);
  s.addText("更正(经核验)", { x:1.72, y:3.68, w:4.4, h:0.34, fontFace:F, fontSize:14, color:INK, bold:true, margin:0 });
  s.addText("GPT-5.2-Codex 实为 2025-12-18(非 1 月);GPT-5.4 确实存在(2026-03-05,非 Codex 专用)。",
    { x:1.72, y:4.02, w:4.45, h:0.95, fontFace:F, fontSize:11, color:INKTX, margin:0 });
  s.addText("基准已从 SWE-bench Verified 迁到 SWE-bench Pro / Terminal-Bench 2.0 —— 跨版本不可直接续比。",
    { x:0.62, y:5.25, w:5.7, h:0.7, fontFace:F, fontSize:10.5, italic:true, color:SLATE, margin:0 });
  // right chart
  s.addText("SWE-bench Verified(%)", { x:6.7, y:1.74, w:6, h:0.3, fontFace:F, fontSize:12, color:INK, bold:true, margin:0 });
  s.addChart(pres.charts.BAR, [{
    name:"SWE-bench Verified", labels:["GPT-5-Codex","5.1-Codex-Max","GPT-5.5*"], values:[74.5,77.9,88.7]
  }], {
    x:6.6, y:2.05, w:6.2, h:3.3, barDir:"col", chartColors:[EMER,TEAL,"5EC9B0"],
    showValue:true, dataLabelPosition:"outEnd", dataLabelColor:INKTX, dataLabelFontFace:F, dataLabelFontSize:11,
    valAxisMinVal:60, valAxisMaxVal:100, catAxisLabelColor:SLATE, valAxisLabelColor:SLATE,
    catAxisLabelFontFace:F, valAxisLabelFontFace:F, catAxisLabelFontSize:10, valAxisLabelFontSize:9,
    valGridLine:{color:MIST,size:0.5}, catGridLine:{style:"none"}, showLegend:false,
    chartArea:{fill:{color:PAPER}},
  });
  s.addText("* GPT-5.5 88.7% 为 OpenAI 自报;独立 harness 约 82.6%。",
    { x:6.6, y:5.45, w:6.2, h:0.35, fontFace:F, fontSize:10, italic:true, color:AMBER, margin:0 });
  foot(s, 4, "数字为厂商自报基准 · 详见研究报告 §1.2");

  // ============================================================ 5 · CODEX ENGINE (light)
  s = pres.addSlide(); s.background={color:PAPER};
  head(s, "CODEX · ENGINEERING CORE", "Codex 的工程内核:模型 × Harness × 沙箱");
  const eng = [
    ["cogs", EMER, "Rust Harness", "零依赖、毫秒级启动、长会话无 GC 抖动;五工具循环走 Responses API。"],
    ["shield", TEAL, "OS 级沙箱", "macOS=Seatbelt · Linux=bubblewrap+seccomp · Windows=受限令牌(+WSL)。"],
    ["clock", AMBER, "Compaction · 长时程", "跨上下文窗口蒸馏关键状态;5.1-Max 内部观察 24h+ 自主迭代。"],
    ["brain", EMER, "自适应推理", "low/medium/high + xhigh;单模型自调步,≠ 系统级路由。"],
  ];
  const gx=[0.62,6.86], gy=[1.78,3.95], gcw=5.85, gch=1.95;
  eng.forEach((e,i)=>{
    const x=gx[i%2], y=gy[Math.floor(i/2)];
    rrect(s, x, y, gcw, gch, CLOUD);
    iconCircle(s, e[0], x+0.32, y+0.34, 0.8, e[1]);
    s.addText(e[2], { x:x+1.32, y:y+0.36, w:4.3, h:0.4, fontFace:F, fontSize:16, color:INK, bold:true, margin:0 });
    s.addText(e[3], { x:x+1.32, y:y+0.82, w:4.35, h:1.0, fontFace:F, fontSize:11.5, color:INKTX, margin:0 });
  });
  foot(s, 5, "架构事实多为官方文档/源码直读(高可信)");

  // ============================================================ 6 · FRONTIER 4 COMPONENTS (light)
  s = pres.addSlide(); s.background={color:PAPER};
  head(s, "FRONTIER · ARCHITECTURE", "OpenAI Frontier:企业 Agent 的控制平面", TEAL);
  s.addText("2026-02-05 发布的端到端企业平台,构建 / 部署 / 管理 AI agent(AI coworkers)。四大核心组件:",
    { x:0.62, y:1.56, w:12.1, h:0.34, fontFace:F, fontSize:12.5, color:INKTX, margin:0 });
  const comp = [
    ["db", EMER, "① 业务上下文(语义层)", "连接 CRM/ERP/工单/数仓/内部应用,统一权限与检索,成为可共享的企业语义层。"],
    ["terminal", TEAL, "② Agent 执行", "推理、处理文件、跑代码、调工具并建立记忆;开放环境,兼容第三方 agent。"],
    ["chart", AMBER, "③ 评估与优化", "内建反馈回路,表现对人类管理者透明,随时间持续改进(如绩效复盘)。"],
    ["lock", INK2, "④ 安全与治理", "对人 + AI coworker 统一 IAM;受限身份、清晰边界、可审计动作。"],
  ];
  const fx=[0.62,6.86], fy=[1.98,4.18], fcw=5.85, fch=2.0;
  comp.forEach((e,i)=>{
    const x=fx[i%2], y=fy[Math.floor(i/2)];
    rrect(s, x, y, fcw, fch, CLOUD);
    iconCircle(s, e[0], x+0.32, y+0.34, 0.8, e[1]);
    s.addText(e[2], { x:x+1.32, y:y+0.36, w:4.35, h:0.4, fontFace:F, fontSize:15, color:INK, bold:true, margin:0 });
    s.addText(e[3], { x:x+1.32, y:y+0.82, w:4.35, h:1.05, fontFace:F, fontSize:11.5, color:INKTX, margin:0 });
  });
  foot(s, 6, "组件清单为官方;『控制平面』为媒体框定 · 研究报告 §5.2");

  // ============================================================ 7 · FRONTIER OPERATING + OPENNESS (light)
  s = pres.addSlide(); s.background={color:PAPER};
  head(s, "FRONTIER · MODEL & OPENNESS", "像管员工一样管 Agent —— 且兼容多供应商", TEAL);
  // left: operating model flow
  rrect(s, 0.62, 1.78, 5.95, 4.5, CLOUD);
  iconCircle(s, "users", 0.92, 2.04, 0.66, TEAL);
  s.addText("运营心智模型", { x:1.74, y:2.06, w:4.6, h:0.4, fontFace:F, fontSize:16, color:INK, bold:true, margin:0 });
  const flow = ["入职 Onboarding —— 学习机构知识与内部语言","在岗学习 + 反馈 —— 表现透明、定向改进","权限与边界 —— 受限身份、与 IAM 打通","→ agent 从『一次性脚本』升级为可治理数字员工"];
  flow.forEach((t,i)=>{
    s.addText((i+1<=3? (i+1)+"." : "▸"), { x:1.0, y:2.7+i*0.82, w:0.4, h:0.4, fontFace:F, fontSize:14, color:i===3?EMER:TEAL, bold:true, margin:0 });
    s.addText(t, { x:1.4, y:2.7+i*0.82, w:4.95, h:0.78, fontFace:F, fontSize:12, color:i===3?EMER:INKTX, bold:i===3, margin:0 });
  });
  // right: openness
  rrect(s, 6.78, 1.78, 5.95, 2.15, PALE_T);
  iconCircle(s, "plug", 7.08, 2.04, 0.66, EMER);
  s.addText("开放性 · 多供应商", { x:7.9, y:2.06, w:4.6, h:0.4, fontFace:F, fontSize:16, color:INK, bold:true, margin:0 });
  s.addText("兼容 OpenAI / 企业自建 / 第三方(Google · Microsoft · Anthropic)agent,经 MCP 集成;不取代 AgentKit/SDK,而是统一上下文+执行+评估。",
    { x:7.08, y:2.72, w:5.5, h:1.1, fontFace:F, fontSize:11.5, color:INKTX, margin:0 });
  rrect(s, 6.78, 4.13, 5.95, 2.15, PALE_A);
  iconCircle(s, "balance", 7.08, 4.39, 0.66, AMBER);
  s.addText("中立悖论(分析)", { x:7.9, y:4.41, w:4.6, h:0.4, fontFace:F, fontSize:16, color:INK, bold:true, margin:0 });
  s.addText("宣称中立,却与自家 GPT-5.x 深绑;批评者主张『LLM 中立的控制平面』。真实中立程度取决于第三方集成深度,仍待观察。",
    { x:7.08, y:5.07, w:5.5, h:1.1, fontFace:F, fontSize:11.5, color:INKTX, margin:0 });
  foot(s, 7, "运营模型为官方表述;中立悖论为分析/批评(部分批评源自身卖中立方案)");

  // ============================================================ 8 · SYNTHESIS (light)
  s = pres.addSlide(); s.background={color:PAPER};
  head(s, "SYNTHESIS", "纵深 × 横向:同一战略的一体两面");
  rrect(s, 0.62, 1.85, 5.7, 3.55, PALE_E);
  iconCircle(s, "code", 0.95, 2.16, 0.8, EMER);
  s.addText("Codex —— 纵深", { x:1.95, y:2.2, w:4.2, h:0.45, fontFace:F, fontSize:19, color:INK, bold:true, margin:0 });
  s.addText([
    {text:"把『软件工程』一个垂直做到极致。",options:{breakLine:true,bold:true}},
    {text:"模型 × harness × 沙箱三位一体;长时程自主、并行、可委派。",options:{breakLine:true}},
    {text:"回答:『单个 agent 能做到多强?』",options:{color:EMER,bold:true}},
  ], { x:0.95, y:3.2, w:5.05, h:2.0, fontFace:F, fontSize:12.5, color:INKTX, paraSpaceAfter:8, margin:0 });
  rrect(s, 6.98, 1.85, 5.7, 3.55, PALE_T);
  iconCircle(s, "ent", 7.31, 2.16, 0.8, TEAL);
  s.addText("Frontier —— 横向", { x:8.31, y:2.2, w:4.2, h:0.45, fontFace:F, fontSize:19, color:INK, bold:true, margin:0 });
  s.addText([
    {text:"把 agent 推广到所有职能。",options:{breakLine:true,bold:true}},
    {text:"补齐上下文 / 治理 / 评估,让 agent 可被像员工一样管理。",options:{breakLine:true}},
    {text:"回答:『一群 agent 如何在企业里安全协作?』",options:{color:TEAL,bold:true}},
  ], { x:7.31, y:3.2, w:5.05, h:2.0, fontFace:F, fontSize:12.5, color:INKTX, paraSpaceAfter:8, margin:0 });
  s.addText("×", { x:6.32, y:3.0, w:0.66, h:0.7, fontFace:F, fontSize:34, color:SLATE, bold:true, align:"center", margin:0 });
  rrect(s, 0.62, 5.6, 12.1, 0.95, INK, {shadow:false});
  s.addText("两条战线共同把竞争从『单轮代码质量』推向『企业级 agent 操作系统』——谁能同时做好『强单体 agent』与『可治理的 agent 群』,谁赢下一阶段。",
    { x:0.95, y:5.6, w:11.45, h:0.95, fontFace:F, fontSize:13, color:PAPER, bold:true, valign:"middle", margin:0 });
  foot(s, 8, "");

  // ============================================================ 9 · COMPETITION (light)
  s = pres.addSlide(); s.background={color:PAPER};
  head(s, "COMPETITIVE LANDSCAPE", "两层竞争格局");
  // left: coding tool layer + chart
  s.addText("编码工具层 · 职场采用率(JetBrains 2026,n>10,000)",
    { x:0.62, y:1.7, w:6.1, h:0.34, fontFace:F, fontSize:12.5, color:INK, bold:true, margin:0 });
  s.addChart(pres.charts.BAR, [{
    name:"采用率 %", labels:["GitHub Copilot","Cursor","Claude Code"], values:[29,18,18]
  }], {
    x:0.55, y:2.05, w:6.05, h:2.7, barDir:"bar", chartColors:[TEAL],
    showValue:true, dataLabelPosition:"outEnd", dataLabelColor:INKTX, dataLabelFontFace:F, dataLabelFontSize:11,
    catAxisLabelColor:INKTX, valAxisLabelColor:SLATE, catAxisLabelFontFace:F, valAxisLabelFontFace:F,
    catAxisLabelFontSize:11, valAxisLabelFontSize:9, valAxisMaxVal:35,
    valGridLine:{color:MIST,size:0.5}, catGridLine:{style:"none"}, showLegend:false, chartArea:{fill:{color:PAPER}},
  });
  s.addText([
    {text:"Claude Code 满意度居首:CSAT 91% / NPS 54", options:{breakLine:true,color:INK,bold:true}},
    {text:"Cursor ~$2B ARR · 一年内 Claude Code 采用 ~6x", options:{}},
  ], { x:0.62, y:4.95, w:6.05, h:0.9, fontFace:F, fontSize:11, color:INKTX, paraSpaceAfter:4, margin:0 });
  // right: enterprise platform layer
  s.addText("企业 Agent 平台层", { x:6.95, y:1.7, w:5.8, h:0.34, fontFace:F, fontSize:12.5, color:INK, bold:true, margin:0 });
  const plats = [
    ["OpenAI Frontier","多供应商控制平面 + 语义层;非 GA、定制定价", EMER],
    ["Salesforce Agentforce","18,000+ 公司;最贴近 CRM,锁定最紧", SLATE],
    ["MS Agent 365 / Frontier Suite","M365 E7,$99/用户;分发优势(撞名!)", SLATE],
    ["IBM watsonx Orchestrate","多厂商控制平面;受监管行业", SLATE],
    ["ServiceNow AI Control Tower","ITSM/工作流就近治理", SLATE],
  ];
  let py=2.1;
  plats.forEach((p)=>{
    rrect(s, 6.92, py, 5.8, 0.82, p[2]===EMER?PALE_E:CLOUD);
    s.addText(p[0], { x:7.16, y:py+0.1, w:5.4, h:0.34, fontFace:F, fontSize:12.5, color:p[2]===EMER?EMER:INK, bold:true, margin:0 });
    s.addText(p[1], { x:7.16, y:py+0.43, w:5.4, h:0.32, fontFace:F, fontSize:10, color:INKTX, margin:0 });
    py += 0.94;
  });
  foot(s, 9, "JetBrains 2026 调查(强)· 平台数据多源(强/中)");

  // ============================================================ 10 · BUSINESS (light)
  s = pres.addSlide(); s.background={color:PAPER};
  head(s, "BUSINESS & STRATEGY", "商业:企业收入引擎(多为自报 / 单源)", AMBER);
  const stats = [
    [">40%→~50%","企业收入占比(目标年底)", EMER],
    ["$57–60亿","OpenAI 2026 Q1 营收", TEAL],
    ["+~$10亿","季度领先 Anthropic", AMBER],
    ["500万+","Codex 周活(6-02)", EMER],
  ];
  const sx0=0.62, scw=2.95, sgap=0.13;
  stats.forEach((st,i)=>{
    const x=sx0+i*(scw+sgap);
    rrect(s, x, 1.85, scw, 1.95, CLOUD);
    s.addText(st[0], { x:x+0.12, y:2.0, w:scw-0.24, h:0.9, fontFace:F, fontSize:30, color:st[2], bold:true, align:"center", margin:0 });
    s.addText(st[1], { x:x+0.12, y:3.0, w:scw-0.24, h:0.65, fontFace:F, fontSize:11.5, color:INKTX, align:"center", margin:0 });
  });
  // codex users growth line chart
  s.addText("Codex 周活增长(万,2026)", { x:0.62, y:4.05, w:6, h:0.32, fontFace:F, fontSize:12, color:INK, bold:true, margin:0 });
  s.addChart(pres.charts.LINE, [{
    name:"周活(万)", labels:["2月","3月","4月初","4-21","6-02"], values:[160,200,300,400,500]
  }], {
    x:0.55, y:4.35, w:6.1, h:2.45, lineSize:3, lineSmooth:true, chartColors:[EMER],
    showValue:true, dataLabelColor:INKTX, dataLabelFontFace:F, dataLabelFontSize:10, dataLabelPosition:"t",
    catAxisLabelColor:SLATE, valAxisLabelColor:SLATE, catAxisLabelFontFace:F, valAxisLabelFontFace:F,
    catAxisLabelFontSize:10, valAxisLabelFontSize:9, valGridLine:{color:MIST,size:0.5}, showLegend:false, chartArea:{fill:{color:PAPER}},
  });
  // right narrative
  rrect(s, 6.95, 4.05, 5.78, 2.75, PALE_A);
  iconCircle(s, "dollar", 7.25, 4.32, 0.66, AMBER);
  s.addText("战略与口径提醒", { x:8.07, y:4.34, w:4.4, h:0.4, fontFace:F, fontSize:15, color:INK, bold:true, margin:0 });
  s.addText([
    {text:"Frontier 是企业营收『载体』,与 Salesforce/ServiceNow/Microsoft 正面相撞。",options:{breakLine:true}},
    {text:"⚠ 口径:OpenAI 季度营收领先,但 ARR 口径 Anthropic 更高——勿混用。",options:{color:AMBER}},
  ], { x:7.25, y:4.98, w:5.25, h:1.7, fontFace:F, fontSize:11, color:INKTX, paraSpaceAfter:6, margin:0 });
  foot(s, 10, "营收/周活均为 OpenAI 自报或单一媒体(The Information),无第三方审计");

  // ============================================================ 11 · RISKS & VERIFICATION (light)
  s = pres.addSlide(); s.background={color:PAPER};
  head(s, "RISKS & VERIFICATION", "核验结论与重要更正(诚实披露)", AMBER);
  const ver = [
    ["times", RED, "✗ 已剔除(证伪)", "「Codex 2.0 / Orchestrator / SafeDeploy / OpenRepo」无可信来源;真实编排规范 = Symphony。"],
    ["sync", AMBER, "🔄 重要更正", "Frontier Alliances / Deployment Co / Codex Labs 实为真实;5.2-Codex=2025-12-18;3M 周活在 4 月初;GPT-5.4 存在。"],
    ["check", TEAL, "~ 已降级 / 已证实", "「省 4x token」为 n≈1 任务级;营收/成效/周活多为自报。Frontier 多供应商开放性已多源证实。"],
  ];
  const vx=0.62, vcw=3.92, vgap=0.17;
  ver.forEach((v,i)=>{
    const x=vx+i*(vcw+vgap);
    rrect(s, x, 1.82, vcw, 3.1, CLOUD);
    iconCircle(s, v[0], x+0.28, 2.08, 0.7, v[1]);
    s.addText(v[2], { x:x+0.26, y:2.86, w:vcw-0.5, h:0.4, fontFace:F, fontSize:14.5, color:INK, bold:true, margin:0 });
    s.addText(v[3], { x:x+0.26, y:3.28, w:vcw-0.5, h:1.55, fontFace:F, fontSize:11, color:INKTX, margin:0 });
  });
  rrect(s, 0.62, 5.18, 12.1, 1.25, INK, {shadow:false});
  iconCircle(s, "warn", 0.92, 5.46, 0.66, AMBER);
  s.addText("⚠ 三重『Frontier』撞名,引用务必区分", { x:1.74, y:5.42, w:10.8, h:0.36, fontFace:F, fontSize:14, color:PAPER, bold:true, margin:0 });
  s.addText("① OpenAI Frontier(企业 agent 平台)  ·  ② Microsoft Frontier Suite(M365 E7,$99/用户)  ·  ③ OpenAI『frontier risk』(前沿模型安全政策)",
    { x:1.74, y:5.84, w:10.8, h:0.5, fontFace:F, fontSize:11.5, color:ICE, margin:0 });
  foot(s, 11, "对抗式核验:2/3 证伪即剔除 · 详见研究报告 §6");

  // ============================================================ 12 · SELECTION GUIDE (light)
  s = pres.addSlide(); s.background={color:PAPER};
  head(s, "RECOMMENDATIONS", "选型建议:按场景对号入座");
  const guide = [
    ["bolt","异步委派 / 并行批量 / CI","Codex(云)", EMER, "隔离容器 + worktree + 免审批自主更可控"],
    ["shield","数据敏感 / 隔离 / 合规(编码)","Claude Code", TEAL, "本地优先 + 应用层细粒度权限 + hooks"],
    ["ent","跨职能 agent 投产 + 治理 / 审计","OpenAI Frontier", TEAL, "共享上下文 + 评估 + 权限(早期访问、定价未公开)"],
    ["building","已重度绑定某 SaaS 生态","Agentforce / Agent 365", AMBER, "贴近业务与采购,但编排锁定更紧"],
    ["compass","务实团队","双持 + 平台中立", INK2, "Codex/Claude 跑编码;按生态选编排层"],
  ];
  let yy=1.82;
  guide.forEach((g)=>{
    rrect(s, 0.62, yy, 12.1, 0.92, CLOUD);
    iconCircle(s, g[0], 0.86, yy+0.16, 0.6, g[3]);
    s.addText(g[1], { x:1.66, y:yy+0.06, w:5.0, h:0.8, fontFace:F, fontSize:13, color:INK, bold:true, valign:"middle", margin:0 });
    s.addText(g[2], { x:6.8, y:yy+0.06, w:2.7, h:0.8, fontFace:F, fontSize:13, color:g[3]===INK2?INK:g[3], bold:true, valign:"middle", margin:0 });
    s.addText(g[4], { x:9.55, y:yy+0.06, w:3.1, h:0.8, fontFace:F, fontSize:10.5, color:INKTX, valign:"middle", margin:0 });
    yy += 1.0;
  });
  foot(s, 12, "综合工程事实与社区共识;非厂商背书 · 落地前须验证集成深度与合规");

  // ============================================================ 13 · CLOSING (dark)
  s = pres.addSlide(); s.background={color:INK};
  s.addText("核心判断", { x:0.9, y:0.95, w:11, h:0.7, fontFace:F, fontSize:32, color:PAPER, bold:true, margin:0 });
  const take = [
    ["Codex 护城河 = 模型×harness 协同设计","RL 让模型与 agent 循环协同;代价是 OpenAI 模型锁定。"],
    ["Frontier 的赌注 = 治理与上下文,而非模型","若兑现,护城河来自企业系统集成深度,而非单纯模型领先。"],
    ["竞争主战场 = 企业级 Agent 操作系统","从『单轮代码质量』转向『跨窗口状态 + 多 agent 治理』。"],
    ["认知风险 = 别被『谁第一』与『撞名』带偏","任何能力数字都须绑定『产品 + 版本 + 日期 + 来源』。"],
  ];
  take.forEach((t,i)=>{
    const y=1.95+i*1.02;
    iconCircle(s, ["bolt","ent","crown","warn"][i], 0.9, y, 0.62, [EMER,TEAL,EMER,AMBER][i]);
    s.addText(t[0], { x:1.78, y:y-0.04, w:10.8, h:0.4, fontFace:F, fontSize:15.5, color:PAPER, bold:true, margin:0 });
    s.addText(t[1], { x:1.78, y:y+0.36, w:10.8, h:0.5, fontFace:F, fontSize:11.5, color:ICE, margin:0 });
  });
  s.addText("来源:OpenAI 官方 · TechCrunch/CNBC/Fortune/VentureBeat/SiliconANGLE/The Information · JetBrains 2026 · Epoch AI · Aikido · 详见配套大纲与研究报告。本次 WebFetch 受限,以多源搜索摘要交叉验证。",
    { x:0.9, y:6.35, w:11.6, h:0.7, fontFace:F, fontSize:9.5, color:SLATE, margin:0 });

  const out = "/home/user/openai-chatgpt-codex/OpenAI_Codex_Frontier_洞察PPT.pptx";
  await pres.writeFile({ fileName: out });
  console.log("saved:", out);
})();
