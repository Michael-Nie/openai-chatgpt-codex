// Build the OpenAI Codex × Frontier insight-report OUTLINE (.docx) via docx-js.
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, LevelFormat,
  HeadingLevel, TableOfContents, PageBreak, BorderStyle, PageNumber,
  Header, Footer,
} = require("docx");

const NAVY = "0E142E", GREEN = "10A37F", TEAL = "0B7A75", GRAY = "6B7280", AMBER = "B45309";

// ---- helpers ----
const T = (text, opts = {}) => new TextRun({ text, font: "Arial", ...opts });
const P = (children, opts = {}) => new Paragraph({ children: Array.isArray(children) ? children : [children], ...opts });

const h1 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [T(text, { bold: true })] });
const h2 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [T(text, { bold: true })] });
const h3 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_3, children: [T(text, { bold: true })] });

// bullet with optional bold lead and a trailing confidence tag
const bul = (lead, body, level = 0, tag = null, tagColor = GRAY) => {
  const runs = [];
  if (lead) runs.push(T(lead, { bold: true }));
  if (body) runs.push(T(body));
  if (tag) runs.push(new TextRun({ text: "  〔" + tag + "〕", font: "Arial", italics: true, color: tagColor, size: 18 }));
  return new Paragraph({ numbering: { reference: "bullets", level }, children: runs });
};
const num = (lead, body) => new Paragraph({
  numbering: { reference: "execnums", level: 0 },
  children: [T(lead, { bold: true }), T(body)],
});

const rule = (color = GREEN) => new Paragraph({
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color, space: 1 } },
  spacing: { after: 120 }, children: [T("")],
});

// ---- numbering ----
const numbering = {
  config: [
    {
      reference: "bullets",
      levels: [
        { level: 0, format: LevelFormat.BULLET, text: "▪", alignment: AlignmentType.LEFT,
          style: { run: { color: GREEN }, paragraph: { indent: { left: 460, hanging: 260 }, spacing: { after: 40 } } } },
        { level: 1, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 920, hanging: 280 }, spacing: { after: 30 } } } },
        { level: 2, format: LevelFormat.BULLET, text: "·", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 1380, hanging: 280 }, spacing: { after: 30 } } } },
      ],
    },
    {
      reference: "execnums",
      levels: [
        { level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { run: { bold: true, color: GREEN }, paragraph: { indent: { left: 520, hanging: 320 }, spacing: { after: 60 } } } },
      ],
    },
  ],
};

// ---- styles ----
const styles = {
  default: { document: { run: { font: "Arial", size: 21 } } }, // 10.5pt body
  paragraphStyles: [
    { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 30, bold: true, font: "Arial", color: NAVY },
      paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 0,
        border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: GREEN, space: 4 } } } },
    { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 25, bold: true, font: "Arial", color: TEAL },
      paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 1 } },
    { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 22, bold: true, font: "Arial", color: NAVY },
      paragraph: { spacing: { before: 140, after: 60 }, outlineLevel: 2 } },
  ],
};

// ---- title block ----
const titleBlock = [
  new Paragraph({ spacing: { before: 1200, after: 0 }, children: [T("OpenAI CODEX × FRONTIER", { bold: true, color: GREEN, size: 26 })] }),
  new Paragraph({ spacing: { before: 60, after: 0 }, children: [T("洞察报告大纲", { bold: true, color: NAVY, size: 64 })] }),
  new Paragraph({ spacing: { before: 120, after: 0 }, children: [T("Coding Agent(纵深) × Enterprise Agent Platform(横向)— 同一 Agent 战略的一体两面", { color: GRAY, size: 24 })] }),
  rule(GREEN),
  new Paragraph({ spacing: { before: 80 }, children: [
    T("研究日期:", { bold: true }), T("2026-06-22      "),
    T("方法:", { bold: true }), T("5 角度并行检索 + 对抗式核验(2/3 证伪即剔除)+ 可信度分层"),
  ] }),
  new Paragraph({ spacing: { after: 80 }, children: [
    T("可信度图例:", { bold: true }),
    T("强 = 官方+多源交叉 · ", { color: GREEN }),
    T("中 = 单一媒体/厂商自报 · ", { color: AMBER }),
    T("弱 = 单源/已降级 · ", { color: GRAY }),
    T("✗ = 已剔除(证伪)", { color: "B91C1C" }),
  ] }),
  new Paragraph({ spacing: { after: 0 }, children: [
    T("用途:", { bold: true }),
    T("本大纲为洞察报告与配套 PPT 的结构蓝本——每个一级标题约对应 1 个章节 / 一组幻灯片,每条要点约对应 1 个幻灯片要点。"),
  ] }),
  new Paragraph({ children: [new PageBreak()] }),
];

// ---- TOC ----
const toc = [
  new Paragraph({ children: [T("目录", { bold: true, color: NAVY, size: 30 })], spacing: { after: 120 } }),
  new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-2" }),
  new Paragraph({ children: [new PageBreak()] }),
];

// ---- body ----
const body = [
  h1("一、执行摘要:七条核心洞察"),
  num("一个战略,两条战线 — ", "2026-02-05 同日发布 GPT-5.3-Codex(纵深的自主编码 agent)与 Frontier(横向的企业 agent 平台)。"),
  num("Codex 技术凶猛迭代 — ", "codex-1(o3 后训练)→ GPT-5-Codex → 5.1-Codex-Max(compaction、24h+)→ 5.2(2025-12-18)→ 5.3(2026-02-05)→ GPT-5.5(2026-04-23)。"),
  num("Frontier = 企业 agent 的「操作系统」 — ", "四组件(业务上下文/执行/评估/治理),「像管员工一样管 agent」;并兼容第三方(Google/MS/Anthropic)agent。"),
  num("企业营收引擎 — ", "企业收入 >40%、目标年底约 50%,Frontier 是「载体」;Q1 营收约 $57–60 亿、季度领先 Anthropic ~$10 亿(自报/单源)。"),
  num("格局两层叠加 — ", "编码工具层(Copilot 29% / Cursor ~$2B ARR / Claude Code CSAT 91%)+ 企业平台层(Frontier vs Agentforce / Agent 365 / watsonx / ServiceNow)。"),
  num("三重「Frontier」撞名 — ", "OpenAI Frontier(平台)≠ Microsoft Frontier Suite(M365 E7)≠ OpenAI「frontier risk」(模型安全政策)。"),
  num("对前作的关键更正 — ", "「Frontier Alliances / Deployment Company / Codex Labs」实为真实计划(此前误判虚构);仅「Codex 2.0/Orchestrator/SafeDeploy/OpenRepo」仍判虚构(真实编排规范 = Symphony)。"),

  h1("二、研究方法与可信度框架"),
  bul("流程 — ", "问题拆为 5 角度并行检索 → 去重抓取 → 对抗式核验(2/3 证伪即剔除)→ 按可信度分层综合。"),
  bul("可信度分层 — ", "强 / 中 / 弱 全程标注;凡中/弱多为厂商自报或单源,引用前须二次核对。"),
  bul("环境限制 — ", "本次 WebFetch 对 openai.com / 主流媒体全程 403,结论基于 WebSearch 摘要的 10+ 源交叉验证;数字以官方 system card 为准。", 0, "重要", AMBER),

  h1("三、OpenAI Agent 栈:Codex 与 Frontier 的定位"),
  bul("五层栈 — ", "① 模型(GPT-5.x 含 5.3-Codex)→ ② 构建(AgentKit / Agents SDK)→ ③ Codex(专用执行)→ ④ Frontier(编排治理)→ ⑤ ChatGPT 超级应用(分发)。"),
  bul("纵深 vs 横向 — ", "Codex 回答「单个 agent 能多强」;Frontier 回答「一群 agent 如何在企业里安全协作」。"),
  bul("同源同日 — ", "2026-02-05 协调发布,是同一「agent 攻势」的一体两面。", 0, "强", GREEN),

  h1("四、Codex 深度解析"),
  h2("4.1 身份澄清(2021 ≠ 2025)"),
  bul(null, "初代 Codex(2021,GPT-3 衍生)驱动首版 Copilot,API 已于 2023-03-23 关停。", 0, "强", GREEN),
  bul(null, "2025-05 重启为自主软件工程 agent(云 + CLI + IDE),由 codex-1(o3 的 RL 后训练)驱动。", 0, "强", GREEN),
  h2("4.2 模型谱系与基准(注意基准迁移)"),
  bul(null, "SWE-bench Verified:74.5%(5-Codex)→ 77.9%(5.1-Max)→ 88.7% 自报(GPT-5.5,独立 ~82.6%)。", 0, "强/中", AMBER),
  bul(null, "基准从 Verified 迁到 SWE-bench Pro / Terminal-Bench 2.0,跨版本不可直接续比。", 0, "强", GREEN),
  bul("更正 — ", "GPT-5.2-Codex 实为 2025-12-18(非 2026-01-14);GPT-5.4 确实存在(2026-03-05,非 Codex 专用)。", 0, "更正", AMBER),
  h2("4.3 Harness、沙箱与长时程自主"),
  bul("Rust 重写 — ", "零依赖、毫秒级启动、长会话无 GC 抖动、原生沙箱绑定;开源。", 0, "强", GREEN),
  bul("平台化沙箱 — ", "macOS=Seatbelt;Linux=bubblewrap+seccomp;Windows=受限令牌/私有桌面(+WSL)。"),
  bul("Compaction — ", "跨上下文窗口蒸馏关键状态,支撑长时程;5.1-Max 内部观察 24h+ 自主。", 0, "强/中", AMBER),
  bul("自适应推理 — ", "low/medium/high + xhigh;单模型自调步 ≠ 系统级路由。"),
  h2("4.4 计费、入口与编排"),
  bul("按 token 计费 — ", "Plus/Pro/Business 2026-04-02、Enterprise 04-23 起;捆绑各档不单售。", 0, "强", GREEN),
  bul("入口 — ", "Web/云、CLI、IDE、iOS、Slack、GitHub(@Codex PR 审查)。"),
  bul("编排 — ", "AGENTS.md(分层)、MCP(双向)、subagents;开源编排规范 Symphony。"),
  h2("4.5 采用与分发(均自报)"),
  bul(null, "周活轨迹更正:2 月 ~160 万 → 4 月初 300 万 → 6-02 500 万+;非开发者 ~20%、增速约 3 倍。", 0, "中", AMBER),
  bul(null, "2026-06-02 起前沿模型与 Codex 上架 AWS Bedrock;2026-04-27 微软–OpenAI 模型独家终结。", 0, "强", GREEN),

  h1("五、Frontier 深度解析"),
  h2("5.1 定义与定位"),
  bul(null, "2026-02-05 发布的端到端企业平台,构建/部署/管理 AI agent(AI coworkers);被框定为企业 agent 的「控制平面/操作系统」。", 0, "强", GREEN),
  h2("5.2 四大核心组件"),
  bul("业务上下文(语义层) — ", "连接数仓/CRM/工单/内部应用,共享「与人相同的信息」+ 机构记忆。"),
  bul("Agent 执行 — ", "推理、处理文件、跑代码、调工具;可跨本地/企业云/OpenAI 托管运行。"),
  bul("评估与优化 — ", "内建反馈回路,表现对人类管理者透明,随时间改进。"),
  bul("安全与治理 — ", "对人 + AI coworker 统一 IAM;受限身份、边界、可审计动作。", 0, "强", GREEN),
  h2("5.3 运营模型:像管员工一样管 agent"),
  bul(null, "借鉴企业「规模化人」:入职 → 传授机构知识 → 反馈中学习 → 授予受限权限与边界。", 0, "强", GREEN),
  h2("5.4 开放性与多供应商(已核验)"),
  bul(null, "兼容 OpenAI/企业自建/第三方(Google/MS/Anthropic)agent,经 MCP 集成,不强制重平台化。", 0, "强", GREEN),
  bul("中立悖论 — ", "宣称中立,却与自家 GPT-5.x 深绑;完整跨厂商互操作的成熟度仍待观察。", 0, "中", AMBER),
  bul("与 AgentKit/SDK 关系 — ", "不取代,而是把上下文+执行+评估统一进一个平台。"),
  h2("5.5 客户、成效、可用性与定价"),
  bul("早期客户 — ", "Uber、State Farm、Intuit、Thermo Fisher、HP、Oracle;试点 BBVA/Cisco/T-Mobile。", 0, "强/中", AMBER),
  bul("成效宣称 — ", "生产优化 6 周→1 天;能源 +5% 产出(增收 >$10 亿);均 OpenAI 自报、匿名、未审计。", 0, "中-弱", AMBER),
  bul("可用性/定价 — ", "非 GA(限量+试点);定价未公开、定制(FDE 模式)。", 0, "强/中", AMBER),

  h1("六、Codex × Frontier 综合与企业战略"),
  h2("6.1 互补关系"),
  bul(null, "GPT-5.3-Codex 作复杂 agentic 推理引擎;Codex 式编码 coworker 可作 Frontier 治理下的一类 agent。", 0, "中", AMBER),
  h2("6.2 企业营收战略"),
  bul(null, "企业收入 >40%→目标 ~50%(Friar);Q1 营收领先 Anthropic ~$10 亿,但 ARR 口径 Anthropic 更高——勿混用。", 0, "中", AMBER),
  h2("6.3 真实的落地组织(对前作的更正)"),
  bul(null, "Frontier Alliances(McKinsey/BCG/Accenture/Capgemini)、OpenAI Deployment Company(FDE,~$4B)、Codex Labs —— 均为真实计划。", 0, "强/中", AMBER),

  h1("七、竞争格局(两层)"),
  h2("7.1 编码工具层(JetBrains 2026,n>10,000)"),
  bul(null, "采用率 Copilot ~29% / Cursor ~18% / Claude Code ~18%(一年 ~6x);Claude Code CSAT 91% 居首。", 0, "强", GREEN),
  bul(null, "Cursor ~$2B ARR(传 $50B 估值);Google 退役 Gemini CLI 转 Antigravity;Cognition 收购 Windsurf→Devin Desktop。", 0, "强/中", AMBER),
  h2("7.2 企业 agent 平台层"),
  bul(null, "Frontier(多供应商控制平面)vs Salesforce Agentforce(18k+ 公司,锁定最紧)vs Microsoft Agent 365 / Frontier Suite(M365 E7,$99/用户)vs IBM watsonx Orchestrate vs ServiceNow。", 0, "强", GREEN),
  bul("框定 — ", "Frontier 以「overlay」模式与 SaaS 在位者「正面相撞」;在位者反击=重定价 + 自建治理层。", 0, "中", AMBER),

  h1("八、风险、批评与安全"),
  h2("8.1 Frontier:锁定与中立悖论"),
  bul(null, "批评者主张「LLM 中立的控制平面」(注:Okta/TrueFoundry 自身卖中立方案,有利益冲突);质疑 OpenAI 缺企业 know-how。", 0, "中", AMBER),
  h2("8.2 Codex:计费与网络安全分级"),
  bul(null, "按 token 计费引发限流抱怨;GPT-5.3-Codex 是首个被评「High」网络安全能力的模型,设分类器把高风险流量路由到 5.2。", 0, "强", GREEN),
  h2("8.3 供应链与漏洞事件"),
  bul(null, "恶意 npm 包 codexui-android 窃 ~/.codex token(Aikido 5-27 披露);命令注入漏洞泄露 GitHub token(已修)。", 0, "强", GREEN),
  h2("8.4 基准方法论"),
  bul(null, "88.7% 为自报(独立 ~82.6%);n=477 非全 500;pass@1=4 次平均;换更难基准就换冠军。", 0, "中", AMBER),

  h1("九、核验结论与重要更正"),
  bul("✗ 已剔除 — ", "「Codex 2.0 / Orchestrator / SafeDeploy / OpenRepo」无可信来源;真实编排规范 = Symphony。", 0, "证伪", "B91C1C"),
  bul("🔄 重要更正 — ", "Frontier Alliances / Deployment Company / Codex Labs 实为真实;GPT-5.2-Codex=2025-12-18;Codex 3M 周活在 4 月初;GPT-5.4 存在。", 0, "更正", AMBER),
  bul("~ 已降级 — ", "「省 4x token」为 n≈1 任务级;营收/利润率/成效/周活多为自报或单源。", 0, "降级", AMBER),
  bul("⚠ 三重撞名 — ", "OpenAI Frontier(平台)/ Microsoft Frontier Suite(M365 E7)/ OpenAI frontier risk(模型安全政策)。", 0, "易错", AMBER),

  h1("十、选型建议(按场景)"),
  bul("异步委派/并行批量/CI — ", "Codex(云):隔离容器 + worktree + 免审批自主更可控。"),
  bul("数据敏感/隔离/合规(编码) — ", "Claude Code:本地优先 + 应用层细粒度权限。"),
  bul("跨职能 agent 投产 + 治理 — ", "OpenAI Frontier:共享上下文 + 评估 + 权限/审计(早期访问、定价未公开)。"),
  bul("已重度绑定某 SaaS 生态 — ", "Agentforce / Copilot Studio / Agent 365:贴近业务与采购,但编排锁定更紧。"),
  bul("务实团队 — ", "双持 + 平台中立:Codex/Claude 跑编码;Frontier/在位者按生态选编排层。"),

  h1("十一、主要来源与方法学声明"),
  bul("OpenAI 官方 — ", "introducing-openai-frontier · business/frontier · introducing-gpt-5-3-codex · introducing-gpt-5-5 · open-source-codex-orchestration-symphony · frontier-alliance-partners · developers.openai.com/codex/*"),
  bul("主流媒体 — ", "TechCrunch · CNBC · Fortune · VentureBeat · Computerworld · SiliconANGLE · The Information · The Hacker News · SecurityWeek"),
  bul("调查/独立基准 — ", "JetBrains Research 2026 · Epoch AI · swebench.com · Vals.ai"),
  bul("声明 — ", "凡标中/弱者多为自报或单源;本次 WebFetch 受限,以多源搜索摘要交叉验证;最终数字以官方一手页面为准。", 0, "方法学", GRAY),
];

const doc = new Document({
  styles, numbering,
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    footers: { default: new Footer({ children: [ new Paragraph({
      alignment: AlignmentType.CENTER,
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: "D5D9DE", space: 6 } },
      children: [ T("OpenAI Codex × Frontier 洞察报告大纲   ·   ", { color: GRAY, size: 16 }),
                  new TextRun({ children: [PageNumber.CURRENT], font: "Arial", color: GRAY, size: 16 }) ],
    }) ] }) },
    children: [...titleBlock, ...toc, ...body],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  const out = "/home/user/openai-chatgpt-codex/OpenAI_Codex_Frontier_洞察报告大纲.docx";
  fs.writeFileSync(out, buf);
  console.log("saved:", out, "bytes:", buf.length);
});
