# OpenAI Codex × OpenAI Frontier — 深度研究报告

> **研究日期**:2026-06-22 · **方法**:5 角度并行检索(Codex 技术 / Frontier 平台 / 战略与商业 / 竞争格局 / 风险与核验)→ 去重抓取 → 对抗式核验(2/3 证伪即剔除)→ 按可信度分层综合。
> **可信度图例**:🟢 强(官方 + 多源交叉) · 🟡 中(单一媒体或厂商自报、无第三方审计) · 🔴 弱(单源/已降级) · ✗ 已剔除(证伪)。
> **环境限制**:本次对 `openai.com` 与多数主流媒体的 `WebFetch` 全程返回 **HTTP 403**;结论基于 WebSearch 摘要在 10+ 来源间交叉验证,以及可直读的 GitHub/个人博客。2026 年数据超模型知识截止(2026-01),正式引用前请对照官方 system card 与一手页面。

---

## 0. 执行摘要(七条核心结论)

1. **一个战略,两条战线。** 2026-02-05,OpenAI 同日推出 **GPT-5.3-Codex** 与 **Frontier**:Codex 是把「软件工程」一个垂直做到极致的自主编码 agent(纵深);Frontier 是把 agent 推广到全企业、并补齐治理/上下文/评估的「控制平面」(横向)。🟢
2. **Codex 技术凶猛迭代。** codex-1(o3 后训练)→ GPT-5-Codex → 5.1-Codex-Max(compaction、24h+ 自主)→ 5.2-Codex(**2025-12-18**)→ 5.3-Codex(2026-02-05)→ GPT-5.5(2026-04-23)。基准从 SWE-bench Verified 迁到 SWE-bench Pro / Terminal-Bench 2.0,跨版本不可直接续比。🟢
3. **Frontier = 企业 agent 的「操作系统」。** 四大组件:业务上下文(语义层)/ Agent 执行 / 评估与优化 / 安全与治理;核心叙事是「像管员工一样管 agent」(入职→反馈→权限)。**开放**:可管理 OpenAI、企业自建及第三方(Google/MS/Anthropic)agent。🟢
4. **企业营收引擎。** CFO Sarah Friar 称企业收入已占 OpenAI >40%、目标年底约 50%,Frontier 是「载体」;The Information 称 OpenAI 2026 Q1 营收约 $57–60 亿、季度领先 Anthropic ~$10 亿(均自报/单源,🟡)。
5. **格局是「两层」。** 编码工具层(Copilot 基数 29% / Cursor 营收 ~$2B ARR / Claude Code 口碑 CSAT 91%)+ 企业 agent 平台层(Frontier vs Agentforce / Microsoft Agent 365 / watsonx Orchestrate / ServiceNow)。🟢
6. **三重「Frontier」撞名,必须区分。** ① OpenAI **Frontier**(企业 agent 平台,2/5)② Microsoft **Frontier Suite**(= M365 E7,$99/用户,3/9)③ OpenAI「**frontier risk**」(前沿*模型*安全政策)。三者完全不同物。🟢
7. **对前作的重要更正。** 此前把「Frontier Alliances / Deployment Company(DeployCo) / Codex Labs」当作虚构——**本次核验证实它们是真实的 OpenAI 计划**;仅「Codex 2.0 / Orchestrator / SafeDeploy / OpenRepo」仍无可信来源、判为虚构。🟢

---

## 1. OpenAI Codex —— 产品、谱系与技术

### 1.1 身份澄清(2021 ≠ 2025)
- 初代 Codex(2021)是 GPT-3 衍生代码模型,驱动首版 GitHub Copilot;OpenAI **2023-03-23 关停其 API**。🟢
- 2025-05 重启的 Codex 是**自主软件工程 agent**(云 + CLI + IDE),由 **codex-1**(o3 的 RL 后训练变体)驱动。🟢
  - 来源:`openai.com/index/introducing-codex/`、`en.wikipedia.org/wiki/Codex_(AI_agent)`

### 1.2 模型谱系与基准(注意基准迁移)
| 日期 | 里程碑 | 谱系 | 关键基准(厂商自报) | 置信度 |
|---|---|---|---|---|
| 2025-05-16 | 云端 Codex 预览 | codex-1 ← o3 | SWE-bench ~70.2%(8 次 ~85%) | 🟡 |
| 2025-09-15 | GPT-5-Codex | ← GPT-5 | SWE-bench Verified **74.5%**;自适应思考时间;~7h 自主 | 🟢 |
| 2025-11-19 | GPT-5.1-Codex-Max | 更新基座 | SWE-bench **77.9%**、Terminal-Bench 2.0 58.1%;**compaction**;内部 **24h+** | 🟢(基准🟢/24h🟡) |
| **2025-12-18** | **GPT-5.2-Codex** | ← GPT-5.2 | SWE-bench Pro 56.4%、Terminal-Bench 2.0 64.0%;Windows 改进、安全增强 | 🟢(**日期更正**) |
| 2026-02-05 | GPT-5.3-Codex | 编码+推理统一 | SWE-bench Pro Public 56.8%、Terminal-Bench 2.0 77.3%;约快 25%;1M context | 🟡(77.3% 跳幅大,⚠) |
| 2026-04-23 | GPT-5.5 | 首个重训基座 | Terminal-Bench 2.0 82.7%、SWE-bench Verified **88.7%(自报)** | 🟢(日期)/🟡(基准) |

> **更正**:我此前的 PPT 把 GPT-5.2-Codex 标为 **2026-01-14**;多源(GIGAZINE、OpenAI system card update、媒体)确认实为 **2025-12-18**。
> **另**:**GPT-5.4** 确实存在(2026-03-05,Thinking/Pro;mini/nano 3-17),是面向企业/通用的重要模型,但**非专门的 Codex 变体**。
> 来源:`openai.com/index/introducing-gpt-5-2-codex/`、`gigazine.net/.../20251219-openai-gpt-5-2-codex/`、`techcrunch.com/2026/03/05/openai-launches-gpt-5-4-...`、`fortune.com/2026/03/05/...gpt5-4...`

### 1.3 Harness、沙箱与长时程自主
- **Rust 重写**(codex-rs,2025 末):零依赖安装、毫秒级启动、长会话无 GC 抖动、原生沙箱绑定;开源,npm `@openai/codex` / Homebrew 安装。🟢(`github.com/openai/codex`)
- **平台化沙箱**:macOS = Apple Seatbelt;Linux = bubblewrap + seccomp(自带 bwrap);Windows = 受限令牌 / 私有桌面隔离(+ WSL 跑 Linux 沙箱)。🟢(Windows 原生细节部分单源,🟡;`developers.openai.com/codex/concepts/sandboxing`)
- **Compaction**(5.1-Codex-Max 头条):接近窗口上限时把关键状态蒸馏为摘要带入新窗口,支持跨窗口/长时程;所有 5.1-Max 基准均在 compaction 开启下测得。🟢
- **自适应推理**:模型按难度「自调步」(简单回合省 token、难题多花算力),离散档 low/medium/high,5.1-Max 起新增 **xhigh**;此为单模型自定推理,**≠ 系统级快/慢路由**。🟢/🟡

### 1.4 计费、入口与编排
- **按 token 计费**:Plus/Pro/Business **2026-04-02** 起,Enterprise **2026-04-23** 起;按百万 input/cached/output token 计信用;捆绑 Free/Go/Plus/Pro/Business/Edu/Enterprise,不单售。🟢(`help.openai.com/.../codex-rate-card`)
- **入口**:Web/云、CLI、IDE 扩展、iOS、Slack、GitHub(`@Codex` PR 审查)。🟢
- **编排**:AGENTS.md(分层指令)、MCP(client/server 双向)、subagents(共享工作区、独立线程);开源编排规范 **Symphony**(把 Linear 等看板变成 Codex 控制平面)。🟢(`openai.com/index/open-source-codex-orchestration-symphony/`)

### 1.5 采用(均为 OpenAI 自报,🟡)
- 周活轨迹(更正版):~60 万(年初)→ ~160 万(2 月,桌面 App 后)→ 200 万(~3 月初)→ **300 万(~4-08)** → 400 万(4-21)→ **500 万+(6-02)**。
  > **更正**:此前简报的「2 月 300 万」有误——300 万是 **4 月初** 达到,2 月约 160 万。
- 非开发者占比 ~20%(分析/研究/报告),增速约为工程师 3 倍;OpenAI 内部 ~95% 工程师周用 Codex、PR 量 +~70%。
- 来源:`constellationr.com/.../5-million-weekly-active-users`、`fortune.com/2026/03/04/openai-codex-growth-...`、`openai.com/index/codex-for-knowledge-work/`
- **新分发**:2026-06-02 起,OpenAI 前沿模型与 Codex 上架 **AWS Bedrock**。🟢(`openai.com/index/openai-frontier-models-and-codex-are-now-available-on-aws/`)

---

## 2. OpenAI Frontier —— 企业 AI Agent 平台

### 2.1 定义与定位
- **2026-02-05 发布**,端到端企业平台,用于**构建/部署/管理 AI agent(AI coworkers)**;被广泛框定为企业 agent 的「**控制平面 / 操作系统 / 编排器**」。🟢
  - 来源:`openai.com/index/introducing-openai-frontier/`、`openai.com/business/frontier/`、`techcrunch.com/2026/02/05/...`、`cnbc.com/2026/02/05/...`、`computerworld.com/article/4135372/...`

### 2.2 四大核心组件
1. **业务上下文 / Business Context(语义层)** — 连接数仓、CRM、工单、内部应用,让 agent 共享「与人相同的信息」并建立持久机构记忆。🟢
2. **Agent 执行 / Agent Execution** — 推理、处理文件、跑代码、调工具;可在本地、企业云、OpenAI 托管运行时执行。🟢
3. **评估与优化 / Evaluation & Optimization** — 内建反馈回路,使 agent 表现对人类管理者透明并随时间改进。🟢
4. **安全与治理 / Security & Governance** — 对「人类员工 + AI coworker」统一 IAM;agent 获得受限身份/权限、明确边界、可审计动作。🟢
   - 来源:`openai.com/business/frontier/`、`the-decoder.com/openais-frontier-gives-ai-agents-employee-like-identities-...`、`datacamp.com/blog/openai-frontier`

### 2.3 运营心智模型:「像管员工一样管 agent」
- OpenAI 明确借鉴企业「规模化人」的做法:入职、传授机构知识/内部语言、在反馈中学习、授予受限访问与边界。🟢(`openai.com/index/introducing-openai-frontier/`)

### 2.4 开放性与多供应商(已核验)
- **CONFIRMED**:Frontier 管理 **OpenAI / 企业自建 / 第三方(Google、Microsoft、Anthropic)** 的 agent,经开放标准(**MCP**)集成,不强制重平台化。OpenAI 表态「无法构建企业需要的每一个 agent」。🟢
  - 来源:`aibusinessweekly.net/.../openai-frontier-...`、`venturebeat.com/orchestration/openai-launches-centralized-agent-platform-...`、`futurumgroup.com/insights/openai-frontier-...`
  - **caveat**:开放是「定位+早期能力」,完整跨厂商互操作的成熟度仍待观察(VentureBeat 早前曾对「是否真正开放」存疑)。🟡
- **与 AgentKit/SDK 的关系**:OpenAI 明确 Frontier **不取代** Agents SDK / AgentKit / API——AgentKit(含 Agent Builder、Connector Registry、ChatKit)是「开发构建」工具包,Frontier 是「部署/管理/治理」层。🟢(`venturebeat.com/...`、`openai.com/index/introducing-agentkit/`)

### 2.5 客户、成效、可用性与定价
- **早期客户(多源一致,🟢)**:Uber、State Farm、Intuit、Thermo Fisher Scientific、HP、Oracle;**试点(🟡)**:BBVA、Cisco、T-Mobile。
  - 来源:`openai.com/index/introducing-openai-frontier/`、`cnbc.com/2026/02/05/...`、`fortune.com/2026/02/05/...`、`artificialintelligence-news.com/news/intuit-uber-and-state-farm-...`
- **成效宣称(OpenAI 自报、匿名、未审计,🟡-🔴)**:某制造商「生产优化」6 周 → 1 天;某能源企业产出最高 +5%(≈ 增收 >$10 亿);某投资公司销售对客时间释放 >90%。→ 按**厂商精选案例**对待。
- **可用性**:**非 GA**——限量客户 + 试点,「未来数月」逐步放开。🟢
- **定价**:未公开,定制(企业销售 / Forward Deployed Engineer 模式)。🟢/🟡(`openai.com/business/frontier/`、`eesel.ai/blog/openai-frontier-pricing`)
- **与 ChatGPT Enterprise 区分**:Enterprise 是对话/生产力应用;Frontier 是部署/治理 agent 群的基础设施层。「ChatGPT Enterprise 叠在 Frontier 之上」的说法来自二手博客,**未在官方原文确认**,⚠。🟡

---

## 3. Codex × Frontier 关系 + OpenAI Agent 栈 + 企业战略

### 3.1 OpenAI Agent 栈(本报告分析框架)
```
⑤ 分发层    ChatGPT「超级应用」(ChatGPT + Codex + agentic browsing + 合作伙伴应用)
④ 编排治理   OpenAI Frontier —— 全企业 AI coworker 控制平面(横向 + 治理)
③ 专用执行   Codex —— 软件工程垂直的自主编码 coworker(纵深)
② 构建层    AgentKit / Agents SDK(Agent Builder · Connector Registry · ChatKit)
① 模型层    GPT-5.x 家族(含 GPT-5.3-Codex 推理引擎)
```
- **同日推进**:2026-02-05 GPT-5.3-Codex + Frontier 一并发布,是协调的「agent 攻势」。🟢
- **互补**:GPT-5.3-Codex 作为复杂 agentic 推理引擎;Codex 式编码 coworker 可作为 Frontier 治理下的一类 agent 被部署/监控/评估。🟡(「engine/fleet」框定为分析语,组件清单为官方)
- 来源:`siliconangle.com/2026/02/05/openai-introduces-frontier-agent-management-platform-gpt-5-3-codex/`、`openai.com/index/next-phase-of-enterprise-ai/`

### 3.2 企业营收战略(均自报/单源,🟡)
- Sarah Friar:企业收入 >40%、目标年底 ~50%,Frontier 为载体。(`decrypt.co/363844/...`)
- The Information:OpenAI 2026 Q1 营收约 $57–60 亿、季度领先 Anthropic ~$10 亿,Codex 为关键驱动;另有「Q1 调整后经营利润率 ≈ −122%」的衍生说法(单一底层来源,🔴)。(`theinformation.com/.../openai-held-1-billion-revenue-lead-anthropic-...`、`pymnts.com/.../openais-codex-helps-drive-nearly-6-billion-quarter/`)
- **口径提醒**:OpenAI「季度营收」领先,但 ARR/run-rate 口径下 Anthropic 数字更高(~$30B run-rate,4 月);**勿混用季度营收与 ARR**。🟡

### 3.3 真实的企业落地组织(对前作的关键更正,🟢)
- **Frontier Alliances**:真实的咨询伙伴计划(McKinsey、BCG、Accenture、Capgemini)。(`openai.com/index/frontier-alliance-partners/`、`mckinsey.com/.../mckinsey-and-openai-...frontier-alliance`)
- **OpenAI Deployment Company(俗称「DeployCo」)**:真实——Forward Deployed Engineers 模式,~$4B 初始投入。(`openai.com/index/openai-launches-the-deployment-company/`)
- **Codex Labs**:真实但来源较薄——借集成商(Accenture、Capgemini、Cognizant、Infosys、PwC、TCS)推进 Codex 落地。🟡(`techgenyz.com/openai-partner-network-enterprise-ai-consultants/`)
> 这三者在我此前的 PPT 中被误列为「虚构、已剔除」。**本次多源核验予以更正:它们是真实计划。**

---

## 4. 竞争格局(两层)

### 4.1 编码工具层(JetBrains 2026 调查,n>10,000,🟢)
- 职场采用率:**GitHub Copilot ~29%**(增长停滞)· **Cursor ~18%** · **Claude Code ~18%**(一年内从 ~3% 增至 18%,~6x)。
- 满意度:**Claude Code CSAT 91% / NPS 54**(品类最高)。
- 营收:**Cursor ~$2B ARR(2026-02)**,传 ~$50B 估值募资(未定,🟡);**Claude Code 据报 ~$2.5B run-rate**(博客源,🟡),Anthropic 整体 ~$30B run-rate(4 月,🟢)。
- 2026 格局剧变:**微软–OpenAI 模型独家于 2026-04-27 结束**(GPT-5.5 次日上 AWS Bedrock);Copilot 转多模型(并推 Claude,自研「Polaris」🟡);**Google 退役 Gemini CLI 转 Antigravity(CLI 6-18 停服)**;**Cognition(Devin)2025-07-14 收购 Windsurf**,2026-06-02 更名「Devin Desktop」。🟢
  - 来源:`blog.jetbrains.com/research/2026/04/...`、`blogs.microsoft.com/blog/2026/04/27/...`、`developers.googleblog.com/an-important-update-transitioning-gemini-cli-to-antigravity-cli/`、`techcrunch.com/2025/07/14/cognition-...-acquires-windsurf/`

### 4.2 企业 Agent 平台层
| 平台 | 归属/定位 | 关键事实 | 锁定/软肋 |
|---|---|---|---|
| **OpenAI Frontier** | 多供应商「控制平面」+ 语义层 | 2/5 发布,非 GA,定制定价 | 企业经验浅;与自家 GPT-5.x 深绑的「中立悖论」 |
| **Salesforce Agentforce** | 贴近 CRM | 18,000+ 公司 / 121 国;ARR ~$800M(+169% YoY) | Atlas 编排锁定最紧 |
| **Microsoft Agent 365 / Frontier Suite** | M365 治理层 | **Frontier Suite = M365 E7,$99/用户(5/1)**;Agent 365 单售 $15/用户 | 与 OpenAI 既合作又竞争;**名称撞名** |
| **IBM watsonx Orchestrate** | 受监管行业 | Think 2026(5/5)重定位为多厂商「agentic 控制平面」 | 模型前沿性/生态偏弱 |
| **ServiceNow(AI Control Tower)** | ITSM/工作流就近 | 集中式 agent 治理;与 OpenAI 多年协议 | 跨域通用性有限 |
- **框定**:Frontier/Anthropic 以「overlay」模式把 agent 推上与 SaaS 在位者「正面相撞」的轨道(Fortune/AIN);Salesforce 股价被指 YTD 下挫(精确「−27%」为单源,🔴)。🟡
  - 来源:`blogs.microsoft.com/blog/2026/03/09/introducing-the-first-frontier-suite-...`、`newsroom.ibm.com/2026-05-05-think-2026...`、`fortune.com/2026/02/05/...saas-salesforce-workday/`

---

## 5. 风险、批评与安全

### 5.1 Frontier:锁定与中立悖论
- **LLM 厂商锁定**:批评者认为不该让单一厂商同时拥有身份、数据、模型路由、编排、可观测的控制平面;主张「LLM 中立的控制平面」。🟡(注意:Okta、TrueFoundry 等最大声的批评者**自身在卖中立方案**,存在利益冲突。)(`infoq.com/news/2026/02/openai-frontier-agent-platform/`、`okta.com/blog/ai/openai-agentic-ai-operational-gap/`)
- **企业经验**:Salesforce/ServiceNow/Workday/SAP 认为 OpenAI 缺乏对系统记录、访问控制、治理的多年积累,须先证明编排能力。🟢(`fortune.com/2026/02/05/...`)
- **保护 agent 而非仅 API**:agentic 部署的安全是未解运营缺口,建议把 agent 作为身份系统中的一等身份并授予委派权限。🟡

### 5.2 Codex:计费与网络安全分级
- **按 token 计费上线引发不满**(2026-04;限流抱怨集中,4-28 各付费档限额重置)。🟢/🟡(`help.openai.com/.../codex-rate-card`、`community.openai.com/...`)
- **GPT-5.3-Codex 是首个被 OpenAI 按 Preparedness Framework 评为「High」网络安全能力**的模型,称其「若被自动化/规模化或足以放大现实网络危害」;缓解含拒答训练 + 分类器把高风险流量路由到能力较弱的 GPT-5.2。🟢(`openai.com/index/introducing-gpt-5-3-codex/`、`developers.openai.com/codex/concepts/cyber-safety`、`fortune.com/2026/02/05/openai-gpt-5-3-codex-warns-...`)

### 5.3 Codex:供应链与漏洞事件
- **恶意 npm 包 `codexui-android`**(伪装成 Codex 远程 UI,~2.7–2.9 万周下载)窃取 `~/.codex/auth.json` token 约一个月;Aikido Security 2026-05-27 披露;OpenAI refresh token 不过期 → 可长期冒充。🟢(`aikido.dev/blog/codex-remote-ui-steals-ai-tokens`、`thehackernews.com/2026/06/openai-codex-authentication-tokens.html`)
- **关联恶意 Android 应用**「OpenClaw …」(下载量各源不一,🟡)。
- **命令注入漏洞**(GitHub 分支名未净化,`${IFS}`+U+3000+`|| true` 隐藏 payload),可泄露容器内明文 GitHub token;2025-12-16 报告,~2026-02-05 修复,无恶意利用证据。🟢(`securityweek.com/critical-vulnerability-in-openai-codex-...`)

### 5.4 基准方法论(谨慎看「谁第一」)
- GPT-5.5「**88.7% SWE-bench Verified」为自报**,独立/第三方测得约 **~82.6–82.7%**(~6 分差;精确独立值多来自二手博客,方向成立、数值谨慎)。🟡
- OpenAI 在 **n=477**(非全 500)子集上算 SWE-bench Verified,且「**pass@1 = 每例 4 次平均**」。🟢(`epoch.ai/benchmarks/swe-bench-verified`、`swebench.com/verified.html`)
- 「换更难基准就换冠军」:有博客称 GPT-5.5 在 SWE-bench Pro 上落后 Claude Opus 4.7 ~5.7 分(单源,🔴)。

---

## 6. 核验结果与对抗式裁决

### ✗ 已剔除(证伪 / 无可信来源)
- **「Codex 2.0 / Orchestrator / SafeDeploy / OpenRepo」** —— 无任何可信来源;真实的开源编排规范名为 **Symphony**。✗

### 🔄 重要更正(此前误判)
- **「Frontier Alliances / Deployment Company(DeployCo) / Codex Labs」** —— **真实** OpenAI 计划(见 §3.3),**撤销此前「虚构」判定**。
- **GPT-5.2-Codex 日期** —— **2025-12-18**(非 2026-01-14)。
- **Codex「2 月 300 万周活」** —— 实为 **4 月初** 达 300 万,2 月约 160 万。
- **GPT-5.4** —— **确实存在**(2026-03-05),非 Codex 专用变体。

### ~ 已降级(传闻 / 自报 / 单源)
- 「Codex 比 Claude Code 省 **4x token**」—— n≈1 任务级对比(如 Figma-to-code 1.5M vs 6.2M),非受控研究;且反方指出 Claude 多花 token 换取更多推理质量。🔴
- **营收 / 利润率 / 客户成效 / 周活** —— 多为 OpenAI 自报或单一媒体(The Information),无第三方审计。🟡-🔴

### ⚠ 三重「Frontier」撞名(务必区分)
1. **OpenAI Frontier** = 企业 agent 平台(2026-02-05)。
2. **Microsoft Frontier Suite** = M365 E7 套餐($99/用户,2026-03-09 宣布、5-01 起售)。
3. **OpenAI「frontier risk」/ Preparedness Framework** = 前沿*模型*的灾难性风险安全政策(CBRN/网络/自主等),与平台无关。(`openai.com/global-affairs/our-approach-to-frontier-risk/`)

---

## 7. 主要来源(分层节选)

**OpenAI 官方**:`introducing-codex` · `introducing-upgrades-to-codex` · `gpt-5-1-codex-max` · `introducing-gpt-5-2-codex` · `introducing-gpt-5-3-codex` · `introducing-gpt-5-5` · `introducing-gpt-5-4` · `open-source-codex-orchestration-symphony` · `introducing-openai-frontier` · `business/frontier` · `introducing-agentkit` · `next-phase-of-enterprise-ai` · `frontier-alliance-partners` · `openai-launches-the-deployment-company` · `openai-frontier-models-and-codex-are-now-available-on-aws` · `next-phase-of-microsoft-partnership` · `global-affairs/our-approach-to-frontier-risk` · `developers.openai.com/codex/*` · `help.openai.com/.../codex-rate-card`

**主流媒体**:TechCrunch · CNBC · Fortune · VentureBeat · Computerworld · SiliconANGLE · The Information · PYMNTS · The Decoder · The Register · The Hacker News · SecurityWeek · CSO Online · Axios

**官方调查 / 独立基准**:JetBrains Research 2026(`blog.jetbrains.com/research/2026/04/...`)· Epoch AI(`epoch.ai/benchmarks/swe-bench-verified`)· `swebench.com` · Vals.ai

**厂商/安全研究**:Aikido Security · Okta · IBM Newsroom · Microsoft Blog · Google Developers Blog · Cognition Blog · McKinsey

**技术博客 / 社区(较低置信)**:simonwillison.net · codex.danielvaughan.com · digitalapplied.com · datacamp.com · nxcode.io · futurumgroup.com · constellationr.com · infoq.com · rierino.com · mindstudio.ai

---

### 置信度与方法学声明
本报告以 5 角度并行检索 + 对抗式核验生成;凡标 🟡/🔴 者多为厂商自报或单源,引用前请二次核对。本次 `WebFetch` 对 `openai.com`/主流媒体全程 403,结论基于 WebSearch 摘要的多源交叉验证;2026 年数据超模型知识截止,最终数字以官方 system card 与一手页面为准。
