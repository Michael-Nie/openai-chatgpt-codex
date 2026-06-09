# -*- coding: utf-8 -*-
"""Build the OpenAI Codex deep-insight deck (PPTX)."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

# ---------- palette ----------
NAVY   = RGBColor(0x0E, 0x14, 0x2E)   # deep background
INK    = RGBColor(0x1F, 0x29, 0x37)   # body text
GREEN  = RGBColor(0x10, 0xA3, 0x7F)   # OpenAI-ish accent
TEAL   = RGBColor(0x0B, 0x7A, 0x75)
GRAY   = RGBColor(0x6B, 0x72, 0x80)
LGRAY  = RGBColor(0xE5, 0xE7, 0xEB)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
AMBER  = RGBColor(0xB4, 0x53, 0x09)
RED    = RGBColor(0xB9, 0x1C, 0x1C)
CARD   = RGBColor(0xF3, 0xF5, 0xF7)

FONT = "Microsoft YaHei"

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]

def set_font(run, name=FONT):
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ('a:latin', 'a:ea', 'a:cs'):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set('typeface', name)

def rect(slide, x, y, w, h, color, line=None):
    sp = slide.shapes.add_shape(1, x, y, w, h)  # rectangle
    sp.fill.solid(); sp.fill.fore_color.rgb = color
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line; sp.line.width = Pt(1)
    sp.shadow.inherit = False
    return sp

def textbox(slide, x, y, w, h):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame; tf.word_wrap = True
    return tb, tf

def style(p, text, size, color=INK, bold=False, align=PP_ALIGN.LEFT, space_after=4, italic=False):
    p.alignment = align
    p.space_after = Pt(space_after)
    p.space_before = Pt(0)
    r = p.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = color
    set_font(r)
    return r

def header(slide, kicker, title):
    rect(slide, 0, 0, SW, Inches(1.18), NAVY)
    rect(slide, 0, Inches(1.18), SW, Pt(3), GREEN)
    tb, tf = textbox(slide, Inches(0.55), Inches(0.16), Inches(12.2), Inches(0.95))
    p = tf.paragraphs[0]; style(p, kicker, 11, GREEN, bold=True, space_after=2)
    p2 = tf.add_paragraph(); style(p2, title, 24, WHITE, bold=True)

def footer(slide, n, note=""):
    tb, tf = textbox(slide, Inches(0.55), Inches(7.06), Inches(11.0), Inches(0.36))
    p = tf.paragraphs[0]
    style(p, note, 8, GRAY)
    tb2, tf2 = textbox(slide, Inches(12.4), Inches(7.06), Inches(0.7), Inches(0.36))
    p2 = tf2.paragraphs[0]
    style(p2, str(n), 9, GRAY, align=PP_ALIGN.RIGHT)

def content_slide(kicker, title, n, note="可信度: 强=官方+多源 · 中=单一媒体/厂商自报 · 弱=已剔除/降级"):
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, SW, SH, WHITE)
    header(s, kicker, title)
    footer(s, n, note)
    return s

def bullets(slide, items, x=Inches(0.6), y=Inches(1.5), w=Inches(12.1), h=Inches(5.4),
            size=14, gap=7):
    tb, tf = textbox(slide, x, y, w, h)
    first = True
    for it in items:
        lvl = it.get("lvl", 0)
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.level = lvl
        p.space_after = Pt(it.get("gap", gap)); p.space_before = Pt(0)
        p.alignment = PP_ALIGN.LEFT
        marks = {0: "▍ ", 1: "•  ", 2: "–  "}
        prefix = it.get("bullet", marks.get(lvl, "•  "))
        col = it.get("color", INK if lvl == 0 else GRAY if lvl >= 2 else INK)
        sz = it.get("size", size if lvl == 0 else size-2)
        # prefix run
        rp = p.add_run(); rp.text = prefix
        rp.font.size = Pt(sz); rp.font.bold = (lvl == 0)
        rp.font.color.rgb = (GREEN if lvl == 0 else col)
        set_font(rp)
        # label (bold lead)
        if "lead" in it:
            rl = p.add_run(); rl.text = it["lead"]
            rl.font.size = Pt(sz); rl.font.bold = True; rl.font.color.rgb = col
            set_font(rl)
        rt = p.add_run(); rt.text = it["t"]
        rt.font.size = Pt(sz); rt.font.bold = it.get("bold", lvl == 0 and "lead" not in it)
        rt.font.color.rgb = col
        set_font(rt)
        if "tag" in it:
            rg = p.add_run(); rg.text = "  〔" + it["tag"] + "〕"
            rg.font.size = Pt(sz-2); rg.font.italic = True
            rg.font.color.rgb = it.get("tagc", GRAY)
            set_font(rg)
    return tb

def table_slide(kicker, title, n, headers, rows, col_w, note="", fs=11, hfs=11,
                y=Inches(1.45), row_h=Inches(0.42)):
    s = content_slide(kicker, title, n, note)
    nrows, ncols = len(rows)+1, len(headers)
    total_w = sum(col_w)
    x = Inches((13.333 - total_w)/2)
    gtab = s.shapes.add_table(nrows, ncols, x, y, Inches(total_w), row_h*nrows)
    tbl = gtab.table
    # disable banding style by setting first row only
    tbl.first_row = True; tbl.horz_banding = True
    for j, cw in enumerate(col_w):
        tbl.columns[j].width = Inches(cw)
    # header
    for j, htext in enumerate(headers):
        c = tbl.cell(0, j)
        c.fill.solid(); c.fill.fore_color.rgb = NAVY
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        c.margin_left = Inches(0.07); c.margin_right = Inches(0.05)
        c.margin_top = Inches(0.02); c.margin_bottom = Inches(0.02)
        p = c.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
        r = p.add_run(); r.text = htext; r.font.size = Pt(hfs); r.font.bold = True
        r.font.color.rgb = WHITE; set_font(r)
    # body
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            c = tbl.cell(i, j)
            c.fill.solid(); c.fill.fore_color.rgb = WHITE if i % 2 else CARD
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            c.margin_left = Inches(0.07); c.margin_right = Inches(0.05)
            c.margin_top = Inches(0.02); c.margin_bottom = Inches(0.02)
            p = c.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
            txt = val; col = INK; bold = (j == 0)
            if isinstance(val, tuple):
                txt, col = val[0], val[1]
                bold = val[2] if len(val) > 2 else (j == 0)
            r = p.add_run(); r.text = txt; r.font.size = Pt(fs)
            r.font.bold = bold; r.font.color.rgb = col; set_font(r)
    return s

# ============================================================ SLIDE 1 — COVER
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, SW, SH, NAVY)
rect(s, 0, Inches(4.05), SW, Pt(3.5), GREEN)
rect(s, Inches(0.0), 0, Inches(0.22), SH, GREEN)
tb, tf = textbox(s, Inches(0.9), Inches(1.5), Inches(11.5), Inches(2.6))
p = tf.paragraphs[0]; style(p, "OpenAI CODEX", 17, GREEN, bold=True, space_after=6)
p = tf.add_paragraph(); style(p, "深度技术洞察分析", 46, WHITE, bold=True, space_after=4)
p = tf.add_paragraph(); style(p, "Product · Architecture · Competition · Strategy · Developer Voice", 15, LGRAY)
tb, tf = textbox(s, Inches(0.9), Inches(4.35), Inches(11.5), Inches(2.2))
for t in [
    "产品形态 · 模型与训练机制 · Agent Harness 与沙箱 · Benchmark 方法论 · 架构级竞品对比 · 商业与口碑",
]:
    p = tf.add_paragraph(); style(p, t, 13, LGRAY, space_after=10)
p = tf.add_paragraph(); style(p, "2026 年 6 月   |   5 路并行检索 + 3 路对抗式核验(2/3 证伪即剔除)", 12, GREEN, bold=True, space_after=4)
p = tf.add_paragraph(); style(p, "方法学提示:本次环境 WebFetch 全程受限,结论基于权威页面的搜索摘要并经多源交叉验证;2026 年数据超模型知识截止,引用前建议对照官方 system card。", 9.5, GRAY)

# ============================================================ SLIDE 2 — EXEC SUMMARY
s = content_slide("EXECUTIVE SUMMARY", "执行摘要:五个核心洞察", 2)
bullets(s, [
 {"lead":"品牌是三代复用,切勿混淆 — ", "t":"2021 初代模型(驱动 Copilot,API 已于 2023-03-23 关停)≠ 2025 重启的「云 + CLI + IDE」自主软件工程 agent。本报告聚焦后者。", "tag":"强"},
 {"lead":"技术主线凶猛迭代 — ", "t":"codex-1(o3)→ GPT-5-Codex → 5.1-Codex-Max(compaction / 24h+)→ 5.2 / 5.3 → 5.5;能力从「结对」走向「长时程自主 + 多 agent 编排」。", "tag":"强"},
 {"lead":"真正壁垒是「模型×harness×沙箱」三位一体 — ", "t":"RL 让模型与 agent 循环协同设计;Rust harness 五工具循环走 Responses API;OS 内核级沙箱使「免审批自主」可控。这是 Codex 的工程内核。", "tag":"强"},
 {"lead":"商业上已成增长引擎 — ", "t":"据 The Information,2026 Q1 OpenAI 营收近 $60 亿「受 Codex 推动」,季度领先 Anthropic ~$10 亿;周活 200 万 → 500 万+(~20% 为非开发者)。", "tag":"中"},
 {"lead":"格局是「三王分治 + 多工具叠用」 — ", "t":"Copilot 赢用户基数、Cursor 赢营收(~$2B ARR)、Claude Code 赢口碑;Codex 差异化=长时程沙箱自主 + 顶尖跑分,软肋=锁定 OpenAI 模型。共识是「双持」而非二选一。", "tag":"强"},
 {"t":"关键风险:版本/跑分极易张冠李戴,任何能力数字都须绑定「模型版本 + 日期 + 评测 harness」才有意义。", "lvl":1, "color":AMBER, "bullet":"⚠  "},
], y=Inches(1.5), size=14.5, gap=13)

# ============================================================ SLIDE 3 — METHOD & CONFIDENCE
s = content_slide("METHODOLOGY", "研究方法与可信度框架", 3)
bullets(s, [
 {"lead":"流程 — ", "t":"将问题拆为 5 个角度并行检索 → 去重抓取 → 对每条关键论断做 3 票对抗式核验(2/3 证伪即剔除)→ 按可信度排序综合。"},
 {"lead":"可信度分层 — ", "t":"强(官方 + 多源交叉);中(单一媒体或 OpenAI 自报、无第三方审计);弱(已剔除或降级处理)。"},
 {"t":"核验结果:已证伪并剔除", "lvl":1, "color":RED, "bullet":"✗  ", "bold":True},
 {"t":"「Codex 2.0 / Orchestrator / SafeDeploy / OpenRepo」 — 仅单一来源(windowsnews.ai),与 OpenAI 实际命名(Symphony、subagents)矛盾,判为不实。", "lvl":2, "color":GRAY},
 {"t":"核验结果:已降级为传闻", "lvl":1, "color":AMBER, "bullet":"~  ", "bold":True},
 {"t":"「Codex 比 Claude Code 省 4x token」 — 源自单次 Figma-to-code 测试(1.5M vs 6.2M),n=1 且含「产出质量不同」的混淆;方向或成立,非严谨基准。", "lvl":2, "color":GRAY},
 {"t":"「GPT-5.5 SWE-bench 88.7%」为 OpenAI 自报;独立 harness 仅 ~82.6%。", "lvl":2, "color":GRAY},
 {"lead":"核心限制 — ", "t":"本次 WebFetch 对 openai.com/主流媒体全程 403,结论基于搜索摘要(GitHub 部分 README、badlogic compaction 研究可直读);2026 年项超知识截止(2026-01)。发布前建议对照原始 system card PDF。", "color":AMBER, "tag":"重要"},
], y=Inches(1.45), size=13, gap=8)

# ============================================================ SLIDE 4 — PRODUCT SURFACES
s = content_slide("PRODUCT", "产品形态全景:一个 agent,六个入口", 4)
bullets(s, [
 {"lead":"多端形态 — ", "t":"云端 agent(chatgpt.com/codex)· Codex CLI(开源 Apache-2.0,Rust)· VS Code/IDE 扩展 · 桌面 App · GitHub 集成 · Slack · iOS 移动端(2026-05)。"},
 {"lead":"三种交互范式 — ", "t":"CLI 重隐私/控制与 CI/CD;IDE 在编辑器内最小干扰;云端做异步、并行、可委派的多文件任务。2026 主流是「混合」:App 作指挥中心规划/审查 + CLI/IDE 本地落地。"},
 {"lead":"关键特性 — ", "t":"审批模式 × 沙箱模式 双层安全;AGENTS.md 项目指令(分层覆盖);MCP(既是 client 又是 server);PR 代码审查(仅报 P0/P1);子 agent 并行编排。"},
 {"lead":"计费 — ", "t":"不单独售卖,捆绑进 Plus($20)/Pro($200)/Business/Enterprise/Edu;2026-04-02 由「按消息」改为「按 token」信用计费。"},
 {"lead":"规模(截至 2026-06)— ", "t":"周活 500 万+,其中 ~20% 为非开发者(分析/市场/运营),非开发者增速约为工程师的 3 倍。", "tag":"中"},
], y=Inches(1.55), size=14, gap=13)

# ============================================================ SLIDE 5 — MODEL LINEAGE / TIMELINE
table_slide(
 "MODEL LINEAGE & TIMELINE", "模型谱系与演进时间线", 5,
 ["日期", "里程碑", "基座 / 谱系", "关键技术点", "SWE-bench Verified"],
 [
  ["2025-04-16", "Codex CLI 开源", "—", "终端 agent;发布即 TS/Node", "—"],
  ["2025-05-16", "云端 Codex 预览", ("codex-1 ← o3", TEAL), "RL 训练于真实编码任务;codex-mini ← o4-mini", "72.1%"],
  ["2025-09-15", "GPT-5-Codex", ("← GPT-5", TEAL), "自适应「思考时间」;代码审查训练", "74.5%"],
  ["2025-11-19", "GPT-5.1-Codex-Max", ("← 更新基座", TEAL), "原生 compaction 跨上下文窗口;内部 24h+", ("76.5% / 77.9%(xhigh)", INK, True)],
  ["2026-01-14", "GPT-5.2-Codex", ("← GPT-5.2", TEAL), "Terminal-Bench 2.0 新高", "—(转向 SWE-Pro)"],
  ["2026-02-05", "GPT-5.3-Codex", ("统一编码+推理", TEAL), "约快 25%;附严格网络安全管控", "~80% / Pro 56.8%"],
  ["2026-04-23", "GPT-5.5", ("首个重训基座", TEAL), "Codex 400K / API 1M 上下文", ("88.7%(厂商自报)", AMBER, True)],
 ],
 col_w=[1.6, 2.25, 2.2, 4.0, 2.55],
 note="谱系:Codex 模型均为基座推理模型的后训练变体,非独立架构 · 2025-09 数字为强证据;2026 项依赖检索",
 fs=10.5, hfs=10.5, row_h=Inches(0.72), y=Inches(1.42))

# ============================================================ SLIDE 6 — RL TRAINING
s = content_slide("TECH DEEP-DIVE 1/6 · TRAINING", "技术深钻①:RL 训练方法论 —「人类风格的 PR」从何而来", 6)
bullets(s, [
 {"lead":"本质 — ", "t":"codex-1 是 o3 的后训练变体,用强化学习在「多环境的真实编码任务」上训练:写功能/加测试/调试/大规模重构/做代码审查。"},
 {"lead":"奖励信号(可验证 + 偏好)— ", "t":"①测试通过(可验证奖励,模型迭代跑测直到通过)②人类编码风格 ③PR/提交规范 ④指令遵从 ⑤代码审查偏好(由资深工程师对真实 OSS commit 的评审打分)。"},
 {"lead":"「人类风格 PR」≠ 新架构 — ", "t":"它是把输出对齐到上述奖励的结果:相比 o3,codex-1「产出更干净、可直接 review 的补丁」。", "tag":"强"},
 {"lead":"代码审查的可量化收益 — ", "t":"GPT-5-Codex 的错误评审意见比 GPT-5 少约 70%(资深工程师在真实 commit 上评估);OpenAI 内部已让 Codex 在人类 reviewer 之前自动审查大多数变更。", "tag":"强"},
 {"lead":"营销 vs 机制(去噪)— ", "t":"「GPT-5.3-Codex 参与创造自己」= 工程师把早期 checkpoint 当工具用于调试训练/诊断 eval,而非自主自训练。", "color":AMBER, "tag":"去噪"},
 {"t":"未公开:奖励权重配比、审查 eval 的样本量/分母;MoE/稀疏激活仅为第三方推测,OpenAI 未披露。", "lvl":1, "color":GRAY, "bullet":"○  "},
], y=Inches(1.5), size=13.5, gap=10)

# ============================================================ SLIDE 7 — ADAPTIVE COMPUTE
s = content_slide("TECH DEEP-DIVE 2/6 · INFERENCE", "技术深钻②:自适应推理 —「思考时间」如何按难度分配", 7)
bullets(s, [
 {"lead":"两个头条数字 — ", "t":"对最简单 10% 的对话,GPT-5-Codex 比 GPT-5 少用约 93.7% 的「思考」token;对最难 10%,思考/行动时长约为 2 倍。", "tag":"强"},
 {"lead":"机制(关键澄清)— ", "t":"这是模型自身按任务复杂度「自调步」分配推理,而非 GPT-5 聊天产品里那个「快/慢模型实时路由器」。单模型自定推理 ≠ 系统级路由。", "color":TEAL, "tag":"中"},
 {"lead":"推理档位外显 — ", "t":"reasoning effort 以离散档暴露:low / medium / high;自 5.1-Codex-Max 起新增 xhigh(为最难的异步 agentic 任务/评测「想更久」)。"},
 {"lead":"工程含义 — ", "t":"简单回合省 token=降成本/提速;难题多花算力=提质量。这也解释了「Codex 在简单任务更省、在复杂任务更彻底」的口碑分裂。"},
 {"t":"不确定:无法独立证实该模型内部「绝无路由」;各版本确切档位集合(none/minimal/…/xhigh)未完全锁定。", "lvl":1, "color":GRAY, "bullet":"○  "},
], y=Inches(1.55), size=14, gap=12)

# ============================================================ SLIDE 8 — COMPACTION
s = content_slide("TECH DEEP-DIVE 3/6 · LONG-HORIZON", "技术深钻③:Compaction — 跨上下文窗口的长时程自主", 8)
bullets(s, [
 {"lead":"定义 — ", "t":"GPT-5.1-Codex-Max 是「首个原生训练以跨多个上下文窗口工作」的模型,通过 compaction 在单任务内连贯处理数百万 token;内部观察到 24h+ 自主迭代。", "tag":"强"},
 {"lead":"机制 — ", "t":"接近窗口上限时,把关键状态(架构决策、测试失败、当前目标、文件位置)蒸馏成密集摘要带入新窗口,丢弃原始 token、保留状态,循环至完成。", "tag":"中"},
 {"lead":"原生 vs 脚手架(核心争议)— ", "t":"CLI 早就有「摘要再续」的 app 层 auto-compact;新意在于模型被训练于这些跨窗口轨迹,使摘要为「自己将来要用」而优化,并经 Responses API 一等公民暴露。", "color":TEAL, "tag":"去噪"},
 {"lead":"Harness 实测(badlogic / danielvaughan)— ", "t":"Codex CLI 触发阈≈ effective_window − 13k,effective = 窗口 − min(max_out, 20k);Claude Code 约 95% 容量才触发(常被诟病「太晚」),且额外保留最近 5 个文件。"},
 {"lead":"标称 ≠ 实际 — ", "t":"GitHub issue 报告会话实际窗口与官方不符(如 GPT-5.5 报 258,400 vs 标称 400K),可能绕过 auto-compaction。把官方窗口当「名义值」。", "color":AMBER, "tag":"强"},
], y=Inches(1.5), size=13, gap=9)

# ============================================================ SLIDE 9 — AGENT HARNESS
s = content_slide("TECH DEEP-DIVE 4/6 · HARNESS", "技术深钻④:Agent Harness 与工具循环", 9)
bullets(s, [
 {"lead":"Agent 循环 — ", "t":"组装历史 → 带工具调用模型 → 若有工具调用则在沙箱执行并回填 → 再查询 → 直到无工具调用的最终回复。一个「turn」可含多次模型/工具迭代。"},
 {"lead":"五个核心工具 — ", "t":"shell(在沙箱跑任意命令,读/写/搜文件即靠它)· apply_patch(结构化多文件 diff)· update_plan(维护 TODO/计划)· view_image · web_search(可并行)。工具三源:核心工具 + MCP + 迁移插件。"},
 {"lead":"为何用 Rust 重写(2025-06,~96% Rust)— ", "t":"①零依赖安装(去掉 Node 运行时)②原生安全绑定(沙箱已是 Rust)③长会话无 GC 停顿/堆膨胀、毫秒级启动(CI 并行起多 agent 关键)④语言无关线协议便于多语言扩展。"},
 {"lead":"通信 — ", "t":"经 OpenAI Responses API(parallel_tool_calls、tool_choice、reasoning 控制、流式事件);ToolRouter 在起进程前强制审批策略并选定沙箱。"},
 {"lead":"非交互 — ", "t":"codex exec PROMPT 跑到模型自判完成即退;hooks 可注入脚本做日志/提示扫描/工具调用校验。"},
], y=Inches(1.5), size=13, gap=9)

# ============================================================ SLIDE 10 — SANDBOX & SECURITY
s = content_slide("TECH DEEP-DIVE 5/6 · SANDBOX", "技术深钻⑤:沙箱与安全架构 — OS 级强制隔离", 10)
bullets(s, [
 {"lead":"双层模型 — ", "t":"沙箱能力(OS 物理允许什么:read-only / workspace-write / danger-full-access)与 审批策略(何时暂停问人:on-request / on-failure / never)解耦。"},
 {"lead":"macOS — ", "t":"Apple Seatbelt(sandbox-exec),默认 deny;允许 file-read*、写限于可写根;省略网络权限即默认拒网;.git/.codex/.agents 元数据受保护。"},
 {"lead":"Linux — ", "t":"现以 bubblewrap(bwrap)为主做命名空间隔离 + PR_SET_NO_NEW_PRIVS + seccomp 网络过滤(挡 connect/bind/sendto 等、放行 AF_UNIX);Landlock+seccomp 为显式 legacy 回退(需内核 ≥5.13)。", "tag":"强(README 直读)"},
 {"lead":"Windows — ", "t":"OpenAI 自研:弃用 AppContainer,采用「写限制令牌 + 合成 SID」(写须同时满足用户身份与受限 SID 列表);用 capability SID 授文件访问;以覆盖代理环境变量 + 桩可执行文件禁出网。"},
 {"lead":"云端两阶段 — ", "t":"setup 阶段有网(跑安装脚本、可见 secret)→ agent 阶段默认断网(secret 已移除),全程经 HTTP(S) 代理走 allowlist;改脚本/env/secret 即失效缓存。理由:防注入驱动的数据外泄、限爆炸半径。"},
], y=Inches(1.48), size=12.5, gap=7)

# ============================================================ SLIDE 11 — EXTENSIBILITY
s = content_slide("TECH DEEP-DIVE 6/6 · EXTENSIBILITY", "技术深钻⑥:扩展性 — MCP、Symphony、子 agent、SDK", 11)
bullets(s, [
 {"lead":"MCP 双向 — ", "t":"作 client:config.toml 内 [mcp_servers.<name>] 声明,codex mcp add/list/remove 管理;作 server:codex mcp-server 跑 stdio JSON-RPC,客户端调 codex 工具即经 ThreadManager 起会话并流式回事件。per-tool 审批模式(auto/prompt/approve)。"},
 {"lead":"Symphony(真实)— ", "t":"OpenAI 开源的 Codex 编排规范,本质是一份 SPEC.md(问题/方案定义,非代码),交给编码 agent 物化成任意语言;把 issue tracker(如 Linear)连到 Codex,每任务从指派到 PR 自成闭环。", "tag":"中"},
 {"lead":"内建子 agent — ", "t":".agents/subagents/<name>.md 配置;编排含 spawn/路由/等待/收束;agents.max_depth 默认 1(直接子可派生,默认不再深嵌)。"},
 {"lead":"AGENTS.md 分层 — ", "t":"~/.codex 级 → 项目根至 cwd 逐级拼接(深者后置=胜出,近因偏置);AGENTS.override.md 在其层级「替换」而非追加;总量上限 project_doc_max_bytes(默认 32 KiB)。"},
 {"lead":"SDK — ", "t":"Codex TypeScript SDK 实为「包裹 codex CLI」(spawn 进程、JSONL over stdin/stdout):startThread/run/runStreamed、JSON-schema 结构化输出、~/.codex/sessions 持久化、resumeThread;原生 GitHub Actions。"},
], y=Inches(1.48), size=12, gap=7)

# ============================================================ SLIDE 12 — BENCHMARK METHODOLOGY
s = content_slide("BENCHMARKS", "Benchmark 方法论:跑分背后的四个陷阱", 12)
bullets(s, [
 {"lead":"陷阱① 分母 477 vs 500 — ", "t":"SWE-bench Verified 是 500 例人工校验子集;OpenAI 早期只在能跑通的 477 例上报告,后改 500 例。跨模型对比必须核对分母。", "tag":"强"},
 {"lead":"陷阱② pass@1 是「4 次平均」 — ", "t":"OpenAI 的 pass@1 对每例平均 4 次尝试(平滑,非真单发);且不把单测喂给模型。", "tag":"强"},
 {"lead":"陷阱③ harness/脚手架敏感 — ", "t":"同模型换脚手架分数可差 5–15 分(例:Auggie 在同一 Opus 上比 Claude Code 多解 17/731);「agent 榜」≠「模型榜」。"},
 {"lead":"陷阱④ 评测迁移 + 标称窗口 — ", "t":"2026 年新模型主跑 SWE-bench Pro / Terminal-Bench 2.0,与早期 Verified 不可直接续比;且会话实际上下文常小于官方标称。"},
 {"lead":"厂商 vs 独立 — ", "t":"GPT-5.5 的 88.7% 为 OpenAI 自报,独立 harness 仅 ~82.6%;GPT-5.3-Codex 的 Terminal-Bench 2.0 有 75.1% vs 77.3% 的来源冲突。", "color":AMBER, "tag":"中"},
 {"t":"结论:任何百分比都须绑「模型版本 + 日期 + 评测集 + harness + 推理档」。脱离上下文的「谁第一」基本无意义。", "lvl":1, "color":TEAL, "bullet":"➔  ", "bold":True},
], y=Inches(1.5), size=13, gap=8)

# ============================================================ SLIDE 13 — ARCH COMPARISON TABLE
table_slide(
 "ARCHITECTURE COMPARISON", "架构级对比:Codex vs Claude Code", 13,
 ["维度", "OpenAI Codex", "Anthropic Claude Code"],
 [
  ["执行模型", ("云沙箱默认 + 本地 Rust CLI;并行容器 + git worktree", INK, False), ("本地优先终端 agent;可见本机服务/DB/env", INK, False)],
  ["数据驻留", ("代码上传 OpenAI 基础设施(fire-and-forget 优)", INK, False), ("默认本地执行(合规/隔离场景优)", INK, False)],
  ["上下文记忆", ("AGENTS.md;原生 compaction(模型级)", INK, False), ("CLAUDE.md;~95% 才 compact + 保留 5 个近期文件", INK, False)],
  ["权限模型", ("OS 内核沙箱为主(Seatbelt/bwrap/令牌)", INK, False), ("应用层权限引擎 + hooks 为主,OS 沙箱为辅", INK, False)],
  ["子 agent", ("Workflow 并行编排,汇总单一结果", INK, False), ("各子 agent 独立上下文 + 可路由廉价模型 + MEMORY.md", INK, False)],
  ["模型锁定", ("仅 OpenAI(GPT-5.x-Codex)— 模型×harness 协同设计", AMBER, True), ("仅 Anthropic(Opus/Sonnet/Haiku)", AMBER, True)],
  ["技术擅场", ("长时程、并行、可委派的自主执行 + token 节俭", GREEN, True), ("交互式、多文件深推理 + 细粒度治理", GREEN, True)],
 ],
 col_w=[1.7, 5.3, 5.3],
 note="架构事实多为官方文档直读(高可信);compaction 差异据 badlogic 代码级独立研究 · Cursor=模型中立,可换各家但用不到专用 GPT-5-Codex",
 fs=10.5, hfs=11.5, row_h=Inches(0.73), y=Inches(1.42))

# ============================================================ SLIDE 14 — COMPETITION
s = content_slide("COMPETITIVE LANDSCAPE", "竞争格局:三王分治 + 多工具叠用", 14)
bullets(s, [
 {"lead":"三个「第一」按口径划分 — ", "t":"Copilot 赢用户基数(~470 万付费);Cursor 赢营收(~$2B ARR,传 $50B 估值);Claude Code 赢满意度/口碑(JetBrains 2026 调查最受欢迎,CSAT 91%)。", "tag":"强"},
 {"lead":"JetBrains 2026 调查(>1 万开发者)— ", "t":"职场采用率 Copilot ~29% / Cursor ~18% / Claude Code ~18%;Claude Code 一年内从 3% 增至 18%(6x)。", "tag":"强"},
 {"lead":"叠用成常态 — ", "t":"~70% 工程师同时用 2–4 个 AI 编码工具;创业公司偏 Claude Code(~75%),万人以上大厂偏 Copilot(~56%,采购/分发优势)。"},
 {"lead":"格局剧变(2026)— ", "t":"微软与 OpenAI 结束独家绑定(~4 月),Copilot 转多模型(接 OpenAI+Anthropic+自研 MAI/Polaris);Google 下线 Gemini CLI 转 Antigravity(5 月);Windsurf 更名 Devin Desktop(6 月)。", "tag":"强"},
 {"t":"已剔除:网传「Codex 2.0 Orchestrator」无官方佐证,与 OpenAI 实际命名矛盾。", "lvl":1, "color":RED, "bullet":"✗  "},
], y=Inches(1.52), size=13.5, gap=11)

# ============================================================ SLIDE 15 — BUSINESS
s = content_slide("BUSINESS & STRATEGY", "商业与市场策略", 15)
bullets(s, [
 {"lead":"定价 — ", "t":"捆绑 Plus/Pro/Business/Enterprise/Edu,不单售;2026-04-02 由按消息改按 token 信用计费(社区不满导火索之一)。登录态走套餐额度,切 API-key 则按 API 单价。", "tag":"强"},
 {"lead":"采用 — ", "t":"周活 200 万(3 月)→ 400 万(4.21)→ 500 万+(6 月),~20% 非开发者;OpenAI 内部用 Codex 的工程师从 <50% 升至 >90%,PR 量 +~70%。", "tag":"中(自报)"},
 {"lead":"营收战略 — ", "t":"编码是 LLM 最高价值变现场景(有分析称占企业生成式 AI 用量约半);The Information 称 OpenAI 2026 Q1 营收近 $60 亿、季度领先 Anthropic ~$10 亿。Claude Code 据报 >$2.5B 年化,是对标标尺。", "tag":"中"},
 {"lead":"Windsurf 始末 — ", "t":"OpenAI 拟 ~$30 亿收购 → 2025-07-11 因(据报)与微软 IP 摩擦告吹 → Google 以 ~$24 亿「逆向收编」CEO 等核心 → 余下由 Cognition(Devin)接手。", "tag":"强"},
 {"lead":"争议 — ", "t":"GPT-5.3-Codex 罕见附严格管控(称或首次「足以放大现实网络危害」);2026-06 冒充 Codex 的恶意 npm 包窃 token;2025 末 GitHub token 泄露漏洞(已修)。", "tag":"强"},
], y=Inches(1.5), size=13, gap=9)

# ============================================================ SLIDE 16 — DEVELOPER VOICE
s = content_slide("DEVELOPER VOICE", "开发者实践与口碑", 16)
bullets(s, [
 {"lead":"被称赞 — ", "t":"彻底/全面(不止改目标函数,连带修调用方、统一日志、标记弃用);「委派—回看」异步流(早上排 3–5 个任务,回来已有审好的 PR);干净 diff;自动 PR 审查。", "color":GREEN},
 {"lead":"被吐槽 — ", "t":"沙箱默认无网阻断装依赖;2026-04 转 token 计费后限流抱怨集中爆发(「5 小时额度一小时用光」);容器内跑测试偏慢;默认过度工程;不能手动选模型。", "color":AMBER},
 {"lead":"口碑拐点 — ", "t":"GPT-5-Codex(2025-09)是关键转折:Builder.io 测得用户评分高出替代品约 40%。但成因有争议——是 Codex 真变强,还是 Claude Code 同期「被阉割/harness 变怪」,两种叙事并存。", "tag":"中"},
 {"lead":"当前共识:双持(dual-wield)— ", "t":"Codex 做后端/重测试/重构与「可辩护的 diff」;Claude Code 做 UI/探索式规划/自主部署。「问题不是谁更聪明,而是你在哪种工作流里。」", "color":TEAL},
 {"lead":"最佳实践 — ", "t":"用好 AGENTS.md(含测试命令)· prompt 显式约束范围以抑过度工程 · 合理设审批档(--full-auto 常为平衡点)· 用子 agent/REQUIREMENTS.md 做结构化交接。"},
], y=Inches(1.5), size=13, gap=9)

# ============================================================ SLIDE 17 — KEY JUDGMENTS
s = content_slide("KEY JUDGMENTS", "核心判断(分析层)", 17)
bullets(s, [
 {"lead":"护城河在「模型×harness 协同设计」 — ", "t":"GPT-5.x-Codex 与其 agent 循环共同 RL 调优,带来更紧的工具使用与 token 效率;但代价是 OpenAI 模型锁定。Cursor/Copilot 以可移植性换取这种垂直整合。"},
 {"lead":"结构性权衡:云 vs 本地 — ", "t":"Codex 云沙箱适合 fire-and-forget、并行、可复现(pinned 镜像);数据敏感/隔离场景,Claude Code 的本地优先是结构性优势。"},
 {"lead":"长时程自主是新战线 — ", "t":"compaction + 24h+ 把竞争从「单轮代码质量」推向「跨窗口状态管理 + 多 agent 编排」。谁能稳定地长时间不跑偏,谁赢异步委派场景。"},
 {"lead":"成熟度的反面是攻击面 — ", "t":"Slack/SDK/插件/Sites 提升渗透,同时带来供应链(恶意 npm)、OAuth token 等新安全面——企业采用须纳入评估。"},
 {"lead":"认知风险:别被「谁第一」带偏 — ", "t":"品牌跨三代、版本月月变、跑分口径混乱;务实做法是按工作流组合工具,而非追逐排行榜。", "color":TEAL},
], y=Inches(1.55), size=13.5, gap=12)

# ============================================================ SLIDE 18 — SELECTION GUIDE
table_slide(
 "RECOMMENDATIONS", "选型建议:按场景对号入座", 18,
 ["你的场景", "首选", "理由"],
 [
  ["异步委派 / 并行批量 / CI 自动化", ("Codex(云)", GREEN, True), "隔离容器 + worktree + best-of-N;免审批自主更可控"],
  ["数据敏感 / 隔离网 / 强合规", ("Claude Code", TEAL, True), "本地优先执行 + 应用层细粒度权限 + hooks 治理"],
  ["交互式多文件深推理 / 探索规划", ("Claude Code", TEAL, True), "独立上下文子 agent + JIT 检索 + 主动追问"],
  ["长时程自主重构 / 长任务", ("Codex(5.1-Max+)", GREEN, True), "原生 compaction 跨窗口,24h+ 连贯"],
  ["要多模型自由 / 价格套利", ("Cursor / Copilot", AMBER, True), "模型中立可切换;但用不到专用 GPT-5-Codex"],
  ["务实团队", ("双持", INK, True), "Codex 跑后端/测试/重构,Claude Code 跑 UI/规划/部署"],
 ],
 col_w=[4.4, 2.5, 5.4],
 note="基于本次研究的工程事实与社区共识综合;非厂商背书 · 具体取舍仍需结合团队栈与合规要求验证",
 fs=11.5, hfs=12, row_h=Inches(0.66), y=Inches(1.5))

# ============================================================ SLIDE 19 — RISKS & UNCERTAINTIES
s = content_slide("RISKS & UNCERTAINTIES", "风险与不确定性(诚实披露)", 19)
bullets(s, [
 {"lead":"已证伪(剔除)— ", "t":"「Codex 2.0 / Orchestrator / SafeDeploy / OpenRepo」单一来源虚构,与官方命名矛盾。", "color":RED, "bullet":"✗  "},
 {"lead":"已降级(传闻)— ", "t":"「4x token 效率」为 n=1 博客且含产出质量混淆;「88.7% SWE-bench」为厂商自报(独立 ~82.6%)。", "color":AMBER, "bullet":"~  "},
 {"lead":"机制存疑 — ", "t":"GPT-5-Codex 是否「绝无内部路由」无法证实;compaction 是否「不止于训练版摘要再续」未完全公开;MoE/稀疏激活仅第三方推测。", "color":GRAY, "bullet":"○  "},
 {"lead":"数据口径 — ", "t":"周活/营收多为 OpenAI 自报或单一媒体(The Information),无第三方审计;标称上下文 ≠ 会话实际上下文。", "color":GRAY, "bullet":"○  "},
 {"lead":"环境限制 — ", "t":"本次 WebFetch 全程 403,结论基于搜索摘要(GitHub README、badlogic 研究可直读);2026 项超知识截止。", "color":GRAY, "bullet":"○  "},
 {"t":"建议:正式引用前对照官方 system card PDF 与 developers.openai.com/codex 原文核对确切数字与日期。", "lvl":1, "color":TEAL, "bullet":"➔  ", "bold":True},
], y=Inches(1.5), size=13, gap=10)

# ============================================================ SLIDE 20 — SOURCES
s = content_slide("SOURCES", "主要来源(节选)", 20, note="完整 URL 见随附报告;以下为各维度最权威/最具价值来源")
bullets(s, [
 {"lead":"OpenAI 官方 — ", "t":"introducing-codex · introducing-upgrades-to-codex · gpt-5-1-codex-max(+system card PDF)· introducing-gpt-5-3-codex · introducing-gpt-5-5 · codex-now-generally-available · building-codex-windows-sandbox · unrolling-the-codex-agent-loop · developers.openai.com/codex/*", "size":12},
 {"lead":"代码 / 一手 — ", "t":"github.com/openai/codex(codex-rs README、linux-sandbox README 直读)· cdn.openai.com 各 system card PDF · cookbook.openai.com 提示指南", "size":12},
 {"lead":"独立技术分析 — ", "t":"simonwillison.net(5.1-Codex-Max)· infoq.com(Rust 重写 / Windows 沙箱)· badlogic gist(compaction 代码级研究)· codex.danielvaughan.com · epoch.ai / swebench.com(基准方法论)", "size":12},
 {"lead":"市场 / 商业 — ", "t":"The Information(Q1 营收)· JetBrains Research 2026(开发者调查)· TechCrunch / Bloomberg(Windsurf)· Google Developers Blog(Antigravity)· devin.ai(Windsurf→Devin)", "size":12},
 {"lead":"开发者口碑 — ", "t":"Hacker News 多帖 · dev.to(2 月实测)· builder.io · OpenAI 社区(限流帖)· anthonymaio.substack(「Codex 变好因 Claude 变怪」)", "size":12},
 {"t":"研究方式:5 路并行检索 + 3 路对抗式核验(2/3 证伪即剔除);可信度全程分层标注。", "lvl":1, "color":GREEN, "bullet":"✔  ", "bold":True, "size":12},
], y=Inches(1.5), size=12, gap=11)

out = "/home/user/openai-chatgpt-codex/OpenAI_Codex_深度技术洞察.pptx"
prs.save(out)
print("saved:", out, "slides:", len(prs.slides._sldIdLst))
