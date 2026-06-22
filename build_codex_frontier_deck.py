# -*- coding: utf-8 -*-
"""Build the merged OpenAI Codex x Frontier deep-insight deck (PPTX).

Extends the original 20-slide Codex deck (build_codex_deck.py): reuses the same
visual system and all Codex content, and weaves in a Frontier section, a
Codex x Frontier synthesis, an OpenAI agent-stack map, an enterprise agent
platform competition table, and an honest Frontier risk/verification slide.
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
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
BLUE   = RGBColor(0x2B, 0x59, 0x9A)
LGREEN = RGBColor(0xE7, 0xF4, 0xEF)
LTEAL  = RGBColor(0xE2, 0xF0, 0xEF)

FONT = "Microsoft YaHei"

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]

# ---------- auto page numbering ----------
_pg = [1]  # cover is page 1 (not stamped)
def P():
    _pg[0] += 1
    return _pg[0]

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

def roundrect(slide, x, y, w, h, color, line=None, line_w=1):
    sp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    sp.fill.solid(); sp.fill.fore_color.rgb = color
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line; sp.line.width = Pt(line_w)
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
    tbl.first_row = True; tbl.horz_banding = True
    for j, cw in enumerate(col_w):
        tbl.columns[j].width = Inches(cw)
    for j, htext in enumerate(headers):
        c = tbl.cell(0, j)
        c.fill.solid(); c.fill.fore_color.rgb = NAVY
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        c.margin_left = Inches(0.07); c.margin_right = Inches(0.05)
        c.margin_top = Inches(0.02); c.margin_bottom = Inches(0.02)
        p = c.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
        r = p.add_run(); r.text = htext; r.font.size = Pt(hfs); r.font.bold = True
        r.font.color.rgb = WHITE; set_font(r)
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

def card(slide, x, y, w, h, hcolor, title, body, body_size=11.5):
    """A titled card: rounded body + colored header strip + body text."""
    roundrect(slide, x, y, w, h, CARD, line=LGRAY, line_w=1)
    rect(slide, x+Inches(0.05), y+Inches(0.05), w-Inches(0.10), Inches(0.5), hcolor)
    tb, tf = textbox(slide, x+Inches(0.2), y+Inches(0.05), w-Inches(0.35), Inches(0.5))
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    style(tf.paragraphs[0], title, 13, WHITE, bold=True, space_after=0)
    tb2, tf2 = textbox(slide, x+Inches(0.22), y+Inches(0.66), w-Inches(0.42), h-Inches(0.78))
    style(tf2.paragraphs[0], body, body_size, INK, space_after=0)

# ============================================================ SLIDE 1 — COVER
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, SW, SH, NAVY)
rect(s, 0, Inches(4.05), SW, Pt(3.5), GREEN)
rect(s, Inches(0.0), 0, Inches(0.22), SH, GREEN)
tb, tf = textbox(s, Inches(0.9), Inches(1.35), Inches(11.5), Inches(2.8))
p = tf.paragraphs[0]; style(p, "OpenAI  CODEX  ×  FRONTIER", 17, GREEN, bold=True, space_after=6)
p = tf.add_paragraph(); style(p, "深度洞察分析", 46, WHITE, bold=True, space_after=4)
p = tf.add_paragraph(); style(p, "Coding Agent(纵深) × Enterprise Agent Platform(横向) — 同一 Agent 战略的一体两面", 15, LGRAY)
tb, tf = textbox(s, Inches(0.9), Inches(4.3), Inches(11.5), Inches(2.3))
for t in [
    "产品形态 · 模型与训练 · Harness 与沙箱 · 企业 Agent 平台架构 · 治理与评估 · 竞争格局 · 商业战略",
]:
    p = tf.add_paragraph(); style(p, t, 13, LGRAY, space_after=10)
p = tf.add_paragraph(); style(p, "2026 年 6 月   |   多源并行检索 + 对抗式核验(2/3 证伪即剔除)· 可信度全程分层标注", 12, GREEN, bold=True, space_after=4)
p = tf.add_paragraph(); style(p, "方法学提示:本次 openai.com / 主流媒体 WebFetch 全程 403;Frontier(2026-02 新品)结论基于 10+ 来源的搜索摘要交叉验证;2026 年数据超模型知识截止(2026-01),引用前请对照官方 system card 与 openai.com/business/frontier。", 9.5, GRAY)

# ============================================================ SLIDE 2 — EXEC SUMMARY
s = content_slide("EXECUTIVE SUMMARY", "执行摘要:七个核心洞察", P())
bullets(s, [
 {"lead":"一个 Agent 战略,两条战线 — ", "t":"Codex(把『软件工程』做到极致的自主编码 agent)与 Frontier(2026-02-05 发布的企业 AI agent 平台)同属 OpenAI『agent 攻势』,同日推进,互为纵深与横向。", "tag":"强"},
 {"lead":"Codex:品牌三代复用,切勿混淆 — ", "t":"2021 初代(驱动 Copilot,API 2023-03 关停)≠ 2025 重启的『云+CLI+IDE』自主软件工程 agent;codex-1→GPT-5-Codex→5.1-Max→5.3→5.5 凶猛迭代。", "tag":"强"},
 {"lead":"Codex 壁垒=模型×harness×沙箱三位一体 — ", "t":"RL 让模型与 agent 循环协同设计;Rust harness 五工具循环走 Responses API;OS 内核级沙箱使『免审批自主』可控。", "tag":"强"},
 {"lead":"Frontier=企业 Agent 的『控制平面』 — ", "t":"四大组件(业务上下文/执行/评估/治理),把 agent 当『可管理的 AI 同事』(入职→反馈→权限),补齐把 agent『投产』所缺的拼图;并兼容第三方 agent。", "tag":"强"},
 {"lead":"商业:企业收入引擎 — ", "t":"企业收入已占 OpenAI >40%、目标年底约 50%,Frontier 是『载体』;Codex 推动 2026 Q1 营收。战略上与 Salesforce / ServiceNow / Microsoft 正面相撞。", "tag":"中"},
 {"lead":"格局是『两层叠加』 — ", "t":"编码工具层(Copilot 基数 · Cursor 营收 · Claude Code 口碑;Codex 差异化=长时程沙箱自主)+ 企业 agent 平台层(Frontier vs Agentforce / Copilot Studio / watsonx)。", "tag":"强"},
 {"t":"关键风险:版本/跑分极易张冠李戴;且『OpenAI Frontier』≠『Microsoft Frontier Suite』——任何能力数字都须绑定『产品+版本+日期+来源』才有意义。", "lvl":1, "color":AMBER, "bullet":"⚠  "},
], y=Inches(1.45), size=12.5, gap=8)

# ============================================================ SLIDE 3 — METHOD & CONFIDENCE
s = content_slide("METHODOLOGY", "研究方法与可信度框架", P())
bullets(s, [
 {"lead":"流程 — ", "t":"将问题拆为多个角度并行检索 → 去重抓取 → 对每条关键论断做对抗式核验(2/3 证伪即剔除)→ 按可信度排序综合。"},
 {"lead":"可信度分层 — ", "t":"强(官方 + 多源交叉);中(单一媒体或 OpenAI 自报、无第三方审计);弱(已剔除或降级处理)。"},
 {"t":"核验结果:已证伪并剔除", "lvl":1, "color":RED, "bullet":"✗  ", "bold":True},
 {"t":"Codex 侧:「Codex 2.0 / Orchestrator / SafeDeploy / OpenRepo」· Frontier 侧:「Frontier Alliance / DeployCo / Codex Labs」 — 均仅见单一非权威页面,与 OpenAI 实际命名矛盾,判为不实。", "lvl":2, "color":GRAY},
 {"t":"核验结果:已降级为传闻 / 厂商自报", "lvl":1, "color":AMBER, "bullet":"~  ", "bold":True},
 {"t":"「省 4x token」源自单次 Figma-to-code(n=1);「GPT-5.5 SWE-bench 88.7%」为自报(独立 ~82.6%);Frontier「6 周→1 天 / +5% 产出」为 OpenAI 精选案例、无第三方审计。", "lvl":2, "color":GRAY},
 {"lead":"核心限制 — ", "t":"本次 openai.com / 主流媒体 WebFetch 全程 403,结论基于多源搜索摘要交叉验证(GitHub README、badlogic 研究可直读);Frontier 为 2026-02 新品、超模型知识截止(2026-01)。发布前请对照官方 system card 与 openai.com/business/frontier 原文。", "color":AMBER, "tag":"重要"},
], y=Inches(1.45), size=13, gap=9)

# ============================================================ SLIDE 4 — OPENAI AGENT STACK (visual)
s = content_slide("STRATEGIC MAP", "一张图看懂:Codex 与 Frontier 在 OpenAI Agent 栈中的位置", P(),
                  note="栈为本报告分析框架(综合官方表述 + 媒体框定)· Codex 与 Frontier 同日推进:2026-02-05")
bands = [
 ("①  模型层 · Models", "GPT-5.x 家族 —— 含 GPT-5.3-Codex,作为复杂 agentic 任务的推理引擎", NAVY, CARD, INK),
 ("②  构建层 · Build", "AgentKit / Agents SDK:Agent Builder · Connector Registry · ChatKit;内建 tracing 与 eval", GRAY, CARD, INK),
 ("③  专用执行层 · Codex", "把『软件工程』一个垂直做到极致的自主编码 coworker —— 模型×harness×沙箱三位一体(纵深)", GREEN, LGREEN, INK),
 ("④  企业编排与治理层 · Frontier", "面向全企业的 AI coworker 控制平面:业务上下文 + 执行 + 评估 + 治理(横向)", TEAL, LTEAL, INK),
 ("⑤  分发层 · ChatGPT 超级应用", "整合 ChatGPT + Codex + agentic browsing + 合作伙伴应用,触达消费者与工作区", GRAY, CARD, INK),
]
bx, bw = Inches(0.6), Inches(12.13)
btop, bh, bgap = 1.40, 0.84, 0.135
for i, (lab, desc, accent, fill, txtc) in enumerate(bands):
    y0 = Inches(btop + i*(bh+bgap))
    roundrect(s, bx, y0, bw, Inches(bh), fill, line=(accent if fill != CARD else LGRAY),
              line_w=(1.5 if fill != CARD else 1))
    rect(s, bx, y0, Inches(0.13), Inches(bh), accent)
    tb, tf = textbox(s, bx+Inches(0.32), y0, Inches(3.95), Inches(bh))
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    style(tf.paragraphs[0], lab, 14.5, accent if fill != CARD else INK, bold=True, space_after=0)
    tb2, tf2 = textbox(s, bx+Inches(4.35), y0, Inches(7.5), Inches(bh))
    tf2.vertical_anchor = MSO_ANCHOR.MIDDLE
    style(tf2.paragraphs[0], desc, 11.5, txtc, space_after=0)
tb, tf = textbox(s, Inches(0.6), Inches(6.5), Inches(12.13), Inches(0.5))
style(tf.paragraphs[0],
      "核心结论:Codex = 纵深(一个垂直做到极致) · Frontier = 横向(把 agent 推广到全企业并加治理);二者是同一战略的一体两面。",
      12, TEAL, bold=True, space_after=0)

# ============================================================ SLIDE 5 — PRODUCT SURFACES (Codex)
s = content_slide("CODEX · PRODUCT", "Codex 产品形态全景:一个 agent,六个入口", P())
bullets(s, [
 {"lead":"多端形态 — ", "t":"云端 agent(chatgpt.com/codex)· Codex CLI(开源 Apache-2.0,Rust)· VS Code/IDE 扩展 · 桌面 App · GitHub 集成 · Slack · iOS 移动端(2026-05)。"},
 {"lead":"三种交互范式 — ", "t":"CLI 重隐私/控制与 CI/CD;IDE 在编辑器内最小干扰;云端做异步、并行、可委派的多文件任务。2026 主流是「混合」:App 作指挥中心规划/审查 + CLI/IDE 本地落地。"},
 {"lead":"关键特性 — ", "t":"审批模式 × 沙箱模式 双层安全;AGENTS.md 项目指令(分层覆盖);MCP(既是 client 又是 server);PR 代码审查(仅报 P0/P1);子 agent 并行编排。"},
 {"lead":"计费 — ", "t":"不单独售卖,捆绑进 Plus($20)/Pro($200)/Business/Enterprise/Edu;2026-04-02 由「按消息」改为「按 token」信用计费。"},
 {"lead":"规模(截至 2026-06)— ", "t":"周活 500 万+,其中 ~20% 为非开发者(分析/市场/运营),非开发者增速约为工程师的 3 倍。", "tag":"中"},
], y=Inches(1.55), size=14, gap=13)

# ============================================================ SLIDE 6 — MODEL LINEAGE / TIMELINE (Codex)
table_slide(
 "CODEX · MODEL LINEAGE & TIMELINE", "Codex 模型谱系与演进时间线", P(),
 ["日期", "里程碑", "基座 / 谱系", "关键技术点", "SWE-bench Verified"],
 [
  ["2025-04-16", "Codex CLI 开源", "—", "终端 agent;发布即 TS/Node", "—"],
  ["2025-05-16", "云端 Codex 预览", ("codex-1 ← o3", TEAL), "RL 训练于真实编码任务;codex-mini ← o4-mini", "72.1%"],
  ["2025-09-15", "GPT-5-Codex", ("← GPT-5", TEAL), "自适应「思考时间」;代码审查训练", "74.5%"],
  ["2025-11-19", "GPT-5.1-Codex-Max", ("← 更新基座", TEAL), "原生 compaction 跨上下文窗口;内部 24h+", ("76.5% / 77.9%(xhigh)", INK, True)],
  ["2026-01-14", "GPT-5.2-Codex", ("← GPT-5.2", TEAL), "Terminal-Bench 2.0 新高", "—(转向 SWE-Pro)"],
  ["2026-02-05", "GPT-5.3-Codex", ("统一编码+推理", TEAL), "约快 25%;与 Frontier 同日发布;附严格网络安全管控", "~80% / Pro 56.8%"],
  ["2026-04-23", "GPT-5.5", ("首个重训基座", TEAL), "Codex 400K / API 1M 上下文", ("88.7%(厂商自报)", AMBER, True)],
 ],
 col_w=[1.6, 2.25, 2.2, 4.0, 2.55],
 note="谱系:Codex 模型均为基座推理模型的后训练变体,非独立架构 · 2026-02-05 为 Codex×Frontier 协调发布日 · 2026 项依赖检索",
 fs=10.5, hfs=10.5, row_h=Inches(0.72), y=Inches(1.42))

# ============================================================ SLIDE 7 — RL TRAINING
s = content_slide("CODEX · TECH 1/6 · TRAINING", "技术深钻①:RL 训练方法论 —「人类风格的 PR」从何而来", P())
bullets(s, [
 {"lead":"本质 — ", "t":"codex-1 是 o3 的后训练变体,用强化学习在「多环境的真实编码任务」上训练:写功能/加测试/调试/大规模重构/做代码审查。"},
 {"lead":"奖励信号(可验证 + 偏好)— ", "t":"①测试通过(可验证奖励,模型迭代跑测直到通过)②人类编码风格 ③PR/提交规范 ④指令遵从 ⑤代码审查偏好(由资深工程师对真实 OSS commit 的评审打分)。"},
 {"lead":"「人类风格 PR」≠ 新架构 — ", "t":"它是把输出对齐到上述奖励的结果:相比 o3,codex-1「产出更干净、可直接 review 的补丁」。", "tag":"强"},
 {"lead":"代码审查的可量化收益 — ", "t":"GPT-5-Codex 的错误评审意见比 GPT-5 少约 70%(资深工程师在真实 commit 上评估);OpenAI 内部已让 Codex 在人类 reviewer 之前自动审查大多数变更。", "tag":"强"},
 {"lead":"营销 vs 机制(去噪)— ", "t":"「GPT-5.3-Codex 参与创造自己」= 工程师把早期 checkpoint 当工具用于调试训练/诊断 eval,而非自主自训练。", "color":AMBER, "tag":"去噪"},
 {"t":"未公开:奖励权重配比、审查 eval 的样本量/分母;MoE/稀疏激活仅为第三方推测,OpenAI 未披露。", "lvl":1, "color":GRAY, "bullet":"○  "},
], y=Inches(1.5), size=13.5, gap=10)

# ============================================================ SLIDE 8 — ADAPTIVE COMPUTE
s = content_slide("CODEX · TECH 2/6 · INFERENCE", "技术深钻②:自适应推理 —「思考时间」如何按难度分配", P())
bullets(s, [
 {"lead":"两个头条数字 — ", "t":"对最简单 10% 的对话,GPT-5-Codex 比 GPT-5 少用约 93.7% 的「思考」token;对最难 10%,思考/行动时长约为 2 倍。", "tag":"强"},
 {"lead":"机制(关键澄清)— ", "t":"这是模型自身按任务复杂度「自调步」分配推理,而非 GPT-5 聊天产品里那个「快/慢模型实时路由器」。单模型自定推理 ≠ 系统级路由。", "color":TEAL, "tag":"中"},
 {"lead":"推理档位外显 — ", "t":"reasoning effort 以离散档暴露:low / medium / high;自 5.1-Codex-Max 起新增 xhigh(为最难的异步 agentic 任务/评测「想更久」)。"},
 {"lead":"工程含义 — ", "t":"简单回合省 token=降成本/提速;难题多花算力=提质量。这也解释了「Codex 在简单任务更省、在复杂任务更彻底」的口碑分裂。"},
 {"t":"不确定:无法独立证实该模型内部「绝无路由」;各版本确切档位集合(none/minimal/…/xhigh)未完全锁定。", "lvl":1, "color":GRAY, "bullet":"○  "},
], y=Inches(1.55), size=14, gap=12)

# ============================================================ SLIDE 9 — COMPACTION
s = content_slide("CODEX · TECH 3/6 · LONG-HORIZON", "技术深钻③:Compaction — 跨上下文窗口的长时程自主", P())
bullets(s, [
 {"lead":"定义 — ", "t":"GPT-5.1-Codex-Max 是「首个原生训练以跨多个上下文窗口工作」的模型,通过 compaction 在单任务内连贯处理数百万 token;内部观察到 24h+ 自主迭代。", "tag":"强"},
 {"lead":"机制 — ", "t":"接近窗口上限时,把关键状态(架构决策、测试失败、当前目标、文件位置)蒸馏成密集摘要带入新窗口,丢弃原始 token、保留状态,循环至完成。", "tag":"中"},
 {"lead":"原生 vs 脚手架(核心争议)— ", "t":"CLI 早就有「摘要再续」的 app 层 auto-compact;新意在于模型被训练于这些跨窗口轨迹,使摘要为「自己将来要用」而优化,并经 Responses API 一等公民暴露。", "color":TEAL, "tag":"去噪"},
 {"lead":"Harness 实测(badlogic / danielvaughan)— ", "t":"Codex CLI 触发阈≈ effective_window − 13k,effective = 窗口 − min(max_out, 20k);Claude Code 约 95% 容量才触发(常被诟病「太晚」),且额外保留最近 5 个文件。"},
 {"lead":"标称 ≠ 实际 — ", "t":"GitHub issue 报告会话实际窗口与官方不符(如 GPT-5.5 报 258,400 vs 标称 400K),可能绕过 auto-compaction。把官方窗口当「名义值」。", "color":AMBER, "tag":"强"},
], y=Inches(1.5), size=13, gap=9)

# ============================================================ SLIDE 10 — AGENT HARNESS
s = content_slide("CODEX · TECH 4/6 · HARNESS", "技术深钻④:Agent Harness 与工具循环", P())
bullets(s, [
 {"lead":"Agent 循环 — ", "t":"组装历史 → 带工具调用模型 → 若有工具调用则在沙箱执行并回填 → 再查询 → 直到无工具调用的最终回复。一个「turn」可含多次模型/工具迭代。"},
 {"lead":"五个核心工具 — ", "t":"shell(在沙箱跑任意命令,读/写/搜文件即靠它)· apply_patch(结构化多文件 diff)· update_plan(维护 TODO/计划)· view_image · web_search(可并行)。工具三源:核心工具 + MCP + 迁移插件。"},
 {"lead":"为何用 Rust 重写(2025-06,~96% Rust)— ", "t":"①零依赖安装(去掉 Node 运行时)②原生安全绑定(沙箱已是 Rust)③长会话无 GC 停顿/堆膨胀、毫秒级启动(CI 并行起多 agent 关键)④语言无关线协议便于多语言扩展。"},
 {"lead":"通信 — ", "t":"经 OpenAI Responses API(parallel_tool_calls、tool_choice、reasoning 控制、流式事件);ToolRouter 在起进程前强制审批策略并选定沙箱。"},
 {"lead":"非交互 — ", "t":"codex exec PROMPT 跑到模型自判完成即退;hooks 可注入脚本做日志/提示扫描/工具调用校验。"},
], y=Inches(1.5), size=13, gap=9)

# ============================================================ SLIDE 11 — SANDBOX & SECURITY
s = content_slide("CODEX · TECH 5/6 · SANDBOX", "技术深钻⑤:沙箱与安全架构 — OS 级强制隔离", P())
bullets(s, [
 {"lead":"双层模型 — ", "t":"沙箱能力(OS 物理允许什么:read-only / workspace-write / danger-full-access)与 审批策略(何时暂停问人:on-request / on-failure / never)解耦。"},
 {"lead":"macOS — ", "t":"Apple Seatbelt(sandbox-exec),默认 deny;允许 file-read*、写限于可写根;省略网络权限即默认拒网;.git/.codex/.agents 元数据受保护。"},
 {"lead":"Linux — ", "t":"现以 bubblewrap(bwrap)为主做命名空间隔离 + PR_SET_NO_NEW_PRIVS + seccomp 网络过滤(挡 connect/bind/sendto 等、放行 AF_UNIX);Landlock+seccomp 为显式 legacy 回退(需内核 ≥5.13)。", "tag":"强(README 直读)"},
 {"lead":"Windows — ", "t":"OpenAI 自研:弃用 AppContainer,采用「写限制令牌 + 合成 SID」(写须同时满足用户身份与受限 SID 列表);用 capability SID 授文件访问;以覆盖代理环境变量 + 桩可执行文件禁出网。"},
 {"lead":"云端两阶段 — ", "t":"setup 阶段有网(跑安装脚本、可见 secret)→ agent 阶段默认断网(secret 已移除),全程经 HTTP(S) 代理走 allowlist;改脚本/env/secret 即失效缓存。理由:防注入驱动的数据外泄、限爆炸半径。"},
], y=Inches(1.48), size=12.5, gap=7)

# ============================================================ SLIDE 12 — EXTENSIBILITY
s = content_slide("CODEX · TECH 6/6 · EXTENSIBILITY", "技术深钻⑥:扩展性 — MCP、Symphony、子 agent、SDK", P())
bullets(s, [
 {"lead":"MCP 双向 — ", "t":"作 client:config.toml 内 [mcp_servers.<name>] 声明,codex mcp add/list/remove 管理;作 server:codex mcp-server 跑 stdio JSON-RPC,客户端调 codex 工具即经 ThreadManager 起会话并流式回事件。per-tool 审批模式(auto/prompt/approve)。"},
 {"lead":"Symphony(真实)— ", "t":"OpenAI 开源的 Codex 编排规范,本质是一份 SPEC.md(问题/方案定义,非代码),交给编码 agent 物化成任意语言;把 issue tracker(如 Linear)连到 Codex,每任务从指派到 PR 自成闭环。", "tag":"中"},
 {"lead":"内建子 agent — ", "t":".agents/subagents/<name>.md 配置;编排含 spawn/路由/等待/收束;agents.max_depth 默认 1(直接子可派生,默认不再深嵌)。"},
 {"lead":"AGENTS.md 分层 — ", "t":"~/.codex 级 → 项目根至 cwd 逐级拼接(深者后置=胜出,近因偏置);AGENTS.override.md 在其层级「替换」而非追加;总量上限 project_doc_max_bytes(默认 32 KiB)。"},
 {"lead":"SDK — ", "t":"Codex TypeScript SDK 实为「包裹 codex CLI」(spawn 进程、JSONL over stdin/stdout):startThread/run/runStreamed、JSON-schema 结构化输出、~/.codex/sessions 持久化、resumeThread;原生 GitHub Actions。"},
], y=Inches(1.48), size=12, gap=7)

# ============================================================ SLIDE 13 — BENCHMARK METHODOLOGY
s = content_slide("CODEX · BENCHMARKS", "Benchmark 方法论:跑分背后的四个陷阱", P())
bullets(s, [
 {"lead":"陷阱① 分母 477 vs 500 — ", "t":"SWE-bench Verified 是 500 例人工校验子集;OpenAI 早期只在能跑通的 477 例上报告,后改 500 例。跨模型对比必须核对分母。", "tag":"强"},
 {"lead":"陷阱② pass@1 是「4 次平均」 — ", "t":"OpenAI 的 pass@1 对每例平均 4 次尝试(平滑,非真单发);且不把单测喂给模型。", "tag":"强"},
 {"lead":"陷阱③ harness/脚手架敏感 — ", "t":"同模型换脚手架分数可差 5–15 分(例:Auggie 在同一 Opus 上比 Claude Code 多解 17/731);「agent 榜」≠「模型榜」。"},
 {"lead":"陷阱④ 评测迁移 + 标称窗口 — ", "t":"2026 年新模型主跑 SWE-bench Pro / Terminal-Bench 2.0,与早期 Verified 不可直接续比;且会话实际上下文常小于官方标称。"},
 {"lead":"厂商 vs 独立 — ", "t":"GPT-5.5 的 88.7% 为 OpenAI 自报,独立 harness 仅 ~82.6%;GPT-5.3-Codex 的 Terminal-Bench 2.0 有 75.1% vs 77.3% 的来源冲突。", "color":AMBER, "tag":"中"},
 {"t":"结论:任何百分比都须绑「模型版本 + 日期 + 评测集 + harness + 推理档」。脱离上下文的「谁第一」基本无意义。", "lvl":1, "color":TEAL, "bullet":"➔  ", "bold":True},
], y=Inches(1.5), size=13, gap=8)

# ============================================================ SLIDE 14 — ARCH COMPARISON (Codex vs Claude Code)
table_slide(
 "CODEX · ARCHITECTURE COMPARISON", "架构级对比:Codex vs Claude Code(编码 agent 层)", P(),
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

# ============================================================ SLIDE 15 — FRONTIER OVERVIEW
s = content_slide("FRONTIER 1/6 · OVERVIEW", "OpenAI Frontier:从「聊天助手」到「可被管理的 AI 同事」", P())
bullets(s, [
 {"lead":"定义 — ", "t":"2026-02-05 发布的企业级平台,用于构建/部署/管理能做「真实工作」的 AI agent(OpenAI 称之为 AI coworkers);与 GPT-5.3-Codex 同日发布。", "tag":"强"},
 {"lead":"一句话定位 — ", "t":"介于「原始模型」与「业务应用」之间的平台层——被媒体框定为企业 AI agent 的「操作系统 / 控制平面 / 编排器」。", "tag":"强"},
 {"lead":"为何是现在 — ", "t":"企业把 agent 投产时,缺的往往不是模型,而是「共享业务上下文 + 治理 + 持续改进」这套支撑;Frontier 把它们收进一个平台。", "tag":"强"},
 {"lead":"与 ChatGPT Enterprise 区分 — ", "t":"Enterprise 是围绕聊天的「生产力产品」(约 $60/用户、150 席起);Frontier 面向「自主多步 agent 在生产环境规模化」,定价更高且定制。", "tag":"中"},
 {"lead":"战略野心 — ", "t":"CFO Sarah Friar:企业收入已占 OpenAI 总营收 >40%,目标年底约 50%、与消费持平,Frontier 是「载体」。", "tag":"中"},
], y=Inches(1.55), size=13.5, gap=12)

# ============================================================ SLIDE 16 — FRONTIER ARCHITECTURE (4 cards)
s = content_slide("FRONTIER 2/6 · ARCHITECTURE", "四大核心组件:把 agent「投产」所缺的拼图", P(),
                  note="据 VentureBeat / 多家媒体对官方表述的转述;openai.com 原文 WebFetch 受限,细节以官方页面为准")
cw, ch = Inches(5.85), Inches(2.42)
x1, x2 = Inches(0.62), Inches(6.86)
y1, y2 = Inches(1.45), Inches(4.05)
card(s, x1, y1, cw, ch, GREEN, "①  业务上下文 · Business Context(语义层)",
     "连接 CRM / ERP / 工单 / 数仓 / 内部应用;统一权限与检索逻辑,成为『所有 AI coworker 都能引用』的企业语义层,让 agent 理解信息如何流动、决策在哪发生。")
card(s, x2, y1, cw, ch, TEAL, "②  Agent 执行 · Agent Execution",
     "可推理、调工具、跑代码、处理文件,并从历史交互建立记忆;『开放的』agent 执行环境——兼容 OpenAI、企业自建及第三方 agent。")
card(s, x1, y2, cw, ch, BLUE, "③  评估与优化 · Evaluation & Optimization",
     "内建反馈回路,让 agent 表现对人类管理者透明,清晰显示『什么有效、什么需打磨』,随时间持续改进——如同对员工做绩效复盘。")
card(s, x2, y2, cw, ch, NAVY, "④  安全与治理 · Security & Governance",
     "身份管理、权限、合规与审计为『一等公民』;agent 自主访问数据须有清晰边界与可审计的动作,确保在限定范围内运行。")
tb, tf = textbox(s, Inches(0.62), Inches(6.56), Inches(12.1), Inches(0.4))
style(tf.paragraphs[0],
      "要点:Frontier 不取代 AgentKit / Agents SDK / API,而是把『上下文 + 执行 + 评估』收进同一平台,减少企业拼装多套系统的负担。",
      11, TEAL, bold=True, space_after=0)

# ============================================================ SLIDE 17 — FRONTIER OPERATING MODEL
s = content_slide("FRONTIER 3/6 · OPERATING MODEL", "运营心智模型:像招募与管理「员工」一样管理 Agent", P())
bullets(s, [
 {"lead":"设计起点 — ", "t":"OpenAI 借鉴企业「如何规模化人」:建立入职流程、传授制度知识与内部语言、允许在实践中学习、用反馈提升表现、授予恰当系统权限并设边界——AI coworker 需要同样这些。", "tag":"强"},
 {"lead":"入职(Onboarding)— ", "t":"agent 被「入职」:学习机构知识、内部术语与流程,并获授「正确的工具与系统」访问。"},
 {"lead":"在岗学习 + 反馈 — ", "t":"内建评估让「什么是高质量」对 agent 可学习;表现对人类管理者透明,可针对「对组织最重要的任务」做定向改进。", "tag":"强"},
 {"lead":"权限与边界 — ", "t":"明确「能做 / 不能做」,与企业身份系统打通,行为可审计。"},
 {"lead":"分析含义 — ", "t":"这把 agent 从「一次性脚本」升级为「可治理、可问责、可改进的数字员工」,也是 Frontier 区别于「裸 API / 框架」的叙事核心与销售抓手。", "color":TEAL, "tag":"分析"},
], y=Inches(1.55), size=13.5, gap=12)

# ============================================================ SLIDE 18 — FRONTIER OPENNESS / MULTI-VENDOR
s = content_slide("FRONTIER 4/6 · OPENNESS", "开放性与多供应商:既要「控制平面」,又不锁死模型?", P())
bullets(s, [
 {"lead":"多供应商兼容 — ", "t":"Frontier 兼容 OpenAI 自建 agent、企业自建 agent,以及第三方(Google / Microsoft / Anthropic)agent——主打「开放的 agent 执行环境」。", "tag":"强"},
 {"lead":"与现有工具关系 — ", "t":"不取代 AgentKit(Agent Builder · Connector Registry · ChatKit)/ Agents SDK / API;而是把「上下文 + 执行 + 评估」统一进一个平台。", "tag":"强"},
 {"lead":"连接器 — ", "t":"主流 CRM / 工单 / 数仓 / 内部应用有预建集成;自有/专有系统经 API/SDK 接入。", "tag":"强"},
 {"lead":"中立悖论(分析)— ", "t":"OpenAI 宣称跨平台中立,但 Frontier 与自家 GPT-5.x 深度绑定;批评者问:为何把「工作流自动化平台」绑死在「你的 LLM 供应商」上?→ 主张「LLM 中立的控制平面」。", "color":AMBER, "tag":"中"},
 {"lead":"执行风险 — ", "t":"中立究竟减少摩擦、还是「又一个要被理顺的平台」,取决于集成深度、运营简洁度与生态规模。", "tag":"中"},
], y=Inches(1.55), size=13.5, gap=11)

# ============================================================ SLIDE 19 — FRONTIER CUSTOMERS & RESULTS
s = content_slide("FRONTIER 5/6 · EVIDENCE", "早期客户与「成效」:信号强,但口径需谨慎", P())
bullets(s, [
 {"lead":"早期客户(多源一致)— ", "t":"Uber · State Farm · Intuit · Thermo Fisher Scientific;部分报道另列 HP · Oracle。先向少量客户限量开放,「未来数月」逐步放开。", "tag":"强(核心名单)"},
 {"lead":"成效宣称(OpenAI 自报,单一来源,无第三方审计)— ", "t":"某制造商把「生产优化」从 6 周缩到 1 天;某能源企业产出提升最高 5%,「增收超十亿美元」。", "color":AMBER, "tag":"中-弱"},
 {"t":"解读:这些是厂商精选案例(cherry-picked),服务于「企业机会规模」叙事;方向性可参考,不可当审计后的基准。", "lvl":1, "color":AMBER, "bullet":"⚠  "},
 {"lead":"定价 / 可用性 — ", "t":"未公开,定制报价,引导联系销售;据报高于 ChatGPT Enterprise。", "tag":"中"},
 {"t":"待核实:成效的分母 / 基线 / 可复现性;客户是「已投产」还是「试点」;GA 时间表与席位门槛。", "lvl":1, "color":GRAY, "bullet":"○  "},
], y=Inches(1.5), size=13.5, gap=11)

# ============================================================ SLIDE 20 — FRONTIER BUSINESS / SAAS COLLISION
s = content_slide("FRONTIER 6/6 · STRATEGY", "商业逻辑:与 SaaS「正面相撞」的企业收入引擎", P())
bullets(s, [
 {"lead":"收入转向 — ", "t":"企业收入已占 OpenAI 总营收 >40%,目标年底约 50%、与消费持平;Frontier 被定位为这一目标的「载体」。", "tag":"中"},
 {"lead":"与 SaaS 的碰撞 — ", "t":"分析普遍认为 Frontier 把 OpenAI 推上与 Salesforce / ServiceNow / Workday 乃至「最亲密伙伴」Microsoft 正面竞争的轨道——agent 直接在「系统记录」上干活,可能侵蚀传统 SaaS「人坐在 UI 前操作」的价值。", "tag":"强"},
 {"lead":"「拥有企业 agent 栈」— ", "t":"媒体框定:OpenAI 想从模型一路向上「拥有」编排层(Frontier)+ 超级应用(ChatGPT + Codex),贯通建模到分发。", "tag":"中"},
 {"lead":"软肋 — ", "t":"OpenAI 缺少在位者多年的「行业 know-how + 企业销售 / 合规」积累;复杂度越高,越吃领域经验。", "color":AMBER, "tag":"中"},
 {"lead":"安全门槛 — ", "t":"必须证明能「保护 agent(而非仅 API 调用)」、对接客户身份体系并支撑 agentic 工作负载,才能赢得受监管行业。", "tag":"中"},
], y=Inches(1.5), size=13, gap=10)

# ============================================================ SLIDE 21 — CODEX x FRONTIER SYNTHESIS
s = content_slide("SYNTHESIS", "Codex × Frontier:同一 Agent 战略的「纵深」与「横向」", P())
bullets(s, [
 {"lead":"同源同日 — ", "t":"2026-02-05,GPT-5.3-Codex 与 Frontier 同日发布,是协调的「agent 攻势」一体两面,而非孤立产品。", "tag":"强"},
 {"lead":"纵深(Codex)— ", "t":"把「软件工程」这一个垂直做到极致:模型×harness×沙箱三位一体的自主编码 coworker;回答「单个 agent 能多强」。", "tag":"强"},
 {"lead":"横向(Frontier)— ", "t":"把 agent 推广到「所有职能」,补齐上下文 / 治理 / 评估,让它们可被企业像员工一样管理;回答「一群 agent 如何在企业里安全协作」。", "tag":"强"},
 {"lead":"衔接点 — ", "t":"GPT-5.3-Codex 作为复杂 agentic 推理引擎,可驱动 Frontier 上更复杂的 agent 操作;Codex 式「编码 coworker」可作为 Frontier 治理下的一类 agent 被部署、监控与评估。", "tag":"中"},
 {"lead":"一句话 — ", "t":"两条战线共同把竞争从「单轮代码质量」推向「企业级 agent 操作系统」——谁能把『强单体 agent』与『可治理的 agent 群』同时做好,谁赢下一阶段。", "color":TEAL, "tag":"分析"},
], y=Inches(1.5), size=13.5, gap=11)

# ============================================================ SLIDE 22 — ENTERPRISE AGENT PLATFORM COMPETITION
table_slide(
 "ENTERPRISE COMPETITION", "企业 Agent 平台之争:Frontier vs 在位者", P(),
 ["平台", "归属 / 定位", "优势", "锁定 / 软肋"],
 [
  ["OpenAI Frontier", ("OpenAI;多供应商「控制平面」+ 语义层", INK, False), ("顶尖模型 + 中立叙事 + 治理一等公民", GREEN, False), ("企业经验浅;与自家 LLM 深绑的中立悖论", AMBER, False)],
  ["Salesforce Agentforce", ("Salesforce;2024 秋 GA", INK, False), ("部署面最广、最贴近 CRM 业务", INK, False), ("Atlas Reasoning Engine 编排锁定最紧", AMBER, False)],
  ["MS Copilot Studio / Frontier Suite", ("Microsoft;M365 E7($99/用户,5/1 GA)", INK, False), ("分发与采购优势、深植 M365", INK, False), ("与 OpenAI 既合作又竞争;名称撞名", AMBER, False)],
  ["IBM watsonx Orchestrate", ("IBM;受监管行业", INK, False), ("治理 / 合规与混合部署见长", INK, False), ("模型前沿性与生态偏弱", AMBER, False)],
  ["Snowflake Cortex / ServiceNow", ("数据 / 工作流就近编排", INK, False), ("贴近数仓 / ITSM 工作流", INK, False), ("跨域通用性有限", AMBER, False)],
 ],
 col_w=[3.0, 3.7, 3.1, 3.2],
 note="命名提醒:Microsoft 亦有「Frontier Suite」,与 OpenAI Frontier 同名不同物 · 市场普遍列五大竞争者:Frontier / Agentforce / watsonx Orchestrate / Copilot Studio / Snowflake Cortex",
 fs=10, hfs=10.5, row_h=Inches(0.82), y=Inches(1.45))

# ============================================================ SLIDE 23 — CODING-TOOL COMPETITION (three kings)
s = content_slide("CODING-TOOL LANDSCAPE", "竞争格局(编码工具层):三王分治 + 多工具叠用", P())
bullets(s, [
 {"lead":"三个「第一」按口径划分 — ", "t":"Copilot 赢用户基数(~470 万付费);Cursor 赢营收(~$2B ARR,传 $50B 估值);Claude Code 赢满意度/口碑(JetBrains 2026 调查最受欢迎,CSAT 91%)。", "tag":"强"},
 {"lead":"JetBrains 2026 调查(>1 万开发者)— ", "t":"职场采用率 Copilot ~29% / Cursor ~18% / Claude Code ~18%;Claude Code 一年内从 3% 增至 18%(6x)。", "tag":"强"},
 {"lead":"叠用成常态 — ", "t":"~70% 工程师同时用 2–4 个 AI 编码工具;创业公司偏 Claude Code(~75%),万人以上大厂偏 Copilot(~56%,采购/分发优势)。"},
 {"lead":"格局剧变(2026)— ", "t":"微软与 OpenAI 结束独家绑定(~4 月),Copilot 转多模型(接 OpenAI+Anthropic+自研 MAI/Polaris);Google 下线 Gemini CLI 转 Antigravity(5 月);Windsurf 更名 Devin Desktop(6 月)。", "tag":"强"},
 {"t":"注意:此为「编码工具层」之争;企业 agent 平台层(Frontier vs Agentforce 等)见上一页,两层不可混为一谈。", "lvl":1, "color":TEAL, "bullet":"➔  "},
], y=Inches(1.52), size=13.5, gap=11)

# ============================================================ SLIDE 24 — BUSINESS (Codex)
s = content_slide("CODEX · BUSINESS & STRATEGY", "Codex 商业与市场策略", P())
bullets(s, [
 {"lead":"定价 — ", "t":"捆绑 Plus/Pro/Business/Enterprise/Edu,不单售;2026-04-02 由按消息改按 token 信用计费(社区不满导火索之一)。登录态走套餐额度,切 API-key 则按 API 单价。", "tag":"强"},
 {"lead":"采用 — ", "t":"周活 200 万(3 月)→ 400 万(4.21)→ 500 万+(6 月),~20% 非开发者;OpenAI 内部用 Codex 的工程师从 <50% 升至 >90%,PR 量 +~70%。", "tag":"中(自报)"},
 {"lead":"营收战略 — ", "t":"编码是 LLM 最高价值变现场景(有分析称占企业生成式 AI 用量约半);The Information 称 OpenAI 2026 Q1 营收近 $60 亿、季度领先 Anthropic ~$10 亿。Claude Code 据报 >$2.5B 年化,是对标标尺。", "tag":"中"},
 {"lead":"Windsurf 始末 — ", "t":"OpenAI 拟 ~$30 亿收购 → 2025-07-11 因(据报)与微软 IP 摩擦告吹 → Google 以 ~$24 亿「逆向收编」CEO 等核心 → 余下由 Cognition(Devin)接手。", "tag":"强"},
 {"lead":"争议 — ", "t":"GPT-5.3-Codex 罕见附严格管控(称或首次「足以放大现实网络危害」);2026-06 冒充 Codex 的恶意 npm 包窃 token;2025 末 GitHub token 泄露漏洞(已修)。", "tag":"强"},
], y=Inches(1.5), size=13, gap=9)

# ============================================================ SLIDE 25 — DEVELOPER VOICE
s = content_slide("CODEX · DEVELOPER VOICE", "Codex 开发者实践与口碑", P())
bullets(s, [
 {"lead":"被称赞 — ", "t":"彻底/全面(不止改目标函数,连带修调用方、统一日志、标记弃用);「委派—回看」异步流(早上排 3–5 个任务,回来已有审好的 PR);干净 diff;自动 PR 审查。", "color":GREEN},
 {"lead":"被吐槽 — ", "t":"沙箱默认无网阻断装依赖;2026-04 转 token 计费后限流抱怨集中爆发(「5 小时额度一小时用光」);容器内跑测试偏慢;默认过度工程;不能手动选模型。", "color":AMBER},
 {"lead":"口碑拐点 — ", "t":"GPT-5-Codex(2025-09)是关键转折:Builder.io 测得用户评分高出替代品约 40%。但成因有争议——是 Codex 真变强,还是 Claude Code 同期「被阉割/harness 变怪」,两种叙事并存。", "tag":"中"},
 {"lead":"当前共识:双持(dual-wield)— ", "t":"Codex 做后端/重测试/重构与「可辩护的 diff」;Claude Code 做 UI/探索式规划/自主部署。「问题不是谁更聪明,而是你在哪种工作流里。」", "color":TEAL},
 {"lead":"最佳实践 — ", "t":"用好 AGENTS.md(含测试命令)· prompt 显式约束范围以抑过度工程 · 合理设审批档(--full-auto 常为平衡点)· 用子 agent/REQUIREMENTS.md 做结构化交接。"},
], y=Inches(1.5), size=13, gap=9)

# ============================================================ SLIDE 26 — KEY JUDGMENTS
s = content_slide("KEY JUDGMENTS", "核心判断(分析层)", P())
bullets(s, [
 {"lead":"Codex 护城河在「模型×harness 协同设计」 — ", "t":"GPT-5.x-Codex 与其 agent 循环共同 RL 调优,带来更紧的工具使用与 token 效率;代价是 OpenAI 模型锁定。Cursor/Copilot 以可移植性换取这种垂直整合。"},
 {"lead":"Frontier 的赌注是「治理与上下文」而非模型 — ", "t":"企业把 agent 投产的瓶颈在权限/审计/共享上下文/持续评估;Frontier 把这些做成一等公民——若兑现,护城河来自「企业系统集成深度」,而非单纯模型领先。"},
 {"lead":"长时程自主 + 企业编排是新战线 — ", "t":"Codex 的 compaction/24h+ 与 Frontier 的多 agent 治理,把竞争从「单轮代码质量」推向「跨窗口状态管理 + 企业级 agent 编排」。"},
 {"lead":"成熟度的反面是攻击面 — ", "t":"Slack/SDK/插件 + 企业系统记录的自主访问,提升渗透的同时带来供应链、OAuth/身份、数据外泄等新安全面——企业采用须纳入评估。"},
 {"lead":"认知风险:别被「谁第一」与「撞名」带偏 — ", "t":"品牌跨三代、版本月月变、跑分口径混乱,且『OpenAI Frontier』≠『Microsoft Frontier Suite』;务实做法是按工作流与治理需求组合工具。", "color":TEAL},
], y=Inches(1.5), size=13, gap=9)

# ============================================================ SLIDE 27 — SELECTION GUIDE
table_slide(
 "RECOMMENDATIONS", "选型建议:按场景对号入座", P(),
 ["你的场景", "首选", "理由"],
 [
  ["异步委派 / 并行批量 / CI 自动化", ("Codex(云)", GREEN, True), "隔离容器 + worktree + best-of-N;免审批自主更可控"],
  ["数据敏感 / 隔离网 / 强合规(编码)", ("Claude Code", TEAL, True), "本地优先执行 + 应用层细粒度权限 + hooks 治理"],
  ["长时程自主重构 / 长任务", ("Codex(5.1-Max+)", GREEN, True), "原生 compaction 跨窗口,24h+ 连贯"],
  ["跨职能 agent 投产 + 治理/审计", ("OpenAI Frontier", BLUE, True), "共享业务上下文 + 评估 + 权限/审计一等公民(早期访问)"],
  ["已重度绑定某 SaaS 生态", ("Agentforce / Copilot Studio", AMBER, True), "贴近 CRM / M365 业务与采购;但编排锁定更紧"],
  ["务实团队", ("双持 + 平台中立", INK, True), "Codex/Claude 跑编码;Frontier/在位者按生态选编排层"],
 ],
 col_w=[4.4, 2.6, 5.3],
 note="基于本次研究的工程事实与社区共识综合;非厂商背书 · Frontier 为早期访问、定价未公开,落地前须验证集成深度与合规",
 fs=11, hfs=12, row_h=Inches(0.66), y=Inches(1.5))

# ============================================================ SLIDE 28 — FRONTIER RISKS & VERIFICATION
s = content_slide("FRONTIER · RISKS & VERIFICATION", "Frontier 风险与待核验项(诚实披露)", P(),
                  note="与 Codex 风险页(下一页)并读 · 标记:✗ 已剔除 / ~ 已降级 / ○ 存疑 / ⚠ 易错点")
bullets(s, [
 {"lead":"名称撞名(易错点)— ", "t":"Microsoft「Frontier Suite」(M365 E7,$99/用户)与 OpenAI Frontier 同名不同物;引用务必区分厂商,否则结论张冠李戴。", "color":AMBER, "bullet":"⚠  "},
 {"lead":"疑似虚构来源(已剔除)— ", "t":"网传「Frontier Alliance / DeployCo / Codex Labs」等名目仅见于单一非权威页面,与官方命名不符,判为不可信(与此前『Codex 2.0/Orchestrator』同类模式)。", "color":RED, "bullet":"✗  "},
 {"lead":"成效未经审计(已降级)— ", "t":"「6 周→1 天」「+5% 产出 / 增收十亿」均为 OpenAI 自报、单一来源,无第三方核验;按厂商精选案例对待。", "color":AMBER, "bullet":"~  "},
 {"lead":"中立悖论(存疑)— ", "t":"「LLM 中立的控制平面」与「深绑 GPT-5.x」存在张力;真实中立程度取决于第三方 agent/连接器的集成深度,尚待观察。", "color":GRAY, "bullet":"○  "},
 {"lead":"信息时效(限制)— ", "t":"openai.com / 主流媒体 WebFetch 全程 403,Frontier 结论基于多源搜索摘要交叉验证;定价 / GA 时间 / 客户名单请以 openai.com/business/frontier 与官方 system card 为准。", "color":GRAY, "bullet":"○  "},
], y=Inches(1.5), size=13, gap=10)

# ============================================================ SLIDE 29 — RISKS & UNCERTAINTIES (Codex)
s = content_slide("CODEX · RISKS & UNCERTAINTIES", "Codex 风险与不确定性(诚实披露)", P())
bullets(s, [
 {"lead":"已证伪(剔除)— ", "t":"「Codex 2.0 / Orchestrator / SafeDeploy / OpenRepo」单一来源虚构,与官方命名矛盾。", "color":RED, "bullet":"✗  "},
 {"lead":"已降级(传闻)— ", "t":"「4x token 效率」为 n=1 博客且含产出质量混淆;「88.7% SWE-bench」为厂商自报(独立 ~82.6%)。", "color":AMBER, "bullet":"~  "},
 {"lead":"机制存疑 — ", "t":"GPT-5-Codex 是否「绝无内部路由」无法证实;compaction 是否「不止于训练版摘要再续」未完全公开;MoE/稀疏激活仅第三方推测。", "color":GRAY, "bullet":"○  "},
 {"lead":"数据口径 — ", "t":"周活/营收多为 OpenAI 自报或单一媒体(The Information),无第三方审计;标称上下文 ≠ 会话实际上下文。", "color":GRAY, "bullet":"○  "},
 {"t":"建议:正式引用前对照官方 system card PDF 与 developers.openai.com/codex 原文核对确切数字与日期。", "lvl":1, "color":TEAL, "bullet":"➔  ", "bold":True},
], y=Inches(1.55), size=13.5, gap=12)

# ============================================================ SLIDE 30 — SOURCES
s = content_slide("SOURCES", "主要来源(节选)", P(), note="完整 URL 见随附报告;以下为各维度最权威/最具价值来源")
bullets(s, [
 {"lead":"OpenAI 官方(Codex)— ", "t":"introducing-codex · introducing-upgrades-to-codex · gpt-5-1-codex-max(+system card)· introducing-gpt-5-3-codex · introducing-gpt-5-5 · codex-now-generally-available · building-codex-windows-sandbox · developers.openai.com/codex/*", "size":11.5},
 {"lead":"OpenAI 官方(Frontier / 企业)— ", "t":"introducing-openai-frontier · openai.com/business/frontier · next-phase-of-enterprise-ai · introducing-agentkit · a-business-that-scales-with-the-value-of-intelligence", "size":11.5},
 {"lead":"媒体 / 分析(Frontier)— ", "t":"TechCrunch · CNBC · Fortune · VentureBeat · Computerworld · InfoQ · SiliconANGLE(Frontier + GPT-5.3-Codex 同日)· Constellation Research · Futurum · artificialintelligence-news", "size":11.5},
 {"lead":"Codex 技术 / 一手 — ", "t":"github.com/openai/codex(codex-rs / linux-sandbox README)· cdn.openai.com system card · simonwillison.net · infoq.com · badlogic gist(compaction)· swebench.com / epoch.ai", "size":11.5},
 {"lead":"竞争 / 市场 — ", "t":"JetBrains Research 2026 · The Information(Q1 营收)· softwareseni / smartbridge(平台对比)· devin.ai(Windsurf→Devin)· Google Developers Blog(Antigravity)", "size":11.5},
 {"t":"研究方式:多路并行检索 + 对抗式核验(2/3 证伪即剔除);可信度全程分层标注;本次 WebFetch 受限,以搜索摘要交叉验证为主。", "lvl":1, "color":GREEN, "bullet":"✔  ", "bold":True, "size":11.5},
], y=Inches(1.5), size=11.5, gap=9)

out = "/home/user/openai-chatgpt-codex/OpenAI_Codex_Frontier_深度洞察.pptx"
prs.save(out)
print("saved:", out, "slides:", len(prs.slides._sldIdLst))
